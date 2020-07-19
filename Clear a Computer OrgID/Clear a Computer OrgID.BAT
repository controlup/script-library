@ECHO OFF
REM  When a Computer was previously managed by another organization
REM      it will not accept connections form another (new) organization.
REM   To resolve this, the script will stop a running agent 
REM     and clear the registry settings locking a running agent to an Organization. 
REM   This script will also clear any previous ControlUp Agent customizations
REM   The script will restart the ControlUp Service at the end
REM   Documentation:
REM      https://support.controlup.com/hc/en-us/articles/207234285-Computer-already-belongs-to-another-ControlUp-organization
REM   Limitation: Target computers need to be resolved by hostname and the user executing the script must be a local admin in that computer
REM   Tags:  $Status="Error"

SET Target=%1

IF z%Target% == z (
	ECHO Missing argument of the computer to act on >&2
	TIMEOUT 5
	EXIT 1
)

SET ENABLEDREMOTEREG=No
SET STARTEDREMOTEREG=No

REM query the Target for Remote Registry service state - if not enable it and start it

sc \\%Target% query remoteregistry|findstr /r "STATE.*STOPPED" >NUL && (
	ECHO Remote registry service is not running on %Target% so attempting to start ...
	sc \\%Target% qc remoteregistry|findstr "START_TYPE.*DISABLED" >NUL && (
		ECHO Remote registry service is disabled so enabling
		sc \\%target%  config remoteregistry start= demand
		SET ENABLEDREMOTEREG=Yes
	)
	sc \\%target% start remoteregistry
	IF %ERRORLEVEL% NEQ 0 ECHO Failed to start remote registry service on %Target% >&2
	SET STARTEDREMOTEREG=Yes
)


ECHO Stopping ControlUp Agent ...
sc \\%target% stop cuAgent
IF %ERRORLEVEL% NEQ 0 ECHO Failed to stop cuAgent service on %Target% >&2

TIMEOUT /T 5

ECHO.
ECHO Attempting to delete the ControlUp Agent registry settings ...
reg delete \\%target%\HKLM\Software\Smart-X\ControlUp\Agent /f
IF %ERRORLEVEL% NEQ 0 ECHO Problem deleting ControlUp Agent registry key on %Target% >&2

ECHO.

REM Revert Remote Registry Service if needed
IF %ENABLEDREMOTEREG% == Yes (
	ECHO Remote registry service was disabled so disabling
	sc \\%target% config remoteregistry start= disabled
)

IF %STARTEDREMOTEREG% == Yes (
	ECHO Remote registry service was stopped so stopping
	sc \\%target% stop remoteregistry
)



ECHO Starting ControlUp Agent ...
sc \\%target% start cuAgent

IF %ERRORLEVEL% NEQ 0 (
	ECHO Failed to start cuAgent service on %Target% >&2
)

