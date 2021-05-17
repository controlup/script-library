@ECHO OFF
REM    Analyze dirve contents with WinDirStat
REM    Run WinDirStat (https://windirstat.net/) to analyze the usage of a drive.
REM    If executed for a Logical drive that will be used.
REM    If executed on a computer or session will be run for ALL Drives on that computer

SET targetComputer=%1
SET DriveName=%2

REM Hardcoding the executable path for automatic actions:
SET execPath="C:\Program Files (x86)\WinDirStat\WinDirStat.exe"

if exist %execPath% (
        ECHO.
        ECHO Found %execPath% and will use it for Drive %DriveName% on Computer %targetComputer%
		GOTO runCommand
	) else   (
		ECHO . 1>&2
		ECHO .     windirstat.exe not foundat %execPath% 1>&2
		ECHO . 1>&2
		ECHO .           Download from https://windirstat.net/ 1>&2
		EXIT
		REM GOTO doneFlag
	)

	
:runCommand
	SET driveLetter=%DriveName:~0,1%
	ECHO.
	ECHO running %execPath%   -accepteula  \\%targetComputer%\%driveLetter%$
	REM Start command should have released the script, but not TO-DO
	%execPath%  -accepteula  \\%targetComputer%\%driveLetter%$

:doneFlag
	ECHO.
	ECHO Done
