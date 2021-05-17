@ECHO OFF
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::  Synopsis:  Windows license activation check
::  Details:   Check the windows licenses activation using the slmgr.vbs script and parse significant info
::             Verbose=No will read the Status line and include the Notification if not properly licensed
::  Credit:    https://superuser.com/users/942348/zer0byt3 for how to display in the output instead of a popup
::                                    invoke slmgr using: cscript //nologo
::             https://stackoverflow.com/questions/31035636/batch-store-command-output-to-a-variable-multiple-lines
::  Context:   Machine
::  Tags:      $OperatingSystem="*Windows*"
::	History:   2020-06-01 Marcel Calef - initial release
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

SET OpSYS=%1    
SET verboseLevel=%2

if "%verboseLevel%"=="Yes" (
	ECHO %OpSYS%
	ECHO.
	ECHO ======================================================================
	cscript //nologo c:\windows\System32\slmgr.vbs /dlv | findstr /C:"Name:" /C:"Key Channel:" /C:"Product Key:" /C:"Status" /C:"Licensed" /C:"Notification"
	)
	
if "%verboseLevel%"=="No"  (
	:: Read the Status line and include the Notification if not properly licensed
	cscript //nologo c:\windows\System32\slmgr.vbs /dlv | findstr /C:"Licensed" /C:"Notification"
	)
