@ECHO OFF
REM This script leverages OfficeIns.exe from https://www.nirsoft.net/utils/officeins.html

REM Create a random temporary directory
SET /a _rand=(%RANDOM%*30000/32768)+10001 
set tmpDir=%TEMP%\tmp_%_rand%
SET zip=officeins.zip
MKDIR %tmpDir%

if exist %TEMP%\OfficeIns.exe (
                    ECHO Found OfficeIns.exe - will NOT verify bitness....
					COPY %TEMP%\OfficeIns.exe %tmpDir%\.
					GOTO runOfficeIns
) else (
		rem file doesn't exist, download and extract
		SET zip=officeins.zip
		REM check if Office is 64bit and adjust download zip
		if %PROCESSOR_ARCHITECTURE% == AMD64 (
				REM ECHO Winx64 - check Office (Outlook path) and if (x86) NOT found i.e. Office is 64 bit
				reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe"  -d -f " (x86)"
				if %ERRORLEVEL% NEQ 0  SET zip=officeins-x64.zip
				)
		
		ECHO DEBUG: downloading %ZIP%
		bitsadmin /transfer downloadOfficeIns /download "https://www.nirsoft.net/utils/%ZIP%" "%tmpDir%\%ZIP%"
		ECHO Calling unzip
		Call :UnZipFile "%tmpDir%" "%tmpDir%\%ZIP%"
		GOTO runOfficeIns
	)

:runOfficeIns
DIR  %tmpDir%
ECHO running OfficeIns.exe
REM dump to file. see documentation for OfficeIns for possible output formats
%tmpDir%\OfficeIns.exe /stext %tmpDir%\add-ins.txt

REM Print out
type %tmpDir%\add-ins.txt

REM Clean up
    RMDIR  /Q /S %tmpDir%

exit /b

:UnZipFile <ExtractTo> <newzipfile>
set vbs="%temp%\_.vbs"
if exist %vbs% del /f /q %vbs%
>%vbs%  echo Set fso = CreateObject("Scripting.FileSystemObject")
>>%vbs% echo If NOT fso.FolderExists(%1) Then
>>%vbs% echo fso.CreateFolder(%1)
>>%vbs% echo End If
>>%vbs% echo set objShell = CreateObject("Shell.Application")
>>%vbs% echo set FilesInZip=objShell.NameSpace(%2).items
>>%vbs% echo objShell.NameSpace(%1).CopyHere(FilesInZip)
>>%vbs% echo Set fso = Nothing
>>%vbs% echo Set objShell = Nothing
cscript //nologo %vbs%
if exist %vbs% del /f /q %vbs%


REM Credits:
REM      https://www.nirsoft.net/utils/officeins.html
REM      https://docs.microsoft.com/en-us/windows/desktop/bits/bitsadmin-examples
REM      Extract in CMD: https://superuser.com/questions/1314420/how-to-unzip-a-file-using-the-cmd
