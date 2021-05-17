@ECHO OFF
REM  ====================  Retrieve the Parent Process  ================================
REM  Retrieve the details of a process's parent process. 
REM  A WMI command is invoked to collect the Parent's PID and then used to filter tasklist results
REM  Additional output options REMarked

set PID=%1
ECHO.
ECHO Retriving the Parent process for PID %PID%
ECHO.
for /f "usebackq tokens=2 delims==" %%a in (`wmic process where ^(processid^=%PID%^) get parentprocessid /value`) do set parent=%%a

REM add the /NH switch if you want ot omit the headers
tasklist /FI "PID eq %PARENT%"


REM to get just the executalbe printed out un-REMark the following line:
REM for /f "tokens=1 delims= " %%a in ('tasklist /NH /FI "PID eq %PARENT%"') do ECHO %%a
	
