@ECHO OFF
REM =======  Reduce Bandwidth consumed by an Citrix HDX session  =======
REM When connecting to a remote session, the more graphics data sent the more bandwidth is needed.
REM While modern protocols like HDX and Blast adjust, when task workers have extremely low bandwidth available
REM   (less than 1 or 2 Mbps) they might benefit from a simple trick: Reduce the maximum frames per second sent instead
REM   of sacrificing image quality and/or responsiveness.
 
REM This script modifies the VMware Blast_VDI and Citrix sesison registry keys and therefore it is recommended
REM  to implement policies that would revert this to the desired defaults (periodically or upon reconnect)
 
REM INPUTS
set Workstation=%1
set SessionID=%2
set fps=%3
set /A minfps=%fps%/2

REM only the fourth or fifth value will come be present (in either case as the 4th variable)
set SelectProto=%4
ECHO VMware Protocol or Citrix Receiver version received is  %SelectProto%
ECHO.

IF [%SelectProto%] EQU [] GOTO NONE
IF [%SelectProto%] EQU [BLAST] GOTO BLAST
GOTO HDX

:HDX
set ctxPol=HKLM\SOFTWARE\Policies\Citrix

ECHO Reduce frame rate to %fps%
reg add %ctxPol%\%SessionID%\User\GraphicsPolicies /v FramesPerSecond /t REG_DWORD /d %fps% /f
reg add %ctxPol%\%SessionID%\User\GraphicsPolicies /v TargetedMinimumFramesPerSecond /t REG_DWORD /d %minfps% /f

REM Additional options to reduce bandwidth
REM reg add %ctxPol%\%SessionID%\User\GraphicsPolicies /v VisualQuality /t REG_DWORD /d 4 /f
REM reg add %ctxPol%\%SessionID%\User\MultimediaPolicies /v SoundQuality /t REG_DWORD /d 0 /f

REM if desired, restart the ctxgfx for the values to take effect
REM for /f "tokens=2 delims= " %%a in ('tasklist /NH /FI "Session eq %SessionID%" ^| findstr ctxgfx') do SET CtxGfxPID=%%a
REM taskkill /F /PID %CtxGfxPID%

GOTO END

:NONE
ECHO.
ECHO Unsopported session selected (not a HDX or VMware Blast_VDI session)
GOTO END

:BLAST
set VMWPol="HKLM\SOFTWARE\Policies\VMware, Inc.\VMware Blast\Config"

REM check if this is a VDI running on a Workstation and not a RDS/Server
IF %Workstation% NEQ True GOTO NONE

ECHO Reduce frame rate to %fps%
reg add %VMWPol% /v EncoderMaxFPS /t REG_SZ /d %fps% /f

:END
