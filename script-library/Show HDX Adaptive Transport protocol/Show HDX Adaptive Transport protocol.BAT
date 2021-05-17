@ECHO OFF
REM ========== Verify HDX Adaptive transport - TCP or UDP ==========
REM
REM credit:   https://goo.gl/uAGCgd stackoverflow 
REM                 how-to-remove-trailing-and-leading-whitespace...
REM
REM Author: Marcel Calef - 2018-12-20
REM Parameters:    session ID
REM                       Detail Level (not implemented yet)
REM ================================================================

REM Setting to remove leading spaces from the output
setlocal enabledelayedexpansion

REM populate in case there is no active HDX session
SET EDT=Not an active Citrix HDX session

REM populate if Citrix VDA is not installed in this computer
WHERE ctxsession > nul
IF %ERRORLEVEL% NEQ 0 ECHO Session is on a computer without Citrix VDA installed & Exit

REM  ctxsession command for the specific session 
REM       ctxsession -s %1 

for /f "tokens=2 delims=:^>-" %%G in ('ctxsession -s %1 ^|  findstr ICA') DO SET EDT=%%G
      REM ECHO DEBUG: before trim %EDT%
CALL :TRIM %EDT% EDT

ECHO %EDT%

GOTO :EOF

:TRIM
    SET %2=%1
    exit /b
