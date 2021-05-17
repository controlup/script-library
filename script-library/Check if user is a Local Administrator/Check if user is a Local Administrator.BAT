@echo off
::  NAME   : Check if user is a Local Administrator
::           This script will check if the user is a member of BUILTIN\Administrators, directly named 
::           or inherited from other local or domain groups
::  CONTEXT: User Session
::  HISTORY: 2020-06-10 Marcel Calef - initial release

:checl_localAdmin

whoami /groups | findstr "BUILTIN\Administrators">NUL
if %errorLevel% == 0 (
        echo   User is a Local Administrator.
    ) else (
        echo   No local admin rights.
    )
:: timeout 15>nul  
:: pause >nul
