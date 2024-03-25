@echo off

REM Some messages go to error so redirect to standard output

chglogon.exe /%1 2>&1
