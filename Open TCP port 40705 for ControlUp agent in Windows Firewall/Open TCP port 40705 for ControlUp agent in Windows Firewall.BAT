@ECHO OFF

REM .NAME 		    Deploy ControlUp firewall rule from the Console computer
REM 
REM .DESCRIPTION    This script is designed to run from a management computer (ControlUp Console).
REM                 Using a WMI command generates on a remote computer an inbound connections rule in
REM                  Windows Firewall to allow a ControlUp agent deployment stuck at 85% to complete.
REM 
REM .ASSIGNED TO:	Computer
REM .EXEC CONTEXT:	ControlUp Console
REM .SCRIPT TYPE:	CMD
REM .ARGUMENTS:		%1 -> Computer Name

SET remoteNODE=%1


ECHO.
ECHO Creating the firewall allow rule to accept TCP 40705 on the selected computer

wmic /node:"%remoteNODE%" process call create "netsh advfirewall firewall add rule dir=in name=ControlUp__40705 action=allow protocol=TCP localport=40705"

