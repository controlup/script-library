@ECHO OFF
:: NAME: 	Stuck at 85%? Allow in the Firewall ControlUp TCP port 40705
:: 
:: DESCR:   This script is designed to run from a management computer (ControlUp Console).
::          Using a WMI command generates on a remote computer an inbound connections rule in
::          Windows Firewall to allow a ControlUp agent deployment stuck at 85% to complete.
::
:: RUN ON :	ControlUp Console
:: .ARGS:	%1 -> TCP Port
::      	%2 -> Computer Name

SET tcpPort=%1
SET remoteNODE=%2

ECHO Creating the firewall allow rule to accept TCP %tcpPort% on the selected computer &echo.

wmic /node:"%remoteNODE%" process call create "netsh advfirewall firewall add rule dir=in name=ControlUp__%tcpPort% action=allow protocol=TCP localport=%tcpPort%"

ECHO.
ECHO Please try connecting again. 
