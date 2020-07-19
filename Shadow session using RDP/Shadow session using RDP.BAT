@ECHO OFF
REM  .SYNOPSYS  Leverage MSTSC for Shadow and remote control 
REM  .AUTHOR    Marcel Calef 2019-11-09
REM  .EXAMPLE   shadow.bat 1 userPCname /control no
REM  .REQUIREMENTS
REM 	Allow shadow via GPO
REM 	    Policies -> ADM Templates -> Windows components -> Remote Desktop Services -> Remote Session Host -> Connections
REM 	        'Set rules for remote control of Remote Desktop Services user sessions'
REM 	or via Registry:
REM 	    HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\Shadow [DWORD]
REM 	        0=not allowed, 1=control with consent, 2= control w/o consent, 3=view with consent, 4=view only w/o consent
REM	
REM 	see https://support.controlup.com/hc/en-us/articles/360000602929-Shadow-RDS-Session-without-prompt
REM	
REM 	Firewall (particularly Win10):   Make sure you are allowing "Remote Desktop -Shadow (TCP-In)" in the Firewall
REM 
REM  .TAGS $ID,$State="Active" 

SET SessionID=%1
SET RemoteComputer=%2
SET addControl=%3
SET consent=%4
SET FQDN=%2.%5

REM adjusting the input to the proper MSTSC switches: /control and /noConsentPrompt  or omitting them
IF /I NOT %addControl%==/control SET addControl=
IF /I %consent%==no (SET consent=/noConsentPrompt) ELSE (SET consent= )

REM addig %6 if provided to prompt for credentials
mstsc /shadow:%SessionID% /v:%FQDN% %addControl% %consent% %6

