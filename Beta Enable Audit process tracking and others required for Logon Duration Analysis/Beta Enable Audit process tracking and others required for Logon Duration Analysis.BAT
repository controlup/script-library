REM Enable Command Line Auditing
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Audit" /v ProcessCreationIncludeCmdLine_Enabled /d 0x1 /f

REM Enable Print Service logging, no retention, size 50MB
wevtutil sl Microsoft-Windows-PrintService/Operational /ms:52428800 /rt:false /ab:false /e

REM Enable Group Policy logging, no retention, size 50MB
wevtutil sl  Microsoft-Windows-GroupPolicy/Operational  /ms:52428800 /rt:false /ab:false /e

REM Enable Process Termination
Auditpol /set /subcategory:"Process Termination" /success:enable

REM Enable Process Creation
Auditpol /set /subcategory:"Process Creation" /success:enable

REM Enable Logon Audit
Auditpol /set /subcategory:"Logon" /success:enable

REM Additional - not sure if 100% required:

Auditpol /set /subcategory:"Network Policy Server" /success:enable

