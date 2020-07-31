:: NAME: Enable requirements for Logon Duration Analysis
::   This script is designed to enhance the login measurements in a computer using auditpol, 
::   wevutil and reg commands and complete the requisites for the Analyze Logon Duration Script.
::   It will additionally increase the size of the Group Policy and Print Service Event logs 
::   to retain the information for a couple of hours after login.
::   Important note: This script modifies the audit policies on a computer and should be used with caution. 
::   If in doubt consult with your IT Security team.
::
:: CREDIT:       Guy Leech, Trentent Tye 
:: PUBLISHED BY: Marcel Calef

:: Enable Command Line Auditing
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Audit" /v ProcessCreationIncludeCmdLine_Enabled /d 0x1 /f

:: Enable Print Service logging, no retention, size 50MB
wevtutil sl Microsoft-Windows-PrintService/Operational /ms:52428800 /rt:false /ab:false /e

:: Enable Group Policy logging, no retention, size 50MB
wevtutil sl  Microsoft-Windows-GroupPolicy/Operational  /ms:52428800 /rt:false /ab:false /e

:: Enable Process Termination audit policy
Auditpol /set /subcategory:"Process Termination" /success:enable

:: Enable Process Creation audit policy
Auditpol /set /subcategory:"Process Creation" /success:enable

:: Enable Logon Audit policy
Auditpol /set /subcategory:"Logon" /success:enable

