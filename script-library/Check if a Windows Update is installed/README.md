# Name: Check if a Windows Update is installed

Description: The script checks if a Windows Update is installed, using Get-WMIObject commandlet and WMI information under 'win32_quickfixengineering' and 'win32_ReliabilityRecords'. 

About Win32_ReliabilityRecords:
They are enabled by default on Windows 7 but Group Policy has to be used to enable them on the server side. See Computer Settings – Administrative Templates – Windows Components – Windows Reliability Analysis. 

MSDN Win32_ReliabilityRecords class https://msdn.microsoft.com/en-us/library/ee706630(v=vs.85).aspx


Version: 1.12.57

Creator: ajal

Date Created: 02/21/2015 14:14:25

Date Modified: 04/05/2016 21:25:29

Scripting Language: ps1

