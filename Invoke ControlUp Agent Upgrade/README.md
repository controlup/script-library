# Name: Invoke ControlUp Agent Upgrade

Description: This script uses the ControlUp Automation PowerShell Module to perform an in-place upgrade of the ControlUp Agent. To execute the upgrade from the agent itself, the script creates an upgrade script in C:\Windows\Temp and a scheduled task. After the installation, the scheduled task will be deleted, but the script will remain.

The script requires the ControlUp Automation PowerShell Module to be installed. If it is not installed, the script will install it automatically.

When the TAGS parameter is set to true, any installation errors will be written to ControlUp Tags. To view these errors, you may need to enable the Tags column in ControlUp. If the installation completes successfully, no tag will be created.

Version: 3.8.38

Creator: Chris Twiest

Date Created: 10/15/2024 16:29:27

Date Modified: 02/27/2025 14:09:57

Scripting Language: ps1

