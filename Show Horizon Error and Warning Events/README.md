# Name: Show Horizon Error and Warning Events

Description: Uses the Horizon PowerCLI api's to pull all Error, Warning and Audit_Fail events from the Horizon Event database for all pods. If there is no cloud pod setup it will only process the local pod. After pulling the events it will translate the id's for the various objects to names to show the proper names where needed. Requires Horizon 7.5 or later

Output is displayed in the console but also saved to a default location of c:windows\temp\CU_Horizon_error_log.csv

This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console). The script uses the target machine to determine the connection server address, and is executed on the machine running ControlUp Console.
This script requires Horizon Credentials to be set for the account running the scipt on the target machine, these need to be created using the 'Create credentials for Horizon scripts' Script Action
This script requires VMWare PowerCLI  to be installed on the machine running the script.
PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers' or by using the 'Install and configure VMware PowerCLI'Script Action

Version: 1.3.6

Creator: Wouter Kursten

Date Created: 08/27/2020 09:54:22

Date Modified: 09/15/2020 09:18:32

Scripting Language: ps1

