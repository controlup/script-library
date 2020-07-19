# Name: Show Horizon admin audit trail

Description: Requires Horizon 7.5 or later

Uses the Horizon PowerCLI api's to pull all admin related events from the Horizon Event database for all pods. If there is no cloud pod setup it will only process the local pod. After pulling the events it will translate the id's for the various objects to names to show the proper names where needed.

Output is displayed in the console but also saved to a default location of c:windows\temp\CU_Horizon_audit_log.csv

This script requires VMware PowerCLI to be installed on the machine running the script. PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMware.PowerCLI -Force -AllowCLobber -Scope AllUsers'



Version: 1.4.11

Creator: Wouter Kursten

Date Created: 04/28/2020 13:00:55

Date Modified: 06/03/2020 12:11:28

Scripting Language: ps1

