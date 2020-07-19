# Name: Assign Horizon User Desktop

Description: This script assigns a user to a Horizon desktop machine. This will only work with dedicated desktop pools.It will receive the connection server fqdn, Desktop pool and machine, login and domain names from the CU Console.

This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console). The script uses the target machine to determine the connection server address, and is executed on the machine running ControlUp Console.

This script requires VMWare PowerCLI  to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'

Version: 1.3.5

Creator: Wouter Kursten

Date Created: 02/20/2020 12:32:31

Date Modified: 02/25/2020 20:20:53

Scripting Language: ps1

