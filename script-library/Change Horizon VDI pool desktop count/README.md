# Name: Change Horizon VDI pool desktop count

Description: Changes the amount VDI desktops in a desktop pool

This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console). The script uses the target machine to determine the connection server address, and is executed on the machine running ControlUp Console.

This script requires VMWare PowerCLI  to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'


Version: 2.5.8

Creator: Wouter Kursten

Date Created: 01/09/2020 09:25:40

Date Modified: 02/13/2020 08:34:24

Scripting Language: ps1

