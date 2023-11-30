# Name: Change Horizon VDI pool desktop count

Description: Changes the amount VDI desktops in a desktop pool. Use UP_FRONT or ON_DEMAND for the Provisioningtype depending if you want to provision all desktops up front. If UP_FRONT is used minNumberOfMachines and numberOfSpareMachines will be ignored.

This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console). The script uses the target machine to determine the connection server address, and is executed on the machine running ControlUp Console.

This script requires VMWare PowerCLI  to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'


Version: 4.7.13

Creator: Wouter Kursten

Date Created: 01/09/2020 09:25:40

Date Modified: 11/24/2023 10:57:23

Scripting Language: ps1

