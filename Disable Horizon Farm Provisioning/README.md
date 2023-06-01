# Name: Disable Horizon Farm Provisioning

Description: Disables provisioning on a Horizon RDS Farm

Can be used as an automated or manual action to disable provisioning for a Horizon RDS Farm

This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console) which is part of the relevant Desktop Pool. The script uses the target machine to determine the connection server address and the Desktop Pool name, and is executed on the machine running ControlUp Console.

This script requires VMWare PowerCLI to be installed on the machine running the script.

PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'

Version: 2.1.9

Creator: Wouter Kursten

Date Created: 04/06/2023 14:10:02

Date Modified: 05/23/2023 07:07:55

Scripting Language: ps1

