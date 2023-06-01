# Name: Horizon - Enable RDS Server

Description: Enables a Horizon RDS Server

Can be used as an automated or manual action to enableHorizon RDS Server after troubleshooting or maintenance.

This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console) which is part of the relevant Desktop Pool. The script uses the target machine to determine the connection server address and the Desktop Pool name, and is executed on the machine running ControlUp Console.

This script requires VMWare PowerCLI to be installed on the machine running the script.

PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'

Version: 2.0.3

Creator: Wouter Kursten

Date Created: 04/06/2023 13:34:44

Date Modified: 05/23/2023 06:14:25

Scripting Language: ps1

