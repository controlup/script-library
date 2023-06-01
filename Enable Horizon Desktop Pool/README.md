# Name: Enable Horizon Desktop Pool

Description: Enables a Horizon Desktop pool

Can be used as an automated or manual action to enable a Horizon Desktop pool for planned maintenance. 

This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console) which is part of the relevant Desktop Pool. The script uses the target machine to determine the connection server address and the Desktop Pool name, and is executed on the machine running ControlUp Console.

This script requires VMWare PowerCLI to be installed on the machine running the script.

PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'

Version: 5.11.55

Creator: Ton de Vreede

Date Created: 08/22/2019 12:38:01

Date Modified: 05/24/2023 13:04:39

Scripting Language: ps1

