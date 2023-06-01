# Name: Enable Horizon Pool Provisioning

Description: Enables VMware Horizon Virtual Desktop pool provisioning

Can be used as a manual or automated action to disable Horizon View Virtual Desktop pool provisioning if a resource shortage is detected. This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console) which is part of the relevant Desktop Pool. The script uses the target machine to determine the connection server address and the Desktop Pool name, and is executed on the machine running ControlUp Console.

This script requires VMWare PowerCLI module to be installed on the machine running the script.
PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'


Version: 3.12.23

Creator: Ton de Vreede

Date Created: 08/22/2019 12:26:28

Date Modified: 05/24/2023 13:43:39

Scripting Language: ps1

