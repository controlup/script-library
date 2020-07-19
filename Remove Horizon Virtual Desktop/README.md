# Name: Remove Horizon Virtual Desktop

Description: This script deletes a machine from an Horizon desktop pool. If it is a manual pool the machine will only be removed from the pool but not deleted. If it is an automated pool the user can be forcefully logged off (otherwise the script will fail, Horizon 7.7 or newer required) and the machine will be deleted.


This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console). The script uses the target machine to determine the connection server address, and is executed on the machine running ControlUp Console.

This script requires VMWare PowerCLI  to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'

Version: 1.3.4

Creator: Wouter Kursten

Date Created: 02/20/2020 12:27:59

Date Modified: 02/25/2020 20:21:09

Scripting Language: ps1

