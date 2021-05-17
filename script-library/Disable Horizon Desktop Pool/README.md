# Name: Disable Horizon Desktop Pool

Description: Disables a Horizon Desktop pool
Can be used as an automated or manual action to disable a Horizon Desktop pool for planned maintenance. 
This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console) which is part of the relevant Desktop Pool. The script uses the target machine to determine the connection server address and the Desktop Pool name, and is executed on the machine running ControlUp Console.
This script requires VMWare PowerCLI and the Vmware.Hv.Helper module to be installed on the machine running the script.
PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'
Vmware.Hv.Helper can be installed using the 'Prepare machine for Horizon View scripts' script. It can also be found on Github (see LINK). Download the module and place it in your systemdrive Program Files\WindowsPowerShell\Modules folder 

Version: 2.6.8

Creator: Ton de Vreede

Date Created: 02/02/2020 15:03:25

Date Modified: 02/04/2020 09:08:32

Scripting Language: ps1

