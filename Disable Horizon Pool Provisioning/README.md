# Name: Disable Horizon Pool Provisioning

Description: Disables Horizon View Virtual Desktop pool provisioning
Can be used as a manual or automated action to disable Horizon View Virtual Desktop pool provisioning if a resource shortage is detected. This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console) which is part of the relevant Desktop Pool. The script uses the target machine to determine the connection server address and the Desktop Pool name, and is executed on the machine running ControlUp Console.

This script requires VMWare PowerCLI and the Vmware.Hv.Helper module to be installed on the machine running the script.
PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'
Vmware.Hv.Helper can be installed using the 'Prepare machine for Horizon View scripts' script. It can also be found on Github (see LINK). Download the module and place it in your systemdrive Program Files\WindowsPowerShell\Modules folder 

Version: 1.11.17

Creator: Ton de Vreede

Date Created: 08/22/2019 12:24:32

Date Modified: 02/04/2020 10:23:08

Scripting Language: ps1

