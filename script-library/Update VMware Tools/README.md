# Name: Update VMware Tools

Description: The script uses PowerCLI Update-Tools with the -NoReboot flag to update VMware Client tools. The update command -NoReboot flag should prevent the target guest machine rebooting. However, the VM might still reboot after updating VMware Tools, depending on the currently installed VMware Tools version, the VMware Tools version to which you want to upgrade, and the vCenter Center/ESX versions.
ControlUp Console may disconnect from the target VM's agent while updating the tools due to the NIC drivers updating.

Version: 1.3.9

Creator: Ton de Vreede

Date Created: 01/11/2021 15:48:49

Date Modified: 01/18/2021 11:29:00

Scripting Language: ps1

