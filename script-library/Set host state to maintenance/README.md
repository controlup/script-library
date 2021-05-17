# Name: Set host state to maintenance

Description: Placing a host in Maintenance will migrate the powerred on VMs to other hosts in the cluster. If the Evacuate switch is passed all offline machines a migrated too.
This script will only place a host in Maintenance if the cluster it is part of is DRSFullyAutomated.
If the host uses a VSAN and the PowerCLI version is high enough the default VsanEvacuationMode setting will be used. This is only supported with PowerCLI 6 or higher.


Version: 1.4.10

Creator: Ton de Vreede

Date Created: 02/21/2019 15:41:54

Date Modified: 03/31/2019 17:21:16

Scripting Language: ps1

