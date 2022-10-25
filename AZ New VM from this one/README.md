# Name: AZ New VM from this one

Description: Create a number of Azure VMs using the selected machine as a template in terms of machine, disk sizes/types and network interfaces.
If the machine name contains # characters, these will be replaced, with leading zeroes, by the next available machine name
It does not clone or copy any disks other than any gallery image the original was built from.
If domain details are provided, it will attempt to join the given domain when creation is complete.
Multiple tags can be added in a comma separated list of the form Tag name=Tag text

Version: 1.4.23

Creator: Guy Leech

Date Created: 06/27/2022 15:46:24

Date Modified: 10/07/2022 11:46:07

Scripting Language: ps1

