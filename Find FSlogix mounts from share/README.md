# Name: Find FSlogix mounts from share

Description: Searches the given list of shares, or pulls them from the registry if * specified and the script is running on a machine with FSlogix installed, for .metadata files and extracts the name of the machine where the corresponding vhd/vhdx file is mounted along with other useful data about the disk.

The last boot time of the machine where the disk is mounted can be retrieved where if this is empty/missing it likely means that the user running the script does not have WMI/CIM permissions or that machine is not powered on. 

Both profile and Office disks will be reported on.

Version: 2.0.12

Creator: Guy Leech

Date Created: 10/06/2022 10:48:22

Date Modified: 12/14/2022 15:18:25

Scripting Language: ps1

