# Name: Get PVS Write Cache Size

Description: A Script to obtain the Pool Non Paged Bytes which ultimately is an indication of the amount of Cache in Ram used by Citrix Provisioning Services. 
This script will output the amount of PVS Ram Cache in use. This is done by enumarating the amount of Pool nonPaged Bytes in use, which is a measure of the RamCache. Action should be taken if this figure is close to the amount allocated within PVS. The script also supports examining the PVS configuration of disk cache only.

Version: 2.12.18

Creator: Matthew Nichols

Date Created: 03/29/2015 14:22:27

Date Modified: 04/14/2015 12:01:55

Scripting Language: ps1

