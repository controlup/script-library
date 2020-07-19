# Name: Set VM resource allocation level

Description: This script will change the resource allocation  of a given vSphere VM. By default, the allocation is increased one 'SharesLevel' for CPU, HDD, Memory or all three.
If the SharesLevel for a resource is 'Custom' this will not be changed. ALL hard disks will be set to a new SharesLevel based on the current level of the FIRST disk. Example, script is set to Increase level for All resources:
    CPU SharesLevel 'Normal' ---> CPU ShareLevel 'High'
    Memory SharesLevel 'Custom' ---> Memory SharesLevel 'Custom'
    FIRST HDD SharesLevel 'Low', SECOND HDD SharesLevel 'Normal' ---> ALL HDD SharesLevel 'Normal'

Version: 1.0.2

Creator: Ton de Vreede

Date Created: 03/19/2019 09:30:28

Date Modified: 03/19/2019 09:33:37

Scripting Language: ps1

