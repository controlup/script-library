# Name: Extend partition to use all available free space

Description: This script uses standard PowerShell commands to expand the disk of Windows machine. For safety, this script only works if the following conditions are met:
    - At least 100Mb of free space available (this needs to be directly AFTER the chosen partition)
    - The Disk State must be Healthy
    - You must specify a drive letter of the disk to be expanded

Version: 1.2.7

Creator: Ton de Vreede

Date Created: 03/17/2020 12:46:05

Date Modified: 03/19/2020 01:03:18

Scripting Language: ps1

