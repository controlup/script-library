# Name: Get AD Domain Controller Status

Description: List the synchronization status and replication errors of all domain controllers in the domain. 

This script can be executed on a monitor and will request the required data via a PSSession to the domain controller(s). 

This script requires the ActiveDirectory PowerShell module to function. 

If errors are found, this could be a long running script. Increase the timeout if required.

Version: 2.0.0

Creator: Rein Leen

Date Created: 04/11/2023 10:42:42

Date Modified: 05/23/2023 12:34:59

Scripting Language: ps1

