# Name: Check/Fix DNS Permissions

Description: Check the permissions of the records that match the computer name specified and optionally change them so that the computer account for the current AD object for that computer has rights to update the record.
Run on a computer that has the ActiveDirectory PowerShell mdoule installed like a domain controller. Permissions for Domain controller DNS records will not be changed,
Use when machines have been rebuilt and they cannot update their own DNS record because it is owned by the account for the previous build which no longer exists in AD so the current computer does not have permissions when it wants to register the IP address against its own DNS record
NOTE: It is recommended you make a backup of your DNS with the Backup DNS Zone script before running this script with Fix set to Yes.

Version: 1.1.7

Creator: Guy Leech

Date Created: 10/03/2019 20:11:04

Date Modified: 10/06/2022 14:56:05

Scripting Language: ps1

