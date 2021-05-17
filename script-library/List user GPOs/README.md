# Name: List user GPOs

Description: This SBA runs under the session context of a selected user and shows, 
every "User Group Policy" applied based on the records inside the "Operational" log under "Microsoft-Windows-GroupPolicy".

By default the log size is configured to 4MB,
That means that this SBA can look back this much.
Consider increasing the log size to view older entries.

Version: 2.2.7

Creator: Niron Koren

Date Created: 07/07/2014 12:21:32

Date Modified: 02/12/2015 09:09:59

Scripting Language: ps1

