﻿# Name: Analyze GPO Extensions Load Time

Description: This SBA runs under the session context of a selected user and shows
how long each "Group Policy Client Side Extension" took to complete based on the records inside the "Operational" log under "Microsoft-Windows-GroupPolicy".

By default the log size is configured to 4MB,
That means that this SBA can look back this much.
Consider increasing the log size to view older entries.

Version: 9.21.52

Creator: Niron Koren

Date Created: 02/09/2015 18:04:57

Date Modified: 11/13/2020 12:36:23

Scripting Language: ps1

