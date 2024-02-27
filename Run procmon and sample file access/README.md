# Name: Run procmon and sample file access

Description: Run the Sysinternals Process Monitor (procmon) utility for a specified amount of time for a selected process and see which files are most frequently accessed. If a path to an existing procmon executable is not given, it will be downloaded securely from the live.sysinternals.com site.
Arguments:
  Monitor Period - the time in seconds to run the monitoring for. Monitoring for more than 60 seconds is not recommended as this can potentially impact system performance and disk space.
  Backing file - if not specified this will be in the \windows\temp folder on the system drive which on Citrix PVS booted systems can cause performance issues so specifying a file on a persistent local drive can help alleviate this potential issue.
  Procmon Location - the location of an existing copy of procmon.exe. If not specified and internet connectivity is available, it will be downloaded. 

Version: 3.10.47

Creator: Guy Leech

Date Created: 10/01/2018 18:43:19

Date Modified: 01/26/2024 14:09:57

Scripting Language: ps1

