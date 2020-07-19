# Name: Show or change StoreFront logging level

Description: Query or change the logging levels in Citrix StoreFront.
Enable more verbose logging when troubleshooting potential StoreFront issues but turn the level back to off or error when finished to reduce the impact on CPU and free disk space.
Arguments:
  Cluster - ff set to true then all servers in the cluster will be discovered and reported/operated on rather than just the selected computer (default is false)
  Trace Level - specify one of: 'Off', 'Error','Warning','Info','Verbose' or leave empty to query rather than change the current settings (default is empty)

Version: 1.4.14

Creator: Guy Leech

Date Created: 10/17/2018 16:12:53

Date Modified: 11/26/2018 16:13:03

Scripting Language: ps1

