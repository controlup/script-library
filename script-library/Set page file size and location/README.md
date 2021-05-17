# Name: Set page file size and location

Description: Set the page file to a specific drive or full path and its initial and maximum sizes. If the sizes are not specified, or specified as zero, then the page file sizes will be system managed. If the page file location is empty or set to "auto" then the page file will be automatically managed.
A reboot will be required for the changes to take effect.
Arguments:
  Page File Location - Either a drive, drive and folder or empty or set to "auto" where the latter two will result in the page file being automatically managed
  Initial Size - The initial size of the page file in MB. If zero then sizes will be system managed.
  Maximum Size - The maximum size of the page file in MB. If zero then sizes will be system managed.

Version: 1.3.9

Creator: Guy Leech

Date Created: 10/14/2018 22:05:04

Date Modified: 11/26/2018 11:47:35

Scripting Language: ps1

