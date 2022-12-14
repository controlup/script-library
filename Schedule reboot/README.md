# Name: Schedule reboot

Description: Schedule a reboot for a number of minutes/hours in the future by creating a scheduled task which runs once only. By default the reboot will not occur if there are logged on users or disconnected sessions in existence at the scheduled reboot time. This script will not message users, log them off or prevent further logons from occurring.
Arguments:
  Minutes:hours - the number of hours/minutes in the future at which to schedule the reboot, e.g. 2:30 for 2 hours and 30 minutes in the future (default is 30 minutes)
  Force - if set to true then the reboot will occur even if there are connected or disconnected user sessions at the time of the reboot (default is false)
  Reason for Reboot - optional text which be placed in the event log.

Version: 3.5.11

Creator: Guy Leech

Date Created: 09/29/2018 11:04:54

Date Modified: 12/14/2022 15:17:00

Scripting Language: ps1

