# Name: Change local group membership

Description: Add or remove domain or local accounts to/from local groups on selected computers. Can either be done immediately or at a given date/time in the future via a scheduled task, e.g. remove specific users from the local admininstrators group in 1 day's time.
Arguments:
  Users - a comma separated list of AD user accounts to add/remove to/from the specified group
  Local group - the name of the local group which will have the users added or removed
  Remove from group - if true then the specified users will be removed from the group, if false then the users will be added to the group (default is false)
  When - If nothing is specified, the action is taken immediately otherwise a scheduled task is created to perform the action at the data/time specified which can also be a number followed by a time unit, e.g. 8h for 8 hours or 1d for 1 day. If specifying a date/time, it must be enclosed in double quotes.

Version: 1.4.12

Creator: Guy Leech

Date Created: 10/22/2018 18:44:28

Date Modified: 11/26/2018 20:53:26

Scripting Language: ps1

