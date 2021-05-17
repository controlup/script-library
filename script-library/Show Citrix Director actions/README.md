# Name: Show Citrix Director actions

Description: Show the Citrix Director actions performed in a given time window, optionally filtered on a specific user. Use this SBA to see what actions have been taken which might affect end users such as logging them off.
Arguments:
  Start - optional time to show actions from. Can be specified as a date/time or as a number of units of time back from the present such as 7d or 1w where s=second,m=minute,h=hour,d=day,w=week,y=year (default is 7 days)
  End - time to stop showing actions after. Can be specified either as a date/time or a number of units of time from the start value specified. (default is the current time)
If date/time values are used, they must be enclosed in double quotes, e.g. "02/02/2018 08:00:00"
  Username - optional name of a user to just show changes for

Version: 1.4.6

Creator: Guy Leech

Date Created: 10/29/2018 12:49:26

Date Modified: 11/26/2018 21:15:31

Scripting Language: ps1

