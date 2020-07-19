# Name: Show Citrix Studio changes

Description: Show the configuration changes made in Citrix Studio in a given time window, optionally filtered on a specific user. Use this SBA to see if any changes have been made which might be affecting end users.
Arguments:
  Start - optional time to show changes  from. Can be specified as a date/time or as a number of units of time back from the present such as 7d or 1w where s=second,m=minute,h=hour,d=day,w=week,y=year (default is 7 days)
  End - time to stop showing changes after. Can be specified either as a date/time or a number of units of time from the start value specified. (default is the current time)
If date/time values are used, they must be enclosed in double quotes, e.g. "02/02/2018 08:00:00"
  Username - optional name of a user to just show changes for

Version: 1.4.14

Creator: Guy Leech

Date Created: 10/29/2018 12:28:11

Date Modified: 11/26/2018 21:16:02

Scripting Language: ps1

