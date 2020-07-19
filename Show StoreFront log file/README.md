# Name: Show StoreFront log file

Description: Pull all log separate Citrix StoreFront log files into a single time sorted csv file. You may need to change the StoreFront logging levels first which can be done with the "Show or Change StoreFront Logging" SBA. Use to help diagnose StoreFront issues.
Arguments:
  Output csv file - full path to a local or remote file to store the log entries in
  Start - optional time to start log export from, e.g. just before the problem started or was reproduced. Can be specified as a date/time or as a number of units of time back from the present such as 7d or 1w where s=second,m=minute,h=hour,d=day,w=week,y=year
  End - optional time to stop log export at. Can be specified either as a date/time or a number of units of time from the start value specified.
If date/time values are used, they must be enclosed in double quotes, e.g. "02/02/2018 08:00:00"

Version: 1.3.6

Creator: Guy Leech

Date Created: 10/23/2018 15:57:48

Date Modified: 11/26/2018 16:39:47

Scripting Language: ps1

