# Name: Disconnect or log off idle sessions

Description: Disconnect or logoff sessions on the selected computer(s) which have been idle in excess of a given period, specified in minutes.
Arguments:
  Idle Period - the time in minutes after which a disconnected session will be logged off or disconnected depending on the value for the "Logoff" argument (default is 30 minutes)
  Logoff - if true, sessions idle in excess of the idle period will be logged off, otherwise they will be disconnected (default is false so sessions will be disconnected)

Version: 1.5.7

Creator: Guy Leech

Date Created: 10/13/2018 19:08:29

Date Modified: 11/23/2018 18:20:03

Scripting Language: ps1

