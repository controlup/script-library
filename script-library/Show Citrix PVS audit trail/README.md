# Name: Show Citrix PVS audit trail

Description: Retrieve Citrix Provisioning Services audit events, if auditing has been enabled although it can be enabled by this SBA if required. Use to see what administrative actions were performed in an optional time window, by whom and what was done. It must be run with the credentials of a user who has been granted PVS access.
Arguments:
  Enable Auditing - if auditing is not enabled, specifying true for this parameter will enable auditing if the user running the SBA has sufficient privilege.
  Start - optional time to show audit events from. Can be specified as a date/time or as a number of units of time back from the present such as 7d or 1w where s=second,m=minute,h=hour,d=day,w=week,y=year
  End - optional time to stop showing audit events after. Can be specified either as a date/time or a number of units of time from the start value specified.
If date/time values are used, they must be enclosed in double quotes, e.g. "02/02/2018 08:00:00"

Version: 1.3.6

Creator: Guy Leech

Date Created: 10/23/2018 18:10:13

Date Modified: 11/26/2018 16:52:31

Scripting Language: ps1

