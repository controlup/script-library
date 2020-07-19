# Name: Show frequent warning events

Description: Show all warning events from all event logs on the selected computer, ordered on the most frequent, in the specified time period where there is more than one instance of that specific event.
Arguments:
  Minutes Back - the number of minutes back from the current time to search (default is 60)
  Log Level - the event log level to search for events (default is Warning)
  Excluded log names - an optional regular expression, e.g. "Application", which will exclude events from any event logs where the event log name matches the regular expression.

Version: 1.6.23

Creator: Guy Leech

Date Created: 07/06/2018 18:50:25

Date Modified: 11/21/2018 14:09:57

Scripting Language: ps1

