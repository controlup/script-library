# Name: Show frequent error events

Description: Show all error events from all event logs on the selected computer, ordered on the most frequent, in the specified time period where there is more than one instance of that specific event.
Arguments:
  Minutes Back - the number of minutes back from the current time to search (default is 60)
  Log Level - the event log level to search for events (default is Error)
  Excluded log names - an optional regular expression, e.g. "Application", which will exclude events from any event logs where the event log name matches the regular expression.

Version: 1.4.8

Creator: Guy Leech

Date Created: 07/09/2018 23:00:04

Date Modified: 11/21/2018 14:12:37

Scripting Language: ps1

