# Name: Message and schedule logoff

Description: Send a message to all connected sessions and then log them all off after a specified amount of time. The message can be repeated at a specified interval if required.
For instance, a delay of 15 minutes can be set and the users messaged every 5 minutes with the text specified in the "Message Text" argument.
When the configured script timeout of 60 seconds is reached, an error will be displayed but the script will keep running. If you need to cancel the logoff once the script has timed out, run the "Cancel Logoffs" SBA for the same computer(s)
Arguments:
  Message Text - The text of the message to display to all connected users on the selected computer(s)
  Delay Before Logoff - The period in minutes from when the script is invoked to when all users will be automatically logged off by the SBA
  Message Every - how often, in minutes, the same message is displayed to the remaining users. Specify zero or a blank value to not repeat the message after the initial display

Version: 1.3.13

Creator: Guy Leech

Date Created: 10/12/2018 16:20:49

Date Modified: 11/23/2018 18:48:46

Scripting Language: ps1

