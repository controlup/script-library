# Name: Send Slack message on Local Admin logon

Description: This is an example of how the 'Send Slack message on Session condition' script can be used in a Trigger follow up action.
Configure the defaults for your Slack environment and set the script as a follow up action on User Logon on a machine.
NOTE: In Settings the following changes have been made (compared to the 'Send Slack message' script)
	Action Assigned to: Session
	Execution Context: User Session
	Security Context: Default (Session's User)

Sends a message to Slack using an Incoming Webhook with an option to include a clickable button with a URI link. Customize the Message input to use this as an Automated Action for alerts on session metrics.
If you want to include a clickable button with your message the ButtonExplanation, ButtonText and ButtonURI must be provided. If any of these is missing the script will return an error.
Useful for Triggered scripts, fill the title and message with data from the console as required.
This script requires a webhook to be configured in your Slack site. See the link on how to do this: https://slack.com/help/articles/115005265063-Incoming-webhooks-for-Slack

Version: 2.15.44

Creator: Ton.de.Vreede

Date Created: 12/25/2021 14:06:45

Date Modified: 02/08/2022 13:29:18

Scripting Language: ps1

