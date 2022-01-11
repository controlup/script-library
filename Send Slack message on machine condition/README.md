# Name: Send Slack message on machine condition

Description: Sends a message to Slack using an Incoming Webhook with an option to include a clickable button with a URI link. Customize the Message input to use this as an Automated Action for alerts on machine metrics.
If you want to include a clickable button with your message the ButtonExplanation, ButtonText and ButtonURI must be provided. If any of these is missing the script will return an error.
Useful for Triggered scripts, fill the title and message with data from the console as required.
This script requires a webhook to be configured in your Slack site. See the link on how to do this: https://slack.com/help/articles/115005265063-Incoming-webhooks-for-Slack

Version: 1.3.25

Creator: Ton.de.Vreede

Date Created: 12/25/2021 14:06:45

Date Modified: 01/04/2022 09:20:35

Scripting Language: ps1

