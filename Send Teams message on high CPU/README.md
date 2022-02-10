# Name: Send Teams message on high CPU

Description: Example script for an automated action when CPU threshold is exceeded using the 'Send Teams message on machine condition' script.When configured as an automated action triggered from high CPU use a message will be posted in the configured Teams environment, with a button that can be clicked to open the console in the machine location.

Sends a message to Teams using an Incoming Webhook, with an option to include a clickable button with a URI link.Customize the Message input to use this as an Automated Action for alerts on machine metrics.
If you want to include a clickable button with your message, ButtonText and ButtonURI must be provided. If either of these is missing the script will return an error.
Useful for Triggered scripts, fill the Title and message with data from the console as required.
This script requires a webhook to be configured in your Teams site. See the link on how to do this: https://docs.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-incoming-webhook

Version: 2.3.37

Creator: Ton.de.Vreede

Date Created: 12/25/2021 14:06:45

Date Modified: 02/07/2022 15:58:17

Scripting Language: ps1

