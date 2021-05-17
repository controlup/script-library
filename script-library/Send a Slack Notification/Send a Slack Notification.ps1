<#

    .SYNOPSIS
	This script will send a User session logon notification to a slack channel using Slack's Incoming Webhooks app

    .DESCRIPTION
	This script will use the Slack Incoming Webhooks API to send a message to a pre-defined Slack channel
	As a prerequisite you need to add and set up the Incoming Webhooks app in your Slack workspace and retrieve the Webhook URL (See NOTES below)
	
	This script can be modified for other scenarios by modifying the following sections below:
	- The CmdLetBinding section to match your Arguments
	- The Slack message text 

    .NOTES
	Instructions for setting up the Incoming Webhooks app on Slack.
		Follow these steps:
		Step 1: Add the app to your workspace by going to https://my.slack.com/apps and search for Incoming Webhooks
		Step 2: After selecting the Incoming Webhooks app click on "Add to Slack" on the left side of the page
		Step 3: Follow the on-screen instructions to configure the Incoming Webhooks app

		For more information on how to add and configure apps on Slack, please read: https://slack.com/help/articles/202035138-Add-an-app-to-your-workspace
		For more details on the Incoming Webhooks app, please read: https://api.slack.com/messaging/webhooks

	Addtional notes:
		I am using a custom "ControlUp icon" in the $body section below. If you want to use a different emoji please change to your preference
		More details: https://slack.com/slack-tips/upload-custom-slack-emoji-to-express-your-unique-office-culture

    .CONTEXT
	Session

    .MODIFICATION_HISTORY
	Created: 2020-03-04

    .AUTHOR
	Joel Stocker
#>

# This section sets the variables that are being used for both the API call body (next section) and the Rest API call URI (bottom section)
# It uses values from the "Arguments" tab in the Scripts Action interface in subsequent order (i.e. 1st line corresponds with $args[0], 2nd is $args[1], etc.)
# You can replace/add to match your needs. In this sample script we use the Full Name of a session user (1st line) and the Machine Name (2nd line)
# The third line set your Slack's Incoming Webhook URL. The script will prompt for your Incoming Webhook URL when running a manual Script Action
# Make sure that you change the default value for the Incoming Webhook URL argument to reflect your URL if you want to use this script in an automated Trigger
[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='User Full Name')][ValidateNotNullOrEmpty()] [string]$username,
    [Parameter(Mandatory=$true,HelpMessage='Machine Name')][ValidateNotNullOrEmpty()] [string]$machinename,
    [Parameter(Mandatory=$true,HelpMessage='Slack Webhook URL')][ValidateNotNullOrEmpty()] [string]$SlackIncomingWebhookUri
)

# This section constructs the $body value that will be send to Slack's webhook API using the variables set in the previous section
$body = @"
    {
        "username": "ControlUp Notification",
        "text": "User *$username* has logged on to machine *$machinename*",
        "icon_emoji":":controlup:"
    }
"@

# This section will send the API call using Powershell to Slack and Slack will process the request and send the notification
Invoke-RestMethod -uri $SlackIncomingWebhookUri -Method Post -body $body -ContentType 'application/json'
