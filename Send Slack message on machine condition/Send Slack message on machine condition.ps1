#requires -Version 3.0
<#
	.SYNOPSIS
	Send a message to Slack

	.DESCRIPTION
	Sends a message to Slack using an Incoming Webhook with an option to include a clickable button with a URI link.
	Useful for Triggered scripts, fill the Title and message with data from the console as required.

	.EXAMPLE
	& '.\Send Slack message with optional button.ps1' -WebhookUri 'https://myslack/MyWebhook' -UserName 'ControlUp Automation' -Proxy DoNotUse -Title 'My title' -Message 'My message'
	This will post a message in Slack in the location configured in the webhook with a title and a message.

	& '.\Send Slack message with optional button.ps1' -WebhookUri 'https://myslack/MyWebhook' -UserName 'ControlUp Automation' -Title 'My title' -Message 'My message' -Proxy 'DoNotUse' -ButtonExplanation 'This explains the button' -ButtonText 'Push me!' -ButtonURI 'http://www.controlup.com'
	This will post a message in Slack in the location configured in the webhook, with a title, message, explanation for the button use, button text and the actual button URI.

	& '.\Send Slack message with optional button.ps1' -WebhookUri 'https://myslack/MyWebhook' -UserName 'ControlUp Automation' -Title 'My title' -Message 'My message' -Proxy 'best.proxy.ever' -ButtonExplanation 'This explains the Button' -ButtonText 'Push me!' -ButtonURI 'http://www.controlup.com' -Proxy best.proxy.ever
	Similar, except using a proxy server

	.PARAMETER WebhookUri
	Slack webhook URI
	This parameter must be provided.

	.PARAMETER UserName
	Enter the author of your Slack message
	This parameter must be provided.

	.PARAMETER Title
	Enter a title for your Slack message
	This parameter must be provided.

	.PARAMETER Message
	Enter the main text of the Slack message
	This parameter must be provided.

	.PARAMETER Proxy
	If you are using a proxy enter the FQDN or IP number of the proxy server. IF YOU ARE NOT USING A PROXY SERVER, SET THIS TO 'DoNotUse'!
	This parameter is mandatory.

	.PARAMETER ButtonExplanation
	Enter the body of the second part of the Slack message for the optional button
	This parameter is optional, but if provided so must ButtonText and ButtonURI.

	.PARAMETER ButtonText
	Enter the text for the optional button in your Slack message
	This parameter is optional, but if provided so must ButtonExplanation and ButtonURI.

	.PARAMETER ButtonURI
	Enter the link for the optional button in your Slack message
	This parameter is optional, but if provided so must ButtonExplanation and ButtonText.

	.LINK
	https://slack.com/help/articles/115005265063-Incoming-webhooks-for-Slack

	.NOTES
	This script requires a webhook to be configured in your Slack site. See the provided link on how to do this.
	If you want to include a clickable button with your message the ButtonExplanation, ButtonText and ButtonURI must be provided. If any of these is missing the script will return an error.
	The ButtonURI can only be a valid URI. Invalid URIs are rejected.
	Valid URI : 'http://www.controlup.com' or 'controlup://MyOrganization/Machines'
	Invalid URI : 'www.controlup.com', as it does not contain either http:// or https:// to indicate this is a website.

	Author:		 Samuel Legrand
	Version:		1.2
	Ton de Vreede - Refactor, added error handling, converted string JSON to hash table, comment block additions.
#>
[CmdletBinding()]
Param(
	[Parameter(Mandatory = $true, HelpMessage = 'Slack webhook URI')]
	[ValidateNotNullOrEmpty()]
	[string]$WebhookUri,
	[Parameter(Mandatory = $true, HelpMessage = 'Enter the author of your Slack message')]
	[ValidateNotNullOrEmpty()]
	[string]$UserName,
	[Parameter(Mandatory = $true, HelpMessage = 'Enter a title for your Slack message')]
	[ValidateNotNullOrEmpty()]
	[string]$Title,
	[Parameter(Mandatory = $true, HelpMessage = 'Enter the main text of the Slack message')]
	[ValidateNotNullOrEmpty()]
	[string]$Message,
	[Parameter(Mandatory = $true, HelpMessage = 'If you are using a proxy enter the FQDN or IP number of the proxy server, otherwise leave at DoNotUse')]
	[ValidateNotNullOrEmpty()]
	[string]$Proxy,
	[Parameter(Mandatory = $false, HelpMessage = '(Optional) Enter the body of the second part of the Slack message, to explain what the button does' )]
	[string]$ButtonExplanation,
	[Parameter(Mandatory = $false, HelpMessage = '(Optional) Enter the text on the button in your Slack message')]
	[string]$ButtonText,
	[Parameter(Mandatory = $false, HelpMessage = '(Optional) Enter the URI for the button in your Slack message')]
	[string]$ButtonURI
)

$ErrorActionPreference = 'Stop'

# Set security protocol in this session to Transport Layer Security 1.2 in case TLS 1.0 is set on the machine.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Create hashtables for body and REST parameters
[hashtable]$hshBody = @{
	'username'   = $UserName
	'icon_emoji' = ':robot_face:'
	'blocks'     = @(
		@{
			'type' = 'section'
			'text' = @{
				'type' = 'mrkdwn'
				'text' = $Title
			}
		},
		@{
			'type' = 'divider'
		},
		@{
			'type' = 'section'
			'text' = @{
				'type' = 'mrkdwn'
				'text' = $Message
			}
		}
	)
}

# If more than 5 arguments were passed, the button should be used. But if fewer than 8 were passed there is not enough information to add the button.
If ($PSBoundParameters.Count -in 6..7) {
	Throw "One of the arguments for adding a button was used, but to add a button three values must be entered. Please check you have entered a value for the button explanation text, the text on the button and the button URI. If it was not your intention to add a button, please ensure all three of these arguments are left empty."
}
ElseIf ($PSBoundParameters.Count -eq 8) {
	# Test if a valid URI was passed for the Button
	If (!($ButtonURI -as [System.URI]).IsAbsoluteUri) {
		Throw "The Button $ButtonURI is not a valid URI. Please check the syntax and try again."
	}

	# Add the 'button section'
	$hshBody.blocks += (
		@{
			'type'      = 'section'
			'text'      = @{
				'type' = 'mrkdwn'
				'text' = $ButtonExplanation
			}
			'accessory' = @{
				'type' = 'button'
				'text' = @{
					'type'  = 'plain_text'
					'text'  = $ButtonText
					'emoji' = $true
				}
				'url'  =	$ButtonURI
			}
		}
	)
}

# Create parameters, convert Body to JSON already to avoid possible Depth problems
[hashtable]$hshParameters = @{
	'Uri'         = $WebhookUri
	'Method'      = 'POST'
	'Body'        = $hshBody | ConvertTo-Json -Depth 5
	'ContentType' = 'application/json'
}

# Add the proxy if passed
If (!$Proxy -eq 'DoNotUse') {
	$hshParameters.Add('Proxy', $Proxy)
}

Write-Verbose -Message $hshParameters

# Send the message to Slack
try {
	$return = Invoke-RestMethod @hshParameters
}
catch {
	Throw "Failed to send message $Title to Slack. Exception:`n$_"
}

# Test the result
If ($return -eq 'ok') {
	Write-Output -InputObject "Slack message sent."
	Exit 0
}
Else {
	Throw "There was an issue sending the Slack message. The expected return from the REST call is 'ok', but $return was received."
}
