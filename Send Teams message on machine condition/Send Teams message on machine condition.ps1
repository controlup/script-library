#requires -Version 3.0
<#
	.SYNOPSIS
	Send a message to Teams with a clickable Action button

	.DESCRIPTION
	Sends a message to Teams using an Incoming Webhook, that includes a clickable Action button.
	Useful for Triggered scripts, fill the Title and message with data from the console as required.

	.EXAMPLE
	& '.\Send Teams message with optional button.ps1' -WebhookUri 'https://myTeams/MyWebhook' -Proxy DoNotUse -Title 'My title' -Message 'My message'
	This will post a message in Teams in the location configured in the webhook with a title and a message.

	& '.\Send Teams message with optional button.ps1' -WebhookUri 'https://myTeams/MyWebhook' -Title 'My title' -Message 'My message' -Proxy 'DoNotUse' -ButtonText 'Push me!' -ButtonURI 'http://www.controlup.com'
	This will post a message in Teams in the location configured in the webhook, with a title, message, explanation for the button use, button text and the actual button URI.

	& '.\Send Teams message with optional button.ps1' -WebhookUri 'https://myTeams/MyWebhook' -Title 'My title' -Message 'My message' -Proxy 'best.proxy.ever' -ButtonText 'Push me!' -ButtonURI 'http://www.controlup.com' -Proxy best.proxy.ever
	Similar, except using a proxy server

	.PARAMETER WebhookUri
	Teams webhook URI
	This parameter must be provided.

	.PARAMETER Title
	Enter a title for your Teams message
	This parameter must be provided.

	.PARAMETER Message
	Enter the message content for your Teams message
	This parameter must be provided.

	.PARAMETER Proxy
	If you are using a proxy enter the FQDN or IP number of the proxy server. IF YOU ARE NOT USING A PROXY SERVER, SET THIS TO 'DoNotUse'!
	This parameter is mandatory.

	.PARAMETER ButtonText
	Enter the text for the optional button in your Teams message
	This parameter is optional, but if provided so must ButtonURI.

	.PARAMETER ButtonURI
	Enter the link for the optional button in your Teams message
	This parameter is optional, but if provided so must ButtonText.

	.LINK
	https://docs.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-incoming-webhook

	.NOTES
	This script requires a webhook to be configured in your Teams site. See the provided link on how to do this.
	If you want to include a clickable button with your message, ButtonText and ButtonURI must be provided. If either of these is missing the script will return an error.
	The ButtonURI can only be a valid URI. Invalid URIs are rejected.
	Valid URI : 'http://www.controlup.com' or 'controlup://MyOrganization/Machines'
	Invalid URI : 'www.controlup.com', as it does not contain either http:// or https:// to indicate this is a website.

	Author:         Samuel Legrand
	Version:        1.2
	Ton de Vreede - Refactor, added error handling, converted string JSON to hashtable, comment block additions.
#>

[CmdletBinding()]
Param(
	[Parameter(Mandatory = $true, HelpMessage = 'Teams webhook URI')]
	[ValidateNotNullOrEmpty()]
	[string]$WebhookUri,
	[Parameter(Mandatory = $true, HelpMessage = 'Enter a title for your Teams message')]
	[ValidateNotNullOrEmpty()]
	[string]$Title,
	[Parameter(Mandatory = $true, HelpMessage = 'Enter the message content for your Teams message')]
	[ValidateNotNullOrEmpty()]
	[string]$Message,
	[Parameter(Mandatory = $true, HelpMessage = 'If you are using a proxy enter the FQDN or IP number of the proxy server')]
	[ValidateNotNullOrEmpty()]
	[string]$Proxy,
	[Parameter(Mandatory = $false, HelpMessage = '(Optional) Enter the text on the button in your Teams message')]
	[string]$ButtonText,
	[Parameter(Mandatory = $false, HelpMessage = '(Optional) Enter the URI for the button in your Teams message')]
	[string]$ButtonURI
)

$ErrorActionPreference = 'Stop'

# Set security protocol in this session to Transport Layer Security 1.2 in case TLS 1.0 is set on the machine.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Create hashtables for body and REST parameters
[hashtable]$hshBody = @{
	'@type'      = 'MessageCard'
	'@context'   = 'http: //schema.org/extensions'
	'themeColor' = 'FF0000'
	'summary'    = 'Summary'
	'sections'   = @(
		@{
			'activityTitle'    = $Title
			'activitySubtitle' = $Message
			'markdown'         = $true
		}
	)
}

# If more than 4 arguments were passed, the button should be used. But if fewer than 6 were passed there is not enough information to add the button.
If ($PSBoundParameters.Count -eq 5) {
	Throw "One of the arguments for adding a button was used, but to add a button two values must be entered. Please check you have entered a value for the text on the button and the button URI. If it was not your intention to add a button, please ensure both of these arguments are left empty."
}
ElseIf ($PSBoundParameters.Count -eq 6) {
	# Test if a valid URI was passed for the Button
	If (!($ButtonURI -as [System.URI]).IsAbsoluteUri) {
		Throw "The Button $ButtonURI is not a valid URI. Please check the syntax and try again."
	}

	# Add the 'button section'
	$hshBody.Add('potentialAction', @(
			@{
				'@type'   = 'OpenUri'
				'name'    = $ButtonText
				'targets' = @(
					@{
						'os'  = 'default'
						'uri' = $ButtonURI
					}
				)
			}
		)
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

# Send the message to Teams
try {
	$return = Invoke-RestMethod @hshParameters
}
catch {
	Throw "Failed to send message $Title to Teams. Exception:`n$_"
}

# Test the result
If ($return -eq 1) {
	Write-Output -InputObject "Teams message sent."
	Exit 0
}
Else {
	Throw "There was an issue sending the Teams message. The message may have been sent but the expected return 1 was not received. The return was $return."
}
