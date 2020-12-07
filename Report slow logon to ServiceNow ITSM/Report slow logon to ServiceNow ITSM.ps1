<#
    .SYNOPSIS
	This Script Action uses the ServiceNow API to create an incident in ServiceNow IT Service Management. 
	    
    .DESCRIPTION
	The script as presented is for a specific use case (it creates an incident containing the user's full name, the user's logon duration and the machine name).
	Tthis Script Action can also serve as an example/template for your own ServiceNow integration needs.
		
	This script can be modified for other scenarios by modifying the following sections below:
	- The CmdLetBinding section to match your Arguments
	- The ServiceNow API call body 
    
    .NOTES
	More details on how this script works can be found here https://www.controlup.com/10-simple-steps-to-build-your-own-integration-in-controlup/
	For more information on how to use the ServiceNow API works, please read https://docs.servicenow.com/bundle/madrid-application-development/page/integrate/inbound-rest/concept/c_GettingStartedWithREST.html
				
    .CONTEXT
	Session
    
    .MODIFICATION_HISTORY
	Created: 2020-10-21
    
    .AUTHOR
	Joel Stocker https://twitter.com/joelinthecloud
#>

# This section sets the variables that are being used for both the API call body (next section) and the Rest API call URI (bottom section)
# It uses values from the "Arguments" tab in the Scripts Action interface in subsequent order (i.e. 1st line corresponds with $args[0], 2nd is $args[1], etc.)
# You can replace/add to match your needs. In this sample script we use the Full Name of the user (1st line), the Machine Name (2nd line) and the Logon Duration (3rd line)
# Lines 4-6 are the details we need to authenticate to the ServiceNow API. The script will prompt for this when running a manual Script Action
# Make sure that you change the default values for the the 3 ServiceNow arguments to reflect your information if you want to use this script in an Automated Action (Trigger)
[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='User Full Name')][ValidateNotNullOrEmpty()] [string]$caller,
    [Parameter(Mandatory=$true,HelpMessage='Logon Duration')][ValidateNotNullOrEmpty()] [string]$logonduration,
    [Parameter(Mandatory=$true,HelpMessage='Machine Name')][ValidateNotNullOrEmpty()] [string]$machinename,
    [Parameter(Mandatory=$true,HelpMessage='ServiceNow Username')][ValidateNotNullOrEmpty()] [string]$username,
    [Parameter(Mandatory=$true,HelpMessage='ServiceNow Password')][ValidateNotNullOrEmpty()] [string]$password,
    [Parameter(Mandatory=$true,HelpMessage='ServiceNow Instance')][ValidateNotNullOrEmpty()] [string]$instanceid
)

# Put API Call URL in parameter consuming $instanceid from above
$apicallurl = "https://$instanceid.service-now.com/api/now/table/incident"

$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $username, $password)))

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/json")
$headers.Add("Authorization", "Basic $base64AuthInfo")

# This section contains the body for the ServiceNow API call. 
# The first part of each line is the ServiceNow Field parameter. The second part of each line is the value for the corresponsing Field parameter.
# Update this section to fit your needs and to match your ServiceNow instance fields (as well in case of adding additional metrics/information send to ServiceNow)
# For full details on the ServiceNow REST API please read https://docs.servicenow.com/bundle/madrid-application-development/page/integrate/inbound-rest/concept/c_GettingStartedWithREST.html
$body = "{
`n `"short_description`": `"Slow User Logon Incident`",
`n `"caller_id`": '$caller',
`n `"description`": `"User $caller's Logon Duration was $logonduration seconds on machine $machinename`"
`n}"

$response = Invoke-RestMethod $apicallurl -Method 'POST' -Headers $headers -Body $body
$IncidentNumber = $response.result.number
Write-Host "ServiceNow Incident $IncidentNumber Created"
