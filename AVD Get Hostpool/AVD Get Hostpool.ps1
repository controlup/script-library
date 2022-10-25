#requires -Version 4.0
<#
.SYNOPSIS
    AVD Get Hostpool
.DESCRIPTION
    Gets the details of the AVD Hostpool the target machine/session is in.
.CONTEXT
    Azure Virtual Desktops
.MODIFICATION_HISTORY
    Esther Barthel, MSc - 22/03/20 - Original code
    Ton de Vreede - 01/06/2022
		- Removed PowerShell module dependency (changed to REST)
		- Using new metrics available from the Console
		- Complete refactor
.LINK
https://docs.microsoft.com/en-us/rest/api/desktopvirtualization/
https://docs.microsoft.com/en-us/azure/templates/microsoft.desktopvirtualization/applicationgroups?tabs=json
.COMPONENT
Set-AzSPCredentials - The required Azure Service Principal (Subscription level) and tenantID information need to be securely stored in a Credentials File. The AZ Store Azure Credentials Script Action will ensure the file is created according to ControlUp standards
.NOTES
Azure functions created by Guy Leech and Esther Barthel
#>

[CmdletBinding()]
Param
(
	[Parameter(Mandatory = $true, HelpMessage = 'SBA parameter auto entry: Session Host NetBIOS Name')]
	[string]$MachineName,
	[Parameter(Mandatory = $true, HelpMessage = 'Azure Tenant ID')]
	[string]$AzTenantId,
	[Parameter(Mandatory = $true, HelpMessage = 'Azure Subscription')]
	[string]$AzSubscription,
	[Parameter(Mandatory = $true, HelpMessage = 'Azure Resource Group')]
	[string]$AzResourceGroup,
	[Parameter(Mandatory = $true, HelpMessage = 'Azure VM Id')]
	[string]$AzVmId,
	[Parameter(Mandatory = $true, HelpMessage = 'Detail level: summary, normal or high.')]
	[string]$OutputDetail
)

# Set up some defaults
$ErrorActionPreference = 'Stop'
# Configure a larger output width for the ControlUp PowerShell console
[int]$outputWidth = 400
# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

# Set array with properties for default output
[array]$arrOutputDetails = @(
	@{N = 'id'; E = { $_.id } }
	@{N = 'kind'; E = { $_.kind } }
	@{N = 'location'; E = { $_.location } }
	@{N = 'name'; E = { $_.name } }
	@{N = 'tags'; E = { $_.tags } }
	@{N = 'type'; E = { $_.type } }
	@{N = 'properties - cloudPcResource'; E = { $_.properties.cloudPcResource } }
	@{N = 'properties - description'; E = { $_.properties.description } }
	@{N = 'properties - friendlyName'; E = { $_.properties.friendlyName } }
	@{N = 'properties - hostPoolType'; E = { $_.properties.hostPoolType } }
	@{N = 'properties - loadBalancerType'; E = { $_.properties.loadBalancerType } }
	@{N = 'properties - maxSessionLimit'; E = { $_.properties.maxSessionLimit } }
	@{N = 'properties - validationEnvironment'; E = { $_.properties.validationEnvironment } }

)

[array]$arrDetailsNormal = @(
	@{N = 'systemData - createdAt'; E = { $_.systemData.createdAt } }
	@{N = 'systemData - createdBy'; E = { $_.systemData.createdBy } }
	@{N = 'systemData - createdByType'; E = { $_.systemData.createdByType } }
	@{N = 'systemData - lastModifiedAt'; E = { $_.systemData.lastModifiedAt } }
	@{N = 'systemData - lastModifiedBy'; E = { $_.systemData.lastModifiedBy } }
	@{N = 'systemData - lastModifiedByType'; E = { $_.systemData.lastModifiedByType } }
	@{N = 'properties - vmTemplate - namePrefix'; E = { $_.properties.vmTemplate.namePrefix } }
	@{N = 'properties - agentUpdate - type'; E = { $_.properties.agentUpdate.type } }
	@{N = 'properties - agentUpdate - maintenanceWindows - dayOfWeek'; E = { $_.properties.agentUpdate.maintenanceWindows.dayOfWeek } }
	@{N = 'properties - agentUpdate - maintenanceWindows - hour'; E = { $_.properties.agentUpdate.maintenanceWindows.hour } }
	@{N = 'properties - vmTemplate - domain'; E = { $_.properties.vmTemplate.domain } }
	@{N = 'properties - vmTemplate - imageType'; E = { $_.properties.vmTemplate.imageType } }
	@{N = 'properties - vmTemplate - imageUri'; E = { $_.properties.vmTemplate.imageUri } }
	@{N = 'properties - vmTemplate - vmSize'; E = { $_.properties.vmTemplate.vmSize } }
	@{N = 'properties - startVMOnConnect'; E = { $_.properties.startVMOnConnect } }
)

[array]$arrDetailsHigh = @(
	@{N = 'properties - agentUpdate - useSessionHostLocalTime'; E = { $_.properties.agentUpdate.useSessionHostLocalTime } }
	@{N = 'properties - agentUpdate - maintenanceWindowTimeZone'; E = { $_.properties.agentUpdate.maintenanceWindowTimeZone } }
	@{N = 'properties - customRdpProperty - audiomode'; E = { $_.properties.customRdpProperty.audiomode } }
	@{N = 'properties - customRdpProperty - autoreconnection_enabled'; E = { $_.properties.customRdpProperty.autoreconnection_enabled } }
	@{N = 'properties - customRdpProperty - bandwidthautodetect'; E = { $_.properties.customRdpProperty.bandwidthautodetect } }
	@{N = 'properties - customRdpProperty - devicestoredirect'; E = { $_.properties.customRdpProperty.devicestoredirect } }
	@{N = 'properties - customRdpProperty - drivestoredirect'; E = { $_.properties.customRdpProperty.drivestoredirect } }
	@{N = 'properties - customRdpProperty - enablecredsspsupport'; E = { $_.properties.customRdpProperty.enablecredsspsupport } }
	@{N = 'properties - customRdpProperty - networkautodetect'; E = { $_.properties.customRdpProperty.networkautodetect } }
	@{N = 'properties - customRdpProperty - redirectclipboard'; E = { $_.properties.customRdpProperty.redirectclipboard } }
	@{N = 'properties - customRdpProperty - redirectcomports'; E = { $_.properties.customRdpProperty.redirectcomports } }
	@{N = 'properties - customRdpProperty - redirectlocation'; E = { $_.properties.customRdpProperty.redirectlocation } }
	@{N = 'properties - customRdpProperty - redirectprinters'; E = { $_.properties.customRdpProperty.redirectprinters } }
	@{N = 'properties - customRdpProperty - redirectsmartcards'; E = { $_.properties.customRdpProperty.redirectsmartcards } }
	@{N = 'properties - customRdpProperty - usbdevicestoredirect'; E = { $_.properties.customRdpProperty.usbdevicestoredirect } }
	@{N = 'properties - customRdpProperty - use_multimon'; E = { $_.properties.customRdpProperty.use_multimon } }
	@{N = 'properties - customRdpProperty - videoplaybackmode'; E = { $_.properties.customRdpProperty.videoplaybackmode } }
	@{N = 'properties - vmTemplate - customImageId'; E = { $_.properties.vmTemplate.customImageId } }
	@{N = 'properties - vmTemplate - diskSizeGB'; E = { $_.properties.vmTemplate.diskSizeGB } }
	@{N = 'properties - vmTemplate - galleryImageOffer'; E = { $_.properties.vmTemplate.galleryImageOffer } }
	@{N = 'properties - vmTemplate - galleryImagePublisher'; E = { $_.properties.vmTemplate.galleryImagePublisher } }
	@{N = 'properties - vmTemplate - galleryImageSKU'; E = { $_.properties.vmTemplate.galleryImageSKU } }
	@{N = 'properties - vmTemplate - galleryItemId'; E = { $_.properties.vmTemplate.galleryItemId } }
	@{N = 'properties - vmTemplate - hibernate'; E = { $_.properties.vmTemplate.hibernate } }
	@{N = 'properties - vmTemplate - osDiskType'; E = { $_.properties.vmTemplate.osDiskType } }
	@{N = 'properties - migrationRequest'; E = { $_.properties.migrationRequest } }
	@{N = 'properties - objectId'; E = { $_.properties.objectId } }
	@{N = 'properties - personalDesktopAssignmentType'; E = { $_.properties.personalDesktopAssignmentType } }
	@{N = 'properties - preferredAppGroupType'; E = { $_.properties.preferredAppGroupType } }
	@{N = 'properties - privateEndpointConnections'; E = { $_.properties.privateEndpointConnections } }
	@{N = 'properties - publicNetworkAccess'; E = { $_.properties.publicNetworkAccess } }
	@{N = 'properties - registrationInfo'; E = { $_.properties.registrationInfo } }
	@{N = 'properties - ring'; E = { $_.properties.ring } }
	@{N = 'properties - ssoadfsAuthority'; E = { $_.properties.ssoadfsAuthority } }
	@{N = 'properties - ssoClientId'; E = { $_.properties.ssoClientId } }
	@{N = 'properties - ssoClientSecretKeyVaultPath'; E = { $_.properties.ssoClientSecretKeyVaultPath } }
	@{N = 'properties - ssoSecretType'; E = { $_.properties.ssoSecretType } }
)

# Set the required details
If ($OutputDetail -ne 'Summary') { $arrOutputDetails += $arrDetailsNormal }
If ($OutputDetail -eq 'High') { $arrOutputDetails += $arrDetailsHigh }

#region AzureFunctions
function Get-AzSPStoredCredentials {
	<#
    .SYNOPSIS
        Retrieve the Azure Service Principal Stored Credentials
    .EXAMPLE
        Get-AzSPStoredCredentials
    .CONTEXT
        Azure
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-08-03
        Purpose:        WVD Administration, through REST API calls
    #>
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true)]
		[string]$system ,
		[string]$tenantId
	)

	$strAzSPCredFolder = [System.IO.Path]::Combine( [environment]::GetFolderPath('CommonApplicationData') , 'ControlUp' , 'ScriptSupport' )
	$AzSPCredentials = $null

	Write-Verbose -Message "Get-AzSPStoredCredentials $system"

	[string]$credentialsFile = $(if ( -Not [string]::IsNullOrEmpty( $tenantId ) ) {
			[System.IO.Path]::Combine( $strAzSPCredFolder , "$($env:USERNAME)_$($tenantId)_$($System)_Cred.xml" )
		}
		else {
			[System.IO.Path]::Combine( $strAzSPCredFolder , "$($env:USERNAME)_$($System)_Cred.xml" )
		})

	Write-Verbose -Message "`tCredentials file is $credentialsFile"

	If (Test-Path -Path $credentialsFile) {
		try {
			if ( ( $AzSPCredentials = Import-Clixml -Path $credentialsFile ) -and -Not [string]::IsNullOrEmpty( $tenantId ) -and -Not $AzSPCredentials.ContainsKey( 'tenantid' ) ) {
				$AzSPCredentials.Add(  'tenantID' , $tenantId )
			}
		}
		catch {
			Write-Error -Message "The required PSCredential object could not be loaded from $credentialsFile : $_"
		}
	}
	Elseif ( $system -eq 'Azure' ) {
		## try old Azure file name 
		$azSPCredentials = Get-AzSPStoredCredentials -system 'AZ' -tenantId $AZtenantId 
	}
    
	if ( -not $AzSPCredentials ) {
		Write-Error -Message "The Azure Service Principal Credentials file stored for this user ($($env:USERNAME)) cannot be found at $credentialsFile.`nCreate the file with the Set-AzSPCredentials script action (prerequisite)."
	}
	return $AzSPCredentials
}

function Get-AzBearerToken {
	<#
    .SYNOPSIS
        Retrieve the Azure Bearer Token for an authentication session
    .EXAMPLE
        Get-AzBearerToken -SPCredentials <PSCredentialObject> -TenantID <string>
    .CONTEXT
        Azure
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-03-22
        Updated:        2020-05-08
                        Created a separate Azure Credentials function to support ARM architecture and REST API scripted actions
        Purpose:        WVD Administration, through REST API calls
    #>
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory = $true, HelpMessage = 'Azure Service Principal credentials' )]
		[ValidateNotNullOrEmpty()]
		[System.Management.Automation.PSCredential] $SPCredentials,
		[Parameter(Mandatory = $true, HelpMessage = 'Azure Tenant ID' )]
		[ValidateNotNullOrEmpty()]
		[string] $TenantID,
		[Parameter(Mandatory = $false, HelpMessage = 'Base URL for the scope' )]
		[ValidateNotNullOrEmpty()]
		[string] $baseURL = 'https://management.azure.com'
	)

    
	## https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
	[string]$uri = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"

	[hashtable]$body = @{
		grant_type    = 'client_credentials'
		client_Id     = $SPCredentials.UserName
		client_Secret = $SPCredentials.GetNetworkCredential().Password
		scope         = "$baseURL/.default"
	}

	[hashtable]$invokeRestMethodParams = @{
		Uri         = $uri
		Body        = $body
		Method      = 'POST'
		ContentType = 'application/x-www-form-urlencoded'
	}

	Invoke-RestMethod @invokeRestMethodParams | Select-Object -ExpandProperty access_token -ErrorAction SilentlyContinue
}


function Invoke-AzureRestMethod {
	<#
    .SYNOPSIS
    Invoke a REST method on Azure

    .EXAMPLE
    Invoke-AzureRestMethod -BearerToken $myBearerToken -uri $Uri -method GET -propertyToReturn 'Value'
    Returns the content of the 'value' property of the return of the REST call

    .CONTEXT
    Azure

    .PARAMETER BearerToken
    Your Azure bearer token

    .PARAMETER uri
    The REST call URI
    
    .PARAMETER body
    The REST call body, if required.
    
    .PARAMETER propertyToReturn
    The property of the return FROM THE REST CALL to output.
    Many Azure REST calls contain the most relevant data in a child property of the main return, the 'property' setting is used to only return the relevant data.

    .PARAMETER contentType
    The REST call content type. Default is 'application/json'
        
    .PARAMETER norest
    Use this switch to make the REST call with Invoke-WebRequest instead of Invoke-RestMethod
    
    .NOTES
        Version:        1.1
        Author:         Guy Leech
        Creation Date:  17-06-2022
    #>

	[CmdletBinding()]
	Param(
		[Parameter( Mandatory = $true, HelpMessage = 'A valid Azure bearer token' )]
		[ValidateNotNullOrEmpty()]
		[string]$BearerToken ,
		[string]$uri ,
		[ValidateSet('GET', 'POST', 'PUT', 'DELETE', 'PATCH')] ## add others as necessary
		[string]$method = 'GET' ,
		$body , ## not typed because could be hashtable or pscustomobject
		[string]$propertyToReturn,
		[string]$contentType = 'application/json' ,
		[switch]$norest
	)

	[hashtable]$header = @{
		'Authorization' = "Bearer $BearerToken"
	}

	if ( ! [string]::IsNullOrEmpty( $contentType ) ) {
		$header.Add( 'Content-Type'  , $contentType )
	}

	[hashtable]$invokeRestMethodParams = @{
		Uri     = $uri
		Method  = $method
		Headers = $header
	}

	if ( $PSBoundParameters[ 'body' ] ) {
		## convertto-json converts certain characters to codes so we convert back as Azure doesn't like them
		$invokeRestMethodParams.Add( 'Body' , (( $body | ConvertTo-Json -Depth 20 ) -replace '\\u003e' , '>' -replace '\\u003c' , '<' -replace '\\u0027' , '''' -replace '\\u0026' , '&' ))
	}

	$responseHeaders = $null

	if ( $PSVersionTable.PSVersion -ge [version]'7.0.0.0' ) {
		$invokeRestMethodParams.Add( 'ResponseHeadersVariable' , 'responseHeaders' )
	}

	## cope with pagination where get 100 results at a time
	do {
		$result = $null
		if ( $norest ) {
			$result = Invoke-WebRequest @invokeRestMethodParams
		}
		else {
			$result = Invoke-RestMethod @invokeRestMethodParams
		}

		if ( -not [String]::IsNullOrEmpty( $propertyToReturn ) ) {
			$result | Select-Object -ErrorAction SilentlyContinue -ExpandProperty $propertyToReturn
		}
		else {
			$result  ## don't pipe through select as will slow script down for large result sets if processed again after return
		}
		## now see if more data to fetch
		if ( $result ) {
			if ( $result.PSObject.Properties[ 'nextLink' ] ) {
				if ( $invokeRestMethodParams.Uri -eq $result.nextLink ) {
					Write-Warning -Message "Got same uri for nextLink as current $($result.nextLink)"
					break
				}
				else {
					## nextLink is different
					$invokeRestMethodParams.Uri = $result.nextLink
				}
			}
			else {
				$invokeRestMethodParams.Uri = $null ## no more data
			}
		}
	} while ( $result -and $null -ne $invokeRestMethodParams.Uri )
}

Function New-AVDRestParameters {
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory = $true, HelpMessage = 'The type of parameter set you wish to create.')]
		[ValidateSet('DeleteMSIXPackage', 'DeleteUserSession', 'DisconnectUserSession', 'ExpandMSIXImage', 'GetApplication', 'GetApplicationGroup', 'GetDesktop', 'GetHostpool', 'GetMSIXPackage', 'GetScalingPlan',
			'GetSessionHost', 'GetUserSession', 'GetWorkSpace', 'ListApplicationGroupsByResourceGroup', 'ListApplicationGroupsByResourceGroupAndType', 'ListApplicationGroupsBySubscription',
			'ListApplicationGroupsBySubscriptionAndType', 'ListApplicationsByApplicationGroup', 'ListDesktopsByApplicationGroup', 'ListHostPoolsByResourceGroup', 'ListHostPoolsBySubscription', 'ListMSIXPackagesByHostPool', 'ListScalingPlansByHostPool',
			'ListScalingPlansByResourceGroup', 'ListScalingPlansBySubscription', 'ListSessionHostsByHostPool', 'ListStartMenuItemsByApplicationGroup', 'ListUserSessionsBySessionHost', 'ListUserSessionsByHostPool', 'ListWorkSpacesByResourceGroup',
			'ListWorkSpacesBySubscription', 'SendUserMessage', 'SetSessionHostDrainMode', 'SetSessionHostAssignedUser')]
		[string]$Type,
		[Parameter(Mandatory = $true, HelpMessage = 'Enable to create an entire comment block with standard fields.')]
		[string]$Subscription,
		[Parameter(Mandatory = $false, HelpMessage = 'Azure Resource Group.')]
		[string]$ResourceGroup,
		[Parameter(Mandatory = $false, HelpMessage = 'AVD Application Group.')]
		[string]$ApplicationGroup,
		[Parameter(Mandatory = $false, HelpMessage = 'AVD Hostpool.')]
		[string]$HostPool,
		[Parameter(Mandatory = $false, HelpMessage = 'AVD Session Host.')]
		[string]$SessionHost,
		[Parameter(Mandatory = $false, HelpMessage = 'AVD Workspace.')]
		[string]$WorkSpace,
		[Parameter(Mandatory = $false, HelpMessage = 'MSIX package full name.')]
		[string]$MSIXPackage,
		[Parameter(Mandatory = $false, HelpMessage = 'Machine name (Session Host) FQDN.')]
		[string]$MachineFQDN,
		[Parameter(Mandatory = $false, HelpMessage = '(User) Session ID.')]
		[string]$SessionID,
		[Parameter(Mandatory = $false, HelpMessage = 'Message body for Send message to user.')]
		[string]$MessageBody,
		[Parameter(Mandatory = $false, HelpMessage = 'Message Title for Send message to user.')]
		[string]$MessageTitle,
		[Parameter(Mandatory = $false, HelpMessage = 'AVD Application Group.')]
		[ValidateSet('Desktop', 'RemoteApp')]
		[string]$ApplicationGroupType,
		[Parameter(Mandatory = $false, HelpMessage = 'Session Host allow new session ''true'' or ''false'' ')]
		[ValidateSet('true', 'false')]
		[string]$AllowNewSession,
		[Parameter(Mandatory = $false, HelpMessage = 'User to assign to Session Host (myuser@mydomain.com)')]
		[string]$AssignedUser,
		[Parameter(HelpMessage = 'Include the required propertyToReturn parameter to return only the subset of relevant values.')]
		[switch]$UsePropertyToReturn
	)

	# Set some defaults
	[string]$AzBase = "https://management.azure.com"
	[string]$strProv = 'providers/Microsoft.DesktopVirtualization'
	[string]$strAPI = '?api-version=2022-02-10-preview'

	# AVD Rest calls
	[hashtable]$hshAVDRestCalls = @{
		'DeleteMSIXPackage'                           = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/msixPackages/$MSIXPackage$strAPI"; 'Method' = 'DELETE' } 
		'DeleteUserSession'                           = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/sessionHosts/$SessionHost/userSessions/$SessionID$strAPI"; 'Method' = 'DELETE' }
		'DisconnectUserSession'                       = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/sessionHosts/$SessionHost/userSessions/$SessionID/disconnect$strAPI"; 'Method' = 'POST' }
		'ExpandMSIXImage'                             = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/expandMsixImage$strAPI"; 'Method' = 'POST'; 'PropertyToReturn' = 'value' }
		'GetApplication'                              = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/applicationGroups/$ApplicationGroup/applications/$AZApplicationName$strAPI" }
		'GetApplicationGroup'                         = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/applicationGroups/$ApplicationGroup$strAPI" }
		'GetDesktop'                                  = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/applicationGroups/$ApplicationGroup/desktops/$AZDesktopName$strAPI" }
		'GetHostpool'                                 = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool$strAPI"; 'PropertyToReturn' = 'value' }
		'GetMSIXPackage'                              = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/msixPackages/$MSIXPackage$strAPI" }
		'GetScalingPlan'                              = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/scalingPlans/$AzScalingPlan$strAPI" }
		'GetSessionHost'                              = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/sessionHosts/$MachineFQDN$strAPI" }
		'GetUserSession'                              = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/sessionHosts/$SessionHost/userSessions/$SessionID$strAPI" }
		'GetWorkSpace'                                = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/workspaces/$WorkSpace$strAPI" }
		'ListApplicationsByApplicationGroup'          = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/applicationGroups/$ApplicationGroup/applications$strAPI"; 'PropertyToReturn' = 'value'; 'norest' = $true }
		'ListApplicationGroupsByResourceGroup'        = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/applicationGroups$strAPI"; 'PropertyToReturn' = 'value' }
		'ListApplicationGroupsBySubscription'         = @{'Uri' = "$AzBase/subscriptions/$Subscription/$strProv/applicationGroups$strAPI"; 'PropertyToReturn' = 'value' }
		'ListApplicationGroupsByResourceGroupAndType' = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/applicationGroups$strAPI&$filter=applicationGroupType eq ''$ApplicationGroupType''"; 'PropertyToReturn' = 'value' }
		'ListApplicationGroupsBySubscriptionAndType'  = @{'Uri' = "$AzBase/subscriptions/$Subscription/$strProv/applicationGroups$strAPI&$filter=applicationGroupType eq ''$ApplicationGroupType''"; 'PropertyToReturn' = 'value' }
		'ListDesktopsByApplicationGroup'              = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/applicationGroups/$ApplicationGroup/desktops$strAPI" ; 'PropertyToReturn' = 'value' }
		'ListHostPoolsByResourceGroup'                = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools$strAPI" ; 'PropertyToReturn' = 'value' }
		'ListHostPoolsBySubscription'                 = @{'Uri' = "$AzBase/subscriptions/$Subscription/$strProv/hostPools$strAPI"; 'PropertyToReturn' = 'value' }
		'ListMSIXPackagesByHostPool'                  = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/msixPackages$strAPI" ; 'PropertyToReturn' = 'value' }
		'ListScalingPlansByHostPool'                  = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/scalingPlans$strAPI"; 'PropertyToReturn' = 'value' }
		'ListScalingPlansByResourceGroup'             = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/scalingPlans$strAPI"; 'PropertyToReturn' = 'value' }
		'ListScalingPlansBySubscription'              = @{'Uri' = "$AzBase/subscriptions/$Subscription/$strProv/scalingPlans$strAPI"; 'PropertyToReturn' = 'value' }
		'ListSessionHostsByHostPool'                  = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/sessionHosts$strAPI"; 'PropertyToReturn' = 'value' }
		'ListStartMenuItemsByApplicationGroup'        = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/applicationGroups/$ApplicationGroup/startMenuItems$strAPI"; 'PropertyToReturn' = 'value' }
		'ListUserSessionsBySessionHost'               = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/sessionHosts/$SessionHost/userSessions$strAPI"; 'PropertyToReturn' = 'value' }
		'ListUserSessionsByHostPool'                  = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/userSessions$strAPI"; 'PropertyToReturn' = 'value' }
		'ListWorkSpacesByResourceGroup'               = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/workspaces$strAPI"; 'PropertyToReturn' = 'value' }
		'ListWorkSpacesBySubscription'                = @{'Uri' = "$AzBase/subscriptions/$Subscription/$strProv/workspaces?api-version=$strAPI" ; 'PropertyToReturn' = 'value' } 
		'SetSessionHostDrainMode'                     = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/sessionHosts/$MachineFQDN$strAPI"
			'Method'                           = 'PATCH'
			'Body'                             = @{'Properties' = @{'allowNewSession' = $AllowNewSession
				}
			}
		}
		'SetSessionHostAssignedUser'                  = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/sessionHosts/$MachineFQDN$strAPI"
			'Method'                              = 'PATCH'
			'Body'                                = @{'Properties' = @{'assignedUser' = $AssignedUser }
			}
		}
		'SendUserMessage'                             = @{'Uri' = "$AzBase/subscriptions/$Subscription/resourceGroups/$ResourceGroup/$strProv/hostPools/$HostPool/sessionHosts/$SessionHost/userSessions/$SessionID/sendMessage$strAPI"
			'Method'                   = 'POST'
			'Body'                     = @{'messageBody' = $MessageBody; 'messageTitle' = $MessageTitle }
		}
	}

	# Get the required param set
	$hshOutput = $hshAVDRestCalls.$Type

	# Test if the URL is valid, there should be only ONE instance of '/ / ' in the Uri
	If (([regex]::Matches($hshOutput.Uri, "//" )).count -ne 1) {
		Write-Error -Message "The URI appears to be malformed. Check if you passed all the required parameters, look for double forward slashes in the Uri: $($hshOutput.Uri)"
	}
	Else {
		# Add the method if it is not there
		If (-not $hshOutput.ContainsKey('Method')) {
			Write-Verbose -Message "REST parameters did not contain Method, which means it must just be GET, add that."
			$hshOutput.Add('Method', 'GET')
		}

		# Remove property subset if this was not specified
		If (($hshOutput.ContainsKey('PropertyToReturn')) -and (!($usePropertyToReturn))) {
			Write-Verbose -Message "REST parameters contains a PropertyToReturn value, but UsePropertyToReturn was not specified so it will be removed."
			$hshOutput.Remove('PropertyToReturn')
		}
		Write-Verbose -Message "URI: $($hshOutput.Uri)"
		# Return the corresponding hashtable
		$hshOutput
	}
}
#endregion AzureFunctions

# Authenticate to Azure tenant
If ($azSPCredentials = Get-AzSPStoredCredentials -system 'Azure' -tenantId $AzTenantId ) {
	# Sign in to Azure with a Service Principal with Contributor Role at Subscription level and retrieve the bearer token
	Write-Verbose -Message "Authenticating to tenant $($azSPCredentials.tenantID) as $($azSPCredentials.spCreds.Username)"
	if ( -Not ( $azBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID ) ) {
		Throw "Failed to get Azure bearer token"
	}
}

# Get the hostpools by subscription and resource groups
try {
	Write-Verbose -Message "Retrieving host pools..."
	$hshRESTParams = New-AVDRestParameters -type ListHostpoolsByResourceGroup -Subscription $AzSubscription -ResourceGroup $AzResourceGroup -UsePropertyToReturn
	$objHostPools = Invoke-AzureRestMethod -BearerToken $azBearerToken @hshRESTParams
}
catch {
	[string]$strMessage = ($_.errordetails.message | ConvertFrom-Json).error.message
	[int]$intCode = ($_.errordetails.message | ConvertFrom-Json).error.code
	Switch ($intCode) {
		404 { Throw "The host pools could not be found.`nPlease ensure you are running this script on an AVD SessionHost.`n$strMessage" }
		default { Throw "There was an unexpected error while trying to retrieve the host pools.`nPlease ensure you are running this script on an AVD SessionHost.`n$strMessage" }
	}
}

# Go through the hostpools to find the one with the SessionHost
Foreach ($pool in $objHostPools) {
	# The name has to match
	$pool.properties.vmTemplate = $pool.properties.vmTemplate | ConvertFrom-Json
	If ($pool.properties.vmTemplate.namePrefix -eq $MachineName.Substring(0, $MachineName.LastIndexOf('-'))) {
		Write-Verbose -Message "Pool namePrefix matches machine name, looks for the SessionHost in that pool."
		# Get the session hosts in pool and see if the VM is there
		$hshRESTParams = New-AVDRestParameters -type ListSessionHostsByHostPool -Subscription $AzSubscription -ResourceGroup $AzResourceGroup -HostPool $pool.name -UsePropertyToReturn
		$objSessionHosts = Invoke-AzureRestMethod -BearerToken $azBearerToken @hshRESTParams 
		
		# Check the machine we're looking for is actually in there.
		Foreach ($SessionHost in $objSessionHosts) {
			If ($SessionHost.properties.virtualMachineId -contains $AzVmId) {
				# SessionHost found, this is the right pool, output the details
				Write-Verbose -Message "SessionHost found in $($pool.name), this IS the hostpool we are looking for."
				# Expand some of the nested properties
				[array]$arrProperties = $pool.properties.customRdpProperty.Split(';')
				$hshCustomRdpProperty = [ordered]@{}
				Foreach ($rdpProperty in $arrProperties) {
					If (!([string]::IsNullOrEmpty($rdpProperty))) {	$hshCustomRdpProperty.Add($rdpProperty.Split(':')[0], $rdpProperty.Split(':')[2]) }
				}
				$pool.properties.customRdpProperty = $hshCustomRdpProperty
				$pool | Select-Object -Property $arrOutputDetails | Sort-Object
				Exit 0
			}
		}
		Else {
			Write-Verbose -Message "$MachineName not found in hostpool $($pool.name)"
		}
	}
}

# If the script got here, looks like the host pool could not be found.
Write-Output -InputObject "Pool containing SessionHost $MachineName could not be found."
Exit 1

