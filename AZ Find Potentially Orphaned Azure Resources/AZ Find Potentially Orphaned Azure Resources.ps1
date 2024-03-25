#requires -version 3.0

<#
.SYNOPSIS
    Find Azure resources which should probably be assigned to resources but are not, eg disks, network interfaces, network security groups , public IP addresses

.DESCRIPTION
    Using REST API calls

.PARAMETER azid
    The relative URI of the Azure VM
    
.PARAMETER AZtenantId
    The azure tenant ID
    
.PARAMETER resourceGroupOnly
    Only return results for the resource group containing the AZid
    
.PARAMETER sortby
    Sort the results by this property
    
.PARAMETER raw
    Output raw objects to pipeline rather than text

.NOTES
    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2022-09-16
    Updated:        2022-06-17  Guy Leech  Added public IP addresses and changed mechanisms for determining unattached
                    2024-02-08  Guy Leech  Added setting of TLS12 & TLS13
                    2024-02-15  Guy Leech  Improved mechanism for determinging API version to use
                    2024-02-16  Guy Leech  Switched to checking managedBy rather than getting full resource details
                    2024-02-21  Guy Leech  Error 429 handling. API version caching moved to function that uses it. Fix for public IP attached to NAT gateway & network interfaces with private end points
                    2024-02-22  Guy Leech  Fixed bug where unattached network resources not flagged as orphans. Added code to deal with less than perfect certificates (code moved to AZ functions)
#>

## TODO Need to check that the parent/associated resource still exists - e.g. does VM still exist for a NIC?

[CmdletBinding()]

Param
(
    [string]$AZid ,## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId ,
    [ValidateSet('Yes','No')]
    [string]$resourceGroupOnly = 'Yes',
    [string]$sortby = 'type' ,
    [switch]$raw
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 400
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

## exclude resource types that we cannot determine if have been used or not in this script (AVD resources can be checked by looking at session usage but we don't do that currently)

[string[]]$excludedResourceTypes = @(
    'Microsoft.AAD/DomainServices'
    'Microsoft.Compute/virtualMachines/extensions'
    'Microsoft.Insights/activityLogAlerts'
    'Microsoft.DesktopVirtualization/workspaces' ## since we do not know if used or not and have no runnable/bootable resources
    'Microsoft.DesktopVirtualization/hostpools'  ## as above
    'Microsoft.DesktopVirtualization/applicationgroups' ## ""
)

[string]$providersApiVersion = '2021-04-01'
[string]$resourceManagementApiVersion = '2021-04-01'
[string]$baseURL = 'https://management.azure.com'
[string]$credentialType = 'Azure'
[hashtable]$script:apiversionCache = @{}

Write-Verbose -Message "AZid is $AZid"

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
        [Parameter(Mandatory=$true)]
        [string]$system ,
        [string]$tenantId
    )

    $strAzSPCredFolder = [System.IO.Path]::Combine( [environment]::GetFolderPath('CommonApplicationData') , 'ControlUp' , 'ScriptSupport' )
    $AzSPCredentials = $null

    Write-Verbose -Message "Get-AzSPStoredCredentials $system"

    [string]$credentialsFile = $(if( -Not [string]::IsNullOrEmpty( $tenantId ) )
        {
            [System.IO.Path]::Combine( $strAzSPCredFolder , "$($env:USERNAME)_$($tenantId)_$($System)_Cred.xml" )
        }
        else
        {
            [System.IO.Path]::Combine( $strAzSPCredFolder , "$($env:USERNAME)_$($System)_Cred.xml" )
        })

    Write-Verbose -Message "`tCredentials file is $credentialsFile"

    If (Test-Path -Path $credentialsFile)
    {
        try
        {
            if( ( $AzSPCredentials = Import-Clixml -Path $credentialsFile ) -and -Not [string]::IsNullOrEmpty( $tenantId ) -and -Not $AzSPCredentials.ContainsKey( 'tenantid' ) )
            {
                $AzSPCredentials.Add(  'tenantID' , $tenantId )
            }
        }
        catch
        {
            Write-Error -Message "The required PSCredential object could not be loaded from $credentialsFile : $_"
        }
    }
    Elseif( $system -eq 'Azure' )
    {
        ## try old Azure file name 
        $azSPCredentials = Get-AzSPStoredCredentials -system 'AZ' -tenantId $AZtenantId 
    }
    
    if( -not $AzSPCredentials )
    {
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
                        2022-06-28
                        Added -scope as argument so can authenticate for Graph as well as Azure
                        2022-07-04
                        Added optional retry mechanism in case of transient Azure errors
        Purpose:        WVD Administration, through REST API calls
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='Azure Service Principal credentials' )]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCredential] $SPCredentials,

        [Parameter(Mandatory=$true, HelpMessage='Azure Tenant ID' )]
        [ValidateNotNullOrEmpty()]
        [string] $TenantID ,

        [Parameter(Mandatory=$true, HelpMessage='Authentication scope' )]
        [ValidateNotNullOrEmpty()]
        [string] $scope
    )

    
    ## https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
    [string]$uri = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"

    [hashtable]$body = @{
        grant_type    = 'client_credentials'
        client_Id     = $SPCredentials.UserName
        client_Secret = $SPCredentials.GetNetworkCredential().Password
        scope         = "$scope/.default"
    }
    
    [hashtable]$invokeRestMethodParams = @{
        Uri             = $uri
        Body            = $body
        Method          = 'POST'
        ContentType     = 'application/x-www-form-urlencoded'
    }
    
    if( -Not (Get-Variable -Scope script -name internetSettingsSet -ErrorAction SilentlyContinue ) )
    {
        # see https://stackoverflow.com/questions/11696944/powershell-v3-invoke-webrequest-https-error
        #     https://stackoverflow.com/questions/2859790/the-request-was-aborted-could-not-create-ssl-tls-secure-channel
        if (-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type )
        {
            Add-Type -ErrorAction SilentlyContinue -TypeDefinition @"
                using System.Net;
                using System.Security.Cryptography.X509Certificates;
                public class TrustAllCertsPolicy : ICertificatePolicy {
                    public bool CheckValidationResult(
                        ServicePoint srvPoint, X509Certificate certificate,
                        WebRequest request, int certificateProblem) {
                        return true;
                    }
                }
"@
        }
        [System.Net.ServicePointManager]::CertificatePolicy = New-Object -TypeName TrustAllCertsPolicy
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13
        $script:internetSettingsSet = [datetime]::Now
    }
    ## else already done

    Invoke-RestMethod @invokeRestMethodParams | Select-Object -ExpandProperty access_token -ErrorAction SilentlyContinue
}

function Invoke-AzureRestMethod {

    [CmdletBinding()]
    Param(
        [Parameter( Mandatory=$true, HelpMessage='A valid Azure bearer token' )]
        [ValidateNotNullOrEmpty()]
        [string]$BearerToken ,
        [string]$uri ,
        [ValidateSet('GET','POST','PUT','DELETE','PATCH')] ## add others as necessary
        [string]$method = 'GET' ,
        $body , ## not typed because could be hashtable or pscustomobject
        [string]$propertyToReturn = 'value' ,
        [string]$contentType = 'application/json' ,
        [switch]$norest ,
        [switch]$newestApiVersion = $true ,
        [string]$type , ## help us with looking up API versions & caching
        [int]$retries = 0 ,
        [int]$retryIntervalMilliseconds = 2500 ,
        [int]$tooBusyRetries = 2 ,
        $warningCollection
    )

    ## pseudo static variable , at least means we don't need to define it outside of this function although is script scope not function only
    if( -Not ( Get-Variable -Scope script -Name lastBusyRetry -ErrorAction SilentlyContinue ) )
    {
        $script:lastBusyRetry = [datetime]::MinValue
    }
    ## else have called function previously and will have created the script scope variable
 
    if( -Not ( Get-Variable -Scope script -Name totalRequests -ErrorAction SilentlyContinue ) )
    {
        [int]$script:totalRequests = 0
    }
    ## else have called function previously and will have created the script scope variable
    
    if( -Not (Get-Variable -Scope script -name internetSettingsSet -ErrorAction SilentlyContinue ) )
    {
        # see https://stackoverflow.com/questions/11696944/powershell-v3-invoke-webrequest-https-error
        #     https://stackoverflow.com/questions/2859790/the-request-was-aborted-could-not-create-ssl-tls-secure-channel
        if (-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type )
        {
            Add-Type -ErrorAction SilentlyContinue -TypeDefinition @"
                using System.Net;
                using System.Security.Cryptography.X509Certificates;
                public class TrustAllCertsPolicy : ICertificatePolicy {
                    public bool CheckValidationResult(
                        ServicePoint srvPoint, X509Certificate certificate,
                        WebRequest request, int certificateProblem) {
                        return true;
                    }
                }
"@
        }
        [System.Net.ServicePointManager]::CertificatePolicy = New-Object -TypeName TrustAllCertsPolicy
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13
        $script:internetSettingsSet = [datetime]::Now
    }
    ## else already done

    [hashtable]$header = @{
        'Authorization' = "Bearer $BearerToken"
    }

    if( ! [string]::IsNullOrEmpty( $contentType ) )
    {
        $header.Add( 'Content-Type'  , $contentType )
    }

    [hashtable]$invokeRestMethodParams = @{
        Uri             = $uri
        Method          = $method
        Headers         = $header
    }

    if( $PSBoundParameters[ 'body' ] )
    {
        ## convertto-json converts certain characters to codes so we convert back as Azure doesn't like them
        $invokeRestMethodParams.Add( 'Body' , (( $body | ConvertTo-Json -Depth 20 ) -replace '\\u003e' , '>' -replace '\\u003c' , '<' -replace '\\u0027' , '''' -replace '\\u0026' , '&' ))
    }

    $responseHeaders = $null

    if( $PSVersionTable.PSVersion -ge [version]'7.0.0.0' )
    {
        $invokeRestMethodParams.Add( 'ResponseHeadersVariable' , 'responseHeaders' )
    }

    [string]$apiversion = $null

    # see if call requires an Api version and doesn't have one
    if( $uri -match '/providers/([^/]+)/([^/]+)' )
    {
        if( $uri -match '\?api\-version=20\d\d-\d\d-\d\d' )
        {
            ## nothing to do as call already has an api version
        }
        elseif( -Not [string]::IsNullOrEmpty( $type ) -and ( $cached = $script:apiversionCache[ $type ] ))
        {
            $apiversion = $cached
            $invokeRestMethodParams.uri += "?api-version=$apiversion"
        }
        else
        {
            $apiversion = '42' ## force error which will return list of valid api versions (as of 2024/02/16 this will not get used as we get api version from 1 off providers query)
            $resourceType = $null
            $provider = $null
            ## see if we have cached entry already for this provider and use that to save a REST call
            if( [string]::IsNullOrEmpty( $type ) )
            {
                if( $uri -match '/providers/([^/]+)/([^/]+)' )
                {
                    $provider = $Matches[ 1 ]
                    $resourceType = $Matches[ 2 ] -replace '[?&].*$' ## get rid of any trailing parameters
                }
                else
                {
                    Write-Warning -Message "No provider for API version determination matched in $uri"
                }
            }
            else ## $type passed from caller so use that
            {
                $provider,$resourceType = $type -split '/' , 2
            }
            
            if( -Not ( Get-Variable -Scope script -Name allProviders -ErrorAction SilentlyContinue ) )
            {
                ## get all provider details so we can pick version if we don't have one. This must have a version number otherwise could go infinitely recursive
                [array]$script:allProviders = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/providers?api-version=$providersApiVersion" -retries 2 )

                Write-Verbose -Message "$([datetime]::Now.ToString( 'G' )): Got $($script:allProviders.count) providers"
            }
            ## else have called function previously and needed a provider so will have populated the script scope variable with all provider versions

            $providerDefinition = $script:allProviders | Where-Object namespace -ieq $provider
            if( $null -ne $providerDefinition )
            {
                ## TODO do we need to look for subtypes which could come from passed $type parameter ?
                $allApiVersions = @( $providerDefinition.resourceTypes | Where-Object resourceType -ieq $resourceType | Select-Object -ExpandProperty apiVersions )
                $releasedApiVersions = @( $allApiVersions | Where-Object { $_ -notmatch '(preview|beta|alpha)' } )
                if( $null -ne $releasedApiVersions -and $releasedApiVersions.Count -gt 0 )
                {
                    $apiversion = $releasedApiVersions[ $(if( $newestApiVersion ) { 0 } else { -1 }) ]
                }
                elseif( $null -ne $allApiVersions -and $allApiVersions.Count -gt 0 )
                {
                    $apiversion = $allApiVersions[ $(if( $newestApiVersion ) { 0 } else { -1 }) ]
                    Write-Verbose -Message "** no release api version for $($providerDefinition.namespace) / $resourceType so having to use $apiversion"
                }
                else
                {
                    [string]$warning = $null
                    $warning =  "No api versions retrieved for provider $($providerDefinition.Namespace) resource type $resourceType in URI $uri"
                    
                    if( $null -ne $warningCollection )
                    {
                        $null = $warningCollection.Add( $warning ) ## don't know what type of collection it is so could return a value
                    }
                    else
                    {
                        Write-Warning -Message $warning
                    }
                }
            }
            else
            {
                Write-Verbose -Message "*** no provider found for namespace `"$type`""
            }
            ## deal with uri that already has ? in it as ther can be only one
            [string]$link = '?'
            if( $invokeRestMethodParams.uri.IndexOf( '?' ) -gt 0 )
            {
                $link = '&'
            }
            $invokeRestMethodParams.uri += "$($link)api-version=$apiversion"
        }
    }
    ## else no provider specified so hopefully doesn't need an API version if not already specified
 
    [string]$lastURI = $null

    ## cope with pagination where get 100 results at a time
    do
    {
        $thisretry = $retries
        $thisBusyRetry = $tooBusyRetries
        $error.Clear()
        $exception = $null
        do
        {
            $exception = $null
            $result = $null

            try
            {
                $script:totalRequests++
                if( $norest )
                {
                    $result = Invoke-WebRequest @invokeRestMethodParams
                }
                else
                {
                    $result = Invoke-RestMethod @invokeRestMethodParams
                }
            }
            catch
            {
                $exception = $_
                if( $exception -match $script:tooBusyErrorRegex)
                {
                    if( $thisBusyRetry -ge 0 )
                    {
                        [int]$waitForSeconds = $Matches[2]
                        [datetime]$now = [datetime]::Now
                        Write-Verbose -Message "  $($now.ToString('G')): Too busy error so waiting $waitForSeconds seconds as suggested by Azure response ($script:totalRequests total requests)"
                        if( $script:lastBusyRetry -gt [datetime]::MinValue )
                        {
                            Write-Verbose -Message "    last busy retry $([math]::Round( ($now - $script:lastBusyRetry).TotalSeconds )) seconds ago"
                        }
                        $script:lastBusyRetry = $now
                        Start-Sleep -Seconds $waitForSeconds
                        $thisBusyRetry--
                    }
                    else
                    {
                        break
                    }
                }
                elseif( $thisretry -ge 1 ) ## do not sleep if no retries requested or this was the last retry
                {
                    Write-Verbose -Message "  Error so sleeping $retryIntervalMilliseconds ms : $exception"
                    Start-Sleep -Milliseconds $retryIntervalMilliseconds
                }
            }
            if( -not $exception )
            {
                break
            }
        } while( --$thisretry -ge 0 ) ## exiting when 429 error retries are over is handled, inelegantly, within the loop code

        ## $result -eq $null does not mean there was an exception so we need to track that separately to know whether to throw an exception here
        if( $exception )
        {
            ## last call gave an exception
            Throw "Exception $($exception.ToString()) originally occurred on line number $($exception.InvocationInfo.ScriptLineNumber)"
        }
        elseif( $error.Count -gt 0 -and $error[0].ToString -notmatch $script:tooBusyErrorRegex ) ## don't report 429 errors as there could be a lot
        {
            [string]$warning = $null
            if( $error[0].ToString() -match '^{.*}$' )
            {
                $warning = "Transient errors on request $($invokeRestMethodParams.Uri) - $($error[0].ToString() | ConvertFrom-Json | Select-Object -ExpandProperty error|Select-Object -ExpandProperty message)"
            }
            else ## not json
            {
                $warning = "Transient errors on request $($invokeRestMethodParams.Uri) - $($error[0].ToString())"
            }
            if( $null -ne $warningCollection )
            {
                $null = $warningCollection.Add( $warning ) ## don't know what type of collection it is so could return a value
            }
            else
            {
                Write-Warning -Message $warning
            }
        }

        ## cache the api version if we had to figure it out
        if( -Not [string]::IsNullOrEmpty( $apiversion ) -and -not [string]::IsNullOrEmpty( $type ) -and -Not $script:apiversionCache.ContainsKey( $type ) )
        {
            try
            {
                $script:apiversionCache.Add( $type , $apiversion )
            }
            catch
            {
                $null
            }
        }

        $lastURI = $invokeRestMethodParams.uri

        if( -not [String]::IsNullOrEmpty( $propertyToReturn ) )
        {
            $result | Select-Object -ErrorAction SilentlyContinue -ExpandProperty $propertyToReturn
        }
        else
        {
            $result  ## don't pipe through select as will slow script down for large result sets if processed again after return
        }
        ## now see if more data to fetch
        if( $result )
        {
            if( ( $nextLink = $result.PSObject.Properties[ 'nextLink' ] ) -or ( $nextLink = $result.PSObject.Properties[ '@odata.nextLink' ] ) )
            {
                if( $invokeRestMethodParams.Uri -eq $nextLink.value )
                {
                    Write-Warning -Message "Got same uri for nextLink as current $($nextLink.value)"
                    break
                }
                else ## nextLink is different
                {
                    $invokeRestMethodParams.Uri = $nextLink.value
                }
            }
            else
            {
                $invokeRestMethodParams.Uri = $null ## no more data
            }
        }
    } while( $result -and $null -ne $invokeRestMethodParams.Uri )
}

#endregion AzureFunctions

[datetime]$startTime = [datetime]::Now

$azSPCredentials = $null
$azSPCredentials = Get-AzSPStoredCredentials -system $credentialType -tenantId $AZtenantId

If ( -Not $azSPCredentials )
{
    Exit 1 ## will already have output error
}

## TLS and certificate handling now set in the Azure functions

# Sign in to Azure with the Service Principal retrieved from the credentials file and retrieve the bearer token
Write-Verbose -Message "Authenticating to tenant $($azSPCredentials.tenantID) as $($azSPCredentials.spCreds.Username)"
if( -Not ( $azBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID -scope $baseURL ) )
{
    Throw "Failed to get Azure bearer token"
}

[string]$subscriptionId = $null
[string]$resourceGroupName = $null

## subscriptions/58ffa3cb-2f63-4242-a06d-deadbeef/resourceGroups/WVD/providers/Microsoft.Compute/virtualMachines/GLMW10WVD-0
if( $AZid -match '\bsubscriptions/([^/]+)/resourceGroups/([^/]+)/providers/Microsoft\.' )
{
    $subscriptionId = $Matches[1]
    $resourceGroupName = $Matches[2]
}
else
{
    Throw "Failed to parse subscription id and resource group from $AZid"
}

## https://docs.microsoft.com/en-us/rest/api/resources/resources/list-by-resource-group
[string]$resourcesURL = $null
if( $resourceGroupOnly -ieq 'yes' )
{
    $resourcesURL = "resourceGroups/$resourceGroupName"
}
## else ## will be for the subscription

[array]$allResources = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/$resourcesURL/resources`?`$expand=createdTime,changedTime,lastusedTime,managedBy&api-version=$resourceManagementApiVersion" -retries 2 | Where-Object type -NotIn $excludedResourceTypes )

[hashtable]$resourceDetails = @{}
[hashtable]$resourceGroups = @{}

[string[]]$typesOfInterest = @( 'Microsoft.Network/networkInterfaces' , 'Microsoft.Compute/disks' , 'Microsoft.Network/networkSecurityGroups' , 'Microsoft.Network/publicIPAddresses' )

ForEach( $resource in $allResources )
{
    ## only cache the resource types that could be orphaned from a deleted parent
    if( $resource.type -in $typesOfInterest )
    {
        if( $resource.PSObject.Properties[ 'managedBy' ] ) ## use this to save another AZ request
        {
            if( -Not [string]::IsNullOrEmpty( $resource.ManagedBy ))
            {
                if( -Not $resourceDetails.ContainsKey( $resource.managedBy ) )
                {
                    if( $resource.type -ieq 'Microsoft.Compute/disks' )
                    {
                        [hashtable]$diskState = @{ 'DiskState' = 'Attached' }
                        if( $resource.psObject.Properties[ 'properties' ] )
                        {
                            Add-Member -InputObject $resource.properties -NotePropertyMembers $diskState -Force
                        }
                        else
                        {
                            Add-Member -InputObject $resource -MemberType NoteProperty -Name properties -Value ([pscustomobject]$diskState)
                        }
                    }
                    else ## not a disk
                    {
                        $null = $null
                    }
                    $resourceDetails.Add( $resource.id , $resource )
                }
                ## else resource is already in dictionary
            }
            else ## managedby present but empty which we assume means not managed
            {
                $null
            }
        }
        elseif( $resourceDetail = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$($resource.id)" -retries 2 -newestApiVersion -propertyToReturn $null -type $resource.type )
        {      
            $resourceDetails.Add( $resource.id , $resourceDetail ) ## use in 2nd pass
        }
        else
        {
            Write-Warning -Message "Failed to get details for resource $($resource.id)"
        }
        
        if( $resourceGroupOnly -ieq 'no')
        {
            if( $resource.id -match '\bsubscriptions/([^/]+)/resourceGroups/([^/]+)/providers/Microsoft\.' )
            {
                [string]$thisResourceGroupName = $Matches[2]

                if( -Not $resourceGroups.ContainsKey( $thisResourceGroupName ))
                {
                    try
                    {
                        $resourceGroups.Add( $thisResourceGroupName , $resource )
                    }
                    catch
                    {
                        ## already got it
                    }
                }
                ## else already got it
                Add-Member -InputObject $resource -MemberType NoteProperty -Name 'Resource Group' -Value $thisResourceGroupName
            }
            else
            {
                Write-Warning -Message "Failed to determine resource group from $($resource.id)"
            }
        }
    }
    ## else not an orphanable type
}

## second pass if Microsoft.Compute/disks or Microsoft.Network/networkinterfaces or Microsoft.Network/networkSecurityGroups then check parent VM (if there is one) as we will have it in the unused collection
## cache VM details so if disk is for a machine we already have, we don't need to get it from Azure again
$potentiallyUnassignedResources = New-Object -TypeName System.Collections.Generic.List[object] ## make a generic list so we can add extra items later as necessary
[array]$parents = @( ForEach( $potentiallyUnassignedResource in $allResources )
{
    [bool]$childResource = $false
    $parent = $null

    if( $potentiallyUnassignedResource.type -ieq 'Microsoft.Network/networkInterfaces' )
    {
        $childResource = $true
        if( $resourceDetail = $resourceDetails[ $potentiallyUnassignedResource.id ] )
        {
            if( -Not ( $parent = $resourceDetail | Select-Object -ExpandProperty properties | Select-Object -ExpandProperty virtualMachine -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id -ErrorAction SilentlyContinue ) )
            {
                $parent = $resourceDetail | Select-Object -ExpandProperty properties -ErrorAction SilentlyContinue | Select-Object -ExpandProperty privateEndpoint -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id -ErrorAction SilentlyContinue
            }
        }
        else
        {
            Write-Warning -Message "Unable to get details for network interface `"$($potentiallyUnassignedResource.id)`""
        }    
    }
    elseif( $potentiallyUnassignedResource.type -ieq 'Microsoft.Compute/disks' )
    {
        $childResource = $true
        ## if there is a parent, check that disk state is not unattached
        ##if( ( $parent = $resourceDetails[ $potentiallyUnassignedResource.id ] | Select-Object -ExpandProperty managedBy -ErrorAction SilentlyContinue ) -and ( $resourceDetail = $resourceDetails[ $potentiallyUnassignedResource.id ] ) -and $resourceDetail.properties.DiskState -ieq 'Unattached' )
        if( ( $resourceDetail = $resourceDetails[ $potentiallyUnassignedResource.id ] ) -and $resourceDetail.psobject.properties[ 'properties' ] -and $resourceDetail.properties.psobject.properties[ 'DiskState' ]  )
        {
            $parent = ( $resourceDetail.properties.DiskState -ine 'Unattached' ) ## don't need id
        }
        else
        {
            Write-Verbose -Message "Unable to get details for disk `"$($potentiallyUnassignedResource.id)`" so assuming unattached" ## there was a managedby property but it was empty
            $parent = $null
        }    
    }
    elseif( $potentiallyUnassignedResource.type -ieq 'Microsoft.Network/networkSecurityGroups' )
    {
       $childResource = $true
       if( $resourceDetail = $resourceDetails[ $potentiallyUnassignedResource.id ] )
       {
            if( $resourceDetail.properties.psobject.properties[ 'subnets' ] -and $resourceDetail.properties.subnets.Count -gt 0 )
            {
                $parent = $resourceDetail.properties | Select-Object -ExpandProperty subnets | Select-Object -ExpandProperty id -ErrorAction SilentlyContinue
            }
            elseif( $resourceDetail.properties.psobject.properties[ 'networkInterfaces' ] -and $resourceDetail.properties.networkInterfaces.Count -gt 0 )
            {
                $parent = $resourceDetail.properties | Select-Object -ExpandProperty networkInterfaces | Select-Object -ExpandProperty id -ErrorAction SilentlyContinue
            }
        }
        else
        {
            Write-Warning -Message "Unable to get details for network security group `"$($potentiallyUnassignedResource.id)`""
        }

    }
    elseif( $potentiallyUnassignedResource.type -ieq 'Microsoft.Network/publicIPAddresses' )
    {
       $childResource = $true
       if( $resourceDetail = $resourceDetails[ $potentiallyUnassignedResource.id ] )
       {
            if( -Not ( $parent = $resourceDetail | Select-Object -ExpandProperty properties -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ipConfiguration -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id -ErrorAction SilentlyContinue ) )
            {
                $parent = $resourceDetail | Select-Object -ExpandProperty properties -ErrorAction SilentlyContinue | Select-Object -ExpandProperty natGateway -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id -ErrorAction SilentlyContinue
            }
        }
        else
        {
            Write-Warning -Message "Unable to get details for public IP address `"$($potentiallyUnassignedResource.id)`""
        }
    }

    if( $childResource )
    {
        if( $null -eq $parent -or $false -eq $parent )
        {
            $potentiallyUnassignedResources.Add( $potentiallyUnassignedResource )
        }
        elseif( -Not ( $parent -is [bool] ) ) ## record the parent, if not disk, so that we can later check if the parent is orphaned (network interfaces and network security groups) and add it to the orphans collection
        {
            ForEach( $resource in $parent )
            {
                [pscustomobject]@{
                    Parent = $resource
                    Child  = $potentiallyUnassignedResource
                }
            }
        }
    }
})

## look through discovered parents to see if actually orphaned so can add their children to the orphan list if not already present
ForEach( $parent in $parents )
{
    if( ( $potentiallyUnassignedResources | Where-Object id -ieq $parent.parent ) -and ( -Not ( $potentiallyUnassignedResources | Where-Object id -ieq $parent.Child.Id ) ) )
    {
        $potentiallyUnassignedResources.Add( $parent.child )
    }
}

if( -Not $raw )
{
    [string]$message = "Found $($potentiallyUnassignedResources.Count) resources potentially orphaned in "

    if( $resourceGroupOnly -ieq 'no' )
    {
        [array]$resourceGroupsInvolved = @( $potentiallyUnassignedResources | Group-Object -Property 'Resource Group' )
        $message += "$($resourceGroupsInvolved.Count) resource groups"
   
        if( -Not $PSBoundParameters[ 'sortby' ] -or $sortby -imatch '^resource' ) ## convert "resourcegroup" to "resource group"
        {
            $sortby = 'Resource Group'
        }
    }
    else
    {
        $message += "resource group $resourceGroupName"
        if( $sortby -imatch '^resource' ) ## if only processing resource group then cannot sort on resource group
        {
            $sortby = 'type'
        }
    }

    $message += ". Output sorted by $sortby"

    Write-Output -InputObject $message
}

[array]$outputObjects = @( $potentiallyUnassignedResources | Select-Object -Property *,
    @{name='Created';expression={ $_.createdTime -as [datetime] }} ,
    @{name='Changed';expression={ $_.changedTime -as [datetime] }}  -ExcludeProperty id,tags,managedBy,sku,createdTime,changedTime )

if( $raw )
{
    $outputObjects
}
else
{
    $outputObjects | Sort-Object -Property $sortby | Format-Table -AutoSize  -Wrap
}


[datetime]$endTime = [datetime]::Now
Write-Verbose -Message "$($endtime.ToString( 'G')) : total run time $(($endTime - $startTime).TotalSeconds) seconds, making $($script:totalRequests) requests"

