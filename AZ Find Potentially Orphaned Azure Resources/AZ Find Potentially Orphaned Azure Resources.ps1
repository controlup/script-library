    #require -version 3.0

<#
.SYNOPSIS
    Find Azure resources which should probably be assigned to resources but are not, eg disks, network interfaces

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

.NOTES
    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2022-09-16
    Updated:        2022-06-17  Guy Leech  Added public IP addresses and changed mechanisms for determining unattached
#>

[CmdletBinding()]

Param
(
    [string]$AZid ,## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId ,
    [ValidateSet('Yes','No')]
    [string]$resourceGroupOnly = 'Yes',
    [string]$sortby = 'type'
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

[string]$computeApiVersion = '2021-07-01'
[string]$insightsApiVersion = '2015-04-01'
[string]$resourceManagementApiVersion = '2021-04-01'
[string]$OperationalInsightsApiVersion = '2022-10-01'
[string]$desktopVirtualisationApiVersion = '2021-07-12'
[string]$baseURL = 'https://management.azure.com'
[string]$credentialType = 'Azure'
[hashtable]$script:apiversionCache = @{}

Write-Verbose -Message "AZid is $AZid"

#region AzureFunctions

Function Get-CurrentLineNumber
{ 
    $MyInvocation.ScriptLineNumber 
}

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

    Invoke-RestMethod @invokeRestMethodParams | Select-Object -ExpandProperty access_token -ErrorAction SilentlyContinue
}

[hashtable]$script:cachedApiVersions = @{}
[int]$script:versionCacheHits = 0

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
        [switch]$newestApiVersion ,
        [switch]$oldestApiVersion ,
        [string]$type , ## help us with looking up API versions & caching
        [int]$retries = 0 ,
        [int]$retryIntervalMilliseconds = 2500
    )

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
    
    [bool]$correctedApiVersion = $false

    if( $newestApiVersion -or $oldestApiVersion )
    {
        if( $uri -match '\?api\-version=20\d\d-\d\d-\d\d' )
        {
            Write-Warning -Message "Uri $uri already has an api version"
            $correctedApiVersion = $true
        }
        else
        {
            [string]$apiversion = '42' ## force error which will return list of valid api versions
            ## see if we have cached entry already for this provider and use that to save a REST call
            if( [string]::IsNullOrEmpty( $type ) -and $uri -match '\w/providers/([^/]+/[^/]+)\w' )
            {
                $type = $Matches[ 1 ]
            }

            if( -Not [string]::IsNullOrEmpty( $type ) -and ( $cached = $script:apiversionCache[ $type ] ))
            {
                $correctedApiVersion = $true
                $apiversion = $cached
                $script:versionCacheHits++
            }
            $invokeRestMethodParams.uri += "?api-version=$apiversion"
        }
    }
    else
    {
        $correctedApiVersion = $true
    }

    [string]$lastURI = $null

    ## cope with pagination where get 100 results at a time
    do
    {
        [datetime]$requestStartTime = [datetime]::Now
        $thisretry = $retries
        $error.Clear()
        $exception = $null
        do
        {
            $exception = $null
            $result = $null

            try
            {
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
                if( ( $_ | Select-Object -ExpandProperty ErrorDetails | Select-Object -ExpandProperty Message | ConvertFrom-Json | Select-Object -ExpandProperty error | Select-Object -ExpandProperty message) -match 'for type ''([^'']+)''\. The supported api-versions are ''([^'']+)''')
                {
                    [string]$requestType = $Matches[ 1 ]
                    [string[]]$apiVersionList = $Matches[2] -split ',\s?'
                    ## 2021-12-01
                    [datetime[]]$apiversions =@( $apiVersionList | Where-Object { $_ -notmatch '(preview|beta|alpha)$' } ) | Sort-Object

                    
                    if( $correctedApiVersion )
                    {
                        ## we have already tried to correct the api version but sometimes there is a sub-provider that we can't easily determine
                        ## https://management.azure.com//subscriptions/<subsriptionid>/resourceGroups/WVD/providers/Microsoft.Automation/automationAccounts/automation033926z?api-version
                        ## and
                        ## https://management.azure.com//subscriptions/<subscriptionid>/resourceGroups/WVD/providers/Microsoft.Automation/automationAccounts/automation033926z/runbooks/inputValidationRunbook?
                        ## where latter provider is AutomationAccounts/
                    }
                    [int]$apiVersionIndex = $(if( $newestApiVersion ) { -1 } else { 0 } ) ## pick first or last version from sorted array
                    [string]$apiversion = $(Get-Date -Date $apiversions[ $apiVersionIndex ] -Format 'yyyy-MM-dd') 
                    $invokeRestMethodParams.uri = "$uri`?api-version=$apiversion"
           
                    ## seems to be too simplistic eg /WVD/providers/Microsoft.Automation/automationAccounts/automation033926z/runbooks/ is type 'automationAccounts/runbooks' not '/Microsoft.Automation/automationAccounts'
                    ##if( $uri -match  '\w/providers/([^/]+/[^/]+)\w' )
                    if( $true )
                    {
                        try
                        {
                            $script:apiversionCache.Add( $type , $apiversion )
                        }
                        catch
                        {
                            ## already have it
                            $null
                        }
                    }

                    $correctedApiVersion = $true
                    $exception = $true ## so we don't break out of loop
                    $thisretry++ ## don't count this as a retry since was not a proper query
                    $error.Clear()
                }
                else
                {
                    $exception = $_
                    if( $thisretry -ge 1 ) ## do not sleep if no retries requested or this was the last retry
                    {
                        Start-Sleep -Milliseconds $retryIntervalMilliseconds
                    }
                }
            }
            if( -not $exception )
            {
                break
            }
        } while( --$thisretry -ge 0)

        ## $result -eq $null does not mean there was an exception so we need to track that separately to know whether to throw an exception here
        if( $exception )
        {
            ## last call gave an exception
            Throw "Exception $($exception.ToString()) originally occurred on line number $($exception.InvocationInfo.ScriptLineNumber)"
        }
        elseif( $error.Count -gt 0 )
        {
            Write-Warning -Message "Transient errors on request $($invokeRestMethodParams.Uri) - $($error.ToString() | ConvertFrom-Json | Select-Object -ExpandProperty error|Select-Object -ExpandProperty message)"
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

$azSPCredentials = $null
$azSPCredentials = Get-AzSPStoredCredentials -system $credentialType -tenantId $AZtenantId

If ( -Not $azSPCredentials )
{
    Exit 1 ## will already have output error
}

# Sign in to Azure with the Service Principal retrieved from the credentials file and retrieve the bearer token
Write-Verbose -Message "Authenticating to tenant $($azSPCredentials.tenantID) as $($azSPCredentials.spCreds.Username)"
if( -Not ( $azBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID -scope $baseURL ) )
{
    Throw "Failed to get Azure bearer token"
}

[string]$vmName = ($AZid -split '/')[-1]
    
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

[array]$allResources = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/$resourcesURL/resources`?`$expand=createdTime,changedTime,lastusedTime&api-version=$resourceManagementApiVersion" -retries 2 | Where-Object type -NotIn $excludedResourceTypes )

[hashtable]$resourceDetails = @{}
[hashtable]$resourceGroups = @{}

[string[]]$typesOfInterest = @( 'Microsoft.Network/networkInterfaces' , 'Microsoft.Compute/disks' , 'Microsoft.Network/networkSecurityGroups' , 'Microsoft.Network/publicIPAddresses' )

ForEach( $resource in $allResources )
{
    ## TODO do we need to throttle the requests ?
    ## only cache the resource types that could be orphaned
    if( $resource.type -in $typesOfInterest )
    {
        if( $resourceDetail = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$($resource.id)" -retries 2 -newestApiVersion -propertyToReturn $null -type $resource.type )
        {
            if( $resourceGroupOnly -ieq 'no')
            {
                if( $resource.id -match '\bsubscriptions/([^/]+)/resourceGroups/([^/]+)/providers/Microsoft\.' )
                {
                    [string]$thisResourceGroupName = $Matches[2]

                    try
                    {
                        $resourceGroups.Add( $thisResourceGroupName , $resource )
                    }
                    catch
                    {
                        ## already got it
                    }
                    Add-Member -InputObject $resource -MemberType NoteProperty -Name 'Resource Group' -Value $thisResourceGroupName
                }
                else
                {
                    Write-Warning -Message "Failed to determine resource group from $($resource.id)"
                }
            }
        
            $resourceDetails.Add( $resource.id , $resourceDetail ) ## use in 2nd pass
        }
        else
        {
            Write-Warning -Message "Failed to get details for resource $($resource.id)"
        }
    }
    ## else not an orphanable type
}

## second pass if Microsoft.Compute/disks or Microsoft.Network/networkinterfaces or Microsoft.Network/networkSecurityGroups then check parent VM (if there is one) as we will have it in the unused collection
## cache VM details so if disk is for a machine we already have, we don't need to get it from Azure again
##[array]$potentiallyUnassignedResources = @( ForEach( $item in $resourceDetails.GetEnumerator() )
[array]$potentiallyUnassignedResources = @( ForEach( $potentiallyUnassignedResource in $allResources )
{
    ##$potentiallyUnassignedResource = $item.Value
    [bool]$orphaned = $false
    [bool]$childResource = $false
    $parentVM = $null

    if( $potentiallyUnassignedResource.type -ieq 'Microsoft.Network/networkInterfaces' )
    {
        $childResource = $true
        if( $resourceDetail = $resourceDetails[ $potentiallyUnassignedResource.id ] )
        {
            $parentVM = $resourceDetail.properties | Select-Object -ExpandProperty virtualMachine -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id -ErrorAction SilentlyContinue
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
        ##if( ( $parentVM = $resourceDetails[ $potentiallyUnassignedResource.id ] | Select-Object -ExpandProperty managedBy -ErrorAction SilentlyContinue ) -and ( $resourceDetail = $resourceDetails[ $potentiallyUnassignedResource.id ] ) -and $resourceDetail.properties.DiskState -ieq 'Unattached' )
        if( $resourceDetail = $resourceDetails[ $potentiallyUnassignedResource.id ] )
        {
            $parentVM = ( $resourceDetail.properties.DiskState -ine 'Unattached' )
        }
        else
        {
            Write-Warning -Message "Unable to get details for disk `"$($potentiallyUnassignedResource.id)`""
        }    
    }
    elseif( $potentiallyUnassignedResource.type -ieq 'Microsoft.Network/networkSecurityGroups' )
    {
       ## $orphaned = -Not $VMsNSGs[ $potentiallyUnassignedResource.id ]
       $childResource = $true
       if( $resourceDetail = $resourceDetails[ $potentiallyUnassignedResource.id ] )
       {
            if( $resourceDetail.properties.psobject.properties[ 'subnets' ] -and $resourceDetail.properties.subnets.Count -gt 0 )
            {
                $parentVM = $true ## don't actually need it
            }
            elseif( $resourceDetail.properties.psobject.properties[ 'networkInterfaces' ] -and $resourceDetail.properties.networkInterfaces.Count -gt 0 )
            {
                $parentVM = $true ## don't actually need it
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
            $parentVM = ( $resourceDetail.properties.psobject.properties[ 'ipConfiguration' ] -and $null -ne $resourceDetail.properties.ipConfiguration )
        }
        else
        {
            Write-Warning -Message "Unable to get details for public IP address `"$($potentiallyUnassignedResource.id)`""
        }
    }

    if( $childResource -and -Not $parentVM )
    {
        #$orphaned = $true ## could it be in a different resource group?
        $potentiallyUnassignedResource
    }

    #Add-Member -InputObject $potentiallyUnassignedResource -MemberType NoteProperty -Name Orphaned -Value $orphaned
})

[string]$message = "Found $($potentiallyUnassignedResources.Count) resources potentially orphaned in "

if( $resourceGroupOnly -ieq 'no' )
{
    $message += "$($resourceGroups.Count) resource groups"
   
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

Write-Verbose -Message "Version cache entries is $($script:apiversionCache.Count), hits were $script:versionCacheHits"

Write-Output -InputObject $message

$potentiallyUnassignedResources | Select-Object -Property * -ExcludeProperty id,tags,managedBy,sku | Sort-Object -Property $sortby | Format-Table -AutoSize  -Wrap

