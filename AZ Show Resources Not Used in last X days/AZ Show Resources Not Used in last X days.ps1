#require -version 3.0

<#
.SYNOPSIS
    Find Azure resources which have not been used in x days by looking for events for them in the activity log

.DESCRIPTION
    Using REST API calls

.PARAMETER azid
    The relative URI of the Azure VM
    
.PARAMETER AZtenantId
    The azure tenant ID
    
.PARAMETER daysback
    The number of days to search back in the logs

.PARAMETER resourceGroupOnly
    Only return results for the resource group containing the AZid
    
.PARAMETER VMOnly
    Only return results for the resource in the AZid

.NOTES
    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2021-10-30
    Updated:        2022-06-17  Guy Leech  Added code to deal with paging in results
#>

[CmdletBinding()]

Param
(
    [string]$AZid ,## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId ,
    [double]$daysBack = 30 ,
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

Function ConvertTo-Object
{
    [CmdletBinding()]
    Param
    (
        $Tables , ## TODO work with multiple tables - array of array ?
        $ExtraFields
    )
    ForEach( $table in $tables )
    {
        ForEach( $row in $table.rows )
        {
            [hashtable]$result = @{ TableName = $table.Name }
            if( $ExtraFields -and $ExtraFields.Count -gt 0 )
            {
                $result += $ExtraFields
            }
            For( [int]$index = 0 ; $index -lt $row.Count ; $index++ )
            {
                if( $table.columns[ $index ].type -ieq 'long' )
                {
                    $result.Add( $table.columns[ $index ].name , $row[ $index ] -as [long] )
                }
                elseif( $table.columns[ $index ].type -ieq 'datetime' )
                {
                    $result.Add( $table.columns[ $index ].name , $row[ $index ] -as [datetime] )
                }
                else
                {
                    if( $table.columns[ $index ].type -ine 'string' )
                    {
                        Write-Warning -Message "Unimplemented type $($table.columns[ $index ].type), treating as string"
                    }
                    $result.Add( $table.columns[ $index ].name , $row[ $index ] )
                }
            }
            [pscustomobject]$result
        }
    }
}

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

## https://docs.microsoft.com/en-us/rest/api/monitor/activity-logs/list

[datetime]$startFrom = (Get-Date).AddDays( -$daysBack )
[string]$filter = "eventTimestamp ge '$(Get-Date -Date $startFrom -Format s)'"
if( $resourceGroupOnly -eq 'yes' )
{
    $filter = "$filter and resourceGroupName eq '$resourceGroupName'"
}

## https://docs.microsoft.com/en-us/rest/api/resources/resources/list-by-resource-group
[string]$resourcesURL = $null
if( $resourceGroupOnly -ieq 'yes' )
{
    $resourcesURL = "resourceGroups/$resourceGroupName"
}
## else ## will be for the subscription

[array]$allResources = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/$resourcesURL/resources`?`$expand=createdTime,changedTime,lastusedTime&api-version=$resourceManagementApiVersion" -retries 2 | Where-Object type -NotIn $excludedResourceTypes )

[array]$allevents = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/providers/Microsoft.Insights/eventtypes/management/values`?api-version=$insightsApiVersion&`$filter=$filter" -retries 2 )

Write-Verbose -Message "Got $($allevents.Count) events in total"

## produce hash table keyed on resource id
[hashtable]$resourcesInActivityLog = @{}

[hashtable]$hostpools = @{}
[array]$logSinsights = @()

<# ## Basic level of "used" for now

$allResources.Where( { $_.type -ieq 'Microsoft.DesktopVirtualization/hostpools' } ).ForEach( { $hostpools.Add( $_.name , $_ ) } )

Write-Verbose -Message "Got $($hostpools.Count) hostpools"

[hashtable]$sessionhostsUsed = @{} ## New-Object -TypeName System.Collections.Generic.List[object]

if( $hostpools.Count -gt 0 )
{
    ## if we have any AVD, see if we have Log Insights
    [array]$workspaces = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.OperationalInsights/workspaces?api-version=$OperationalInsightsApiVersion" -retries 2 )

    if( -Not ( $loganalyticsBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID -scope 'https://api.loganalytics.io' ) )
    {
        Write-Warning -Message "Unable to get log analytics berer token for $($azSPCredentials.spCreds)"
    }
    else
    {
        ForEach( $workspace in $workspaces )
        {
            if( $workspace.properties.retentionInDays -lt $daysBack )
            {
                Write-Warning -Message "Days retention is $($workspace.properties.retentionInDays) in logs `"$($workspace.name)`" is less than days back requested of $daysBack so may miss some AVD sessions"
            }
            ## $usages = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.OperationalInsights/workspaces/$($workspace.Name)/usages?api-version=$OperationalInsightsApiVersion" -propertyToReturn $null -retries 2 )
            $query = "WVDConnections | where TimeGenerated > ago($($daysBack)d) and State == `"Connected`""

            ## https://docs.microsoft.com/en-us/rest/api/loganalytics/dataaccess/query/get?tabs=HTTP
            $queryResult = Invoke-AzureRestMethod -BearerToken $loganalyticsBearerToken -uri "https://api.loganalytics.io/v1/workspaces/$($workspace.properties.customerId)/query?query=$query" -propertyToReturn $null -retries 2
            if( $null -ne $queryResult -and $queryResult.PSObject.Properties[ 'tables' ] )
            {
                ## sessionhostname that comes back is fqdn but we need simple name so group on that so we can look it up
                $sessionhostsUsed += ConvertTo-Object -Tables $queryResult.tables -ExtraFields @{ Workspace = $workspace } | Select-Object -Property *,@{n='hostname';e={ ($_.SessionHostName -split '\.')[0] }} | Group-Object -Property hostName -AsHashTable -AsString
                Write-Verbose -Message "Got $($sessionhostsUsed.count) session hosts that have been used since $(Get-Date -Date $startFrom -Format G)"
            }
        }
    }
    ## iterate over all host pools to get session hosts in them and thenlook them up in the log analytics data to see if they have been logged on to or not so we can mark as such in output
    [hashtable]$allSessionHosts = @{}

    ForEach( $hostpool in $hostpools.GetEnumerator() )
    {
        [array]$sessionHosts = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.DesktopVirtualization/hostPools/$($hostpool.name)/sessionHosts?api-version=$desktopVirtualisationApiVersion" -retries 2 )
        if( $null -eq $sessionHosts -or $sessionHosts.Count -eq 0 )
        {
            Write-Warning -Message "Host pool `"$($hostpool.name)`" contains no session hosts"
        }
        else
        {
            ForEach( $sessionHost in $sessionHosts )
            {
                ## session host is hostpool/sessionhost.fqdn
                [string]$sessionHostname = ($($sessionHost.name -split '/')[-1] -split '\.')[0]
                $used = $sessionhostsUsed[ $sessionHostname ]

                $allSessionHosts.Add( $sessionHostname , $(if( $used ) { 'Used' } else { 'Unused' }) )
            }
        }
    }
}
#>

ForEach( $event in $allevents )
{
    try
    {
        $resourcesInActivityLog.Add( $event.resourceId , $event.resourceType )
    }
    catch
    {
        ## already have it
    }
}

Write-Verbose -Message "Got $($resourcesInActivityLog.Count) resources from activity log and $($allResources.Count) resources in resource group $resourceGroupName"

[hashtable]$usedVMs = @{}
[hashtable]$resourceDetails = @{}
[hashtable]$vmDetails = @{}
[hashtable]$resourceGroups = @{}

[System.Collections.generic.list[object]]$unusedResources = @( ForEach( $resource in $allResources )
{
    if( -Not $resourcesInActivityLog[ $resource.id ] )
    {
        [bool]$include = $null -ne $resource.psobject.properties[ 'changedTime' ] -and $null -ne $resource.changedTime -and [datetime]$resource.changedTime -lt $startFrom

        if( $include )
        {
            ## last modified time on resource itself can be newer than we were told in the /resources query
            ## TODO do we need to throttle the requests ?
            if( $resourceDetail = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$($resource.id)" -retries 2 -newestApiVersion -propertyToReturn $null -type $resource.type )
            {
                if( $systemData = $resourceDetail | Select-Object -ExpandProperty systemData -ErrorAction SilentlyContinue )
                {
                    if( $systemData.PSObject.properties[ 'lastModifiedAt' ] -and $systemData.lastModifiedAt -and [datetime]$systemData.lastModifiedAt -ge $startFrom )
                    {
                        $include = $false
                    }
                }
                elseif( $properties = $resourceDetail | Select-Object -ExpandProperty properties -ErrorAction SilentlyContinue )
                {
                    if( $properties.PSObject.properties[ 'lastModifiedTime' ] -and $properties.lastModifiedTime -and [datetime]$properties.lastModifiedTime -ge $startFrom )
                    {
                        $include = $false
                    }
                }
                if( $resource.type -ieq 'Microsoft.Compute/virtualMachines' )
                {
                    $vmDetails.Add( $resource.Id , $resourceDetail ) ## will check these later for orphaned resources
                }
                $resourceDetails.Add( $resource.id , $resourceDetail ) ## use in 2nd pass
            }

            [string]$avd = $null

            if( $include )
            {
                ## what if VM has been running the whole time so no start/stop? Check if running now
                if( $resource.type -ieq 'Microsoft.Compute/virtualMachines' )
                {
                    ## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/instance-view
                    if( $null -ne ( $instanceView = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$($resource.id)/instanceView`?api-version=$computeApiVersion" -property $null) ) )
                    {
                        if( ( $line = ( $instanceview.Statuses.code -match 'PowerState/' )) -and ( $powerstate = ($line -split '/' , 2 )[-1] ))
                        {
                            if( $powerstate -ieq 'running' )
                            {
                                $include = $false
                                ##$usedVMs.Add( $resource.id , $instanceView ) ## disks could be useful in 2nd pass
                            }
                        }
                        else
                        {
                            Write-Warning -Message "Failed to determine if vm $($resource.Name) is powered up"
                        }
                    }
                    else
                    {
                        Write-Warning -Message "Failed to get state of VM $($resource.name)"
                    }
                    <#
                    $avd = $sessionhostsUsed[ $resource.name ]
                    if( $avd )
                    {
                        $include = $true ## trumps VM being running the whole period as AVD session host powered on without users is potentially wasteful
                    }
                    #>
                }

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
                        Add-Member -InputObject $resource -MemberType NoteProperty -Name 'Resource Group' -Value $thisResourceGroupName -PassThru
                    }
                    else
                    {
                        Write-Warning -Message "Failed to determine resource group from $($resource.id)"
                    }
                }

                if( $include )
                {
                    $resource
                    ##Add-Member -InputObject $resource -MemberType NoteProperty -Name AVD -Value $avd -PassThru
                }
            }

            if( -Not $include )
            {
                Write-Verbose -Message "`tUsed : $($resource.id)"
            }
        }
    }
})

[hashtable]$VMsNSGs = @{}

## TODO for all network interfaces, get the network security groups and virtual machine(s) for each so we can check on NSGs later which are potentially unused
ForEach( $nic in $resourceDetails.GetEnumerator().Where( { $_.value.PSObject.Properties[ 'type' ] -and $_.value.type -eq 'Microsoft.Network/networkInterfaces' } ))
{
    ## TODO could there be more than one of either?
    if( ( $NSG = $nic.value.properties | Select-Object -ExpandProperty networkSecurityGroup -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id ) `
        -and ( $VM = $nic.value.properties | Select-Object -ExpandProperty virtualMachine -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id ) )
    {
        try
        {
            $VMsNSGs.Add( $NSG , ([System.Collections.Generic.List[object]]@( $VM ) ))
        }
        catch
        {
            $VMsNSGs[ $NSG ].Add( $VM )
        }
    }
}

Write-Verbose -Message "Got $($VMsNSGs.Count) network security groups"

## second pass if Microsoft.Compute/disks or Microsoft.Network/networkinterfaces or Microsoft.Network/networkSecurityGroups then check parent VM (if there is one) as we will have it in the unused collection
## cache VM details so if disk is for a machine we already have, we don't need to get it from Azure again
For( [int]$index = $unusedResources.Count -1 ; $index -ge 0 ; $index-- )
{
    $potentionallyUnusedResource = $unusedResources[ $index ]
    [bool]$remove = $false
    [bool]$orphaned = $false
    [bool]$childResource = $false
    $parentVM = $null

    if( $potentionallyUnusedResource.type -eq 'Microsoft.Network/networkInterfaces' )
    {
        $childResource = $true
        $parentVM = $resourceDetails[ $potentionallyUnusedResource.id ] | Select-Object -ExpandProperty properties | Select-Object -ExpandProperty virtualMachine -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id
    }
    elseif( $potentionallyUnusedResource.type -eq 'Microsoft.Compute/disks' )
    {
        $childResource = $true
        $parentVM = $resourceDetails[ $potentionallyUnusedResource.id ] | Select-Object -ExpandProperty managedBy -ErrorAction SilentlyContinue
    }
    elseif( $potentionallyUnusedResource.type -eq 'Microsoft.Network/networkSecurityGroups' )
    {
       if( $VMs = $VMsNSGs[ $potentionallyUnusedResource.id ] )
       {
            ## not orphaned but need to see if any of the VMs using it are in use
            ForEach( $VM in $VMs )
            {
                if( -Not ( $unused = $unusedResources.Where( { $_.id -ieq $VM } ) )) 
                {
                    $remove = $true
                    break
                }
            }
        }
        else
        {
            $orphaned = $true
        }
    }

    ## TODO Microsoft.DesktopVirtualization/applicationgroups Microsoft.DesktopVirtualization/hostpools Microsoft.Network/publicIPAddresses Microsoft.Compute/availabilitySets Microsoft.Network/applicationSecurityGroups 
    ## Need to see if associated session hosts have had logons and mark as such

    if( $childResource )
    {
        if( $parentVM )
        {
            ## now see if in our unused VMs and if not we remove from this list as must've been in use
            if( -Not ( $unused = $unusedResources.Where( { $_.id -ieq $parentVM } ) ))
            {
                $remove = $true
            }
        }
        else
        {
           $orphaned = $true ## could it be in a different resource group?
        }
    }
    if( $remove )
    {
        $unusedResources.RemoveAt( $index )
    }
    else
    {
        Add-Member -InputObject $potentionallyUnusedResource -Force -MemberType NoteProperty -Name 'Orphaned' -Value $(if( $orphaned ) { 'Possibly' } else { 'No' } )
    }
}

[string]$message = "Found $($unusedResources.Count) resources potentially not used since $(Get-Date -Format G -Date $startFrom) in "

if( $resourceGroupOnly -ieq 'no' )
{
    $message += "$($resourceGroups.Count) resource groups"
    if( -Not $PSBoundParameters[ 'sortby' ] -or $sortby -match 'resource' ) ## deal with spaces 
    {
        $sortby = 'Resource Group'
    }
}
else
{
    $message += "resource group $resourceGroupName"
}

$message += ". Output sorted by $sortby"

Write-Verbose -Message "Version cache entries is $($script:apiversionCache.Count), hits were $script:versionCacheHits"

Write-Output -InputObject $message

$unusedResources | Select-Object -Property *,@{n='Parent';e={$_.managedby -replace '^/subscriptions/[^/]+/' }} -ExcludeProperty id,tags,managedBy | Sort-Object -Property $sortby | Format-Table -AutoSize  -Wrap

