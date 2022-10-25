#require -version 3.0

<#
.SYNOPSIS
    Get Azure logs from the number of days back to the present, optionally just for this VM or resource group

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
    [double]$daysBack = 1 ,
    [ValidateSet('Yes','No')]
    [string]$resourceGroupOnly = 'Yes',
    [ValidateSet('Yes','No')]
    [string]$VMOnly = 'No'
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 250
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

[string]$computeApiVersion = '2021-07-01'
[string]$insightsApiVersion = '2015-04-01'
[string]$baseURL = 'https://management.azure.com'
[string]$credentialType = 'Azure'

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
                $exception = $_
                if( $thisretry -ge 1 ) ## do not sleep if no retries requested or this was the last retry
                {
                    Start-Sleep -Milliseconds $retryIntervalMilliseconds
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
            Throw $exception
        }
        elseif( $error.Count -gt 0 )
        {
            Write-Warning -Message "Transient errors on request $($invokeRestMethodParams.Uri) - $($error.ToString() | ConvertFrom-Json | Select-Object -ExpandProperty error|Select-Object -ExpandProperty message)"
        }

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

## subscriptions/58ffa3cb-2f63-4242-a06d-deadbeef/resourceGroups/WVD/providers/Microsoft.Compute/virtualMachines/MYMACHINE-0
if( $AZid -match '\bsubscriptions/([^/]+)/resourceGroups/([^/]+)/providers/Microsoft\.' )
{
    $subscriptionId = $Matches[1]
    $resourceGroupName = $Matches[2]
}
else
{
    Throw "Failed to parse subscription id and resource group from $AZid"
}

## get service principals so we can map caller to meaningful name via MS Graph
$graphBearerToken = $null
[hashtable]$servicePrincipals = @{}

try
{
    if( -Not ( $graphBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID -scope 'https://graph.microsoft.com' ) )
    {
        Write-Warning -Message "Failed to get Microsoft Graph bearer token"
    }
    else
    {
        [array]$allServicePrincipals = @(  Invoke-AzureRestMethod -BearerToken $graphBearerToken -uri "https://graph.microsoft.com/v1.0/servicePrincipals" -method GET )
        Write-Verbose -Message "Got $($allServicePrincipals.Count) service principals"
        ForEach( $servicePrincipal in $allServicePrincipals )
        {
            $servicePrincipals.Add( $servicePrincipal.Id , $servicePrincipal.DisplayName )
        }
    }
}
catch
{
    Write-Warning -Message "Unable to get service principal list via Microsoft Graph. Due to this some GUIDs cannot be resolved to an account name. Please ensure the service principal you use for this script has Reader access to Azure Active Directory."
}

## https://docs.microsoft.com/en-us/rest/api/monitor/activity-logs/list

[datetime]$startFrom = (Get-Date).AddDays( -$daysBack )
[string]$filter = "eventTimestamp ge '$(Get-Date -Date $startFrom -Format s)'"
if( $VMOnly -eq 'yes' )
{
    $filter = "$filter and resourceUri eq '$AZid'"
}
elseif( $resourceGroupOnly -eq 'yes' )
{
    $filter = "$filter and resourceGroupName eq '$resourceGroupName'"
}

[array]$allevents = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/providers/Microsoft.Insights/eventtypes/management/values`?api-version=$insightsApiVersion&`$filter=$filter" -propertyToReturn 'value' -retries 2 )

Write-Verbose -Message "Got $($allevents.Count) events in total"

if( -Not ( $logs = @( $allevents | Group-Object -Property CorrelationId) ) -or -Not $logs.Count )
{
    Write-Output -InputObject "No correlated events found since $(Get-Date -Date (Get-Date).AddDays( -$daysBack ) -Format G)"
}
else
{
    Write-Verbose -Message "Got $($logs.Count) correlated log entries"
    [array]$results = @( $logs | ForEach-Object `
    {
        ## correlate the start and end so we can show a duration - there could be 3 events - status stated, accepted, succeeded/failed
        $activity = $_.Group
        ForEach( $resourceGroup in ($activity | Group-Object -Property resourceid ) )
        {
            $events = $resourceGroup.Group
            $start = $events | Where-Object { $_.eventName.value -eq 'BeginRequest' -and $_.status.value -eq 'Started' }
            ## may be 2 events where the properties are different (and contain a detailed error message which we could extract). Sort on substatus so we hopefully get the non-null one
            if( $failed = $events | Where-Object { $_.eventName.value -eq 'EndRequest' -and $_.status.value -eq 'Failed' } )
            {
                $end = $failed
            }
            else
            {
                $end = $events | Where-Object { $_.eventName.value -eq 'EndRequest' -and $_.status.value -eq 'Succeeded' } | Sort-Object -Property eventTimestamp | Select-Object -Last 1
            }
            $duration = $null
            $endtime = $null
            $starttime = $null
            $substatus = $null
            $status = 'Unknown'
            $detail = $null
            try
            {
                $starttime = $start.eventTimestamp -as [datetime]
                $endtime = ($end.eventTimestamp | Select-Object -Last 1) -as [datetime]
                $duration = ( $endtime - $starttime ).TotalSeconds
                $status = $end.status.localizedValue
                ##$substatus = $end.substatus.localizedValue | Where-Object { $null -ne $_ } | Select-Object -First 1
                if( $failed )
                {
                    if( -Not ( $detail = $failed.properties.statusMessage | ConvertFrom-Json -ErrorAction SilentlyContinue | Select-Object -ErrorAction SilentlyContinue -ExpandProperty error | Select-Object -ExpandProperty details | Select-Object -ExpandProperty message ))
                    {
                        $detail = $failed.properties.statusMessage | ConvertFrom-Json -ErrorAction SilentlyContinue | Select-Object -ErrorAction SilentlyContinue -ExpandProperty error | Select-Object -ExpandProperty message
                    }
                }
                <#
                elseif( $end )
                {
                    $detail = $end.properties.message
                }
                #>
            }
            catch
            {
                ## lazy way to ensure we have decent properties
                $null
            }
            if( $endtime -and $starttime -and $duration -ge 0 )
            {
                [string]$caller = $events[0].caller
                if( $caller -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$' ) ## if it's a GUID, see if we can resolve it
                {
                    if( $thisServiceprincipal = $servicePrincipals[ $caller ] )
                    {
                        $caller = $thisServiceprincipal
                    }
                }
                [pscustomobject]@{
                    'User' = $caller
                    'Operation' = $(if( $start.PSobject.Properties[ 'OperationName' ] ) { $start.OperationName.localizedValue } )
                    'Resource' = Split-Path -Path $resourceGroup.Name -Leaf
                    'ResourceGroup' = $(if( $events[0].PSObject.Properties[ 'resourceGroupName' ] ) { $events[0].resourceGroupName.ToUpper() })
                    'Start' = $starttime
                    'End' = $endtime
                    'Duration (s)' = $duration
                    'Level' = $end | Select-Object -ExpandProperty Level -ErrorAction SilentlyContinue ## in case end is null. Can't use start as may be "informational" but will have changed to "error" if errored
                    'Status' = $status
                    ##'Sub Status' = $substatus
                    'Detail' = $detail
                }
            }
        }
    } )
    [string]$title = "$($logs.Count) events found since $(Get-Date -Date (Get-Date).AddDays( -$daysBack ) -Format G)"
    if( $VMOnly -ieq 'yes' )
    {
        $title += " for $vmName"
    }
    elseif( $resourceGroupOnly -ieq 'yes' )
    {
        $title += " for resource group `"$resourceGroupName`""
    }

    $results | Sort-Object -Property Start -Descending
    if( $chosen -and $chosen.Count -gt 0 )
    {
        $chosen | Set-Clipboard
    }
}

