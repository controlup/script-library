#require -version 3.0

<#
.SYNOPSIS
    Get Azure logs and windows event logs for the specified AVD VM to show how long it took to boot

.DESCRIPTION
    Using REST API calls

.PARAMETER azid
    The relative URI of the Azure VM

.PARAMETER AZtenantId
    The azure tenant ID

.NOTES
    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2021-10-30
    Updated:        2022-06-17  Guy Leech  Added code to deal with paging in results
                    2023-01-20  Guy Leech  Removed logon success event, tidied up output, added total duration
                    2023-01-25  Added log analytics query to get AVD launch request time
                    2023-01-26  Fix for ignoring log insights connection start event time
                    2023-01-31  Changed detection f first session connection as was measuring first succesful one
                                Fixed bug where previous reboot health events were being retrieved and skewing results
                    2023-02-13  Added UTC message
                    2023-03-24  Workaround for different log retention periods causing fatal exception
                                Select specific properties only to be returned by activity log query
                    2023-03-27  Moved code for getting logs to after get VM event logs so can be more precise about time/date
                                Added code to deal with local time in VM being different to where script run
                    2023-03-28  Added -TimeoutSec argument capability to Invoke-RestMethod after Log Analytics hangs
                    2023-03-29  Added code to detect and warn when health events appear to be missing
                    2023-03-30  Try/catch around decoding of json of remote event logs as sometimes gives exception
                    2023-04-26  Added code to deal with no resource health events returned, eg booted days ago
                    2023-04-28  Show difference in hours if health events occur before connection attempt
#>

[CmdletBinding()]

Param
(
    [string]$AZid ,## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId ,
    [int]$maxWaitTimeSeconds = 120 ,
    [decimal]$maxRequestWaitTimeSeconds = 90
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 250
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    try
    {
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions
    }
    catch
    {
        ## Nothing we can do but shouldn't cause script to end
    }
}

[string]$computeApiVersion  = '2022-08-01'
[string]$insightsApiVersion = '2015-04-01'
[string]$resourceHealthApiVersion = '2022-10-01'
[string]$avdApiVersion = '2021-07-12'
[string]$operationalInsightsApiVersion = '2022-10-01'
[string]$baseURL = 'https://management.azure.com'
[string]$credentialType = 'Azure'
[string]$outputTimeFormat = 'HH:mm:ss.fff'
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
        [string] $scope ,

        [int]$timeOutSeconds
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

    if( $timeOutSeconds -gt 0 )
    {
        $invokeRestMethodParams.Add( 'TimeoutSec' , $timeOutSeconds )
    }

    Invoke-RestMethod @invokeRestMethodParams | Select-Object -ExpandProperty access_token -ErrorAction SilentlyContinue ## return
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
        [switch]$newestApiVersion ,
        [switch]$oldestApiVersion ,
        [switch]$rawException ,
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

    if( $maxRequestWaitTimeSeconds -gt 0 )
    {
        $invokeRestMethodParams.Add( 'TimeoutSec' , $maxRequestWaitTimeSeconds ) 
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
                Write-Verbose -Message "$($requestStartTime.ToString( 'G' )): sending $($invokeRestMethodParams.Method) to $($invokeRestMethodParams.Uri)"
                if( $norest )
                {
                    $result = Invoke-WebRequest @invokeRestMethodParams
                }
                else
                {
                    $result = Invoke-RestMethod @invokeRestMethodParams
                }

                Write-Verbose -Message "$([datetime]::Now.ToString( 'G' )): received response to $($invokeRestMethodParams.Method) from $($invokeRestMethodParams.Uri)"
            }
            catch
            {
                if( ( $newestApiVersion -or $oldestApiVersion ) -and ( $_ | Select-Object -ExpandProperty ErrorDetails | Select-Object -ExpandProperty Message | ConvertFrom-Json | Select-Object -ExpandProperty error | Select-Object -ExpandProperty message) -match 'for type ''([^'']+)''\. The supported api-versions are ''([^'']+)''')
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

                    try
                    {
                        $script:apiversionCache.Add( $type , $apiversion )
                    }
                    catch
                    {
                        ## already have it
                        $null
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
                        Write-Verbose -Message "$(Get-Date -Format G): retries $thisretry exception $_"
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
            if( $rawException )
            {
                Throw $exception
            }
            else ## hopefully more human readable & meaningful
            {
                Throw "Exception $($exception.ToString()) originally occurred on line number $($exception.InvocationInfo.ScriptLineNumber) for request $uri ($method)"
            }
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

function Wait-AsyncAzureOperation
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$asyncURI ,
        [Parameter(Mandatory=$true)]
        [string]$azBearerToken ,
        [string]$operation ,
        [string]$returnProperty ,
        [int]$maxWaitTimeSeconds = 0 ,
        [double]$sleepSeconds = 10.0
    )

    $return = $null
    [datetime]$startTime = [datetime]::Now
    [datetime]$endTime = [datetime]::MaxValue
    if( $maxWaitTimeSeconds -gt 0 )
    {
        $endTime = $startTime.AddSeconds( $maxWaitTimeSeconds )
    }

    $status = $null

    Write-Verbose -Message "$(Get-Date -Date $startTime -Format G): starting wait for operation $operation to complete via $asyncURI"

    do
    {
        $status = $null
        $status = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $asyncURI -property $null -method GET
        if( -Not $status -or $status.status -ine 'Succeeded')
        {   
            Write-Verbose -Message "`t$(Get-Date -Format G): waiting for script to complete or until $(Get-Date -Date $endTime -Format G) - status is $($status|Select-Object -ExpandProperty status), sleeping $($sleepSeconds)s"
            Start-Sleep -Seconds $sleepSeconds
        }
    } while ( $status -and $status.status -eq 'InProgress' -and [datetime]::Now -le $endTime)

    Write-Verbose -Message "$(Get-Date -Format G): finished wait for operation $operation to complete via $asyncURI"

    if( $status )
    {
        if( $status.status -eq 'InProgress' )
        {
            Write-Warning -Message "Timed out after $maxWaitTimeSeconds seconds waiting for completion"
        }
        elseif( $status.status -ne 'Succeeded' )
        {
            Write-Error -Message "Bad status $($status.status) from operation"
            <#
            $status.properties.output.value | Where-Object { $_.Code -ieq 'ComponentStatus/StdOut/succeeded' } | Select-Object -ExpandProperty Message

            if( $errors = $status.properties.output.value | Where-Object { $_.Code -imatch 'ComponentStatus/StdErr/' } | Select-Object -ExpandProperty Message -ErrorAction SilentlyContinue )
            {
                Write-Warning -Message "Errors: $errors"
            }
            #>
        }
        else
        {
            if( $PSBoundParameters[ 'returnProperty' ] )
            {
                $return = $status | Select-Object -ExpandProperty properties | Select-Object -ExpandProperty $returnProperty
            }
            else
            {
                $return = $true
            }
        }
    }
    else
    {
        Write-Warning -Message "Failed to get status from $asyncURI"
    }

    return $return
}

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

#endregion AzureFunctions

#region authentication

$azSPCredentials = $null
$azSPCredentials = Get-AzSPStoredCredentials -system $credentialType -tenantId $AZtenantId

If ( -Not $azSPCredentials )
{
    Exit 1 ## will already have output error
}

# Sign in to Azure with the Service Principal retrieved from the credentials file and retrieve the bearer token
Write-Verbose -Message "$([datetime]::Now.ToString( 'G' )) : authenticating to tenant $($azSPCredentials.tenantID) as $($azSPCredentials.spCreds.Username)"
if( -Not ( $azBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID -scope $baseURL -timeOutSeconds $maxRequestWaitTimeSeconds ) )
{
    Throw "Failed to get Azure bearer token"
}

Write-Verbose -Message "$([datetime]::Now.ToString( 'G' )) : authenticated"

[string]$vmName = ($AZid -split '/')[-1]

[string]$subscriptionId = $null
[string]$resourceGroupName = $null

## subscriptions/58ffa3cb-baff-b0ff-eeee-deadbeef/resourceGroups/WVD/providers/Microsoft.Compute/virtualMachines/GLMW10WVD-0
if( $AZid -match '\bsubscriptions/([^/]+)/resourceGroups/([^/]+)/providers/Microsoft\.' )
{
    $subscriptionId = $Matches[1]
    $resourceGroupName = $Matches[2]
}
else
{
    Throw "Failed to parse subscription id and resource group from $AZid"
}
#endregion authentication

## get instance view so we can get provisioning times so we know where to search in activity logs
## UPDATE: instance view provisioning times get updated if changes are made to VM so we cannot use it for start up time, we will have to use logs

## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/instance-view
[string]$instanceViewURI = "$baseURL/$azid/instanceView`?api-version=$computeApiVersion"
if( $null -eq ( $vmInstanceView = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $instanceViewURI -property $null ) )
{
    Throw "Failed to get VM instance view via $instanceViewURI : $_"
}

<#
code                        level displayStatus          time                             
----                        ----- -------------          ----                             
ProvisioningState/succeeded Info  Provisioning succeeded 2022-07-04T07:47:46.6809294+00:00
#>

if( ($powerState = ( $vminstanceview.Statuses.Where( {$_.code -match 'PowerState/' } )) ))
{
    if( ( $status = ($powerState.code -split '/')[-1] ))
    {
        if( $status -ne 'running' )
        {
            Throw "Power state of VM $VMname is $status, not running"
        }
    }
    else
    {
        Write-Warning -Message "Failed to get power state of VM $vmname from $($powerState | Select-object -ExpandProperty code)"
    }
}
else
{
    Write-Warning -Message "Failed to get VM $vmname status from $vmInstanceView"
}

## get virtual machine so we can get its id to match to a session host
if( -Not ( $virtualMachine = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid/?api-version=$computeApiVersion" -property $null ) )
{
    Throw "Failed to get VM for $azid"
}

[string]$subscriptionId = $null
[string]$resourceGroupName = $null
## subscriptions/baffa3cb-2f63-4242-a06d-badbadcebbf5/resourceGroups/WVD/providers/Microsoft.Compute/virtualMachines/GLMW10WVD-0
if( $AZid -match '\bsubscriptions/([^/]+)/resourceGroups/([^/]+)/providers/Microsoft\.' )
{
    $subscriptionId = $Matches[1]
    $resourceGroupName = $Matches[2]
}
else
{
    Throw "Failed to parse subscription id and resource group from `"$AZid`""
}

$sessionHost = $null
$parentHostPool = $null
$warnings = New-Object -TypeName System.Collections.Generic.List[string]

## https://learn.microsoft.com/en-us/rest/api/desktopvirtualization/host-pools/list?tabs=HTTP
[array]$hostpools = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/providers/Microsoft.DesktopVirtualization/hostPools?api-version=$avdApiVersion" )

Write-Verbose -Message "Got $($hostpools.Count) host pools"

ForEach( $hostpool in $hostpools )
{ 
    ## https://learn.microsoft.com/en-us/rest/api/desktopvirtualization/session-hosts/list?tabs=HTTP
    if( $sessionHost = ( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$($hostpool.id)/sessionHosts?api-version=$avdApiVersion" ) | Where-Object { $_.Properties.virtualMachineId -eq $virtualMachine.properties.vmid } )
    {
        $parentHostPool = $hostpool
        Write-Verbose -Message "Found vm in host spool $($parentHostPool.name)"
        if( -Not $parentHostPool.properties.startVMOnConnect )
        {
            Write-Warning -Message "Host pool $($parentHostPool.name) containing this VM is not set to start VM on connect"
        }
        break ## there can be only one
    }
}

$avdConnections = New-Object -TypeName System.Collections.Generic.List[object]

## see what we can get from log analytics, if available
if( $parentHostPool )
{
    ## if we have any AVD, see if we have Log Insights
    [array]$workspaces = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.OperationalInsights/workspaces?api-version=$OperationalInsightsApiVersion" -retries 2 )

    if( -Not ( $loganalyticsBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID -scope 'https://api.loganalytics.io' ) )
    {
        Write-Warning -Message "Unable to get log analytics bearer token for $($azSPCredentials.spCreds)"
    }
    else
    {
        ## name is qualified with hostpool/ so strip that off
        [string]$query = "WVDConnections | where SessionHostName == `"$(($sessionHost.Name -split '/')[-1])`"" ##TimeGenerated > ago($($daysBack)d)"

        ForEach( $workspace in $workspaces )
        {
            ##$usages = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.OperationalInsights/workspaces/$($workspace.Name)/usages?api-version=$OperationalInsightsApiVersion" -propertyToReturn $null -retries 2 )
             ## https://docs.microsoft.com/en-us/rest/api/loganalytics/dataaccess/query/get?tabs=HTTP
            $queryResult = Invoke-AzureRestMethod -BearerToken $loganalyticsBearerToken -uri "https://api.loganalytics.io/v1/workspaces/$($workspace.properties.customerId)/query?query=$query" -propertyToReturn $null -retries 2
            if( $null -ne $queryResult -and $queryResult.PSObject.Properties[ 'tables' ] )
            {
                $avdConnections += @( ConvertTo-Object -Tables $queryResult.tables -ExtraFields @{ Workspace = $workspace } )
            }
        }
        $avdConnections = @( $avdConnections | Sort-Object -Property TimeGenerated -Descending )
    }
}

## https://learn.microsoft.com/en-us/rest/api/resourcehealth/2022-10-01/events/list-by-single-resource?tabs=HTTP
## there may be multiple allocate/start events so we want to get the most recent - there must be an easier way than this!
[hashtable]$groupedHealthEvents = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid/providers/Microsoft.ResourceHealth/events?api-version=$resourceHealthApiVersion" -propertyToReturn 'value' `
    | Select-Object -ExpandProperty properties -ErrorAction SilentlyContinue | Where-Object title -NotMatch 'deallocat|stop|reboot' | Select-Object -Property *,@{n='TimeOccurred';e={ $_.ImpactStartTime -as [datetime]}} | Sort-Object -Property TimeOccurred -Descending | Group-Object -Property title -AsHashTable

 [array]$healthEvents = @()

if( $null -ne $groupedHealthEvents -and $groupedHealthEvents.Count -gt 0 )
{
    $healthEvents = @( $groupedHealthEvents.GetEnumerator() | ForEach-Object `
        {
            $_.value | Select-Object -First 1
        } | Sort-Object -Property TimeOccurred )
}
else
{
    $warnings.Add( "No start health events returned for VM" )
}

## https://docs.microsoft.com/en-us/rest/api/monitor/activity-logs/list

## As we can't get a reliable initial provisioning time from the instance view, we'll work backwards in time until we get a start time as that will be the most recent
[int]$daysBack = 0
$events = $null
## we will only get back properties we need since this cuts down on the amount of data transferred
[string]$selectedProperties = 'correlationId,eventTimestamp,operationName,status'
do
{
    $daysBack += 7
    [string]$filter = "eventTimestamp ge '$(Get-Date -Date ([datetime]::Now.AddDays( -$daysBack )) -Format s)' and resourceUri eq '$AZid'"
    try
    {
        ## events come back as most recent first so first correlationId should be most recent as in the last boot
        $events = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/providers/Microsoft.Insights/eventtypes/management/values`?api-version=$insightsApiVersion&`$filter=$filter&`$select=$selectedProperties" -propertyToReturn 'value' ` | Where-Object { $_.OperationName.value -eq 'Microsoft.Compute/virtualMachines/start/action' }  )
    }
    catch
    {
        Write-Verbose -Message "Log retention period of $daysback gave exception $_"
        break
    }
} while( $null -eq $events -or $events.Count -eq 0 )

[array]$startEvents = @( $events | Select-Object -Property @{n='Time';e={ $_.eventTimestamp -as [datetime] }},* -ExcludeProperty eventTimestamp `
            | Group-Object -Property correlationId | Select-Object -first 1 -ExpandProperty Group | Sort-Object -Property Time )

Write-Verbose -Message "Got $($startEvents.Count) start events and $($healthEvents.Count) health events"

## remote to VM to get specific information we need

##https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/run-command?tabs=HTTP
[datetime]$earliestAZevent = [datetime]::MinValue
if( $healthEvents.Count -gt 0 -and $healthEvents[0].TimeOccurred -lt $startEvents[0].Time )
{
    $earliestAZevent = $healthEvents[0].TimeOccurred
}
else
{
    $earliestAZevent = $startEvents[0].Time
}

[string]$start = Get-Date -Date $earliestAZevent -Format u
[string]$end   = Get-Date -Date $earliestAZevent.AddMinutes( 10 ) -Format u

## Seems to be an undocumented limitation that stdout is limited to 4096 bytes. Also using aliases & positional parameters to reduce input size
## delimit different output sections with a blank/empty line
##          Get-Process -Name WindowsAzureGuestAgent | Select-Object -Property Name,Id,StartTime | ConvertTo-Json -Depth 1 -Compress
## not getting event id 21 for Microsoft-Windows-TerminalServices-LocalSessionManager as that is "Session logon succeeded" which pits uss against ALD
[string]$remoteCode = @"
        [datetime]::now.ToString( 'zzz' );
        ''
        `$boot = (gcim win32_operatingsystem).LastBootUpTime ; `$boot.ToString( 's' ) ;
        ''
        . {
            `$endTime = `$boot.AddMinutes( 15 )
            Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable @{ ProviderName = 'Microsoft-Windows-RemoteDesktopServices-RdpCoreCDV' ; Id = 65 ; StartTime = `$boot ; EndTime = `$endtime }
            ##Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable @{ ProviderName = 'Microsoft-Windows-Kernel-General' ; Id = 12 ; StartTime = `$boot ; EndTime = `$endtime }
            Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable @{ ProviderName = 'Service Control Manager' ; Id = 7036 ; StartTime = `$boot ; EndTime = `$endtime } | Where-Object Message -match 'WindowsAzureGuestAgent|TermService'
            Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable @{ ProviderName = 'Microsoft-Windows-TerminalServices-LocalSessionManager' ; Id = @( 41,42 ) ; StartTime = `$boot ; EndTime = `$endtime }
            } | Select @{n='T';e={ `$_.TimeCreated }},Id,@{n='P';e={`$_.ProviderName}},@{n='L';e={`$_.Level}},@{n='M';e={`$_.Message}} | Sort T | ConvertTo-Json -Depth 1 -Compress
"@

[hashtable]$body = @{
    'commandId' = 'RunPowerShellScript'
    'script' = @( $remoteCode -split "`r`n" )
}

[string]$runURI = "$baseURL/$azid/runCommand`?api-version=$computeApiVersion"

$waitResult = $null
$result = $null
[string]$stdout = $null

## this is an async operation so we need the response headers to get the URI to check the status
try
{
    $result = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $runURI -property $null -method POST -body $body -norest
}
catch
{
    $thisError = $null
    $thisError = $_ | Select-Object -ExpandProperty ErrorDetails | Select-Object -ExpandProperty Message |ConvertFrom-Json|Select-Object -ExpandProperty error
    if( $thisError )
    {
        if( $powerstate -ieq 'running' -or $thisError -notmatch 'The operation requires the VM to be running' )
        {
            Write-Warning -Message "Problem running $($body.script) in $sourceVMName : $thisError"
        }
    }
}

if( $result )
{
    if( $result.StatusCode -eq 202 ) ## accepted
    {
        [string]$asyncOperation = 'Azure-AsyncOperation'
        if( -Not $result.Headers -or -Not $result.Headers[ $asyncOperation ] )
        {
            Write-Warning -Message "Result headers missing $asyncOperation so unable to monitor command: $($body.script)"
        }
        else
        {
            $waitResult = Wait-AsyncAzureOperation -asyncURI $result.Headers[ $asyncOperation ] -azBearerToken $azBearerToken -operation "Run" -maxWaitTimeSeconds $maxWaitTimeSeconds -returnProperty output

            if( $null -ne $waitResult -and $false -ne $waitResult )
            {
                $stdout = $waitResult.value | Where-Object { $_.Code -ieq 'ComponentStatus/StdOut/succeeded' } | Select-Object -ExpandProperty Message -ErrorAction SilentlyContinue
                [string]$stderr = $waitResult.value | Where-Object { $_.Code -ieq 'ComponentStatus/StdErr/succeeded' } | Select-Object -ExpandProperty Message -ErrorAction SilentlyContinue
                if( [String]::IsNullOrEmpty( $stdout ) -or -Not [string]::IsNullOrEmpty( $stderr ) )
                {
                    Throw "Command `"$($body.script)`" errored : $stderr"
                }
            }
            else
            {
                Write-Warning -Message "Failed to get status from running $($body.script) in $sourceVMName"
            }
        }
    }
    else
    {
        Write-Warning -Message "Unexpected status $($result|Select-Object -ExpandProperty StatusCode) returned when trying to run `"$($body.script)`" in $sourceVMName"
    }
}

if( $stdout )
{
    [decimal]$timeDifference = 0
    ## first segment should be Get-Process information followed by empty/blank line then events
    [string[]]$segments = ($stdout -split '^\s*$' , 0 , 'Multiline').Trim() | Where-object { -Not [string]::IsNullOrEmpty( $_ ) }
    Write-Verbose -Message "Got $($segments.Count ) segments from text of $($stdout.Length)"
    ##$processInfo = ( $segments[ 0 ] | ConvertFrom-Json ) ## details of WindowsAzureGuestAgent process
    [string]$remoteTimZoneOffset = $segments[ 0 ] -replace ':' , '.'
    [string]$localTimZoneOffset = [datetime]::now.ToString( 'zzz' ) -replace ':'  , '.' 
    Write-Verbose -Message "Local time zone offset is $localTimZoneOffset , remote is $remoteTimZoneOffset"
    [array]$remoteEvents = @()
    try
    {
        $remoteEvents = $segments[ -1 ] | ConvertFrom-Json  ## | Select-Object -Property @{n='Time';e={$_.T.value}},* -ExcludeProperty T )
    }
    catch
    {
        Write-Warning -Message "Problem converting segment text to json : $($segments[ -1 ])"
    }
    Write-Verbose -Message "Got $($remoteEvents.Count) remote events"

    if( $localTimZoneOffset -ne $remoteTimZoneOffset )
    {
            ## Check both time zone offsets are valid

            $timeDifference = [decimal]$localTimZoneOffset - [decimal]$remoteTimZoneOffset

            ## change timestamps in all remote events to be local time or should that be UTC ??
            ForEach( $remoteEvent in $remoteEvents )
            {
            $remoteEvent.T.value = $remoteEvent.T.value.AddHours( $timeDifference ) ## also has string DateTime but we don't use this so no point changing
            }
    }

    [array]$AZeventLogs = @(
        $startEvents  | Select-Object -Property @{n='Time';e={$_.Time.ToString( $outputTimeFormat ) }},@{n='__Time';e={$_.Time}},@{n='Operation';e={$_.OperationName.localizedValue}},@{n='Status';e={$_.status.localizedValue}}
        $healthEvents | Select-Object -Property @{n='Time';e={$_.TimeOccurred.ToString( $outputTimeFormat ) }},@{n='__Time';e={$_.TimeOccurred}},@{n='Operation';e={$_.Summary}},@{n='Status';e={'N/A'}}
    ) | Sort-Object -Property Time
 
    ## time can come through as  ?2023?-?01?-?20T15:20:33.500000000Z
    [array]$remoteEventsOrdered = @( $remoteEvents | Select-Object -Property @{n='Time';e={Get-Date -Date $_.T.value }},@{n='Operation';e={$_.M -replace '\?' -replace "`r`n" , ' ' -replace '\s+' , ' ' }} -ExcludeProperty T )

    ## get the closest session start time before the first from Insight logs
    $avdConnectionStartEvent = $null
    $avdConnectionConnected = $null
    $bootTime = ($segments[ 1 ] -as [datetime]).AddHours( $timeDifference )

    if( $null -ne $avdConnections -and $avdConnections.Count -and $remoteEventsOrdered -and $remoteEventsOrdered.Count -gt 0 )
    {
        ## already sorted descended 
        if( $avdConnectionStartEvent = $avdConnections.Where( { $_.TimeGenerated -le $remoteEventsOrdered[0].Time -and $_.State -ieq 'Started'} , 1 ) | Select-Object -First 1 )
        {
            ## in case there is a disconnect/reconnect we need to get the first connected or completed event after boot
            $avdConnectionConnected  = $avdConnections.Where( { $_.TimeGenerated -ge $avdConnectionStartEvent.TimeGenerated -and $_.State -imatch 'Connected|Completed' -and $_.CorrelationId -eq $avdConnectionStartEvent.CorrelationId } ) | Select-Object -Last 1
        }
    }

    if( $avdConnectionStartEvent -or $avdConnectionConnected )
    {
        ## have seen where health events are missing so we have previous ones but could be that connection attempt that started VM failed & user didn't try again until now
        if( $avdConnectionStartEvent -and $healthEvents -and ($earlyEvents = $healthEvents.Where( { $_.TimeOccurred -lt $avdConnectionStartEvent.TimeGenerated } ) ) )
        {
            [datetime]$earliestEvent = ($earlyEvents | Measure-Object -Property TimeOccurred -Minimum).Minimum
            [decimal]$hoursDifference = [math]::round( ($avdConnectionStartEvent.TimeGenerated - $earliestEvent).TotalHours , 2 )
            
            $warnings.Add( "$($earlyEvents.Count) health events occur $hoursDifference hours before the user requested a connection so cannot be trusted"  )
        }
        Write-Output -InputObject 'Log Analytics'
        Write-Output -InputObject '============='
        @(
            $avdConnectionStartEvent , $avdConnectionConnected
        ) | Select-Object -Property @{n='Time';e={Get-Date -Date $_.TimeGenerated -Format $outputTimeFormat }},@{n='Operation';e={"AVD connection $($_.State.ToLower()) for $($_.Username)"}} | Format-Table -AutoSize
    }

    Write-Output -InputObject "Azure Logs"
    Write-Output -InputObject "=========="
    $AZeventLogs | Select-Object -Property * -ExcludeProperty __* | Format-Table -AutoSize

    Write-Output -InputObject "`nVM Event Logs"
    Write-Output -InputObject "============="
    @(
        if( $bootTime )
        {
            [pscustomobject]@{ Time = $bootTime ; Operation = 'Operating System started' }
        }
        $remoteEventsOrdered ) | Format-Table -Property @{n='Time';e={$_.Time.ToString( $outputTimeFormat ) }},Operation -AutoSize

    ## '\d#\d+ created' will match the first winstation created which is what we measure because user might not have had a successful logon at tht point but we are measuring time to availability, not successful logon
    $firstConnection = $remoteEvents | Where-Object { $_.Id -eq 65 -and $_.P -ieq 'Microsoft-Windows-RemoteDesktopServices-RdpCoreCDV' -and $_.M -match '^Connection.*\d+#\d+\s*created' } | Select-Object -First 1

    [datetime]$earliestEvent = $AZeventLogs[0].Time
    if( $avdConnectionStartEvent )
    {
        $earliestEvent = $avdConnectionStartEvent.TimeGenerated
    }
    [datetime]$latestEvent = $remoteEventsOrdered[-1].Time

    if( $firstConnection )
    {
        $latestEvent = $firstConnection.T.value
    }

    Write-Output -InputObject "Total elapsed time to first remote connection establishment = $(($latestEvent - $earliestEvent).TotalSeconds -as [int]) seconds`n"

    $connectionsSinceBoot = @( $avdConnections | Where-Object { $_.State -ieq 'Started' -and $_.TimeGenerated -gt $bootTime } )

    if( $null -ne $connectionsSinceBoot -and $connectionsSinceBoot.Count -gt 0 )
    {
        Write-Output -InputObject "Note that there have been $($connectionsSinceBoot.Count) connections initiated since boot for $(@($connectionsSinceBoot | Group-Object -Property UserName).Count) different users, first one at $($connectionsSinceBoot[-1].TimeGenerated.ToString('G'))"
    }

    Write-Output -InputObject 'Note that all times are local'
    if(  $localTimZoneOffset -ne $remoteTimZoneOffset )
    {
        $warnings.Add( "There is a time zone difference of $timeDifference hours between the VM and locally (VM is $remoteTimZoneOffset hours ahead of UTC, local is $localTimZoneOffset hours ahead)" )
    }
}
else
{
    Write-Warning -Message "Failed to get event log information from VM"
}

if( $null -ne $warnings -and $warnings.Count -gt 0 )
{
    ''
    $warnings | Write-Warning
}

