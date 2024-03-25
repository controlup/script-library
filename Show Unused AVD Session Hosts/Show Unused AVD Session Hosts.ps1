#requires -version 3.0

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

.PARAMETER hostpoolOnly
    Only return results for the host pool containing the AZid
    
.PARAMETER resourceGroupOnly
    Only return results for the resource group containing the AZid
    
.PARAMETER includeUsedWithUsers
    Incude session hosts which have had users on them otherwise only show machines which have had no users

.PARAMETER includeNotPoweredOn
    Incude session hosts which have not been powered on in the period otherwise include all

.PARAMETER VMOnly
    Only return results for the resource in the AZid

.NOTES
    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2021-10-30
    Updated:        2022-06-17  Guy Leech  Added code to deal with paging in results
                    2024-02-06  Guy Leech  Updated API version numbers. Added -includeUsedWithUsers and -includeNotPoweredOn. Implemented -hostpoolOnly
                    2024-02-07  Guy Leech  Summary added
                    2024-02-08  Guy Leech  Added setting of TLS12 & TLS13
                    2024-02-23  Guy Leech  Updated shared Azure functions imported
#>

[CmdletBinding()]

Param
(
    [string]$AZid ,## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId ,
    [double]$daysBack = 30 ,
    [ValidateSet('Yes','No','Only')]
    [string]$summary = 'Yes',
    [ValidateSet('Yes','No')]
    [string]$resourceGroupOnly = 'Yes',
    [ValidateSet('Yes','No')]
    [string]$hostpoolOnly = 'yes',
    [ValidateSet('Yes','No')]
    [string]$includeUsedWithUsers = 'No',
    [ValidateSet('Yes','No')]
    [string]$includeNotPoweredOn = 'No'
    ##[string]$sortby = 'type' ,
    ##[switch]$raw
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

[string]$insightsApiVersion = '2015-04-01'
[string]$OperationalInsightsApiVersion = '2022-10-01'
[string]$desktopVirtualisationApiVersion = '2023-09-05'
[string]$baseURL = 'https://management.azure.com'
[string]$credentialType = 'Azure'
[hashtable]$script:apiversionCache = @{}
$warnings = New-Object -TypeName System.Collections.Generic.List[string]
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
                        $warnings.Add(  "Unimplemented type $($table.columns[ $index ].type), treating as string" )
                    }
                    $result.Add( $table.columns[ $index ].name , $row[ $index ] )
                }
            }
            [pscustomobject]$result
        }
    }
}

Function Out-PassThru
{
    Process
    {
        $_
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
$hostPoolForAzId = $null

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
[datetime]$scriptStartTime = [datetime]::Now
[datetime]$startFrom = $scriptStartTime.AddDays( -$daysBack )
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

[array]$hostpools = @()

## TODO filter if host pool only parameter passed to script 
if( $resourceGroupOnly -ieq 'yes' )
{
    ## https://learn.microsoft.com/en-us/rest/api/desktopvirtualization/host-pools/list-by-resource-group?view=rest-desktopvirtualization-2022-02-10-preview&tabs=HTTP
    $hostpools = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/resourcegroups/$resourceGroupName/providers/Microsoft.DesktopVirtualization/hostPools?api-version=$desktopVirtualisationApiVersion" )
}
else ## get all host pools
{
    $hostpools = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/providers/Microsoft.DesktopVirtualization/hostPools?api-version=$desktopVirtualisationApiVersion" )
}

Write-Verbose -Message "Retrieved $($hostpools.Count) host pools"
    
[array]$logSinsights = @()

Write-Verbose -Message "Got $($hostpools.Count) hostpools"

if( $null -eq $hostpools -or $hostpools.Count -eq 0 )
{
    Throw "No host pools found"
}

[hashtable]$sessionhostsUsed = @{} ## New-Object -TypeName System.Collections.Generic.List[object]

## if we have any AVD, see if we have Log Insights
[array]$workspaces = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.OperationalInsights/workspaces?api-version=$OperationalInsightsApiVersion" -retries 2 )

if( -Not ( $loganalyticsBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID -scope 'https://api.loganalytics.io' ) )
{
    Throw "Unable to get log analytics bearer token for $($azSPCredentials.spCreds)"
}
elseif( $null -eq $workspaces -or $workspaces.Count -eq 0 )
{
    Throw "No log analytics workspaces found so no AVD session history available"
}
else
{
    ForEach( $workspace in $workspaces )
    {
        if( $workspace.properties.retentionInDays -lt $daysBack )
        {
            $warnings.Add( "Days retention is $($workspace.properties.retentionInDays) in logs `"$($workspace.name)`" is less than days back requested of $daysBack so may miss some AVD sessions" )
        }
        ## $usages = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.OperationalInsights/workspaces/$($workspace.Name)/usages?api-version=$OperationalInsightsApiVersion" -propertyToReturn $null -retries 2 )
        $query = "WVDConnections | where TimeGenerated > ago($($daysBack)d) and State == `"Connected`""

        ## https://docs.microsoft.com/en-us/rest/api/loganalytics/dataaccess/query/get?tabs=HTTP
        $queryResult = Invoke-AzureRestMethod -BearerToken $loganalyticsBearerToken -uri "https://api.loganalytics.io/v1/workspaces/$($workspace.properties.customerId)/query?query=$query" -propertyToReturn $null -retries 2
        if( $null -ne $queryResult -and $queryResult.PSObject.Properties[ 'tables' ] )
        {
            ## sessionhostname that comes back is fqdn and we prepend host pool name so we can cross reference to session hosts
            $AVDconnections = ConvertTo-Object -Tables $queryResult.tables -ExtraFields @{ Workspace = $workspace } | Select-Object -Property *,@{n='hostname';e={ ($_.SessionHostName -split '\.')[0] }},@{n='hostpool' ; e = { $script:thishostpool = $_._ResourceId -replace '^/subscriptions/[a-z0-9\-]+/resourcegroups/wvd/providers/microsoft\.desktopvirtualization/hostpools/' ; $script:thishostpool }},@{n='hostPoolSessionHost';e={ "$($script:thishostpool)/$($_.SessionHostName)"}} | Group-Object -Property hostPoolSessionHost -AsHashTable -AsString
            $sessionhostsUsed += $AVDconnections
            Write-Verbose -Message "Got $($sessionhostsUsed.count) session hosts that have been used since $(Get-Date -Date $startFrom -Format G)"
        }
    }
}
## iterate over all host pools to get session hosts in them and then look them up in the log analytics data to see if they have been logged on to or not so we can mark as such in output

if( $sessionhostsUsed.Count -eq 0 )
{
    $warnings.Add( "No session connect events found since $($startFrom.ToString('G'))" )
}

[array]$allSessionHosts = @( ForEach( $hostpool in $hostpools )
{
    ## https://learn.microsoft.com/en-us/rest/api/desktopvirtualization/session-hosts/list?tabs=HTTP
    ## /subscriptions/58ffa3cb-DEAD-BEEF-DADA-369c1fcebbf5/resourcegroups/WVD/providers/Microsoft.DesktopVirtualization/hostpools/host-pool-server
    [string]$thisResourceGroup = $hostpool.id -replace '^/subscriptions/[^/]+/resourcegroups/([^/]+)/providers/.*$' , '$1'
    [array]$sessionHosts = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/resourceGroups/$thisResourceGroup/providers/Microsoft.DesktopVirtualization/hostPools/$($hostpool.name)/sessionHosts?api-version=$desktopVirtualisationApiVersion" -retries 2 )
    if( $null -eq $sessionHosts -or $sessionHosts.Count -eq 0 )
    {
        $warnings.Add( "Host pool `"$($hostpool.name)`" in resource group `"$thisResourceGroup`" contains no session hosts" )
    }
    else
    {
        ForEach( $sessionHost in $sessionHosts )
        {
            if( $sessionHost.properties.resourceId -ieq $AZid )
            {
                $hostPoolForAzId = $hostpool
            }
            $result = Add-Member -PassThru -InputObject $sessionHost -NotePropertyMembers @{              
                ## session host is hostpool/sessionhost.fqdn
                SessionHostName = $sessionhost.Name -replace '^.*/([^.]+)\..*$' , '$1'
                HostPool = $hostpool.name
                ResourceGroup = $thisResourceGroup
                Used = $sessionhostsUsed[ $sessionHost.Name ]
            }
            $result
        }
    }
    if( $hostpoolOnly -ieq 'yes' -and $null -ne $hostPoolForAzId )
    {
        break ## we have found the hostpool for the passed AZid but there may be earlier hostpools/sessionhosts that we still need to ignore
    }
})

Write-Verbose -Message "Got $($allSessionHosts.Count) session hosts in $($hostpools.Count) host pools"

if( $hostpoolOnly -ieq 'yes' )
{
    if( $null -eq $hostPoolForAzId )
    {
        Throw "Unable to find a host pool containing resource id $AZid"
    }
    if( $hostpools.Count -gt 1 )
    {
        $hostpools = @( $hostPoolForAzId ) ## the only host pool we are interested in
        $allSessionHosts = @( $allSessionHosts | Where-Object HostPool -ieq $hostPoolForAzId.Name )
    }
}

Write-Verbose -Message "Now have $($hostpools.Count) host pools and $($allSessionHosts.Count) session hosts"

## https://docs.microsoft.com/en-us/rest/api/monitor/activity-logs/list

[datetime]$startFrom = (Get-Date).AddDays( -$daysBack )
[string]$filter = "eventTimestamp ge '$(Get-Date -Date $startFrom -Format s)' and resourceProvider eq 'Microsoft.Compute'"
## cannot filter on ResourceProvider and ResourceGroup so post filter on the latter
## TODO filter on host pool only

[array]$startDealllocateEvents = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/providers/Microsoft.Insights/eventtypes/management/values`?api-version=$insightsApiVersion&`$filter=$filter&`$select=operationname,resourcegroupname,eventTimestamp,resourceid,status,eventname" -propertyToReturn 'value' -retries 2 | Where-Object { ( $_.OperationName.value -ieq 'Microsoft.Compute/virtualMachines/deallocate/action' -or $_.OperationName.value -ieq 'Microsoft.Compute/virtualMachines/start/action' ) -and $_.status.value -ieq 'Succeeded' -and $_.eventName.value -ieq 'EndRequest' -and ( $resourceGroupOnly -ieq 'no' -or $_.ResourceGroupName -ieq $resourceGroupName ) } `
    | Select-Object -Property @{n='time';e={([datetime]$_.eventTimestamp).ToUniversalTime() }}, ## maintain uTC as casting string to datetime makes local time
        @{n='operation';e={$_.operationname.value -replace 'Microsoft.Compute/virtualMachines/(.+)/action' , '$1' }},
        resourcegroupname,
        @{n='VM';e={ $_.resourceid -replace '^.*/providers/Microsoft\.Compute/virtualMachines/' }} | Sort-Object -Property time | Group-Object -Property VM,resourcegroupname )

[int]$counter = 0
[array]$upTimes = @( ForEach( $machine in $startDealllocateEvents ) `
{
    $counter++
    [uint64]$secondsRunning = 0
    $VM,$resourcegroup = $machine.name -split ',\s?'
    Write-Verbose -Message "$counter : $($startDealllocateEvents.Count) : $resourcegroup / $VM"
    [datetime]$lastStartTime = $startFrom ## assume machine was already running before logging if we don't see a start event before a deallocate
    ForEach( $event in $machine.group )
    {
        if( $event.operation -ieq 'start' )
        {
            $lastStartTime = $event.Time
        }
        elseif( $event.operation -ieq 'deallocate' )
        {
            if( $lastStartTime -ne [datetime]::MinValue )
            {
                $running = ($event.time - $lastStartTime).TotalSeconds
                $secondsRunning += $running
                $lastStartTime = [datetime]::MinValue
            }
            else
            {
                $warnings.Add( "Got another deallocate for $($machine.machinename) without a start event in between" )
            }
        }
        else
        {
            $warnings.Add( "Unexepected event $($event) for $($machine.name)" )
        }
    }
    if( $lastStartTime -ne [datetime]::MinValue )
    {
        ## still running as not seen deallocate for previous start - ## TODO do we cross check that it is still running?
        $running = ($scriptStartTime - $lastStartTime ).TotalSeconds
        if( $running -gt 0 )
        {
            $secondsRunning += $running
        }
    }
    $result = [pscustomobject]@{
        'VM' = $VM
        'ResourceGroupName' = $resourcegroup
        'UpTimeSeconds' = $secondsRunning
    }
    Write-Verbose -Message "$resourcegroup/$VM up $($result.UpTimeSeconds) seconds"
    $result
})

[array]$usage = @( ForEach( $machine in $allSessionHosts )
{
    [array]$uniqueUsers = @()
    [int]$totalSessions = 0
    [uint64]$VMupTimeSeconds = 0
    $VMupTimeSeconds =  $upTimes | Where-Object { $_.VM -ieq $machine.SessionHostName -and $_.ResourceGroupName -ieq $machine.ResourceGroup } | Select-Object -ExpandProperty UpTimeSeconds -First 1
    $used = $sessionhostsUsed[ $machine.name ]
    if( $used ) ##  VM has had user
    {
        if( $includeUsedWithUsers -ieq 'yes' )
        {
            $uniqueUsers = @( $used | Group-Object -Property username )
            $totalSessions = $used.Count
        }
        else ## used but not including used session hosts
        {
            continue
        }
    }
    elseif( $includeNotPoweredOn -ieq 'no' -and $VMupTimeSeconds -eq 0 )
    {
        continue
    }

    $result = [pscustomobject]@{
        'Host Pool' = $machine.HostPool
        'Session Host Name' = $machine.SessionHostName
        'Unique Users' = $uniqueUsers.Count
        'Total Sessions' = $totalSessions
        'Run time (hours)' = [math]::Round( $VMupTimeSeconds / 3600 , 2 )
    }
    if( $resourceGroupOnly -ine 'yes' )
    {
        Add-Member -InputObject $result -MemberType NoteProperty -Name 'Resource Group' -Value $machine.ResourceGroup
    }
    $result
} )

if( $summary -ine 'no' )
{
    [array]$usageByHostPool = @( $usage | Group-Object -Property 'Host Pool' )
    [array]$allSessionHostsByHostPool = @( $allSessionHosts | Group-Object -Property HostPool )

    [array]$summaryData = @( ForEach( $hostpool in $hostpools ) `
    {
        $hostpoolUsage = $usageByHostPool | Where-Object Name -ieq $hostpool.Name
        $sessionHostsForHostPool = $allSessionHostsByHostPool | Where-Object Name -ieq $hostpool.name
        $sessionLess = @( $usage | Where-Object { $_.'Host Pool' -ieq $hostpool.Name -and $_.'Total Sessions' -eq 0 -and $_.'Run time (hours)' -gt 0 } )
        $hadSession  = @( $usage | Where-Object { $_.'Host Pool' -ieq $hostpool.Name -and $_.'Total Sessions' -gt 0 } )

        if( $includeNotPoweredOn -ieq 'yes' -or $sessionLess.Count -gt 0 )
        {
            $result = [pscustomobject]@{
                'Host Pool' = $hostpool.name
                'Total Session Hosts' = $sessionHostsForHostPool.Count
                'Session Hosts with User Sessions' = $hadSession.Count 
                'Powered Up but Unused' = $sessionLess.Count
                'Powered Up but Unused - Hours Total' = $sessionLess | Measure-Object -Property 'Run time (hours)' -Sum | Select-Object -ExpandProperty Sum
            }
            $result ## output
        }
    })
    
    if( $null -ne $summaryData -and $summaryData.Count -gt 0 )
    {
        Write-Output -InputObject "Session host usage summary across $($hostpools.Count) host pools & $($allSessionHosts.Count) session hosts :"
        $summaryData | Sort-Object -property 'Session Hosts Powered Up Unused Hours' -Descending | Format-Table -AutoSize
    }
    ## else no summary data which we inform about below

    if( $summary -ieq 'only' )
    {
        exit 0
    }
}

if( $null -eq $usage -or $usage.Count -eq 0 )
{
    Write-Output -InputObject "None of the $($allsessionHosts.Count) session hosts in the $($hostpools.Count) host pools meet the criteria to be displayed here"
}
else ## some usage
{
    Write-Output -InputObject "Indvidual Session Host usage and run time since $($startFrom.ToString('G'))"

    $usage | Sort-Object -Property @{ expression = 'Unique Users' ; descending = $false } , @{ expression = 'Run time (hours)' ; descending = $true } | Format-Table -AutoSize
}

[datetime]$endTime = [datetime]::Now
Write-Verbose -Message "$($endtime.ToString( 'G')) : total run time $(($endTime - $startTime).TotalSeconds) seconds, making $($script:totalRequests) requests"

if( $null -ne $warnings -and $warnings.Count -gt 0 )
{
    $warnings | Write-Warning
}

