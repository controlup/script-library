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

.PARAMETER includeProviderRegex
    Only include Azure resources where the providers match this regular expression

.PARAMETER excludeProviderRegex
    Exclude Azure resources where the providers match this regular expression

.PARAMETER sortby
    Which output property to sort (group) by

.PARAMETER raw
    Output objects rather than text

.NOTES
    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2021-10-30
    Updated:        2022-06-17  Guy Leech  Added code to deal with paging in results
                    2022-09-30  Guy Leech  Added setting of TLS12 & TLS13
                    2024-02-09  Guy Leech  Fixed bug around auto selection of api versions when no release version available. Refactored to query api version available if not passed
                    2024-02-14  Guy Leech  Added code to ignore "delete" activities in log & filter out other non-relevant entries. Logic to use managedby for resource if available otherwise get resource details.
                    2024-02-16  Guy Leech  PS script analyser run. Superfluous code removed
                    2024-02-21  Guy Leech  Error 429 handling. API version caching moved to function that uses it.
                    2024-02-23  Guy Leech  Updated shared Azure functions imported
#>

[CmdletBinding()]

Param
(
    [string]$AZid ,## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId ,
    [double]$daysBack = 30 ,
    [ValidateSet('Yes','No')]
    [string]$resourceGroupOnly = 'Yes',
    [string]$sortby = 'type' ,
    [string]$includeProviderRegex ,
    [string]$excludeProviderRegex ,
    [switch]$raw
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 400
try
{
    if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
    {
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions
    }
}
catch
{
    ## not fatal
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
[string]$computeApiVersion = '2021-07-01'
[string]$insightsApiVersion = '2015-04-01'
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

[datetime]$startTime = [datetime]::Now
$azSPCredentials = $null
$azSPCredentials = Get-AzSPStoredCredentials -system $credentialType -tenantId $AZtenantId

If ( -Not $azSPCredentials )
{
    Exit 1 ## will already have output error
}

[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13

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
    $resourcesURL = "resourceGroups/$resourceGroupName/"
}
## else ## will be for the subscription

## get all provider details so we can pick version if we don't have one
##[array]$allProviders = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/providers?api-version=$providersApiVersion" -retries 2 )

##Write-Verbose -Message "$([datetime]::Now.ToString( 'G' )): Got $($allProviders.count) providers"

[array]$allResources = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/$($resourcesURL)resources`?`$expand=createdTime,changedTime,managedBy`&api-version=$resourceManagementApiVersion" -retries 2 | Where-Object type -NotIn $excludedResourceTypes )

Write-Verbose -Message "$([datetime]::Now.ToString( 'G' )): Got $($allResources.Count) resources in total"

## https://docs.microsoft.com/en-us/rest/api/monitor/activity-logs/list
## we only need the resource id  so only get that for efficiency plus operation name so we can ignore delete. Also, resource helath and advisor operations show but with no tenant id so remove those
[array]$allevents = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/providers/Microsoft.Insights/eventtypes/management/values`?api-version=$insightsApiVersion&`$filter=$filter&`$select=resourceid,tenantid,channels,operationname" -retries 2 | Where-Object { $_.tenantId -ieq $AZtenantId -and $_.channels -ieq 'Operation' -and $_.operationname.value -notmatch '^Microsoft\.Advisor/' -and $_.operationName.value -notmatch '/delete$' | Select-Object -Property resourceid })

Write-Verbose -Message "$([datetime]::Now.ToString( 'G' )): Got $($allevents.Count) events in total"

## produce hash table keyed on resource id
[hashtable]$resourcesInActivityLog = $allevents | Group-Object -Property resourceId -AsHashTable

if( $null -eq $resourcesInActivityLog )
{
    $resourcesInActivityLog = @{}
}

Write-Verbose -Message "$([datetime]::Now.ToString( 'G' )): Got $($resourcesInActivityLog.Count) resources from activity log and $($allResources.Count) resources via resource group $resourceGroupName"

##[hashtable]$usedVMs = @{}
[hashtable]$resourceDetails = @{}
[hashtable]$vmDetails = @{}
[hashtable]$resourceGroups = @{}
[int]$counter = 0
[string[]]$resourceTypesOfInterest  = @( 'Microsoft.Network/networkInterfaces'  , 'Microsoft.Compute/virtualMachines' , 'Microsoft.Compute/disks' ) ## seems that managedBy for disks is not reliable
[System.Collections.generic.list[object]]$unusedResources = @( ForEach( $resource in $allResources )
{
    $counter++
    if( ( -Not [string]::IsNullOrEmpty( $includeProviderRegex ) -and $resource.type -notmatch $includeProviderRegex ) -or ( -Not [string]::IsNullOrEmpty( $excludeProviderRegex ) -and $resource.Type -match $excludeProviderRegex ) )
    {
        Write-Verbose -Message "-- excluding $($resource.type) $($resource.name)"
    }
    elseif( -Not $resourcesInActivityLog[ $resource.id ] )
    {
        [bool]$include = $null -eq $resource.psobject.properties[ 'createdTime' ] -or ( $null -ne $resource.createdTime -and [datetime]$resource.createdTime -lt $startFrom )

        ##[bool]$include = $null -ne $resource.psobject.properties[ 'changedTime' ] -and $null -ne $resource.changedTime -and [datetime]$resource.changedTime -lt $startFrom

        if( $include -and ( $null -eq $resource.psobject.properties[ 'changedTime' ] -or ( $null -ne $resource.changedTime -and [datetime]$resource.changedTime -lt $startFrom) ))
        {
            Write-Verbose -Message "$counter / $($allResources.Count) : including $($resource.id)"
            ## last modified time on resource itself can be newer than we were told in the /resources query - no longer seems to be the case as of 2024/02/12
            ## TODO do we need to throttle the requests ?
            $resourceDetail = $null
            if( $resource.type -in $resourceTypesOfInterest )
            {
                ## if we have managedBy property then no need to fetch resource detail
                if( $resource.PSObject.Properties[ 'managedBy' ] )
                {
                    Write-Verbose "** Got managed by $($resource.managedBy) for $($resource.id)"
                }
                elseif( $resourceDetail = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$($resource.id)" -retries 2 -newestApiVersion -propertyToReturn $null -type $resource.type )
                {
                    if( $resource.type -ieq 'Microsoft.Compute/virtualMachines' )
                    {
                        $vmDetails.Add( $resource.Id , $resourceDetail ) ## will check these later for orphaned resources
                    }
                    $resourceDetails.Add( $resource.id , $resourceDetail ) ## use in 2nd pass
                }
            }
            if( $include )
            {
                ## what if VM has been running the whole time so no start/stop? Check if running now
                if( $resource.type -ieq 'Microsoft.Compute/virtualMachines' )
                {
                    ## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/instance-view
                    if( $null -ne ( $instanceView = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$($resource.id)/instanceView`?api-version=$computeApiVersion" -property $null) ) )
                    {
                        if( $instanceview.Statuses | Where-Object code -match '^PowerState/(.*)$' ) ## -and ( $powerstate = ($line -split '/' , 2 )[-1] ))
                        {
                            ## https://learn.microsoft.com/en-us/dotnet/api/microsoft.azure.management.compute.fluent.powerstate?view=azure-dotnet-legacy
                            if( $matches[1] -ine 'Deallocated' ) ## if stopped then still accumulating cost. unknown warrants inclusion
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
                }

                if( $resourceGroupOnly -ieq 'no')
                {
                    if( $resource.id -match '\bsubscriptions/([^/]+)/resourceGroups/([^/]+)/providers/Microsoft\.' )
                    {
                        [string]$thisResourceGroupName = $Matches[2]

                        if( -Not $resourceGroups.ContainsKey( $thisResourceGroupName ) )
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
                        Add-Member -InputObject $resource -MemberType NoteProperty -Name 'Resource Group' -Value $thisResourceGroupName
                    }
                    else
                    {
                        Write-Warning -Message "Failed to determine resource group from $($resource.id)"
                    }
                }

                if( $include )
                {
                    $resource
                }
            }

            if( -Not $include )
            {
                Write-Verbose -Message "`tUsed : $($resource.id)"
            }
        }
    }
})

Write-Verbose -Message "$([datetime]::Now.ToString( 'G' )): Got $($unusedResources.Count) potentially unused resources out of $($allResources.Count) total"

[hashtable]$VMsNSGs = @{}

ForEach( $nic in $resourceDetails.GetEnumerator().Where( { $_.value.PSObject.Properties[ 'type' ] -and $_.value.type -ieq 'Microsoft.Network/networkInterfaces' } ))
{
    ## TODO could there be more than one of either?
    if( ( $NSG = $nic.value.properties | Select-Object -ExpandProperty networkSecurityGroup -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id ) `
        -and ( $VM = $nic.value.properties | Select-Object -ExpandProperty virtualMachine -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id ) )
    {
        if( -Not $VMsNSGs.ContainsKey( $NSG ) )
        {
            $VMsNSGs.Add( $NSG , ([System.Collections.Generic.List[object]]@( $VM ) ))
        }
        else
        {
            $VMsNSGs[ $NSG ].Add( $VM )
        }
    }
}

Write-Verbose -Message "$([datetime]::Now.ToString( 'G' )): Got $($VMsNSGs.Count) network security groups"

$unusedResourceIds = New-Object -TypeName System.Collections.Generic.HashSet[string] -ArgumentList ([StringComparer]::InvariantCultureIgnoreCase)

## ForEach on a list does not use $_ for the item being iterated on
## Checking existence is far quicker than an array for large numbers of items
ForEach( $unusedResource in $unusedResources )
{
    if( $null -ne $unusedResource -and $unusedResource.Id )
    {
        $null = $unusedResourceIds.Add( $unusedResource.Id )
    }
}
Write-Verbose -Message "$([datetime]::Now.ToString( 'G' )): Got $($unusedResourceIds.Count) resource ids in dictionary"

[int]$originalTotal = $unusedResources.Count
## second pass if Microsoft.Compute/disks or Microsoft.Network/networkinterfaces or Microsoft.Network/networkSecurityGroups then check parent VM (if there is one) as we will have it in the unused collection
## cache VM details so if disk is for a machine we already have, we don't need to get it from Azure again
For( [int]$index = $unusedResources.Count -1 ; $index -ge 0 ; $index-- )
{
    $potentionallyUnusedResource = $unusedResources[ $index ]
    Write-Verbose -Message "$([datetime]::Now.ToString( 'G' )): checking unused resource @ $index / $($unusedResources.Count) / $originalTotal"
    [bool]$remove = $false
    [bool]$orphaned = $false
    [bool]$childResource = $false
    $parentVM = $null

    if( $potentionallyUnusedResource.type -ieq 'Microsoft.Network/networkInterfaces' )
    {
        $childResource = $true

        if( -Not ( $parentVM = $potentionallyUnusedResource | Select-Object -ExpandProperty managedBy -ErrorAction SilentlyContinue ) ) ## not seen this property returned for NICs but it may change
        {
            $parentVM = $resourceDetails[ $potentionallyUnusedResource.id ] | Select-Object -ExpandProperty properties | Select-Object -ExpandProperty virtualMachine -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id
        }
    }
    elseif( $potentionallyUnusedResource.type -ieq 'Microsoft.Compute/disks' )
    {
        $childResource = $true
        if( -Not ( $parentVM = $potentionallyUnusedResource | Select-Object -ExpandProperty managedBy -ErrorAction SilentlyContinue ) ) ## managedBy in all resources query was not reliably returning it
        {
            $parentVM = $resourceDetails[ $potentionallyUnusedResource.id ] | Select-Object -ExpandProperty managedBy -ErrorAction SilentlyContinue
        }
    }
    elseif( $potentionallyUnusedResource.type -ieq 'Microsoft.Network/networkSecurityGroups' )
    {
       if( $VMs = $VMsNSGs[ $potentionallyUnusedResource.id ] )
       {
            ## not orphaned but need to see if any of the VMs using it are in use
            ForEach( $VM in $VMs )
            {
                ## if( -Not ( $unused = $unusedResources.Where( { $_.id -ieq $VM } ) ))
                if( -Not ( $unused = $unusedResourceIds.Contains( $VM )))
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

    if( $childResource )
    {
        if( $parentVM )
        {
            ## now see if in our unused VMs and if not we remove from this list as must've been in use
            ## if( -Not ( $unused = $unusedResources.Where( { $_.id -ieq $parentVM } , 1 ) ))
            if( -Not ( $unused = $unusedResourceIds.Contains( $parentVM ) ) )
            {
                $remove = $true
            }
            else
            {
                Add-Member -InputObject $potentionallyUnusedResource -Force -MemberType NoteProperty -Name 'ParentVM' -Value $parentVM
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

[string]$message = "$([datetime]::Now.ToString( 'G' )): Found $($unusedResources.Count) resources, out of $($allResources.count) examined, potentially not used since $(Get-Date -Format G -Date $startFrom) in "

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

if( -not $raw )
{
    Write-Output -InputObject $message
}

$finalOutput = $unusedResources | Select-Object -Property *,
    @{name='Created';expression={ $_.createdTime -as [datetime] }} ,
    @{name='Changed';expression={ $_.changedTime -as [datetime] }} ,
    @{name='Parent';expression={ $( if( $_.PSObject.Properties[ 'parentVM' ] ) { $_.parentVM } else { $_.managedby }) -replace '^/subscriptions/[^/]+/' }} -ExcludeProperty parentVM,id,tags,managedBy,createdTime,changedTime | Sort-Object -Property $sortby

if( $raw )
{
    $finalOutput
}
else
{
    $finalOutput | Format-Table -AutoSize
}


[datetime]$endTime = [datetime]::Now
Write-Verbose -Message "$($endtime.ToString( 'G')) : total run time $(($endTime - $startTime).TotalSeconds) seconds, making $($script:totalRequests) requests"

