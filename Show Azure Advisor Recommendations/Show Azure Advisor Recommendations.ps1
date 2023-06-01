#require -version 3.0

<#
.SYNOPSIS
    Get Azure Advisor Recommendations

.DESCRIPTION
    Using REST API calls

.PARAMETER azid
    The relative URI of the Azure VM
    
.PARAMETER AZtenantId
    The azure tenant ID

.PARAMETER resourceGroupOnly
    Only return results for the resource group containing the AZid
    
.PARAMETER VMOnly
    Only return results for the resource specified in the AZid
    
.PARAMETER sortBy
    Sort output by this field

.NOTES
    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2023-04-04
    Updated:        2023-04-05  Guy Leech  Added Resource Group Name when not resurce group specific query
                                           Added -sortby and -filterby parameters
#>

[CmdletBinding()]

Param
(
    [string]$AZid ,## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId ,
    [ValidateSet('Yes','No')]
    [string]$resourceGroupOnly = 'Yes',
    [ValidateSet('Yes','No')]
    [string]$vmOnly = 'Yes',
    [string]$sortBy = 'type' ,
    [string]$filterCategoryBy = '*' ,
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

[string]$computeApiVersion = '2021-07-01'
[string]$insightsApiVersion = '2015-04-01'
[string]$resourceManagementApiVersion = '2021-04-01'
[string]$OperationalInsightsApiVersion = '2022-10-01'
[string]$desktopVirtualisationApiVersion = '2021-07-12'
[string]$recommendationsApiVersion = '2023-01-01'

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
                if( ( $_ | Select-Object -ExpandProperty ErrorDetails | Select-Object -ExpandProperty Message | ConvertFrom-Json -ErrorAction SilentlyContinue | Select-Object -ExpandProperty error | Select-Object -ExpandProperty message) -match 'for type ''([^'']+)''\. The supported api-versions are ''([^'']+)''')
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
[string]$resourceIdRegex = '/?\bsubscriptions/([^/]+)/resourceGroups/([^/]+)/providers/Microsoft\..*$'

## subscriptions/58ffa3cb-2f63-4242-a06d-deadbeef/resourceGroups/WVD/providers/Microsoft.Compute/virtualMachines/GLMW10WVD-0
if( $AZid -match $resourceIdRegex )
{
    $subscriptionId = $Matches[1]
    $resourceGroupName = $Matches[2]
}
else
{
    Throw "Failed to parse subscription id and resource group from $AZid"
}

if( $sortby -imatch 'type' )
{
    $sortby = 'Resource Type'
}
elseif( $sortby -imatch 'resource|group' )
{
    $sortby = 'Resource Group'
}

##https://learn.microsoft.com/en-us/rest/api/advisor/Recommendations/List?tabs=HTTP

## GET https://management.azure.com/subscriptions/{subscriptionId}/providers/Microsoft.Advisor/recommendations?api-version=2022-10-01&$filter={$filter}&$top={$top}&$skipToken={$skipToken}

[string]$URI = "$baseURL/subscriptions/$subscriptionId/providers/Microsoft.Advisor/recommendations?api-version=$recommendationsApiVersion"
[string]$filter = $null
[System.Collections.Generic.List[object]]$propertiesToOutput = @( 'Category','Impact',@{n='Resource Type';e={$_.impactedField -replace '^.*/'}},@{n='Resource';e={$_.ImpactedValue}},@{n='Problem';e={$_|Select-Object -ExpandProperty ShortDescription -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Problem}} )

if( $vmOnly -ieq 'yes' )
{
    $filter = "&`$filter=ResourceId eq '$Azid'"
}
elseif( $resourceGroupOnly -ieq 'yes' )
{
    $filter = "&`$filter=ResourceGroup eq '$resourceGroupName'"
}
else
{
    ## no filter but need the resource group name from the resource id /subscriptions/05f0d60d-0b58-43e2-a66f-80a6ae691a13/resourceGroups/terraform2/providers/Microsoft.Compute/virtualMachines/avdtf-1
    $propertiesToOutput.Insert( 0 , @{n='Resource Group' ; e = { $_.resourceMetadata.ResourceId -replace $resourceIdRegex , '$2' }} )
}

[array]$results = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$URI$filter" -retries 1 | Select-Object -ExpandProperty properties -ErrorAction SilentlyContinue | Where-Object category -like $filterCategoryBy )

if( $raw )
{
    $results
}
elseif( $null -eq $results -or $results.Count -eq 0 )
{
    Write-Output -InputObject "No recommendations were returned"
}
else
{
    Write-Output -InputObject "Retrieved $($results.Count) recommendations"

    $results | Select-Object -Property $propertiesToOutput | Sort-Object -Property $sortby | Format-Table -AutoSize
}

