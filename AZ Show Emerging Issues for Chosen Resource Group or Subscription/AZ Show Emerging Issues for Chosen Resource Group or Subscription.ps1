#require -version 3.0

<#
.SYNOPSIS
    Get Azure emerging issues list and show detail if the locations impacted in this resource group or subscription

.DESCRIPTION
    Using REST API calls

.PARAMETER azid
    The relative URI of the Azure VM
    
.PARAMETER AZtenantId
    The azure tenant ID
    
.PARAMETER resourceGroupOnly
    Cross reference emerging issues with the locations used by resources only in the resource group containg the azif passed, otherwise do it for the entire subscription.

.NOTES
    Version:        1.2
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2022-07-29
    Updated:        2022-08-30 Fixed resource location scoping issue
#>

[CmdletBinding()]

Param
(
    [string]$AZid ,## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId ,
    [ValidateSet('Yes','No')]
    [string]$resourceGroupOnly = 'No'
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 250
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
    ## not a showstopper, will just may not be wide enough to stop output wrapping
}

[string]$computeApiVersion = '2021-07-01'
[string]$resourceHealthApiVersion = '2018-07-01' ## '2020-05-01'
[string]$resourceManagementApiVersion = '2021-04-01'
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

## get the resource so we know what region(s) are relevant
## TODO do we need to get the resource provider so we know what API version to use?
## https://management.azure.com/subscriptions/58ffa3cb-2f63-4f2e-a06d-369c1fcebbf5/providers?api-version=2021-04-01
$resource =  Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$AZid`?api-version=2021-04-01" -propertyToReturn $null -retries 2
if( -not $resource )
{
    Throw "Failed to retrieve $AZid"
}

## get resource group so can use location as may be different to the resource passed
$resourceGroup = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/resourcegroups/$resourceGroupName`?api-version=$resourceManagementApiVersion" -propertyToReturn $null -retries 2

## get locations and build look up table because region ids returned by emerging issue calls are different to regions used in resources so we need to corrleate via display name
[array]$locations = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/locations`?api-version=$resourceManagementApiVersion" -retries 2 )

[hashtable]$locationLookup = @{}

ForEach( $location in $locations )
{
    $locationLookup.Add( $location.DisplayName , $location.name )
}

## https://docs.microsoft.com/en-us/rest/api/resources/resources/list-by-resource-group
[string]$resourcesURL = $null
if( $resourceGroupOnly -ieq 'yes' )
{
    $resourcesURL = "resourceGroups/$resourceGroupName"
}
## else ## will be for the subscription

[array]$allResources = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/$resourcesURL/resources`?`$expand=createdTime,changedTime,lastusedTime&api-version=$resourceManagementApiVersion" -retries 2 )

## get all locations that we have resources in
[hashtable]$ourResourceLocations = @{}
ForEach( $resource in $allResources )
{
    try
    {
        ## TODO do we need to record type(s) of resource too?
        $ourResourceLocations.Add( $resource.location , $true )
    }
    catch
    {
        ## already got it
    }
}

Write-Verbose -Message "Got $($locationLookup.Count) locations & $($ourResourceLocations.Count) locations with our resources in"

## https://docs.microsoft.com/en-us/rest/api/resourcehealth/emerging-issues/list?tabs=HTTP

[array]$emergingIssues = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/providers/Microsoft.ResourceHealth/emergingIssues`?api-version=$resourceHealthApiVersion" -propertyToReturn 'value' -retries 2 )

[array]$activeEvents = @( $emergingIssues | Select-Object -ExpandProperty properties | Select-Object -ExpandProperty statusActiveEvents )

Write-Verbose -Message "Got $($emergingIssues.Count) emerging issues, $($activeEvents.Count) active events"

[hashtable]$regionsWithIssues = @{}

[string]$ourRegions = ($ourResourceLocations.GetEnumerator() | Select-Object -ExpandProperty Name | Sort-Object ) -join ', '

if( $activeEvents -and $activeEvents.Count -gt 0 )
{
    [array]$relevantEvents = @( ForEach( $activeEvent in $activeEvents )
    {
        [bool]$relevant = $false
        ForEach( $impact in $activeEvent.impact )
        {
            if( -Not $relevant ) ## only want 1 event even if multiple impacts
            {
                ForEach( $region in $impact.Regions )
                {
                    [string]$locationId = $locationLookup[ $region.name ]
                    if( $region -ieq 'global' -or $ourResourceLocations[ $locationId ] -or $ourResourceLocations[ $region.id ] )
                    {
                        $activeEvent
                        $relevant = $true
                        break ## only need 1 instance of this impact
                    }
                    else
                    {
                        try
                        {
                            $regionsWithIssues.Add( $region.name , $true )
                        }
                        catch
                        {
                            ## already got it
                        }
                    }
                }
            }
        }
    })
    
    Write-Output -InputObject "There are currently $($relevantEvents.Count) emerging issues in the $($ourResourceLocations.Count) region(s) used ($ourRegions) Azure  out of $($emergingIssues.Count) issues total"

    if( $relevantEvents -and $relevantEvents.Count -gt 0 )
    {
        ## strip <p> </p> tags
        $relevantEvents | Select-Object @{n='Start';e={ Get-Date -Date $_.startTime -Format G}},@{n='Last Update';e={ Get-Date -Date $_.lastModifiedTime -Format G}},cloud,severity,stage,title,@{n='Description';e={ $_.description -replace '\</?[a-z]{1,4}\>' }}| Format-Table -AutoSize -Wrap
    }
    elseif( $regionsWithIssues.Count -gt 0 )
    {
        Write-Output -InputObject "`nThe $($regionsWithIssues.Count) regions with emerging issues are:"
        $regionsWithIssues.GetEnumerator()| Select-Object -Property @{n='Indent';e={'    '}},@{n='Region';e={$_.Key}} | Sort-Object -Property Region | Format-Table -AutoSize -HideTableHeaders
    }
}
else
{
    Write-Output -InputObject "There are currently no emerging Azure issues"
}

