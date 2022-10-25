#require -version 3.0

<#
.SYNOPSIS
    Create one or more Azure VMs with the same size, disks & networking as the source VM

.DESCRIPTION
    Using REST API calls
    
.PARAMETER AZid
    The Azure id of the VM to use as a template for the new VM
    
.PARAMETER AZtenantId
    The azure tenant ID
    
.PARAMETER VMname
    Name of the new VM to create. # characters will be replaced with numbers
    
.PARAMETER localadminUsername
    Name of the local admin account that will be created

.PARAMETER newLocaladminPassword
    Password for the local admin account which will be created

.PARAMETER domainToJoin
    The FQDN of the Active Directory domain to join. If not specific, domain joining will not be performed

.PARAMETER domainJoinUsername
    The pre-existing domain qualified account which will be used to join the new VMs to the domain

.PARAMETER domainJoinPassword
    The password for the account specified for domain joining

.PARAMETER azureADDomainJoin
    Join the VM to Azure AD

.PARAMETER tags
    Tags to assign to the VM in the form tagname=description

.PARAMETER Count
    Number of VMs to create - must have # in the name if more than one

.PARAMETER StartAtNumber
    The number to start multiple machine creation at - default is 1

.PARAMETER maxWaitTimeSeconds
    The maximum number of seconds to wait for the action to complete. If not specified or less than or equal to zero, no waiting will be done

.PARAMETER sleepMilliseconds
    The period to sleep for in milliseconds between calls to get the status of the operation

.EXAMPLE
    & '.\AZ New VM.ps1' -AZid /subscriptions/58ffa3cb-4242-4f2e-a06d-deadbeefdead/resourceGroups/WVD/providers/Microsoft.Compute/virtualMachines/GLW10WVD-0 -azTenantId 77d1b06d-21f7-456b-8311-87c53f6d053c -name GLAZW10## -count 3 -azureADDomainJoin yes -localadminUsername localadminbod -newLocaladminPassword Jobbery6+

    Create 3 VMs using the same size, disks and network as GLW10WVD-0, naming them GLAZW10## where ## is replaced by the next available numbers such as GLAZW1003 and set the local admin username and password. Join to Azure AD once created

.NOTES
    Saved credentials for the user running the script must be available in the file "C:\ProgramData\ControlUp\ScriptSupport\%username%_AZ_Cred.xml" - there is a ControlUp script to create them

    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2022-03-25
    Updated:        2022-09-28 Added validation of domain to join and credentials
                    2022-09-29 Added Azure AD join option via -azureADDomainJoin. Added -tags
                    2022-09-30 Added TLS12 and TLS13 setting
                    2022-10-07 Used updated version of Invoke-AzureRestMethod, fixed tags bug, added check for vmname too long & containing illegal characters
#>

## TODO Copy auto startup/shutdown settings from original

[CmdletBinding()]

Param
(
    [string]$AZid , ## passed by CU as the URL to the VM minus the FQDN ,
    [string]$AZtenantId ,
    [string]$VMname ,
    [string]$localadminUsername ,
    [string]$newLocaladminPassword , ## clear text because of running via ControlUp
    [ValidateSet('yes','no')]
    [string]$deleteOnFail = 'yes' ,
    [ValidateSet('yes','no')]
    [string]$domainValidation = 'yes' ,
    [ValidateSet('yes','no')]
    [string]$azureADDomainJoin = 'no' ,
    [int]$count = 1 ,
    [int]$startAtNumber = 1 ,
    [string[]]$tags ,
    [string]$domainToJoin ,
    [string]$domainJoinUsername ,
    [string]$domainJoinPassword , ## clear text because of running via ControlUp
    [int]$maxWaitTimeSeconds = 0 ,
    [double]$sleepSeconds = 10.0 ,
    [int]$maxmimumVMNameLength = 15
)

#region initialisation
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[datetime]$scriptStartTime = [datetime]::Now

[int]$outputWidth = 250
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

if( [string]::IsNullOrEmpty( $VMname ) -or $VMname.Length -gt $maxmimumVMNameLength )
{
    Throw "VM name `"$VMname`" length is greater than $maxmimumVMNameLength"
}

## Azure resource names cannot contain special characters \/""[]:|<>+=;,?*@&, whitespace, or begin with '_' or end with '.' or '-'

if( $vmname -match '[\\/"\[\]:\|\<\>\+=;,\?\*@&\s]' )
{
    Throw "Illegal characters in VMname `"$VMname`""
}
if( $vmname -match '^_' )
{
    Throw "VMname `"$VMname`" not allowed to begin with _"
}
if( $vmname -match '[\.\-]$' )
{
    Throw "VMname `"$VMname`" not allowed to end with . or -"
}

if( -Not [string]::IsNullOrEmpty( $domainJoinUsername ) -and ( $domainJoinUsername.IndexOf( '\' ) -le 0 -or $domainJoinUsername.IndexOf( '\' ) -ne $domainJoinUsername.LastIndexOf( '\' ) ) )
{
    Throw "Domain joining account must be in domain\username format"
}

if( -Not [string]::IsNullOrEmpty( $domainToJoin ) )
{
    if( $domainToJoin.IndexOf( '.' ) -le 0 )
    {
        Throw "Domain to join must be a FQDN"
    }
    if( [string]::IsNullOrEmpty( $domainJoinUsername ) -or [string]::IsNullOrEmpty( $domainJoinPassword ) )
    {
        Throw "Must specify domain joining account and password for domain $domainToJoin"
    }
}
else
{
    if( -Not [string]::IsNullOrEmpty( $domainJoinUsername ) -or -Not [string]::IsNullOrEmpty( $domainJoinPassword ) )
    {
        Throw "Must specify domain FQDN when specifying a domain joining account and password"
    }
}

## The supplied password must be between 8-123 characters long and must satisfy at least 3 of password complexity requirements from the following:
## 1) Contains an uppercase character ## 2) Contains a lowercase character 3) Contains a numeric digit 4) Contains a special character 5) Control characters are not allowed
if( $newLocaladminPassword.Length -lt 8 -or $newLocaladminPassword.Length -gt 123 )
{
    Throw "Local admin password must be between 8 and 123 characters"
}
[int]$containsLowerCase = $newLocaladminPassword -cmatch '[a-z]'
[int]$containsUpperCase = $newLocaladminPassword -cmatch '[A-Z]'
[int]$containsNumber    = $newLocaladminPassword -match '\d'
[int]$containsSpecial   = $newLocaladminPassword -match '\W'
[bool]$containsControl  = $newLocaladminPassword -match '[\x00-\x1F]' ## ASCII 0 to 31

if( $containsLowerCase + $containsLowerCase + $containsNumber + $containsNumber + $containsSpecial -lt 3 -or $containsControl)
{
    Throw "Local admin password is not complex enough - must contain no control characters and 3 from: upper and lower case letters, digits or special characters"
}

## mandatory parameters best avoided in CU scripts as can cause scripts to hang if missing since willbe promoptinng, siliently, for missing parameters
if( [string]::IsNullOrEmpty( $AZid ) )
{
    Throw "Missing Azure id parameter"
}

if( [string]::IsNullOrEmpty( $localadminUsername ) -or [string]::IsNullOrEmpty( $newLocaladminPassword ) )
{
    Throw "Must provide username and password for the local admin account to create"
}

if( $localadminUsername.IndexOf( '\' ) -ge 0 )
{
    Throw "Invalid local admin user name $localadminUsername"
}

if( $count -le 0 )
{
    Throw "Invalid count $count"
}

[string]$computeApiVersion = '2022-08-01'
[string]$networkApiVersion = '2021-05-01'
[string]$diskApiVersion = '2022-03-02'
[string]$baseURL = 'https://management.azure.com'
[string]$microsoftLoginURL = 'https://login.microsoftonline.com'
[string]$credentialType = 'Azure'

Write-Verbose -Message "AZid is $AZid"

[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13

#endregion initialisation

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
        Purpose:        WVD Administration, through REST API calls
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='Azure Service Principal credentials' )]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCredential] $SPCredentials,

        [Parameter(Mandatory=$true, HelpMessage='Azure Tenant ID' )]
        [ValidateNotNullOrEmpty()]
        [string] $TenantID
    )

    ## https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
    [string]$uri = "$microsoftLoginURL/$TenantID/oauth2/v2.0/token"

    [hashtable]$body = @{
        grant_type    = 'client_credentials'
        client_Id     = $SPCredentials.UserName
        client_Secret = $SPCredentials.GetNetworkCredential().Password
        scope         = "$baseURL/.default"
    }

    [hashtable]$invokeRestMethodParams = @{
        ErrorAction     = 'Continue'
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
            Throw "Exception $($exception.ToString()) originally occurred on line number $($exception.InvocationInfo.ScriptLineNumber) for request $uri ($method)"
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

#endregion AzureFunctions

$azSPCredentials = $null

If (-Not ( $azSPCredentials = Get-AzSPStoredCredentials -system $credentialType -tenantId $AZtenantId ))
{
    ## will already have given errors
    exit 1
}

Write-Verbose -Message "Authenticating to tenant $($azSPCredentials.tenantID) as $($azSPCredentials.spCreds.Username)"
if( -Not ( $azBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID ) )
{
    Throw "Failed to get Azure bearer token"
}

[string]$sourceVMName = ($AZid -split '/')[-1]
    
## get current VM configuration so we can "clone"
    
if( -Not ( $virtualMachine = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid/?api-version=$computeApiVersion" -property $null ) )
{
    Throw "Failed to get VM for $azid"
}
        
## Get network interfaces so we can create for new VM on the same virtual network/subnet

[string]$sourceVMNameEscaped = [regex]::Escape( $sourceVMName )
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

if( $null -ne $tags -and $tags -is [array] -and $tags.Count -eq 1 -and $tags[0].IndexOf( ',' ) -ge 0 )
{
    $tags = @( $tags -split ',' )
}

[hashtable]$tagsToAdd = @{}

ForEach( $tag in $tags )
{
    [string[]]$thisTag = ($tag -split '=' , 2).Trim()
    if( -Not [string]::IsNullOrEmpty( $thisTag[ 0 ] ) )
    {
        $tagsToAdd.Add( $thisTag[ 0 ] , $(if( $thisTag.Count -gt 1 ) { $thisTag[ -1 ] } ) )
    }
}

## may already have been added so allow the user specified ones to prevail
try
{
    $tagsToAdd.Add( 'Created' , "Added by ControlUp Script Action by $env:USERNAME $(Get-Date -Format G) from $azid")
}
catch
{
    ## already have them so let them stay
}

try
{
    $tagsToAdd.Add( 'Creator' , 'ControlUp Script Action' )
}
catch
{
    ## already have them so let them stay
}

## create window that we can update since running inside CU console doesn't output anything until the script finishes or times out

#borrowed from https://learn-powershell.net/2012/10/14/powershell-and-wpf-writing-data-to-a-ui-from-a-different-runspace/
$syncHash = [hashtable]::Synchronized(@{})
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = 'STA'
$newRunspace.ThreadOptions = 'ReuseThread'
$newRunspace.Open()     
$newRunspace.SessionStateProxy.SetVariable( 'syncHash' , $syncHash )

[scriptblock]$code = {
    Param( [string]$title )
    Add-Type -AssemblyName System.Windows.Forms

    $syncHash.Textbox = New-Object -Typename System.Windows.Forms.RichTextBox
    $syncHash.Textbox.Width = 800
    $syncHash.Textbox.Height = 500
    $syncHash.Textbox.Multiline = $true
    $syncHash.Textbox.AutoSize = $true
    $syncHash.Textbox.ReadOnly = $true
    $syncHash.Textbox.Font = New-Object -Typename System.Drawing.Font( $synchash.Textbox.Font.Name, 12 , [System.Drawing.FontStyle]::Bold )

    $syncHash.Form = New-Object -Typename Windows.Forms.Form
    $syncHash.Form.Width = $syncHash.Textbox.Width
    $syncHash.Form.Height = $syncHash.Textbox.Height
    $syncHash.Form.AutoSize = $true
    $syncHash.Form.controls.add( $syncHash.Textbox )
    $syncHash.Form.Text = $title

    $syncHash.Form.add_Closing( {
        $_.Cancel = $false } )

    ## add a variable that we check for to indicate that dialogue has been activated
    [void]$syncHash.Form.Add_Shown({ 
        $syncHash.Form.Activate()
        $syncHash.ReadyToGo = $true })
    $syncHash.Form.Visible = $false
    ##$syncHash.Form.TopMost = $true

    [void]$syncHash.Form.ShowDialog()
}

$powerShellHandle = $null
[hashtable]$runspaceParameters = @{ 'Title' = "Creating $count VMs from $sourceVMName" }

if( $powershell = [PowerShell]::Create().AddScript( $code ) )
{
    [void]$powershell.AddParameters( $runspaceParameters )
    $powershell.Runspace = $newRunspace
    $powerShellHandle = $powershell.BeginInvoke()
}

## wait for dialog to be available
[datetime]$endWait = [datetime]::Now.AddSeconds( 10 )

## wait for dialogue box in runspace to be ready for input
while( -Not $syncHash[ 'ReadyToGo'] -and [datetime]::Now -lt $endWait )
{
    Start-Sleep -Milliseconds 500
}

## if source  VM powered up and domain join requested, check domain can be reached and credentials are valid
if( -Not [string]::IsNullOrEmpty( $domainToJoin ) -and $domainValidation -ieq 'yes' )
{
    $statusMessage = "$(Get-Date -Format T) checking domain $domainToJoin acessibility and credentials"
    Write-Verbose -Message $statusMessage
    if( $syncHash.PSObject.Properties[ 'ReadyToGo'] -and $syncHash.ReadyToGo )
    {
        $syncHash.Textbox.AppendText( $statusMessage + [System.Environment]::NewLine )
    }

    ## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/instance-view
    [string]$instanceViewURI = "$baseURL/$azid/instanceView`?api-version=$computeApiVersion"
    if( $null -eq ( [array]$virtualMachineStatuses = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $instanceViewURI -property 'statuses' ) ) `
        -or $virtualMachineStatuses.Count -eq 0 )
    {
        Throw "Failed to get VM instance view via $instanceViewURI : $_"
    }

    ## code property will be an array so cannot use $matches (https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comparison_operators?view=powershell-5.1)
    ##   PowerState/running
    
    $powerstate = 'unknown'

    if( ( $line = ( $virtualMachineStatuses.code -match 'PowerState/' )) -and ( $powerstate = ($line -split '/' , 2 )[-1] ))
    {
        ## check not already in requested or similar state
        Write-Verbose -Message "Current VM status is $powerstate"

        if( $powerstate -ne 'running' )
        {
            Write-Warning -Message "Power state of VM $vmName is $powerstate so may not be able to check domain $domainToJoin is accessible"
        }
    }
    else
    {
        Write-Warning -Message "Unable to get power state of VM $vmName"
    }

    ##https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/run-command?tabs=HTTP
    [hashtable]$body = @{
        'commandId' = 'RunPowerShellScript'
        'script' = @(
            "Resolve-DnsName -Name $domainToJoin" 
            ## powershell.exe runs in session zero which means we can't check the credentials due to session zero isolation - get "access denied"
            ## cannot use SecureString/PSCredential created locally as encrypted using machine & user so won't decrypt inn remote machine
            "`$credential = New-Object -TypeName pscredential -ArgumentList @( `"$domainJoinUsername`" , (ConvertTo-SecureString -AsPlainText -String `"$domainJoinPassword`" -Force ))"
            "New-PSDrive -Name netlogon$pid -PSProvider FileSystem -Root `"\\$domainToJoin\netlogon`" -Description 'Testing credentials' -scope Script -Credential `$credential"
            ## not persistent so no need to unmap
         )
    }

    [string]$runURI = "$baseURL/$azid/runCommand`?api-version=$computeApiVersion"

    $result = $null
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
                $waitResult = Wait-AsyncAzureOperation -asyncURI $result.Headers[ $asyncOperation ] -azBearerToken $azBearerToken -operation "Run" -maxWaitTimeSeconds 60 -returnProperty output

                if( $null -ne $waitResult -and $false -ne $waitResult )
                {
                    [string]$stdout = $waitResult.value | Where-Object { $_.Code -ieq 'ComponentStatus/StdOut/succeeded' } | Select-Object -ExpandProperty Message -ErrorAction SilentlyContinue
                    [string]$stderr = $waitResult.value | Where-Object { $_.Code -ieq 'ComponentStatus/StdErr/succeeded' } | Select-Object -ExpandProperty Message -ErrorAction SilentlyContinue
                    if( [String]::IsNullOrEmpty( $stdout ) -or -Not [string]::IsNullOrEmpty( $stderr ) )
                    {
                        Throw "Command `"$($body.script)`" in VM $sourceVMName errored : $stderr"
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
}

## if more than one VM being created get all VMs in subscription to ensure we don't duplicate names
## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/list-all
$allExistingVMs = $null 

[string]$padding = $null
[string]$prefix  = $null
[string]$suffix  = $null
[array]$allExistingVMs = @()

if( $count -gt 1 -or $VMname -match '#' )
{  
    if( $vmName -notmatch '^([^#]*)(#+)([^#]*)$' )
    {
        Throw "Must have # characters in machine name when creating more than one"
    }
    $padding = '0' * $Matches[2].Length
    $prefix = $Matches[1]
    $suffix = $Matches[3]

    try
    {
        $allExistingVMs = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/subscriptions/$subscriptionId/providers/Microsoft.Compute/virtualMachines`?api-version=$computeApiVersion" )
    }
    catch
    {
        $allExistingVMs = @()
    }

    if( -Not $allExistingVMs -or $null -eq $allExistingVMs )
    {
        Write-Warning -Message "Failed to get list of VMs for subscription $subscriptionId"
    }
    else
    {
        Write-Verbose -Message "Got $($allExistingVMs.Count) VMs for subscription $subscriptionId"
    }
}

$osdisk = $null
[string]$osDiskId = $virtualMachine.properties.storageProfile.osDisk | Select-Object -ErrorAction SilentlyContinue -ExpandProperty managedDisk | Select-Object -ExpandProperty id
if( [string]::IsNullOrEmpty( $osDiskId ) )
{
    Throw "Failed to get details of OS disk or is not a managed disk" ## TODO do we cater for unmanaged disks ?
}
$osdisk = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$osDiskId`?api-version=$diskApiVersion" -propertyToReturn $null
if( -Not $osdisk )
{
    Throw "Failed to get OS disk $osDiskId from source $sourceVMName"
}

[array]$sourceDataDisks = @( $virtualMachine.properties.storageProfile | Select-Object -ExpandProperty dataDisks -ErrorAction SilentlyContinue | ForEach-Object `
{
    $sourceDataDisk = $_
    $datadiskId = $sourceDataDisk | Select-Object -ExpandProperty managedDisk -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id
    $datadisk = $null
    $datadisk = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$datadiskId`?api-version=$diskApiVersion" -propertyToReturn $null
    if( -Not $datadisk )
    {
        Throw "Failed to get details of data disk $datadiskId from source $sourceVMName"
    }
    else
    {
        ## need details of source data disk and the managed disk
        Add-Member -InputObject $sourceDataDisk -PassThru -NotePropertyMembers @{
            datadiskproperties = ($datadisk | Select-Object -ExpandProperty properties)
            datadiskSKU = $datadisk | Select-Object -ExpandProperty sku
        }
    }
})

Write-Verbose -Message "Got $($sourceDataDisks.Count) source data disks"

[int]$VMnumber = $startAtNumber - 1
[int]$numberLeftToCreate = $count
[string]$asyncOperation = 'Azure-AsyncOperation'
[string]$retryAfter = 'Retry-After'
[string]$newVMURI = $null
[double]$retryPeriodSeconds = $sleepSeconds
$newVMs = New-Object -TypeName System.Collections.Generic.List[object]
$newVMsCreated = New-Object -TypeName System.Collections.Generic.List[string]

[array]$sourceNetworkInterfaces = @( $virtualMachine | Select-Object -ExpandProperty properties | Select-Object -ExpandProperty networkProfile | Select-Object -ExpandProperty networkInterfaces | Select-Object -ExpandProperty id | ForEach-Object `
{
    $networkInterface = $_
    if( -Not ( $thisNetworkInterface = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$networkInterface/?api-version=$networkApiVersion" -property $null ) )
    {
        Throw "Failed to get information for network interface $networkInterface from source $sourceVMName"
    }
    else
    {
        $thisNetworkInterface
    }
})

Write-Verbose -Message "Got $($sourceNetworkInterfaces.Count) source network interfaces"

$signature = @'
    public enum WindowShowStyle : uint
    {
        Hide = 0,
        ShowNormal = 1,
        ShowMinimized = 2,
        ShowMaximized = 3,
        Maximize = 3,
        ShowNormalNoActivate = 4,
        Show = 5,
        Minimize = 6,
        ShowMinNoActivate = 7,
        ShowNoActivate = 8,
        Restore = 9,
        ShowDefault = 10,
        ForceMinimized = 11
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, WindowShowStyle nCmdShow);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
'@

Add-Type -MemberDefinition $signature -Name 'Windows' -Namespace Win32Functions -Debug:$false

[hashtable]$domainjoinbody = @{}

if( -Not [string]::IsNullOrEmpty( $domainToJoin ) -and -Not [string]::IsNullOrEmpty( $domainJoinUsername ) -and -Not [string]::IsNullOrEmpty( $domainJoinPassword ) )
{
    if( $azureADDomainJoin -ieq 'yes' )
    {
        Throw 'Cannot do Azure AD join and join domain $domainToJoin'
    }

    $domainjoinbody = @{
        "properties" = @{
                "publisher" = "Microsoft.Compute"
                "type" = "JsonADDomainExtension"
                "typeHandlerVersion" = "1.3"
                "autoUpgradeMinorVersion" = $true
                "settings" = @{
                    "Restart" = $true
                    "Options" = 3 ## join domain AND create computer account
                    "Name" = $domainToJoin
                    "User" = $domainJoinUsername
                    ## TODO do we need to get OU ?
                }
                "protectedSettings" = @{
                    "Password" = $domainJoinPassword
                }
            }
        "location" = $virtualMachine.location
    }
}
elseif( $azureADDomainJoin -ieq 'yes' )
{
    $domainjoinbody = @{
        "properties" = @{
                "publisher" = "Microsoft.Azure.ActiveDirectory"
                "type" = "AADLoginForWindows"
                "typeHandlerVersion" = "1.0"
                "autoUpgradeMinorVersion" = $true
                "settings" = @{
                    "mdmId" = ""
                }
            }
        "location" = $virtualMachine.location
    }
}

#<#
while( $numberLeftToCreate-- -gt 0 )
{
    ++$VMnumber
    [string]$thisVMname = $VMname

    [string]$tryThisVM = $null
    do
    {
        if( $VMname -match '#' )
        {
            $tryThisVM = "{0}{1:$padding}{2}" -f $prefix , $VMnumber , $suffix
        }
        else
        {
            $tryThisVM = $VMname
        }

        if( $VMname -match '#' -and ( $existingVM = $allExistingVMs | Where-Object Name -eq $tryThisVM ) )
        {
            $VMnumber++
        }
        else ## double check doesn't exist
        {
            ## URI will be the same as for the source VM except for the name
            $newVMURI = "$baseURL/$azid`?api-version=$computeApiVersion" -replace "/$($sourceVMNameEscaped)\?" , "/$($tryThisVM)?"

            ## check new VM does not already  exist
            $existing = $null
            try
            {
                $existing = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $newVMURI -propertyToReturn $null
            }
            catch
            {
                ## we are expecting this to fail as we don't want the VM to already exist
            }
    
            if( -Not $existing )
            {
                break
            }
            else
            {
                if( $VMname -match '#' )
                {
                    Write-Warning -Message "New VM $tryThisVM unexpectedly already exists" ## wasn't in our original list of existing VMs
                    $VMnumber++
                }
                else
                {
                    $tryThisVM = $null
                    Throw  "New VM $VMname already exists"
                }
            }
        }
    } while( $existingVM )

    if( $tryThisVM )
    {
        $thisVMname = $tryThisVM
    }
    else
    {
        break
    }

    ## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/create-or-update
    ## create deep copy
    $newVMbody = $virtualMachine | ConvertTo-Json -Depth 99 | ConvertFrom-Json
    
    ## remove disk details as they are not detailed enough, eg storage type
    @( 'osDisk', 'dataDisks' ) | ForEach-Object { $newVMbody.properties.storageProfile.psobject.Properties.Remove( $_ ) }

    ## need to analyse each disk so that we duplicate the type, eg HDD, standard or premium SSD
    ## https://docs.microsoft.com/en-us/rest/api/compute/disks/get
    
    [int]$disks = 0
    [int]$nics = 0

    if( $osDisk )
    {
        Add-Member -InputObject $newVMbody.properties.storageProfile -MemberType NoteProperty -Name osdisk -value @{
            #'tags'         = $tagsToAdd
            'caching'      = $virtualMachine.properties.storageProfile.osDisk.caching
            'createOption' = $virtualMachine.properties.storageProfile.osDisk.createOption
            'deleteOption' = 'delete' ## don't want to be blamed for leving orphaned disks if the VM is deleted
            ##'encryptionSettings' = $virtualMachine.properties.storageProfile.osDisk.encryptionSettings
            'managedDisk' = @{
                'storageAccountType' = $osdisk | Select-Object -ExpandProperty sku | Select-Object -ExpandProperty name
            }
        }
        $disks++
    }

    [int]$lun = 0

    [array]$dataDisks = @( ForEach( $sourceDataDisk in $sourceDataDisks )
        {
        [pscustomobject]@{
            #'tags'         = $tagsToAdd
            'caching'      = $sourceDataDisk.caching
            'diskSizeGB'   = $sourceDataDisk | Select-Object -ExpandProperty datadiskproperties -ErrorAction SilentlyContinue | Select-Object -ExpandProperty diskSizeGB
            'lun'          = $lun++ ## don't seem to be able to get this so let's hope it (order/number) doesn't matter
            'createOption' = 'empty' ## we are not cloning
            'deleteOption' = 'delete' ## don't want to be blamed for leaving orphaned disks if the VM is deleted
            ##'encryptionSettings' = $virtualMachine.properties.storageProfile.osDisk.encryptionSettings
            'managedDisk' = @{
                'storageAccountType' = $sourceDataDisk | Select-Object -ExpandProperty datadiskSKU -ErrorAction SilentlyContinue | Select-Object -ExpandProperty name
            }
        }
    })

    if( $null -ne $dataDisks -and $dataDisks.Count -gt 0 )
    {
        $disks += $dataDisks.Count
        Add-Member -InputObject $newVMbody.properties.storageProfile -MemberType NoteProperty -Name datadisks -Value $dataDisks
    }

    $newNetworkInterfaces = New-Object -TypeName System.Collections.Generic.List[object]

    ForEach( $thisNetworkInterface in $sourceNetworkInterfaces )
    {
        if( $IPproperties = $thisNetworkInterface.properties.ipConfigurations | Select-Object -ExpandProperty properties )
        {
            ## check using DHCP otherwise we will not get involved
            if( $IPproperties.psobject.properties[ 'privateIPAllocationMethod' ] -and $IPproperties.privateIPAllocationMethod -ne 'Dynamic' )
            {
                Throw "Network interface $($thisNetworkInterface.id) is not DHCP so unable to replicate"
            }
            ## we don't know where, if at all, the source machine name will be in the network interface name other than will be in the last / portion so we look for the machine name and replace it
            $newNetworkInterfaceId = $thisNetworkInterface.id -replace "$sourceVMNameEscaped([^/]*$)" , "$thisVMname$`1"
            if( $newNetworkInterfaceId -eq $thisNetworkInterface.Id )
            {
                Throw "Unable to create new network interface name from $($thisNetworkInterface.id)"
            }

            Write-Verbose -Message "New network interface is $newNetworkInterfaceId"
            
            [hashtable]$body = @{}
            $thisNetworkInterface.psobject.properties | ForEach-Object { $body.Add( $_.Name , $_.Value ) }
            ## remove unwanted properties either because for source or because created for us
            @( 'type' , 'tags' , 'id' , 'name' , 'resources' , 'etag' ) | ForEach-Object { if( $body.ContainsKey( $_ ) ) { $body.Remove( $_ ) } }
            @( 'resourceGuid', 'provisioningState' , 'macAddress' , 'virtualmachine' , 'etag' ) | ForEach-Object { if( $body.properties.psobject.properties[ $_ ] ) { $body.properties.psobject.Properties.Remove( $_ ) } }

            $body.properties.ipConfigurations.properties.psobject.Properties | Where-Object Name -NotIn @( 'subnet' , 'primary' , 'IPAllocationMethod' ) | ForEach-Object { $body.properties.ipConfigurations.properties.psobject.Properties.Remove( $_.Name ) }
            $body.Add( 'Tags' , $tagsToAdd )

            ## https://docs.microsoft.com/en-us/rest/api/virtualnetwork/network-interfaces/create-or-update
            $result = $null
            $result = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL$newNetworkInterfaceId`?api-version=$networkApiVersion" -propertyToReturn $null -method PUT -body $body
            if( $result )
            {
                $newNetworkInterfaces.Add( ([pscustomobject]@{ id = $result.id ; properties = @{  "deleteOption" = 'delete' } } ))
            }
            else
            {
                Write-Warning -Message "Problem creating network interface $newNetworkInterfaceId"
            }
        }
        else
        {
            Write-Warning -Message "No IP properties for source network interface $($thisNetworkInterface.Id)"
        }
    }
    
    if( $null -eq $newNetworkInterfaces -or $newNetworkInterfaces.Count -eq 0 )
    {
        Throw "No new network interfaces created, cannot proceed - check network interfaces and IP configuration of source VM"
    }

    $nics += $newNetworkInterfaces.Count
    $windowsConfiguration = $newVMbody.properties | Select-Object -ExpandProperty osProfile -ErrorAction SilentlyContinue | Select-Object -ExpandProperty windowsConfiguration -ErrorAction SilentlyContinue

    @( 'type' , 'tags' , 'id' , 'name' , 'resources' ) | ForEach-Object { $newVMbody.psobject.Properties.Remove( $_ ) }
    @( 'vmId', 'provisioningState' ) | ForEach-Object { $newVMbody.properties.psobject.Properties.Remove( $_ ) }
    $newVMbody.properties.psobject.Properties.Remove( 'osprofile' )
    ## for Azure AD joined, this needs removing otherwise get "resource's Identity property must be null or empty for 'SystemAssigned' principalid"
    if( $newVMbody -and $newVMbody.psobject.properties[ 'identity' ] -and $newVMbody.identity.psobject.properties[ 'type' ] -and $newVMbody.identity.type -ieq 'SystemAssigned' )
    {
        $newVMbody.identity.psobject.Properties.Remove( 'principalId' )
        $newVMbody.identity.psobject.Properties.Remove( 'tenantId' )
    }

    [hashtable]$osprofile = @{
            'ComputerName'  = $thisVMname
            'adminUsername' = $localadminUsername
            'adminPassword' = $newLocaladminPassword
        }
    if( $windowsConfiguration )
    {
        $osprofile.Add( 'windowsConfiguration' , $windowsConfiguration )
    }

    Add-Member -InputObject $newVMbody.properties -MemberType NoteProperty -Name osprofile -Value $osprofile
    if( $tagsToAdd.Count -gt 0 )
    {
        Add-Member -InputObject $newVMbody -MemberType NoteProperty -Name 'Tags' -Value $tagsToAdd
    }

    $newVMbody.properties.networkProfile.networkInterfaces = $newNetworkInterfaces
    
    ## TODO can we replicate auto start and stop settings?

    $result = $null
    ## this is an async operation so we need the response headers to get the URI to check the status

    [string]$statusMessage = "$(Get-Date -Format T) starting creation of $thisVMname with $disks disks and $nics network interfaces"
    
    Write-Output -InputObject $statusMessage
    
    $syncHash.Textbox.AppendText( $statusMessage + [System.Environment]::NewLine )

    try
    {
        $result = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $newVMURI -property $null -method PUT -body $newVMbody -norest
    }
    catch
    {
        $null
    }

    ## see if the response header has a retry property and if it is greater than what is passed on the command line, or the default, use that instead
    if( $result.headers[ $retryAfter ])
    {
        [double]$thisRetryPeriodSeconds = $result.headers[ $retryAfter ]
        if( $thisRetryPeriodSeconds -gt $retryPeriodSeconds )
        {
            $retryPeriodSeconds = $thisRetryPeriodSeconds
        }
    }

    if( $result )
    {
        Write-Verbose -Message "REST result $($result.StatusCode)"
    
        if( $result.StatusCode -ge 200 -and $result.StatusCode -lt 400 ) ## accepted/created
        {
            if( -Not $result.Headers -or -Not $result.Headers[ $asyncOperation ] )
            {
                Write-Warning -Message "Result headers missing $asyncOperation so unable to monitor"
            }
            else
            {
                Write-Verbose -Message "VM $thisVMname creation submitted ok"
                ## insert new VMs at start so when we iterate and start at end, the first created are processed first
                $newVMs.Insert( 0 , ([pscustomobject]@{
                    VMname = $thisVMname
                    NewVMUri = $newVMURI
                    Async = $result.headers[ $asyncOperation ]
                    StartTime = [datetime]::Now
                    DomainJoin = $(if( $domainjoinbody -and $domainjoinbody.Count ) { 'Not Started' } else { 'Not Doing' })
                    NetworkInterfaces = $newNetworkInterfaces.ToArray() ## so we can delete them if VM creation fails
                }))
            }
        }
        else
        {
            Write-Warning -Message "Unexpected status $($result | Select-Object -ExpandProperty StatusCode) returned - check Azure to see if VM $thisVMname created"
        }
    }
}

## wait for VM creations to finish so we can domain join if required
do
{
    Write-Verbose -Message "$(get-Date -Format G): checking $($newVMs.Count) new VMs"
    For( [int]$index = $newVMs.Count - 1 ; $index -ge 0 ; $index-- ) ## work backwards so we can delete finished/failed elements
    {
        [bool]$removeElement = $true ## will only set to false and not remove element if operations still in progress and not timed out
        $thisVMname = $newVMs[ $index ].VMname
        Write-Verbose -Message "$(Get-Date -Format G): checking $thisVMname, operation started at $(Get-Date -Date ($newVMs[ $index ].StartTime) -Format G), domain join $($newVMs[ $index ].DomainJoin)"

        if( $newVMs[ $index ].Async ) ## either VM creation or domain join in progress
        {
            $status = $null
            $status = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $newVMs[ $index ].Async -property $null -method GET -Verbose:$false
            if( $status -and $status.status -ine 'InProgress' )
            {
                if( $newVMs[ $index ].DomainJoin -match '^Not' )
                {
                    Write-Verbose -Message "Finished creation of $thisVMname with status $($status.status)"
                    if( $status.status -ieq 'Succeeded' )
                    {
                        if( $domainjoinbody.Count -gt 0 )
                        {
                            Write-Verbose -Message "Joining $thisVMname to domain $domainToJoin"       
                            $result = $null
                            [string]$domainJoinURI = $newVMs[ $index ].NewVMUri -replace '\?' , "/extensions/CUDomainJoin?"
                            $result = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $domainJoinURI -property $null -method PUT -body $domainjoinbody -norest -Verbose:$false
                            if( $result )
                            {
                                if( $result.StatusCode -ge 200 -and $result.StatusCode -lt 400 ) ## accepted/created
                                {
                                    if( -Not $result.Headers -or -Not $result.Headers[ $asyncOperation ] )
                                    {
                                        Write-Warning -Message "Result headers missing $asyncOperation so unable to monitor $thisVMname domain join"
                                    }
                                    else
                                    {
                                        $newVMs[ $index ].Async = $result.Headers[ $asyncOperation ]
                                        $newVMs[ $index ].DomainJoin = 'Started'
                                        $newVMs[ $index ].StartTime = [datetime]::Now
                                        $removeElement = $false

                                        $statusMessage = "$(Get-Date -Format T) started domain join of $thisVMname"
                                        $syncHash.Textbox.AppendText( $statusMessage + [System.Environment]::NewLine )

                                    }
                                }
                                else
                                {
                                    Write-Warning -Message "Unexpected status $($result | Select-Object -ExpandProperty StatusCode) returned - check to see if VM $thisVMname gets domain joined"
                                }
                            }
                            else
                            {
                                Write-Warning -Message "Failed to submit task to join $thisVMname to domain"
                            }
                        }
                        ## else not domain joining
                    }
                    elseif( $deleteOnFail -ieq 'yes' )
                    {                           
                        ForEach( $newNetworkInterface in $newVMs[ $index ].NetworkInterfaces )
                        {
                            Write-Verbose -Message "Deleting NIC $($newNetworkInterface.Id)"
                            ## network interface could have been associated with the VM so we must dissociate it first before we can delete

                            ##https://learn.microsoft.com/en-us/rest/api/virtualnetwork/network-interfaces/delete?tabs=HTTP
                            $deletion = $null
                            try
                            {
                                $deletion = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$($newNetworkInterface.Id)/?api-version=$networkApiVersion" -property $null -method DELETE -norest
                            }
                            catch
                            {
                                ## try deleting VM instead in case it actually made it
                                Write-Warning -Message "Problem deleting network interface $($newNetworkInterface.Id) : $($_ | Select-Object -ExpandProperty ErrorDetails | Select-Object -ExpandProperty Message)"
                                $deleteVM = $null
                                $deleteVM = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $newVMs[ $index ].NewVMUri -method DELETE -propertyToReturn $null -norest
                                if( -Not $deleteVM )
                                {
                                    Write-Warning -Message "Problem submitting delete operation for $($newVMs[ $index ].VMname)"
                                }
                                elseif( $deleteVM.StatusCode -lt 200 -or $deleteVM.StatusCode -ge 400 )
                                {
                                    Write-Warning -Message "Unexpected status $($deleteVM.StatusCode) on submitting delete operation for $($newVMs[ $index ].VMname)"
                                }
                                $deletion = 'Deleted Parent VM'
                            }
                            if( -Not $deletion )
                            {
                                Write-Warning -Message "Problem deleting network interface $($newNetworkInterface.Id)"
                            }
                            elseif( $deletion.StatusCode -lt 200 -or $deletion.StatusCode -ge 400 )
                            {
                                Write-Warning -Message "Unexpected status $($deletion.StatusCode) deleting network interface $($newNetworkInterface.Id) for $($newVMs[ $index ].VMname)"
                            }
                        }
                    }
                    ## else VM creation not succeeded
                }
                elseif( $newVMs[ $index ].DomainJoin -eq 'Started' )
                {
                    Write-Verbose -Message "Finished domain join of $thisVMname with status $($status.status)"
                    if( $status.status -ine 'Succeeded' )
                    {
                        if( $deleteOnFail -ieq 'yes' )
                        {
                            $deleteVM = $null
                            $deleteVM = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $newVMs[ $index ].NewVMUri -method DELETE -propertyToReturn $null -norest
                            if( -Not $deleteVM )
                            {
                                Write-Warning -Message "Problem submitting delete operation for $($newVMs[ $index ].VMname)"
                            }
                            elseif( $deleteVM.StatusCode -lt 200 -or $deleteVM.StatusCode -ge 400 )
                            {
                                Write-Warning -Message "Unexpected status $($deleteVM.StatusCode) on submitting delete operation for $($newVMs[ $index ].VMname)"
                            }
                        }
                        ## TODO try and get log file so we can show why it failed "C:\WindowsAzure\Logs\Plugins\Microsoft.Compute.JsonADDomainExtension\1.3.9\ADDomainExtension.log"
                    }
                }
            }
            elseif( $newVMs[ $index ].StartTime.AddSeconds( $maxWaitTimeSeconds ) -gt [datetime]::Now )
            {
                Write-Warning -Message "Time out on operation on $thisVMname"
            }
            else ## still InProgress and not timed out
            {
                $removeElement = $false
            }
        }
        ## else no async handle so nothing to do

        if( $removeElement )
        {
            $newVMsCreated.Add( $thisVMname )
            Write-Verbose -Message "Removing element $thisVMname at $index"
            $newVMs.RemoveAt( $index )
        }
    }

    if( $newVMs.Count -gt 0 )
    {
        Start-Sleep -Seconds $retryPeriodSeconds
    }
    else
    {
        break
    }
} while( $newVMs -and $newVMs.Count -gt 0 )

Write-Verbose -Message "All finished at $(Get-Date -Format G)"

$statusMessage = "$(Get-Date -Format T) created $($newVMsCreated.Count) VMs in $([math]::Round( ([datetime]::now - $scriptStartTime).TotalMinutes , 1 )) minutes $($newVMsCreated -join ' , ')"

Write-Output -InputObject $statusMessage
$syncHash.Textbox.AppendText( $statusMessage + [System.Environment]::NewLine )

## don't wait for form to be closed since what is in GUI will be in the SBA output window, it was only a progress window
if( $syncHash -and $syncHash[ 'Form' ] )
{
    $syncHash.Form.Close()
    if( $powershell )
    {
        $powerShell.Dispose()
        $powershell = $null
    }
    if( $newRunspace )
    {
        $newRunspace.Close()
        $newRunspace.Dispose()
        $newRunspace = $null
    }
    $syncHash = $null
}

