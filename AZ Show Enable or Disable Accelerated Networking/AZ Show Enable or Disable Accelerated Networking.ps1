#require -version 3.0

<#
.SYNOPSIS
   Get, enable or disable accelrated networking for the given VM

.DESCRIPTION
    Using REST API calls
    
.PARAMETER AZid
    The Azure id of the VM to use as a template for the new VM
    
.PARAMETER AZtenantId
    The azure tenant ID
    
.PARAMETER operation
    The operation to perform on the network interfaces
    
.PARAMETER waitTimeoutSeconds
    Number of seconds to wait for the change in accelerated networking to apply

.EXAMPLE
    & 'C:\AZ accelerated networking.ps1' -VM GLHZAV42 -Operation enable

    Enable accelerated networking on the specified VM

.NOTES
    Saved credentials for the user running the script must be available in the file "C:\ProgramData\ControlUp\ScriptSupport\%username%_AZ_Cred.xml" - there is a ControlUp script to create them

    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2022-10-17
    Updated:        2022-11-08 Changed message when NIC already in the new state
#>

[CmdletBinding()]

Param
(
    [string]$AZid , ## passed by CU as the URL to the VM minus the FQDN ,
    [string]$AZtenantId ,
    [ValidateSet('report','enable','disable')]
    [string]$operation = 'report' ,
    [decimal]$waitTimeoutSeconds = 60
)

#region initialisation
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[datetime]$scriptStartTime = [datetime]::Now

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
    ## nothing we can do but shouldn't be a show stopper
}

## mandatory parameters best avoided in CU scripts as can cause scripts to hang if missing since willbe promoptinng, siliently, for missing parameters
if( [string]::IsNullOrEmpty( $AZid ) )
{
    Throw "Missing Azure id parameter"
}

[string]$computeApiVersion = '2022-08-01'
[string]$networkApiVersion = '2022-07-01'
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


Function Wait-ProvisioningComplete
{
    [CmdletBinding()]

    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$bearerToken ,
        [Parameter(Mandatory=$true)]
        [string]$uri ,
        [int]$sleepMilliseconds = 3000 ,
        [int]$waitForSeconds = 60 
    )

    [datetime]$start = [datetime]::Now
    [datetime]$end = $start.AddSeconds( $waitForSeconds )

    do
    {
        if( ( $state = Invoke-AzureRestMethod -BearerToken $bearerToken -uri $uri -property 'properties') )
        {
            if( -Not $state.PSObject.properties[ 'provisioningState' ] )
            {
                Write-Warning -Message "No state property on response from $uri - $state"
            }
            elseif( $state.provisioningState -eq 'Succeeded' )
            {
                break
            }
            elseif( $state.provisioningState -eq 'Failed' )
            {
                Write-Error -Message "Provisioning failed for $uri"
                break
            }
        }
        else
        {
            Write-Warning -Message "Failed call to $uri"
        }
        Write-Verbose -Message "$(Get-Date -Format G) : provisioning state of $uri is $($state | Select-Object -ExpandProperty provisioningState -ErrorAction SilentlyContinue) so waiting $sleepMilliseconds ms"
        Start-Sleep -Milliseconds $sleepMilliseconds
    } while( [datetime]::Now -le $end )

    if( $state -and $state.PSObject.properties[ 'provisioningState' ] -and $state.provisioningState -eq 'Succeeded' )
    {
        $state ## return
    }
    ## else not succeeded so implicitly return $null
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

if( -Not ( $virtualMachine = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid/?api-version=$computeApiVersion" -property $null ) )
{
    Throw "Failed to get VM for $azid"
}
        
## Get network interfaces so we can create for new VM on the same virtual network/subnet

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

[int]$counter = 0
[hashtable]$NICnames = @{}

[array]$sourceNetworkInterfaces = @( $virtualMachine | Select-Object -ExpandProperty properties | Select-Object -ExpandProperty networkProfile | Select-Object -ExpandProperty networkInterfaces | Select-Object -ExpandProperty id | ForEach-Object `
{
    $networkInterface = $_
    if( -Not ( $thisNetworkInterface = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$networkInterface/?api-version=$networkApiVersion" -property $null ) )
    {
        Throw "Failed to get information for network interface $networkInterface from source $sourceVMName"
    }
    else
    {
        if( $thisNetworkInterface.properties.provisioningState -ieq 'Succeeded' )
        {
            $NICnames.Add( $thisNetworkInterface.Name , "NIC$counter" )
            $counter++
            $thisNetworkInterface
        }
        else
        {
            Write-Warning -Message "Accelerated Networking cannot be determined as network interface $($thisNetworkInterface.Name) is in state $($thisNetworkInterface.properties.provisioningState)"
        }
    }
})

if( $null -eq $sourceNetworkInterfaces -or $sourceNetworkInterfaces.Count -eq 0 )
{
    Throw "No network interfaces found for $($virtualMachine.Name)"
}

Write-Verbose -Message "Got $($sourceNetworkInterfaces.Count) network interfaces for $($virtualMachine.Name)"

if( $operation -ieq 'report' )
{
    if( $sourceNetworkInterfaces.Count -gt 1 )
    {
        $sourceNetworkInterfaces | Select-Object -Property @{n='Network Interface';e={ $NICnames[ $_.name ] }},@{n='Accelerated Networking';e={ if( $_.properties.enableAcceleratedNetworking ) { 'enabled' } else { 'disabled'  } }} | Format-Table -AutoSize
        "Please run the 'AZ Show machine network information' to show more details"
    }
    else
    {
        Write-Output "Accelerated networking is currently $(if( $sourceNetworkInterfaces[0].properties.enableAcceleratedNetworking ) { 'enabled' } else { 'disabled'  } )"
    }
}
else ## enable/disable
{
    [int]$changed = 0
    [bool]$newState = $operation -ieq 'enable'
    [int]$alreadyInDesiredState = 0

    ForEach( $networkInterface in $sourceNetworkInterfaces )
    {
        if( $networkInterface.properties.enableAcceleratedNetworking -ieq $newState )
        {
            $alreadyInDesiredState++
        }
        else
        {
            $networkInterface.properties.enableAcceleratedNetworking = $newState ## TODO could it be missing so we need to add?
            ## https://docs.microsoft.com/en-us/rest/api/virtualnetwork/network-interfaces/create-or-update
            $nicUpdate = $null
            [string]$networkInterfaceURI = "$baseURL/$($networkInterface.id)/?api-version=$networkApiVersion"
            try
            {
                $nicUpdate = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $networkInterfaceURI -body $networkInterface -property $null -method PUT -rawException
            }
            catch
            {
                ## Virtual machine /subscriptions/58ffa3cb-2f63-4f2e-a06d-369c1fcebbf5/resourceGroups/W1ndows365/providers/Microsoft.Compute/virtualMachines/GLAZADW1001 has size Standard_B2s, which is not compatible with enabling accelerated networking on network interface(s) on the VM. Compatible VM sizes: Standard_D2_v4, Standard_D2s_v4,
                if( ( $errorCodeAndMessage = $_.ErrorDetails | Select-Object -ExpandProperty Message -ErrorAction SilentlyContinue | ConvertFrom-Json | Select-Object -ExpandProperty error -ErrorAction SilentlyContinue) )
                {
                    if( $errorCodeAndMessage.code -ieq 'VMSizeIsNotPermittedToEnableAcceleratedNetworking' `
                        -and $errorCodeAndMessage.Message -match "$azid has size ([a-z_0-9]+),\s*(.+)Compatible VM sizes:" )
                    {
                        Write-Warning -Message "VM is size $($matches[1]) $($Matches[2])"
                        break ## is VM size related so will fail for any other NICs
                    }
                    else
                    {
                        ## SizeIsNotPermittedToEnableAcceleratedNetworking
                        [string]$errorString = ($errorCodeAndMessage.Code -csplit '(?=[A-Z])' -join ' ').Trim()
                        Write-Warning -Message "Problem trying to $operation on $($networkInterface.Name) : $errorString"
                    }
                }
                else
                {
                    Write-Warning "Error when trying to change accelerated networking setting on network interface $($networkInterface.name) : $_"
                }
                $nicUpdate = $null ## just in case it is set to something so we don't do checks/output below
            }
            if( $nicUpdate )
            {
                [string]$state = $nicUpdate | Select-Object -ExpandProperty properties -ErrorAction SilentlyContinue | Select-Object -ExpandProperty provisioningState -ErrorAction SilentlyContinue
                if( $state -ine 'Updating' -and $state -ine 'Succeeded' )
                {
                    Write-Warning -Message "Unexpected state `"$state`" returned from updating NIC $($networkInterface.Name)"
                }
                else
                {
                    [string]$prefix = $null
                    if( $sourceNetworkInterfaces.Count -gt 1 )
                    {
                        $prefix = $NICnames[ $networkInterface.name ] + ': '
                    }
                    $provisioningState = $null
                    $provisioningState = Wait-ProvisioningComplete -bearerToken $azBearerToken -uri $networkInterfaceURI -waitForSeconds $waitTimeoutSeconds
                    if( -not $provisioningState -or $provisioningState.provisioningState -ine 'succeeded' )
                    {
                        if( $provisioningState.provisioningState -ieq 'Updating' )
                        {
                            Write-Warning -Message "$($prefix)Accelerated networking provisioning is still updating"
                        }
                        else
                        {
                            Write-Error -Message "$($prefix)Accelerated networking provisioning encountered a problem : $($provisioningState.provisioningState)"
                        }
                    }
                    else
                    {
                        Write-Output -InputObject "$($prefix)Accelerated networking is now set to $($operation)d"
                    }
                }
                $changed++
            }
        }
        $counter++
    }

    if( $alreadyInDesiredState -gt 0 )
    {
        if( $sourceNetworkInterfaces.Count -eq 1 )
        {
            Write-Warning -Message "Accelerated networking is already $($operation)d on the network interface"
        }
        elseif( $sourceNetworkInterfaces.Count -eq $alreadyInDesiredState )
        {
            Write-Warning -Message "Accelerated networking is already $($operation)d on all $($sourceNetworkInterfaces.Count) network interfaces"
        }
        else ## not all NICs in the desired state
        {
            if( $sourceNetworkInterfaces.Count - $alreadyInDesiredState -eq $changed )
            {
                Write-Warning -Message "Accelerated networking was already $($operation)d on $alreadyInDesiredState network interfaces and has been $($operation)d on the others"
            }
            else
            {
                Write-Warning -Message "Accelerated networking was already $($operation)d on $alreadyInDesiredState network interfaces but could only be $($operation)d on $changed of the $($sourceNetworkInterfaces.Count - $alreadyInDesiredState  - $changed) others"
            }
        }
    }
}

