#require -version 3.0

<#
.SYNOPSIS
    If VM doesn't have public IP, add, assign to NIC, allow through firewall, mstsc to it, wait for exit and disable firewall and remove public IP (if added)

.DESCRIPTION
    Using REST API calls

.PARAMETER azid
    The relative URI of the Azure VM
    
.PARAMETER AZtenantId
    Optional Azure tenant id. Specify when there is a need to access multiple tenants with different credentials.

.NOTES
    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2021-10-30
    Updated:        2022-01-18  Added code to re-auth in case mstsc run time exceeds auth duration. Change oauth to use v2 url
                    2022-02-16  Added code to check VM has finished provisioning and is running and that any existing public IP address is not empty
                    2022-02-22  Added check for empty/malformed AZid and moved subscription parsing higher up before REST calls are made
                    2022-03-03  Fix to destination prefix checking which marked as accessible when it wasn't.
                                Fix problem with not calculating rule priority correctly
                    2022-03-07  Added wait for public IP address to appear on VM's network interface
                    2022-09-26  Added -force to continue on possible non-fatal errors
                    2022-09-29  Added retries when getting external IP address
                    2022-09-30  Added setting of TLS12 & TLS13
#>

[CmdletBinding()]

Param
(
    [string]$AZid ,## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId ,
    [int]$rdpPort = 3389 ,
    [boolean]$force
    ## TODO do we have an option to not delete the IP address and rule or leave open for a given amount of time and then remove?
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
[string]$networkApiVersion = '2021-05-01'
[string]$baseURL = 'https://management.azure.com'
[string]$credentialType = 'Azure'
[int]$nearTimeoutSeconds = 300 ## some tidy up operations can take a while so factor this into the check if we are near auth token expiry
[int]$highestPriorityRuleAllowed = 100
[int]$rulePriorityGap = 5
[int]$newRulePriority = 2500
[int]$provisioningWaitTimeSeconds = 180

Write-Verbose -Message "AZid is $AZid"

#region AzureFunctions
function Get-AzSPStoredCredentials {
    <#
    .SYNOPSIS
        Retrieve the Azure Service Principal Stored Credentials.
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

function Get-AzAuthToken {
    <#
    .SYNOPSIS
        Retrieve the Azure Authentication Token for an authentication session.
    .EXAMPLE
        Get-AzAuthToken -SPCredentials <PSCredentialObject> -TenantID <string>
    .CONTEXT
        Azure
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-03-22
        Updated:        2020-05-08
                        Created a separate Azure Credentials function to support ARM architecture and REST API scripted actions
                        2021-01-18
                        Changed from Get-AzBearerToken to return entire token so we can deal with timeouts
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
    [string]$uri = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"

    [hashtable]$body = @{
        grant_type    = 'client_credentials'
        client_Id     = $SPCredentials.UserName
        client_Secret = $SPCredentials.GetNetworkCredential().Password
        scope         = "$baseURL/.default"
    }

    [hashtable]$invokeRestMethodParams = @{
        Uri             = $uri
        Body            = $body
        Method          = 'POST'
        ContentType     = 'application/x-www-form-urlencoded'
    }

    Invoke-RestMethod @invokeRestMethodParams
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
        [string]$property = 'value' ,
        [string]$contentType = 'application/json'
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
        $invokeRestMethodParams.Add( 'Body' , ( $body | ConvertTo-Json -Depth 20))
    }

    if( -not [String]::IsNullOrEmpty( $property ) )
    {
        Invoke-RestMethod @invokeRestMethodParams | Select-Object -ErrorAction SilentlyContinue -ExpandProperty $property
    }
    else
    {
        Invoke-RestMethod @invokeRestMethodParams ## don't pipe through select as will slow script down for large result sets if processed again after rreturn
    }
}

#endregion AzureFunctions

#region OtherFunctions

<#
.SYNOPSIS
    Test if an IPv4 address is within the given CIDR range

.PARAMETER cidr
    The CIDR to test

.PARAMETER address
    The IP address to test against the CIDR specified

.EXAMPLE
    Test-IPRangeFromCIDR -cidr "192.168.2.1/28" -address 192.168.2.10

    Test if the specified IP address is contained within the given CIDR range

.NOTES

    Modification History:

    2021/11/03  @guyrleech  Initial Release
#>

Function Test-IPRangeFromCIDR
{
    [cmdletbinding()]

    Param
    (
        [Parameter(Mandatory=$true,HelpMessage='IP address range as CIDR')]
        [string]$cidr ,
        [Parameter(Mandatory=$true,HelpMessage='IP address to check in range')]
        [ipaddress]$address
    )

    [ipaddress]$startAddress = [ipaddress]::Any
    [ipaddress]$endAddress   = [ipaddress]::Any

    if( Get-IPRangeFromCIDR -cidr $cidr -startAddress ([ref]$startAddress) -endAddress ([ref]$endAddress) )
    {
        [byte[]]$bytes = $address.GetAddressBytes()
        [uint64]$addressToCompare =  ( ( [uint64]$bytes[0] -shl 24) -bor ( [uint64]$bytes[1] -shl 16) -bor ( [uint64]$bytes[2] -shl 8) -bor  [uint64]$bytes[3])
        $bytes = $startAddress.GetAddressBytes()
        [uint64]$startAddressToCompare =  ( ( [uint64]$bytes[0] -shl 24) -bor ( [uint64]$bytes[1] -shl 16) -bor ( [uint64]$bytes[2] -shl 8) -bor  [uint64]$bytes[3])
        $bytes = $endAddress.GetAddressBytes()
        [uint64]$endAddressToCompare =  ( ( [uint64]$bytes[0] -shl 24) -bor ( [uint64]$bytes[1] -shl 16) -bor ( [uint64]$bytes[2] -shl 8) -bor  [uint64]$bytes[3])

        $addressToCompare -ge $startAddressToCompare -and $addressToCompare -le $endAddressToCompare ## return
    }
}

<#
.SYNOPSIS
    Take a CIDR (Classless Inter-Domain Routing) notation IP v4 range and returns the first and last IPv4 addresses in the range

.PARAMETER cidr
    The CIDR to convert

.PARAMETER startAddress
    Will be set to the start address of the range if the CIDR is valid
    
.PARAMETER endAddress
    Will be set to the end address of the range if the CIDR is valid

.EXAMPLE
    Get-IPRangeFromCIDR -cidr "192.168.2.1/28" -Verbose -startAddress ([ref]$start) -endAddress ([ref]$end)

    Get the starting and ending IPv4 addresses of the specified CIDR range

.NOTES
    Results compared with https://mxtoolbox.com/SubnetCalculator.aspx

    Modification History:

    2021/11/03  @guyrleech  Initial Release
#>

Function Get-IPRangeFromCIDR
{
    [cmdletbinding()]

    Param
    (
        [Parameter(Mandatory=$true,HelpMessage='IP address range as CIDR')]
        [string]$cidr ,
        [Parameter(Mandatory=$true,HelpMessage='IP address range start result')]
        [ref]$startAddress ,
        [Parameter(Mandatory=$true,HelpMessage='IP address range end result')]
        [ref]$endAddress
    )

    [string]$ipaddressPart , [int]$bitsPart = $cidr -split '/'

    if( $bitsPart -eq $null -or $bitsPart -le 0 -or $bitsPart -gt 32 )
    {
        Write-Error -Message "/$bitsPart is invalid"
        return $false
    }

    if( -Not ( $ipaddress = $ipaddressPart -as [ipaddress] ))
    {
        Write-Error -Message "IP address $ipaddressPart is invalid"
        return $false
    }

    [uint64]$mask = ([int64][System.Math]::Pow( 2 , (32 - $bitsPart) ) - 1)
    [byte[]]$bytes = $ipaddress.GetAddressBytes()
    [uint64]$octets =  ( ( [uint64]$bytes[0] -shl 24) -bor ( [uint64]$bytes[1] -shl 16) -bor ( [uint64]$bytes[2] -shl 8) -bor  [uint64]$bytes[3])
    [uint64]$start = $octets -band ($mask -bxor 0xffffffff)
    [uint64]$end = $octets -bor $mask

    $startAddress.Value = [ipaddress]$start
    $endAddress.Value   = [ipaddress]$end

    return $true
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

#endregion OtherFunctions

[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13

[datetime]$authTime = [datetime]::Now

If ($azSPCredentials = Get-AzSPStoredCredentials -system $credentialType -tenantId $AZtenantId )
{
    # Sign in to Azure with a Service Principal with Contributor Role at Subscription level and retrieve the bearer token
    Write-Verbose -Message "Authenticating to tenant $($azSPCredentials.tenantID) as $($azSPCredentials.spCreds.Username)"
    ## save the whole token as we may need it if mstsc lives longer than the expiry time
    if( -Not ( $azAuthToken = Get-AzAuthToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID ) )
    { 
        Throw "Failed to get Azure authentication token"
    }
    
    if( -Not ( $azBearerToken = $azAuthToken | Select-Object -ExpandProperty access_token ) )
    { 
        Throw "Failed to get Azure authentication token"
    }

    [datetime]$tokenExpiryTime = $authTime.AddSeconds( $azAuthToken.expires_in )

    Write-Verbose -Message "$(Get-Date -Format G) : auth token expires at $(Get-Date -Date $tokenExpiryTime -Format G)"

    [string]$vmName = ($AZid -split '/')[-1]
    
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

    if( [string]::IsNullOrEmpty( $vmName ) )
    {
        Throw "Azure id `"$AZid`" does not appear valid - failed to find VM name"
    }

    if( -Not ( $vm = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid/?api-version=$computeApiVersion" -property $null ) )
    {
        Throw "Failed to get VM for $azid"
    }

    if( ! [string]::IsNullOrEmpty( $vm.id ) )
    {
        ## GRL 2021-10-20 appears to be a CU bug/feature that transmogrifies the VM name to lowercase which produces a blob URL that doesn't work so we replace the Az ID with what is returned here
        $AZid = $vm.id
        $vmName = $vm.Name
    }

    ## get instance view so we can check it is powered up
    
    ## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/instance-view
    [string]$instanceViewURI = "$baseURL/$azid/instanceView`?api-version=$computeApiVersion"
    if( $null -eq ( [array]$virtualMachineStatuses = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $instanceViewURI -property 'statuses' ) ) `
        -or $virtualMachineStatuses.Count -eq 0 )
    {
        if( $force )
        {
            Write-Warning -Message "Failed to get VM instance view via $instanceViewURI : $_"
        }
        else
        {
            Throw "Failed to get VM instance view via $instanceViewURI : $_"
        }
    }
    elseif( ( $line = ( $virtualMachineStatuses.code -match 'ProvisioningState/' )) -and ( $status = ($line -split '/' , 2 )[-1] ) )
    {
        ## check not performing an operation already
        Write-Verbose -Message "Current VM provisioning state is $status"
        if( $status -ine 'succeeded' )
        {
            if( $force )
            {
                Write-Warning -Message "VM $vmName has not finished a provisioning operation, it is $status"
            }
            else
            {
                Throw "VM $vmName has not finished a provisioning operation, it is $status"
            }
        }
    }
    else
    {
        Write-Warning -Message "Failed to determine current provisioning state of vm $vmName"
    }

    if( ( $line = ( $virtualMachineStatuses.code -match 'PowerState/' )) -and ( $powerstate = ($line -split '/' , 2 )[-1] ))
    {
        if( $powerstate -ine 'running' )
        {
            Throw "VM $vmName is not running, it is in power state $powerstate"
        }
    }
    else
    {
        Write-Warning -Message "Failed to determine if vm $vmName is powered up"
    }

    ## get its networking so we can see if it already has a public IP
    if( -Not ( [array]$networkInterfaces = @( $vm.properties | Select-Object -ExpandProperty networkProfile | Select-Object -ExpandProperty networkInterfaces | Select-Object -ExpandProperty Id ) ))
    {
        Throw "VM $vmName has no network interfaces"
    }
    
    [hashtable]$applicationSecurityGroups = @{}
    [bool]$alreadyReachable = $false

    $publicIpAddresses = New-Object -TypeName System.Collections.Generic.List[object]

    ForEach( $networkInterface in $networkInterfaces )
    {
        if( ( $thisNetworkInterface = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$networkInterface/?api-version=$networkApiVersion" -property $null ) `
            -and ( $IPproperties = $thisNetworkInterface.properties.ipConfigurations | Select-Object -ExpandProperty properties ) )
        {
            $TCPClient = [System.Net.Sockets.TcpClient]::new()
            ## see if we can connect to its rdport on an internal interface already (e.g. VPN in place) in which case user can use CU console instead
            [ipaddress]$internalAddress = $IPproperties | Select-Object -ExpandProperty privateIPAddress -ErrorAction SilentlyContinue
            $alreadyReachable = ( $internalAddress -and $TCPClient.ConnectAsync( $internalAddress , $rdpPort ).Wait( 2500 ) )

            $TCPClient.Close()
            $TCPClient.Dispose()
            $TCPClient = $null

            if( $alreadyReachable )
            {
                Write-Output -InputObject "Can already access port $rdpport on $($IPproperties.privateIPAddress)"
                break
            }
            elseif( $thisPublicIpAddress = $thisNetworkInterface.properties.ipConfigurations|Select-Object -ExpandProperty properties|Select-Object -ExpandProperty publicIPAddress -ErrorAction SilentlyContinue)
            {
                ## need to record the network interface for which we have the public IP address
                Write-Verbose -Message "Found existing public IP address"
                $publicIpAddresses.Add( ( [PSCustomObject]@{
                    'PublicIPAddress' = $thisPublicIpAddress.Id
                    'NetworkInterface' = $thisNetworkInterface
                    } ))
                $publicIPAddressURI = "$baseURL$($thisPublicIpAddress.Id)`?api-version=$networkApiVersion" ## so we can get the IP address
                ## don't break so we can get all application security groups for all the VMs NICs as they may appear in network security group rules which we check later for port 3389 access
                $thisNetworkInterface.properties.ipConfigurations|Select-Object -ExpandProperty properties | Select-Object -ExpandProperty applicationSecurityGroups -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id | Select-Object -Unique | ForEach-Object `
                {
                    [string]$applicationSecurityGroupName = ($_ -split '/')[-1]
                    try
                    {
                        $applicationSecurityGroups.Add( $applicationSecurityGroupName , $true )
                    }
                    catch
                    {
                        ## already have it which doesn't matter
                    }
                }
            }
        }
    }

    if( $alreadyReachable )
    {
        Add-Type -AssemblyName PresentationFramework
        [void][Windows.MessageBox]::Show( "Can already reach port $rdport on $($vm.name) - use mstsc locally or CU console RDP feature" , 'Script Error' , 'Ok' ,'Information' )
        exit 0
    }

    Write-Verbose -Message "Got $($applicationSecurityGroups.Count) application security groups for this VM's network interfaces"

    $newPublicIpAddress = $null

    if( -Not $publicIpAddresses -or -Not $publicIpAddresses.Count )
    {
        ## get a new public IP address and assign it to the NIC
        ## https://docs.microsoft.com/en-us/rest/api/virtualnetwork/public-ip-addresses/create-or-update

        [hashtable]$body = @{
                "properties" = @{
                    "publicIPAllocationMethod" = "Dynamic"
                    "DeleteOption" = "Delete"
                    "publicIpAddressVersion" = "IPv4"
                }
                "location" = $vm.location
                "tags" = @{
                    'Created' = "Added by ControlUp Script Action by $env:USERNAME $(Get-Date -Format G)"
                    'Creator' = 'ControlUp Script Action'
                }
              }

        if( $vm.psobject.properties[ 'zones' ] -and $vm.zones.Count -gt 0 )
        {
            ## get an error for standard sku if method is dynamic (static means always gets the same IP but as we delete it after, so shouldn't matter)
            $body.properties.publicIPAllocationMethod = "Static"
            $body += @{ "sku" = @{
                        "name" = "Standard"
                        "tier" = "Regional"
                     }
                ##"Zones" = $vm.zones ## this causes errors
                }
        }

        [string]$publicIPAddressName = $thisNetworkInterface.name + "-cu-pip"

        Write-Verbose -Message "Creating public IP address $publicIPAddressName"

        [string]$publicIPAddressURI = "$baseURL/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.Network/publicIPAddresses/$publicIpAddressName`?api-version=$networkApiVersion"
        if( -Not ( $newPublicIpAddress = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $publicIPAddressURI -body $body -property $null -method PUT ) )
        {
            Throw "Error when trying to create public IP address $publicIPAddressName"
        }
        
        if( -Not ( Wait-ProvisioningComplete -BearerToken $azBearerToken -uri $publicIPAddressURI -waitForSeconds $provisioningWaitTimeSeconds ))
        {
            Write-Warning "Error when trying to wait for public IP address $publicIPAddressName to be ready"
        }
        else
        {
            Write-Output -InputObject "Created public IP address $publicIPAddressName"
        }

        ## assign to the network interface now that it's ready
        ## https://docs.microsoft.com/en-us/rest/api/virtualnetwork/network-interfaces/create-or-update

        ## add the new public address to an existing ip configuration on the VM's NIC
        [string]$networkInterfaceURI = "$baseURL/$networkInterface`?api-version=$networkApiVersion"

        $PublicIPAddressId = [pscustomobject]@{ 'Id' = $newPublicIpAddress.id   }
        Add-Member -InputObject $thisNetworkInterface.properties.ipConfigurations[0].properties -MemberType NoteProperty -Name 'publicipaddress' -value $PublicIPAddressId
        if( -Not ( $nicUpdate = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $networkInterfaceURI -body $thisNetworkInterface -property $null -method PUT ) )
        {
            Throw "Error when trying to assign public IP address to network interface"
        }
    }

    $publicIPAddress = $null

    ## need to wait until we have the public IP address now that it is assigned to a NIC, or was already there. The pip will appear on the Public IP Address object
    if( $provisioningState = Wait-ProvisioningComplete -bearerToken $azBearerToken -uri $publicIPAddressURI -waitForSeconds $provisioningWaitTimeSeconds )
    {
        $ipconfiguration = $null
        [datetime]$finishTime = [datetime]::Now.AddSeconds( 120 )

        Write-Verbose -Message "$(Get-Date -Format G): waiting until $(Get-Date -Date $finishTime -Format G) for public IP address to be ready"
        do
        {
            if( -Not ( $ipconfiguration = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$($provisioningState.ipConfiguration.id)`?api-version=$networkApiVersion" -property $null -method GET ) )
            {
                Throw "Failed to get status of public IP $publicIPAddressURI"
            }
            elseif( $ipconfiguration.Properties.provisioningState -eq 'Succeeded' )
            {
                break
            }
            elseif( $ipconfiguration.Properties.provisioningState -ne 'Updating' )
            {
                Throw "Problem creating public IP $publicIPAddressURI - status is $($ipconfiguration.Properties.provisioningState)"
            }
            else
            {
                Start-Sleep -Milliseconds 5000
            }
        } while( [datetime]::Now -lt $finishTime )

        if( $ipconfiguration -and $ipconfiguration.Properties.provisioningState -eq 'Succeeded' -and ($PublicIPAddress = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $publicIPAddressURI -method GET -property $null | Select-Object -ExpandProperty properties | Select-Object -ExpandProperty ipAddress ))
        {
            Write-Verbose -Message "Public IP address is $publicIPAddress"
        }
        else
        {
            Throw "Failed to get public IP address via $publicIPAddressURI"
        }
    }
    else
    {
        Throw "Error waiting for public IP address to be available"
    }

    ## Check network security groups for NIC to see if 3389 will be allowed
    
    [bool]$rdpPortReachable = $false
    [int]$highestDeniedRulePriority = [int]::MaxValue
    [hashtable]$prioritiesUsed = @{}
    $networkSecurityGroup = $null
    [string]$networkSecurityGroupURI = $null
    [string]$networkSecurityGroupId = $null

    [string]$ipURL = 'https://ipinfo.io/ip'
    [datetime]$retryEnd = [datetime]::Now.AddSeconds( 15 )
    $externalIPAddress = $null

    while( -Not $externalIPAddress )
    {
        try
        {
            $externalIPAddress = Invoke-WebRequest -URI $ipURL | Select-Object -ExpandProperty Content -ErrorAction SilentlyContinue
        }
        catch
        {
            Write-Verbose -Message "$(Get-Date -Format G): sleeping before retry to $ipURL until $(Get-Date -Date $retryEnd -Format G) : $_"
            Start-Sleep -Milliseconds 1666
        }
    }

    if( -Not $externalIPAddress )
    {
        Throw "Unable to get external IP address from $ipURL so unable to create a network security group rule ofr just this IP"
    }

    Write-Verbose -Message "External IP address is $externalIPAddress"

    if( $thisNetworkInterface )
    {
        if( $thisNetworkInterface.properties.PSObject.properties[ 'networkSecurityGroup' ] -and $thisNetworkInterface.properties.networkSecurityGroup )
        {
            ## https://docs.microsoft.com/en-us/rest/api/virtualnetwork/network-security-groups/get
            [string]$networkSecurityGroupName = Split-Path -Path $thisNetworkInterface.properties.networkSecurityGroup.Id -Leaf
            Write-Verbose -Message "Getting NSG $networkSecurityGroupName"
            $networkSecurityGroupId = $thisNetworkInterface.properties.networkSecurityGroup.Id
        }
        else ## no NSG on NIC so check on subnet
        {
            $thisNetworkInterface.properties.ipConfigurations | Select-Object -ExpandProperty properties | Select-Object -ExpandProperty subnet -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id -ErrorAction SilentlyContinue | ForEach-Object `
            {
                if( $subnet = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL$_`?api-version=$networkApiVersion" -property 'properties' )
                {
                    $networkSecurityGroupId = $subnet | Select-Object -ExpandProperty networkSecurityGroup -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id -ErrorAction SilentlyContinue
                }

            }
        }
        
        if( $networkSecurityGroupId )
        {
            $networkSecurityGroupURI = "$baseURL/$networkSecurityGroupId`?api-version=$networkApiVersion"
            Write-Verbose -Message "Analysing network security group $networkSecurityGroupURI"
            if( $networkSecurityGroup = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $networkSecurityGroupURI -property 'properties' )
            {
                ## if rule added to NSG or to NIC, it will be in securityRules, not defaultSecurityRules
                ## sort on priority, once made numeric, so that highest priority is processed last
                ForEach( $securityRule in ($networkSecurityGroup.securityRules | Select-Object -ExpandProperty Properties | Select-Object -Property *,@{n='Priority';e={$_.Priority -as [int]}} -ExcludeProperty Priority | Sort-Object -Property Priority -Descending ) )
                {
                    ## need to track rule priorities so we can insert a higher priority rule to allow RDP if we need to
                    $prioritiesUsed.Add( $securityRule.Priority , $securityRule.access )

                    if( $securityRule.provisioningState -eq 'Succeeded' -and $securityRule.direction -eq 'inbound' -and ( $securityRule.protocol -eq 'TCP' -or  $securityRule.protocol -eq '*' ))
                    {
                        [bool]$isRDPPort = $false
                        ## see if port range includes 3389 and it it does we need to see if allow or deny and if sourceAddressPrefix allows us
                        ForEach( $port in $securityRule.destinationPortRange )
                        {
                            if( $port -match '^(\d*)-(\d*)$' )
                            {
                                $isRDPPort = $rdpPort -ge ($matches[1] -as [int]) -and $rdpPort -le ($matches[2] -as [int])
                            }
                            elseif( $port -eq '*' )
                            {
                                $isRDPPort = $true
                            }
                            else
                            {
                                $isRDPPort = $port -as [int] -eq $rdpPort
                            }
                        }
                        if( $isRDPPort )
                        {
                            [bool]$ruleAppliesToUs = $false
                            ## TODO deal with service tags
                            ## check we are in sourceAddressPrefix
                            if( -Not [string]::IsNullOrEmpty( $securityRule.sourceAddressPrefix ))
                            {
                                if( $securityRule.sourceAddressPrefix.IndexOf( '/' ) -gt 0 )
                                {
                                    if( $ruleAppliesToUs = Test-IPRangeFromCIDR -cidr $securityRule.sourceAddressPrefix -address $externalIPAddress )
                                    {
                                        Write-Verbose -Message "Our IP $externalIPAddress is in source CIDR $($securityRule.sourceAddressPrefix)"
                                    }
                                }
                                elseif( $securityRule.sourceAddressPrefix -eq '*' )
                                {
                                    $ruleAppliesToUs = $true
                                }
                                elseif( ($securityRule.sourceAddressPrefix -as [int]) -ne $null ) ## single IP address
                                {
                                    if( $ruleAppliesToUs = ( $securityRule.sourceAddressPrefix -eq $externalIPAddress ))
                                    {
                                        Write-Verbose -Message "Our IP $externalIPAddress is in source address $($securityRule.sourceAddressPrefix)"
                                    }
                                }
                                ## else ## a service tag since not numeric - assuming CU console not running somewhere else in Azure
                            }
                            ## check destinationAddressPrefix includes VM
                            if( $securityRule.PSObject.Properties[ 'destinationAddressPrefix' ] -and -Not [string]::IsNullOrEmpty( $securityRule.destinationAddressPrefix ))
                            {
                                $thisNetworkInterface.properties.ipConfigurations | Select-Object -ExpandProperty properties | Select-Object -ExpandProperty privateIPAddress | ForEach-Object `
                                {
                                    $vmIPAddress = $_
                     
                                    if( $securityRule.destinationAddressPrefix -eq '*' )
                                    {
                                        ## does not mean that it does apply to use but it won't stop access if we are allowed
                                    }
                                    elseif( $securityRule.destinationAddressPrefix.IndexOf( '/' ) -gt 0 )
                                    {
                                        if( -Not ( Test-IPRangeFromCIDR -cidr $securityRule.destinationAddressPrefix -address $vmIPAddress ) )
                                        {
                                            $ruleAppliesToUs = $false
                                            Write-Verbose -Message "VM IP Address $vmIPAddress not in destination CIDR $($securityRule.destinationAddressPrefix)"
                                        }
                                    }
                                    elseif( $securityRule.destinationAddressPrefix -ne $vmIPAddress )
                                    {
                                        $ruleAppliesToUs = $false
                                        Write-Verbose -Message "VM IP Address $vmIPAddress not in destination"
                        
                                    }
                                }
                            }
                            elseif( $securityRule.PSObject.Properties[ 'destinationApplicationSecurityGroups' ] -and $securityRule.destinationApplicationSecurityGroups )
                            {
                                ForEach( $applicationSecurityGroup in $securityRule.destinationApplicationSecurityGroups )
                                {
                                    [string]$applicationSecurityGroupName = ($applicationSecurityGroup.Id -split '/')[-1]
                                    ## see if it is applied to any of the NICs in our VM. If not then this rule does not apply to us.
                                    if( $applicationSecurityGroups[ $applicationSecurityGroupName ] )
                                    {
                                        $ruleAppliesToUs = $true
                                        Write-Verbose -Message "NSG rule contains destination ASG $applicationSecurityGroupName assigned to one of VM's NICs"
                                    }
                                    else
                                    {
                                        Write-Verbose -Message "NSG rule contains destination ASG $applicationSecurityGroupName which is NOT assigned to one of VM's NICs"
                                    }
                                }
                            }
                            if( $ruleAppliesToUs )
                            {
                                ## check if allow or deny
                                if( $securityRule.access -eq 'Deny' )
                                {
                                    $rdpPortReachable = $false
                                    if( $securityRule.Priority -lt $highestDeniedRulePriority )
                                    {
                                        $highestDeniedRulePriority = $securityRule.Priority
                                    }
                                }
                                elseif( $securityRule.access -eq 'Allow' )
                                {
                                    $rdpPortReachable = $true
                                }
                                else
                                {
                                    Write-Warning -Message "Unexpected access rule $($securityRule.access)"
                                }
                            }
                            Write-Verbose -Message "`tSecurity rule '$($securityRule|Select-Object -ExpandProperty Description -ErrorAction SilentlyContinue)' = rdp port access allowed $rdpPortReachable"
                        }
                    }
                }
            }
            else
            {
                Write-Warning -Message "Failed to get network security group $($thisNetworkInterface.properties.networkSecurityGroup.Id)"
            }
        }
        else
        {
            Write-Verbose -Message "No network security group on network interface with public IP address or subnet"
            $rdpPortReachable = $true
        }
    }
    else
    {
        Write-Warning -Message "No network interface to check security on"
    }

    $newrules = $null

    if( -Not $rdpPortReachable )
    {
        Write-Verbose -Message "Highest denied rule priority is $highestDeniedRulePriority"

        if( $highestDeniedRulePriority -le 100 )
        {
            Write-Warning -Message "Highest denied rule priority is already 100 so cannot insert a higher priority one"
        }
        elseif( $networkSecurityGroup )
        {
            ## new rule priority must be higher (so lower number) than highest priority deny rule and must be unique
            if( $highestDeniedRulePriority -lt [int]::MaxValue )
            {
                $newRulePriority = $highestDeniedRulePriority - 1
            }

            while( $newRulePriority -ge $highestPriorityRuleAllowed -and $prioritiesUsed.ContainsKey( $newRulePriority ) )
            {
                $newRulePriority -= $rulePriorityGap ## leave a gap in case someone needs to put a rule between
            }

            Write-Verbose -Message "New rule priority is $newRulePriority"

            if( $newRulePriority -lt $highestPriorityRuleAllowed )
            {
                Throw "Unable to find a priority for the network rule as highest priority allowed is $highestPriorityRuleAllowed"
            }

            ## can only be one NSG per NIC so we will need to edit the NSG already on the NIC rather than creating a new NSG
            [string]$newRuleName = "Allow RDP port $rdpPort (ControlUp) from $($externalIPAddress)"
            $newrules = [pscustomobject]@{ location = $vm.location
                properties = @{ securityrules = ( $networkSecurityGroup.securityRules + 
                    [pscustomobject]@{ 
                        name = $newRuleName
                        properties = @{
                            description                = "Added by ControlUp script by $env:USERNAME $(Get-Date -Format G)"
                            protocol                   = 'TCP'
                            sourcePortRange            = '*'
                            destinationPortRange       = $rdpPort
                            sourceAddressPrefix        = $externalIPAddress
                            destinationAddressPrefix   = $thisNetworkInterface.properties.ipConfigurations | Select-Object -ExpandProperty properties | Select-Object -ExpandProperty privateIPAddress -First 1
                            access                     = 'Allow'
                            priority                   = $newRulePriority
                            direction                  = 'Inbound'
                        }
                    })
                }
            }
            if( -Not ( $newruleResult = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $networkSecurityGroupURI -body $newrules -property $null -method PUT ) )
            {
                Write-Warning "Error when trying to create new rule in network security group"
            }

            if( -Not ( Wait-ProvisioningComplete -bearerToken $azBearerToken -uri $networkSecurityGroupURI -waitForSeconds $provisioningWaitTimeSeconds ))
            {
                Write-Warning -Message "Timed out waiting for network security group update to finish"
            }
            else
            {
                Write-Output -InputObject "Added new rule `"$newRuleName`" at priority $newRulePriority to network security group `"$(Split-Path -Path $networkSecurityGroupId -Leaf)`" for $externalIPAddress"
            }
        }
        else
        {
            Write-Warning -Message "No network securty group to update!"
        }
    }
    else
    {
        Write-Output -InputObject "RDP port $rdpPort appears to be reachable"
    }

    ## wait until we can get to the port
    $retry = 0
    [bool]$connectedRDPPort = $false

    [ipaddress]$externalAddress = $publicIpAddress

    Write-Verbose -Message "$(Get-Date -Format G): checking port $rdpPort on $publicIpAddress"

    $timer = [Diagnostics.Stopwatch]::StartNew()
    do
    {
        $TCPClient = [System.Net.Sockets.TcpClient]::new()
        $connectedRDPPort = $TCPClient.ConnectAsync( $externalAddress , $rdpPort ).Wait( 2500 )
        ## must close socket before we can try opening again
        $TCPClient.Close()
        $TCPClient.Dispose()
        $TCPClient = $null
        if( $connectedRDPPort )
        {
            break
        }
        Start-Sleep -Milliseconds 500
    } while ( $timer.ElapsedMilliseconds -le 60000 )
    $timer.Stop()
    
    Write-Verbose -Message "$(Get-Date -Format G): finished checking port $rdpPort on $publicIpAddress"

    if( -Not $connectedRDPPort )
    {
        Write-Warning -Message "Unable to connect to RDP port $($publicIpAddresses):$rdpPort"
    }

    $process = Start-Process -FilePath 'mstsc.exe' -ArgumentList "/v:$($publicIPAddress):$rdpPort" -Wait -PassThru
    
    Write-Verbose -Message "$(Get-Date -Format G): mstsc process ($($process.Id)) has exited"

    ## if we are close to timeout of auth token, get a new one
    if( ( $tokenExpiryTime - [datetime]::Now ).TotalSeconds -le $nearTimeoutSeconds )
    {
        Write-Verbose -Message "Renewing auth token"
        
        if( $newAuthToken = Get-AzAuthToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID )
        {
            if( $newBearerToken = $newAuthToken | Select-Object -ExpandProperty access_token )
            {
                $azBearerToken = $newBearerToken
            }
            else
            { 
                Write-Warning -Message "$(Get-Date -Format G) : failed to retrieve authentication token, current one expires at $(Get-Date -Date $tokenExpiryTime -Format G)"
            }
        }
        else
        {
            Write-Warning -Message "$(Get-Date -Format G) : failed to renew authentication token, current one expires at $(Get-Date -Date $tokenExpiryTime -Format G)"
        }
    }

    if( $newrules )
    {
         $oldrules = [pscustomobject]@{ 
                location = $vm.location
                properties = @{ 
                    securityrules = ( $networkSecurityGroup.securityRules )
                }
            }
        if( -Not ( $oldruleResult = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $networkSecurityGroupURI -body $oldrules -property $null -method PUT ) )
        {
            Write-Warning "Error when trying to set previous rules in network security group"
        }
        else
        {
            Write-Output -InputObject "Removed new rule from network security group"
        }
    }

    if( $newPublicIpAddress )
    {
        ## delete the public IP address since we created it after unassigning from the VM
        if( ( $stateOfNetworkInterface = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $networkInterfaceURI -property $null ) )
        {
            [bool]$removed = $false
            ForEach( $ipconfig in $stateOfNetworkInterface.properties.ipConfigurations )
            {
                if( $ipconfig.properties.publicipaddress.id -eq $PublicIPAddressId.id )
                {
                    Write-Verbose -Message "Removing public IP address property"
                    $ipconfig.properties.psobject.properties.remove( 'publicIPAddress' )
                    $removed = $true
                }
            }
            if( $removed )
            {
                if( -Not ( $nicUpdate = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $networkInterfaceURI -body $stateOfNetworkInterface -property $null -method PUT ) )
                {
                    Write-Warning "Error when trying to unassign public IP address from network interface"
                }
                else
                {
                    ## need to wait until the public IP address disappears from the NIC
                    [string]$networkInterfaceName = Split-Path -Path ($networkInterfaceURI -replace '\?.*$') -Leaf
                    [datetime]$start = [datetime]::Now
                    ## this can be slow
                    if( ( $stateOfNetworkInterface = Wait-ProvisioningComplete -BearerToken $azBearerToken -uri $networkInterfaceURI -sleepMilliseconds 5000 -waitForSeconds $provisioningWaitTimeSeconds ) )
                    {
                        if( -Not ($stateOfNetworkInterface.ipConfigurations | Select-Object -ExpandProperty properties | Select-Object -ExpandProperty ipAddress -ErrorAction SilentlyContinue))
                        {
                            Write-Output -InputObject "Public IP address $publicIPAddress removed ok from network interface $networkInterfaceName"
                        }
                    }
                    else
                    {
                        Write-Warning -Message "Public IP Address removal from network interface $networkInterfaceName did not complete in $([int]([datetime]::Now - $start).TotalSeconds) seconds"
                    }
                }
            }
            else
            {
                Write-Warning -Message "Unable to find ip configuration with the newly created public IP address"
            }
        }
        ## https://docs.microsoft.com/en-us/rest/api/virtualnetwork/public-ip-addresses/delete
        ## no return data so can't test
        try
        {
            Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $publicIPAddressURI -property $null -method DELETE
            Write-Output -InputObject "Public IP address $publicIPAddressName deleted"
        }
        catch
        {
            Throw $_
        }
    }
}

Write-Verbose -Message "$(Get-Date -Format G): script finished"

