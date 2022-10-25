#require -version 3.0

<#
.SYNOPSIS
    Get and display network info for specified VM

.DESCRIPTION
    Using REST API calls

.PARAMETER AZid
    The relative URI of the Azure VM

.PARAMETER AZtenantId
    The tenanit id containing the Azure VM
    
.NOTES
    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2021-10-30
    Updated:        2022-01-17  Guy Leech    Fix for tenant id handling
                    2022-01-17  Guy Leech    Fix for errors from missing public IP address properties when VM deallocated
#>

[CmdletBinding()]

Param
(
    [string]$AZid ,## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId
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
[int]$rdpPort = 3389

Write-Verbose -Message "AZid is $AZid in tenant $AZtenantId"

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

function Get-AzBearerToken {
    <#
    .SYNOPSIS
        Retrieve the Azure Bearer Token for an authentication session.
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

    Invoke-RestMethod @invokeRestMethodParams | Select-Object -ExpandProperty access_token -ErrorAction SilentlyContinue
}

function Invoke-AzureRestMethod {

    [CmdletBinding()]
    Param(
        [Parameter( Mandatory=$true, HelpMessage='A valid Azure bearer token' )]
        [ValidateNotNullOrEmpty()]
        [string]$BearerToken ,
        [string]$uri ,
        [ValidateSet('GET','POST','PUT','DELETE')] ## add others as necessary
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

If ($azSPCredentials = Get-AzSPStoredCredentials -system $credentialType -tenantId $AZtenantId )
{
    # Sign in to Azure with a Service Principal with Contributor Role at Subscription level and retrieve the bearer token
    Write-Verbose -Message "Authenticating to tenant $($azSPCredentials.tenantID) as $($azSPCredentials.spCreds.Username)"
    if( -Not ( $azBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID ) )
    { 
        Throw "Failed to get Azure bearer token"
    }

    [string]$vmName = ($AZid -split '/')[-1]

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
    
    ## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/instance-view
    if( -Not ( $virtualMachineStatus = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid/instanceView`?api-version=$computeApiVersion" -property 'statuses' ) )
    {
        Write-Warning "Failed to get VM instance view for $azid"
    }
        
    ## get its networking so we can see if it already has a public IP
    if( -Not ( [array]$networkInterfaces = @( $vm.properties | Select-Object -ExpandProperty networkProfile | Select-Object -ExpandProperty networkInterfaces | Select-Object -ExpandProperty Id ) ))
    {
        Throw "VM $vmName has no network interfaces"
    }
    
    Write-Verbose -Message "Got $($networkInterfaces.Count) network interfaces"

    [array]$networkInfo = @( ForEach( $networkInterface in $networkInterfaces )
    {
        $result = [pscustomobject]@{ 'Interface' = Split-Path -Path $networkInterface -Leaf }

        if( ( $thisNetworkInterface = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$networkInterface/?api-version=$networkApiVersion" -property 'properties' ) )
        {
            if( $IPproperties = $thisNetworkInterface.ipConfigurations | Select-Object -ExpandProperty properties )
            {
                if( $IPproperties.provisioningState -ne 'Succeeded' )
                {
                    Write-Warning -Message "Provisioning state of $networkInterface is $($IPproperties.provisioningState)"
                }
                
                ## don't break so we can get all application security groups for all the VMs NICs as they may appear in network security group rules which we check later for port 3389 access
                $IPproperties | Select-Object -ExpandProperty applicationSecurityGroups -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id | Select-Object -Unique | ForEach-Object `
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
                Add-Member -InputObject $result -NotePropertyMembers @{
                    'Nic Type' = $thisNetworkInterface.nicType
					'Accelerated Networking' = $thisNetworkInterface.enableAcceleratedNetworking
                    'Subnet' = (( $IPproperties |Select-Object -ExpandProperty subnet -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id -ErrorAction SilentlyContinue) -split '/') | Select-Object -Last 1
                    'Private IP' = $IPproperties | Select-Object -ExpandProperty privateIPAddress
                    'Private IP Allocation' = $IPproperties | Select-Object -ExpandProperty privateIPAllocationMethod
                    'Private IP Type' = $IPproperties | Select-Object -ExpandProperty privateIPAddressVersion
                    'MAC Address' = $thisNetworkInterface.macAddress
                    'Network Security Group' = (( $thisNetworkInterface | Select-Object -ExpandProperty networkSecurityGroup -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id -ErrorAction SilentlyContinue) -split '/') | Select-Object -Last 1
                }
                
                if( $IPproperties.PSObject.properties[ 'publicIPAddress' ] )
                {
                    if( ( $thisPublicIpAddressDetail = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL$($IPproperties.publicIPAddress.Id)`?api-version=$networkApiVersion" -property $null ) )
                    {
                        Add-Member -InputObject $result -NotePropertyMembers @{
                            'Public IP' = $thisPublicIpAddressDetail.properties | Select-Object -ExpandProperty ipAddress -ErrorAction SilentlyContinue
                            'Public IP Allocation' = $thisPublicIpAddressDetail.properties.publicIPAllocationMethod
                            'Public IP Type' = $thisPublicIpAddressDetail.properties.publicIPAddressVersion
                            'Public IP Location' = $thisPublicIpAddressDetail.location
                            'Public IP SKU' = $thisPublicIpAddressDetail.sku.name
                            'Public IP Tier' = $thisPublicIpAddressDetail.sku.tier
                            'Public IP Timeout' = "$($thisPublicIpAddressDetail.properties.idleTimeoutInMinutes) minutes"
                            'Public IP Tags' = $thisPublicIpAddressDetail | Select-Object -ExpandProperty Tags -ErrorAction SilentlyContinue
                        }
                    }
                    else
                    {
                        Write-Warning -Message "Failed to get details of public IP address $($thisPublicIpAddress.Id)"
                    }
                }
            }
        }
        else
        {
            Write-Warning -Message "Failed to get properties for network interface $networkInterface"
        }
        $result
    })

    Write-Output "Details for $($networkInfo.Count) network interfaces, $(($virtualMachineStatus | Where-Object code -match '^PowerState/' | Select-Object -ExpandProperty code) -replace '/' , ' is ') :"

    $networkInfo | Format-List
}
