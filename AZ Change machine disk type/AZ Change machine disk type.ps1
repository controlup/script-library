#require -version 3.0

<#
.SYNOPSIS
    Change the type of disk in the Azure VM passed (eg premium to standard when powered off)

.DESCRIPTION
    Using REST API calls

.PARAMETER AZid
    The relative URI of the Azure VM
    
.PARAMETER AZtenantId
    The tenanit id containing the Azure VM
    
.PARAMETER newDiskSKU
    The SKU of the disk to change to

.NOTES
    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2021-11-09
    Updated:
#>

[CmdletBinding()]

Param
(
    [string]$AZid , ## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId ,
    [ValidateSet('Standard_LRS','Premium_LRS','StandardSSD_LRS','UltraSSD_LRS','Premium_ZRS','StandardSSD_ZRS')]
    [string]$newDiskSKU 
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

Write-Verbose -Message "AZid is $AZid in tenant $AZtenantId"

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

If ($azSPCredentials = Get-AzSPStoredCredentials -system $credentialType -tenantId $AZtenantId )
{
    # Sign in to Azure with a Service Principal with Contributor Role at Subscription level and retrieve the bearer token
    Write-Verbose -Message "Authenticating to tenant $($azSPCredentials.tenantID) as $($azSPCredentials.spCreds.Username)"
    if( -Not ( $azBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID ) )
    {
        Throw "Failed to get Azure bearer token"
    }

    [string]$vmName = ($AZid -split '/')[-1]
    
    [string]$subscriptionId = $null
    [string]$resourceGroupName = $null

    ## subscriptions/58ffa3cb-2f63-4f2e-a06d-369c1fcebbf5/resourceGroups/WVD/providers/Microsoft.Compute/virtualMachines/GLMW10WVD-0
    if( $AZid -match '\bsubscriptions/([^/]+)/resourceGroups/([^/]+)/providers/Microsoft\.' )
    {
        $subscriptionId = $Matches[1]
        $resourceGroupName = $Matches[2]
    }
    else
    {
        Throw "Failed to parse subscription id and resource group from $AZid"
    }

    ## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/instance-view
    if( -Not ( $virtualMachineStatus = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid/instanceView`?api-version=$computeApiVersion" -property $null ) )
    {
        Throw "Failed to get VM instance view for $azid"
    }

    if( $virtualMachineStatus.statuses.Where( { $_.code -match '^PowerState/(.*)$' } ) )
    {
        if( $matches[1] -ne 'deallocated' )
        {
            Throw "Machine must be deallocated - $vmName is $($Matches[1])"
        }
    }
    else
    {
        Write-Warning -Message "Failed to get powerstate of vm $vmName - $($virtualMachineStatus.statuses|Format-Table -AutoSize|Out-String)"
    }
    
    if( -Not ( $virtualMachine = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid`?api-version=$computeApiVersion" -property $null ) )
    {
        Throw "Failed to get VM for $azid"
    }

    ## get disks
    if( $managedDisk = $virtualMachine.properties.storageProfile.osDisk | Select-Object -expandproperty managedDisk )
    {
        Write-Verbose -Message "Disk for $vmName is type $($managedDisk|Select-Object -ExpandProperty storageAccountType -ErrorAction SilentlyContinue)"
        ## https://docs.microsoft.com/en-us/rest/api/compute/disks/get
        ## if we use later API we get error "No registered resource provider found for location 'westeurope' and API version '2021-07-01' for type 'disks'. The supported api-versions are ..."
        if( $disk = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$($managedDisk.id)`?api-version=2021-04-01" -property $null )
        {
            Write-Verbose -Message "Disk tier is $($disk.sku.tier), sku $($disk.sku.name)"

            if( $disk.sku.name -eq $newDiskSKU )
            {
                Throw "Disk is already sku $($disk.sku.name)"
            }

            ## https://docs.microsoft.com/en-us/rest/api/compute/disks/create-or-update
            [hashtable]$body = @{
                'location' = $virtualMachine.location
                'sku' = @{
                    'name' = $newDiskSKU
                }
                'properties' = @{
                    'creationdata' = @{
                        'createOption' = $disk.properties.creationData.createOption
                        'imageReference' = $disk.properties.creationData.imageReference
                    }
                }
            }
            if( -Not ( $updateddisk = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$($managedDisk.id)`?api-version=2021-04-01" -property $null -body $body -method PUT) )
            {
                Write-Error -Message "Failed to update disk from sku $($disk.sku.name) to $newDiskSKU"
            }
            else
            {
                Write-Output -InputObject "Updated disk from sku $($disk.sku.name) to $newDiskSKU ok"
            }
        }
        else
        {
            Write-Warning -Message "Failed to get disk $($managedDisk.id)"
        }
    }
    else ## blob storage ? can we do anything with it?
    {
    }
}

