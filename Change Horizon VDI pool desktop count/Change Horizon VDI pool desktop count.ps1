
<#
    .SYNOPSIS
    Changes the amount of Desktops in a Horizon Desktop Pool

    .DESCRIPTION
    This script changes the amount of Desktops in a Horizon Desktop Pool. 

    .PARAMETER HVDesktopPoolname
    Name of the Desktop Pool to update

    .PARAMETER Provisioningtype
    Use ON_DEMAND to provision all desktops up front (will ignore minNumberOfMachines and numberOfSpareMachines

    .PARAMETER maxNumberOfMachines
    Maximum number of desktops in the pool

    .PARAMETER minNumberOfMachines
    Minimum number of desktops in the pool

    .PARAMETER numberOfSpareMachines
    Minimum number of powered on desktops in the pool

    .PARAMETER HVConnectionServerFQDN
    FQDN for a connectionserver in the pod the pool belongs to.

    .NOTES
    This script requires VMWare PowerCLI to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers' Or by using the 'Install VMware PowerCLI' script.
    Credentials can be set using the 'Prepare machine for Horizon View scripts' script.

    Modification history:   12/12/2019 - Wouter Kursten - First version
                            26/03/2022 - Wouter Kursten - Added options for on demand provisioning

    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli


    .COMPONENT
    VMWare PowerCLI

#>

[CmdletBinding()]
Param
(
    [Parameter(
        Position = 0,
        Mandatory=$true,
        HelpMessage='Name of the Desktop Pool'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $HVDesktopPoolname,

    [Parameter(
        Position = 1,
        Mandatory=$true,
        HelpMessage='FQDN for the connection server'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $HVConnectionServerFQDN,

    [Parameter(
        Position = 2,
        Mandatory=$true,
        HelpMessage='Provisioning type'
    )]
    [ValidateSet("UP_FRONT","ON_DEMAND")]
    [string] $Provisioningtype,

    [Parameter(
        Position = 3,
        Mandatory=$true,
        HelpMessage='Change Provisioning type?'
    )]
    [ValidateSet("True","False")]
    [string] $ChangeProvisioningtype,

    [Parameter(
        Position = 4,
        Mandatory=$true,
        HelpMessage='Maximum number of machines in the desktop.'
    )]
    [ValidateNotNullOrEmpty()]
    [int] $maxNumberOfMachines,

    [Parameter(
        Position = 5,
        Mandatory=$false,
        ParameterSetName = 'ondemand',
        HelpMessage='The minimum number of machines to have provisioned if on demand provisioning is selected. Will be ignored if provisioningtype is set to UP_FRONT.'
    )]
    [ValidateNotNullOrEmpty()]
    [int] $minNumberOfMachines,

    [Parameter(
        Position = 6,
        Mandatory=$false,
        ParameterSetName = 'ondemand',
        HelpMessage='Number of spare powered on machines. Will be ignored if provisioningtype is set to UP_FRONT.'
    )]
    [ValidateNotNullOrEmpty()]
    [int] $numberOfSpareMachines
)

$ErrorActionPreference = 'Stop'

function Load-VMWareModules {
    <# Imports VMware modules
    NOTES:
    - The required modules to be loaded are passed as an array.
    - In versions of PowerCLI below 6.5 some of the modules can't be imported (below version 6 it is Snapins only) using so Add-PSSnapin is used (which automatically loads all VMWare modules)
    #>

    param (    
        [parameter(Mandatory = $true,
            HelpMessage = "The VMware module to be loaded. Can be single or multiple values (as array).")]
        [array]$Components
    )

    # Try Import-Module for each passed component, try Add-PSSnapin if this fails (only if -Prefix was not specified)
    # Import each module, if Import-Module fails try Add-PSSnapin
    foreach ($component in $Components) {
        try {
            $null = Import-Module -Name VMware.$component
        }
        catch {
            try {
                $null = Add-PSSnapin -Name VMware
            }
            catch {
                write-error "The required VMWare modules were not found as modules or snapins. Please check the .NOTES and .COMPONENTS sections in the Comments of this script for details."
                exit
            }
        }
    }
}

function Get-CUStoredCredential {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The system the credentials will be used for.")]
        [string]$System
    )
    # Get the stored credential object
    [string]$strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
    Import-Clixml $strCUCredFolder\$($env:USERNAME)_$($System)_Cred.xml
}

function Connect-HorizonConnectionServer {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The FQDN of the Horizon View Connection server. IP address may be used.")]
        [string]$HVConnectionServerFQDN,
        [parameter(Mandatory = $true,
            HelpMessage = "The PSCredential object used for authentication.")]
        [PSCredential]$Credential
    )
    # Try to connect to the Connection server
    try {
        Connect-HVServer -Server $HVConnectionServerFQDN -Credential $Credential
    }
    catch {
        write-error "There was a problem connecting to the Horizon View Connection server: $_."
        exit
    }
}

function Disconnect-HorizonConnectionServer {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    # Try to connect from the connection server
    try {
        Disconnect-HVServer -Server $HVConnectionServer -Confirm:$false
    }
    catch {
        write-error  "There was a problem disconnecting from the Horizon View Connection server: $_"
        exit
    }
}

function Get-HVDesktopPool {
    param (
        [parameter(Mandatory = $true,
        HelpMessage = "Name of the Desktop Pool.")]
        [string]$HVPoolName,
        [parameter(Mandatory = $true,
        HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    # Try to get the Desktop pools in this pod
    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'DesktopSummaryView'
        # Filter oud rds desktop pools since they don't contain machines
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='desktopSummaryData.displayName'; 'value' = "$HVPoolname"}
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        if (!$queryResults){
            write-error  "Can't find $HVPoolName, exiting"
            exit
        }
        elseif (($queryResults).desktopsummarydata.type -eq "MANUAL"){
            write-output  "This a manual Horizon View Desktop Pool, cannot change the amount of desktops"
            exit
        }
        elseif (($queryResults).desktopsummarydata.source -eq "VIRTUAL_CENTER"){
            write-output  "This a Full Clone Horizon View Desktop Pool, if the amount of desktops has been reduced the extra systems need to be removed manually"
            return $queryResults
        }
        else {
            return $queryResults
        }
    }
    catch {
        write-error  "There was a problem retreiving the Horizon View Desktop Pool: $_"
        exit
    }
}

function get-hvpoolspec{
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "ID of the Desktop Pool.")]
        [VMware.Hv.DesktopId]$HVPoolID,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    try {
        $HVConnectionServer.ExtensionData.Desktop.Desktop_Get($HVPoolID)
    }
    catch {
        write-error "There was a problem retreiving the desktop pool details: $_"
        exit
    }
}

function Set-HVPool {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "ID of the Desktop Pool.")]
        [VMware.Hv.DesktopId]$HVPoolID,
        [parameter(Mandatory = $true,
        HelpMessage = "Provisioning type UP_FRONT or ON_DEMAND")]
        [ValidateSet("UP_FRONT","ON_DEMAND")]
        [string] $Provisioningtype,
        [parameter(Mandatory = $true,
            HelpMessage = "Desired amount of desktops in the pool.")]
        [int]$maxNumberOfMachines,
        [parameter(Mandatory = $false,
        HelpMessage = "Desired amount of spare desktops in the pool.")]
        [int]$numberOfSpareMachines,
        [parameter(Mandatory = $false,
        HelpMessage = "Desired minimum amount of desktops in the pool.")]
        [int]$minNumberOfMachines,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    if($Provisioningtype -eq "UP_FRONT"){
        try {
            # First define the Service we need
            [VMware.Hv.DesktopService]$desktopservice=new-object vmware.hv.DesktopService
            # Fill the helper for this service with the application information
            $desktophelper=$desktopservice.read($HVConnectionServer.extensionData, $HVPoolID)
            # Change the state of the application in the helper
            $desktophelper.getAutomatedDesktopDataHelper().getVmNamingSettingsHelper().getPatternNamingSettingsHelper().setMaxNumberOfMachines($maxNumberOfMachines)
            $desktophelper.getAutomatedDesktopDataHelper().getVmNamingSettingsHelper().getPatternNamingSettingsHelper().setProvisioningTime("UP_FRONT")
            # Apply the helper to the actual object
            $desktopservice.update($HVConnectionServer.extensionData, $desktophelper)
        }
        catch {
            write-error "There was a problem changing the desktop count: $_"
            exit
        }
    }
    else{
        try {
            # First define the Service we need
            [VMware.Hv.DesktopService]$desktopservice=new-object vmware.hv.DesktopService
            # Fill the helper for this service with the application information
            $desktophelper=$desktopservice.read($HVConnectionServer.extensionData, $HVPoolID)
            # Change the state of the application in the helper
            $desktophelper.getAutomatedDesktopDataHelper().getVmNamingSettingsHelper().getPatternNamingSettingsHelper().setminNumberOfMachines($minNumberOfMachines)
            $desktophelper.getAutomatedDesktopDataHelper().getVmNamingSettingsHelper().getPatternNamingSettingsHelper().setMaxNumberOfMachines($maxNumberOfMachines)
            $desktophelper.getAutomatedDesktopDataHelper().getVmNamingSettingsHelper().getPatternNamingSettingsHelper().setnumberOfSpareMachines($numberOfSpareMachines)
            $desktophelper.getAutomatedDesktopDataHelper().getVmNamingSettingsHelper().getPatternNamingSettingsHelper().setProvisioningTime("ON_DEMAND")
            # Apply the helper to the actual object
            $desktopservice.update($HVConnectionServer.extensionData, $desktophelper)
        }
        catch {
            write-error "There was a problem changing the desktop count: $_"
            exit
        }
    }
}

if($Provisioningtype -eq "ON_DEMAND"){
    write-verbose "Checking if numberOfSpareMachines or minNumberOfMachines is missing"
    if(!$minNumberOfMachines -or !$numberOfSpareMachines){
        write-error "numberOfSpareMachines and minNumberOfMachines are required when using provisioningtype: $provisioningtype"
        exit
    }
}

# Set the credentials location
[string]$strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"

# Import the VMware PowerCLI modules
Load-VMwareModules -Components @('VimAutomation.HorizonView')

# Get the stored credentials for running the script
[PSCredential]$CredsHorizon = Get-CUStoredCredential -System 'HorizonView'

# Connect to the Horizon View Connection Server

[VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $HVConnectionServerFQDN -Credential $CredsHorizon

# Retreive the desktop pool
$HVPool=Get-HVDesktopPool -HVPoolName $HVDesktopPoolname -HVConnectionServer $objHVConnectionServer
write-verbose  "Retreived information about $HVDesktopPoolname"

# But we only need the ID
$HVPoolID=($HVPool).id

# Retreive the pool spec
$hvpoolspec=Get-HVPoolSpec -HVConnectionServer $objHVConnectionServer -HVPoolID $HVPoolID
$ProvisioningTime=($hvpoolspec).AutomatedDesktopData.VmNamingSettings.PatternNamingSettings.ProvisioningTime
write-verbose "Checking if provisioningtype matches the current setting and if I am allowed to change it."
if($ProvisioningTime -ne $provisioningtype -and $changeprovisioningtype -eq "False"){
    write-error "Provisioningtype of $provisioningtype does not match the current provisioningtype. Set changeprovisioningtype to True to change the provisioningtype"
    exit
}

# We cannot change manual pools so we give a warning about this and exit the script.
if ($hvpoolspec.Type -eq "MANUAL"){
    write-error "Could not execute, this a manual Horizon View Desktop Pool, cannot change the amount of desktops."
    exit
}

# When not all vm's are provisioned up front the max amount of machines can't be lower that the minimum amount or the number of spare machines.
if ($Provisioningtype -eq "ON_DEMAND"){
    if ($numberOfSpareMachines -ge $maxNumberOfMachines -or $minNumberOfMachines -ge $maxNumberOfMachines){
        write-error "Could not execute, the number of desktops cannot be smaller than the minimum amount of desktops or the number of spare desktops"
        exit
    }
}

# Change the desktop count in the pool

if($Provisioningtype -eq "UP_FRONT"){
    write-verbose "Provisioningtype is $Provisioningtype so ignoring minNumberOfMachines and numberOfSpareMachines if they have been added."
    write-verbose  "Trying to change $HVDesktopPoolname to $maxNumberOfMachines desktops."
    Set-HVPool -HVConnectionServer $objHVConnectionServer -HVPoolID $HVPoolID -maxNumberOfMachines $maxNumberOfMachines -Provisioningtype $Provisioningtype
    write-output  "Changed $HVDesktopPoolname to $maxNumberOfMachines desktops all provisioned up front."
}
else{
    write-verbose "Provisioningtype is $Provisioningtype so using minNumberOfMachines and numberOfSpareMachines."
    write-verbose  "Trying to change $HVDesktopPoolname to $maxNumberOfMachines desktops with a minimum of $minNumberOfMachines machines and $numberOfSpareMachines spares."
    Set-HVPool -HVConnectionServer $objHVConnectionServer -HVPoolID $HVPoolID -maxNumberOfMachines $maxNumberOfMachines -Provisioningtype $Provisioningtype -minNumberOfMachines $minNumberOfMachines -numberOfSpareMachines $numberOfSpareMachines
    write-output  "Changed $HVDesktopPoolname to $maxNumberOfMachines desktops with a minimum of $minNumberOfMachines machines and $numberOfSpareMachines spares."
}

# Disconnect from the connection server
Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
