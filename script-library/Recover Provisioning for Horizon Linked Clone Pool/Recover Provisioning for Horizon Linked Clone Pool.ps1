$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Acts when provisioning is disabled for a Linked Clone Desktop Pool.

    .DESCRIPTION
    This script acts when provisioning gets disabled for linked clones desktop pools because the overcommit ratio is set too low. It will calculate the correct ratio and set it to that.
    After changing the ratio it will enable provisioning and when set to true it can also force a rebalance of the datastores.

    .NOTES
    This script requires VMWare PowerCLI to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers' Or by using the 'Install VMware PowerCLI' script.
    Credentials can be set using the 'Prepare machine for Horizon View scripts' script.

    please be aware of this when installing powercli from the powershell gallery:
    - https://devblogs.microsoft.com/powershell/powershell-gallery-tls-support/

    Modification history:   21/04/2020 - Wouter Kursten - First version

    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli


    .COMPONENT
    VMWare PowerCLI

#>

# Name of the Horizon View published Application. Passed from the ControlUp Console.
[string]$HVDesktopPool = $args[0]
# Name of the Horizon View connection server. Passed from the ControlUp Console.
[string]$HVConnectionServerFQDN = $args[1]
# Rebalance required or not
[string]$rebalance=$args[2]

Function Out-CUConsole {
    <# This function provides feedback in the console on errors or progress, and aborts if error has occured.
      If only Message is passed this message is displayed
      If Warning is specified the message is displayed in the warning stream (Message must be included)
      If Stop is specified the stop message is displayed in the warning stream and an exception with the Stop message is thrown (Message must be included)
      If an Exception is passed a warning is displayed and the exception is thrown
      If an Exception AND Message is passed the Message message is displayed in the warning stream and the exception is thrown
    #>

    Param (
        [Parameter(Mandatory = $false)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [switch]$Warning,
        [Parameter(Mandatory = $false)]
        [switch]$Stop,
        [Parameter(Mandatory = $false)]
        $Exception
    )

    # Throw error, include $Exception details if they exist
    if ($Exception) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        If ($Message) {
            Write-Warning -Message "$Message`n$($Exception.CategoryInfo.Category)`nPlease see the Error tab for the exception details."
            Write-Error "$Message`n$($Exception.Exception.Message)`n$($Exception.CategoryInfo)`n$($Exception.Exception.ErrorRecord)" -ErrorAction Stop
        }
        Else {
            Write-Warning "There was an unexpected error: $($Exception.CategoryInfo.Category)`nPlease see the Error tab for details."
            Throw $Exception
        }
    }
    elseif ($Stop) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        Write-Warning -Message "There was an error.`n$Message"
        Throw $Message
    }
    elseif ($Warning) {
        # Write the warning to Warning stream, thats it. It's a warning.
        Write-Warning -Message $Message
    }
    else {
        # Not an exception or a warning, output the message
        Write-Output -InputObject $Message
    }
}

Function Test-ArgsCount {
    <# This function checks that the correct amount of arguments have been passed to the script. As the arguments are passed from the Console or Monitor, the reason this could be that not all the infrastructure was connected to or there is a problem retreiving the information.
    This will cause a script to fail, and in worst case scenarios the script running but using the wrong arguments.
    The possible reason for the issue is passed as the $Reason.
    Example: Test-ArgsCount -ArgsCount 3 -Reason 'The Console may not be connected to the Horizon View environment, please check this.'
    Success: no ouput
    Failure: "The script did not get enough arguments from the Console. The Console may not be connected to the Horizon View environment, please check this.", and the script will exit with error code 1
    Test-ArgsCount -ArgsCount $args -Reason 'Please check you are connectect to the XXXXX environment in the Console'
    #>
    Param (
        [Parameter(Mandatory = $true)]
        [int]$ArgsCount,
        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    # Check all the arguments have been passed
    if ($args.Count -ne $ArgsCount) {
        Out-CUConsole -Message "The script did not get enough arguments from the Console. $Reason" -Stop
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
                Out-CUConsole -Message 'The required VMWare modules were not found as modules or snapins. Please check the .NOTES and .COMPONENTS sections in the Comments of this script for details.' -Stop
            }
        }
    }
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
        Out-CUConsole -Message 'There was a problem connecting to the Horizon View Connection server.' -Exception $_
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
        Out-CUConsole -Message 'There was a problem disconnecting from the Horizon View Connection server.' -Exception $_
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
        Out-CUConsole -Message 'There was a problem disconnecting from the Horizon View Connection server.' -Exception $_
    }
}

function Get-HVDesktopPool {
    param (
        [parameter(Mandatory = $true,
        HelpMessage = "Displayname of the Desktop Pool.")]
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
        # Filter on the correct displayname
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='desktopSummaryData.displayName'; 'value' = "$HVPoolname"}
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        if (!$queryResults){
            Out-CUConsole -message "Can't find $HVPoolName, exiting." -Exception "$HVPoolname not found" 
            exit
        }
        else {
            return $queryResults
        }
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the Horizon View Desktop Pool.' -Exception $_
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
    # Retreive the details of the desktop pool
    try {
        $HVConnectionServer.ExtensionData.Desktop.Desktop_Get($HVPoolID)
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the desktop pool details.' -Exception $_
    }
}

function Set-HVPoolProvisioningState {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "ID of the Desktop Pool.")]
        [VMware.Hv.DesktopId]$HVPoolID,
        [parameter(Mandatory = $true,
            HelpMessage = "Provisioningstateenabled: true or false")]
        [bool]$Provisioningstateenabled,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    try {
        # First define the Service we need
        [VMware.Hv.DesktopService]$desktopservice=new-object vmware.hv.DesktopService
        # Fill the helper for this service with the application information
        $desktophelper=$desktopservice.read($HVConnectionServer.extensionData, $HVPoolID)
        # Change the state of the application in the helper
        $desktophelper.getAutomatedDesktopDataHelper().getVirtualCenterProvisioningSettingshelper().setEnableProvisioning($Provisioningstateenabled)
        # Apply the helper to the actual object
        $desktopservice.update($HVConnectionServer.extensionData, $desktophelper)
    }
    catch {
        Out-CUConsole -Message 'There was a problem changing the provisioning state.' -Exception $_
    }
}

function Get-HVDesktopMachines {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "ID of the Horizon View Desktop Pool.")]
        [VMware.Hv.DesktopId]$HVPoolID,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'MachineSummaryView'
        # Filter for only the machines within the provided desktop pool
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='base.desktop'; 'value' = $HVPoolID}
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        return $queryResults
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the Horizon View machines.' -Exception $_
    }
}

# Test arguments
Test-ArgsCount -ArgsCount 3 -Reason 'The Console or Monitor may not be connected to the Horizon View environment, please check this.'

# Set the credentials location
[string]$strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"

# Import the VMware PowerCLI modules
Load-VMwareModules -Components @('VimAutomation.HorizonView')

# Get the stored credentials for running the script
[PSCredential]$CredsHorizon = Get-CUStoredCredential -System 'HorizonView'

# Connect to the Horizon View Connection Server

[VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $HVConnectionServerFQDN -Credential $CredsHorizon

# Retreive the Desktop Pool details

$HVPool=Get-HVDesktopPool -HVPoolName $HVDesktopPool -HVConnectionServer $objHVConnectionServer

# But we only need the ID
$HVPoolID=($HVPool).id
$hvpoolspec = Get-HVPoolSpec -HVConnectionServer $objHVConnectionServer -HVPoolID $HVPoolID
$diskspaceerror=@()
# yes there is an actual space behind the errors, add this or things will fail
$diskspaceerror +="Datastores unable to accommodate the new virtual machine because of one or more errors. Space reserved for expansion of VMs. "
$provisioningerror=$hvpoolspec.AutomatedDesktopData.ProvisioningStatusData.LastProvisioningError

# We need to check if there's an error and if it applies to having not enough disk space.
if (!$provisioningerror){
    Out-CUConsole -message "Exiting, no provisioning error found"
    exit
}
elseif ($diskspaceerror -notcontains $provisioningerror){
    Out-CUConsole -message "Exiting, provisioning error not caused by disk space reservation issues."
    exit
}

# Create an array of the available datastores
$pooldatastores=$hvpoolspec.AutomatedDesktopData.VirtualCenterProvisioningSettings.VirtualCenterStorageSettings.Datastores

# Get the configured datastores
$datastorelistspec=new-object VMware.Hv.DatastoreSpec
$datastorelistspec.DesktopId=$HVPoolID
$datastorelist=$objHVConnectionServer.ExtensionData.Datastore.Datastore_ListDatastoresByDesktopOrFarm($datastorelistspec)
$datastores=($datastorelist | where-object {($pooldatastores).datastore.id -contains $_.id.id})

# The amount of required machines in the desktop pool
[int]$requiredmachinecount=$hvpoolspec.AutomatedDesktopData.VmNamingSettings.PatternNamingSettings.MaxNumberOfMachines

# Golden Image vm
$parentvmid=$hvpoolspec.AutomatedDesktopData.VirtualCenterProvisioningSettings.VirtualCenterProvisioningData.ParentVm

# Golden image snapshot
$snapshotid=$hvpoolspec.AutomatedDesktopData.VirtualCenterProvisioningSettings.VirtualCenterProvisioningData.Snapshot
$baseimagesnapshot=$objHVConnectionServer.extensionData.BaseImageSnapshot.BaseImageSnapshot_List($parentvmid) | where-object {$_.id.id -eq $snapshotid.id}

# Disk Size for a single VM
$singlevmdisksize=($baseimagesnapshot.diskSizeInBytes)/1024/1024
$singlevmmemsize=$baseimagesnapshot.MemoryMB

# Total capacity of the configured Datastores
$datastorestotalsize=[Linq.Enumerable]::Sum([int[]]@($datastores).datastoredata.capacityMB)

# Calculate what's required per vm and per overcommit setting
[int]$diskrequiredpervm=$singlevmdisksize+$singlevmmemsize
[int]$diskrequiredtotal=$diskrequiredpervm*($requiredmachinecount+1)
[int]$conservativeallowed=$datastorestotalsize*4
[int]$moderateallowed=$datastorestotalsize*7
[int]$aggressiveallowed=$datastorestotalsize*15

$changerequired=0
$newovercommitsetting=$null

if (($diskrequiredtotal -ge $datastorestotalsize) -AND ($diskrequiredtotal -lt $conservativeallowed)) {
    $newovercommitsetting = "CONSERVATIVE"
}
elseif (($diskrequiredtotal -ge $conservativeallowed) -AND ($diskrequiredtotal -lt $moderateallowed)) {
    $newovercommitsetting = "MODERATE"
}
elseif (($diskrequiredtotal -ge $moderateallowed) -AND ($diskrequiredtotal -lt $aggressiveallowed)) {
    $newovercommitsetting = "AGGRESSIVE"
}
elseif ($diskrequiredtotal -ge $datastorestotalsize){
    out-cuconsole -Message "Not enough room to fit all the machines."
}

foreach ($pooldatastore in $pooldatastores){
    $currentovercommitsetting=$pooldatastore.StorageOvercommit

    if  ($currentovercommitsetting -eq $newovercommitsetting){
        out-cuconsole -message "Overcommitsetting is already at the desired level, not changimg this datastore"
    }
    else {
        $pooldatastore.StorageOvercommit = $newovercommitsetting
        $changerequired=1
    }
}
if ($changerequired -eq 1){
    [VMware.Hv.DesktopService]$desktopservice=new-object vmware.hv.DesktopService
    # Fill the helper for this service with the pool information
    $desktophelper=$desktopservice.read($objHVConnectionServer.extensionData, $HVPoolID)
    # Change the state of the pool in the helper
    $desktophelper.getAutomatedDesktopDataHelper().getVirtualCenterProvisioningSettingsHelper().getVirtualCenterStorageSettingsHelper().setdatastores($pooldatastores)
    $desktopservice.update($objHVConnectionServer.extensionData, $desktophelper)
}

# Enable provisioning
Set-HVPoolProvisioningState -HVConnectionServer $objHVConnectionServer -HVPoolID $hvpoolid -Provisioningstateenabled $true

# Rebalance if required
if ($rebalance -eq "True"){
    # Get a list of all the machines in the pool
    $machines = Get-HVDesktopMachines -HVConnectionServer $objHVConnectionServer -HVPoolID $HVPoolID

    # Create a spec for the rebalancing of the datastores
    $rebalancespec = new-object VMware.Hv.DesktopRebalanceSpec
    $rebalancespec.logoffsetting = "WAIT_FOR_LOGOFF"
    $rebalancespec.StopOnFirstError = $true
    $rebalancespec.Machines = $machines.ID

    # Rebalance the datastores
    $objHVConnectionServer.extensiondata.Desktop.Desktop_Rebalance($hvpoolid, $rebalancespec)
}

# Disconnect from the connection server
Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer

