#requires -version 3
$ErrorActionPreference = 'Stop'

<#
    .SYNOPSIS
    Migrate VM to new datastore

    .DESCRIPTION
    This script will migrate the ALL the VM (local) disks to a new datastore. If all the disks are already on the target datastore the script will exit.

    .PARAMETER HypervisorPlatform
    The hypervisor platform, will be passed from ControlUp Console

    .PARAMETER strVCenter
    The name of the vcenter server that will be connected to to run the PowerCLI commands

    .PARAMETER strVMName
    The name of the virtual machine the action is to be performed on/for.

    .PARAMETER strDatastore
    The name of the datastore the VM has to be migrated to.

    .EXAMPLE
    Example is not relevant as this script will be called through ControlUp Console

    .NOTES
    VMware PowerCLI Core needs to be installed on the machine running the script.
    Loading VMWare PowerCLI will result in a 'Join our CEIP' message. In order to disable these in the future run the following commands on the target system:
    Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false (or $true, if that's your kind of thing)
    
 
#>

[string]$strHypervisorPlatform = $args[0]
[string]$strVCenter = $args[1]
[string]$strVMName = $args[2]
[string]$strDatastore = $args[3]

Function Feedback {
    Param (
        [Parameter(Mandatory = $true,
            Position = 0)]
        [string]$Message,
        [Parameter(Mandatory = $false,
            Position = 1)]
        $Exception,
        [switch]$Oops
    )

    # This function provides feedback in the console on errors or progress, and aborts if error has occured.
    If (!$Exception -and !$Oops) {
        # Write content of feedback string
        Write-Host $Message -ForegroundColor 'Green'
    }

    # If an error occured report it, and exit the script with ErrorLevel 1
    Else {
        # Write content of feedback string but to the error stream
        $Host.UI.WriteErrorLine($Message) 
        
        # Display error details
        If ($Exception) {
            $Host.UI.WriteErrorLine("Exception detail:`n$Exception")
        }
        
        # Exit errorlevel 1
        Exit 1
    }
}

Function Test-HypervisorPlatform ([string]$strHypervisorPlatform) {
    # This function checks if the hypervisor is supported by this script.
    If ($strHypervisorPlatform -ne 'VMWare') {
        Feedback -Message "Currently this script based action only supports VMWare, selected guest is not running on VMWare" -Oops
    }
}

function Load-VMWareModules {
    <# Imports VMware PowerCLI modules, with a -Prefix $Prefix is supplied (desirable to avoid conflict with Hyper-V cmdlets)
      NOTES:
      - The required modules to be loaded are passed as an array.
      - If the PowerCLI versions is below 6.5 some of the modules can't be imported (below version 6 it is Snapins only) using so Add-PSSnapin is used (which automatically loads all VMWare modules)
    #>

    param (    
        [parameter(Mandatory = $true,
            ValueFromPipeline = $false)]
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
                Write-Host 'The required VMWare PowerCLI components were not found as modules or snapins. Please make sure VMWare PowerCLI (version 6.5 or higher preferred) is installed and available for the user running the script.'
                Exit 1
            }
        }
    }
}

Function Connect-VCenterServer {
    Param (
        [Parameter(Mandatory = $true,
            Position = 0)]
        [string]$VCenterName
    )
    Try {
        # Connect to VCenter server
        Connect-VIServer -Server $VCenterName -WarningAction SilentlyContinue -Force
    }
    Catch {
        Feedback -Message "There was a problem connecting to VCenter server $VCenterName. Please correct the error and try again." -Exception $_
    }
}

Function Disconnect-VCenterServer {
    Param (
        [Parameter(Mandatory = $true,
            Position = 0)]
        $VCenter
    )
    # This function closes the connection with the VCenter server 'VCenter'
    try {
        # Disconnect from the VCenter server
        Disconnect-VIServer -Server $VCenter -Confirm:$false
    }
    catch {
        Feedback -Message "There was a problem disconnecting from VCenter server $($VCenter.name)" -Exception $_
    }
}

Function Get-VMWareVirtualMachine {
    param (
        [parameter(Mandatory = $true)]
        [string]$strVMName,
        [parameter(Mandatory = $true)]
        $objVCenter
    )
    # This function retrieves the VMware VM
    try {
        Get-VM -Name $strVMName -Server $objVCenter
    }
    catch {
        Feedback -Message ("There was a problem retrieving virtual machine $strVMName.") -Exception $_
    }
}
  

function Get-VirtualMachineHarddisks {
    param (
        [parameter(Mandatory = $true)]
        $objVM,
        [parameter(Mandatory = $true)]
        $objVCenter
    )
    try {
        Get-HardDisk -VM $objVM -Server $objVCenter
    }
    catch {
        Feedback -Message 'The virtual machine hard disks could not be retreived' -Oops
    }
}

function Get-VMwareDS {
    param (
        [parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            Position = 0)]
        [string]$DatastoreName,
        [parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            Position = 1)]
        $VCenter
    )

    # Get the Datastore
    try {
        Get-Datastore -Name $DatastoreName -Server $VCenter
    }
    catch {
        Feedback -Message "Datastore $DatastoreName could not be retreived." -Exception $_
    }
}

function Migrate-VMWareVMMachineStorage {
    param (
        [parameter(Mandatory = $true)]
        $objVM,
        [parameter(Mandatory = $true)]
        $objVCenter,
        [parameter(Mandatory = $true)]
        $objStorageResource
    )

    try {
        Move-VM -VM $objVM -Server $objVCenter -Confirm:$false -Datastore $objStorageResource
    }
    catch {
        Feedback -Message 'The VM storage could not be migrated to the requested Datastore.' -Exception $_
    }
}


# Check all the arguments have been passsed
if ($args.Count -ne 4) {
    Feedback -Message  "The script did not get the correct amount of arguments from the Console. This can occur if you are not connected to the VM's hypervisor.`nPlease connect to the hypervisor in the ControlUp Console and try again." -Oops
}

# Check that the host is a supported hypervisor
Test-HypervisorPlatform $strHypervisorPlatform

# Import the VMWare PowerCLI module
Load-VMwareModules -Components @('VimAutomation.Core')

# Connect to VCenter server for VMWare
$objVCenter = Connect-VCenterServer $strVCenter

# Get the virtual machine
$objVM = Get-VMWareVirtualMachine $strVMName $objVCenter

# Get the virtual machine's hard disks to check their current location
$objVMHDDs = Get-VirtualMachineHarddisks $objVM $objVCenter

# Create counter to store amount of HDDs already on the target cluster
[int]$intHDDsOnTargetDatastore = 0

# Check each HDD storage location
foreach ($hdd in $objVMHDDs) {
    if ($hdd.Filename.Split(']')[0].Replace('[', '') -eq $strDatastore) {
        Feedback -Message "Hard disk `'$($hdd.Name)`' will not be moved to datastore $strDatastore because it is already stored there."
        $intHDDsOnTargetDatastore ++   
    }
    
}

# If ALL the hard disks are already stored on the target Datastore, exit with error
if ($objVMHDDs.Count -eq $intHDDsOnTargetDatastore) {
    Feedback -Message "No hard disks were moved to target Datastore $strDatastore because they were all already stored there." -Oops
}


# Get the target datastore
$objDatastore = Get-VMwareDS $strDatastore $objVCenter

# Move the VM storage
Migrate-VMWareVMMachineStorage $objVM $objVCenter $objDatastore | Out-Null

# Output the current disk(s) location
foreach ($hdd in Get-VirtualMachineHarddisks $objVM $objVCenter) {
    Feedback "`'$($hdd.Name)`' file: $($hdd.Filename)"
}

# Disconnect from the VCenter server
Disconnect-VCenterServer $objVCenter
