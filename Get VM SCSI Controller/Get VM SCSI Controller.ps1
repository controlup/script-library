$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Retreives the SCSI Controller used by the VM

    .DESCRIPTION
    Gets the VM through Vcenter and then enumerates the controller Name, BusSharingMode and UnitNumber

    .EXAMPLE
    For checking the storage configuration of a VM
    
    .NOTES
    Credit to fpacheco
    Context: Computer
    Modification history: 29062019 - Ton de Vreede - Updated fucntions, added some error handling, changes details for output. Also removed forcing TLS 1.2 setting for this script as this should be default on any machine for security.
    
    .COMPONENT
    Requires VMware PowerCLI on the Console:
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli
#>

$strHypervisorPlatform = $args[0]
$strVMName = $args[1]
$strVCenter = $args[2]

If ($args.count -ne 3) {
    Write-Host "ControlUp must be connected to the VM's hypervisor for this command to work. Please connect to the appropriate hypervisor and try again."
    Exit 1
}

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

Function Test-HypervisorPlatform {
    Param (
        [Parameter(Mandatory = $true,
            Position = 0)]
        [string]$strHypervisorPlatform
    )
    # This function checks if the hypervisor is supported by this script.
    If ($strHypervisorPlatform -ne 'VMware') {
        Feedback -Message "Currently this script based action only supports VMware, selected host $strVMwareHost is not running on VMware." -Oops
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

# Check that the host is a supported hypervisor
Test-HypervisorPlatform $strHypervisorPlatform

# Import the VMWare PowerCLI module
Load-VMwareModules -Components @('VimAutomation.Core')

# Connect to VCenter server for VMWare
$objVCenter = Connect-VCenterServer $strVCenter

# Get the VM
$objVM = Get-VMWareVirtualMachine $strVMName $objVCenter

# Output the VM SCSI controller
try {
    Get-ScsiController -VM $objVM
}
catch {
    Feedback -Message 'The SCSI Controller details could not be retreived.' -Exception $_
}

# Disconnect from the VCenter server
Disconnect-VCenterServer $objVCenter

