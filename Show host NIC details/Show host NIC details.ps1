<#
    .SYNOPSIS
    Reports the details of VMWare host's physical network adapters

    .DESCRIPTION
    This script will retreive the network adapters of the hypervisor machine and output their details

    .PARAMETER HypervisorPlatform
    The hypervisor platform, will be passed from ControlUp Console

    .PARAMETER VCenter
    The name of the vcenter server that will be connected to to run the PowerCLI commands

    .PARAMETER Host
    The name of the host the action is to be performed on/for.

    .EXAMPLE
    Example is not relevant as this script will be called through ControlUp Console

    .NOTES
    VMware PowerCLI needs to be installed on the machine running the script.
 
#>

[string]$strHypervisorPlatform = $args[0]
[string]$strVCenter = $args[1]
[string]$strVMWareHost = $args[2]

Function Feedback {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$strFeedbackString,
        [switch]$Oops
    )

    # This function provides feedback in the console on errors or progress, and aborts if error has occured.
    If (!($error) -and !($Oops)) {
        # Write content of feedback string
        Write-Host $strFeedbackString -ForegroundColor 'Green'
    }
  
    # If an error occured report it, and exit the script with ErrorLevel 1
    Else {
        # Write content of feedback string but in red
        $Host.UI.WriteErrorLine($strFeedbackString) 
    
        # Display error details
        If ($error) {
            $Host.UI.WriteErrorLine("Exception detail: $($error[0].Exception.Message)")
        }

        # Exit errorlevel 1
        Exit 1
    }
}

Function Test-HypervisorPlatform ([string]$strHypervisorPlatform) {
    # This function checks if the hypervisor is supported by this script.
    If ($strHypervisorPlatform -ne 'VMWare') {
        Feedback -Oops "Currently this script based action only supports VMWare, selected guest $strVMName is not running on VMWare"
    }
}

function Import-VMWareVIMAutomationCoreModule {
    # Imports VMware PowerCLI module WITH -Prefix VMWare to avoid conflict with Hyper-V cmdlets
    # Minimum version number can be passed as string, default is '5.1.0.0'
    param (    
        [parameter(Mandatory = $false,
            ValueFromPipeline = $true)]
        [string]$strMinimumVersion = '5.1.0.0'
    )
    # Import VMWare PowerCLI Core module
    try {
        Import-Module VMware.VimAutomation.Core -Prefix VMWare -MinimumVersion $strMinimumVersion | Out-Null
    }
    catch {
        # Check module
        try {
            $module = Get-InstalledModule VMware.VimAutomation.Core
            # Check the installed version meets the requirement.
            if ($($module.version) -lt [System.Version]"$strMinimumVersion") {
                Feedback -Oops "The VMware.VimAutomation.Core module version on this system is $($module.version). This is too low, the minimum required version is $strMinimumVersion"
            }
        }
        catch {
            Feedback 'No VMware.VimAutomation.Core module was found. Perhaps VMWare PowerCLI is not installed on this system?'
        }
        Feedback "There was an error loading the VMware.VimAutomation.Core module."
    }
}

Function Connect-VCenterServer ([string]$strVCenter) {
    Try {
        # Connect to VCenter server
        Connect-VMWareVIServer $strVCenter -WarningAction SilentlyContinue -Force
    }
    Catch {
        Feedback "There was a problem connecting to VCenter server $strVCenter. Please correct the error and try again."
    }
}

function Get-VMWareHost {
    param (
        [parameter(Mandatory = $true)]
        [string]$strVMWareHost,
        [parameter(Mandatory = $true)]
        $objVCenter
    )
    try {
        Get-VMWareVMHost -Name $strVMWareHost -Server $objVCenter
    }
    catch {
        Feedback -Oops "There was a problem retrieving the VMWare host machine $strHostName."
    }
}

function Get-VMWareHostPhysicalNICs {
    param (
        [parameter(Mandatory = $true)]
        $objVMHost,
        [parameter(Mandatory = $true)]
        $objVCenter
    )
    try {
        Get-VMWareVMHostNetworkAdapter -VMHost $objVMHost -Physical
    }
    catch {
        Feedback -Oops "The VMWare host machine network adapter details could not be retreived."
    }
}

Function Disconnect-VCenterServer ($objVCenter) {
    # This function closes the connection with the VCenter server 'VCenter'
    try {
        # Disconnect from the VCenter server
        Disconnect-VMWareVIServer -Server $objVCenter -Confirm:$false
    }
    catch {
        Feedback "There was a problem disconnecting from VCenter server $objVCenter."
    }
}

# Check all the arguments have been passsed
if ($args.Count -ne 3) {
    Feedback -Oops  "The script did not get the correct amount of arguments from the Console. This can occur if you are not connected to the VM's hypervisor.`nPlease connect to the hypervisor in the ControlUp Console and try again."
}

# Check that the host is a supported hypervisor
Test-HypervisorPlatform $strHypervisorPlatform

# Import the VMWare PowerCLI module
Import-VMWareVIMAutomationCoreModule '5.1.0.0'

# Connect to VCenter server for VMWare
$objVCenter = Connect-VCenterServer $strVCenter

# Get the VMWAre Host
$objVMHost = Get-VMWareHost $strVMWareHost $objVCenter

# Create VMWare host network adapter object, physical adapters only
$objVMHostNICs = Get-VMWareHostPhysicalNICs $objVMHost $objVCenter

# Create array for desired details
[array]$arrNICProperties = @(
  'DeviceName',
  @{Name='SpeedMB';Expression={$_.ExtensionData.LinkSpeed.SpeedMB}},
  @{Name='Duplex';Expression={$_.ExtensionData.LinkSpeed.Duplex}},
  'MAC',
  'DhcpEnabled',
  @{Name='Driver';Expression={$_.ExtensionData.Driver}},
  @{Name='VmDirectPathGen2Supported';Expression={$_.ExtensionData.VmDirectPathGen2Supported}}
)

# Output the required details
$objVMHostNICS | Select-Object $arrNICProperties | Format-Table


# Disconnect from the VCenter server
Disconnect-VCenterServer $objVCenter

