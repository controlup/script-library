#requires -version 3
$ErrorActionPreference = 'Stop'

<#
    .SYNOPSIS
    Shutdown VM Guest OS

    .DESCRIPTION
    This script will gracefully shut down the VM Guest OS using VMWare Client Tools.

    .PARAMETER HypervisorPlatform
    The hypervisor platform, will be passed from ControlUp Console

    .PARAMETER strVCenter
    The name of the vcenter server that will be connected to to run the PowerCLI commands

    .PARAMETER strVMName
    The name of the virtual machine the action is to be performed on/for.

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
$intTimeOut = 180

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

Function Test-VMwareToolsStatus {
    param (
        [parameter(Mandatory = $true)]
        $objVM
    )
    # This function checks if VMware Tools are installed. If no version of tools is returned this means they are not installed.
    If ((Get-VMGuest -VM $objVM).ToolsVersion -eq '') {
        Feedback -Message 'There are no VMware Guest Tools installed in the guest OS. Gracefull shutdown of the OS can only be done if the Guest Tools are installed.' -Oops
    }
}

function Shutdown-VMWareVMGuestOS {
    param (
        [parameter(Mandatory = $true)]
        $objVM,
        [parameter(Mandatory = $true)]
        $objVCenter,
        [parameter(Mandatory = $true)]
        [int]$intTimeOut
    )
    try {
        Stop-VMGuest -VM $objVM -Server $objVCenter -Confirm:$false
        
        # Wait until machine is down, or fail after timeout
        # Check if machine is down every 5 seconds
        for ($i = 5; $i -le $intTimeOut; $i += 5) {

            # If machine is Off, return
            [string]$strPowerState = (Get-VMWareVirtualMachine $strVMName $objVCenter | Select-Object PowerState).PowerState
            if ($strPowerState -eq 'PoweredOff') { return }

            # Wait 5 seconds before checking again
            Start-Sleep -Seconds 5
        } 

        # Power off took longer than timeout, exit with error
        Feedback -Message "Shutting down the target machine took more than the timeout $intTimeOut seconds, script will exit." -Oops
    }
    catch {
        Feedback -Message 'The guest OS could not be shut down.' -Exception $_
    }
}

# Check all the arguments have been passsed
if ($args.Count -ne 3) {
    Feedback -Message "The script did not get enough arguments from the Console. This can occur if you are not connected to the VM's hypervisor.`nPlease connect to the hypervisor in the ControlUp Console and try again." -Oops
}

# Check that the host is a supported hypervisor
Test-HypervisorPlatform $strHypervisorPlatform

# Import the VMWare PowerCLI module
Load-VMwareModules -Components @('VimAutomation.Core')

# Connect to VCenter server for VMWare
$objVCenter = Connect-VCenterServer $strVCenter

# Get the VM
$objVM = Get-VMWareVirtualMachine $strVMName $objVCenter

# Check if VMWare Tools are installed
Test-VMwareToolsStatus $objVM

# Shutdown VM Guest
Shutdown-VMWareVMGuestOS $objVM $objVCenter $intTimeOut | Out-Null

# Report the PowerState
Get-VMWareVirtualMachine $strVMName $objVCenter | Select-Object PowerState | Format-List

# Disconnect from the VCenter server
Disconnect-VCenterServer $objVCenter
