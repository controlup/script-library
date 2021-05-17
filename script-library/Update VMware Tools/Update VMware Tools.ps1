#requires -Version 3.0
$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    This script will update the VMware Client tools on the target machine

    .DESCRIPTION
    The script uses PowerCLI Update-Tools with the -noreboot flag to update VMware Client tools (where applicable and possible)

    .PARAMETER strHypervisorPlatform
    The hypervisor platform, will be passed from ControlUp Console

    .PARAMETER strVCenter
    The name of the VCenter server that will be used to run the PowerCLI commands, will be passed from ControlUp Console

    .PARAMETER strVMName
    The name of the virtual machine that will be updated, will be passed from ControlUp Console

    .PARAMETER strIgnoreCertificateError
    If set to True invalid VCenter certificate errors will be ignored

    .NOTES
    The update command -NoReboot flag should prevent the target guest machine rebooting. However, VMware states the virtual machine might still reboot after updating VMware Tools,
    depending on the currently installed VMware Tools version, the VMware Tools version to which you want to upgrade, and the vCenter Center/ESX versions.

    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli
    
    .COMPONENT
    VMWare PowerCLI 6.5.0R1 or higher
#>

# Hypervisor platform
[string]$strHypervisorPlatform = $args[0]
# Name of Vcenter, parsed from Console input
[string]$strVCenter = $args[1].Replace('https://', '').Split('/')[0]
# Name of virtual machine
[string]$strVMName = $args[2]
# Ignore certificate errors for connecting the environment
[string]$strIgnoreCertificateError = $args[3]

Function Out-CUConsole {
    <# This function provides feedback in the console on errors or progress, and aborts if error has occurred.
      If only Message is passed this message is displayed
      If Warning is specified, the message is displayed in the warning stream (Message must be included)
      If Stop is specified, the stop message is displayed in the warning stream and an exception with the Stop message is thrown (Message must be included)
      If an Exception is passed a warning is displayed and the exception is thrown
      If an Exception AND Message is passed the Message is displayed in the warning stream and the exception is thrown
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
        Write-Warning -Message "There was a problem.`n$Message"
        Throw $Message
    }
    elseif ($Warning) {
        # Write the warning to Warning stream, that's it. It is a warning.
        Write-Warning -Message $Message
    }
    else {
        # Not an exception or a warning, output the message
        Write-Output -InputObject $Message
    }
}

Function Test-ArgsCount {
    <# This function checks that the correct number of arguments have been passed to the script. As the arguments are passed from the Console or Monitor, the reason this could be that not all the infrastructure was connected to or there is a problem retrieving the information.
      This will cause a script to fail, and in worst case scenarios the script running but using the wrong arguments.
      The possible reason for the issue is passed as the $Reason.
      Example: Test-ArgsCount -ArgsCount 3 -Reason 'The Console may not be connected to the Horizon View environment, please check this.'
      Success: no output
      Failure: "The script did not get enough arguments from the Console. The Console may not be connected to the Horizon View environment, please check this.", and the script will exit with error code 1
      Test-ArgsCount -ArgsCount $args -Reason 'Please check you are connected to the XXXXX environment in the Console'
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

function Load-VMWareModules {
    <# Imports VMware modules
      NOTES:
      - The required modules to be loaded are passed as an array.
      - In versions of PowerCLI below 6.5 some of the modules cannot be imported (below version 6 it is Snapins only) using so Add-PSSnapin is used (which automatically loads all VMWare modules)
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

Function Connect-VCenterServer {
    param (
        [parameter(Mandatory = $true)]
        [string]$VcenterName
    )
    Try {
        # Connect to VCenter server
        Connect-VIServer $VCenterName
    }
    Catch {
        Out-CUConsole -Message "There was a problem connecting to VCenter server $VCenterName. Please correct the error and try again." -Exception $_
    }
}

Function Disconnect-VCenterServer {
    # This function closes the connection with the VCenter server 'VCenter'
    param (
        [parameter(Mandatory = $true)]
        $Vcenter
    )
    try {
        # Disconnect from the VCenter server
        Disconnect-VIServer -Server $VCenter -Confirm:$false
    }
    catch {
        Out-CUConsole -Message "There was a problem disconnecting from VCenter server $VCenter. This is not a serious problem as the connection will close when this script ends but keep an eye on this recurring." -Warning
        Exit 0
    }
}

Function Get-VMWareVirtualMachine {
    param (
        [parameter(Mandatory = $true)]
        [string]$VMName,
        [parameter(Mandatory = $true)]
        $VCenter
    )
    # This function retrieves the VMware VM
    try {
        Get-VM -Name $VMName -Server $VCenter
    }
    catch {
        Out-CUConsole -Message "There was a problem retrieving virtual machine $VMName." -Exception $_
    }
}

Function Get-VMWareVIMException {
    param (
        [parameter(Mandatory = $true)]
        [Management.Automation.ErrorRecord]$Exception
    )
    If ($_.Exception.MethodFault.GetType().Name -eq 'NoPermission') {
        [string]$Privilege = $_.Exception.MethodFault.PrivilegeID
        [string]$Explanation = (Get-VIPrivilege -Id $Privilege).Description
        [string]$ObjectType = $_.Exception.MethodFault.Object.Type
        [string]$ObjectName = (Get-View -Id $_.Exception.MethodFault.Object).Name | Get-Unique
        Write-Output "The account used to run this is script is missing a required vCenter privilege for $ObjectType $ObjectName`:`n$Privilege ($Explanation)."
    }
    else {
        Write-Output $Exception.Exception.Message
    }
}


# Test correct number of arguments was passed
Test-ArgsCount -ArgsCount 4 -Reason 'The Console or Monitor may not be connected to the ESX environment, please check this.'

# Test correct Hypervisor platform
If ($strHypervisorPlatform -ne 'VMWare') {
    Out-CUConsole -Message "Currently this script based action only supports VMWare, selected guest $strVMName is not running on VMWare" -Stop
}

# Import the VMware PowerCLI modules
Load-VMwareModules -Components @('VimAutomation.Common', 'VimAutomation.SDK', 'VimAutomation.Core', 'VimAutomation.Cis.Core')

# Set PowerCLI configuration
if ($strIgnoreCertificateError -eq 'True') {
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -ParticipateInCeip:$false -Scope Session -Confirm:$false | Out-Null
}
else {
    Set-PowerCLIConfiguration -InvalidCertificateAction Fail -ParticipateInCeip:$false -Scope Session -Confirm:$false | Out-Null
}

# Connect to VCenter server for VMWare
$objVCenter = Connect-VCenterServer -VcenterName $strVCenter

# Get the virtual machine
$objVM = Get-VMWareVirtualMachine -VMName $strVMName -VCenter $objVCenter

# Update the tools, NoReboot (NOTE: the virtual machine might still reboot after updating VMware Tools, depending on the currently installed VMware Tools version, the VMware Tools version to which you want to upgrade, and the vCenter Center/ESX versions.)
try {
    $tskUpdate = Update-Tools -VM $objVM -Server $objVCenter -NoReboot -RunAsync
}
catch {
    Get-VMWareVIMException -Exception $_
}

# Wait for job to complete, max 5 minutes
For ($i = 0; $i -lt 30; $i++) {
    # Wait 10 seconds
    Start-Sleep -Seconds 10

    # Get the current task status
    try {
        $tskRunning = Get-Task -Id $tskUpdate.Id -Server $objVCenter
    }
    catch {
        Get-VMWareVIMException -Exception $_
    }
    If ($tskRunning.State -in 'Success', 'Error') {
        # Task has completed one way or the other, exit the loop and check the results
        Break
    }
}

# Output results
Switch ($tskRunning.State) {
    'Queued' { Write-Output -InputObject "The VMware Client Tools update task is in the queue waiting to be executed and the 5 minute timeout for this script has been exceeded.`nPlease check the Console or VCenter later to ensure it has run successfully." }
    'Running' { Write-Output -InputObject "The VMware Client Tools update task is still running and the 5 minute timeout for this script has been exceeded.`nPlease check the Console or VCenter later to ensure it has run successfully." }
    'Success' { Write-Output -InputObject 'VMware Client Tools updated successfully. Please be aware that the ControlUp Console may have disconnected due to a NIC driver update.' }
    'Error' {
        Write-Output -InputObject 'There was an error in the VMware Client Tools update task. Please check the VM in VCenter for details.'
        # Disconnect Vcenter server
        Disconnect-VCenterServer $objVCenter
        # Exit with error
        Exit 1
    }
}

# Disconnect Vcenter server
Disconnect-VCenterServer $objVCenter
