#requires -version 3
$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Changes the VM resource allocation

    .DESCRIPTION
    This script will change the resource allocation  of a given VM in VMWare. By default, the allocation is increased one 'SharesLevel' for CPU, HDD, Memory or all three.
    If the SharesLevel for a resource is 'Custom' this will not be changed. ALL hard disks will be set to a new SharesLevel based on the current level of the FIRST disk. Example, script is set to Increase level for All resources:
    CPU SharesLevel 'Normal' ---> CPU ShareLevel 'High'
    Memory SharesLevel 'Custom' ---> Memory SharesLevel 'Custom'
    FIRST HDD SharesLevel 'Low', SECOND HDD SharesLevel 'Normal' ---> ALL HDD SharesLevel 'Normal'

    .PARAMETER strHypervisorPlatform
    The hypervisor platform, will be passed from ControlUp Console

    .PARAMETER strVCenter
    The name of the vcenter server that will be connected to to run the PowerCLI commands

    .PARAMETER strVMName
    The name of the virtual machine the action is to be performed on/for.

    .PARAMETER strAction
    The desired action to be performed. Increase, Decrease or Report

    .PARAMETER strResourceType
    The resource to be changed, CPU, Memory, HDD or All

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
[string]$strAction = $args[3]
[string]$strResourceType = $args[4]

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

Function Get-VMWareVirtualMachine {
    param (
        [parameter(Mandatory = $true)]
        [string]$strVMName,
        [parameter(Mandatory = $true)]
        $objVCenter
    )
    # This function retrieves the VMware VM
    try {
        Get-VMWareVM -Name $strVMName -Server $objVCenter
    }
    catch {
        Feedback ("There was a problem retrieving virtual machine $strVMName.")
    }
}

function Get-VMWareVirtualMachineResourceConfiguration {
    param (
        [parameter(Mandatory = $true)]
        $objVM,
        [parameter(Mandatory = $true)]
        $objVCenter
    )
    try {
        Get-VMWareVMResourceConfiguration -VM $objVM -Server $objVCenter
    }
    catch {
        Feedback 'The resource configuration of the virtual machine could not be retreived.'
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
if ($args.Count -ne 5) {
    Feedback -Oops  "The script did not get enough arguments from the Console. This can occur if you are not connected to the VM's hypervisor.`nPlease connect to the hypervisor in the ControlUp Console and try again."
}

# Check that the host is a supported hypervisor
Test-HypervisorPlatform $strHypervisorPlatform

# Import the VMWare PowerCLI module
Import-VMWareVIMAutomationCoreModule '5.1.0.0'

# Connect to VCenter server for VMWare
$objVCenter = Connect-VCenterServer $strVCenter

# Get the VM
$objVM = Get-VMWareVirtualMachine $strVMName $objVCenter

# Get the current SharesLevel configuration
$objVMResourceConfig = Get-VMWareVirtualMachineResourceConfiguration $objVM $objVCenter

# Create command string start
[string]$cmdSetConfig = 'Set-VMWareVMResourceConfiguration -Configuration $objVMResourceConfig'

# Set the resources to be changed if required
switch ($strResourceType) {
    'All' {
        [bool]$bolCPU = $true
        [bool]$bolMem = $true
        [bool]$bolHDD = $true
        break
    }
    'CPU' {
        [bool]$bolCPU = $true
        [bool]$bolMem = $false
        [bool]$bolHDD = $false
        break
    }
    'Memory' {
        [bool]$bolCPU = $false
        [bool]$bolMem = $true
        [bool]$bolHDD = $false
        break
    }
    'HDD' {
        [bool]$bolCPU = $false
        [bool]$bolMem = $false
        [bool]$bolHDD = $true
        break
    }
}

# Check the current SharesLevels and perform requested operation. Only perform on requested type of resource
switch ($strAction) {
    'Increase' {
        [string]$cmdSetConfig = 'Set-VMWareVMResourceConfiguration -Configuration $objVMResourceConfig'
        # Set new levels, increasing level if possible

        if ($bolCPU) {
            # Get the current CPU Level, add portion to command string if level is not Custom or remains the same
            [string]$lvlCPULevel = $objVMResourceConfig.CPUSharesLevel

            # If level is Low or Normal, add portion to command string to increase level
            if ($lvlCPULevel -eq 'Low') {$cmdSetConfig += ' -CPUSharesLevel Normal'}
            if ($lvlCPULevel -eq 'Normal') {$cmdSetConfig += ' -CPUSharesLevel High'}
            
        }

        if ($bolMem) {
            # Get the current Memory Level, add portion to command string if level is not Custom or High
            [string]$lvlMemLevel = $objVMResourceConfig.MemSharesLevel

            # If level is Low or Normal, add portion to command string to increase level
            if ($lvlMemLevel -eq 'Low') {$cmdSetConfig += ' -MemSharesLevel Normal'}
            if ($lvlMemLevel -eq 'Normal') {$cmdSetConfig += ' -MemSharesLevel High'}
        }

        if ($bolHDD) {
            # Add command portion for the hard disks. If there is more than one disk, level will be based on the first hard disk
            [string]$lvlHDDLevel = $objVMResourceConfig.DiskResourceConfiguration[0].DiskSharesLevel

            # If level is Low or Normal, add portion to command string to increase level
            if ($lvlHDDLevel -eq 'Low') {$cmdSetConfig += ' -Disk (Get-VMWareHardDisk -VM $objVM) -DiskSharesLevel Normal'}
            if ($lvlHDDLevel -eq 'Normal') {$cmdSetConfig += ' -Disk (Get-VMWareHardDisk -VM $objVM) -DiskSharesLevel High'}
        }

        # Run the reconfiguration command
        try {
            Invoke-Expression $cmdSetConfig | Out-Null 
        }
        catch {
            Feedback 'The Resource ShareLevels could not be set.'
        }
    }
    'Decrease' {
        # Set new levels, decreasing level if possible
        [string]$cmdSetConfig = 'Set-VMWareVMResourceConfiguration -Configuration $objVMResourceConfig'

        if ($bolCPU) {
            # Get the current CPU Level, add portion to command string if level is not Custom or remains the same
            [string]$lvlCPULevel = $objVMResourceConfig.CPUSharesLevel

            # If level is Normal or High, add portion to command string to decrease level
            if ($lvlCPULevel -eq 'High') {$cmdSetConfig += ' -CPUSharesLevel Normal'}
            if ($lvlCPULevel -eq 'Normal') {$cmdSetConfig += ' -CPUSharesLevel Low'}
        }

        if ($bolMem) {
            # Get the current Memory Level, add portion to command string if level is not Custom or High
            [string]$lvlMemLevel = $objVMResourceConfig.MemSharesLevel

            # If level is Normal or High, add portion to command string to decrease level
            if ($lvlMemLevel -eq 'High') {$cmdSetConfig += ' -MemSharesLevel Normal'}
            if ($lvlMemLevel -eq 'Normal') {$cmdSetConfig += ' -MemSharesLevel Low'}
        }

        if ($bolHDD) {
            # Add command portion for the hard disks. If there is more than one disk, level will be based on the first hard disk
            [string]$lvlHDDLevel = $objVMResourceConfig.DiskResourceConfiguration[0].DiskSharesLevel

            # If level is Normal or High, add portion to command string to decrease level
            if ($lvlHDDLevel -eq 'High') {$cmdSetConfig += ' -Disk (Get-VMWareHardDisk -VM $objVM) -DiskSharesLevel Normal'}
            if ($lvlHDDLevel -eq 'Normal') {$cmdSetConfig += ' -Disk (Get-VMWareHardDisk -VM $objVM) -DiskSharesLevel Low'}
        }
        
        # Run the reconfiguration command
        try {
            Invoke-Expression $cmdSetConfig | Out-Null 
        }
        catch {
            Feedback 'The Resource ShareLevels could not be set.'
        }
    }
}

# Refresh the configuration and report, if 'Report' option was selected config does not need to be refreshed
If ($strAction -ne 'Report') {$objVMResourceConfig = Get-VMWareVirtualMachineResourceConfiguration $objVM $objVCenter}

# Output with some formatting
$objVMResourceConfig | Select-Object CpuSharesLevel, MemSharesLevel, @{Name = 'Disk(s)SharesLevel'; Expression = {$_.DiskResourceConfiguration.DiskSharesLevel}} | Format-Table -AutoSize

# Disconnect from the VCenter server
Disconnect-VCenterServer $objVCenter
