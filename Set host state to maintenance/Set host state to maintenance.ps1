#requires -version 3
$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    This script sets a VMware host in Maintenance state.

    .DESCRIPTION
    Placing a host in Maintenance will migrate the powerred on VMs to other hosts in the cluster. If the Evacuate switch is passed all offline machines a migrated too.
    This script will only place a host in Maintenenance if the cluster it is part of is DRSFullyAutomated.
    
    .PARAMETER strHypervisorPlatform
    The hypervisor platform, will be passed from ControlUp Console

    .PARAMETER strVCenter
    The name of the vcenter server that will be connected to to run the PowerCLI commands

    .PARAMETER strVMwareHost
    The name of the virtual machine the action is to be performed on/for.

    .PARAMETER bolEvacuate
    Should the host be evacuated as well?

    .PARAMETER bolVSAN
    Does this host use VSAN?
        - If the host uses a VSAN the default VsanEvacuationMode setting will be used

    .EXAMPLE
    Example is not relevant as this script will be called through ControlUp Console

    .NOTES
    VMware PowerCLI Core needs to be installed on the machine running the script.
    Loading VMware PowerCLI will result in a 'Join our CEIP' message. In order to disable these in the future run the following commands on the target system:
    Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false (or $true, if that's your kind of thing)
    
 
#>

[string]$strHypervisorPlatform = $args[0]
[string]$strVCenter = $args[1]
[string]$strVMwareHost = $args[2]
If ($args[3] -eq 'True') { [bool]$bolEvacuate = $true } Else { [bool]$bolEvacuate = $false }
If ($args[4] -eq 'True') { [bool]$bolVSAN = $true } Else { [bool]$bolVSAN = $false }

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

Function Connect-VCenterServer {
    Param (
        [Parameter(Mandatory = $true,
            Position = 0)]
        [string]$VCenterName
    )
    Try {
        # Connect to VCenter server
        Connect-VMwareVIServer -Server $VCenterName -WarningAction SilentlyContinue -Force
    }
    Catch {
        Feedback -Message "There was a problem connecting to VCenter server $VCenterName. Please correct the error and try again." -Exception $_
    }
}

Function Disconnect-VCenterServer {
    Param (
        [Parameter(Mandatory = $true,
            Position = 0)]
        [string]$VCenter
    )
    # This function closes the connection with the VCenter server 'VCenter'
    try {
        # Disconnect from the VCenter server
        Disconnect-VMwareVIServer -Server $VCenter -Confirm:$false
    }
    catch {
        Feedback -Message "There was a problem disconnecting from VCenter server $VCenter." -Exception $_
    }
}

function Get-VMwareHost {
    param (
        [parameter(Mandatory = $true,
            Position = 0)]
        [string]$strVMwareHost,
        [parameter(Mandatory = $true,
            Position = 1)]
        $VCenter
    )
    try {
        Get-VMwareVMHost -Name $strVMwareHost -Server $VCenter
    }
    catch {
        Feedback -Message "There was a problem retrieving the VMware host machine $strVMwareHost." -Exception $_
    }
}

function Get-VMwareVMCluster {
    param (
        [Parameter(Mandatory = $true,
            Position = 0)]
        $VMHost,
        [parameter(Mandatory = $true,
            Position = 1)]
        $VCenter
    )
    try {
        Get-VMwareCluster -VMHost $VMHost -Server $VCenter
    }
    catch {
        Feedback -Message 'The VM Cluster (used for checking if cluster is Fully Automated) could not be retreived.' -Exception $_
    }
}

function Load-VMwareModules {
    <# Imports VMware PowerCLI modules, with a -Prefix $Prefix is supplied (desirable to avoid conflict with Hyper-V cmdlets)
      NOTES:
      - The required modules to be loaded are passed as an array.
      - If the PowerCLI versions is below 6.5 some of the modules can't be imported (below version 6 it is Snapins only) using so Add-PSSnapin is used (which automatically loads all VMware modules) BUT 
        If a -Prefix has been specified AND Add-PSSnapin has to be used because the PowerCLI version is too low loading the module will fail because the -Prefix option can't be used with Add-PSSnapin
    #>
    param (    
        [parameter(Mandatory = $true,
            ValueFromPipeline = $false)]
        [array]$Components,
        [parameter(Mandatory = $false,
            ValueFromPipeline = $false)]
        [string]$Prefix
    )
    # Try Import-Module for each passed component, try Add-PSSnapin if this fails (only if -Prefix was not specified)
    if ($Prefix) {
        try {
            # Import each specified module
            foreach ($component in $Components) {
                $null = Import-Module -Name VMware.$component -Prefix $Prefix
            }
        }
        catch {
            Write-Host 'The required VMware PowerCLI components were not found as PowerShell modules. Because a -Prefix is used in loading the components Add-PSSnapin can not be used. Please make sure VMware PowerCLI version 6.5 or higher is installed and available for the user running the script.'
            Exit 1
        }
    }
    else {
        # Import each module, if Import-Module fails try Add-PSSnapin
        foreach ($component in $Components) {
            try {
                $null = Import-Module -Name VMware.$component
            }
            catch {
                try {
                    $null = Add-PSSnapin -Name VMware.$component
                }
                catch {
                    Write-Host 'The required VMware PowerCLI components were not found as modules or snapins. Please make sure VMware PowerCLI (version 6.5 or higher preferred) is installed and available for the user running the script.'
                    Exit 1
                }
            }
        }
    }
}


function Set-VMwareHostMaintenance {
    param (
        [parameter(Mandatory = $true,
            Position = 0)]
        $VMHost,
        [parameter(Mandatory = $true,
            Position = 1)]
        $VCenter
    )

    # Set host state in Maintenance, Evacuate as well if $bolEvacuate is True
    try {
        If ($bolEvacuate) {
            Set-VMwareVMHost -VMHost $VMHost -Server $VCenter -State Maintenance -Evacuate
        }
        Else {
            Set-VMwareVMHost -VMHost $VMHost -Server $VCenter -State Maintenance
        }
    }
    catch {
        Feedback -Message 'The Host could not be placed in Maintenance.' -Exception $_
    }
}

# Check all the arguments have been passsed
if ($args.Count -ne 5) {
    Feedback  -Message 'The script did not get enough arguments from the Console. This can occur if you are not connected to the hypervisor.' -Oops
}

# Check that the host is a supported hypervisor
Test-HypervisorPlatform -strHypervisorPlatform $strHypervisorPlatform

# Import the VMware PowerCLI modules
Load-VMwareModules -Components @('VimAutomation.Core') -Prefix 'VMware'

# Connect to VCenter server for VMware$DefaultVIServer
$VCenter = Connect-VCenterServer -VCenterName $strVCenter

# Get the VMware Host
$VMHost = Get-VMwareHost -strVMwareHost $strVMwareHost -VCenter $VCenter

# Get the cluster to check if cluster is Fully Automated
$VMCluster = Get-VMwareVMCluster -VMHost $VMHost -VCenter $VCenter

<# Check if the cluster is DRSFullyAutomated
    Before entering maintenance mode, if the host is fully automated, the cmdlet first relocates all powered on virtual machines. If the host is not automated or partially automated, you must first generate a DRS recommendation
    and wait until all powered on virtual machines are relocated or powered off. In this case, you must specify the RunAsync parameter, otherwise an error is thrown.
#>
if ($VMCluster.DRSAutomationLevel -ne 'FullyAutomated') {
    Feedback -Message "Host $strVMwareHost State could not be set to maintenance as the cluster it is part of ($($VMCluster.Name)) is not DRS Fully Automated." -Oops
}

# Check if VSAN support is required, if so module VMware.VimAutomation.Core needs to be at least 6
if ($bolVSAN) {
    if ((Get-Module -Name 'Vmware.VimAutomation.Core' -ListAvailable).Version.Major -lt 6) {
        Feedback -Message 'You have specified this is a host with VSAN, but your PowerCLI version seems to be too low to fully support VSAN. Please ensure you have PowerCLI version 6 or greater installed.' -Oops
    }
    else {
        If (!$VMCluster.VsanEnabled) {
            # Module version is high enough, but is this a VSANEnabled cluster?
            Feedback -Message 'You have specified this is a host WITH VSAN, but the cluster is not VSANEnabled.' -Oops
        }
    }
}

else {
    If ($VMCluster.VsanEnabled) {
        Feedback -Message 'You have specified this is a host WITHOUT VSAN, but the cluster is VSANEnabled.' -Oops
    }
    elseif ($VMCluster.VsanEnabled -eq $null) {
        Feedback -Message 'Unable to determine if this is a VSANEnabled cluster (PowerCLI version probably too low, version 6 or greater required). For safety reasons this script will exit.' -Oops
    }
}

# Set the host in Maintenance
Set-VMwareHostMaintenance -VMHost $VMHost -VCenter $VCenter

# Disconnect from the VCenter server
Disconnect-VCenterServer -VCenter $VCenter
