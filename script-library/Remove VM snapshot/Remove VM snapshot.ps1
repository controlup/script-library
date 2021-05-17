$ErrorActionPreference = 'Stop'
<# This script removes a named snapshot from the selected VMware virtual machine(s).
    If the name is left blank it will remove the latest snapshot. 
    Steps:
    1. Check hypervisor type, only VMware is supported
    2. Connect to VCenter server
    3. Delete snapshot
    - If snapshot nam was entered, this snapshot will be deleted
    - If snapshot is left blank, the LATEST snapshot will be deleted
    - If snapshot has children, AND the parameter $bolRemoveWithChildren is $true the snapshot and child snapshots will be deleted, if it is false the script stops.
    4. Disconnect from the VCenter server

    Script requires VMWare PowerCLI to be installed on the machine it runs on.

    Author: Ton de Vreede, 7/12/2016
#>

$strHypervisorPlatform = $args[0]
$strVCenter = $args[1]
$strVMName = $args[2]
$strRemoveWithChildren = $args[3]
$strSnapshotName = $args[4]

# Set bool $bolRemoveWithChildren
[bool]$bolRemoveWithChildren = $false
If ($strRemoveWithChildren -eq 'Yes') { $bolRemoveWithChildren = $true }

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


function Load-VMWareModules {
    <# Imports VMware PowerCLI modules, with a -Prefix $Prefix is supplied (desirable to avoid conflict with Hyper-V cmdlets)
      NOTES:
      - The required modules to be loaded are passed as an array.
      - If the PowerCLI versions is below 6.5 some of the modules can't be imported (below version 6 it is Snapins only) using so Add-PSSnapin is used (which automatically loads all VMWare modules) BUT 
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
            Write-Host 'The required VMWare PowerCLI components were not found as PowerShell modules. Because a -Prefix is used in loading the components Add-PSSnapin can not be used. Please make sure VMWare PowerCLI version 6.5 or higher is installed and available for the user running the script.'
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
                    Write-Host 'The required VMWare PowerCLI components were not found as modules or snapins. Please make sure VMWare PowerCLI (version 6.5 or higher preferred) is installed and available for the user running the script.'
                    Exit 1
                }
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
        Connect-VMWareVIServer -Server $VCenterName -WarningAction SilentlyContinue -Force
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
        Disconnect-VMWareVIServer -Server $VCenter -Confirm:$false
    }
    catch {
        Feedback -Message "There was a problem disconnecting from VCenter server $VCenter." -Exception $_
    }
}

Function Test-HypervisorPlatform {
    Param (
        [Parameter(Mandatory = $true,
            Position = 0)]
        [string]$strHypervisorPlatform
    )
    # This function checks if the hypervisor is supported by this script.
    If ($strHypervisorPlatform -ne 'VMWare') {
        Feedback -Message "Currently this script based action only supports VMWare, selected guest $strVMName is not running on VMWare" -Oops
    }
}

Function Check-VMSnapshot ($strVMName, $bolRemoveWithChildren, $strSnapshotName) {
    # This function handles checking the snapshot name (remove latest snapshot if left blank) and what to do if the snapshot has children
    If (!$strSnapshotName) {
        # Check if snapshot name is blank, in which case $strSnapshotname will be the latest snapshot on the machine
        try {
            $objSnapshot = Get-VMWareSnapshot -VM $strVMName |
            Sort-Object -Property Created |
            Select-Object -Last 1
        }
        Catch {
            Feedback -Message "There was a problem retrieving the snapshots for machine $strVMName." -Exception $_
        }
        If ($objSnapshot) {
            Remove-VMSnapshot -objSnapshot $objSnapshot
        }
        Else {
            Feedback -Message 'There are no snapshots for this VM.'
        }
    }
    Else {
        # Retrieve the snapshot and check if it is a parent
        try {
            $objSnapshot = Get-VMWareSnapshot -VM $strVMName -Name $strSnapshotName
        }
        Catch {
            Feedback -Message "There was a problem retrieving the snapshot $strSnapshotName for machine $strVMName." -Exception $_
        } 
    
        # Check, only remove snapshot if it has no child OR if it does then only remove snapshot if $bolRemoveWithChildern is $true
        If ($objSnapshot.Children) {
            If ($bolRemoveWithChildren -eq $true) {
                Remove-VMSnapshotChain -objSnapshot $objSnapshot
            }
            Else {
                Feedback -Message "Snapshot $strSnapshotName for machine $strVMName could not be deleted as it is a parent snapshot and you have chosen not to delete snapshots that have child snapshots." -Oops
                Exit 0
            }
        }
        Else {
            Remove-VMSnapshot -objSnapshot $objSnapshot
        }
    }
}

Function Remove-VMSnapshot ($objSnapshot) {
    try {
        # Remove the snapshot
        $objSnapshot | Remove-VMWareSnapshot -RemoveChildren:$false -Confirm:$false >$null
    
        # Report the result
        Feedback -Message "Snapshot $objSnapshot (size $([math]::round($objSnapshot.SizeMB,2)) MB) was deleted."
    }
    catch {
        Feedback -Message "There was a problem deleting snapshot $strSnapshotName." -Exception $_
    }
}


Function Remove-VMSnapshotChain ($objSnapshot) {
    # This function removes a chain of snapshots of machine '$strVMName'
    try {
        # Remove the snapshots

        $objSnapshot | Remove-VMWareSnapshot -RemoveChildren:$true -Confirm:$false >$null
    
        # Report the result
        Feedback -Message "$strSnapshotName and its child snapshot(s) were deleted from machine $strVMName"
    }
    catch {
        Feedback -Message "There was a problem deleting $strSnapshotName and the child snapshot(s) of machine $strVMName" -Exception $_
    }
}

# Check all the arguments have been passsed
If ($args.count -lt 5) {
    Feedback  -Message "The script did not get enough arguments from the Console. This can occur if you are not connected to the VM's hypervisor.`nPlease connect to the hypervisor in the ControlUp Console and try again." -Oops
}

# Check the Hypervisor is supported
Test-HypervisorPlatform -strHypervisorPlatform $strHypervisorPlatform

# Import the VMWare PowerCLI modules
Load-VMWareModules -Components @('VimAutomation.Core') -Prefix 'VMWare'

# Connect to VCenter server for VMWare
$VCenter = Connect-VCenterServer -VCenterName $strVCenter

# Remove the snapshots
Check-VMSnapshot  -strVMName $strVMName -bolRemoveWithChildren $bolRemoveWithChildren -strSnapshotName $strSnapshotName

# Disconnect from the VCenter server
Disconnect-VCenterServer -VCenter $VCenter
