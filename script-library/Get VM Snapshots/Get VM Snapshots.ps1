$ErrorActionPreference = 'Stop'
<# This script retrieves the snapshots from the selected VMware virtual machine(s).  
    Steps:
    1. Check hypervisor type, only VMware is supported
    2. Connect to VCenter server
    3. Retrieve snapshot information and output
    4. Disconnect from the VCenter server

    Script requires VMware PowerCLI to be installed on the machine it runs on.
    Author: Ton de Vreede, 5/12/2016
#>

$strHypervisorPlatform = $args[0]
$strVMName = $args[1]
$strVCenter = $args[2]

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

Function Get-VMSnapshots {
    Param
    (
        [parameter(mandatory = $true, Position = 0)]
        [string]$strVMName 
    )

    # This function retrieves the snapshots of the machine 'strVMName'
    try {
        $objSnapshots = Get-VMWareSnapshot -VM $strVMName -Server $VCenter | Select-Object -Property Name, Description, Created, SizeMB, ParentSnapshot
        if (!$objSnapshots) {
            Feedback -Message 'No snapshots were found.'
            Exit 0
        } 
        else {
            $objSnapshots |
            Select-Object -Property Name, Description, Created, @{
                Name       = 'SizeMB'
                Expression = {
                    [math]::round($_.SizeMB, 2)
                }
            }, ParentSnapshot |
            Format-Table -AutoSize
        }
    }
    catch {
        Feedback -Message "There was a problem retrieving the snapshots of VM $strVMName" -Exception $_
    }
}

# Check all the arguments have been passsed
If ($args.count -lt 3) {
    Feedback  -Message "The script did not get enough arguments from the Console. This can occur if you are not connected to the VM's hypervisor.`nPlease connect to the hypervisor in the ControlUp Console and try again." -Oops
}

# Check the VM is using a supported hypervisor platform
Test-HypervisorPlatform -strHypervisorPlatform $strHypervisorPlatform

# Import the VMWare PowerCLI modules
Load-VMWareModules -Components @('VimAutomation.Core') -Prefix 'VMWare'

# Connect to VCenter server for VMWare
$VCenter = Connect-VCenterServer -VCenterName $strVCenter

# Get the snapshots
Get-VMSnapshots -strVMName $strVMName

# Disconnect from the VCenter server
Disconnect-VCenterServer -VCenter $VCenter
