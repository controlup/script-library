#requires -Version 3.0
$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Place Horizon View machine in Maintenance

    .DESCRIPTION
    This script places a Horizon View machine in Maintenance through the VMware.Hv.Helper module

    .EXAMPLE
    Can be used as an Automated action to place a machine in maintenance if counters indicate there is an issue with the machine, or for planned maintenance windows.
    
    .NOTES
    This script requires VMWare PowerCLI and the Vmware.Hv.Helper module to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'
    Vmware.Hv.Helper can be installed using the 'Prepare machine for Horizon View scripts' script. It can also be found on Github (see LINK). Download the module and place it in your systemdrive Program Files\WindowsPowerShell\Modules folder 

    Credits to the various contributors to the Hv.Helper module.
    
    Context: Can be triggered from the Horizon View Machines view
    Modification history: 13/08/2019 - Anthonie de Vreede - First version
    
    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli
    https://github.com/vmware/PowerCLI-Example-Scripts/tree/master/Modules/VMware.Hv.Helper
    
    .COMPONENT
    VMWare PowerCLI
    VMWare Hv.Helper 
#>

# Name of the Horizon View machine. Passed from the ControlUp Console.
[string]$strHVMachineName = $args[0]
# Name of the Horizon View connection server. Passed from the ControlUp Console.
[string]$strHVConnectionServerFQDN = $args[1]

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
    $strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
    try {
        Import-Clixml -LiteralPath $strCUCredFolder\$($env:USERNAME)_$($System)_Cred.xml
    }
    catch {
        Out-CUConsole -Message "The required PSCredential object could not be loaded. Please make sure you have run the 'Create credentials for Horizon View scripts' script on the target machine." -Exception $_
    }
}

function Load-VMWareModules {
    <# Imports VMware PowerCLI modules
      NOTES:
      - The required modules to be loaded are passed as an array.
      - If the PowerCLI versions is below 6.5 some of the modules can't be imported (below version 6 it is Snapins only) using so Add-PSSnapin is used (which automatically loads all VMWare modules)
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
                Out-CUConsole -Message 'The required VMWare PowerCLI components were not found as modules or snapins. Please make sure VMWare PowerCLI (version 6.5 or higher preferred) is installed and available for the user running the script.' -Stop
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

    try {
        Connect-HVServer -Server $HVConnectionServerFQDN -Credential $Credential
    }
    catch {
        if ($_.Exception.Message.StartsWith('Could not establish trust relationship for the SSL/TLS secure channel with authority')) {
            Out-CUConsole -Message 'There was a problem connecting to the Horizon View Connection server. It looks like there may be a certificate issue. Please ensure the certificate authority used for the Horizon View server is trusted by the machine running this script.' -Exception $_
        }
        else {
            Out-CUConsole -Message 'There was a problem connecting to the Horizon View Connection server.' -Exception $_
        }
    }
}

function Disconnect-HorizonConnectionServer {
    param (    
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    
    try {
        Disconnect-HVServer -Server $HVConnectionServer -Confirm:$false
    }
    catch {
        Out-CUConsole -Message 'There was a problem disconnecting from the Horizon View Connection server. If not running in a persistent session (ControlUp scripts do not run in a persistant session) this is not a problem, the session will eventually be deleted by Horizon View.' -Warning
    }
}

function Get-HorizonViewMachine {
    param (    
        [parameter(Mandatory = $true,
            HelpMessage = "Name of the Horizon View machine.")]
        [string]$HVMachineName,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )

    try {
        Get-HVMachine -MachineName $HVMachineName -HvServer $HVConnectionServer
    }
    catch {
        Out-CUConsole -Message "There was an issue retrieving the Horzion View machine." -Exception $_
    }
}

function Set-HorizonViewMachineMaintenanceMode {
    param (   

        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View machine object.")]
        [Object]$HVMachine,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer,
        [parameter(Mandatory = $true,
            HelpMessage = "Enter or Leave.")]
        [string]$MaintenanceMode
    )

    if ($MaintenanceMode -eq 'Enabled') { 
        try {
            Set-HVMachine -Machine $HVMachine -HvServer $HVConnectionServer -Maintenance ENTER_MAINTENANCE_MODE
            Out-CUConsole -Message 'Machine entered Maintenance mode.'
        }
        catch {
            Out-CUConsole -Message 'There was a problem enabling Maintenance for the machine.' -Exception $_
        }
    }
    else {

        try {
            Set-HVMachine -Machine $HVMachine -HvServer $HVConnectionServer -Maintenance EXIT_MAINTENANCE_MODE
            Out-CUConsole -Message 'Machine left Maintenance mode.'
        }
        catch {
            Out-CUConsole -Message 'There was a problem disabling Maintenance for the machine.' -Exception $_
        }
    }
}

# Test arguments
Test-ArgsCount -ArgsCount 2 -Reason 'The Console or Monitor may not be connected to the Horizon View environment, please check this.'

# Import the VMware PowerCLI modules
Load-VMwareModules -Components @('VimAutomation.HorizonView', 'Hv.Helper')

# Get the stored credentials for running the script
[PSCredential]$CredsHorizon = Get-CUStoredCredential -System 'HorizonView'

# Connect to the Horizon View Connection Server
[VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $strHVConnectionServerFQDN -Credential $CredsHorizon

# Get the Horizon View machine
[Object]$objHVMachine = Get-HorizonViewMachine -HVMachineName $strHVMachineName -HVConnectionServer $objHVConnectionServer

# Set Horizon View machine in Maintenance
Set-HorizonViewMachineMaintenanceMode -HVMachine $objHVMachine -HVConnectionServer $objHVConnectionServer -MaintenanceMode Enabled

# Disconnect from the Horizon View Connection Center
Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
