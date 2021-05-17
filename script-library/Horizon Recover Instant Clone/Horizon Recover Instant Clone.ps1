#requires -Version 3.0
$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Recovers a Horizon View machine

    .DESCRIPTION
    This script Recovers a Horizon View Instant Clone through the VMware.Hv.Helper module

    .EXAMPLE
    You can use this to 'rebuild' an Instant Clone if there is an issue with the machine.
    
    .NOTES
    This script requires VMWare PowerCLI and the Vmware.Hv.Helper module to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'
    Vmware.Hv.Helper can be installed using the 'Install Hv.Helper module for Horizon View scripts' script. It can also be found on Github (see LINK). Download the module and place it in your systemdrive Program Files\WindowsPowerShell\Modules folder.

    Before running this script you will also need to have a PSCredential object available on the target machine. This can be created by running the 'Create credentials for Horizon View scripts' script in ControlUp on the target machine.
    
    Credits to the various contributors to the Hv.Helper module.
    
    Context: Can be triggered from the Horizon View Machines view
    Modification history: 09/01/2020 - Anthonie de Vreede - First version

    .PARAMETER strHVMachineName
    Name of the Horizon View machine. Passed from the ControlUp Console.
    .PARAMETER strHVMachinePool
    Name of the Horizon View machine Pool. Passed from the ControlUp Console.
    .PARAMETER strHVMachineSource
    Type of machine. Passed from ControlUp Console.
    .PARAMETER strHVConnectionServerFQDN
    Name of the Horizon View connection server. Passed from the ControlUp Console.
    
    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli
    https://github.com/vmware/PowerCLI-Example-Scripts/tree/master/Modules/VMware.Hv.Helper
    
    .COMPONENT
    VMWare PowerCLI 6.5.0R1 or higher
    VMWare Hv.Helper 1.1 or higher
#>

# Name of the Horizon View machine. Passed from the ControlUp Console.
[string]$strHVMachineName = $args[0]
# Name of the Horizon View machine Pool. Passed from the ControlUp Console.
[string]$strHVMachinePool = $args[1]
# Type of machine. Passed from ControlUp Console.
[string]$strHVMachineSource = $args[2]
# Name of the Horizon View connection server. Passed from the ControlUp Console.
[string]$strHVConnectionServerFQDN = $args[3]

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
    <# Imports VMware modules
    NOTES:
    - The required modules to be loaded are passed as an array.
    - In versions of PowerCLI below 6.5 some of the modules can't be imported (below version 6 it is Snapins only) using so Add-PSSnapin is used (which automatically loads all VMWare modules)
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
            Out-CUConsole -Message 'There was a problem connecting to the Horizon View Connection server. It looks like there may be a certificate issue. Please ensure the certificate used on the Horizon View server is trusted by the machine running this script.' -Exception $_
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
function Recover-HorizonViewMachine {
    param (   
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View machine object.")]
        [object]$HVMachine,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View services object.")]
        [object]$HVServices
    )

    try {
        $HVServices.Machine.Machine_Recover($HVMachine.Id)
        Out-CUConsole -Message 'Recover command has been sent to Horizon View.'
    }
    catch {
        Out-CUConsole -Message 'There was a problem Recovering the machine.' -Exception $_
    }
}

# Test arguments
Test-ArgsCount -ArgsCount 4 -Reason 'The Console or Monitor may not be connected to the Horizon View environment, please check this.'

# Check if the machine is an Instant Clone. Recovery only works for Instant and Linked Clones
if ($strHVMachineSource -ne 'vCenter (instant clone)') {
    Out-CUConsole -Message 'This machine is not an Instant Clone, it cannot be recoverred.' -Stop
}

# Import the VMware PowerCLI modules
Load-VMwareModules -Components @('VimAutomation.HorizonView', 'Hv.Helper')

# Get the stored credentials for running the script
[PSCredential]$CredsHorizon = Get-CUStoredCredential -System 'HorizonView'

# Connect to the Horizon View Connection Server
[VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $strHVConnectionServerFQDN -Credential $CredsHorizon

# Get the machine
try {
    $objHorizonViewMachine = Get-HvMachine -MachineName $strHVMachineName -PoolDisplayName $strHVMachinePool
}
catch {
    Out-CUConsole -Message "There was a problem retreiving the Horizon View machine. Please check the exception for details" -Exception $_
}

# Create Horizon View services object
$objHVServices = $objHVConnectionServer.ExtensionData

# Recover the Horizon View Machine
Recover-HorizonViewMachine -HVMachine $objHorizonViewMachine -HVServices $objHVServices

# Disconnect from the Horizon View Connection Center
Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
