$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Changes the amount of Desktops in a Horizon Desktop Pool

    .DESCRIPTION
    This script changes the amount of Desktops in a Horizon Desktop Pool. 

    .NOTES
    This script requires VMWare PowerCLI to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers' Or by using the 'Install VMware PowerCLI' script.
    Credentials can be set using the 'Prepare machine for Horizon View scripts' script.

    Modification history:   12/12/2019 - Wouter Kursten - First version

    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli


    .COMPONENT
    VMWare PowerCLI

#>

# Name of the Horizon View published Application. Passed from the ControlUp Console.
[string]$HVDesktopPoolname = $args[0]
# Name of the Horizon View connection server. Passed from the ControlUp Console.
[string]$HVConnectionServerFQDN = $args[1]
# Desired enablement state of the published application.
[int]$HVDesktopCount = $args[2]

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
    [string]$strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
    Import-Clixml $strCUCredFolder\$($env:USERNAME)_$($System)_Cred.xml
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
    # Try to connect to the Connection server
    try {
        Connect-HVServer -Server $HVConnectionServerFQDN -Credential $Credential
    }
    catch {
        Out-CUConsole -Message 'There was a problem connecting to the Horizon View Connection server.' -Exception $_
    }
}

function Disconnect-HorizonConnectionServer {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    # Try to connect from the connection server
    try {
        Disconnect-HVServer -Server $HVConnectionServer -Confirm:$false
    }
    catch {
        Out-CUConsole -Message 'There was a problem disconnecting from the Horizon View Connection server.' -Exception $_
    }
}

function Get-HVDesktopPool {
    param (
        [parameter(Mandatory = $true,
        HelpMessage = "Name of the Desktop Pool.")]
        [string]$HVPoolName,
        [parameter(Mandatory = $true,
        HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    # Try to get the Desktop pools in this pod
    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'DesktopSummaryView'
        # Filter oud rds desktop pools since they don't contain machines
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='desktopSummaryData.displayName'; 'value' = "$HVPoolname"}
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        if (!$queryResults){
            Out-CUConsole -message "Can't find $HVPoolName, exiting." -Exception "$HVPoolname not found"
            exit
        }
        elseif (($queryResults).desktopsummarydata.type -eq "MANUAL"){
            Out-CUConsole -Message 'This a manual Horizon View Desktop Pool, cannot change the amount of desktops.' -Exception
        }
        elseif (($queryResults).desktopsummarydata.source -eq "VIRTUAL_CENTER"){
            Out-CUConsole -Message 'This a Full Clone Horizon View Desktop Pool, if the amount of desktops has been reduced the extra systems need to be removed manually.'
            return $queryResults
        }
        else {
            return $queryResults
        }
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the Horizon View Desktop Pool.' -Exception $_
    }
}

function get-hvpoolspec{
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "ID of the Desktop Pool.")]
        [VMware.Hv.DesktopId]$HVPoolID,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    try {
        $HVConnectionServer.ExtensionData.Desktop.Desktop_Get($HVPoolID)
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the desktop pool details.' -Exception $_
    }
}

function Set-HVPool {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "ID of the Desktop Pool.")]
        [VMware.Hv.DesktopId]$HVPoolID,
        [parameter(Mandatory = $true,
            HelpMessage = "Desired amount of desktops in the pool.")]
        [int]$HVDesktopCount,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    try {
        # First define the Service we need
        [VMware.Hv.DesktopService]$desktopservice=new-object vmware.hv.DesktopService
        # Fill the helper for this service with the application information
        $desktophelper=$desktopservice.read($HVConnectionServer.extensionData, $HVPoolID)
        # Change the state of the application in the helper
        $desktophelper.getAutomatedDesktopDataHelper().getVmNamingSettingsHelper().getPatternNamingSettingsHelper().setMaxNumberOfMachines($HVDesktopCount)
        # Apply the helper to the actual object
        $desktopservice.update($HVConnectionServer.extensionData, $desktophelper)
    }
    catch {
        Out-CUConsole -Message 'There was a problem changing the desktop count.' -Exception $_
    }
}


# Test arguments
Test-ArgsCount -ArgsCount 3 -Reason 'The Console or Monitor may not be connected to the Horizon View environment, please check this.'

# Set the credentials location
[string]$strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"

# Import the VMware PowerCLI modules
Load-VMwareModules -Components @('VimAutomation.HorizonView')

# Get the stored credentials for running the script
[PSCredential]$CredsHorizon = Get-CUStoredCredential -System 'HorizonView'

# Connect to the Horizon View Connection Server

[VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $HVConnectionServerFQDN -Credential $CredsHorizon

# Retreive the desktop pool
$HVPool=Get-HVDesktopPool -HVPoolName $HVDesktopPoolname -HVConnectionServer $objHVConnectionServer
Out-CUConsole -Message "Retreived information about $HVDesktopPoolname" -Verbose

# But we only need the ID
$HVPoolID=($HVPool).id

# Retreive the pool spec
$hvpoolspec=Get-HVPoolSpec -HVConnectionServer $objHVConnectionServer -HVPoolID $HVPoolID
$ProvisioningTime=($hvpoolspec).AutomatedDesktopData.VmNamingSettings.PatternNamingSettings.ProvisioningTime

# We cannot change manual pools so we give a warning about this and exit the script.
if ($hvpoolspec.Type -eq "MANUAL"){
    Out-CUConsole -Message 'Could not execute, this a manual Horizon View Desktop Pool, cannot change the amount of desktops.' -warning $_
    exit
}

# When not all vm's are provisioned up front the max amount of machines can't be lower that the minimum amount or the number of spare machines.
if ($ProvisioningTime -eq "ON_DEMAND"){
    $NumberOfSpareMachines=($hvpoolspec).AutomatedDesktopData.VmNamingSettings.PatternNamingSettings.NumberOfSpareMachines
    $MinNumberOfMachines=($hvpoolspec).AutomatedDesktopData.VmNamingSettings.PatternNamingSettings.MinNumberOfMachines
    if ($NumberOfSpareMachines -ge $HVDesktopCount -or $MinNumberOfMachines -ge $HVDesktopCount){
        Out-CUConsole -Message 'Could not execute, the number of desktops cannot be smaller than the minimum amount of desktops or the number of spare desktops' -warning $_
        exit
    }
}

# Change the desktop count in the pool
Out-CUConsole -Message "Trying to change $HVDesktopPoolname to $HVDesktopCount desktops."
Set-HVPool -HVConnectionServer $objHVConnectionServer -HVPoolID $HVPoolID -HVDesktopCount $HVDesktopCount
Out-CUConsole -Message "Changed $HVDesktopPoolname to $HVDesktopCount desktops."

# Disconnect from the connection server
Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
