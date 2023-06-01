#requires -Version 3

<#
    .SYNOPSIS
    Enables a VMware Horizon RDS Server

    .DESCRIPTION
    This script enables a VMware Horizon RDS Server using the VMware Horizon SOAP APIs via PowerCLI

    .EXAMPLE
    You can use this to 'enable' a RDS Server if there is an issue with the machine.

    .NOTES
    This script requires VMWare PowerCLI to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'

    Before running this script you will also need to have a PSCredential object available on the target machine. This can be created by running the 'Create credentials for VMware Horizon scripts' script in ControlUp on the target machine.

    Context: Can be triggered from the VMware Horizon Machines view

    Modification history: 06/04/2023 - Wouter Kursten - First Version
                          22/05/2023 - Guy Leech      - Added scripting standards & combined to single script

    .PARAMETER strHVMachineName
    Name of the VMware Horizon machine. Passed from the ControlUp Console.
    .PARAMETER strHVMachineFarm
    Name of the VMware Horizon machine Pool. Passed from the ControlUp Console.
    .PARAMETER strHVConnectionServerFQDN
    Name of the VMware Horizon connection server. Passed from the ControlUp Console.

    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli

    .COMPONENT
    VMWare PowerCLI 11 or higher
#>

[CmdletBinding()]

Param
(
    # Name of the VMware Horizon machine. Passed from the ControlUp Console.
    [string]$strHVMachineName ,

    # Name of the VMware Horizon machine Pool. Passed from the ControlUp Console.
    [string]$strHVMachineFarm ,

    # Type of machine. Passed from ControlUp Console.
    # Name of the VMware Horizon connection server. Passed from the ControlUp Console.
    [string]$strHVConnectionServerFQDN ,

    ## So that we can use same script after this declaration for enabling or disabling, by changing this value.
    ## Can't pass as a non-defaulted parameter otherwise would not work as an automated action

    ##             VVVVVVVV
    [bool]$enable = $false
    ##             ^^^^^^^^
)

## Map bool to stub for disable or enable messages
[array]$operationStubs = @( 'Disabl' , 'Enabl' )

#region ControlUp_Standards
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 250
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    try
    {
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions
    }
    catch
    {
        ## Nothing we can do but shouldn't cause script to end
    }
}
#endregion ControlUp_Standards

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
        Out-CUConsole -Message "The required PSCredential object could not be loaded. Please make sure you have run the 'Create credentials for VMware Horizon scripts' script on the target machine." -Exception $_
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
            $null = Import-Module -Name VMware.$component -Verbose:$false
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
            HelpMessage = "The FQDN of the VMware Horizon Connection server. IP address may be used.")]
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
            Out-CUConsole -Message 'There was a problem connecting to the VMware Horizon Connection server. It looks like there may be a certificate issue. Please ensure the certificate used on the VMware Horizon server is trusted by the machine running this script.' -Exception $_
        }
        else {
            Out-CUConsole -Message 'There was a problem connecting to the VMware Horizon Connection server.' -Exception $_
        }
    }
}

function Disconnect-HorizonConnectionServer {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The VMware Horizon Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )

    try {
        Disconnect-HVServer -Server $HVConnectionServer -Confirm:$false
    }
    catch {
        Out-CUConsole -Message 'There was a problem disconnecting from the VMware Horizon Connection server. If not running in a persistent session (ControlUp scripts do not run in a persistant session) this is not a problem, the session will eventually be deleted by VMware Horizon.' -Warning
    }
}

function Get-HVRDSFarm {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "Displayname of the Desktop Farm.")]
        [string]$HVFarmName,
        [parameter(Mandatory = $true,
            HelpMessage = "The VMware Horizon Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    # Try to get the Desktop pools in this pod
    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'FarmSummaryView'
        # Filter on the correct displayname
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName' = 'data.displayName'; 'value' = "$HVFarmName" }
        # Perform the actual query
        [array]$queryResults = ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        if (!$queryResults) {
            Out-CUConsole -message "Can't find $HVFarmName, exiting."
            exit
        }
        else {
            return $queryResults
        }
    }
    catch {
        Out-CUConsole -message 'There was a problem retreiving the VMware Horizon RDS Farm.' -Exception $_
    }
}

function Get-HVRDSMachine {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "ID of the RDS Farm.")]
        [VMware.Hv.FarmId]$hvfarmid,
        [parameter(Mandatory = $true,
            HelpMessage = "Name of the RDS machine.")]
        [string]$HVMachineName,
        [parameter(Mandatory = $true,
            HelpMessage = "The VMware Horizon Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )

    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'RDSServerSummaryView'
        # Filter so we get the correct machine in the correct pool
        $poolfilter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName' = 'base.farm'; 'value' = $hvfarmid }
        $machinefilter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName' = 'base.name'; 'value' = "$HVMachineName" }
        $filterlist = @()
        $filterlist += $poolfilter
        $filterlist += $machinefilter
        $filterAnd = New-Object VMware.Hv.QueryFilterAnd
        $filterAnd.Filters = $filterlist
        $defn.Filter = $filterAnd
        # Perform the actual query
        [array]$queryResults = ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        if (!$queryResults) {
            Out-CUConsole -message "Can't find $HVMachineName, exiting."
            exit
        }
        else {
            return $queryResults
        }
    }
    catch {
        Out-CUConsole -message 'There was a problem retreiving the VMware Horizon Desktop Pool.' -Exception $_
    }
}

function Set-HVRDSMachine {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "RDS Server object.")]
        [VMware.Hv.RDSServerSummaryView]$hvRDSServer,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer,
        [parameter(Mandatory = $true,
            HelpMessage = "True for Enable, False for Disable.")]
        [bool]$Enable ,[parameter(Mandatory = $true,
            HelpMessage = "Operation string for messages.")]
        [string]$operation
    )

    try {
        $hvRDSServerId = $hvRDSServer.id
        $hvRDSServername = $hvRDSServer.base.name
        # First define the Service we need
        [VMware.Hv.rdsserverservice]$rdsserverservice = new-object vmware.hv.rdsserverservice
        # Fill the helper for this service with the application information
        $rdsserverservicehelper = $rdsserverservice.read($HVConnectionServer.extensionData, $hvRDSServerId)
        if( $rdsserverservicehelper.getSettingsHelper().getEnabled() -eq $enable ) {
            Out-CUConsole -Message "RDS server $hvRDSServername is already $($operation)ed." -Warning
        }
        else {
            # Change the state of the application in the helper
            $rdsserverservicehelper.getSettingsHelper().setEnabled($enable)
            # Apply the helper to the actual object
            $rdsserverservice.update($HVConnectionServer.extensionData, $rdsserverservicehelper)
            Out-CUConsole -Message "Successfully $($operation)ed $hvRDSServername."
        }
    }
    catch {
        Out-CUConsole -Message "There was a problem $($operatin)ing $hvRDSServername." -Exception $_
    }
}

# Test arguments
if( [string]::IsNullOrEmpty( $strHVMachineName ) -or [string]::IsNullOrEmpty( $strHVMachineFarm ) -or [string]::IsNullOrEmpty( $strHVConnectionServerFQDN ) ) {
    Out-CUConsole -Message 'The Console or Monitor may not be connected to the Horizon environment, please check this.' -Stop
}

## For reasons unknown, CEIP warnings still come from import-module even when redirected
[string]$vmwareFolder = Join-Path -Path ([environment]::GetFolderPath( [Environment+SpecialFolder]::ApplicationData )) -ChildPath 'VMware\PowerCLI'
[string]$powerCLISettingsFile = Join-Path -Path $vmwareFolder -ChildPath 'PowerCLI_Settings.xml'

if( -Not ( Test-Path -Path $vmwareFolder -PathType Container ) ) {
    $null = New-Item -Path $vmwareFolder -ItemType Directory -Force
}

if( -Not ( Test-Path -Path $powerCLISettingsFile -PathType Leaf ) ) {
    '<Settings><Setting Name="InvalidCertificateAction" Value="Ignore" /><Setting Name="ParticipateInCEIP" Value="False" /><Setting Name="DisplayObsoleteWarnings" Value="False" /></Settings>' | Set-Content -Path $powerCLISettingsFile
}
else {
    $powerCLISettingsFile = $null ## flag for later that we don't delete it
}

try {
    # Import the VMware PowerCLI modules
    Load-VMwareModules -Components @('VimAutomation.HorizonView')

    # Get the stored credentials for running the script
    Out-CUConsole -Message "Loading Credentials."
    [PSCredential]$CredsHorizon = Get-CUStoredCredential -System 'HorizonView'

    # Connect to the VMware Horizon Connection Server
    Out-CUConsole -Message "Connecting to $strHVConnectionServerFQDN"
    [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = $null
    $objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $strHVConnectionServerFQDN -Credential $CredsHorizon
    if( $objHVConnectionServer ) {
        # get the pool
        Out-CUConsole -Message "Getting details for farm $strHVMachineFarm."
        $rds_farm = Get-HVRDSFarm -HVFarmName $strHVMachineFarm -HVConnectionServer $objHVConnectionServer
        $farmid = $rds_farm.id

        # Get the machine
        Out-CUConsole -Message "Getting details for $strHVMachineName."
        $machine = $null
        $machine = Get-HVRDSMachine -HVConnectionServer $objHVConnectionServer -HVMachineName $strHVMachineName -hvfarmid $farmid
        if( $machine ) {
            [string]$actionBase = $operationStubs[ ($enable -as [int] ) ]
            
            # Enable/Disable the RDS Server
            Out-CUConsole -Message "$($actionBase)ing $strHVMachineName."
            Set-HVRDSMachine -hvrdsserver $machine -HVConnectionServer $objHVConnectionServer -enable:$enable -operation $actionBase.ToLower()

            # Disconnect from the VMware Horizon Connection Center
            Out-CUConsole -Message "Disconnecting from $strHVConnectionServerFQDN"
            Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
        }
    }
}
catch {
    throw $_
}
finally {
    if( $powerCLISettingsFile -and ( Test-Path -Path $powerCLISettingsFile -PathType Leaf ) ) { ## will be null if we didn't create it and we can leave the folder if we created it
        Remove-Item -Path $powerCLISettingsFile
    }
}
