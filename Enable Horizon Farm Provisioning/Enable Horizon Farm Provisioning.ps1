#requires -Version 3

<#
    .SYNOPSIS
    Enables Horizon RDS Farm provisioning

    .DESCRIPTION
    This script enables Horizon RDS Farm provisioning through the VMware.Hv.Helper module

    .EXAMPLE
    Can be used as an Automated action to disable Horizon RDS Farm provisioning if a resource shortage is detected.

    .NOTES
    This script requires VMWare PowerCLI  module to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'

    Before running this script you will also need to have a PSCredential object available on the target machine. This can be created by running the 'Create credentials for Horizon scripts' script in ControlUp on the target machine.

    Context: Can be triggered from the Horizon Machines view
    Modification history: 06/04/2023 - Wouter Kursten - First Version
                          16/05/2023 - Guy Leech      - Brought to CU scripting standards & into single script

    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli

    .COMPONENT
    VMWare PowerCLI 6.5.0R1 or higher
#>

[CmdletBinding()]

Param
(
    # Name of the Horizon RDS Farm.
    [string]$HVRDSFarmname ,
    # Name of the Horizon connection server. Passed from the ControlUp Console.
    [string]$strHVConnectionServerFQDN ,
    
    ## So that we can use same script after this declaration for enabling or disabling, by changing this value.
    ## Can't pass as a non-defaulted parameter otherwise would not work as an automated action

    ##             VVVVVVVV
    [bool]$enable = $true
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
        Out-CUConsole -Message "The required PSCredential object could not be loaded. Please make sure you have run the 'Create credentials for Horizon scripts' script on the target machine." -Exception $_
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
            HelpMessage = "The FQDN of the Horizon Connection server. IP address may be used.")]
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
            Out-CUConsole -Message 'There was a problem connecting to the Horizon Connection server. It looks like there may be a certificate issue. Please ensure the certificate used on the Horizon server is trusted by the machine running this script.' -Exception $_
        }
        else {
            Out-CUConsole -Message 'There was a problem connecting to the Horizon Connection server.' -Exception $_
        }
    }
}

function Disconnect-HorizonConnectionServer {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )

    try {
        Disconnect-HVServer -Server $HVConnectionServer -Confirm:$false
    }
    catch {
        Out-CUConsole -Message 'There was a problem disconnecting from the Horizon Connection server. If not running in a persistent session (ControlUp scripts do not run in a persistant session) this is not a problem, the session will eventually be deleted by Horizon.' -Warning
    }
}

function Get-HVFarm {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "Name of the RDS Farm.")]
        [string]$HVFarmname,
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
        $defn.queryEntityType = 'FarmSummaryView'
        # Filter oud rds desktop pools since they don't contain machines
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName' = 'data.name'; 'value' = "$HVFarmname" }
        # Perform the actual query
        [array]$queryResults = ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        if (!$queryResults) {
            Out-CUConsole -message "Can't find $HVFarmname, exiting." -Exception "$HVFarmname not found"
            exit
        }
        else {
            return $queryResults
        }
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the Horizon View RDS Farm.' -Exception $_
    }
}

function Set-HVFarm {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "ID of the RDS Farm.")]
        [VMware.Hv.FarmSummaryView]$hvfarm,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer,
        [parameter(Mandatory = $true,
            HelpMessage = "True for Enable, False for Disable.")]
        [bool]$Enable ,
        [parameter(Mandatory = $true,
            HelpMessage = "Text for operation for messages.")]
        [string]$operation
    )
    if ($HVFarm.data.Type -ieq "MANUAL") {
        Out-CUConsole -Message 'Could not execute! This a manual Horizon RDS farm, cannot change the number of RDS hosts' -warning $_
        exit
    }
    try {
        $hvfarmid = $hvfarm.id
        $hvfarmname = $hvfarm.data.name
        # First define the Service we need
        [VMware.Hv.FarmService]$farmservice = New-Object -Typename vmware.hv.FarmService
        # Fill the helper for this service with the application information
        $farmhelper = $farmservice.read( $HVConnectionServer.extensionData , $HVFarmID )
        if( $farmhelper.getAutomatedFarmDataHelper().getVirtualCenterProvisioningSettingsHelper().getEnableProvisioning() -eq $enable ) {
            Out-CUConsole -Message "Farm $hvfarmname is already $($operation)ed." -Warning
        }
        else {
            # Change the state of the application in the helper
            $farmhelper.getAutomatedFarmDataHelper().getVirtualCenterProvisioningSettingsHelper().setEnableProvisioning( $enable )
            # Apply the helper to the actual object
            $farmservice.update( $HVConnectionServer.extensionData , $farmhelper )
            Out-CUConsole -Message "Successfully $($operation)ed the provisioning state for farm $hvfarmname."
        }
    }
    catch {
        Out-CUConsole -Message "There was a problem $($operation)ing the provisioning state for farm $hvfarmname." -Exception $_
    }
}

# Test arguments
if( [string]::IsNullOrEmpty( $HVRDSFarmname ) -or [string]::IsNullOrEmpty( $strHVConnectionServerFQDN ) ) {
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
    Out-CUConsole -Message "Loading Powershell module(s)"
    Load-VMwareModules -Components @('VimAutomation.HorizonView')

    # Get the stored credentials for running the script
    Out-CUConsole -Message "Loading credentials"
    [PSCredential]$CredsHorizon = Get-CUStoredCredential -System 'HorizonView'

    # Connect to the Horizon Connection Server
    Out-CUConsole -Message "Connecting to $strHVConnectionServerFQDN"
    [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = $null
    $objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $strHVConnectionServerFQDN -Credential $CredsHorizon

    if( $objHVConnectionServer ) {
        # Retreive the RDS Farm
        Out-CUConsole -Message "Getting details for farm $HVRDSFarmname"
        $HVFarm = $null
        $HVFarm = Get-HVFarm -HVFarmname $HVRDSFarmname -HVConnectionServer $objHVConnectionServer

        if( $HVFarm ) {         
            [string]$actionBase = $operationStubs[ ($enable -as [int] ) ]

            # Enable or Disable Horizon Farm provisioning
            Out-CUConsole -Message "$($actionbase)ing provisioning for Farm $HVRDSFarmname"
            Set-HVFarm -HVConnectionServer $objHVConnectionServer -HVFarm $HVFarm -enable:$enable -Operation $actionBase.ToLower()

            # Disconnect from the connection server
            Out-CUConsole -Message "Disconnecting from $strHVConnectionServerFQDN"
            Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
        }
    }
}
catch {
    throw $_
}
finally{
    if( $powerCLISettingsFile -and ( Test-Path -Path $powerCLISettingsFile -PathType Leaf ) ) { ## will be null if we didn't create it and we can leave the folder if we created it
        Remove-Item -Path $powerCLISettingsFile
    }
}
