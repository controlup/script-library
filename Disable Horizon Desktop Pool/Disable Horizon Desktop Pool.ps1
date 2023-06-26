#requires -Version 3.0

<#
    .SYNOPSIS
    Enabled or Disables Horizon Virtual Desktop pool

    .DESCRIPTION
    This script disables or enables Horizon Virtual Desktop pool using Horizon SOAP API via PowerCLI

    .EXAMPLE
    Can be used as an Automated action to disable or enable Horizon Virtual Desktop pool.

    .NOTES
    This script requires VMWare PowerCLI  module to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'

    Before running this script you will also need to have a PSCredential object available on the target machine. This can be created by running the 'Create credentials for Horizon scripts' script in ControlUp on the target machine.

    Context: Can be triggered from the Horizon Machines view
    Modification history:   13/08/2019 - Anthonie de Vreede - First version
                            06/04/2023 - Wouter Kursten - Removed VMware.hv.helper dependency
                            12/05/2023 - Guy Leech - Brought into single script for enable & disable
                            16/05/2023 - Guy Leech - Added code to hide VMware CEIP warnings temporarily
                            23/05/2023 - Guy Leech - Merged code from pool provisioning script

    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli

    .COMPONENT
    VMWare PowerCLI 6.5.0R1 or higher
#>

[CmdletBinding()]

Param
(
    # Name of the Horizon Virtual Desktop pool.
    [string]$strHVPoolName ,

    # Name of the Horizon connection server. Passed from the ControlUp Console.
    [string]$strHVConnectionServerFQDN ,

    ## +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
    ##
    ## CHANGE THE DEFAULT VALUES BELOW DEPENDING ON WHETHER SCRIPT IS ENABLE OR DISABLE VARIANT AND FOR THE POOL ITSELF OR PROVISIONINING IN THAT POOL
    ##
    ## +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

    ## Can't pass the following parameters as non-defaulted parameter otherwise would not work as an automated action
    
    ## So that we can use same script after this param block for enabling or disabling, by changing the default value for $enable
    [bool]$enable = $false ,
    
    ## So that we can use same script after this param block for enabling/disabling pools or provisioning, by changing the default value for $doProvisioning
    [bool]$doProvisioning = $false
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
    $strCUCredFolder = "$([environment]::GetFolderPath( [Environment+SpecialFolder]::CommonApplicationData ))\ControlUp\ScriptSupport"
    try {
        Import-Clixml -LiteralPath (Join-Path -Path $strCUCredFolder -ChildPath "$($env:USERNAME)_$($System)_Cred.xml")
    }
    catch {
        Out-CUConsole -Message "The required PSCredential object could not be loaded fron `"$strCUCredFolder`". Please make sure you have run the 'Create credentials for Horizon scripts' script on the target machine." -Exception $_
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
            $warnings = $null
            Import-Module -Name "VMware.$component" -Verbose:$false -WarningVariable warnings 3>$null >$null
        }
        catch {
            try {
                $null = Add-PSSnapin -Name 'VMware'
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
        Out-CUConsole -Message 'There was a problem disconnecting from the Horizon Connection server. If not running in a persistent session (ControlUp scripts do not run in a persistent session) this is not a problem, the session will eventually be deleted by Horizon.' -Warning
    }
}

function Get-HVDesktopPool {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "Displayname of the Desktop Pool.")]
        [string]$HVPoolName,
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
        $defn.queryEntityType = 'DesktopSummaryView'
        # Filter on the correct displayname
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{ 'memberName' = 'desktopSummaryData.displayName' ; 'value' = $HVPoolname }
        # Perform the actual query
        [array]$queryResults = ($queryService.queryService_create( $HVConnectionServer.extensionData , $defn )).results
        # Remove the query
        $queryService.QueryService_DeleteAll( $HVConnectionServer.extensionData )
        # Return the results
        if (!$queryResults) {
            Out-CUConsole -Message "Can't find $HVPoolName, exiting." -Stop
            exit
        }
        else {
            return $queryResults
        }
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the VMware Horizon Desktop Pool.' -Exception $_
        exit
    }
}

function Set-PoolEnablement {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon Pool object.")]
        [object]$HVPool,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer,
        [parameter(Mandatory = $true,
            HelpMessage = "True for Enable, False for Disable.")]
        [bool]$Enable ,
        [parameter(Mandatory = $true,
            HelpMessage = "Text for operation for messages.")]
        [string]$Operation ,
        [string]$ProvisioningText 
    )
    
    try {
        $hvpoolid = $HVPool.id
        $poolname = $HVPool.DesktopSummaryData.name
        # First define the Service we need
        [VMware.Hv.desktopService]$desktopservice = new-object vmware.hv.desktopService
        # Fill the helper for this service with the Desktop information
        $desktophelper = $desktopservice.read($HVConnectionServer.extensionData, $hvpoolid)
        [bool]$doUpdate = $false
        [bool]$existingSetting = $false

        if( -Not [string]::IsNullOrEmpty( $ProvisioningText ) ) {         
            $existingSetting = $desktophelper.getAutomatedDesktopDataHelper().getVirtualCenterProvisioningSettingsHelper().getEnableProvisioning()
        }
        else {
            $existingSetting = $desktophelper.getDesktopSettingsHelper().getEnabled()
        }

        if( $existingSetting -eq $Enable ) {
            Out-CUConsole -Message "Desktop Pool `"$poolname`"$ProvisioningText is already $($Operation)ed." -Warning
        }
        elseif( -Not [string]::IsNullOrEmpty( $ProvisioningText )) {
            # Change the state of provisioning in the helper
            $desktophelper.getAutomatedDesktopDataHelper().getVirtualCenterProvisioningSettingsHelper().setEnableProvisioning($Enable)
            $doUpdate = $true
        }
        else {
            # Change the state of the Desktop in the helper
            $desktophelper.getDesktopSettingsHelper().setEnabled($Enable)
            $doUpdate = $true
        }
        
        if( $doUpdate ) {
            # Apply the helper to the actual object
            $desktopservice.update($HVConnectionServer.extensionData, $desktophelper)
            Out-CUConsole -Message "Successfully $($Operation)ed Desktop Pool$ProvisioningText `"$poolname`"."
        }
    }
    catch {
        Out-CUConsole -Message "There was a problem $($Operation)ing Desktop pool `"$poolname`"$provisioningText." -Exception $_
    }
}

# Test arguments
if( [string]::IsNullOrEmpty( $strHVPoolName ) -or [string]::IsNullOrEmpty( $strHVConnectionServerFQDN ) ) {
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

    ## Stop long warning message about participating in Customer Experience Improvement Program 
    $null = Set-PowerCLIConfiguration -Scope Session -ParticipateInCEIP $false -Confirm:$false -DisplayDeprecationWarnings $false

    # Connect to the Horizon Connection Server
    Out-CUConsole -Message "Connecting to $strHVConnectionServerFQDN"
    [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = $null
    Write-Verbose -Message "Connecting to $strHVConnectionServerFQDN as user $($CredsHorizon.Username), local user $env:USERNAME"
    $objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $strHVConnectionServerFQDN -Credential $CredsHorizon

    if( -Not $objHVConnectionServer ) {
        Out-CUConsole -Message "Failed to connect to Horizon Connection Server $strHVConnectionServerFQDN as user $($CredsHorizon.Username)" -Stop
    }
    
    [string]$provisioningText = $null
    if( $doProvisioning ) {
        $provisioningText = ' provisioning'
    }

    # Get the Horizon Virtual Desktop Pool
    Out-CUConsole -Message "Getting details for Desktop Pool `"$strHvPoolName`""
    [object]$objHVPool = Get-HVDesktopPool -HVPoolName $strHvPoolName -HVConnectionServer $objHVConnectionServer

    [string]$actionBase = $operationStubs[ ($enable -as [int] ) ]

    if( $objHVPool ) {
        # Enable Horizon Virtual Desktop Pool provisioning
        Out-CUConsole -Message "$($actionBase)ing Desktop Pool `"$strHvPoolName`""
        Set-PoolEnablement -HVPool $objHVPool -HVConnectionServer $objHvConnectionServer -Enable:($actionBase -ieq 'enabl') -Operation $actionBase.ToLower() -provisioningText $provisioningText
    }

    # Disconnect from the Horizon Connection Center
    Out-CUConsole -Message "Disconnecting from $strHVConnectionServerFQDN"
    Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
}
catch
{
    throw $_
}
finally
{
    if( $powerCLISettingsFile -and ( Test-Path -Path $powerCLISettingsFile -PathType Leaf ) ) { ## will be null if we didn't create it and we can leave the folder if we created it
        Remove-Item -Path $powerCLISettingsFile
    }
}
