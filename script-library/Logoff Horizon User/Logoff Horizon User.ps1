#requires -Version 3.0
$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Logs a user of with Horizon

    .DESCRIPTION
    This script finds a user session based on username, machine and pool. This user is then logged off, even if the session is locked.

    .EXAMPLE
    You can automate this script to trigger on sessions that have been in a Disconnected state too long, so they will be logged off.
    
    .NOTES
    This script requires the VMWare PowerCLI (minimum 6.5R1) module to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'

    Before running this script you will also need to have a PSCredential object available on the target machine. This can be created by running the 'Create credentials for Horizon View scripts' script in ControlUp on the target machine.

    Context: Can be triggered from the Horizon Sessions and Machines view
    Modification history: 11/01/2020 - Anthonie de Vreede - 1.0

    .PARAMETER strHVUserName
    Username be logged off (domain\username). Passed from the ControlUp Console.

    .PARAMETER strHVMachineName
    DNS Name of the machine the session is on. Passed from the ControlUp Console.

    .PARAMETER strHVConnectionServerFQDN
    FQDN of the Primary Connection Server. This is a parameter to be filled in by the user.

    .PARAMETER strIgnoreCertificateError
    If set to 'True' the -Force parameter is used when connecting to the Horizon environment so certificate errors are ignored. Default is 'False'
    
    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli
    
    .COMPONENT
    VMWare PowerCLI 6.5.0R1 or higher
#>

# Name of the user to be logged off.  Drop the domain name (if it exists).
[string]$strHVUserName = $args[0].Split('\')[-1]
# Name of the Horizon View machine. Passed from the ControlUp Console. Drop the domain name (if it exists).
[string]$strHVMachineName = $args[1].Split('.')[0]
# Name of the Horizon View machine. Passed from the ControlUp Console.
[string]$strHVConnectionServerFQDN = $args[2]
# Ignore certificate errors for connecting the environment
[string]$strIgnoreCertificateError = $args[3]

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
        [string]$HorizonServerFQDN,
        [parameter(Mandatory = $true,
            HelpMessage = "The PSCredential object used for authentication.")]
        [PSCredential]$Credential
    )

    try {
        Connect-HVServer -Server $HorizonServerFQDN -Credential $Credential
    }
    catch {
        if ($_.Exception.Message.StartsWith('Could not establish trust relationship for the SSL/TLS secure channel with authority')) {
            Out-CUConsole -Message "There was a problem connecting to the Horizon View Connection server. It looks like there may be a certificate issue.`nYou can either set the Ignore Certificate Error parameter to True, or ensure the certificate used on the Horizon View server is trusted by the machine/account running this script." -Exception $_
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

function New-HorizonQueryDefinition {
    Param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Name of the type of object to be returned.')]
        [ValidateSet('ADUserOrGroupSummaryView', 'ApplicationIconInfo', 'ApplicationInfo', 'DesktopSummaryView', 'EntitledUserOrGroupGlobalSummaryView', 'EntitledUserOrGroupLocalSummaryView', 'EventSummaryView', `
                'FarmHealthInfo', 'FarmSummaryView', 'GlobalApplicationEntitlementInfo', 'GlobalEntitlementSummaryView', 'MachineNamesView', 'MachineSummaryView', 'PersistentDiskInfo', 'PodAssignmentInfo', 'RDSServerInfo', `
                'RDSServerSummaryView', 'RegisteredPhysicalMachineInfo', 'SessionGlobalSummaryView', 'SessionLocalSummaryView', 'TaskInfo', 'UserHomeSiteInfo')]
        [string]$queryEntityType,
        [Parameter(Mandatory = $false,
            HelpMessage = 'Filter of type VMware.Hv.Filter')]
        [Object]$Filter,
        [Parameter(Mandatory = $false,
            HelpMessage = 'Member names to sort by, if any.')]
        [string]$sortBy,
        [Parameter(Mandatory = $false,
            HelpMessage = 'Sort order, false (ascending) by default.')]
        [bool]$sortDescending,
        [Parameter(Mandatory = $false,
            HelpMessage = '0-based starting offset for returned results, defaults to 0.')]
        [int]$startingOffset,
        [Parameter(Mandatory = $false,
            HelpMessage = 'Maximum count of items this query should produce, defaults to all.')]
        [int]$limit,
        [Parameter(Mandatory = $false,
            HelpMessage = 'Maximum page size to return (the server may use a smaller size).')]
        [int]$maxPageSize
    )
    New-Object -TypeName Vmware.Hv.QueryDefinition -Property $PSBoundParameters
}

function Invoke-HorizonQuery {
    Param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'The Horizon View query')]
        $HVQueryService,
        [Parameter(Mandatory = $true,
            HelpMessage = 'The Horizon View Services object')]
        $HVServices,
        [Parameter(Mandatory = $true,
            HelpMessage = 'Horizon query definintion object.')]
        $QueryDefinition

    )
    # Run the query
    [pscustomobject]$Query = $HVQueryService.QueryService_Create($HVServices, $QueryDefinition)

    # Place result in PSCustomObject
    [pscustomobject]$Results = $Query.results

    # Delete the server side query
    $HVQueryService.QueryService_Delete($HVServices, $Query.id)

    # Output the object
    $Results
}

# Check enough arguments have been passed from the Console
Test-ArgsCount -ArgsCount 4 -Reason "There are several reasons this can happen:`nThe Console or Monitor may not be connected to the Horizon environment, please check this.`nYou may also have selected a machine without any users on it.`nIf started from the machine view the script runs on all sessions, including system and sessions that are not Horizon sessions."

# Import the VMware PowerCLI modules
Load-VMwareModules -Components @('VimAutomation.Core','VimAutomation.HorizonView')

# Set PowerCLI configuration
if ($strIgnoreCertificateError -eq 'True') {
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -ParticipateInCeip:$false -Scope Session -Confirm:$false | Out-Null
}
else {
    Set-PowerCLIConfiguration -InvalidCertificateAction Fail -ParticipateInCeip:$false -Scope Session -Confirm:$false | Out-Null
}

# Get the stored credentials for running the script
[PSCredential]$CredsHorizon = Get-CUStoredCredential -System 'HorizonView'

# Connect to the Horizon View Connection Server
[VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HorizonServerFQDN $strHVConnectionServerFQDN -Credential $CredsHorizon

# Create Horizon View services object
$objHVServices = $objHVConnectionServer.ExtensionData

# Create query service object
[System.Object]$objHVQueryService = New-Object -TypeName VMware.Hv.QueryServiceService

# Create filter on user, machine and pool for triple check
$objFilter = New-Object VMware.Hv.QueryFilterAnd -Property @{filters = 
    @(
        New-Object VMware.Hv.QueryFilterContains -Property @{memberName = 'namesData.userName'; value = $strHVUserName }
        New-Object VMware.Hv.QueryFilterStartsWith -Property @{memberName = 'namesData.machineOrRDSServerName'; value = $strHVMachineName }
    )
}

# Create the query
[System.Object]$objHVQuery = New-HorizonQueryDefinition -queryEntityType SessionLocalSummaryView -filter $objFilter

# Get the SessionID object(s)
$objSessions = Invoke-HorizonQuery -HVServices $objHVServices -HVQueryService $objHVQueryService -QueryDefinition $objHVQuery | Select-Object Id, NamesData

# Check the username and machine name are correct. (to avoid issues with username being part of another username, ditto for machine)
# Create array to hold selected Session IDs
[VMware.Hv.SessionId[]]$arrSessionIDs = @()
Foreach ($Session in $objSessions) {
    If (( ($Session.NamesData.UserName).Split('`\')[-1] -eq $strHVUserName) -and (($Session.NamesData.MachineOrRDSServerName).Split('.')[0] -eq $strHVMachineName)) {
        $arrSessionIDs += $Session.Id
    }
}

If ($arrSessionIDs.Count -eq 0) {
    Out-CUConsole -Message "The session was not found, either it is not a Horizon session or perhaps the session was already logged off." -Warning
    Exit 0
}
elseif ($arrSessionIDs.Count -gt 1) {
    Out-CUConsole -Message "More than one session matching username AND machine was found. To err on the side of caution, no sessions were logged off." -Stop
}

else {
    try {
        $objHVServices.Session.Session_Logoffforced($arrSessionIDs[0])
        Out-CUConsole -Message "`nLogoff command was issued for the user."
    }
    catch {
        Out-CUConsole -Message "The user could not be logged off." -Exception $_
    }
}

# Disconnect from the Horizon View
Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
