#requires -Version 3.0
$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Gets the administrative roles in a Horizon View environment

    .DESCRIPTION
    This script retreives the administrative users and groups in a Horizon View environment.

    .EXAMPLE
    You can this script to make sure administrators have the right permissions in Horizon View.
    
    .NOTES
    This script requires the VMWare PowerCLI module (minimum version 6.5.0R1) to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'
    Before running this script you will also need to have a PSCredential object available on the target machine. This can be created by running the 'Create credentials for Horizon View scripts' script in ControlUp on the target machine. For details, see LINKTOREPLACE
    
    Context: Can be triggered from the Horizon View Machines view
    Modification history: 02/01/2020 - Anthonie de Vreede - First version
   
    .PARAMETER strHVConnectionServerFQDN
    Name of the Horizon View connection server to be querried. Passed from the ControlUp Console.

    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli
    
    .COMPONENT
    VMWare PowerCLI 6.5.0R1
#>

# Name of the Horizon View connection server. Passed from the ControlUp Console.
[string]$strHVConnectionServerFQDN = $args[0]

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
            Write-Error -Message "$Message`n$($Exception.Exception.Message)`n$($Exception.CategoryInfo)`n$($Exception.Exception.ErrorRecord)" -ErrorAction Stop
        }
        Else {
            Write-Warning -Message "There was an unexpected error: $($Exception.CategoryInfo.Category)`nPlease see the Error tab for details."
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
            HelpMessage = 'The system the credentials will be used for.')]
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
            HelpMessage = 'The FQDN of the Horizon View Connection server. IP address may be used.')]
        [string]$HVConnectionServerFQDN,
        [parameter(Mandatory = $true,
            HelpMessage = 'The PSCredential object used for authentication.')]
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
            HelpMessage = 'The Horizon View Connection server object.')]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    
    try {
        Disconnect-HVServer -Server $HVConnectionServer -Confirm:$false
    }
    catch {
        Out-CUConsole -Message 'There was a problem disconnecting from the Horizon View Connection server. If not running in a persistent session (ControlUp scripts do not run in a persistant session) this is not a problem, the session will eventually be deleted by Horizon View.' -Warning
    }
}

# Test arguments
Test-ArgsCount -ArgsCount 1 -Reason 'The Console or Monitor may not be connected to the Horizon View environment, please check this.'

# Import the VMware PowerCLI modules
Load-VMWareModules -Components @('VimAutomation.HorizonView')

# Get the stored credentials for running the script
[PSCredential]$CredsHorizon = Get-CUStoredCredential -System 'HorizonView'

# Connect to the Horizon View Connection Server
[VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $strHVConnectionServerFQDN -Credential $CredsHorizon

# Create Horizon View services object
$objHVServices = $objHVConnectionServer.ExtensionData

# Get the list of permissions from the service
try {
    [array]$arrPermissionList = $objHVServices.Permission.Permission_List()


    # Get role names
    [hashtable]$hshRoles = @{ }
    Foreach ($Role in $arrPermissionList.Base.Role) {
        [string]$strRoleID = $Role.Id
        If (!($hshRoles.ContainsKey($strRoleID))) {
            $hshRoles[$strRoleID] = ($objHVServices.Role.Role_Get($Role)).Base.Name
        }
    }

    # Get access groups names
    [hashtable]$hshAccessGroups = @{ }
    Foreach ($AccessGroup in $arrPermissionList.Base.AccessGroup) {
        [string]$strAccessGroupID = $AccessGroup.Id
        If (!($hshAccessGroups.ContainsKey($strAccessGroupID))) {
            $hshAccessGroups[$strAccessGroupID] = ($objHVServices.AccessGroup.AccessGroup_Get($AccessGroup)).base.name
        }
    }

    # Get user and AD group names
    [hashtable]$hshAdminUserOrGroups = @{ }
    Foreach ($AdminUserOrGroup in $arrPermissionList.Base.UserOrGroup) {
        [string]$strAdminUserOrGroupID = $AdminUserOrGroup.Id
        If (!($hshAdminUserOrGroups.ContainsKey($strAdminUserOrGroupID))) {
            $hshAdminUserOrGroups[$strAdminUserOrGroupID] = ($objHVServices.AdminUserOrGroup.AdminUserOrGroup_Get($AdminUserOrGroup)).base.displayname
        }
    }
}
catch {
    Out-CUConsole -Message "There was an issue using the Horizon View Services object. Please see the exception for details." -Exception $_
}

# Put it all together and create the output
[array]$objOutput = @()
Foreach ($Permission in $arrPermissionList) {
    $objOutput += (New-Object -TypeName PSObject -Property @{
            'User or Group'     = $hshAdminUserOrGroups[$Permission.Base.UserOrGroup.Id]
            'Role; AccessGroup' = "$($hshRoles[$Permission.Base.Role.Id]); $($hshAccessGroups[$Permission.Base.AccessGroup.Id])"
        } ) 
}

# We are done with the Horizon Connection server, stop the connection
Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer

# Output the results
$objOutput | Sort-Object 'User or Group', 'Role; AccessGroup' | Select-Object 'User or Group', 'Role; AccessGroup'
