$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Sends messages to users

    .DESCRIPTION
    This script can be used to send messages to a single user session using the Horizon SOAP API's. It can also be used as an automated action with a fixed message and severitylevel

    .NOTES
    For the severity level these are allowed: INFO,WARNING, ERROR

    This script requires VMWare PowerCLI to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers' Or by using the 'Install VMware PowerCLI' script.
    If you get TLS/SSL errors use this command Set-PowerCLIConfiguration -InvalidCertificateAction ignore
        or Set-PowerCLIConfiguration -InvalidCertificateAction warn
    To get rid of the CEIP warning use Set-PowerCLIConfiguration -ParticipateInCeip $true 
        or Set-PowerCLIConfiguration -ParticipateInCeip $false
    Credentials can be set using the 'Create credentials for Horizon View scripts' script.

    Modification history:   07/07/2020 - Wouter Kursten - First version

    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli


    .COMPONENT
    VMWare PowerCLI

#>

# Name of the Horizon View connection server. Passed from the ControlUp Console.
[string]$HVConnectionServerFQDN = $args[0]
# username
[string]$HVUsername = $args[1]
# Name of the Desktop pool
[string]$HVDesktopPoolname = $args[2]
# The message as a string
[string]$HVMessage = $args[3]
# Message severity
[string]$HVMessageSeverity = $args[4]


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
        else {
            return $queryResults
        }
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the Horizon View Desktop Pool.'
    }
}

function Get-HVUser {
    param (
        [parameter(Mandatory = $true,
        HelpMessage = "Username")]
        [string]$UserName,
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
        $defn.queryentitytype='ADUserOrGroupSummaryView'
        # this builds the filter
        $split=$UserName.split('\')
        $domain=$split[0]
        $user=$split[1]
        $notgroupfilter = New-Object VMware.Hv.QueryFilterEquals -Property @{ 'memberName' = 'base.group'; 'value' = $false }
        $usernamefilter = New-Object VMware.Hv.QueryFilterEquals -Property @{ 'memberName' = 'base.loginName'; 'value' = $user }
        $domainfilter = New-Object VMware.Hv.QueryFilterEquals -Property @{ 'memberName' = 'base.domain'; 'value' = $domain }
        $filterarray=@()
        $filterarray+=$notgroupfilter
        $filterarray+=$usernamefilter
        $filterarray+=$domainfilter
        $filterAnd = New-Object VMware.Hv.QueryFilterAnd
        $filterAnd.Filters = $filterarray
        $defn.Filter = $filterAnd
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        if (!$queryResults){
            Out-CUConsole -message "Can't find $UserName, exiting." -Exception "$UserName not found"
            exit
        }
        else {
            return $queryResults
        }
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the user object.' -Exception $_
    }
}

function Get-HVSession {
    param (
        [parameter(Mandatory = $true,
        HelpMessage = "ID for the user object")]
        [VMware.Hv.UserOrGroupId]$UserID,
        [parameter(Mandatory = $true,
            HelpMessage = "ID of the Desktop Pool.")]
        [VMware.Hv.DesktopId]$PoolID,
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
        $defn.queryEntityType = 'SessionLocalSummaryView'
        # Filter oud rds desktop pools since they don't contain machines
        $poolfilter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='referenceData.desktop'; 'value' = $PoolID}
        $userfilter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='referenceData.user'; 'value' = $UserID}
        out-cuconsole -message $poolfilter
        out-cuconsole -message $userfilter
        $filterarray=@()
        $filterarray+=$poolfilter
        $filterarray+=$userfilter
        $filterAnd = New-Object VMware.Hv.QueryFilterAnd
        $filterAnd.Filters = $filterarray
        $defn.Filter = $filterAnd
        out-cuconsole $defn.filter
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        if (!$queryResults){
            Out-CUConsole -message "Can't find session, exiting." -Exception "Session not found"
            exit
        }
        else {
            return $queryResults
        }
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the Horizon View session.' -Exception $_
    }
}

function Send-HVMessage {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "ID for the session")]
        [VMware.Hv.SessionId]$HVSessionID,
        [parameter(Mandatory = $true,
        HelpMessage = "The actual message")]
        [string]$HVMessage,
        [parameter(Mandatory = $true,
        HelpMessage = "Severity level for the message")]
        [ValidateSet("WARNING","ERROR","INFO")]
        [string]$Messageseveritylevel,
        [parameter(Mandatory = $true,
        HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    # Try to get the Desktop pools in this pod
    try {
        $HVConnectionServer.ExtensionData.Session.Session_SendMessage($HVSessionID,$Messageseveritylevel,$HVMessage)
    }
    catch {
        Out-CUConsole -Message 'There was a problem sending the message.' -Exception $_
    }
}

# Test arguments
Test-ArgsCount -ArgsCount 5 -Reason 'The Console or Monitor may not be connected to the Horizon View environment, please check this.'

# Set the credentials location
[string]$strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"

# Import the VMware PowerCLI modules
Load-VMwareModules -Components @('VimAutomation.HorizonView')

# Get the stored credentials for running the script
[PSCredential]$CredsHorizon = Get-CUStoredCredential -System 'HorizonView'

# Connect to the Horizon View Connection Server

[VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $HVConnectionServerFQDN -Credential $CredsHorizon

# retreives the pool object
$HVPool=Get-HVDesktopPool -HVPoolName $HVDesktopPoolname -HVConnectionServer $objHVConnectionServer
# Retreives the user object
$HVUser=Get-HVUser -Username $hvusername -HVConnectionServer $objHVConnectionServer
# But we only need the ID for both
$HVPoolID=($HVPool).id
$HVUserID=($hvuser).id

# retreives the session object
$HVSession=Get-HVSession -UserID $HVUserID -PoolID $HVPoolID -HVConnectionServer $objHVConnectionServer

# Again we only need the ID
$HVSessionID=($HVSession).id

# Send the message
Send-HVMessage -Messageseveritylevel $HVMessageSeverity -HVSessionID $hvsessionID -HVMessage $HVMessage -HVConnectionServer $objHVConnectionServer

# Disconnect from the connection server
Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
