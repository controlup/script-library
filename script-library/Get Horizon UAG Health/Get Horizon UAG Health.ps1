$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Gets Unified Access Gateway(UAG) health information from all UAG in the Cloud Pod federation 

    .DESCRIPTION
    This script connects to all pod's in a Cloud Pod Architecture(CPA) or only the local one if CPA hasn't been initialized and pulls all health information for configured UAG's.

    .NOTES
    This script requires VMWare PowerCLI to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers' Or by using the 'Install VMware PowerCLI' script.
    Credentials can be set using the 'Prepare machine for Horizon View scripts' script.

    Some functions require Powershell 11.4 or higher and Horizon 7.10 or Higher

    Modification history:   01/09/2020 - Wouter Kursten - First version
                            23/09/2020 - Wouter Kursten - Second version

    Changelog
        23/09/2020  -   Minor fix

    .LINK
    https://code.vmware.com/web/tool/11.4.0/vmware-powercli

    .COMPONENT
    VMWare PowerCLI
#>
# Name of the Horizon View connection server. Passed from the ControlUp Console.
[string]$HVConnectionServerFQDN = $args[0]

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

function Get-HVUAGGatewayZone {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "Boolean for the GatewayZoneInternal property of the GatewayHealth.")]
        [bool]$GatewayZoneInternal
    )
    try {
        if ($GatewayZoneInternal -eq $False) {
            $GatewayZoneType="External"
        }
        elseif ($GatewayZoneInternal -eq $True) { 
            $GatewayZoneType="Internal"
        }
        # Return the results
        return $GatewayZoneType
    }
    catch {
        Out-CUConsole -Message 'There was a problem determining the gateway zone type.' -Exception $_
    }
}

function Get-HVGatewayType {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "String for the GatewayZoneInternal property of the GatewayHealth.")]
        [string]$HVGatewayType
    )
    try {
        if ($HVGatewayType -eq "AP") {
            $GatewayType="UAG"
        }
        elseif ($HVGatewayType -eq "F5") { 
            $GatewayType="F5 Load Balancer"
        }
        elseif ($HVGatewayType -eq "SG") { 
            $GatewayType="Security Server"
        }
        elseif ($HVGatewayType -eq "SG-cohosted") { 
            $GatewayType="Cohosted CS"
        }
        elseif ($HVGatewayType -eq "Unknown") { 
            $GatewayType="Unknownr"
        }
        # Return the results
        return $GatewayType
    }
    catch {
        Out-CUConsole -Message 'There was a problem determining the gateway type.' -Exception $_
    }
}

# Test arguments
Test-ArgsCount -ArgsCount 1 -Reason 'The Console or Monitor may not be connected to the Horizon View environment, please check this.'

# Set the credentials location
[string]$strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"

# Make Verbose messages green
$host.privatedata.VerboseForegroundColor="Green"

# Import the VMware PowerCLI modules
Load-VMwareModules -Components @('VimAutomation.HorizonView')

# Get the stored credentials for running the script
[PSCredential]$CredsHorizon = Get-CUStoredCredential -System 'HorizonView'

# Connect to the Horizon View Connection Server
[VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $HVConnectionServerFQDN -Credential $CredsHorizon

# checks if this connectionserver is member of a cloud pod federation

[VMware.Hv.PodFederationLocalPodStatus]$HVpodstatus=($objHVConnectionServer.ExtensionData.PodFederation.PodFederation_Get()).localpodstatus

if ($HVpodstatus.status -eq "ENABLED"){
    # Retreives all pods
    [array]$HVpods=$objHVConnectionServer.ExtensionData.Pod.Pod_List()
    # retreive the first connection server from each pod
    $HVPodendpoints=@()
    if ($localsiteonly -eq $true){
        $hvlocalpod=$hvpods | where-object {$_.LocalPod -eq $true}
        $hvlocalsite=$objHVConnectionServer.ExtensionData.Site.Site_Get($hvlocalpod.site)
        foreach ($hvpod in $hvlocalsite.pods){$HVPodendpoints+=$objHVConnectionServer.ExtensionData.PodEndpoint.PodEndpoint_list($hvpod) | select-object -first 1}
        }

    else {
            [array]$HVPodendpoints = foreach ($hvpod in $hvpods) {$objHVConnectionServer.ExtensionData.PodEndpoint.PodEndpoint_List($hvpod.id) | select-object -first 1}
    }
    # Convert from url to only the name
    [array]$hvconnectionservers=$HVPodendpoints.serveraddress.replace("https://","").replace(":8472/","")
    # Disconnect from the current connection server
    Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
}
else {
    # Create list with one entry
    $hvconnectionservers=$hvConnectionServerfqdn
    # Disconnect from the current connection server
    Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
}

$uaghealthstatuslist=@()

foreach ($hvconnectionserver in $hvconnectionservers){
    if ($HVpodstatus.status -eq "ENABLED"){
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $hvconnectionserver -Credential $CredsHorizon
        # Retreive the name of the pod
        [string]$podname=$objHVConnectionServer.extensionData.pod.Pod_list() | where-object {$_.localpod -eq $True} | select-object -expandproperty Displayname
        Out-CUConsole -message "Processing Pod $podname"
        [array]$uaglist=$objHVConnectionServer.extensiondata.Gateway.Gateway_List()
        foreach ($uag in $uaglist){
            [VMware.Hv.GatewayHealthInfo]$uaghealth=$objHVConnectionServer.extensiondata.GatewayHealth.GatewayHealth_Get($uag.id)
            $uaghealthstatuslist+=New-Object PSObject -Property @{
                "Podname" = $podname;
                "Gateway_Name" = $uaghealth.name;
                "Gateway_Address" = $uaghealth.name;
                "Gateway_GatewayZone" = (Get-HVUAGGatewayZone -GatewayZoneInternal ($uaghealth.GatewayZoneInternal));
                "Gateway_Version" = $uaghealth.Version;
                "Gateway_Type" = (Get-HVGatewayType -HVGatewayType ($uaghealth.type));
                "Gateway_Active" = $uaghealth.GatewayStatusActive;
                "Gateway_Stale" = $uaghealth.GatewayStatusStale;
                "Gateway_Contacted" = $uaghealth.GatewayContacted;
                "Gateway_Active_Connections" = $uaghealth.ConnectionData.NumActiveConnections;
                "Gateway_Blast_Connections" = $uaghealth.ConnectionData.NumBlastConnections;
                "Gateway_PCOIP_Connections" = $uaghealth.ConnectionData.NumPcoipConnections;
            }
        }
    }
    else {
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $hvconnectionserver -Credential $CredsHorizon
        # Get the Horizon View desktop Pools within the pod
        Out-CUConsole -message "Processing the local pod"
        [array]$uaglist=$objHVConnectionServer.extensiondata.Gateway.Gateway_List()
        $podname="No CPA"
        foreach ($uag in $uaglist){
            [VMware.Hv.GatewayHealthInfo]$uaghealth=$objHVConnectionServer.extensiondata.GatewayHealth.GatewayHealth_Get($uag.id)
            $uaghealthstatuslist+=New-Object PSObject -Property @{
                "Podname" = $podname;
                "Gateway_Name" = $uaghealth.name;
                "Gateway_Address" = $uaghealth.Address;
                "Gateway_GatewayZone" = (Get-HVUAGGatewayZone -GatewayZoneInternal ($uaghealth.GatewayZoneInternal));
                "Gateway_Version" = $uaghealth.Version;
                "Gateway_Type" = (Get-HVGatewayType -HVGatewayType ($uaghealth.type));
                "Gateway_Active" = $uaghealth.GatewayStatusActive;
                "Gateway_Stale" = $uaghealth.GatewayStatusStale;
                "Gateway_Contacted" = $uaghealth.GatewayContacted;
                "Gateway_Active_Connections" = $uaghealth.ConnectionData.NumActiveConnections;
                "Gateway_Blast_Connections" = $uaghealth.ConnectionData.NumBlastConnections;
                "Gateway_PCOIP_Connections" = $uaghealth.ConnectionData.NumPcoipConnections;
            }
        }
    }
    Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
}
Out-CUConsole -message ($uaghealthstatuslist | select-object Podname,Gateway_Name,Gateway_Address,Gateway_GatewayZone,Gateway_Version,Gateway_Type,Gateway_Active,Gateway_Stale,Gateway_Contacted,Gateway_Active_Connections,Gateway_Blast_Connections,Gateway_PCOIP_Connections | Out-String)
