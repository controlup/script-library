$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Pulls all admin events from the Horizon Event Database

    .DESCRIPTION
    Uses the Horizon PowerCLI api's to pull all admin related events from the Horizon Event database for all pods. If there is no cloud pod setup it will only process the local pod.
    After pulling the events it will translate the id's for the various objects to names to show the proper names where needed.

    .PARAMETER HVConnectionserverFQDN
    Passes as the Primary COnnection server object from a machine

    .PARAMETER daysback
    Amount of days to go back to gather the audit logs

    .PARAMETER csvlocation
    Path to where the output csv file will be stored ending in \.

    .PARAMETER csvfilename
    Name of the csv file where the output will be stored.


    .NOTES
    This script is based on the work done here: https://www.retouw.nl/powercli/new-view-api-query-services-in-powercli-10-1-1-pulling-event-information-without-the-sql-password/
    
    This script requires VMWare PowerCLI to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers' Or by using the 'Install VMware PowerCLI' script.
    Credentials can be set using the 'Prepare machine for Horizon View scripts' script.

    This script require Powershell 11.4 or higher and Horizon 7.5 or Higher

    Modification history:   06/05/2020 - Wouter Kursten - First version

    .LINK
    https://code.vmware.com/web/tool/11.4.0/vmware-powercli
    https://www.retouw.nl/powercli/new-view-api-query-services-in-powercli-10-1-1-pulling-event-information-without-the-sql-password/

    .COMPONENT
    VMWare PowerCLI
#>
# Name of the Horizon View connection server. Passed from the ControlUp Console.
[string]$HVConnectionServerFQDN = $args[0]

# Amount of days to go back for the logs
[int]$daysback = $args[1]

# output location
[string]$csvlocation = $args[2]

# output filename
[string]$csvfilename = $args[3]

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

function get-hvadminevents {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "Start date.")]
        $startDate,
        [parameter(Mandatory = $true,
            HelpMessage = "End Date.")]
        $endDate,
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
        $defn.queryEntityType = 'EventSummaryView'
        # Filter on just the user and the vlsi module
        $modulefilter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='data.module'; 'value' = "Vlsi"}
        $timeFilter = new-object VMware.Hv.QueryFilterBetween -property @{'memberName'='data.time'; 'fromValue' = $startDate; 'toValue' = $endDate}
        $filterlist = @()
        $filterlist += $modulefilter
        $filterlist += $timeFilter
        $filterAnd = New-Object VMware.Hv.QueryFilterAnd
        $filterAnd.Filters = $filterlist
        $defn.Filter = $filterAnd

        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        return $queryResults
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving event data from the Horizon View Connection server.' -Warning
    }
}


# Test arguments
Test-ArgsCount -ArgsCount 4 -Reason 'The Console or Monitor may not be connected to the Horizon View environment, please check this.'

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
    [array]$HVPodendpoints = foreach ($hvpod in $hvpods) {$objHVConnectionServer.ExtensionData.PodEndpoint.PodEndpoint_List($hvpod.id) | select-object -first 1}

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

$date = get-date

$sinceDate = (get-date).AddDays(-$daysback)

$auditlog = @()


$auditlog



foreach ($hvconnectionserver in $hvconnectionservers){
    if ($HVpodstatus.status -eq "ENABLED"){
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $hvconnectionserver -Credential $CredsHorizon
        # Retreive the name of the pod
        [string]$podname=$objHVConnectionServer.extensionData.pod.Pod_list() | where-object {$_.localpod -eq $True} | select-object -expandproperty Displayname
        Out-CUConsole -message "Processing Pod $podname"
        $events= get-hvadminevents -HVConnectionServer $objHVConnectionServer -startDate $sinceDate -endDate $date

        foreach ($event in $events){
            if($event.data.RDSServerid){
                try{
                    $rdsserver=$objHVConnectionServer.ExtensionData.rdsserver.rdsserver_Get($event.data.rdsserverid).base
                    $rdsservername = $rdsserver.Name
                }
                catch{
                    $rdsservername = "RDS Server not found"
                    $farmname = "RDS Farm could not be found because the RDS server could not be resolved"
                }
            if($rdsserver){
                try{
                    $farmname=($objHVConnectionServer.ExtensionData.farm.farm_get($rdsserver.farm)).data.displayname
                }
                catch{
                    $farmname = "RDS Farm not found"
                }
        }
            }
            else{
                $rdsservername = $null
                $farmname = $event.namesdata.farmdisplayname
            }

            $auditlog+=New-Object PSObject -Property @{"Event Type" = $event.data.eventType;
                "Pod" = $podname;
                "Severity" = $event.data.severity;
                "Time" = $event.data.time;
                "Message" = $event.data.message;
                "Node" = $event.data.node;
                "User" = $event.namesdata.userdisplayname;
                "Machinename" = $event.namesdata.Machinename;
                "RDSServer" = $rdsservername;
                "RDSFarm" = $farmname;
                "Poolname" = $event.namesdata.DesktopDisplayName;
                "Application" = $event.namesdata.ApplicationDisplayName
                "Persistentdisk" = $event.namesdata.PersistentdiskName
            }
        }
    }

    else {
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $hvconnectionserver -Credential $CredsHorizon
        # Get the Horizon View desktop Pools within the pod
        Out-CUConsole -message "Processing the local pod"
        [string]$podname="Local"
        Out-CUConsole -message "Processing Pod $podname"
        $events= get-hvadminevents -HVConnectionServer $objHVConnectionServer -startDate $sinceDate -endDate $date

        foreach ($event in $events){
            if($event.data.RDSServerid){
                try{
                    $rdsserver=$objHVConnectionServer.ExtensionData.rdsserver.rdsserver_Get($event.data.rdsserverid).base
                    $rdsservername = $rdsserver.Name
                }
                catch{
                    $rdsservername = "RDS Server not found"
                    $farmname = "RDS Farm could not be found because the RDS server could not be resolved"
                }
            if($rdsserver){
                try{
                    $farmname=($objHVConnectionServer.ExtensionData.farm.farm_get($rdsserver.farm)).data.displayname
                }
                catch{
                    $farmname = "RDS Farm not found"
                }
        }
            }
            else{
                $rdsservername = $null
                $farmname = $event.namesdata.farmdisplayname
            }

            $auditlog+=New-Object PSObject -Property @{"Event Type" = $event.data.eventType;
                "Pod" = $podname;
                "Severity" = $event.data.severity;
                "Time" = $event.data.time;
                "Message" = $event.data.message;
                "Node" = $event.data.node;
                "User" = $event.namesdata.userdisplayname;
                "Machinename" = $event.namesdata.Machinename;
                "RDSServer" = $rdsservername;
                "RDSFarm" = $farmname;
                "Poolname" = $event.namesdata.DesktopDisplayName;
                "Application" = $event.namesdata.ApplicationDisplayName
                "Persistentdisk" = $event.namesdata.PersistentdiskName
            }
        }
    }
    Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
    #Out-CUConsole -message ($uaghealthstatuslist | select-object Podname,Gateway_Name,Gateway_Address,Gateway_GatewayZone,Gateway_Version,Gateway_Type,Gateway_Active,Gateway_Stale,Gateway_Contacted,Gateway_Active_Connections,Gateway_Blast_Connections,Gateway_PCOIP_Connections | Out-String)
}
$csvpath = $csvlocation+$csvfilename
$auditlog | sort-object -property Pod,Time | ConvertTo-Csv -NoTypeInformation | Out-File $csvpath -Encoding utf8
write-output $auditlog | select-object -property Time,Node,message | sort-object -property Pod,Time | format-table * -autosize -wrap
