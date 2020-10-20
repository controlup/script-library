<#
    .SYNOPSIS
    Shows information about snapshot status for VMware Horizon Instant and Linked Clone desktop pools and farms.

    .DESCRIPTION
    Uses the Horizon PowerCLI api's to pull all snapshot information for Horizon Linked Clones and Instant Clones Desktops pools and RDS farms. The script uses this information to find VDI machines
    and RDS hosts that are not running on the same Golden Image and Snapshot that are configured in the Desktop Pool settings.

    The script also uses the api's to poll the Cloud Pod status of the system and connects to other pods if Cloud Pod has been enabled.

    .PARAMETER HVConnectionserverFQDN
    Passes as the Primary COnnection server object from a machine

    .NOTES
    This script is based on the work done here: https://www.retouw.nl/2020/09/19/horizonapifinding-vdi-or-rds-machines-based-on-wrong-old-golden-image/

    This script requires VMWare PowerCLI to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers' Or by using the 'Install VMware PowerCLI' script.
    Credentials can be set using the 'Prepare machine for Horizon View scripts' script.

    This script require Powershell 11.4 or higher and Horizon 7.5 or Higher

    Modification history:   22/95/2020 - Wouter Kursten - First version

    .LINK
    https://code.vmware.com/web/tool/11.4.0/vmware-powercli
    https://www.retouw.nl/powercli/new-view-api-query-services-in-powercli-10-1-1-pulling-event-information-without-the-sql-password/

    .COMPONENT
    VMWare PowerCLI
#>
# Name of the Horizon View connection server. Passed from the ControlUp Console.

## GRL this way allows script to be run with debug/verbose without changing script
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[string]$HVConnectionServerFQDN = $args[0]
[int]$outputWidth = 400

# Altering the size of the PS Buffer
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ($WideDimensions = $PSWindow.BufferSize) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

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

function Get-HVDesktopPools {
    param (
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
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='desktopSummaryData.type'; 'value' = "AUTOMATED"}
        # Filter oud rds desktop pools since they don't contain machines
        
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        $queryResults = foreach ($queryResult in $queryResults){$HVConnectionServer.extensionData.desktop.desktop_get($queryResult.id) }
        $queryResults=$queryResults |  where-object {$_.automateddesktopdata.provisioningtype -ne "VIRTUAL_CENTER"}
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        return $queryResults
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the Horizon View Desktop Pool(s).' -Exception $_
    }
}

function Get-HVFarms {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'FarmSummaryView'
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='data.type'; 'value' = "AUTOMATED"}
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        $queryResults = foreach ($queryResult in $queryResults){$HVConnectionServer.extensionData.farm.farm_get($queryResult.id)}
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        return $queryResults
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the Horizon View RDS Farm(s).' -Exception $_
    }
}

function Get-HVDesktopMachines {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "ID of the Horizon View Desktop Pool.")]
        [VMware.Hv.DesktopId]$HVPoolID,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'MachineSummaryView'
        # Filter for only the machines within the provided desktop pool
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='base.desktop'; 'value' = $HVPoolID}
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        if($queryResults.count -ge 1){
            $queryResults=$HVConnectionServer.extensionData.machine.machine_getinfos($queryResults.id)
        }
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        return $queryResults
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the Horizon View machines.' -Exception $_
    }
}

function Get-HVRDSMachines {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "ID of the Horizon View RDS Farm.")]
        [VMware.Hv.FarmId]$HVFarmID,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'RDSServerInfo'
        # Filter for only the machines within the provided desktop pool
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='base.farm'; 'value' = $HVFarmID}
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        return $queryResults
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the Horizon View machines.' -Exception $_
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

$wrongsnapdesktops=@()
$wrongsnaphosts=@()

foreach ($hvconnectionserver in $hvconnectionservers){
    [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $hvconnectionserver -Credential $CredsHorizon
    # Retreive the name of the pod
    if ($HVpodstatus.status -eq "ENABLED"){
        [string]$podname=$objHVConnectionServer.extensionData.pod.Pod_list() | where-object {$_.localpod -eq $True} | select-object -expandproperty Displayname
    }
    else{
        $podname="N/A"
    }
    [array]$HVPools = Get-HVDesktopPools -HVConnectionServer $objHVConnectionServer
    foreach ($hvpool in $hvpools){
        $poolname=($hvpool).base.Name
        $HVDesktopmachines = Get-HVDesktopMachines -HVPoolID $HVPool.id -HVConnectionServer $objHVConnectionServer
        if ($HVDesktopmachines.count -ge 1){
            $wrongsnaps=$HVDesktopmachines | where {$_.managedmachinedata.viewcomposerdata.baseimagesnapshotpath -notlike  $hvpool.automateddesktopdata.VirtualCenternamesdata.snapshotpath -OR $_.managedmachinedata.viewcomposerdata.baseimagepath -notlike $hvpool.automateddesktopdata.VirtualCenternamesdata.parentvmpath}
            if ($wrongsnaps){
                foreach ($wrongsnap in $wrongsnaps){
                    $wrongsnapdesktops+= New-Object PSObject -Property @{
                        "Pod Name"                  = $podname;
                        "Desktop Name"              = $poolname;
                        "VM Name"                   = $wrongsnap.base.name;
                        "Status"                    = $wrongsnap.base.basicstate
                        "VM Snapshot"               = ($wrongsnap.managedmachinedata.viewcomposerdata.baseimagesnapshotpath).split("/")[-1];
                        "VM Golden image"           = ($wrongsnap.managedmachinedata.viewcomposerdata.baseimagepath).split("/")[-1];
                        "Pool Snapshot"             = ($hvpool.automateddesktopdata.VirtualCenternamesdata.snapshotpath).split("/")[-1];
                        "Pool Golden image"         = ($hvpool.automateddesktopdata.VirtualCenternamesdata.parentvmpath).split("/")[-1];
                    }
                }
            }
        }
    }
    [array]$HVFarms = Get-HVfarms -HVConnectionServer $objHVConnectionServer
    foreach ($HVFarm in $HVFarms){
        $farmname=($hvfarm).Data.Name
        $HVfarmmachines = Get-HVRDSMachines -HVFarmID $HVfarm.id -HVConnectionServer $objHVConnectionServer
        if ($HVfarmmachines.count -ge 1){
            $wrongsnaps=$HVfarmmachines | where {$_.rdsservermaintenancedata.baseimagesnapshotpath -notlike  $HVFarm.automatedfarmdata.VirtualCenternamesdata.snapshotpath -OR $_.rdsservermaintenancedata.baseimagepath -notlike $HVFarm.automatedfarmdata.VirtualCenternamesdata.parentvmpath}
            if ($wrongsnaps){
                foreach ($wrongsnap in $wrongsnaps){
                    $wrongsnaphosts+= New-Object PSObject -Property @{
                        "Pod Name"                  = $podname;
                        "Farm Name"                 = $farmname;
                        "RDS Name"                  = $wrongsnap.base.name;
                        "Status"                    = $wrongsnap.RuntimeData.Status
                        "Active Sessions"           = $wrongsnap.RuntimeData.SessionCount
                        "VM Snapshot"               = ($wrongsnap.rdsservermaintenancedata.baseimagesnapshotpath).split("/")[-1];
                        "VM Golden Image"           = ($wrongsnap.rdsservermaintenancedata.baseimagepath).split("/")[-1];
                        "Farm Snapshot"             = ($HVFarm.automatedfarmdata.VirtualCenternamesdata.snapshotpath).split("/")[-1];
                        "Farm Golden Image"         = ($HVFarm.automatedfarmdata.VirtualCenternamesdata.parentvmpath).split("/")[-1];
                    }
                }
            }
        }
    }
    Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
}
if($wrongsnapdesktops){
    Out-CUConsole -Message "VDI Machines based on wrong snapshot"
    $wrongsnapdesktops | format-table -groupby "Pod Name" -property "Desktop Name","VM Name","Status","VM Golden image","VM Snapshot","Pool Golden image","Pool Snapshot"
}
if($wrongsnaphosts){
    Out-CUConsole -Message "RDS Hosts based on wrong snapshot"
    $wrongsnaphosts | format-table -groupby "Pod Name" -property "Farm Name","RDS Name","Status","Active Sessions","VM Golden image","VM Snapshot","Farm Golden Image","Farm Snapshot"
}
if ($wrongsnaphosts.count -eq 0 -AND $wrongsnapdesktops.count -eq 0){
    Out-CUConsole -Message "No systems found running on the wrong snapshot."
}
