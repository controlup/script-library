#requires -Version 3.0
<#
    .SYNOPSIS
    Exports Metrics to be used by the COntrolUp Stress Calculator v2

    .DESCRIPTION
    Exports Metrics to be used by the ControlUp Stress Calculator v2 by using the ControlUp Monitor Powershell Module. Uses True or False as string for the benefit of using the script as a ControlUp Script Action.

    .EXAMPLE
    .\Stress Calculator Exports SBA.ps1 -ExportLocation "C:\Temp\exports\" -DatastoreData "True" -HostData "True" -vDiskData "True" -LogicalDiskData "True" -GatewayData "True"

    .EXAMPLE
    .\Stress Calculator Exports SBA.ps1 -ExportLocation "C:\Temp\exports\" -SessionData "False" -DatastoreData "True" -HostData "True" -vDiskData "True" -LogicalDiskData "True" -GatewayData "True"

    .PARAMETER ExportLocation
    Mandatory: Yes
    Type: String
    Location where metric exports will be stored.

    .PARAMETER MaxFileAge
    Mandatory: Yes
    Type: int
    Default: 7
    Maximum age in days for files in Export Location

    .PARAMETER SessionData
    Mandatory: No
    Type: String
    Default: "True"
    "True" or "False" wether to export Session Metrics or not

    .PARAMETER ComputerData
    Mandatory: No
    Type: String
    Default: "True"
    "True" or "False" wether to export Computer Metrics or not

    .PARAMETER FolderData
    Mandatory: No
    Type: String
    Default: "True"
    "True" or "False" wether to export Folder Metrics or not

    .PARAMETER DatastoreData
    Mandatory: No
    Type: String
    Default: "False"
    "True" or "False" wether to export Datastore Metrics or not

    .PARAMETER HostData
    Mandatory: No
    Type: String
    Default: "False"
    "True" or "False" wether to export Host Metrics or not

    .PARAMETER VirtualDiskData
    Mandatory: No
    Type: String
    Default: "False"
    "True" or "False" wether to export Virtual Disk Metrics or not

    .PARAMETER LogicalDiskData
    Mandatory: No
    Type: String
    Default: "False"
    "True" or "False" wether to export Logical Disk Metrics or not

    .PARAMETER GateWayData
    Mandatory: No
    Type: String
    Default: "False"
    "True" or "False" wether to export Gateway Metrics or not


    .PARAMETER LoadBalancerData
    Mandatory: No
    Type: String
    Default: "False"
    "True" or "False" wether to export LoadBalancer Metrics or not

    .PARAMETER LBServiceGroupData
    Mandatory: No
    Type: String
    Default: "False"
    "True" or "False" wether to export Load Balancer ServiceGroup Metrics or not

    .PARAMETER LBServiceData
    Mandatory: No
    Type: String
    Default: "False"
    "True" or "False" wether to export Load Balancer Service Metrics or not

    .PARAMETER NetscalerData
    Mandatory: No
    Type: String
    Default: "False"
    "True" or "False" wether to export NetScaler Metrics or not

    .NOTES
    Created by: Wouter Kursten
    First version: 04-07-2022

    Requires ControlUp 8.6.5 or Higher
    Requires ControlUp PowerShell Module
#>

[CmdletBinding()]
Param
(
    [Parameter(
        Mandatory=$true,
        HelpMessage='Folder to export the files to i.e. c:\temp\'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $ExportLocation,

    [Parameter(
        Mandatory=$false,
        HelpMessage='Maximum age in days for files in Export Location'
    )]
    [ValidateNotNullOrEmpty()]
    [int] $MaxFileAge = 7,

    [Parameter(
        Mandatory=$False,
        HelpMessage='"True" or False to export Session metrics'
    )]
    [ValidateSet("True","False")]
    [string] $SessionData = "True",

    [Parameter(
        Mandatory=$False,
        HelpMessage='"True" or False to export Computer metrics'
    )]
    [ValidateSet("True","False")]
    [string] $ComputerData = "True",

    [Parameter(
        Mandatory=$False,
        HelpMessage='"True" or False to export Folder metrics'
    )]
    [ValidateSet("True","False")]
    [string] $FolderData = "True",

    [Parameter(
        Mandatory=$False,
        HelpMessage='"True" or False to export Datastore metrics'
    )]
    [ValidateSet("True","False")]
    [string] $DatastoreData = "False",

    [Parameter(
        Mandatory=$False,
        HelpMessage='"True" or False to export Hosts metrics'
    )]
    [ValidateSet("True","False")]
    [string] $HostData = "False",

    [Parameter(
        Mandatory=$False,
        HelpMessage='"True" or False to export vDisks metrics'
    )]
    [ValidateSet("True","False")]
    [string] $VirtualDiskData = "False",

    [Parameter(
        Mandatory=$False,
        HelpMessage='"True" or False to export LogicalDisk metrics'
    )]
    [ValidateSet("True","False")]
    [string] $LogicalDiskData = "False",

    [Parameter(
        Mandatory=$False,
        HelpMessage='"True" or False to export Gateway metrics'
    )]
    [ValidateSet("True","False")]
    [string] $GatewayData = "False",

    [Parameter(
        Mandatory=$False,
        HelpMessage='"True" or False to export LoadBalancer metrics'
    )]
    [ValidateSet("True","False")]
    [string] $LoadBalancerData = "False",

    [Parameter(
        Mandatory=$False,
        HelpMessage='"True" or False to export LBServiceGroup metrics'
    )]
    [ValidateSet("True","False")]
    [string] $LBServiceGroupData = "False",

    [Parameter(
        Mandatory=$False,
        HelpMessage='"True" or False to export LBService metrics'
    )]
    [ValidateSet("True","False")]
    [string] $LBServiceData = "False",

    [Parameter(
        Mandatory=$False,
        HelpMessage='"True" or False to export NetScaler metrics'
    )]
    [ValidateSet("True","False")]
    [string] $NetScalerData = "False"
)

$ErrorActionPreference = 'Stop'

function Import-ControlUpModule {
    # Try Import-Module for each passed component.
    try {
        $pathtomodule = (Get-ChildItem "C:\Program Files\Smart-X\ControlUpMonitor\*ControlUp.PowerShell.User.dll" -Recurse | Sort-Object LastWriteTime -Descending)[0]
        write-verbose "Loading ControlUp PS module from $pathtomodule"
        Import-Module $pathtomodule
    }
    catch {
        write-error 'The required module was not found. Please make sure COntrolUP.CLI Module is installed and available for the user running the script.'
        }
}
#region Fieldsselection
[array]$computerfieldsarray = [array]$computerfieldsarray = "ParentFolderPath",
"ObjectGuid", #for the session to folder option
"ActiveMemory",
"AppLoadTime",
"AppResolusionTime",
"ASPRequestQueued",
"ASPRequestRejected",
"AvgDiskQueue",
"AvgDiskReadPerSec",
"AvgDiskWritePerSec",
"AvgProcessorLength",
"AvgTransactionTime",
"AvgUserInputDelay",
"AwsAlarmAlarm",
"AwsAlarmData",
"AwsAlarmOk",
"AwsCatalogPrice",
"AwsCPUUtilsization",
"AwsDiskReadOps",
"AwsDiskWriteOps",
"AwsEC2Price",
"AwsInstanceStatus",
"AwsMonthlyCost",
"AwsNetworkIn",
"AwsNetworkOut",
"AwsRIUpfrontPrice",
"AwsRootDeviceSize",
"AwsSnapshotPrice",
"AwsStatusChecks",
"AwsStoragePrice",
"AwsTotalPrice",
"AzureTotalComputeCost",
"AzureTotalDataDisksCost",
"AzureTotalDisksCost",
"AzureTotalMachineCost",
"AzureTotalOsDiskCost",
"AZVMTotalComputeCostLastMonth",
"CPU",
"CPUExSMPUse",
"CPUReady",
"CPUSwapWait",
"CPUSystemTime",
"CPUUsage",
"CSGConnections",
"DataStoreConnectionFailure",
"DBConnected",
"DBTransactionsErrorRate",
"DiskIOsPerSec",
"DiskReadsPerSec",
"DiskTime",
"DiskTransfersPerSec",
"DiskWritesPerSec",
"DroppedRx",
"DroppedTx",
"DynamicMemoryAvgPressure",
"GPUAvailableMemory",
"GPUFrameBufferSize",
"GPUFrameBufferUsage",
"GPUMemoryUsage",
"GPUUsage",
"HostCPUUsage",
"HostDataStoreReadLatency",
"HostDataStoreWriteLatency",
"HostDroppedRx",
"HostDroppedTx",
"HostMemUsage",
"HostvCPURatio",
"HzMaxConnectionCount",
"hzMaxSessionsConfigured",
"hzRdsServerHealth",
"HzServerHealthStatus",
"HzTotalConnections",
"hzUserSessions",
"ICASessions",
"LicenseConnectionFailure",
"LicenseLastCheckoutTime",
"LogonDurationAvg",
"MaxSpaceDrive",
"MaxUserInputDelay",
"MemoryDemand",
"MemoryInUse",
"MemoryMaximum",
"MemoryMinimum",
"MemUsage",
"MinSpaceDrive",
"NetERRIn",
"NetERROut",
"NetReceivedData",
"NetSentData",
"NetTotalData",
"NonPagedPoolMemory",
"NonZeroAvgUserInputDelay",
"NonZeroMaxUserInputDelay",
"oActiveSessions",
"oDisconnectedSessions",
"oErrorRate",
"oOtherSessions",
"oUserSessions",
"oWarningRate",
"PageFaultsPerSec",
"PagingFile",
"PVSBootRetryCount",
"PVSBootTime",
"PVSFreeSpaceWriteCacheDrive",
"PVSRamCacheUsgae",
"PVSServerReconnectCount",
"PVSTargetDeviceHealth",
"PVSUDPRetryCount",
"QueueReadyCount",
"RAMScore",
"ResolutionQueueReadyCount",
"ServerLoad",
"SessionDisconnectRate",
"SeverityLevel",
"StartupMemory",
"SwapInRateMem",
"SwapOutRateMem",
"SystemDriveFreeSpace",
"TopCitrixlicenseutilization",
"TotalProcesses",
"TotalReadRate",
"TotalSessions",
"TotalWriteRate",
"UptimeInDays",
"UptimeInSeconds",
"vDiskReadLatency",
"vDiskReadOpsPerSec",
"vDiskReadPerSec",
"vDiskWriteLatency",
"vDiskWriteOpsPerSec",
"vDiskWritePerSec",
"VirtualDiskAverageLatency",
"VirtualMachineMemoryBallooning",
"VirtualMachineSnapshotExists",
"VirtualMachineSnapshotSize",
"VMDaysSuspended",
"XDApplicationInstancesInUse",
"xdAvarageLogonDuration",
"xdBrokerHealth",
"xduserSessions"

[array]$hostfieldsarray = "hFolderPath",
"ActiveMemory",
"ConnectionState",
"CPUScore",
"CPUSpeed",
"CPUUsage",
"DataStoreReadLatency",
"DataStoreReadRate",
"DataStoreRWIOps",
"DatastoreTotalLatency",
"DataStoreWriteLatency",
"DataStoreWriteRate",
"DiskDeviceLatency",
"DiskKernelLatency",
"DiskQueueLatency",
"DiskScore",
"DiskTotalLatency",
"DroppedRx",
"DroppedTx",
"EntityType",
"ErrorRx",
"ErrorTx",
"HostCPUUsageMhz",
"MaxFreeSpaceDS",
"MemoryBallooning",
"MemoryCompressed",
"MemoryShared",
"MemorySwapping",
"MemUsage",
"MinFreeSpaceDS",
"NetworkScore",
"nicUsage",
"RAMScore",
"RunningVMCount",
"UnmanagedVMCount",
"vCPURatio",
"VMCount"

[array]$sessionFieldsarray= "ActiveTimePercentage",
"ComputerId", #for the session to folder option
"ApplicationsInUseCount",
"AppLoadTime",
"AssociationMask",
"AverageEncodingTime",
"ClientCPU",
"ClientDeviceScore",
"ClientMetricsSessionErrorCode",
"ClientNICSpeed",
"ClientPacketLoss",
"clientProtocolType",
"ComputerCpu",
"ComputerCpuReady",
"CPU_Usage",
"DisconnectedTimePercentage",
"DiskReadKBs",
"DiskWriteKBs",
"EucGatewayLatency",
"EucPlatform",
"FramesPerSecond",
"GPUDecoderUtilization",
"GPUEncoderUtilization",
"GPUFrameBufferMemoryUtilization",
"GPUUtilization",
"GroupPolicyProccesingTime",
"InternetLatency",
"IOReadOperationsPerSec",
"IOWriteOperationsPerSec",
"IspLatency",
"LANLatency",
"LegalNoticeDuration",
"LogonDuration",
"LogonDurationOther",
"NetworkReceiveKBs",
"NetworkSendKBs",
"NonZeroUserInputDelay",
"PacketLoss",
"PageFaultPerSec",
"PrivateUsage",
"ProtocolRTT",
"SessionLatency",
"SessionLatencyAvg",
"SeverityLevel",
"ShellLoadTime",
"TotalProcesses",
"TotalSessionLatency",
"UserInputDelay",
"UserProfileLoadTime",
"WorkingSet",
"xdAuthenticationDuration",
"xdBrokeringDuration",
"xdGPOLoadTime",
"xdHDXConnectionLoadTime",
"xdInteractiveSessionLoadTime",
"xdLogonDuration",
"xdProfileLoadTime",
"xdVMStartDuration"

[array]$vDiskfieldsarray = "hFolderPath",
"AzureDiskIopsLimit",
"AzureDiskMbpsLimit",
"AzureTotalDiskCost",
"Disk Capacity",
"EntityType",
"LargeSeeks",
"MediumSeeks",
"OutstandingReadRequests",
"OutstandingRequests",
"OutstandingWriteRequests",
"ReadIOPS",
"ReadLatency",
"ReadLoadMetric",
"ReadRate",
"ReadRequestSize",
"SeverityLevel",
"SmallSeeks",
"vDiskShareAllocation",
"WriteIOPS",
"WriteLatency",
"WriteLoadMetric",
"WriteRate",
"WriteRequestSize"

[array]$datastoresFieldsarray = "hFolderPath",
"ActiveTime",
"DatastoreMaxQueueDepth",
"FreeSpace",
"FreeSpacePercentage",
"I/OLatency",
"IOPS",
"MaxQueueDepth",
"ReadCacheHitRate",
"ReadCacheReadLatency",
"ReadCacheWriteLatency",
"ReadIOPS",
"ReadLatency",
"ReadRate",
"SDRSOutstandingReadRequests",
"SDRSOutstandingWriteRequests",
"SDRSReadLatency",
"SDRSReadRate",
"SDRSWriteIOPS",
"SDRSWriteLatency",
"vSANCacheHitIOPS",
"vSANCacheHitRate",
"vSANCongestions",
"vSANInboundIOThroughput",
"vSANInboundPacketLossRate",
"vSANInboundPacketsPerSecond",
"vSANOutboundIOThroughput",
"vSANOutboundPacketLossRate",
"vSANOutboundPacketsPerSecond",
"vSANOutstandingIO",
"WriteBufferMinFreePercentage",
"WriteBufferReadLatency",
"WriteBufferWriteLatency",
"WriteIOPS",
"WriteLatency",
"WriteRate"

[array]$LogicalDiskfieldsarray = "AvgDiskBytesRead",
"ComputerGuid", #for the session to folder option
"AvgDiskBytesTransfer",
"AvgDiskBytesWrite",
"AvgDiskQueueLength",
"AvgDiskReadQueueLength",
"AvgDiskWriteQueueLength",
"CurrentDiskQueueLength",
"DiskKBps",
"DiskReadKBps",
"DiskReadssec",
"DiskTransferssec",
"DiskWriteKBps",
"DiskWritessec",
"FreeInodes",
"FreeSpace",
"FreeSpacePercentage",
"SplitIOSec",
"TotalInodes",
"UsedInodes"

[array]$folderfieldsarray = "AppLoadTime",
"AverageComputerMemoryUtilization",
"AverageGPUFrameBufferUsage",
"AverageGPUUsage",
"AvgHostCPU",
"AvgHostMem",
"Path",
"AvgSessionPerComputer",
"AWSAlarms",
"AwsAvgCpuUtilzation",
"AwsEc2Instances",
"AwsFolderAlarmInAlarm",
"AwsFolderAlarmsInsuffiecientData",
"AwsFolderAlarmsOk",
"AwsFolderAttachedSnapshots",
"AwsFolderAttachedStorage",
"AwsFolderDaysToNextRi",
"AwsFolderDetachedSnapshots",
"AwsFolderDetachedStorage",
"AwsFolderDiskReadOperations",
"AwsFolderDiskWriteOperations",
"AwsFolderEc2Cost",
"AwsFolderMontlyCost",
"AwsFolderNetworkIn",
"AwsFolderNetworkOut",
"AwsFolderNumOfUnUtilized",
"AwsFolderTotalCost",
"AwsFolderUnUtilizedCost",
"AwsFolderUpfrontCost",
"AwsInstancesFailedStatus",
"AwsInstancesWithFailedStatus",
"AwsMonthlyRunRate",
"AwsRunningInstances",
"AwsStoppedInstances",
"AZComputeCostLastMonth",
"AZCostLastMonth",
"AZDiskCostLastMonth",
"AZForecastCost",
"AZLastMonthStoppedMachinesCost",
"AZMachineCostLastMonth",
"AZNetworkCostLastMonth",
"AZScaleSetCostLastMonth",
"AZSnapshotCostLastMonth",
"AzStatus",
"AZStorageAccountsCostLastMonth",
"AzureDisks",
"AzureMachines",
"AzureProvState",
"AzureResourceGroups",
"AzureRunningMachines",
"AzureStoppedMachines",
"AzureStoppedMachinesCost",
"AzureTotalComputeCost",
"AzureTotalCost",
"AzureTotalDiskCapacity",
"AzureTotalDiskCost",
"AzureTotalMachineCost",
"AzureTotalNetworkCost",
"AzureTotalScaleSetCost",
"AzureTotalSnapshotCost",
"AzureTotalStorageAccountCost",
"CitrixAndCloudConnections",
"CitrixCloudConnections",
"CitrixConnections",
"ClientSessionsCount",
"ClientSessionsErrorCount",
"CloudMachines",
"CloudRunningMachines",
"CloudStoppedMachines",
"Clusters",
"ComputerDiskIOAvgLatency",
"ComputerDiskTransfersSec",
"ComputerNetTotal",
"ComputerswithGPU",
"ConnectedCloudConnectors",
"ConnectedSources",
"ConnectorCount",
"CPU",
"CriticalStressedComputers",
"CriticalStressedHosts",
"CriticalStressedProcesses",
"CriticalStressedSessions",
"CriticalStressedVMs",
"CUDCCollectionModifiedErrors",
"CUDCConcecutiveErrorLimit",
"CUDCConcecutiveErrors",
"CUDCErrors",
"CUDCLastCollectionDuration",
"CUDCLongetModelQuery",
"CUDCThrottlingErrors",
"Datacenters",
"DatastoreCount",
"DatastoreFreeSpace",
"DatastoreReadIOPS",
"DatastoreReadLatency",
"DatastoreReadRate",
"DatastoreWriteIOPS",
"DatastoreWriteLatency",
"DatastoreWriteRate",
"DisconnectedComputers",
"FolderxdUserSessions",
"Gateways",
"HAStatus",
"HighStressedClientSessions",
"HighStressedComputers",
"HighStressedHosts",
"HighStressedProcesses",
"HighStressedSessions",
"HighStressedVMs",
"HorizonConnections",
"HostAverageDiskIOPS",
"HostCount",
"HostNICaverageusage",
"HostNICdroppedpackets",
"HostsDatastoreCount",
"hzAutomaticLogoffTimeout",
"HzAvailableDesktopPoolMachines",
"HzAvailableDesktopPools",
"HzAvailableDesktopPoolsPercent",
"HzAvailableFarms",
"HzAvailableFarmsPercent",
"HzAvailableMachines",
"HzAvailableMachinesPercent",
"HzAvailableRdsDesktopPoolMachines",
"HzConnectionServers",
"HzDesktopPools",
"HzDisconnectedSessions",
"HzDisconnectedSessionsFarmsOnly",
"HzDisconnectedSessionsPoolsOnly",
"HzEmptySessionTimeout",
"HzFarms",
"HzHealthyConnectionServers",
"HzHealthyConnectionServersPercent",
"hzMachineCount",
"HzMachinesInUse",
"HzMachinesWithoutConnectionServersCount",
"HzMaxConnectionCount",
"hzMaxNumberOfMonitors",
"HZPreparingMachines",
"HZPreparingMachinesPercent",
"HzProblematicMachines",
"HzTotalComposerMachineConnection",
"HzTotalConnections",
"HzTotalUserSessions",
"hzvRamSize",
"LoadBalancers",
"LogonDurationAvg",
"LowStressedComputers",
"LowStressedHosts",
"LowStressedProcesses",
"LowStressedSessions",
"LowStressedVMs",
"Machines",
"ManagedComputerCount",
"MaxSessionPerComputer",
"MinSessionPerComputer",
"NetScalerAppliances",
"NetScalerApplianceswithUnsavedConfig",
"NetScalerGatewayTotalBytesInsec",
"NetScalerGatewayTotalBytesOutsec",
"NetScalerLBBytesInsec",
"NetScalerLBBytesOutsec",
"NetScalerLBTotalHitRate",
"NetScalerPercentageofGWsUp",
"NetScalerPercentageofLBsUp",
"NetScalerTotalBytesInsec",
"NetScalerTotalBytesOutsec",
"NetScalerTotalHDXSessions",
"NetScalerTotalLBUsersConnections",
"NetScalerTotalSSLUsers",
"NoStressedClientSessions",
"OnlineMachines",
"PageFaultPerSec",
"Processes",
"ReadCacheHitRate",
"ReadWriteAverageDiskLatency",
"RunningVMCount",
"Sessions",
"SessionsAvgProtocolLatency",
"SeverityLevel",
"TotalDatastoresCapacity",
"TotalSessions",
"UserInputDelayAvg",
"UserSessions",
"vdaMachines",
"VMActiveMemory",
"VMCount",
"VMCPUReady",
"VMHostCount",
"VMHostCPUUsage",
"VMPhysicalMemoryUsed",
"VMVirtualDiskReadIOPS",
"VMVirtualDiskReadLatency",
"VMVirtualDiskReadsKBps",
"VMVirtualDiskWriteIOPS",
"VMVirtualDiskWritesKBps",
"VMVirtualDiskWritLatency",
"vSANCacheHitRate",
"vSANCongestions",
"vSANFreeCapacity",
"vSANHealth",
"vSANInboundPacketLossRate",
"vSANOutboundPacketLossRate",
"WriteBufferMinFreePercentage",
"xdAverageLogonDuration",
"xdBrokersCount",
"xdComputers",
"xdDeliveryGroups",
"xdDeliveryGroupsAvailable",
"xdDeliveryGroupsAvailableSum",
"xdDesktopsAvailable",
"xdDesktopsDisconnected",
"xdDesktopsInUse",
"xdDesktopsNeverRegistered",
"xdDesktopsPreparing",
"xdDesktopsUnregistered",
"xdFailedMachines",
"XDFolderDesktopInstancesInUse",
"XDFolderInUsePublishedRatio",
"XDFolderPublishedAppsInUse",
"XDFolderPublishedInstancesInUse",
"xdFolderType",
"xdHealthyBrokers",
"xdHealthyBrokersCount",
"xdMachinesAvailable",
"xdPercentageofDesktopsAvailable",
"xdPublishedApplications",
"xdTotalDesktops",
"xdUnmanagedComputers",
"xdUserConnectionFailuresPerHour"

#endregion


Import-ControlUpModule

write-host "Starting to export data to $ExportLocation "

if($SessionData -eq "True"){
    write-verbose "Exporting Session Data"
    try{
        $export = export-cuquery -scheme Main -table Sessions -Fields $sessionFieldsarray -OutputFolder $ExportLocation
    }
    Catch{
        throw $_
    }
    $Count=$export.Total
    $Location = $export.FullFilePath
    write-host "Exported $Count lines with Session data to $Location"
}
else{
    write-verbose "SessionData was set to $SessionData, skipping Session Data"
}

if($ComputerData -eq "True"){
    write-verbose "Exporting Computer data"
    try{
        $export = export-cuquery -scheme Main -table Computers -Fields $computerfieldsarray -Where "isConnected= 'True'" -OutputFolder $ExportLocation
    }
    Catch{
        throw $_
    }
    $Count=$export.Total
    $Location = $export.FullFilePath
    write-host "Exported $Count lines with Computer data to $Location"
}
else{
    write-verbose "ComputerData was set to $ComputerData, skipping Computer Data"
}

if($FolderData  -eq "True"){
    write-verbose "Exporting Folder data"
    try{
        $export = export-cuquery -scheme Main -table Folders -Fields $folderfieldsarray -OutputFolder $ExportLocation
    }
    Catch{
        throw $_
    }
    $Count=$export.Total
    $Location = $export.FullFilePath
    write-host "Exported $Count lines with Folder data to $Location"
}
else{
    write-verbose "FolderData was set to $FolderData, skipping Folder Data"
}

if($DatastoreData  -eq "True"){
    write-verbose "Exporting Datastore data"
    try{
        $export = export-cuquery -scheme Main -table Datastores -Fields $datastoresFieldsarray -OutputFolder $ExportLocation
        $export2 = export-cuquery -scheme Main -table "Datastores on Hosts" -FIelds $datastoresFieldsarray -OutputFolder $ExportLocation
    }
    Catch{
        throw $_
    }
    $Count=$export.Total
    $Location = $export.FullFilePath
    $Count2=$export2.Total
    $Location2 = $export2.FullFilePath
    write-host "Exported $Count lines with Datastore data to $Location"
    write-host "Exported $Count2 lines with Datastore on Hosts data to $Location2"
}
else{
    write-verbose "DatastoreData was set to $DatastoreData, skipping Datastore Data"
}

if($HostData  -eq "True"){
    write-verbose "Exporting Host data"
    try{
        $export = export-cuquery -scheme Main -table Hosts -Fields $hostfieldsarray -OutputFolder $ExportLocation
    }
    Catch{
        throw $_
    }
    $Count=$export.Total
    $Location = $export.FullFilePath
    write-host "Exported $Count lines with Host data to $Location"
}
else{
    write-verbose "HostData was set to $HostData, skipping Host Data"
}

if($VirtualDiskData  -eq "True"){
    write-verbose "Exporting vDisk data"
    try{
        $export = export-cuquery -scheme Main -table vDisks -Fields $vDiskfieldsarray -OutputFolder $ExportLocation
    }
    Catch{
        throw $_
    }
    $Count=$export.Total
    $Location = $export.FullFilePath
    write-host "Exported $Count lines with vDisk data to $Location"
}
else{
    write-verbose "VirtualDiskData was set to $VirtualDiskData, skipping vDisk Data"
}

if($LogicalDiskData  -eq "True"){
    write-verbose "Exporting Logical Disk data"
    try{
        $export = export-cuquery -scheme Main -table LogicalDisks -Fields $LogicalDiskfieldsarray -OutputFolder $ExportLocation
    }
    Catch{
        throw $_
    }
    $Count=$export.Total
    $Location = $export.FullFilePath
    write-host "Exported $Count lines with Logical Disk data to $Location"
}
else{
    write-verbose "LogicalDiskData was set to $LogicalDiskData, skipping Logical DIsk Data"
}

if($GatewayData  -eq "True"){
    write-verbose "Exporting Gateway data"
    try{
        $export = export-cuquery -scheme Main -table Gateways -Fields * -OutputFolder $ExportLocation
    }
    Catch{
        throw $_
    }
    $Count=$export.Total
    $Location = $export.FullFilePath
    write-host "Exported $Count lines with Gateway data to $Location"
}
else{
    write-verbose "GatewayData was set to $GatewayData, skipping Gateway Data"
}

if($LoadBalancerData  -eq "True"){
    write-verbose "Exporting LoadBalancer data"
    try{
        $export = export-cuquery -scheme Main -table LoadBalancers -FIelds * -OutputFolder $ExportLocation
    }
    Catch{
        throw $_
    }
    $Count=$export.Total
    $Location = $export.FullFilePath
    write-host "Exported $Count lines with LoadBalancer data to $Location"
}
else{
    write-verbose "LoadBalancerData was set to $LoadBalancerData, skipping LoadBalancer Data"
}

if($LBServiceGroupData  -eq "True"){
    write-verbose "Exporting LBServiceGroup data"
    try{
        $export = export-cuquery -scheme Main -table LBServiceGroups -FIelds * -OutputFolder $ExportLocation
    }
    Catch{
        throw $_
    }
    $Count=$export.Total
    $Location = $export.FullFilePath
    write-host "Exported $Count lines with LBServiceGroup data to $Location"
}
else{
    write-verbose "LBServiceGroupData was set to $LBServiceGroupData, skipping LBServiceGroup Data"
}

if($LBServiceData  -eq "True"){
    write-verbose "Exporting LBService data"
    try{
        $export = export-cuquery -scheme Main -table LBServices -FIelds * -OutputFolder $ExportLocation
    }
    Catch{
        throw $_
    }
    $Count=$export.Total
    $Location = $export.FullFilePath
    write-host "Exported $Count lines with LBService data to $Location"
}
else{
    write-verbose "LBServiceData was set to $LBServiceData, skipping LBService Data"
}

if($NetScalerData  -eq "True"){
    write-verbose "Exporting NetScaler data"
    try{
        $export = export-cuquery -scheme Main -table NetScalers -FIelds * -OutputFolder $ExportLocation
    }
    Catch{
        throw $_
    }
    $Count=$export.Total
    $Location = $export.FullFilePath
    write-host "Exported $Count lines with NetScaler data to $Location"
}
else{
    write-verbose "NetScalerData was set to $NetScalerData, skipping NetScaler Data"
}

$oldestfiledate = (get-date).AddDays(-$MaxFileAge)

try{
    write-verbose "Removing files from before $oldestfiledate"
    $filestoremove = Get-ChildItem -Path $ExportLocation | Where-Object {$_.LastWriteTime -lt $oldestfiledate} | Remove-Item -force -Confirm:$False
    write-verbose "Finished removing files"
}
catch{
    throw $_
}

write-host "Finished Exporting metrics"
