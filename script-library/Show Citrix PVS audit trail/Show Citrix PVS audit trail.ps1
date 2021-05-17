#requires -version 3.0
<#
    Show or export Citrix PVS audit logs.
    Enable auditing if not enabled and requested via parameter.

    @guyrleech, 2018

    Modification history:

#>

[string]$outputFile = $null
[bool]$enableAuditing = $false
[int]$outputWidth = 400
[string]$pvsModule = "$env:ProgramFiles\Citrix\Provisioning Services Console\Citrix.PVS.SnapIn.dll"
[int]$ERROR_INVALID_PARAMETER = 87

$startDate = $null
$endDate = $null

if( $args.Count -ge 1 -and $args[0] )
{
    $enableAuditing = $args[0] -eq 'true'
}

if( $args.Count -ge 2 -and $args[1] )
{
    ## This could be specified as a date/time or last days/minutes/hours/etc so figure out which
    $result = New-Object DateTime
    if( [datetime]::TryParse( $args[1] , [ref]$result ) )
    {
        $startDate = $result
    }
    else
    {
        ## see what last character is as will tell us what units to work with
        [string]$last = $args[1]
        [long]$multiplier = 0
        switch( $last[-1] )
        {
            "s" { $multiplier = 1 }
            "m" { $multiplier = 60 }
            "h" { $multiplier = 3600 }
            "d" { $multiplier = 86400 }
            "w" { $multiplier = 86400 * 7 }
            "y" { $multiplier = 86400 * 365 }
            default { Throw "Unknown multiplier `"$($last[-1])`"" }
        }
        $endDate = Get-Date
        if( $last.Length -le 1 )
        {
            $startDate = $endDate.AddSeconds( -$multiplier )
        }
        else
        {
            $startDate = $endDate.AddSeconds( - ( ( $last.Substring( 0 ,$last.Length - 1 ) -as [long] ) * $multiplier ) )
        }
    }
}
else
{
    $startDate = (Get-Date).AddDays( -1 )
    $endDate = Get-Date
}

if( $args.Count -ge 3 -and $args[2] )
{
    ## This could be specified as a date/time or a duration of days/minutes/hours/etc so figure out which
    $result = New-Object DateTime
    if( [datetime]::TryParse( $args[2] , [ref]$result ) )
    {
        $endDate = $result
    }
    else
    { 
        [string]$last = $args[2]
        [long]$multiplier = 0
        switch( $last[-1] )
        {
            "s" { $multiplier = 1 }
            "m" { $multiplier = 60 }
            "h" { $multiplier = 3600 }
            "d" { $multiplier = 86400 }
            "w" { $multiplier = 86400 * 7 }
            "y" { $multiplier = 86400 * 365 }
            default { Throw "Unknown multiplier `"$($last[-1])`"" }
        }
        if( $last.Length -le 1 )
        {
            $endDate = $startDate.AddSeconds( $multiplier )
        }
        else
        {
            $endDate = $startDate.AddSeconds( ( ( $last.Substring( 0 ,$last.Length - 1 ) -as [long] ) * $multiplier ) )
        }
    }
}

[string[]]$audittypes = @(
    'Many' , 
    'AuthGroup' , 
    'Collection' , 
    'Device' , 
    'Disk' , 
    'DiskLocator' , 
    'Farm' , 
    'FarmView' , 
    'Server' , 
    'Site' , 
    'SiteView' , 
    'Store' ,
    'System' , 
    'UserGroup'
)

[hashtable]$auditActions = @{
 1 = 'AddAuthGroup'
 2 = 'AddCollection'
 3 = 'AddDevice'
 4 = 'AddDiskLocator'
 5 = 'AddFarmView'
 6 = 'AddServer'
 7 = 'AddSite'
 8 = 'AddSiteView'
 9 = 'AddStore'
 10 = 'AddUserGroup'
 11 = 'AddVirtualHostingPool'
 12 = 'AddUpdateTask'
 13 = 'AddDiskUpdateDevice'
 1001 = 'DeleteAuthGroup'
 1002 = 'DeleteCollection'
 1003 = 'DeleteDevice'
 1004 = 'DeleteDeviceDiskCacheFile'
 1005 = 'DeleteDiskLocator'
 1006 = 'DeleteFarmView'
 1007 = 'DeleteServer'
 1008 = 'DeleteServerStore'
 1009 = 'DeleteSite'
 1010 = 'DeleteSiteView'
 1011 = 'DeleteStore'
 1012 = 'DeleteUserGroup'
 1013 = 'DeleteVirtualHostingPool'
 1014 = 'DeleteUpdateTask'
 1015 = 'DeleteDiskUpdateDevice'
 1016 = 'DeleteDiskVersion'
 2001 = 'RunAddDeviceToDomain'
 2002 = 'RunApplyAutoUpdate'
 2003 = 'RunApplyIncrementalUpdate'
 2004 = 'RunArchiveAuditTrail'
 2005 = 'RunAssignAuthGroup'
 2006 = 'RunAssignDevice'
 2007 = 'RunAssignDiskLocator'
 2008 = 'RunAssignServer'
 2009 = 'RunWithReturnBoot'
 2010 = 'RunCopyPasteDevice'
 2011 = 'RunCopyPasteDisk'
 2012 = 'RunCopyPasteServer'
 2013 = 'RunCreateDirectory'
 2014 = 'RunCreateDiskCancel'
 2015 = 'RunDisableCollection'
 2016 = 'RunDisableDevice'
 2017 = 'RunDisableDeviceDiskLocator'
 2018 = 'RunDisableDiskLocator'
 2019 = 'RunDisableUserGroup'
 2020 = 'RunDisableUserGroupDiskLocator'
 2021 = 'RunWithReturnDisplayMessage'
 2022 = 'RunEnableCollection'
 2023 = 'RunEnableDevice'
 2024 = 'RunEnableDeviceDiskLocator'
 2025 = 'RunEnableDiskLocator'
 2026 = 'RunEnableUserGroup'
 2027 = 'RunEnableUserGroupDiskLocator'
 2028 = 'RunExportOemLicenses'
 2029 = 'RunImportDatabase'
 2030 = 'RunImportDevices'
 2031 = 'RunImportOemLicenses'
 2032 = 'RunMarkDown'
 2033 = 'RunWithReturnReboot'
 2034 = 'RunRemoveAuthGroup'
 2035 = 'RunRemoveDevice'
 2036 = 'RunRemoveDeviceFromDomain'
 2037 = 'RunRemoveDirectory'
 2038 = 'RunRemoveDiskLocator'
 2039 = 'RunResetDeviceForDomain'
 2040 = 'RunResetDatabaseConnection'
 2041 = 'RunRestartStreamingService'
 2042 = 'RunWithReturnShutdown'
 2043 = 'RunStartStreamingService'
 2044 = 'RunStopStreamingService'
 2045 = 'RunUnlockAllDisk'
 2046 = 'RunUnlockDisk'
 2047 = 'RunServerStoreVolumeAccess'
 2048 = 'RunServerStoreVolumeMode'
 2049 = 'RunMergeDisk'
 2050 = 'RunRevertDiskVersion'
 2051 = 'RunPromoteDiskVersion'
 2052 = 'RunCancelDiskMaintenance'
 2053 = 'RunActivateDevice'
 2054 = 'RunAddDiskVersion'
 2055 = 'RunExportDisk'
 2056 = 'RunAssignDisk'
 2057 = 'RunRemoveDisk'
 2058 = 'RunDiskUpdateStart'
 2059 = 'RunDiskUpdateCancel'
 2060 = 'RunSetOverrideVersion'
 2061 = 'RunCancelTask'
 2062 = 'RunClearTask'
 2063 = 'RunForceInventory'
 2064 = 'RunUpdateBDM'
 2065 = 'RunStartDeviceDiskTempVersionMode'
 2066 = 'RunStopDeviceDiskTempVersionMode'
 3001 = 'RunWithReturnCreateDisk'
 3002 = 'RunWithReturnCreateDiskStatus'
 3003 = 'RunWithReturnMapDisk'
 3004 = 'RunWithReturnRebalanceDevices'
 3005 = 'RunWithReturnCreateMaintenanceVersion'
 3006 = 'RunWithReturnImportDisk'
 4001 = 'RunByteArrayInputImportDevices'
 4002 = 'RunByteArrayInputImportOemLicenses'
 5001 = 'RunByteArrayOutputArchiveAuditTrail'
 5002 = 'RunByteArrayOutputExportOemLicenses'
 6001 = 'SetAuthGroup'
 6002 = 'SetCollection'
 6003 = 'SetDevice'
 6004 = 'SetDisk'
 6005 = 'SetDiskLocator'
 6006 = 'SetFarm'
 6007 = 'SetFarmView'
 6008 = 'SetServer'
 6009 = 'SetServerBiosBootstrap'
 6010 = 'SetServerBootstrap'
 6011 = 'SetServerStore'
 6012 = 'SetSite'
 6013 = 'SetSiteView'
 6014 = 'SetStore'
 6015 = 'SetUserGroup'
 6016 = 'SetVirtualHostingPool'
 6017 = 'SetUpdateTask'
 6018 = 'SetDiskUpdateDevice'
 7001 = 'SetListDeviceBootstraps'
 7002 = 'SetListDeviceBootstrapsDelete'
 7003 = 'SetListDeviceBootstrapsAdd'
 7004 = 'SetListDeviceCustomProperty'
 7005 = 'SetListDeviceCustomPropertyDelete'
 7006 = 'SetListDeviceCustomPropertyAdd'
 7007 = 'SetListDeviceDiskPrinters'
 7008 = 'SetListDeviceDiskPrintersDelete'
 7009 = 'SetListDeviceDiskPrintersAdd'
 7010 = 'SetListDevicePersonality'
 7011 = 'SetListDevicePersonalityDelete'
 7012 = 'SetListDevicePersonalityAdd'
 7013 = 'SetListDiskLocatorCustomProperty'
 7014 = 'SetListDiskLocatorCustomPropertyDelete'
 7015 = 'SetListDiskLocatorCustomPropertyAdd'
 7016 = 'SetListServerCustomProperty'
 7017 = 'SetListServerCustomPropertyDelete'
 7018 = 'SetListServerCustomPropertyAdd'
 7019 = 'SetListUserGroupCustomProperty'
 7020 = 'SetListUserGroupCustomPropertyDelete'
 7021 = 'SetListUserGroupCustomPropertyAdd'
}

[hashtable]$auditParams = @{}

if( ! [string]::IsNullOrEmpty( $startDate ) )
{
    $auditParams.Add( 'BeginDate' , [datetime]::Parse( $startDate ) )
}

if( ! [string]::IsNullOrEmpty( $endDate ) )
{
    $auditParams.Add( 'EndDate' , [datetime]::Parse( $endDate ) )
    if( ! [string]::IsNullOrEmpty( $startDate ) )
    {
        if( $auditParams[ 'EndDate' ] -lt $auditParams[ 'BeginDate' ] )
        {
            Write-Error "End date $endDate earlier than start date $startDate"
            Exit $ERROR_INVALID_PARAMETER
        }
    }
}

if( ! [string]::IsNullOrEmpty( $pvsModule ) )
{
    Import-Module $pvsModule -ErrorAction Stop
}

## Check if auditing is enabled
$farm = Get-PvsFarm
if( $farm -and ! $farm.AuditingEnabled )
{
    Write-Warning "Auditing is not enabled on farm `"$($farm.Name)`""
    if( $enableAuditing )
    {
        Set-PvsFarm -FarmId $farm.FarmId -AuditingEnabled:$true
        if( $? )
        {
            "Auditing succesfully enabled"
        }
        else
        {
            Write-Warning "Failed to enable auditing"
        }
    }
}
elseif( ! $farm )
{
    Write-Warning "Failed to retrieve PVS farm details"
}     

[hashtable]$sites = @{}
[hashtable]$stores = @{}
[hashtable]$collections = @{}

## Lookup table for site id to name
Get-PvsSite | ForEach-Object `
{
    $sites.Add( $_.SiteId , $_.SiteName )
}
Get-PvsCollection | ForEach-Object `
{
    $collections.Add( $_.CollectionId , $_.CollectionName )
}
Get-PvsStore | ForEach-Object `
{
    $stores.Add( $_.StoreId , $_.StoreName )
}

[array]$auditevents = @( Get-PvsAuditTrail @auditParams | ForEach-Object `
{
    $auditItem = $_
    [string]$subItem = $null
    if( ! [string]::IsNullOrEmpty( $auditItem.SubId ) ) ## GUID of the Collection or Store of the action
    {
        $subItem = $collections[ $auditItem.SubId ]
        if( [string]::IsNullOrEmpty( $subItem ) )
        {
            $subItem = $stores[ $auditItem.SubId ]
        }
    }
    [string]$parameters = $null
    [string]$properties = $null
    if( $auditItem.Attachments -band 0x4 ) ## parameters
    {
        $parameters = ( Get-PvsAuditActionParameter -AuditActionId $auditItem.AuditActionId | ForEach-Object `
        {
            "$($_.name)=$($_.value) "
        } )
    }
    if( $auditItem.Attachments -band 0x8 ) ## properties
    {
        $properties = ( Get-PvsAuditActionProperty -AuditActionId $auditItem.AuditActionId | ForEach-Object `
        {
            "$($_.name):$($_.OldValue)=>$($_.NewValue) "
        } )
    }
    [PSCustomObject]@{ 
        'Time' = $auditItem.Time
        'Domain' = $auditItem.Domain
        'User' = $auditItem.UserName
        'Type' = $audittypes[ $auditItem.Type ]
        'Action' = $auditActions[ $auditItem.Action -as [int] ]
        'Object Name' = $auditItem.ObjectName
        'Sub Item' = $subItem
        'Path' = $auditItem.Path
        'Site' = $sites[ $auditItem.SiteId ] 
        'Properties' = $properties
        'Parameters' = $parameters }
} ) | Sort Time -Descending

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

[string]$message = "Got $(if( $auditevents -and $auditevents.Count ) { $auditevents.Count } else { 'no' }) audit events"
if( $startDate )
{
    $message += " from $(Get-Date $startDate -Format G)"
}

if( $endDate )
{
    $message += " until $(Get-Date $endDate -Format G)"
}

if( $auditevents -and $auditevents.Count )
{
    $message

    if( ! [string]::IsNullOrEmpty( $outputFile ) )
    {
        "Writing to `"$outputFile`""
        $auditevents | Export-Csv -Path $outputFile -NoClobber -NoTypeInformation
    }

    $auditevents | Format-Table -AutoSize
}
else
{
    Write-Warning $message
}

