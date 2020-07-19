<#
    Find Target devices not booted off the latest or assigned vdisk

    @guyrleech 2018

    Modification History:

    04/04/19  GRL  Made dual purpose to report all active devices, optionally matching a regex pattern
#>

## change this depending on whether it is the SBA to show just wrongly booted devices or to report all
[bool]$showAll = $true
[string]$pattern = if( $args.Count -and $args[0] ) { $args[0] }

[int]$outputWidth = 400

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

Import-Module -Name  "$env:ProgramFiles\Citrix\Provisioning Services Console\Citrix.PVS.SnapIn.dll" -ErrorAction Stop

## Get Device info in one go as quite slow
[hashtable]$deviceInfos = @{}
Get-PvsDeviceInfo | ForEach-Object `
{
    $deviceInfos.Add( $_.DeviceId , $_ )
}

# Cache all disk and version info
[hashtable]$diskVersions = @{}

Get-PvsSite | ForEach-Object `
{
    Get-PvsDiskInfo -SiteId $_.SiteId | ForEach-Object `
    {
        $diskVersions.Add( $_.DiskLocatorId , @( Get-PvsDiskVersion -DiskLocatorId $_.DiskLocatorId ) )
    }
}

## Cache store locations so we can look up vdisk sizes
[hashtable]$stores = @{}
Get-PvsStore | ForEach-Object `
{
    $stores.Add( $_.StoreName , $_.Path )
}

[string[]]$cacheTypes = 
@(
    'Standard Image' ,
    'Cache on Server', 
    'Standard Image' ,
    'Cache in Device RAM', 
    'Cache on Device Hard Disk', 
    'Standard Image' ,
    'Device RAM Disk', 
    'Cache on Server, Persistent',
    'Standard Image' ,
    'Cache in Device RAM with Overflow on Hard Disk' 
)

[string[]]$accessTypes = 
@(
    'Production', 
    'Maintenance', 
    'Maintenance Highest Version', 
    'Override', 
    'Merge', 
    'MergeMaintenance', 
    'MergeTest'
    'Test'
)

[int]$bootedOffWrongDisk = 0
[int]$bootedOffWrongVersion = 0

## Get boot events so we can report last boot since we may not be able to use remote WMI/CIM because of the account we are running under
[hashtable]$bootTimes = @{}
if( $showAll )
{
    Get-WinEvent -FilterHashtable @{Logname='Application';ID=10;ProviderName='StreamProcess'} -ErrorAction SilentlyContinue | Where-Object { $_.Message -match 'boot time'} | ForEach-Object `
    {
        if( $_.Message -match '^Device (?<Target>[^\s]+) boot time: (?<minutes>\d+) minutes (?<seconds>\d+) seconds\.$' -and ($saveMatches = $Matches.Clone()) -and $Matches[ 'Target' ] -match $pattern  )
        {
            try
            {
                $details = New-Object PSCustomObject
                Add-Member -InputObject $details -MemberType NoteProperty -Name 'UpTime' -Value ([math]::Round( ( New-TimeSpan -Start $_.TimeCreated -End ([datetime]::Now)).TotalDays , 1 ))
                Add-Member -InputObject $details -MemberType NoteProperty -Name 'BootTime' -Value ( ( $saveMatches[ 'minutes' ] -as [int] ) * 60 + ( $saveMatches[ 'seconds' ] -as [int] ) )
                $bootTimes.Add( $saveMatches[ 'Target' ] , $details )
            }
            catch
            {
                ## already got it so not the latest
            }
        }
    }
}

[int]$count = 0
[array]$results = @( Get-PvsDevice | Where-Object { $_.Active -and $_.Name -match $pattern } | ForEach-Object `
{
    $count++
    $device = $_
    [hashtable]$fields = [ordered]@{
        'Device Name' = $device.Name
        'Site' = $device.SiteName
        'Collection' = $device.CollectionName
    }
    
    ## Can't easily cache this since needs each device's deviceid
    $vDisk = Get-PvsDiskInfo -DeviceId $device.DeviceId
    $versions = $null
    if( $vdisk )
    {
        $bootDetails = $bootTimes[ $device.Name ]
        $fields += @{
            'Disk Name' = $vdisk.Name
            'Store Name' = $vdisk.StoreName
            'Store Free Space (GB)' = [math]::Round(( Get-PvsStoreFreeSpace -StoreId $vDisk.StoreId -ServerName $env:COMPUTERNAME ) / 1KB , 1)
            'Disk Description' = $vdisk.Description
            'Uptime (days)' = $(if( $bootDetails ) { $bootDetails.UpTime })
            'Boot time (s)' = $(if( $bootDetails ) { $bootDetails.BootTime })
            'Cache Type' = $cacheTypes[$vdisk.WriteCacheType]
            'Disk Size (GB)' = ([math]::Round( $vdisk.DiskSize / 1GB , 2 ))
            'Write Cache Size (MB)' = $vdisk.WriteCacheSize }

        $versions = $diskVersions[ $vdisk.DiskLocatorId ] ## Get-PvsDiskVersion -DiskLocatorId $vdisk.DiskLocatorId ## 
            
        if( $versions )
        {
            ## Now get latest production version of this vdisk
            $override = $versions | Where-Object { $_.Access -eq 3 } 
            $vdiskFile = $null
            $latestProduction = $versions | Where-Object { $_.Access -eq 0 } | Sort Version -Descending | Select -First 1 
            if( $latestProduction )
            {
                $vdiskFile = $latestProduction.DiskFileName
                $latestProductionVersion = $latestProduction.Version
            }
            else
            {
                $latestProductionVersion = $null
            }
            if( $override )
            {
                $bootVersion = $override.Version
                $vdiskFile = $override.DiskFileName
            }
            else
            {
                ## Access: Read-only access of the Disk Version. Values are: 0 (Production), 1 (Maintenance), 2 (MaintenanceHighestVersion), 3 (Override), 4 (Merge), 5 (MergeMaintenance), 6 (MergeTest), and 7 (Test) Min=0, Max=7, Default=0
                $bootVersion = $latestProductionVersion
            }
            if( $vdiskFile)
            {
                $vdiskFile = Join-Path $stores[ $vdisk.StoreName ] $vdiskFile
                if( ( Test-Path $vdiskFile -ErrorAction SilentlyContinue ) )
                {
                    $fields += @{ 'vDisk Size (GB)' = [math]::Round( (Get-ItemProperty -Path $vdiskFile).Length / 1GB ) }
                }
                else
                {
                    Write-Warning "Could not find disk `"$vdiskFile`" for $($device.name)"
                }
            }
            if( $latestProductionVersion -eq $null -and $override )
            {
                ## No production version, only an override so this must be the latest production version
                $latestProductionVersion = $override.Version
            }
            $fields += @{
                'Override Version' = $( if( $override ) { $bootVersion } else { $null } ) 
                'Vdisk Latest Version' = $latestProductionVersion
                'Correct Boot Version' = $(if( $override ) { $bootVersion } else { $latestProductionVersion } )
                'Latest Version Description' = $( $versions | Where-Object { $_.Version -eq $latestProductionVersion } | Select -ExpandProperty Description )  
            }      
        }
        else
        {
            Write-Output "Failed to get vdisk versions for id $($vdisk.DiskLocatorId) for $($device.Name):$($error[0])"
        }
        $fields.Add( 'Vdisk Production Version' ,$bootVersion )
    }
    else
    {
        Write-Output "Failed to get vdisk for device id $($device.DeviceId) device $($device.Name)"
    }
        
    $deviceInfo = $deviceInfos[ $device.DeviceId ]
    if( $deviceInfo )
    {
        $fields.Add( 'Disk Access' , $accessTypes[ $deviceInfo.DiskVersionAccess ] )
        $fields.Add( 'Booted Off' , $deviceInfo.ServerName )
        ##$fields.Add( 'Device IP' , $deviceInfo.IP )
        if( ! [string]::IsNullOrEmpty( $deviceInfo.Status ) )
        {
            $fields.Add( 'Retries' , ($deviceInfo.Status -split ',')[0] -as [int] ) ## second value is supposedly RAM cache used percent but I've not seen it set
        }
        if( $device.Active )
        {
            ## Check if booting off the disk we should be as previous info is what is assigned, not what is necessarily being used (e.g. vdisk changed for device whilst it is booted)
            $bootedDiskName = (( $diskVersions[ $deviceInfo.DiskLocatorId ] | Select -First 1 | Select -ExpandProperty Name ) -split '\.')[0]
            $fields.Add( 'Booted Disk Version' , $deviceInfo.DiskVersion )
            if( $bootVersion -ge 0 )
            {
                Write-Verbose "$($device.Name) booted off $bootedDiskName, disk configured $($vDisk.Name)"
                $fields.Add( 'Booted off correct version' , ( $bootVersion -eq $deviceInfo.DiskVersion -and $bootedDiskName -eq $vdisk.Name ) )
                $fields.Add( 'Booted off assigned vdisk' , ( $bootedDiskName -eq $vdisk.Name ) )
                if( ! $fields.'Booted off assigned vdisk' )
                {
                    $fields.Add( 'Booted Off Disk' , $bootedDiskName )
                    $bootedOffWrongDisk++
                }
                elseif( ! $fields.'Booted off correct version' )
                {
                    $bootedOffWrongVersion++
                }
            }
        }
        if( $versions )
        {
            try
            {
                $fields.Add( 'Disk Version Created' , (Get-Date ( $versions | Where-Object { $_.Version -eq $deviceInfo.DiskVersion } | select -ExpandProperty CreateDate ) -Format G))
            }
            catch
            {}
        }
        ##$fields.'Boot Time' = (Get-Date ([Management.ManagementDateTimeConverter]::ToDateTime( ( Get-WmiObject -class Win32_OperatingSystem -ComputerName $device.Name | Select -ExpandProperty LastBootUpTime ))) -Format G)
    }
    else
    {
        Write-Warning "Failed to get PVS device info for id $($device.DeviceId) device $($device.Name)"
    }
    if( $showAll -or ! $fields.'Booted off correct version' -or ! $fields.'Booted off assigned vdisk' )
    {
        [pscustomobject]$fields
    }
})

[string]$header = $(if( $showAll )
{
    $version = Get-PvsVersion | select -ExpandProperty MapiVersion
    $uptime = [math]::Round( (New-TimeSpan -Start ([Management.ManagementDateTimeConverter]::ToDateTime( ( Get-WmiObject -class Win32_OperatingSystem | Select -ExpandProperty LastBootUpTime ))) -End ([datetime]::Now)).TotalDays , 1 )
    "Found $($results.Count) active devices"
    if( $pattern )
    {
        "matching pattern `"$pattern`""
    }
    "- server uptime $uptime days, PVS version $version"
}
else
{
    "Out of $count active devices, $bootedOffWrongDisk are booted off the wrong disk and $bootedOffWrongVersion are booted off the wrong disk version"
})

$header

[string[]]$properties = if( $showAll )
{
    @( 'Device Name', 'Site' , 'Collection' , 'Disk Name' , 'Store Name','Disk Size (GB)','Write Cache Size (MB)' , 'Store Free Space (GB)' , 'Disk Access','Disk Version Created','Retries' , 'Uptime (days)' , 'Boot Time (s)' )
}
else
{
    @( 'Device Name', 'Disk Name'  ,'Booted Off Disk' ,'Store Name','Disk Size (GB)','Write Cache Size (MB)' , 'Booted Disk Version','Correct Boot Version' ,'Booted off correct version' , 'Booted off assigned vdisk','Disk Access','Disk Version Created','Retries' )
} 

$results | Sort-Object -Property 'Device Name' | Format-Table -AutoSize -Property $properties

