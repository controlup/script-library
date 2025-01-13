#requires -version 3

<#
.SYNOPSIS
    Look for bad event log entries for StoreFront, check config sync if in cluster, show log file contents and check bindings and certificates

.DESCRIPTION
    https://support.citrix.com/article/CTX214759/storefront-failed-to-run-discovery

.PARAMETER daysBack
    How many days back to look for issues

.PARAMETER certificateWarningDays
    Certificates expiring within this many days are added to the warnings

.PARAMETER maximumTimeDifferenceSeconds
    Maximum time difference between clock and authoritative time source before awarning is generated

.PARAMETER webrequestTmeoutSeconds
    Timeout in seconds for web reqests

.NOTES
    Modification History:

    2023/05/31  Guy Leech  Script born
    2023/06/01  Guy Leech  First version released into the wild
    2023/06/05  Guy Leech  Added checks of time against domain and internet
    2023/06/16  Guy Leech  Added event log fetching from DDC and VDA
    2023/12/18  Guy Leech  Try catch around failing get-winevent -computer
    2024/03/07  Guy Leech  Handling of bad domain controllers improved to stop error/hanging
    2024/03/13  Guy Leech  Checking of STA URLs
    2024/03/15  Guy Leech  Optimised IIS log processing. Service checking implemented. IIS error code summary
    2024/07/22  Guy Leech  Added reading of SF XML like log files
    2024/07/25  Guy Leech  Checking app pools, certificate expiry warning parameter added, more IIS error detail
    2024/07/26  Guy Leech  Added checking of StoreFront groups and config sync
    2024/07/31  Guy Leech  Added checking of URLs. Removed repetition of bindings checking for same site. Filter out ::1 source IP for IIS logs
#>

[CmdletBinding()]

Param
(
    [double]$daysBack = 1 ,
    [decimal]$certificateWarningDays = 45 ,
    [int]$maximumTimeDifferenceSeconds = 60 ,
    [int]$webrequestTmeoutSeconds = 30
)

#region ControlUp_Script_Standards
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
#endregion ControlUp_Script_Standards

[System.Collections.Generic.List[string]]$warnings = @()
[datetime]$startAnalysisDate = [datetime]::Now.AddDays( -$daysBack )
[string]$logfileNamePattern = 'u_ex*.log'
[string]$providerName = 'Citrix Receiver for Web'
[int]$eventId = 17
$storeFrontInstallKey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | Where-Object { $_.PSObject.Properties[ 'DisplayName' ] -and $_.DisplayName -imatch '\bCitrix\b.*\bStoreFront\b' -and ( -Not $_.PSObject.Properties[ 'SystemComponent' ] -or $_.SystemComponent -ne 1 ) }
[hashtable]$driveDetails = @{}
[hashtable]$logFolderDetails = @{}
[hashtable]$logFoldersForWebSites = @{}
[hashtable]$bindingsForWebSites = @{}
[hashtable]$checkedSTAs = @{}
[hashtable]$checkedURLs = @{}
[string[]]$expectedServices = @(
    'Citrix Peer Resolution Service'
    'Citrix Cluster Join Service'
    'Citrix Configuration Replication'
    'Citrix Credential Wallet'
    'Citrix Default Domain Services'
    'Citrix Service Monitor'
    'Citrix Subscriptions Store'
    'Citrix Telemetry Service'
)
[string[]]$expectedAppPools = @(
    'Citrix Delivery Services Resources' 
    'Citrix Delivery Services Authentication' 
    'Citrix Receiver for Web' 
    'Citrix Configuration Api'
)
[string[]]$excludedIISProperties = @( 'cs-uri-query' , 'cs-method' , 'cs(User-Agent)' , 'cs(Referer)' , 'time-taken' )
[string]$logsFolder = "$($env:ProgramFiles)\Citrix\Receiver StoreFront\Admin\trace"

## Read SF log files

Function Process-LogFile
{
    Param( [string]$inputFile , [datetime]$start , [datetime]$end )

    [string]$parent = 'GuyLeech' ## doesn't matter what it is since it only ever lives in memory

    Write-Verbose "Processing $inputFile ..." 
    ## Missing parent context so add
    [xml]$wellformatted = "<$parent>" + ( Get-Content -Path $inputFile ) + "</$parent>"

    <#
    <EventID>0</EventID>
		    <Type>3</Type>
		    <SubType Name="Information">0</SubType>
		    <Level>8</Level>
		    <TimeCreated SystemTime="2016-05-25T14:59:12.8077640Z" />
		    <Source Name="Citrix.DeliveryServices.WebApplication" />
		    <Correlation ActivityID="{00000000-0000-0000-0000-000000000000}" />
		    <Execution ProcessName="w3wp" ProcessID="52168" ThreadID="243" />
    #>

    try
    {
        $wellformatted.$parent.ChildNodes.Where( { [datetime]$_.System.TimeCreated.SystemTime -ge $start -and [datetime]$_.System.TimeCreated.SystemTime -le $end } ).ForEach( `
        {
            $result = [pscustomobject]@{
                'Date'=[datetime]$_.System.TimeCreated.SystemTime
                'File' = (Split-Path $inputFile -Leaf)
                'Type'=$_.System.Type
                'Subtype'=$_.System.Subtype.Name
                'Level'=$_.System.Level
                'SourceName'=$_.System.Source.Name
                'Computer'=$_.System.Computer
                'ApplicationData'=$_.ApplicationData
            }
            ## Not all objects have all properties
            if( ( Get-Member -InputObject $_.system -Name 'Process' -Membertype Properties -ErrorAction SilentlyContinue ) )
            {
                Add-Member -InputObject $result -NotePropertyMembers @{
                    'Process' = $_.System.Process.ProcessName
                    'PID' =  $_.System.Process.ProcessID
                    'TID' = $_.System.Process.ThreadID
                }
                <#
                $result | Add-Member -MemberType NoteProperty -Name 'Process' -Value $_.System.Process.ProcessName
                $result | Add-Member -MemberType NoteProperty -Name 'PID' -Value $_.System.Process.ProcessID;`
                $result | Add-Member -MemberType NoteProperty -Name 'TID' -Value $_.System.Process.ThreadID;`
                #>
            }
            $result
        } ) | Select-Object -Property Date,File,SourceName,Type,SubType,Level,Computer,Process,PID,TID,ApplicationData ## This is the order they will be displayed in
    }
    catch {}
}

if( -Not $storeFrontInstallKey )
{
    $warnings.Add( "Cannot find installation details for Citrix StoreFront - is it installed or is this the wrong machine?" )
}
else
{
    ## Write-Verbose -Message "StoreFront application id from registry is $($storeFrontInstallKey.PSChildName)"
    Write-Output -InputObject "StoreFront installed version is $($storeFrontInstallKey.DisplayVersion)"
}

[string]$storeFrontConfigurationKeyName = 'HKLM:\SOFTWARE\Citrix\DeliveryServices'

if( $storeFrontConfigurationKey = Get-ItemProperty -Path $storeFrontConfigurationKeyName -ErrorAction SilentlyContinue )
{
    if( $storeFrontConfigurationKey.PSObject.Properties[ 'InstallDir' ] )
    {
        if( Test-Path -Path $storeFrontConfigurationKey.InstallDir )
        {
            [string]$moduleScript = [System.IO.Path]::Combine( $storeFrontConfigurationKey.InstallDir , 'Scripts' , 'ImportModules.ps1' )
            if( Test-Path -Path $moduleScript -PathType Leaf )
            {
                $null = . $moduleScript >$null 4>$null 6>$null ## noisy script so redirect stdout and Write-Host channels to oblivion
                if( -Not $? )
                {
                    $warnings.Add( "Bad status from importing StoreFront PS modules using $moduleScript" )
                }
            }
            else
            {
                $warnings.Add( "Expected StoreFront PowerShell module `"$moduleScript`" does not exist" )
            }
        }
        else
        {
            $warnings.Add( "Installation path `"$($storeFrontConfigurationKey.InstallDir)`" configured in registry does not exist" )
        }
    }
    else
    {
        $warnings.Add( "Unable to find InstallDir value in key $storeFrontConfigurationKeyName" )
    }
}
else
{
    $warnings.Add( "StoreFront configuration key $storeFrontConfigurationKeyName not found" )
}

<#
## Get Citrix Delivery Services event logs - get all because also look for good syncs
[array]$allDeliveryServicesEvents = @( Get-WinEvent -FilterHashtable @{ LogName = 'Citrix Delivery Services' ; StartTime = $startAnalysisDate } -ErrorAction SilentlyContinue | Select-Object -ExcludeProperty Message -Property *,
    ## roll failures for different users and desktop groups into a single error to keep number down without hiding completely
    @{n='Message';e={ $_.Message -replace '(No available resource found for user).*$' , '$1' }} )

[array]$deliveryServiceGroupedEvents = @( $allDeliveryServicesEvents | Group-Object -Property { $_.Level -le 3 } )
$deliveryServiceBadEvents  = @( $deliveryServiceGroupedEvents | Where-Object Name -eq $true  | Select-Object -ExpandProperty Group )
$deliveryServiceGoodEvents = @( $deliveryServiceGroupedEvents | Where-Object Name -eq $false | Select-Object -ExpandProperty Group )
#>

[System.Collections.Generic.List[object]]$deliveryServiceBadEvents  = @()
[System.Collections.Generic.List[object]]$deliveryServiceGoodEvents = @()

[array]$deliveryServiceGoodEvents = @( Get-WinEvent -FilterHashtable @{ LogName = 'Citrix Delivery Services' ; StartTime = $startAnalysisDate } -ErrorAction SilentlyContinue | Select-Object -ExcludeProperty Message -Property *,
    ## roll failures for different users and desktop groups into a single error to keep number down without hiding completely
    @{n='Message';e={ $_.Message -replace '(No available resource found for user).*$' , '$1' }} | ForEach-Object `
    {
        if( $_.Level -le 3 )
        {
            $deliveryServiceBadEvents.Add( $_ )
        }
        else
        {
            $_ ## output to pipeline so goes to deliveryServiceGoodEvents array
        }
    })

if( $null -ne $deliveryServiceBadEvents -and $deliveryServiceBadEvents.Count -gt 0 )
{
    $warnings.Add( "$($deliveryServiceBadEvents.Count) Citrix Event Log Reported Issues:" )

    $warnings += @( $deliveryServiceBadEvents | Group-Object -Property Message | Sort-Object -Property count -Descending | Select-Object -Property count,
        @{name='Last Occurrence';expression={$_.Group.TimeCreated|Sort-Object -Descending|Select-Object -first 1}},
        @{name='Level';expression={$_.Group[0].LevelDisplayName}},
        @{name='Id';expression={$_.Group[0].Id}},
        @{name='Provider';expression={$_.Group[0].ProviderName}},
        @{name='Error' ; expression = { $_.Name -replace "`r?`n" , ' ' -replace '\s{2,}' ,' '}} | Format-Table -AutoSize | Out-String )
}
else
{
    Write-Output -InputObject "No problems found in the Delivery Services Event Log since $($startAnalysisDate.ToString( 'G' ))"
}

##[array]$events = @( Get-WinEvent -FilterHashtable @{ ProviderName = $providerName ; Id = $eventId } -ErrorAction SilentlyContinue )
[array]$events = @( $deliveryServiceBadEvents | Where-Object { $_.ProviderName -ieq $providerName -and $_.Id -eq $eventId })
$OS = $null
$OS = Get-CimInstance -ClassName Win32_OperatingSystem

if( $OS )
{
    Write-Output -InputObject "OS last booted $($OS.LastBootUpTime.ToString('G'))"
}
else
{
    $warnings.Add( "Failed to get OS information" )
}

if( $null -eq $events -or $events.Count -eq 0 )
{
    $providerDetail = Get-WinEvent -ListProvider $providerName -ErrorAction SilentlyContinue
    $oldestEvent = $null
    if( -Not $providerDetail )
    {
        if( -Not $storeFrontInstallKey )
        {
            Throw "Cannot find installation details for Citrix StoreFront - is it installed or is this the wrong machine?"
        }
    }
    else
    {
        $oldestEvent = Get-WinEvent -LogName $providerDetail.LogLinks.LogName -Oldest -MaxEvents 1
    }
    Write-Output -InputObject "No instances of event id $eventId for provider `"$providerName`" found suggesting SSL is ok or server not used"
    if( $oldestEvent )
    {
        Write-Output -InputObject "Oldest event in containing event log is $($oldestEvent.TimeCreated.ToString( 'G' ))"
    }
}
<#
else
{
    Write-Output -InputObject "Most recent of $($events.Count) occurrences of event id $eventId for provider `"$providerName`" was $($events[0].TimeCreated.ToString('G')) - see https://support.citrix.com/article/CTX214759/storefront-failed-to-run-discovery"
}
#>

[array]$citrixServices = @( Get-CimInstance -ClassName win32_service -Filter "Name like 'Citrix%'" )

## have to truncate message because different message depending on whether 1st, 2nd, 3rd, etc failure
[array]$serviceProblems = @( Get-WinEvent -FilterHashtable @{ ProviderName = 'Service Control Manager' ; Level = 1,2,3 ; StartTime = [datetime]::Now.AddDays( -$daysBack ) } -ErrorAction SilentlyContinue |Where-Object Message -match Citrix|Select-object -Property *,@{name='FirstSentence';expression = { $_.Message -replace '\.\s+.*$' }}|group-object firstsentence|select-object count,name,@{name='Most Recent Occurrence';expression={$_.Group.TimeCreated|Sort-Object -Descending|Select-Object -first 1}}|Sort-Object -Property count -Descending )

if( $null -ne $serviceProblems -and $serviceProblems.Count -gt 0 )
{
    $warnings.Add( "Citrix Service Problems" )
    $warnings += @( $serviceProblems | Format-Table -AutoSize | Out-String )
}
else
{
    Write-Output -InputObject "No Citrix service problems found in the event log since $($startAnalysisDate.ToString( 'G' ))"
}

ForEach( $expectedService in $expectedServices )
{
    $thisService = $citrixServices | Where-Object DisplayName -ieq $expectedService
    if( $null -eq $thisService  )
    {
        $warnings.Add( "Service `"$expectedService`" not present" )
    }
    else ## found service so check status
    {
        if( $thisService.StartMode -ieq 'Auto' -and $thisService.State -ine 'Running' )
        {
            $warnings.Add( "Service `"$($thisservice.DisplayName)`" is not running but its start is set to $($thisService.StartMode)" )
        }
        elseif( $thisService.State -ieq 'Running' -and $thisService.Status -ine 'OK' )
        {
            $warnings.Add( "Service `"$($thisservice.DisplayName)`" is running but status is $($thisService.Status)" )
        }
    }
}

$stores = $null
if( Get-Command -Name Get-STFStoreService -ErrorAction SilentlyContinue )
{
    $stores = @( Get-STFStoreService )
    if( $null -eq $stores -or $stores.Count -eq 0 )
    {
         $warnings.Add( "Failed to get store details from StoreFront via PowerShell" )
    }
    else
    {
        Write-Verbose -Message "Got $($stores.Count) stores"
        $stores | Select-Object -Property VirtualPath,ConfigurationFile,@{n='Config Last Written';e={ (Get-ItemProperty -Path $_.ConfigurationFile).LastWriteTime }} | Sort-Object -Property VirtualPath | Format-Table -AutoSize
    }
}
else
{
    $warnings.Add( "Unable to find StoreFront cmdlet Get-STFStoreService" )
}

Import-Module -Name WebAdministration,IISAdministration -Verbose:$false -Debug:$false

$webSites = @( $stores | Select-Object -ExpandProperty WebApplication | Select-Object -ExpandProperty WebSite | Group-Object -Property Name )

Write-Verbose -Message "Stores are using $($webSites.count) web sites"

[string]$matchIISregex = '^(\d|#Fields)'
[string]$replaceIISregex = '^#Fields: '
[array]$includeProperties = @( '*',@{name='duration';expression = { $_.'time-taken' -as [int] }}, @{name='datetime';expression={([datetime]"$($_.date) $($_.time)").ToLocalTime()}} )
[scriptblock]$IISFilterScript = { $_.'s-port' -as [int] -ge 0 -and $_.'c-ip' -ne '127.0.0.1' -and $_.'c-ip' -ne '::1' -and $_.'cs-method' -ieq 'GET' }
[int]$missingIISlogfiles = 0
[int]$IISlogfilesCheckedFor = 0

Write-Verbose -Message "$([datetime]::Now.ToString( 'G')): processing $($websites.count) web sites"

ForEach( $website in $webSites )
{
    $logfilefolder = $null
    $websiteProperties = $null
    [string]$webSiteName = $website.Name
    Write-Verbose -Message "Processing logs from web site $webSiteName"
    $websiteProperties = Get-ItemProperty -Path "IIS:\Sites\$webSiteName" ## .WebApplication.WebSite.Name)"
    $allGetRequests = New-Object -TypeName System.Collections.Generic.List[object]
    [string]$webSiteAndId = "$($webSiteName):$($websiteProperties.Id)"

    if( $websiteProperties.State -ine 'Started' )
    {
        $warnings.Add( "Web site `"$webSiteName`" is not running" )
    }
    if( $websiteProperties.logFile -and $websiteProperties.logFile.enabled )
    {
        $logfilefolder = Join-Path -Path ([environment]::ExpandEnvironmentVariables( $websiteProperties.logFile.directory )) -ChildPath "W3SVC$($websiteProperties.Id)"
        if( -Not ( Test-Path -Path $logfilefolder ) )
        {
            $logfilefolder = $null
        }
        else
        {
            if( -Not $logFoldersForWebSites.ContainsKey( $webSiteAndId ) )
            {
                ## use this when we examine the stores later so we can get the correct log file for that web site & id
                $logFoldersForWebSites.Add( $webSiteAndId , $logfilefolder )
            }
            ## get latest log file and total size
            if( -Not $logFolderDetails.ContainsKey( $logfilefolder ) )
            {
                [string]$logfolderVolumeLetter = $logfilefolder -replace '^(\w).*$' , '$1'
                if( -Not $driveDetails.ContainsKey( $logfolderVolumeLetter ) )
                {
                    $latestLogFile = Get-ChildItem -Path $logfilefolder -File -Force -Recurse -Depth 2 -Filter $logfileNamePattern -ErrorAction SilentlyContinue | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1
                    [int64]$logfilesTotalSize = (Get-ChildItem -Path $logfilefolder -File -Force -Recurse -Depth 2 -ErrorAction SilentlyContinue | Measure-Object -Sum Length).Sum 
        
                    Write-Output -InputObject "IIS logging is enabled to folder $logfilefolder. Latest log file is $($latestLogFile.Name), last written $($latestLogFile.LastWriteTime.ToString('G')) ($([math]::Round( ([datetime]::Now - $latestLogFile.LastWriteTime).TotalDays , 2)) days ago)"

                    [string]$logfolderVolumeLetter = $logfilefolder -replace '^(\w).*$' , '$1'
                    $logFolderVolume = $null
                    $logFolderVolume = Get-Volume -DriveLetter $logfolderVolumeLetter
                    [string]$message = "IIS log files consuming $([math]::Round( $logfilesTotalSize / 1GB , 2 ))GB on $($logfolderVolumeLetter):"
                    if( $logFolderVolume )
                    {
                        $message += ", free space is $([math]::Round( $logFolderVolume.SizeRemaining / 1GB , 1))GB"
                    }
                    else
                    {
                       $warnings.Add( "Failed to get volume details for volume $logfolderVolumeLetter containing log folder $logfilefolder" )
                    }
                    Write-Output -InputObject $message

                    For( [int]$dayBack = $daysBack ; $daysBack -ge 0 ; $daysBack-- ) 
                    {
                        [datetime]$logFileDate = [datetime]::Today.AddDays( -$daysBack )
                        [string]$logFile = Join-Path -Path $logfilefolder -ChildPath ( $logfileNamePattern -replace '\*' , $logFileDate.ToString('yyMMdd'))
                        Write-Verbose -Message "Processing IIS log file $logfile"
                        $IISlogfilesCheckedFor++
                        if( Test-Path -Path $logFile -PathType Leaf )
                        {
                            $requests = $null
                            try
                            {
                                ## if you change this line, copy to Get-Content one in the catch block - not assigning file contents to variable for speed (could assign to variables & use those)
                                $requests = @( [System.IO.File]::ReadLines( $logfile ) -match $matchIISregex -replace $replaceIISregex  | ConvertFrom-Csv -Delimiter ' ' | Where-Object -FilterScript $IISFilterScript | Select-Object -Property $includeProperties -ExcludeProperty $excludedIISProperties )
                            }
                            catch ## could catch specific "file in use" so we don't retry in other error conditions
                            {
                                try
                                {
                                    Write-Verbose -Message "Error from ReadLines switching to Get-Content on $logfile - $_"
                                    ## copy of ReadLines line from try block
                                    $requests = @( (Get-Content -Path $logfile ) -match $matchIISregex -replace $replaceIISregex  | ConvertFrom-Csv -Delimiter ' ' | Where-Object -FilterScript $IISFilterScript | Select-Object -Property $includeProperties -ExcludeProperty $excludedIISProperties )
                                }
                                catch
                                {
                                    $warnings.Add( "Failed to read IIS log file $logfile - $_" )
                                }
                            }
                            if( $null -ne $requests -and $requests.Count -gt 0 )
                            {
                                $allGetRequests.AddRange( $requests )
                            }
                        }
                        else
                        {
                            $missingIISlogfiles++
                        }
                    }
                }
                ## else already processed this driver / folder
                $logFolderDetails.Add( $logfilefolder , $allGetRequests )
            }
            ## else already processed this folder
            $driveDetails.Add( $logfolderVolumeLetter , $logFolderVolume )
        }
    }
    else
    {
        Write-Output -InputObject "Logging not enabled for web site `"$webSiteName`""
    }
}

if( $missingIISlogfiles -gt 0 )
{
    $warnings.Add( "$missingIISlogfiles / $IISlogfilesCheckedFor IIS log files do not exist" )
}

[array]$appPools = @( Get-IISAppPool | Where-Object -Property Name -match '^Citrix ' )
[string[]]$appPoolNames = @( $appPools | Select-Object -ExpandProperty Name ) ## make it easier to find missing ones

if( $appPools.Count -ne $expectedAppPools.Count )
{
    $missingAppPools = @( ForEach( $appPool in $expectedAppPools ) 
                        {
                            if( $appPoolNames -notcontains $appPool )
                            {
                                $appPool
                            }
                        })
    $warnings.Add( "Expected $($expectedAppPools.Count) Citrix app pools but only found $($appPools.Count). Missing `"$($missingAppPools -join '" , "')`"" )
}

[array]$notStartedAppPools = @( $appPools | Where-Object State -ine 'Started' )
if( $notStartedAppPools.Count -gt 0 )
{
    $warnings.Add( "$($notStartedAppPools.Count) app pools are not started which are `"$(($notStartedAppPools|Select-Object -ExpandProperty Name) -join '" , "')`"" )
}

[array]$IISevents = @( Get-WinEvent -FilterHashtable @{ ProviderName = ( Get-WinEvent -ListProvider *-IIS-* | Select-Object -ExpandProperty Name ) ; Level = 1,2,3 ;StartTime = $startAnalysisDate } -ErrorAction SilentlyContinue | Select-Object -Property *,
    ## group "The Application Host Helper Service encountered an error trying to delete the history directory 'C:\inetpub\history\CFGHISTORY_0000000009'."
   @{ name='Message';expression={ $_.Message -replace '_\d{8,}' , '??' }} -ExcludeProperty Message )

if( $null -ne $IISevents -and $IISevents.Count -gt 0 )
{
    $warnings.Add( "$($IISevents.Count) IIS warning/error events since $($startAnalysisDate.ToString( 'G' ))" )
    $warnings += @( $IISevents | Group-Object -Property Message | Sort-Object -Property count -Descending | Select-Object -Property count,
        @{name='Last Occurrence';expression={$_.Group.TimeCreated | Sort-Object -Descending|Select-Object -first 1}},
        @{name='Level';expression={$_.Group[0].LevelDisplayName}},
        @{name='Id';expression={$_.Group[0].Id}},
        @{name='Provider';expression={$_.Group[0].ProviderName}},
        @{name='Error' ; expression = { $_.Name -replace "`r?`n" , ' ' -replace '\s{2,}' ,' '}}  | Format-Table -AutoSize | Out-String )
}

Write-Verbose -Message "$([datetime]::Now.ToString( 'G')): finished processing web sites, got $($logFolderDetails.count) log folders"

$domain = $null
$domain = [DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
if( $domain )
{
    ## times from domain controllers are UTC
    [int]$counter = 0
    [int]$insync = 0
    [int]$contacted = 0
    [string]$logonDC = $env:LOGONSERVER -replace '^\\\\'
    $timeOnLogonServer = $null
    [double]$worstTimeDifference = 0
    $worstTime = $null
    
    if( $null -eq $domain.DomainControllers-or $domain.DomainControllers.Count -eq 0 )
    {
        $warnings.Add( "No domain controllers returned for domain $($domain.Name) so cannot check for time skew" )
    }
    else
    {
        ForEach( $domainController in $domain.DomainControllers )
        {
            $counter++
            Write-Verbose -Message "Checking DC $($domainController.Name) $counter / $($domain.DomainControllers.Count)"
            $dcTime = $null
            $dcTime = $domainController.CurrentTime
            if( $null -eq $dcTime )
            {
                $warnings.Add( "Failed to get current time from domain controller $($domainController.Name), IP address $($domainController.IPAddress)" )
            }
            else
            {
                $contacted++
                $timeDifferenceSeconds = [math]::Abs( ([datetime]::now.ToUniversalTime() - $domainController.CurrentTime ).TotalSeconds )
                if( $timeDifferenceSeconds -ge $maximumTimeDifferenceSeconds )
                {
                    if( $timeDifferenceSeconds -gt $worstTimeDifference )
                    {
                        $worstTime = $domainController
                        $worstTimeDifference = $timeDifferenceSeconds
                    }
                }
                else
                {
                    $insync++
                }
            }
        }

        if( $contacted -ne $domain.DomainControllers.Count )
        {
            if( $contacted -eq 0 )
            {
                $warnings.Add( "None of the $($domain.DomainControllers.Count) domain controllers could be contacted" )
            }
            else
            {
                $warnings.Add( "Only $contacted of the $($domain.DomainControllers.Count) domain controllers could be contacted" )
            }
        }
        if( $insync -ne $domain.DomainControllers.Count -and $null -ne $worstTime )  ## don't report time if no DCs could be contacted since there are bigger issues
        {
            $warnings.Add( "Time is out with $( ($domain.DomainControllers.Count - $insync) / $domain.DomainControllers.Count * 100)% of $($domain.DomainControllers.Count) domain controllers - worst is $worstTimeDifference seconds on $worstTime" )
        }
    }
    [string]$proxy = $null
    [hashtable]$proxyParameter = @{ }
    [bool]$retry = $true
    $nowUTC = $null
    $utcTime = $null
    [string]$timeURI = 'https://timeapi.io/api/Time/current/zone?timeZone=UTC'

    While( $retry ) 
    {
        try
        {
            $nowUTC = [datetime]::Now.ToUniversalTime()
            $utcTime = Invoke-RestMethod -Uri $timeURI -TimeoutSec $webrequestTmeoutSeconds -UseBasicParsing @proxyParameter
            $retry = $false
        }
        catch
        {
            ## see if there is a proxy at the system level and try that
            if( -Not $proxy -and ( [string[]]$proxies = @( (netsh.exe winhttp show proxy | Select-String -SimpleMatch 'Proxy Server(s)') -split ':' | Select-Object -Skip 1 ) ) -and $proxies.Count )
            {
                $proxy = $proxies[0]
                if( $proxies[0] -notmatch 'https?:' )
                {
                    $proxy = "http://" + $proxies[0].Trim()
                }
                if( $proxies.Count -gt 1 -and ! [string]::IsNullOrEmpty( $proxies[1] ) )
                {
                    $proxy =  "$($proxy):$($proxies[1])"
                }
                Write-Verbose -Message "Trying proxy $proxy"
                $proxyParameter.Add( 'Proxy' , $proxy )
                $proxyParameter.Add( 'ProxyUseDefaultCredentials' , $true ) ## we don't have any others we can pass
            }
        }
    }
    
    if( $utcTime )
    {
        [datetime]$internetTimeUTC = $utcTime.dateTime
        [int]$secondsDifference = [math]::Abs( ($internetTimeUTC - $nowUTC).TotalSeconds )
        Write-Verbose -Message "UTC time locally $($nowUTC.ToString('G')) versus internet UTC $($internetTimeUTC.ToString('G')) - difference $secondsDifference"
        if( $secondsDifference -ge $maximumTimeDifferenceSeconds )
        {
            $warnings.Add( "Time locally is $secondsDifference seconds adrift compared with time from internet" )
        }
    }
    else
    {
        $warnings.Add( "Failed to get time from $timeURI" )
    }
}
else
{
    $warnings.Add( "Failed to get current domain" )
}

[hashtable]$checkedControllers = @{}

$ddcEvents = New-Object -TypeName System.Collections.Generic.List[object]

$configSyncDetails = $null
$storeFrontGroupDetails = $null
$storeFrontGroupDetails = Get-STFServerGroup
if( $null -eq $storeFrontGroupDetails )
{
    $warnings.Add( "Failed to get details for StoreFront server group" )
}
else
{
    Write-Output -InputObject "Got $($storeFrontGroupDetails.ClusterMembers.Count) members in StoreFront server group"
}
if( $storeFrontGroupDetails.ClusterMembers.Count -gt 1 )
{
    ## check connectivity to each server in cluster
     ForEach( $member in $storeFrontGroupDetails.HostBaseUrl )
     {
        Write-Verbose -Message "Testing $($member.DnsSafeHost):$($member.port)"
        $networkConnectionResult = $null
        $networkConnectionResult = Test-NetConnection -ComputerName $member.DnsSafeHost -Port $member.port -InformationLevel Detailed -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        if( -Not $networkConnectionResult -or -Not $networkConnectionResult.TcpTestSucceeded )
        {
            $warnings.Add( "Failed to contact StoreFront cluster member at $($member.DnsSafeHost):$($member.port)" )
        }
    }
    $configSyncDetails = Get-DSConfigurationReplicationState
    if( $null -ne $configSyncDetails )
    {
        if( $configSyncDetails.LastUpdateStatus -ine 'completed' )
        {
            if( $null -ne $configSyncDetails.ServerStates -and $configSyncDetails.ServerStates.Count -gt 0 )
            {
                ForEach( $state in $configSyncDetails.ServerStates )
                {
                    if( $state.LastUpdateStatus -ine 'completed' )
                    {
                        $warnings.Add( "Last configuration sync with $($state.Servername) was not successful at $($state.LastStartTime.ToString('G')), error was `"$($state.LastError)`"" )
                    }
                }
            }
            else ##no detailed error
            {
                $warnings.Add( "Last configuration sync was not successful at $($configSyncDetails.LastStartTime.ToString('G')), error was `"$($configSyncDetails.LastError)`"" )
            }
        }
        else
        {
            Write-Output -InputObject "Last configuration sync was successful at $($configSyncDetails.LastStartTime.ToString('G')) taking $([math]::round( ($configSyncDetails.LastEndTime - $configSyncDetails.LastStartTime).TotalSeconds,2 )) seconds)"
        }
    }
    else
    {
        $warnings.Add( "Failed to get StoreFront replication state" )
    }
    [array]$installedVersions = @( ForEach( $server in $storeFrontGroupDetails.ClusterMembers )
    {
        $version = $null
        $versionError = $null
        $version = Get-DSConfigurationReplicationVersion -Hostname $server.hostname -ErrorVariable versionError
        [pscustomobject]@{
            Server = $server.hostname
            Version = $version
            Error = $versionError
        }
    })
    [array]$differentVersions = @( $installedVersions | Group-Object -Property Version )
    if( $null -ne $differentVersions -and $differentVersions.Count -gt 1 )
    {
       [int]$badCount = 0
       $failures = $differentVersions | Where-Object { [string]::IsNullOrEmpty( $_.Name ) } ## where version number was null
       if( $null -ne $failures )
       {
            $badCount++
            ForEach( $failure in $failures.Group )
            {
                $warnings.Add( "Failed to get StoreFront version from $($failure.Server), error `"$($failure.Error)`"" )
            }
        }
        if( $differentVersions.Count -ge $badCount )
        {
            $warnings.Add( "Got $($differentVersions.Count - $badCount) different StoreFront versions in cluster across $( ($differentVersions | Measure-Object -Property Count -Sum).Sum - $badCount) servers - $( ($differentVersions | Where-object { -Not [string]::IsNullOrEmpty( $_.Name ) } ).Name -join ' , ')" )
        }
    }
}
else
{
    $warnings.Add( "StoreFront not clustered" )
}
    
ForEach( $store in $stores )
{
    $webSiteAndId = "$($store.WebApplication.WebSite.Name):$($store.WebApplication.WebSite.Id)"
    $logfilefolder = $logFoldersForWebSites[ $webSiteAndId ]
    if( $null -ne $logfilefolder -and ($allGetRequests = $logFolderDetails[ $logfilefolder ] ))
    {
        [array]$requests = ( $allGetRequests | Where-Object 'cs-uri-stem' -match $store.VirtualPath )

        if( $null -eq $requests -or $requests.Count -eq 0 )
        {
            Write-Output -InputObject "No IIS requests matching virtual path $($store.VirtualPath) found in log file $($latestLogFile.Name)"
        }
        else
        {
            [array]$sourceIPs = @( $requests | Group-Object -Property 'c-ip' )
            [array]$errorRequests = @( $requests | Where-Object sc-status -ge 400 )

            Write-Output -InputObject "Got $($requests.Count) IIS requests matching virtual path $($store.VirtualPath) from all log files from $($sourceIPs.Count) different IPs and $($errorRequests.Count) error responses"
            if( $null -ne $errorRequests -and $errorRequests.Count -gt 0 )
            {
                ## errors are ordered so we can get date of first and last from those elements in the group array for each error
                [array]$groupedErrors = @( $errorRequests | Group-Object -Property 'sc-status' | Sort-Object -Property Count -Descending | Select-Object -Property Count,
                    @{name = 'Error' ;expression = {$_.Name}} ,
                    @{name = 'First Occurence' ; expression = { $_.Group[0].datetime.ToString('G') }} ,
                    @{name = 'Last Occurence' ; expression = { if( $_.group.Count -gt 1 ) { $_.Group[-1].datetime.ToString('G') } else { '-' } }} ,
                    @{name = 'Duration (h)' ; expression = { if( $_.group.Count -gt 1 ) { [math]::round( ( $_.Group[-1].datetime - $_.Group[0].datetime ).TotalHours , 1 ) } else { '-' } }} )
                $warnings.Add( "Error summary for $($errorRequests.Count) errors for $($store.VirtualPath):" )
                $warnings += @( $groupedErrors | Format-Table -AutoSize | Out-String  )
            }
        }
    }
    else
    {
        $warnings.Add( "No matching IIS logs for virtual path $($store.VirtualPath)" )
    }

    ## check URLs reachable - not the greatest of tests since we are local but better than nothing
    $store.Routing.ExternalEndpoints | Where-Object Id -ieq 'WebUI' | Select-Object -ExpandProperty Url -Unique -PipelineVariable storeURL | ForEach-Object `
    {
        $response = $null
        $exception = $null
        if( -Not $checkedURLs.ContainsKey( $storeURL ) )
        {
            [string]$warning = $null
            Write-Verbose -Message "Checking URL $storeURL for store $($store.Name)"
            try
            {
                $response = Invoke-WebRequest -UseBasicParsing -Uri $storeURL -TimeoutSec $webrequestTmeoutSeconds
            }
            catch
            {
                $exception = $_.Exception
            }
            if( $null -ne $exception )
            {
                $warning = "Error from URL $storeURL for store $($store.Name) - $($exception|Select-Object -ExpandProperty Message)" 
            }
            elseif( $null -eq $response )
            {
                $warning = "No response or error from URL $storeURL for store $($store.Name)"
            }
            elseif( $response.StatusCode -ge 400 )
            {
                $warning = "Error from URL $storeURL for store $($store.Name) - $($response.StatusCode)"
            }
            ## else good but should we check for some specific content in case redirectedd to an internal page not found page or similar ?
            if( -Not [string]::IsNullOrEmpty( $warning ) )
            {
                $warnings.Add( $warning )
            }
            $checkedURLs.Add( $storeURL , $warning ) ## add regardless of success or fail since we test with ContainsKey so that we only test once
        }
        ## else previously reported so don't repeat
    }

    ## check controllers for stores can be resolved and ports connected to
    [int]$controllersChecked = 0
    ForEach( $farm in $store.FarmsConfiguration.Farms )
    {
        ## check resolution
        ## check port connectivity
        ForEach( $server in $farm.servers )
        {
            [string]$serverAndPort = "$($server):$($farm.Port)"
            if( -Not ( $networkConnectionResult = $checkedControllers[ $serverAndPort ] ) )
            {
                Write-Verbose -Message "Testing $server on port $($farm.port) for store $($store.name)"
                $networkConnectionResult = $null
                $networkConnectionResult = Test-NetConnection -ComputerName $server -Port $farm.port -InformationLevel Detailed -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                $controllersChecked++
                $checkedControllers.Add( $serverAndPort , $networkConnectionResult )
                if( $daysBack -gt 0 )
                {
                    [datetime]$startEventSearchTime = [datetime]::Now.AddDays( -$daysBack )
                    Write-Verbose -Message "Checking $server for events from $server from $($startEventSearchTime.ToString('G'))"
                    $events = @()
                    try
                    {
                        $events = @( Get-WinEvent -ComputerName $server -FilterHashtable @{ LogName = 'Application' ; StartTime = $startEventSearchTime } -ErrorAction SilentlyContinue )
                        $ddcEvents += $events
                        Write-Verbose -Message "Got $($events.Count) events from $server since $($startEventSearchTime.ToString('G'))"
                    }
                    catch
                    {
                        Write-Verbose -Message "Exception getting events from $server - $_"
                    }
                }
            }
            if( -Not $networkConnectionResult -or -Not $networkConnectionResult.TcpTestSucceeded )
            {
                [string]$pingResult = 'unknown'
                if( $networkConnectionResult )
                {
                    if( $networkConnectionResult.PingSucceeded )
                    {
                        $pingResult = 'worked'
                    }
                    else
                    {
                        $pingResult = 'failed'
                    }
                }
                $warnings.Add( "Unable to contact delivery controller $server (resolved to $($networkConnectionResult.ResolvedAddresses -join ' ')) on port $($farm.port), transport $($farm.TransportType). Ping $pingResult" )
                if( -Not ( $server -as [ipaddress] ))
                {
                    $resolvedTo = $null
                    if( $networkConnectionResult )
                    {
                        $resolvedTo = $networkConnectionResult.AllNameResolutionResults | Select-Object -ExpandProperty IPAddress
                    }
                    else
                    {
                        $resolvedTo = Resolve-DnsName -Name $server -DnsOnly -ErrorAction SilentlyContinue | Select-Object -ExpandProperty IPAddress
                    }
                    if( $null -ne $resolvedTo )
                    {
                        Write-Output -InputObject "$server resolved to $($resolvedTo -join ' ')"
                    }
                    else
                    {
                        $warnings.Add( "Unable to DNS resolve $server" )
                    }
                }
            }
        }

        ## check STA if present
        if( $store.Gateways -and $store.Gateways.Count -gt 0 )
        {
            [string[]]$staURLs = @( $store.Gateways.SecureTicketAuthorityUrls | Select-Object -ExpandProperty StaURL -Unique )
            Write-Verbose -Message "Got $($staURLs.Count) STA URLS for store `"$($store.Name)`""
            ForEach( $staURl in $staURLs )
            {
                $response = $null
                $exception = $null
                $warning = $checkedSTAs[ $staURL ] ## don''t make [string] as $null becomes empty string which is special to us (tested previously and was ok)
                if( $null -eq $warning )
                {
                    Write-Verbose -Message "Checking STA URL $staurl for store $($store.Name)"
                    try
                    {
                        $response = Invoke-WebRequest -UseBasicParsing -Uri $staURl -TimeoutSec $webrequestTmeoutSeconds
                    }
                    catch
                    {
                        $exception = $_.Exception
                    }
                    if( $null -eq $response )
                    {
                        ## we are expecting a 406 error
                        if( -Not $exception )
                        {
                            $warning = "No response from STA URL $staURl for store $($store.Name) and no error" 
                        }
                        elseif( -Not $exception.PSobject.Properties[ 'Response' ] -or -Not $exception.Response -or -Not $exception.Response.PSobject.Properties[ 'StatusCode' ] -or $exception.Response.StatusCode -ine 'NotAcceptable' )
                        {
                            $warning = "Error from STA URL $staURl for store $($store.Name) - $($exception.Message)" 
                        }
                        else ## good response so store empty string so doesn't check again but no message output
                        {
                            $warning = ''
                        }
                        $checkedSTAs.Add( $staURl , $warning )
                    }
                    else ## good response but probably shouldn't have been
                    {
                        $warning = "No error from STA URL $staURl for store $($store.Name) but expected one" 
                    }
                    if( -Not [string]::IsNullOrEmpty( $warning ) )
                    {
                        $warnings.Add( $warning )
                    }
                }
                ## else already warned so don't repeat
            }
        }
    }
    
    if( $controllersChecked -eq 0 )
    {
        $warnings.Add( "No delivery controllers checked for store $($store.Name), check configured correctly!" )
    }

    if( $null -eq $store.WebApplication.WebSite.Bindings -or $store.WebApplication.WebSite.Bindings.Count -eq 0 )
    {
        $warnings.Add( "No bindings at all found for store $($store.Name)" )
        continue
    }

    ## TODO check dates versus when cert last changed if possible - check dates of last write to web.config for SF

    [int]$counter = 0

    Write-Output -InputObject "$($store.WebApplication.WebSite.Bindings.Count) bindings found for StoreFront store $($store.name)"

    ## only check if not already checked since at the web site level so if only 1 web site, only 1 set of bindings
    if( -Not ( $bindingsForWebSites.ContainsKey( $webSiteAndId )))
    {
        ForEach( $binding in $store.WebApplication.WebSite.Bindings )
        {
            if( $binding.HostSource -ieq 'Certificate' -and $null -eq $binding.Certificate )
            {
                $warnings.Add( "Certificate for store $($store.name) missing for binding on port $($binding.port)" )
                continue
            }
            if( $binding.port -eq 443 -and $null -eq $binding.Certificate )
            {
                $warnings.Add( "Binding is on port $($binding.port) but there is no certificate configured for store $($store.name)" )
                continue
            }
        
            if( $storeFrontCertificate = $binding.Certificate )
            {     
                $counter++

                [string]$subject = ( $storeFrontCertificate | Select-Object -ExpandProperty Subject ) -replace '^CN='

                if( [string]::IsNullOrEmpty( $subject ) )
                {
                    $warnings.Add(  "Certificate with thumbprint $($storeFrontCertificate.Thumbprint) for store $($store.name) has no subject" )
                }
                if( $storeFrontCertificate.NotBefore -gt [datetime]::Now )
                {
                    $warnings.Add(  "Certificate with thumbprint $($storeFrontCertificate.Thumbprint) not valid until $($storeFrontCertificate.NotBefore.ToString('G')) " )
                }
                if( $storeFrontCertificate.NotAfter -le [datetime]::Now )
                {
                    $warnings.Add(  "Certificate with thumbprint $($storeFrontCertificate.Thumbprint) expired $($storeFrontCertificate.NotAfter.ToString('G'))" )
                }
                else
                {
                    [int]$daysToExpiry = ($storeFrontCertificate.NotAfter - [datetime]::Now).TotalDays
                    [string]$message = "Certificate with thumbprint $($storeFrontCertificate.Thumbprint) expires on $($storeFrontCertificate.NotAfter.ToString('G')) which is $daysToExpiry days from now"
                    if( $daysToExpiry -le $certificateWarningDays )
                    {
                        $warnings.Add( $message )
                    }
                    else
                    {
                        Write-Output -InputObject $message
                    }
                }
                $matchingSubjectAlternativeName = $storeFrontCertificate.DnsNameList | Where-Object Unicode -ieq $subject
                if( $null -eq $matchingSubjectAlternativeName )
                {
                    $warnings.Add(  "None of the $($storeFrontCertificate.DnsNameList.Count) Subject Alternative Name DNS names match the subject `"$subject`" (thumbprint $($storeFrontCertificate.Thumbprint))" )
                }
                else
                {
                    Write-Output -InputObject "Certificate Subject Alternative Name DNS names are ok ($($matchingSubjectAlternativeName.Unicode))"
                }
            }
            elseif( $binding.BaseUrl -match '^https:' )
            {
                $warnings.Add(  "Store $($store.Name) has a binding on port $($binding.port) for $($binding.BaseURL) with no certificate" )
            }
        }
        $bindingsForWebSites.Add( $webSiteAndId , $true ) ## mark as already done for this website
    }
}
## else already reported on these bindings

[string]$logsFolder = "$($env:ProgramFiles)\Citrix\Receiver StoreFront\Admin\trace" ## default but will set later if we have it in the SF config reg key
if( $storeFrontConfigurationKey -and $storeFrontConfigurationKey.PSObject.Properties[ 'InstallDir' ] )
{
    $logsFolder = Join-Path -Path $storeFrontConfigurationKey.InstallDir -ChildPath '\Admin\trace'
}

if( -Not ( Test-Path -Path $logsFolder -PathType Container ) )
{
    $warnings.Add( "StoreFront logs folder `"$logsFolder`" does not exist" )
}

[int]$filesCount = 0
[array]$relevantLogFiles =  @( Get-ChildItem -Path $logsFolder | Where-Object LastWriteTime -ge $startAnalysisDate )

Write-Verbose -Message "Got $($relevantLogFiles.Count) log files modified after $($startAnalysisDate.ToString('G'))"

[array]$logEntries = @( ForEach( $XMLlogFile in $relevantLogFiles )
{
    $filesCount++
    Write-Verbose -Message "$filesCount / $($relevantLogFiles.Count) : $($XMLlogFile.Name)"
    Process-LogFile -inputFile $XMLlogFile.FullName -start $startAnalysisDate -end ([datetime]::Now) | Where-Object Subtype -ine 'Information'
})

if( $null -ne $logEntries -and $logEntries.Count -gt 0 )
{
    ## remove GUIDs and time spans of form 00:00:00.12345. Replace CRLF separator with space
    $groupedLogEntries = @( $logentries|Select-Object -Property Date,SourceName,Subtype,@{n='Message';e={$_.ApplicationData -replace '\n' -replace  '[0-9a-z]{8}-[0-9a-z]{4}-[0-9a-z]{4}-[0-9a-z]{4}-[0-9a-z]{12}' -replace '[\s'']\d\d:\d\d:\d\d\.\d+' -replace "`r?`n" , ' ' }} | Group-Object -property message | ForEach-Object `
    {
        $group = $_
        $timerange = $group.group | Measure-Object -Minimum -Maximum -Property Date
        Add-Member -InputObject $group -PassThru -NotePropertyMembers @{
            'First Occurrence' = $timerange.Minimum.ToString('G')
            'Last Occurrence'  = $(if( $group.Count -gt 1 ) { $timerange.Maximum.ToString('G') } else { '-' })  ## only 1 occurence so don't clutter output with what will be the same date
            'Span (h)' = $(if( $group.Count -gt 1 ) { [math]::round( ($timerange.Maximum -$timerange.Minimum).TotalHours , 2 ) } else { '-' })
        }
    })
    $warnings.Add( "$($logEntries.Count) non-information level entries found in $($relevantLogFiles.Count) StoreFront log files modified since $($startAnalysisDate.ToString('G'))" )
    ## don't wrap as have multi line messages
    $warnings += @( $groupedLogEntries | Sort-Object -Property Count -Descending | Select-Object -Property Count, 'First Occurrence' , 'Last Occurrence' , 'Span (h)' ,
        @{name='Error' ; expression = { $_.Name -replace "`r?`n" , ' ' -replace '\s{2,}' ,' '}}  | Format-Table -AutoSize | Out-String )
}

## invoke FQDN, check result

## check config replication in HKEY_LOCAL_MACHINE\SOFTWARE\Citrix\DeliveryServices\ConfigurationReplication

## last modified time of C:\inetpub\wwwroot\Citrix\<store>\App_Data\default.ica

if( $warnings.Count -gt 0 )
{
    $warnings | Write-Warning
}
else
{
    Write-Output -InputObject "No warnings to report"
}


