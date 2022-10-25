#requires -version 3.0

<#
.SYNOPSIS
    Enable WMI logging for a given period of time, disable and parse events

.PARAMETER seconds
    How long to leave trace running for in seconds
    
.PARAMETER summary
    Produce a summary of activity to highlight highest consumers otherwise show individual WMI activities

.EXAMPLE
   & '.\Trace WMI activity.ps1' -seconds 120

   Trace WMI activity for 2 minutes and output a summary sorted by the processes consuming the most amount of time in WMI calls
   
.EXAMPLE
   & '.\Trace WMI activity.ps1' -seconds 60 -summary no

   Trace WMI activity for 2 minutes and output every operation

.NOTES
    https://docs.microsoft.com/en-us/windows/win32/wmisdk/tracing-wmi-activity
    
    THE SCRIPT IS PROVIDED IN AN 'AS IS' CONDITION, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE
    AND NONINFRINGEMENT. IN NO EVENT SHALL CONTROLUP, ANY AUTHORS OR ANY COPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, 
    ARISING FROM, OUT OF OR IN CONNECTION WITH THE SCRIPT OR THE USE OR OTHER DEALINGS IN THE SCRIPT

    Modification History:

    2022/06/02  @guyrleech  Initial release
    2022/06/02  @guyrleech  Added -summary parameter
    2022/06/08  @guyrleech  Added service information and exit 0 when no events so not an error
    2022/06/08  @guyrleech  Changed summary outputs
    2022/06/20  @guyrleech  Added summary outputs by WMI operation
    2022/06/23  @guyrleech  Changed summary by WMI operation
    2022/06/23  @guyrleech  Added number of different processes count to second summary
    2022/07/04  @guyrleech  Added functionality for detecting remote calls
    2022/07/05  @guyrleech  Made summaries cope with same pid on different machines
    2022/07/06  @guyrleech  Change display of local machine to "  local"
#>

[CmdletBinding()]

Param
(
    [int]$seconds = 30 ,
    [ValidateSet('yes','true','no','false','both')]
    [string]$summary = 'both'
)

[string]$timeFormat = 'HH:mm:ss.fff'

$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { 'Continue' } else { 'SilentlyContinue' })
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { 'Continue' } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'ErrorAction' ] ) { $ErrorActionPreference } else { 'Stop' })

[int]$outputWidth = 520 ## 4K monitor
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

if( $seconds -le 0 )
{
    Throw "Invalid survey period of $seconds seconds"
}

[string]$wmiTraceLog = 'Microsoft-Windows-WMI-Activity/Trace'

[xml]$wmilog = wevtutil.exe get-log $wmiTraceLog /f:xml

if( -Not $wmilog -or -not $wmilog.Channel )
{
    Throw "Failed to get state of event log $wmiTraceLog"
}

[bool]$weEnabledLog = $false
[bool]$warnedForAuditing = $false

try
{
    if( $wmilog.Channel.Enabled -ieq 'false' )
    {
        wevtutil.exe set-log $wmiTraceLog /enabled:true /quiet:true
        if( $? )
        {
            $weEnabledLog = $true
        }
    }
    else
    {
        Write-Warning -Message "$wmiTraceLog was already enabled"
    }

    [datetime]$startTime = [datetime]::Now

    [hashtable]$processesBefore = @{}
    [hashtable]$servicesBefore = @{}

    Get-Process | Where-Object Id -ne $pid | ForEach-Object -Process { $processesBefore.Add( $_.Id , $_ ) }
    Get-CimInstance -ClassName win32_service | Where ProcessId -gt 0 | ForEach-Object -Process `
    {
        ## as multiple services can share a process, we need to capture all services in that pid
        if( $existing = $servicesBefore[ $_.ProcessId ] )
        {
            $servicesBefore[ $_.ProcessId ] = "$existing,$($_.DisplayName)"
        }
        else
        {
            $servicesBefore.Add( $_.ProcessId , $_.DisplayName )
        }
    }

    Start-Sleep -Seconds $seconds

    [datetime]$endTime = [datetime]::Now

    [hashtable]$processesAfter = @{}
    [hashtable]$servicesAfter = @{}

    Get-Process | Where-Object Id -ne $pid | ForEach-Object -Process { $processesAfter.Add( $_.Id , $_ ) }    
    Get-CimInstance -ClassName win32_service | Where ProcessId -gt 0 | ForEach-Object -Process `
    {
        ## as multiple services can share a process, we need to capture all services in that pid
        if( $existing = $servicesAfter[ $_.ProcessId ] )
        {
            $servicesAfter[ $_.ProcessId ] = "$existing,$($_.DisplayName)"
        }
        else
        {
            $servicesAfter.Add( $_.ProcessId , $_.DisplayName )
        }
    }
    ## get events in our window
    [array]$events = @( Get-WinEvent -FilterHashTable @{logname = $wmiTraceLog ; id = 11,13 ; starttime = $startTime ; endtime = $endTime} -Oldest -ErrorAction SilentlyContinue )

    if( $null -eq $events -or $events.Count -eq 0 )
    {
        Write-Warning -Message "No events found"
        Exit 0
    }

    <#
    event 11 - Windows 10
       0 CorrelationId {D2CBD44A-6B43-0000-499F-89D3436BD801} 
       1 GroupOperationId 3631871 
       2 OperationId 3631872 
       3 Operation Start IWbemServices::CreateInstanceEnum - root\cimv2 : Win32_Thread 
       4 ClientMachine CONTOSO 
       5 ClientMachineFQDN CONTOSO 
       6 User CONTOSO\admingl 
       7 ClientProcessId 15244 
       8 ClientProcessCreationTime 132974156514040271 
       9 NamespaceName \\.\root\cimv2 
       10 IsLocal true 

    event 11 - pre-Windows 10
        0 CorrelationId {00000000-0000-0000-0000-000000000000} 
        1 GroupOperationId 640958 
        2 OperationId 640959 
        3 Operation Start IWbemServices::CreateInstanceEnum - root\cimv2 : Win32_SystemDriver 
        4 ClientMachine GRL-CTXCLDCON1 
        5 User GUYRLEECH\admingle 
        6 ClientProcessId 472 
        7 NamespaceName \\.\root\cimv2

    event 12
        GroupOperationId 3631871 
        Operation Provider::CreateInstanceEnum - CIMWin32 : Win32_Thread 
        HostId 6068 
        ProviderName CIMWin32 
        ProviderGuid {d63a5850-8f16-11cf-9f47-00aa00bf345c} 
        Path %systemroot%\system32\wbem\cimwin32.dll 

    event 13 (stop)
       OperationId 
       ResultCode 0x0 
 
    event 17
        CorrelationId {D2CBD44A-6B43-0000-499F-89D3436BD801} 
        ProcessId 15244 
        Protocol DCOM 
        Operation MI_Session::EnumerateInstance 
        User NULL 
        Namespace root\cimv2
    #>

    [array]$correlatedByOperationId = @( $events | Select-Object -Property *,@{n='OperationId';e={if( $_.Id -eq 11 ) { $_.Properties[2].Value}  else { $_.Properties[0].Value }}} | Group-Object -property OperationId )
    $processCreationEvents = New-Object -TypeName System.Collections.Generic.List[psobject]
    [bool]$gotProcessAuditingEvents = $false
    [hashtable]$securityEventFilter = @{ 'Logname' = 'Security' ;  Id = 4688 }

    ## different OSs have different number of properties in event id 11 so we must deal with this
    [int]$clientProcessIdIndex = -1
    [int]$processStartTimeIndex = -1
    [int]$userIndex = -1
    [int]$namespaceIndex = -1
    [int]$machineIndexId = -1

    $os = Get-CimInstance -ClassName win32_operatingsystem

    ##Write-Verbose -Message "Got $($correlatedByOperation.Count) different WMI operations"

    [array]$results = @(ForEach( $eventGroup in $correlatedByOperationId )
    {
        $firstEvent = $lastEvent = $null
        if( $eventGroup.Count -gt 1 -and `
            ( $firstEvent = ( $eventGroup.Group | Where-Object Id -eq 11 | Sort-Object -Property TimeCreated | Select-Object -First 1 ) ) -and ( $lastEvent = ( $eventGroup.Group | Where-Object recordId -ne $firstEvent.recordid | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1 ) ) )
        {
            if( $clientProcessIdIndex -lt 0 )
            {
                ## define the array indexes for the properties for event 11 as they differ depending on OS
                if( $firstEvent.Properties.Count -eq 11 ) # Win 10
                {
                    $userIndex = 6
                    $namespaceIndex = 9
                    $clientProcessIdIndex = 7
                    $processStartTimeIndex = 8
                    $machineIndexId = 4
                }
                elseif( $firstEvent.Properties.Count -eq 8 )
                {
                    $userIndex = 5
                    $namespaceIndex = 7
                    $clientProcessIdIndex = 6
                    $machineIndexId = 4
                    ## no process start time property
                }
                else
                {
                    Throw "Unexpected $($firstEvent.Properties.Count) properties for event id 11"
                }
            }
            if( ($Operation = $firstEvent.Properties[ 3 ].Value) -and $operation -ine 'IWbemServices::Connect' ) ## doesn't seem to tell us much
            {
                try
                {
                    [bool]$isLocal = $(if( $firstEvent.Properties.Count -ge 11) { $firstEvent.Properties[ 10 ].Value -ieq 'true'} else { $true } )
                    [int]$processPid = -1 
                    if( $firstEvent.Properties.Count -lt ($clientProcessIdIndex + 1) -or `
                        -Not ( ($processPid = [int32]$firstEvent.Properties[ $clientProcessIdIndex ].Value ) -ne $pid -and ( $process = $processesAfter[ $processPid ] ) ) -and $lastEvent.Properties.Count -ge ($clientProcessIdIndex + 1)  )
                    {
                        $process = $processesBefore[ [int32]$lastEvent.Properties[ $clientProcessIdIndex ].Value ]
                    }
                    if( $processPid -eq $pid )
                    {
                        ## skip self
                        continue
                    }
                    if( $isLocal ) ## chances of us being able to remote to remote system to get process details is small plus we can't get before snapshots
                    {
                        ## if process has exited we can't get start time but it is stored in event properties so get it from there
                        $processStartTime = $process | Select-Object -ExpandProperty StartTime -ErrorAction SilentlyContinue
                        if( -Not $processStartTime -and $processStartTimeIndex -ge 0 -and $firstEvent.Properties.Count -ge ($processStartTimeIndex + 1) )
                        {
                            $processStartTime = [datetime]::FromFileTime( $firstEvent.Properties[ $processStartTimeIndex ].Value )
                        }
                        if( -Not $process )
                        {
                            ## see if we an find a process started auditing event (4688) to get the information (e.g. process started and died during window or started before and exited during window
                            ## if process start time outside of our run window then look for that event specifically otherwise cache everything in that window
                            ## will have to work backwards through events to cater for pid re-use unless we have an exact start time
                            ## add events outside of the window to our list so that if there is a subsequent instance of this pid, it is already cached
                            ## 
                            if( -Not $gotProcessAuditingEvents )
                            {
                                ## if we have a process start time then search from this and next time we come in here see if we have already searched and if not change search to be from the new process start time until the previous start search time
                                if( $processStartTime )
                                {                        
                                    if( $securityEventFilter[ 'StartTime' ] )
                                    {
                                        ## if $processStartTime after current start time then no point getting any more events
                                        if( $processStartTime -lt $securityEventFilter[ 'StartTime' ] )
                                        {
                                            $securityEventFilter[ 'EndTime' ] = $securityEventFilter[ 'StartTime' ]
                                            $securityEventFilter[ 'StartTime' ] = $processStartTime.AddSeconds( -1 )
                                        }
                                    }
                                    else ## no existing starttime
                                    {
                                        $securityEventFilter[ 'StartTime' ] = $processStartTime.AddSeconds( -1 )
                                    }
                                }
                                else ## no process start time so do we get all 4688 events which could be a lot ?
                                {
                                    $securityEventFilter[ 'StartTime' ] = $startTime.AddSeconds( -60 )
                                }

                                if( $os -and $securityEventFilter[ 'StartTime' ] -lt $os.LastBootUpTime )
                                {
                                   ## have seen instances where the process start time property converts to a date in 1601 which means we would search entire event log which could be large & thus slow if persistent
                                   $securityEventFilter[ 'StartTime' ] = $os.LastBootUpTime
                                }
                                $processCreationEvents = @( Get-WinEvent -FilterHashtable $securityEventFilter -ErrorAction SilentlyContinue )
                                if( $processCreationEvents.Count -gt 0 )
                                {
                                    Write-Verbose -Message "Got $($processCreationEvents.Count) process creation events (earliest $(Get-Date -Date $processCreationEvents[-1].TimeCreated -Format G))" 
                                }
                                else
                                {
                                    if( -Not $warnedForAuditing )
                                    {
                                        Write-Warning -Message "Could not find any process auditing events"
                                        $warnedForAuditing
                                    }
                                    Write-Verbose -Message "Got no process creation events"
                                }
                                $gotProcessAuditingEvents = $true
                            }
                    
                            if( $gotProcessAuditingEvents )
                            {
                                [Int64]$processStartTimeFileTime = 0
                                if( $processStartTime )
                                {
                                    ## process start times will not match exactly so see if "close" (filetime is 100ns granularity)
                                    $processStartTimeFileTime = $processStartTime.ToFileTime()
                                }
                                if( $cachedProcess = $processCreationEvents.Where( { $_.Properties[4].value -eq $processPid -and ( $processStartTimeFileTime -eq 0 -or [math]::Abs( $_.TimeCreated.ToFileTime() - $processStartTimeFileTime ) -lt 10000 ) } , 1))
                                {
                                    $process = [pscustomobject]@{
                                        'Name' = Split-Path -Path $cachedProcess.Properties[5].value -Leaf
                                        'Path' = $cachedProcess.Properties[5].value
                                        'SessionId' = $null
                                        'Company' = Get-ItemProperty -Path $cachedProcess.Properties[5].value | Select-Object -ExpandProperty VersionInfo -ErrorAction SilentlyContinue | Select-Object -ExpandProperty CompanyName
                                        }
                                    if( -not $processStartTime )
                                    {
                                        $processStartTime = $cachedProcess.TimeCreated
                                    }
                                }
                            }
                        }
                    }
                    ## else process is remote so we can't get start time as start time is not something I know how to decode when remote as not a DMTF datetime or seconds since 1/1/1970
                }
                catch
                {
                    Write-Warning -Message "Exception processing $firstevent : $_"
                }
                try
                {
                    [pscustomobject]@{
                        Start = Get-Date -Date $firstEvent.TimeCreated -Format $timeFormat
                        End   = Get-Date -Date $lastEvent.TimeCreated  -Format $timeFormat
                        'Duration (ms)' = [int]($lastEvent.TimeCreated - $firstEvent.TimeCreated).TotalMilliseconds
                        ## strip off Start IWbemServices::CreateInstanceEnum - root\cimv2 : 
                        Operation = $operation -replace '^Start IWbemServices::ExecQuery [^:]+:\s*'
                        User = $(if( $firstEvent.Properties.Count -ge ($userIndex + 1)) { $firstEvent.Properties[ $userIndex ].Value } )
                        Machine = $(if( $firstEvent.Properties.Count -ge ($machineIndexId + 1 )) { if( $machine = $firstEvent.Properties[ $machineIndexId ].Value ) { if( $machine -ieq $env:COMPUTERNAME ) { '  local' } else { $machine } } } )
                        Process = $process | Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue
                        ProcessPath = $process | Select-Object -ExpandProperty Path -ErrorAction SilentlyContinue
                        Company = $process | Select-Object -ExpandProperty Company -ErrorAction SilentlyContinue
                        Pid = $processPid
                        SessionId = $process | Select-Object -ExpandProperty SessionId -ErrorAction SilentlyContinue
                        Service = $(if( $isLocal -and ( $service = $servicesAfter[ [uint32]$processPid ] ) -or ( $service = $servicesBefore[ [uint32]$processPid ] ) ) { $service })
                        ## ProcessStartTime = $processStartTime
                        ResultCode = $(if( $lastEvent.Id = 13 ) { $lastEvent.Properties[ 1 ].Value })
                        NameSpace = $(if( $firstEvent.Properties.Count -ge ($namespaceIndex + 1)) { $firstEvent.Properties[ $namespaceIndex ].Value } )
                        Local = $isLocal
                    }
                }
                catch
                {
                    Write-Warning -Message "Exception creating object from $firstevent : $_"
                }
            }
        }
        elseif( $eventGroup.Group.Count -eq 1 -and $eventGroup.Group[0].Id -eq 13 )
        {
            Write-Verbose -Message "Ignoring single event at $(Get-Date -Date $eventGroup.Group[0].TimeCreated -Format G) as stop where we don't have a start - operation id $($eventGroup.Name)"
        }
        elseif( $firstEvent -or $lastEvent )
        {  
            try
            {
                [bool]$isLocal = $(if( $firstEvent -and $firstEvent.Properties.Count -ge 11) { $firstEvent.Properties[ 10 ].Value -ieq 'true'} else { $true } )
                [pscustomobject]@{
                        Start = $(if( $starttime = $firstEvent | Select-Object -ExpandProperty TimeCreated ) { Get-Date -Format $timeFormat -Date $endTime } )
                        End = $(if( $endtime = $lastEvent | Select-Object -ExpandProperty TimeCreated ) { Get-Date -Format $timeFormat -Date $endTime } )
                        'Duration (ms)' = $(if( $firstEvent -and $lastEvent) { [int]($lastEvent.TimeCreated - $firstEvent.TimeCreated).TotalMilliseconds })
                        ## strip off Start IWbemServices::CreateInstanceEnum - root\cimv2 : 
                        Operation = $operation -replace '^Start IWbemServices::ExecQuery [^:]+:\s*'
                        User = $(if( $firstEvent -and $firstEvent.Properties.Count -ge ($userIndex + 1)) { $firstEvent.Properties[ $userIndex ].Value } )
                        Machine = $(if( $firstEvent.Properties.Count -ge ($machineIndexId + 1 )) { $firstEvent.Properties[ $machineIndexId ].Value } )
                        Process = $process | Select-Object -ExpandProperty Name
                        ProcessPath = $process | Select-Object -ExpandProperty Path
                        Pid = $processPid
                        SessionId = $process | Select-Object -ExpandProperty SessionId
                        Service = $(if( ( $service = $servicesAfter[ [uint32]$processPid ] ) -or ( $service = $servicesBefore[ [uint32]$processPid ] ) ) { $service })
                        ## ProcessStartTime = $processStartTime
                        ResultCode = $null
                        NameSpace = $(if( $firstEvent -and $firstEvent.Properties.Count -ge ($namespaceIndex + 1)) { $firstEvent.Properties[ $namespaceIndex ].Value } )
                        Local = $isLocal
                }
            }
            catch
            {
                Write-Warning -Message "Exception creating object for process $process : $_"
            }
        }
    })
    if( $results -and $results.Count -gt 0 )
    {
        if( $summary -ieq 'yes' -or $summary -imatch 'true$' -or $summary -ieq 'both' )
        {
            [array]$groupedByProcess = @( $results | Group-Object -Property Pid,Machine )
            [array]$groupedByOperation = @( $results.Where( { $null -ne $_.'Duration (ms)' -and $_.'Duration (ms)' -gt 0 } ) | Group-Object -Property Operation,ResultCode )
            [array]$summaryData = @( ForEach( $item in $groupedByProcess )
            {
                [int]$longestDuration = -1
                [long]$totalDuration = 0
                [string]$company = $null
                [int]$sessionId = -1
                [string]$processName = $null
                [string]$processPath = $null
                [string]$longestOperation = $null
                [hashtable]$differentOperations = @{}

                ForEach( $operation in $item.Group )
                {
                    $totalDuration += $Operation.'Duration (ms)'
                    if( $operation.'Duration (ms)' -gt $longestDuration )
                    {
                        $longestDuration = $operation.'Duration (ms)'
                        $longestOperation = $operation.Operation
                    }
                    if( -Not $company -and $operation.PSObject.properties[ 'company' ])
                    {
                        $company = $operation.Company
                    }
                    if( -Not $processName )
                    {
                        $processName = $operation.process
                    }
                    if( $sessionId -lt 0 -and $operation.PSObject.properties[ 'sessionId' ] -and $null -ne $operation.sessionId )
                    {
                        $sessionId = $operation.sessionId
                    }
                    if( -Not $processPath -and $operation.PSObject.properties[ 'processPath' ] )
                    {
                        $processPath = $operation.processPath
                    }
                    try
                    {
                        $differentOperations.Add( $Operation.operation , $true )
                    }
                    catch
                    {
                        ## already have it 
                    }
                }
                Add-Member -InputObject $item -PassThru -NotePropertyMembers @{
                    'Total Duration (s)' = $totalDuration / 1000
                    ##'Longest Duration (ms)' = $(if( $longestDuration -ge 0 ) { $longestDuration } else { $null } )
                    ##'Longest Operation' = $longestOperation
                    'Company' = $company
                    'Process' = $processName
                    'ProcessPath' = $processPath
                    'Service' = $operation.Service
                    'Different Operations' = $differentOperations.Count
                }
            } )

            "Summary for each process, sorted by the highest total time spent in WMI calls"
            "-----------------------------------------------------------------------------"

            $summaryData | Sort-Object -Property 'Total Duration (s)' -Descending | Select-Object Process,@{n='Pid';e={($_.Name -split ',')[0]}},@{n='Total Operations';e={$_.Count}},'Different Operations',*Duration*,ProcessPath,@{n='Machine';e={($_.Name -split ',')[-1].Trim()}},Service,Company <#-ExcludeProperty Name,Group,Values,Count#> | Format-Table -AutoSize -Property *

            ## Find the longest duration for each unique WMI call, split by result code so we can tell if bad result codes cause longer queries
          
            ForEach( $operation in $groupedByOperation )
            {  
                [int]$totalDuration = 0
                [int]$fastest = [int]::MaxValue
                [int]$slowest = 0
                [hashtable]$differentProcesses = @{}

                ForEach( $instance in $operation.Group )
                {
                    if( $instance.'Duration (ms)' -lt $fastest )
                    {
                        $fastest = $instance.'Duration (ms)'
                    }
                    if( $instance.'Duration (ms)' -gt $slowest )
                    {
                        $slowest = $instance.'Duration (ms)'
                    }
                    $totalDuration += $instance.'Duration (ms)'
                    try
                    {
                        ## need to cater for same process id on different machines
                        $differentProcesses.Add( "$($instance.pid)-$($instance.Machine)" , $true )
                    }
                    catch
                    {
                        ## already got it
                        $null
                    }
                }
                Add-Member -InputObject $Operation -NotePropertyMembers @{
                    'Processes'    = $differentProcesses.Count
                    'Fastest (ms)' = $fastest
                    'Slowest (ms)' = $slowest
                    'Average (ms)' = [int]( $totalDuration / $Operation.group.Count )
                    'Result Code'  = $instance.ResultCode ## grouped by this (secondary) so should only be 1 per grouped item set of values
                    'Operation'    = $instance.Operation ## as grouped by result code too, name property will have ", resultcode" on end
                }
            }
            
            "Summary by unique WMI calls across all processes, sorted by highest WMI operation count"
            "---------------------------------------------------------------------------------------"

            $groupedByOperation | Sort-Object -Property 'Count' -Descending | Format-Table -AutoSize -Property Operation,Count,'* (ms)','Processes','Result Code'
        }
        
        if( $summary -ieq 'no' -or $summary -imatch 'false$' -or $summary -ieq 'both' )
        {
            $results | Format-Table -AutoSize -Property *
        }
    }
    else
    {
        Write-Warning -Message "No usable event data found in $($events.Count) events"
    }
}
catch
{
    Throw $_
}
finally
{
    if( $weEnabledLog )
    {
        wevtutil.exe set-log $wmiTraceLog /enabled:false /quiet:true
    }
}
