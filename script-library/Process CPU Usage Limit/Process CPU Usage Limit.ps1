#requires -version 3
<#
.SYNOPSIS

Find over consuming threads of the specified process and pause them for a while to reduce their CPU usage

.DETAILS

Pauses threads within the monitored process when that thread's CPU usage ius excessive based on the aggressiveness and resumes them when the thread's average CPU usage drops below the threshold

.PARAMETER badPid

The process id of the process which is to be monitored/clamped

.PARAMETER aggressiveness

The agressiveness of the CPU throttling where 1 is low and 10 is high

.PARAMETER durationMinutes

The duration in minutes of how long the process will monitor/control CPU for. If set to 0 then will monitor for the life of the monitored process.

.PARAMETER relaunched

Not exposed to the console. This is so the script can tell if it neds to copy itself and respawn with the same parameters and adding -relaunched

.CONTEXT

Process 

.MODIFICATION_HISTORY:

@guyrleech 29/10/19

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='Process Id to monitor')]
    [int]$badPid ,
    [Parameter(Mandatory=$true,HelpMessage='Aggressiveness of CPU clamping, 1 to 10 where 1 is low')]
    [double]$aggressiveness ,
    [double]$durationMinutes = 0 ,
    [switch]$relaunched ## not specified from the console, used to control respawning of script
)

$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { 'Continue' } else { 'SilentlyContinue' })
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { 'Continue' } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'ErrorAction' ] ) { $ErrorActionPreference } else { 'Stop' })

Add-Type @'
using System;
using System.Runtime.InteropServices;

public static class Kernel32
{
    [DllImport( "kernel32.dll",SetLastError = true )]
    public static extern IntPtr OpenThread( 
        UInt32 dwDesiredAccess, 
        bool bInheritHandle, 
        UInt32 dwThreadId );

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool CloseHandle(
        [In] IntPtr hHandle );
            
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern int SuspendThread(
        [In] IntPtr hThread );

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern int Wow64SuspendThread(
        [In] IntPtr hThread );

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern int ResumeThread(
        [In] IntPtr hThread );

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool IsWow64Process(
        [In] IntPtr hProcess ,
        [Out,MarshalAs(UnmanagedType.Bool)] out bool wow64Process );

    public enum ThreadAccess
    {
        THREAD_SUSPEND_RESUME = 0x2 ,
        THREAD_QUERY_INFORMATION = 0x40 ,
        THREAD_QUERY_LIMITED_INFORMATION = 0x800,
    };
}
'@ -ErrorAction Stop

[int]$targetCPU = 100 - ($aggressiveness * 10) ## percent

if( $targetCPU -lt 0 -or $targetCPU -gt 100 )
{
    Throw "Illegal value for aggressiveness specified - must be between 1 and 10"
}

$PSBoundParameters.GetEnumerator() | Write-Verbose

if( ! $relaunched ) ## need to relaunch this script so the ControlUp launched instance can return and not timeout
{
    ## Copy this script since ControlUp will delete the script it launches
    ## Get folder for script and we'll use it too
    [string]$originalScript =  ( & { $myInvocation.ScriptName } )
    $process = Get-Process -Id $badPid| Select -First 1
    if( $process )
    {
        $copiedScript = Join-Path (Split-Path $originalScript -Parent) ('Async_' + $(Split-Path $originalScript -Leaf))
        Copy-Item -Path $originalScript -Destination $copiedScript -Force
        if( Test-Path $copiedScript -PathType Leaf -ErrorAction SilentlyContinue )
        {
            [string]$arguments = "-badPid $badPid -aggressiveness $aggressiveness -durationMinutes $durationMinutes -relaunched"
            Write-Verbose "Arguments are $arguments"
            $invocation = Invoke-CimMethod -ClassName Win32_Process -MethodName create -Arguments @{ 
                CommandLine = ( "Powershell.exe -NonInteractive -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File `"$copiedScript`" $arguments" ) ;
                ProcessStartupInformation = New-CimInstance -CimClass ( Get-CimClass Win32_ProcessStartup ) -Property @{ ShowWindow = 1 } -Local ;  
                CurrentDirectory = $null }
            if( ! $invocation -or ! $invocation.ProcessId )
            {
                Write-Error "Failed to launch `"$copiedScript`", return value was $($invocation.ReturnValue)"
                exit $invocation.ReturnValue
            }
            else
            {
                Write-Output "Launched `"$copiedScript`" asynchronously as process id $($invocation.ProcessId)"
            }
        }
        else
        {
            Throw "Failed to copy `"$originalScript`""
        }
    }
    else
    {
        Write-Error "Unable to get process for pid $badPid"
    }
    ## we have launched our child process so we exit
    exit 0
}

[int]$is32bitProcess = 0
## Get session for thread as we will not operate on session 0 due to deadlock potential
$theProcess = Get-Process -Id $badPid -ErrorAction Stop

if( $theProcess.SessionId -eq 0 )
{
    Throw "Process $badPid is in session 0 which could cause deadlocks if changed"
}

if( ! [kernel32]::IsWow64Process( $theProcess.Handle , [ref]$is32bitProcess ) )
{
    Throw "Failed to determine if process $badPid is 32 or 64 bit"
}

[hashtable]$adjustedThreads = @{}
[int]$adjustmentsMade = 0
[int]$samplePeriod = 100 ## milliseconds
[hashtable]$threadInfo = @{}
[int]$indexSize = 30
$thisProcess = $null
[bool]$pulse = $false

$timer = [Diagnostics.Stopwatch]::StartNew()

$theProcess | select -ExpandProperty Threads | ForEach-Object `
{
    ## Add our own CPU consumption object so that WMI won't update it
    $threadInfo.Add( $_.Id , [pscustomobject]@{ 
        'TotalCPUConsumed' = $(if( $_.TotalProcessorTime.PSObject.Properties[ 'TotalMilliseconds' ] ) { $_.TotalProcessorTime.TotalMilliseconds } else { 0 })
        'Handle' = $null
        'Consumptions' = @( 1..$indexSize | ForEach-Object { 0 } )
        'Index' = [long]0
        'Timer' = $timer.Elapsed.TotalMilliseconds
        'Pulse' = $pulse
        'Paused' = $false } )
}

[datetime]$startTime = [Datetime]::Now

try
{
    do
    {
        Start-Sleep -Milliseconds $samplePeriod 
        $pulse = ! $pulse
        [array]$threadsAfter = @( )
        ## Now see which threads have consumed the most CPU
        $thisProcess = Get-Process -Id $badPid -ErrorAction SilentlyContinue
        if( ! $thisProcess -or $thisProcess.HasExited )
        {
            Write-Warning "Process $badPid has exited"
            break
        }

        $thisProcess | select -ExpandProperty Threads | ForEach-Object `
        {
            $threadAfter = $PSItem
            $existingThread = $threadInfo[ $threadAfter.Id ]
            if( $existingThread )
            {
                [double]$cpuConsumptionNow = $(if(  $threadAfter.TotalProcessorTime.PSObject.Properties -and $threadAfter.TotalProcessorTime.PSObject.Properties[ 'TotalMilliseconds' ] ) { $threadAfter.TotalProcessorTime.TotalMilliseconds } else { 0 })
               
                [double]$threadCPUms = ( $cpuConsumptionNow - $existingThread.TotalCPUConsumed )
                [double]$timeMs = $timer.Elapsed.TotalMilliseconds - $existingThread.Timer
                [double]$totalThreadCPUPercent = $threadCPUms / $timeMs * 100
                ## If the array has not yet filled up then we take the average of the elements that have been added, not the whole array as the zeroes will bring the average down
                [int]$averageConsumption = ( $existingThread.Consumptions | Measure-Object -Sum |Select-Object -ExpandProperty Sum ) / $( if(  $existingThread.Index -and $existingThread.Index -lt $existingThread.Consumptions.Count) { $existingThread.Index } else { $existingThread.Consumptions.Count })
                if( $averageConsumption -gt $targetCPU )
                {
                    if( $existingThread.Paused )
                    {
                        Write-Verbose -Message "Thread id $($threadAfter.Id) average CPU is $($averageConsumption)% but already paused so leaving paused"
                    }
                    else ## not paused yet
                    {
                        Write-Verbose -Message "Thread id $($threadAfter.Id) is consuming $($totalThreadCPUPercent)% now, average $($averageConsumption)% exceeding $($targetCPU)% so reducing it"

                        Write-Verbose -Message "$([int]$timer.Elapsed.TotalMilliseconds)ms : Pausing thread $($threadAfter.Id) @ $($totalThreadCPUPercent)%"

                        if( ! $existingThread.Handle )
                        {
                            $existingThread.Handle = [kernel32]::OpenThread(  [Kernel32+ThreadAccess]::THREAD_SUSPEND_RESUME , $false , $threadAfter.Id ); $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                        }
                
                        if( $existingThread.Handle )
                        {
                            [int]$result = -1
                            if( $is32bitProcess )
                            {
                                $result = [kernel32]::Wow64SuspendThread( $existingThread.Handle ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                            }
                            else
                            {
                                $result = [kernel32]::SuspendThread( $existingThread.Handle ) ;  $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                            }
                            if( $result -ge 0 )
                            {
                                $existingThread.Paused = $true
                                $adjustmentsMade++
                                try
                                {
                                    $adjustedThreads.Add( $threadAfter.Id , $badPid )
                                }
                                catch {}
                            }
                            else
                            {
                                Write-Warning "Failed to suspend thread id $($threadAfter.Id) - $lastError"
                            }
                        }
                    }
                }
                elseif( $existingThread.Paused )
                {
                    Write-Verbose "$([int]$timer.Elapsed.TotalMilliseconds)ms : Resuming thread $($threadAfter.Id)"
                    [int]$result = [kernel32]::ResumeThread( $existingThread.Handle ) ;  $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                    if( $result -ge 0 )
                    {
                        $existingThread.Paused = $false
                    }
                    else
                    {
                        Write-Warning "Failed to resume thread id $($threadAfter.Id) - $lastError"
                    }
                }

                ## update our copy so that CPU counters are from last sample
                $existingThread.TotalCPUConsumed = $cpuConsumptionNow
                $existingThread.Pulse = $pulse
                $existingThread.Timer = $timer.Elapsed.TotalMilliseconds
            }
            else
            {
                $existingThread = [pscustomobject]@{ 
                    'TotalCPUConsumed' = $cpuConsumptionNow 
                    'Handle' = $null
                    'Consumptions' = @( 1..$indexSize | ForEach-Object { 0 } )
                    'Index' = [long]0
                    'Pulse' = $pulse
                    'Timer' = $timer.Elapsed.TotalMilliseconds
                    'Paused' = $false }
                $threadInfo.Add( $threadAfter.Id , $existingThread )
            }

            $existingThread.Consumptions[ $existingThread.Index % $indexSize ] = $totalThreadCPUPercent
            $existingThread.Index++
        }
        ## now remove any threads that haven't been checked as that means they are dead
        [array]$toDelete = @( $threadInfo.GetEnumerator() | ForEach-Object `
        {
            if( $_.Value.Pulse -ne $pulse )
            {
                $_.Key
            }
        })
        $toDelete | ForEach-Object `
        {
            Write-Verbose "`tRemoving thread $_"
            $threadInfo.Remove( $_ )
        }
    } while( ! $durationMinutes -or ([datetime]::Now - $startTime).TotalMinutes -le $durationMinutes ) ## loop until process exits or duration is reached
}
catch
{
    Throw $_
}
finally
{
    $timer.Stop()
    if( $thisProcess )
    {
        ## Check if any threads are still paused by us and resume them
        $threadInfo.GetEnumerator() | ForEach-Object `
        {   
            $existingThread = $_.Value
            if( $existingThread.Handle )
            {
                if( $_.value.Paused )
                {
                    [int]$result = [kernel32]::ResumeThread( $existingThread.Handle )
                    if( $result -ge 0 )
                    {
                        $existingThread.Paused = $false
                    }
                    else
                    {
                        Write-Warning "Failed to final resume of thread id $($threadAfter.Id) - $lastError"
                    }
                }
                [void][kernel32]::CloseHandle( $existingThread.Handle )
                 $existingThread.Handle = $null
            }
        }
    }
}

