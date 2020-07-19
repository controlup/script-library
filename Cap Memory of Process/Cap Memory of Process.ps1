<#
    Set a hard working set limit for the process

    @guyrleech 2019

    Based on code from https://github.com/guyrleech/Microsoft/blob/master/Trimmer.ps1
#>

if( $args.Count -ne 2 -or ! $args[0] -or ! $args[1] )
{
    Throw 'Must pass the pid of the process to act on and working set size in MB as the only arguments'
}

[int]$thePid = $args[0] -as [int]
[int]$maxWorkingSet = ($args[1] -as [int]) * 1MB

$process = Get-Process -Id $thePid -ErrorAction SilentlyContinue

if( ! $process )
{
    Throw "Failed to get a process for pid $thePid"
}

if( ! $process.Handle )
{
    Throw "No process handle for pid $thePid"
}

Add-Type -Debug:$false @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PInvoke.Win32
{
  
    public static class Memory
    {
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern bool SetProcessWorkingSetSizeEx( IntPtr hProcess, int min, int max , int flags );
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern bool GetProcessWorkingSetSizeEx( IntPtr hProcess, ref int min, ref int max , ref int flags );
    }
}
'@

$statsBefore = $null
$statsAfter  = $null

[int]$flags = 4 ## QUOTA_LIMITS_HARDWS_MAX_ENABLE
[int]$thisMinimumWorkingSet = -1 
[int]$thisMaximumWorkingSet = -1 
[int]$thisFlags = -1 ## Grammar alert! :-)

## https://msdn.microsoft.com/en-us/library/windows/desktop/ms683227(v=vs.85).aspx
[bool]$result = [PInvoke.Win32.Memory]::GetProcessWorkingSetSizeEx( $process.Handle, [ref]$thisMinimumWorkingSet , [ref]$thisMaximumWorkingSet , [ref]$thisFlags );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

if( $result )
{
    ## convert flags value - if not hard then will be soft so no point reporting that separately IMHO
    [bool]$hardMinimumWorkingSet = $thisFlags -band 1 ## QUOTA_LIMITS_HARDWS_MIN_ENABLE
    [bool]$hardMaximumWorkingSet = $thisFlags -band 4 ## QUOTA_LIMITS_HARDWS_MAX_ENABLE
    $statsBefore = New-Object -TypeName pscustomobject -ArgumentList @{ 
            'Hard Minimum Working Set Limit' = $hardMinimumWorkingSet ; 'Hard Maximum Working Set Limit' = $hardMaximumWorkingSet ;
            'Working Set (MB)' = [math]::Round( $process.WorkingSet64 / 1MB , 1 ) ; 'Peak Working Set (MB)' = [math]::Round( $process.PeakWorkingSet64 / 1MB , 1 );
            'Commit Size (MB)' = [math]::Round( $process.PagedMemorySize / 1MB , 1 ); 
            'Paged Pool Memory Size (KB)' = [math]::Round( $process.PagedSystemMemorySize64 / 1KB , 1 ); 'Non-paged Pool Memory Size (KB)' = [math]::Round( $process.NonpagedSystemMemorySize64 / 1KB , 1 );
            'Minimum Working Set (KB)' = $thisMinimumWorkingSet / 1KB ; 'Maximum Working Set (KB)' = $thisMaximumWorkingSet / 1KB }
}
else
{                   
    Write-Warning ( "Failed to get working set info for {0} pid {1} - {2}" -f $process.Name , $process.Id , $LastError)
    $thisMinimumWorkingSet = 1KB ## will be set to the minimum by the call
}
                                     
## see https://msdn.microsoft.com/en-us/library/windows/desktop/ms686237(v=vs.85).aspx
$result = [PInvoke.Win32.Memory]::SetProcessWorkingSetSizeEx( $process.Handle , $thisMinimumWorkingSet , $maxWorkingSet , $flags );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

if( ! $result )
{                   
    Write-Error ( "Failed to set working set info for {0} pid {1} to {2}MB - {3}" -f $process.Name , $process.Id , $maxWorkingSet / 1MB , $LastError)
}

$result = [PInvoke.Win32.Memory]::GetProcessWorkingSetSizeEx( $process.Handle, [ref]$thisMinimumWorkingSet , [ref]$thisMaximumWorkingSet , [ref]$thisFlags );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

if( $result )
{
    $process = Get-Process -Id $process.Id -ErrorAction Continue
    ## convert flags value - if not hard then will be soft so no point reporting that separately IMHO
    [bool]$hardMinimumWorkingSet = $thisFlags -band 1 ## QUOTA_LIMITS_HARDWS_MIN_ENABLE
    [bool]$hardMaximumWorkingSet = $thisFlags -band 4 ## QUOTA_LIMITS_HARDWS_MAX_ENABLE
    $statsAfter = New-Object -TypeName pscustomobject -ArgumentList @{ 
            'Hard Minimum Working Set Limit' = $hardMinimumWorkingSet ; 'Hard Maximum Working Set Limit' = $hardMaximumWorkingSet ;
            'Working Set (MB)' = [math]::Round( $process.WorkingSet64 / 1MB , 1 );'Peak Working Set (MB)' = [math]::Round( $process.PeakWorkingSet64 / 1MB , 1 ) ;
            'Commit Size (MB)' = [math]::Round( $process.PagedMemorySize / 1MB , 1 ); 
            'Paged Pool Memory Size (KB)' = [math]::Round( $process.PagedSystemMemorySize64 / 1KB , 1 ); 'Non-paged Pool Memory Size (KB)' = [math]::Round( $process.NonpagedSystemMemorySize64 / 1KB , 1 );
            'Minimum Working Set (KB)' = $thisMinimumWorkingSet / 1KB ; 'Maximum Working Set (KB)' = $thisMaximumWorkingSet / 1KB }
}

"Memory statistics for process '$($process.Name)' ($thePid) started at $(Get-Date -Date $process.StartTime -Format G):"

@( ForEach( $statBefore in $statsBefore.GetEnumerator() )
{
    New-Object -TypeName psobject -Property @{ 'Value' = $statBefore.Name ; 'Before' = $statBefore.Value ; 'After' = $statsAfter[ $statBefore.Name ] }
} ) | Sort-Object -Property Value | Format-Table -AutoSize -Property 'Value','Before','After'

