<#
.SYNOPSIS

Empty working sets or trim them down to a specific value for a given process or all processes in a specified session

.PARAMETER id

Process id or session id to act on

.PARAMETER trimToMB

The size of the working set to trim down to. If 0 or negative then the entire working set will be emptied. The WS can grow above this as long as hard limit is not in place.

.CONTEXT

Computer, Session

.LINK

Based on code from https://github.com/guyrleech/Microsoft/blob/master/Trimmer.ps1

.NOTES

If too many processes are trimmed too frequently, performance can suffer due to hard page faults so use with caution

.MODIFICATION_HISTORY:

@guyrleech 25/07/19

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='Process or session id to act on')]
    [int]$id ,
    [Parameter(Mandatory=$true,HelpMessage='Working set size to trim down to')]
    [int]$trimToMB,
    [Parameter(Mandatory=$false,HelpMessage='Processes to exclude from memory trim (eg, cmd/powershell/winlogon)')]
    [string]$excludedProcesses
)

[bool]$sessionContext = $true ## works on either a single process all all processes in a session

$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

[int]$thePid = -1
[int]$theSessionId = -1
[array]$processes = @()

if( $sessionContext )
{
    $theSessionId = $id
    $processes = @( Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $theSessionId } )

    if( ! $processes -or ! $processes.Count )
    {
        Throw "Failed to get any processes for session $theSessionId"
    }
}
else
{
    $thePid = $id
    $processes = @( Get-Process -Id $thePid -ErrorAction SilentlyContinue )

    if( ! $processes -or ! $processes.Count )
    {
        Throw "Failed to get a process for pid $thePid"
    }
}

[int]$maxWorkingSet = $trimToMB * 1MB

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

[int]$thisMinimumWorkingSet = -1 
[int]$thisMaximumWorkingSet = -1 
[int]$thisFlags = 0
[int]$counter = 0

$listOfExcludedProcesses = $excludedProcesses -split "/"
ForEach( $process in $processes )
{
    if ($listOfExcludedProcesses -like "$($process.ProcessName)") {
        Write-Output "Excluded process from memory trim: $($Process.ProcessName) pid $($Process.Id)"
        continue
    }
    else
    {
        Write-Verbose "Trimming memory for $($process.ProcessName)"
        if( ! $process.Handle )
        {
            Write-Warning -Message "No process handle for process $($process.ProcessName) pid $($process.Id)"
        }
        else
        {
            ## https://msdn.microsoft.com/en-us/library/windows/desktop/ms683227(v=vs.85).aspx 
            [bool]$result = [PInvoke.Win32.Memory]::GetProcessWorkingSetSizeEx( $process.Handle, [ref]$thisMinimumWorkingSet , [ref]$thisMaximumWorkingSet , [ref]$thisFlags );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

            if( ! $result )
            {                
                Write-Warning ( "Failed to get working set info for {0} pid {1} - {2}" -f $process.ProcessName , $process.Id , $LastError)
                $thisMinimumWorkingSet = 1KB ## will be set to the minimum by the call
            }
            else
            {
                [int]$originalMaxWorkingSet = 0
                if( $maxWorkingSet -le 0 )
                {
                    $thisMaximumWorkingSet = $thisMinimumWorkingSet = -1 ## emptying the working set
                }
                else
                {
                    $originalMaxWorkingSet = $thisMaximumWorkingSet
                    $thisMaximumWorkingSet = $maxWorkingSet ## not completely emptying & can grow above this , assuming it doesn't have a hard WS limit flag set
                }
   
                ## see https://msdn.microsoft.com/en-us/library/windows/desktop/ms686237(v=vs.85).aspx
                $result = [PInvoke.Win32.Memory]::SetProcessWorkingSetSizeEx( $process.Handle , $thisMinimumWorkingSet , $thisMaximumWorkingSet , $thisFlags );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

                if( ! $result )
                {                   
                    Write-Error ( "Failed to set working set size for {0} pid {1} to {2}MB - {3}" -f $process.ProcessName , $process.Id , $maxWorkingSet , $LastError)
                }
                else
                {
                    $counter++
                }
            }
        }
    }
}

Write-Output -InputObject ( "{0} working sets of {1} process{2}" -f $(if( $maxWorkingSet -le 0 ) { 'Emptied' } else { 'Trimmed' }) , $counter , $(if( $counter -ne 1 ) { 'es' } ) )

