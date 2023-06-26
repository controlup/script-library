<#
  .SYNOPSIS
  This script lists all the available CU Actions
  .DESCRIPTION
  This script lists all the available CU Actions for the active monitor version. The CU actions are only available from version 8.8 onwards.
  .NOTES
   Version:        0.1
   Context:        Computer, executes on Monitor
   Author:         Bill Powell
   Requires:       Realtime DX 8.8
   Creation Date:  2023-05-23

  .LINK
   https://support.controlup.com/docs/powershell-cmdlets-for-solve-actions
#>

[CmdletBinding()]

#region ControlUpScriptingStandards
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 400
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}
#endregion ControlUpScriptingStandards

#region Load the version of the module to match the running monitor and check that it has the new features

function Get-MonitorDLLs {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][string[]]$DLLList
    )
    [int]$DLLsFound = 0
    Get-CimInstance -Query "SELECT * from win32_Service WHERE Name = 'cuMonitor' AND State = 'Running'" | ForEach-Object {
        $MonitorService = $_
        if ($MonitorService.PathName -match '^(?<op>"{0,1})\b(?<text>[^"]*)\1$') {
            $Path = $Matches.text
            $MonitorFolder = Split-Path -Path $Path -Parent
            $DLLList | ForEach-Object {
                $DLLBase = $_
                $DllPath = Join-Path -Path $MonitorFolder -ChildPath $DLLBase
                if (Test-Path -LiteralPath $DllPath) {
                    $DllPath
                    $DLLsFound++
                }
                else {
                    throw "DLL $DllPath not found in running monitor folder"
                }
            }
        }
    }
    if ($DLLsFound -ne $DLLList.Count) {
        throw "cuMonitor is not installed or not running"
    }
}

$AcceptableModules = New-Object System.Collections.ArrayList
try {
    $DllsToLoad = Get-MonitorDLLs -DLLList @('ControlUp.PowerShell.User.dll')
    $DllsToLoad | Import-Module 
    $DllsToLoad -replace "^.*\\",'' -replace "\.dll$",'' | ForEach-Object {$AcceptableModules.Add($_) | Out-Null}
}
catch {
    $exception = $_
    Write-Error "Required DLLs not loaded: $($exception.Exception.Message)"
}

if (-not ((Get-Command -Name 'Invoke-CUAction' -ErrorAction SilentlyContinue).Source) -in $AcceptableModules) {
   Write-Error "ControlUp version 8.8 commands are not available on this system"
   exit 0
}

#endregion

#region Explore what commands are available

try {
    $AvailableActions = Get-CUAvailableActions
}
catch {
    $exception = $_
    Write-Error "Get-CUAvailableActions failed: $($exception.Exception.Message)"
    exit 0
}
$Tables = $AvailableActions.Table | Sort-Object -Unique

$Tables | ForEach-Object {
    $Table = $_
    Write-Output "======================================================================="
    Write-Output "Actions that operate on objects in the $Table table"
    Write-Output "======================================================================="
    $ActionsForTable = $AvailableActions | Where-Object {$_.Table -eq $Table}
    $ActionsForTable | Sort-Object -Property Category,Title | Format-Table -Property Title,Category,IsSBA,Description
}

#endregion

Write-Output "Done"

