<#
  .SYNOPSIS
  This script performs a Power On on a VM
  .DESCRIPTION
  This script performs a Power On on a VM.  The script uses the CU Actions in version 8.8 to action the power-on, regardless of the underlying hypervisor.
  .PARAMETER hostname
  The hostname to perform the action on, supplied via a trigger or right-click action against a VM
  .NOTES
   Version:        0.1
   Context:        Computer, executes on Monitor
   Author:         Bill Powell
   Requires:       Realtime DX 8.8
   Creation Date:  2023-05-11

  .LINK
   https://support.controlup.com/docs/powershell-cmdlets-for-solve-actions
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = 'hostname of machine to be actioned')]
    [ValidateNotNullOrEmpty()]
    [string]$hostname
)

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

#region Define the CUAction

$RequiredAction = "Power On VM"  # or any of "Force Power Off VM","Force Reset VM","Power On VM","Restart Guest","Shutdown Guest"
$RequiredActionCategory = 'VM Power Management'

#endregion

#region Process Args

#$hostname = $args[0]   # the computername is passed as the only argument

#endregion

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

$AcceptableModules = New-Object System.Collections.Generic.List[object]
try {
    $DllsToLoad = Get-MonitorDLLs -DLLList @('ControlUp.PowerShell.User.dll')
    $DllsToLoad | Import-Module 
    $DllsToLoad -replace "^.*\\",'' -replace "\.dll$",'' | ForEach-Object {$AcceptableModules.Add($_)}
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

#region Perform CUAction on target

$ThisComputer = Get-CimInstance -ClassName Win32_ComputerSystem

Write-Output "Action $RequiredAction applied to $hostname from $($ThisComputer.Name)"

$Action = (Get-CUAvailableActions -DisplayName $RequiredAction | Where-Object {($_.Title -eq $RequiredAction) -and ($_.IsSBA -eq $false) -and ($_.Category -eq $RequiredActionCategory)})[0]

$Allrows = Invoke-CUQuery -Table $Action.Table -Fields * -Where "sName = '${hostname}'"
$Allrows.Data | ForEach-Object {
    $VMTableRow = $_

    $Result = Invoke-CUAction -ActionId $Action.Id `
                              -Table $Action.Table `
                              -RecordsGuids @($VMTableRow.key)

    Write-Output "Action $RequiredAction applied to $hostname, result $($Result.Result)"
}

#endregion

Write-Output "Done"

