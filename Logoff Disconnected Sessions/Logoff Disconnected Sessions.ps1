<#
  .SYNOPSIS
  This script logs off all disconnected sessions for a user
  .DESCRIPTION
  The script logs off all disconnected sessions for a user. The script uses the CU Actions in version 8.8 to action all sessions, regardless of connection type
  .PARAMETER UserAccount
  The hostname to perform the action on, supplied via a trigger or right-click action against a session
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
param (
    [Parameter(Mandatory = $true, HelpMessage = 'user account of machine to be actioned')]
    [ValidateNotNullOrEmpty()]
    [string]$UserAccount
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

$RequiredAction = "LogOff Session"
$RequiredActionCategory = 'Remote Desktop Services'

#endregion

#region Process Args

#$UserAccount = $args[0] # e.g. 'CUEMEA\billp'

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

Write-Output "Action $RequiredAction applied to $UserAccount from $($ThisComputer.Name)"

$Action = (Get-CUAvailableActions -DisplayName $RequiredAction | Where-Object {($_.Title -eq $RequiredAction) -and ($_.IsSBA -eq $false) -and ($_.Category -eq $RequiredActionCategory)})[0]

Write-Output "Action title '$($Action.Title)', ID $($Action.Id)" 

#
# in theory, a user could have 100s of disconnected session, so the code fetches chunks
# (this makes the script easy to adapt to just clear down *all* disconnected sessions, for example)

function Get-CUData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][string]$TableName,
        [Parameter(Mandatory=$false)][string]$Where,
        [Parameter(Mandatory=$false)][string[]]$FieldList,
        [Parameter(Mandatory=$false)][int]$ChunkSize = 100
    )
    if ($FieldList.Count -lt 1) {
        $FieldList = '*'
    }
    $ParamSplat = @{
        Table = $TableName;
        Fields = $FieldList;
        Take = $ChunkSize;
        Skip = 0;
    }
    if (-not [string]::IsNullOrWhiteSpace($Where)) {
        $ParamSplat["Where"] = $Where
    }
    do {
        $CUQueryResult = Invoke-CUQuery @ParamSplat
        $CUQueryResult.Data
        $ParamSplat["Skip"] += $chunkSize
    } while ($ParamSplat["Skip"] -lt $CUQueryResult.Total)
}

[System.Collections.Generic.List[object]]$SessionList = Get-CUData -TableName $Action.Table `
                                                        -Where "sUserAccount = '$UserAccount'"

$SessionList | 
  Where-Object {$_.eConnectState -eq 4} | # only disconnected sessions, omit for all sessions
  ForEach-Object {
    $Session = $_
    Invoke-CUAction -ActionId $Action.Id `
                    -Table $Action.Table `
                    -RecordsGuids @($Session.key) | Out-Null
    Write-Output "Logoff session for user $($Session.sUserAccount) on host $($Session.sServerName)"
}

#endregion

Write-Output "Done"

