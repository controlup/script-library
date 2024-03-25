#region ControlUpScriptingStandards
<#
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'
#>

[int]$outputWidth = 400
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}
#endregion ControlUpScriptingStandards

#region Load the latest version of the module and check that it has the new features

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

$sites = Get-CUSites
foreach ($site in $sites) {
    $siteid = $site.id
    $sitename = $site.name
    [array]$data = (Invoke-CUQuery -Scheme coordinator -Table Nodes -Fields * -Sort ObjectGuid -Search $siteid -SearchField SiteId).data

    $totalcapacity = ($data.capacity | Measure-Object -sum).sum
    $totalusage = [math]::Round((($data.CurrentUsage | Measure-Object -sum).sum), 2)
    $usedpercentage = ($totalusage / $totalcapacity).ToString("P")
    $connectedmachines = $data | Where-Object { $_.status -eq 1 } | Measure-Object | Select-Object -ExpandProperty Count
    $disconnectedmachines = $data | Where-Object { $_.status -eq 0 }  | Measure-Object | Select-Object -ExpandProperty Count
    $rejoiningmachines = $data | Where-Object { $_.status -eq 2 }  | Measure-Object | Select-Object -ExpandProperty Count

    Write-Output "###### Monitor Cluster Status for ControlUp Site: $sitename ######"
    Write-Output "Total Capacity: $totalcapacity"
    Write-Output "Total Used Capacity: $totalusage"
    Write-Output "Total Used Capacity % : $usedpercentage"
    Write-Output "Connected Monitors: $connectedmachines"
    Write-Output "Disconnected Monitors: $disconnectedmachines"
    Write-Output "Rejoining Monitors: $rejoiningmachines"

}

