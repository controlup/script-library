<#
    .SYNOPSIS
        Adjusts the process priority based on session state. 

    .DESCRIPTION
        Adjusts process priority for the session based on the session state.  Session state can be idle, disconnected or active.
        Idle is determined by a combination of Active + Idle time, whereas Active is determined when idle time = 0.

        Active will always equal 0 and any value above that means it's been idle and will cause the lowering of process priority

    .EXAMPLE
        . .\SetProcessPriority.ps1 4 BOTTHEORY\amttye Active 15 BelowNormal
        A session in the active state has been idle for 15 minutes, process priority will be adjusted down to BelowNormal

    .EXAMPLE
        . .\SetProcessPriority.ps1 4 CONTROLUP\trententt Disconnected 0 Idle cmd/powershell
        A session in the disconnect state and process priority will be adjusted down to Idle, except for cmd.exe and powershell.exe

    .EXAMPLE
        . .\SetProcessPriority.ps1 4 LABRAT\guyrleech Active 0 Idle
        A session is in a active state and process priority will be adjusted back to preconfigured values or Normal

    .CONTEXT
        Session

    .MODIFICATION_HISTORY
        Created TTYE : 2019-07-25


    AUTHOR: Trentent Tye
#>


[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='Session ID')][ValidateNotNullOrEmpty()]                                                                [int]$SessionID,
    [Parameter(Mandatory=$true,HelpMessage='Username in the format %DOMAIN%\%username%')][ValidateNotNullOrEmpty()]                                [string]$UserName,
    [Parameter(Mandatory=$true,HelpMessage='Session State')][ValidateNotNullOrEmpty()][ValidateSet("Active","Disconnected")]                       [string]$SessionState,
    [Parameter(Mandatory=$true,HelpMessage='Idle Time (in minutes)')][ValidateNotNullOrEmpty()]                                                    [int]$IdleTime,
    [Parameter(Mandatory=$true,HelpMessage='Process Priority Floor')][ValidateSet("Normal","Idle","High","RealTime","BelowNormal","AboveNormal")]  [string]$Priority,
    [Parameter(Mandatory=$false,HelpMessage='Excluded Processes - without ".exe", and seperated by forward slash eg "cmd/winlogon/powershell"')]   [string]$ExcludedProcesses
)


Set-StrictMode -Version Latest
###$ErrorActionPreference = "Stop"

#Get Processes 
$VerbosePreference = "silentlycontinue"  #change to continue when you need to debug
[array]$processes = @(Get-Process | Where-Object {$_.SessionId -eq $SessionID} -ErrorAction SilentlyContinue)

#Get user sid
$sid = (New-Object System.Security.Principal.NTAccount($username)).Translate([System.Security.Principal.SecurityIdentifier]).value

#map HKU so we can access it
$null = New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS -ErrorAction SilentlyContinue

function Export-ProcessPriorityState ($ProcessList) {
    $sessionPriority = $ProcessList | Select-Object -Property ProcessName,Id,PriorityClass
    New-ItemProperty -Path "HKU:\$sid\Volatile Environment\$sessionID\" -PropertyType MultiString -Name ProcessPriorityState -Value ($sessionPriority | ConvertTo-Csv) -Force | Out-Null
}

if (-not(Test-Path "HKU:\$sid")) {
    Write-Error -Message "Did not map the users hive, storing previous process priorities will not occur"
}

# borrowed Test-RegistryValue from the great https://www.jonathanmedd.net/2014/02/testing-for-the-presence-of-a-registry-key-and-value.html
function Test-RegistryValue {
    param (
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$Path,
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$Value
    )
    try {
        Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

#if session is idle or in a disconnected state, lower process priority
if ((($SessionState -eq "Active") -and ($IdleTime -ge "1")) -or ($SessionState -eq "Disconnected")) {
    Write-Host "Found session in an idle or disconnected state.  Lowering processes in $sessionId to $priority"
    If (-not(Test-RegistryValue -Path "HKU:\$sid\Volatile Environment\$sessionId" -value ProcessPriorityState)) { Export-ProcessPriorityState -ProcessList $processes }
    foreach ($process in $processes) {
        #skip lowering priority on excluded processes
        if (($excludedProcesses -split "/") -like "$($process.ProcessName)") {
            write-host "Excluded from process reprioritization: $($Process.ProcessName) - $($Process.Id)"
        } else {
            try {
                $Process.priorityclass = "$priority" #Ignore error output for some per-user processes that are controlled by the SYSTEM (csrss)
                Write-Verbose -Message "Reduced priority: $($Process.ProcessName) - $($Process.Id) - $($Process.priorityclass)"
            } catch {
                Write-Verbose -Message "Unable to reduce priority for  $($Process.ProcessName) - $($Process.Id)"
            }
        }
    }
}

#Active Session -> Restore process priority
if (($SessionState -eq "Active") -and ($IdleTime -eq "0")) {
    Write-Host "Session $SessionId is an active state, adjusting process priorities to Normal"
    If (Test-RegistryValue -Path "HKU:\$sid\Volatile Environment\$sessionId" -value ProcessPriorityState) {
        Write-Verbose -Message "Process History found"
        #retrieve previous values
        $previousProcessState = (Get-ItemProperty -Path "HKU:\$sid\Volatile Environment\$sessionID\").ProcessPriorityState | ConvertFrom-Csv #Get-ItemPropertyValue does not exist in PS2
        foreach ($process in $previousProcessState) {
            try {
                (Get-Process -Id ($process.Id) -ErrorAction SilentlyContinue).PriorityClass = "$($process.PriorityClass)"
                Write-Verbose -Message "Process priority: $($Process.ProcessName) - $($Process.Id) - $($Process.priorityclass)"
            } catch {
                Write-Verbose -Message "Unable to elevate priority for  $($Process.ProcessName) - $($Process.Id)"
            }
        }
        Remove-ItemProperty -Path "HKU:\$sid\Volatile Environment\$sessionID\" -Name ProcessPriorityState
    } else {
        #no process history so defaulting to set processes to normal priority
        Write-Verbose -Message "No Process History found"
        foreach ($process in $processes) { 
            try {
                (Get-Process -Id ($process.Id)).PriorityClass = "Normal"
                Write-Verbose -Message "Process priority: $($Process.ProcessName) - $($Process.Id) - $($Process.priorityclass)"
            } catch {
                Write-Verbose -Message "Unable to elevate priority for  $($Process.ProcessName) - $($Process.Id)"
            }
        }
    }
}

