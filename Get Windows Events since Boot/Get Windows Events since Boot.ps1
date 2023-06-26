<#
.SYNOPSIS
    Get the number of times an event has been registered in the Windows event log since the system has booted.

.DESCRIPTION
    Get the number of times an event has been registered in the Windows event log since the system has booted.
    This can be used to decide whether to reboot a system to bring it back into a "healthy" state.
    
    This script is intended to be used within ControlUp as an action.

.EXAMPLE
    Parameters:
        EventlogName = 'Application'
        EventID = 16394

    Running the ControlUp Action using the above parameters returns the number of times the eventID 16394 has been found in the Application eventlog since the system has booted.

.PARAMETER EventlogName
    The name of the Eventlog to search in. To find valid values run the "[System.Diagnostics.EventLog]::GetEventLogs().Log" command in PowerShell.

.PARAMETER EventID
    The ID of the event to look for.

.NOTES
    Author: 
        Rein Leen
    Contributor(s):
    Context: 
        Machine
    Modification_history:
        Rein Leen       09-06-2023      Version ready for release
#>

#region [parameters]
[CmdletBinding()]
Param (
    [Parameter(Position = 0, Mandatory = $true, HelpMessage = 'The name of the Eventlog to search in.')]
    [string]$EventlogName,
    # A valid Distinguished name always contains two domainComponents.
    [Parameter(Position = 1, Mandatory = $true, HelpMessage = 'The ID of the event to look for.')]
    [string]$EventID
)
#endregion [parameters]

#region [prerequisites]
# Required dependencies
#Requires -Version 5.1

# Setting error actions
$ErrorActionPreference = 'Stop'
$DebugPreference = 'SilentlyContinue'
#endregion [prerequisites]

#region [functions]
# Function to get the ControlUp engine under which the script is running.
function Get-ControlUpEngine {
    $runtimeEngine = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $PID"
    switch ($runtimeEngine.ProcessName) {
        'cuAgent.exe' {
                return '.NET'
            }
        'powershell.exe' {
                return 'Classic'
            }
    }
}

# Function to assert the parameters are correct
function Assert-ControlUpParameter {
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [object]$Parameter,
        [Parameter(Position = 1, Mandatory = $true)]
        [boolean]$Mandatory,
        [Parameter(Position = 2, Mandatory = $true)]
        [ValidateSet('.NET','Classic')]
        [string]$Engine
    )

    # If a parameter is optional passing using a hyphen (-) or none is required when using the Classic engine. If this is the case return $null.
    if (($Mandatory -eq $false) -and (($Parameter -eq '-') -or ($Parameter -eq 'none'))) {
        return $null
    }

    # If a parameter is optional when using the .NET engine it should be empty. if this is the case return $null.
    if (($Engine -eq '.NET') -and ($Mandatory -eq $false) -and ([string]::IsNullOrWhiteSpace($Parameter))) {
        return $null
    }

    # Check if a mandatory parameter isn't null
    if (($Mandatory -eq $true) -and ([string]::IsNullOrWhiteSpace($Parameter))) {
        throw [System.ArgumentException] 'This parameter cannot be empty'
    }

    # ControlUp can add double quotes when using the .NET engine when a parameter value contains spaces. Remove these.
    if ($Engine -eq '.NET') {
        # Regex used to match double quotes
        $possiblyQuotedStringRegex = '^(?<op>"{0,1})\b(?<text>[^"]*)\1$'
        $Parameter -match $possiblyQuotedStringRegex | Out-Null
        return $Matches.text
    } else {
        return $Parameter
    }
}
#endregion [functions]

#region [variables]
$controlUpEngine = Get-ControlUpEngine

# Validate $$EventlogName
$EventlogName = Assert-ControlUpParameter -Parameter $EventlogName -Mandatory $true -Engine $controlUpEngine

# Validate $Searchbases
$EventID  = Assert-ControlUpParameter -Parameter $EventID  -Mandatory $true -Engine $controlUpEngine
#endregion [variables]

#region [actions]
$lastBootTime = (Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object LastBootUpTime).LastBootUpTime

# Validate if eventlog exists
if ($EventlogName -notin [System.Diagnostics.EventLog]::GetEventLogs().Log) {
    throw 'No eventlog found with the name {0}' -f $EventlogName
}

# Create eventlog instance
$eventlogInstance = [System.Diagnostics.EventLog]::new($EventlogName)
$eventlogEntries = $eventlogInstance.Entries | Where-Object {($_.EventID -eq $EventID) -and ($_.TimeGenerated -gt $lastBootTime)}

if ($null -ne $eventlogEntries) {
    Write-Output ('Event with eventID {0} in log {3} was found {1} times since last boot at {2}' -f $EventID, $eventlogEntries.Count, $lastBootTime, $EventlogName)
} else {
    Write-Output ('Event with eventID {0} in log {2} was not found since last boot at {1}' -f $EventID, $lastBootTime, $EventlogName)
}
#endregion [actions]

