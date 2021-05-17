<#
.SYNOPSIS
        Outputs computer applied group policies
.DESCRIPTION
        This script shows computer applied group policies as shown inside the EventLog
.PARAMETER <paramName>
        None at this point
.EXAMPLE
        <Script Path>\script.ps1
.INPUTS
        None
.OUTPUTS
        List of applied group policies
.LINK
        See http://www.controlup.com
#>

# must be run with elevated privileges

$ErrorActionPreference = "Stop"     #   another way to try to stop the script in case of errors. Important for Try/Catch usage.

# Filters by Event Id '4004' (workstations) or '4006' (servers) and the computer name
$Query = "*[EventData[Data[@Name='PrincipalSamName'] and (Data='$env:userdomain\$env:username')]] and *[System[(EventID='4004') or (EventID='4006')]]"

try {
    # Gets the last (most recent) event matching the criteria by $Query
    $Event = Get-WinEvent -ProviderName Microsoft-Windows-GroupPolicy -FilterXPath "$Query" -MaxEvents 1
    $ActivityId = $Event.ActivityId.Guid
}
catch {
    Write-Host "Could not find relevant events in the Microsoft-Windows-GroupPolicy/Operational log. `nThe default log size (4MB) may not be large enough for the volume of data saved in it. Please increase the log size to support older messages."
    Exit 1
}

# Looks for Event Id '5312' with the relevant 'Activity Id'
$message = Get-WinEvent -ProviderName Microsoft-Windows-GroupPolicy -FilterXPath "*[System[(EventID='5312')]]" | Where-Object{$_.ActivityId -eq $ActivityId}

# Displays the 'Message Property'
Write-Host $message.Message

