#Requires -version 3.0
<#
    .SYNOPSIS
    Displays local time, timezone and sync settings.

    .DESCRIPTION
    Gets the local time, timezone details and the settings used by the Win32Time service for synchronization from the machine.

    .NOTES
    Version:        1.0
    Author:         Ton de Vreede
#>

# Get time, timezone and sync settings
[datetime]$dtNow = Get-Date
[pscustomobject]$objTimeSettings = Get-ItemProperty HKLM:\System\CurrentControlSet\Services\W32Time\Parameters -name NtpServer, Type
[TimeZoneInfo]$tzLocal = [System.TimeZoneInfo]::Local

# Write time and timezone info
[string]$strTimeZone = ($tzLocal.Id + " (" + $tzLocal.DisplayName.Split(')')[0].TrimStart('(') + ")"  )

If ($dtNow.IsDaylightSavingTime()) {
    $strTimeZone += ', currently in Daylight Saving Time'
}

Write-Output -InputObject $dtNow.ToString("HH:mm:ss")
Write-Output -InputObject $strTimeZone

# Check and report on 'Set the time zone automatically'
If ((Get-ItemProperty HKLM:\System\CurrentControlSet\Services\tzautoupdate -name Start).Start -ne '4') {
    Write-Output -InputObject 'Timezone is set automatically by the tzautoupdate service'
}
Else {
    Write-Output -InputObject 'Timezone NOT set automatically'
}


# Output sync method
[string]$strSync = ''
Switch ($objTimeSettings.Type) {
    'NoSync' { $strSync = "The time service does not synchronize with other sources." }
    'NTP' { $strSync = "The time service synchronizes using NTP with following server(s):`n$($objTimeSettings.NtpServer.Split(' ') | Foreach-Object {$_.Split(',')[0]})" }
    'NT5DS' { $strSync = "The time service synchronizes from the domain hierarchy with the following server(s):`n$($objTimeSettings.NtpServer.Split(' ') | Foreach-Object {$_.Split(',')[0]})" }
    'AllSync' { $strSync = "The time service uses all the available synchronization mechanisms with the following server(s):`n$($objTimeSettings.NtpServer.Split(' ') | Foreach-Object {$_.Split(',')[0]})" }
}

Write-Output -InputObject $strSync

Exit 0
