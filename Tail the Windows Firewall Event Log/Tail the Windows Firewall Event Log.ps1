<#
    .SYNOPSIS
        Enables Windows Firewall logging than tails the event log for Firewall events

    .DESCRIPTION
        Enables Windows Firewall logging than tails the event log for Firewall events

    .PARAMETER  <ComputerName <string>>
        Name of the computer to enable/disable logging and tailing

    .PARAMETER  <IncludeApplications <string>>
        List of applications seperated by commas. Ex, svchost.exe,filezilla server.exe,SYSTEM

    .PARAMETER  <ExcludeApplications <string>>
        List of applications seperated by commas. Ex, svchost.exe,filezilla server.exe,SYSTEM. Default value: svchost.exe,SYSTEM

    .EXAMPLE
        . .\TailWindowsFirewall.ps1 -ComputerName W2019-001 -includeApplications "svchost.exe,VirtualDesktopAgent.exe" -excludeApplications ""
        Enables Windows Firewall logging than tails the event log for Firewall events, including svchost.exe and virtualdesktopagent.exe. No applications are excluded.

    .EXAMPLE
        . .\TailWindowsFirewall.ps1 -ComputerName W2019-001 -includeApplications "svchost.exe,VirtualDesktopAgent.exe"
        Enables Windows Firewall logging than tails the event log for Firewall events, including svchost.exe and virtualdesktopagent.exe. Default applications svchost.exe and SYSTEM are excluded and override the inclusion.

    .EXAMPLE
        . .\TailWindowsFirewall.ps1 -ComputerName W2019-001
        Enables Windows Firewall logging than tails the event log for Firewall events. Default applications svchost.exe and SYSTEM are excluded.


    .NOTES
        Requires remote powershell and admin privilege's on the target device

    .CONTEXT
        Console

    .MODIFICATION_HISTORY
        Created TTYE : 2020-09-22


    AUTHOR: Trentent Tye
#>
[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='Enter the SamAccountName of the Server')][ValidateNotNullOrEmpty()]  [string]$ComputerName,
    [Parameter(Mandatory=$false,HelpMessage='Applications to include delimited by a comma. Notepad.exe,Explorer.exe. Default is include everything')]    [string]$IncludeApplications = "",
    [Parameter(Mandatory=$false,HelpMessage='Applications to exclude delimited by a comma. svchost.exe,System. Default is exclude svchost.exe,System')]  [string]$ExcludeApplications = "svchost.exe,System"
)

$Script = 
@'
[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='Enter the SamAccountName of the Server')][ValidateNotNullOrEmpty()]  [string]$ComputerName,
    [Parameter(Mandatory=$false,HelpMessage='Applications to include delimited by a comma. Notepad.exe,Explorer.exe. Default is include everything')]    [string]$IncludeApplications = "",
    [Parameter(Mandatory=$false,HelpMessage='Applications to exclude delimited by a comma. svchost.exe,System. Default is exclude svchost.exe,System')]  [string]$ExcludeApplications = "svchost.exe,System"

)

#Start-Transcript "C:\firewall.txt" -Force
$ErrorActionPreference = "Continue"
Set-StrictMode -Version Latest
#$verbosePreference = "Continue"

Write-Verbose "ComputerName        : $Computername"
Write-Verbose "IncludeApplications : $IncludeApplications"
Write-Verbose "ExcludeApplications : $ExcludeApplications"

Function CheckAuditpolExitCode ([string]$ExitCode) {
    Switch ($ExitCode) {
        "0"       { Write-Verbose "Success" }
        "1314"    { Write-Error "Action requires admin privileges"; pause ; exit}
        default   { Write-Error "An error occurred."; pause ; exit}
    }
}

Function FilterBuilder ([string]$IncludeApplications, [string]$ExcludeApplications) {
    $StringBuilder = New-Object -TypeName "System.Text.StringBuilder"
    if ($IncludeApplications.Length -gt 1) {
        $apps = $IncludeApplications.split(",")
        for ($i=0;$i -lt $apps.count;$i++) {
            if ($i -eq 0) {
                [void]$StringBuilder.Append("`$_.message -like `"*$($apps[$i])*`"") 
            } else {
                [void]$StringBuilder.Append(" -or `$_.message -like `"*$($apps[$i])*`"") 
            }
        }
        Write-Verbose "Include: $($StringBuilder.ToString())"
        return [System.Management.Automation.ScriptBlock]::Create($StringBuilder.ToString())
    }
    if ($ExcludeApplications.Length -gt 1) {
        $apps = $ExcludeApplications.split(",")
        for ($i=0;$i -lt $apps.count;$i++) {
            if ($i -eq 0) {
                [void]$StringBuilder.Append("`$_.message -notlike `"*$($apps[$i])*`"") 
            } else {
                [void]$StringBuilder.Append(" -and `$_.message -notlike `"*$($apps[$i])*`"")  ##Don't ask me why this has to be an and. If you make it an or then only the first one is evaluated.
            }
        }
        Write-Verbose "Exclude: $($StringBuilder.ToString())"
        return [System.Management.Automation.ScriptBlock]::Create($StringBuilder.ToString())
    }
}

if ($IncludeApplications.Length -gt 3) {
    $Global:IncludeFilter = FilterBuilder -IncludeApplications $IncludeApplications
    Write-Verbose "Global:IncludeFilter : $($Global:IncludeFilter)"
}
if ($ExcludeApplications.Length -gt 3) {
    $Global:ExcludeFilter = FilterBuilder -ExcludeApplications $ExcludeApplications
    Write-Verbose "Global:ExcludeFilter : $($Global:ExcludeFilter)"
}

$pshost = Get-Host              # Get the PowerShell Host.
$pswindow = $pshost.UI.RawUI    # Get the PowerShell Host's UI.

$newsize = $pswindow.BufferSize # Get the UI's current Buffer Size.
$newsize.width = 131            # Set the new buffer's width to 131 columns.
$pswindow.buffersize = $newsize # Set the new Buffer Size as active.

$newsize = $pswindow.windowsize # Get the UI's current Window Size.
$newsize.width = 131            # Set the new Window Width to 131 columns.
$pswindow.windowsize = $newsize # Set the new Window Size as active.
$host.ui.RawUI.WindowTitle = "Firewall Tail : $ComputerName"

# Enable Auditing for Firewall
try {
    $FilteringPlatformPacketDrop = Invoke-Command -ComputerName $computerName -ScriptBlock {
        $ExitCode = Start-Process -FilePath auditpol.exe -ArgumentList @("/set","/subcategory:`"Filtering Platform Packet Drop`"","/success:enable","/failure:enable") -PassThru -Wait
        return $exitCode.ExitCode
    }
} catch {
    Write-Error "Failed to connect to $computername.`n`nCheck that WinRM is installed and enabled and that your account has permission to connect remotely." ; pause ; exit
}
CheckAuditpolExitCode -ExitCode $FilteringPlatformPacketDrop
try{
    $FilteringPlatformConnection = Invoke-Command -ComputerName $computerName -ScriptBlock {
        $ExitCode = Start-Process -FilePath auditpol.exe -ArgumentList @("/set","/subcategory:`"Filtering Platform Connection`"","/success:enable","/failure:enable") -PassThru -Wait
        return $exitCode.ExitCode
    }
} catch {
    Write-Error "Failed to connect to $computername.`n`nCheck that WinRM is installed and enabled and that your account has permission to connect remotely." ; pause ; exit
}
CheckAuditpolExitCode -ExitCode $FilteringPlatformConnection

# Paddings
$PaddingTime        = 13
$PaddingOperation   = 10
$PaddingPID         = 7
$PaddingApplication = 25
$PaddingDirection   = 10
$PaddingSrcAddr     = 16
$PaddingSrcPort     = 8
$PaddingDestAddr    = 16
$PaddingDestPort    = 9
$PaddingProtocol    = 7
$PaddingEventId     = 8


function Rewrite-Header {
    $CurrentForegroundColor = $host.ui.RawUI.ForegroundColor
    $currentBackgroundColor = $host.ui.RawUI.BackgroundColor
    $CurrentCursorYPosition = $host.ui.RawUI.CursorPosition.Y 
    $CurrentWindowHeight = ($host.ui.RawUI.WindowSize.Height)
    #move the cursor to the top of the screen
    $NewYCoordinate = $($host.ui.RawUI.CursorPosition.Y - ($host.ui.RawUI.WindowSize.Height-1))
    Write-Verbose "`n`n$NewYCoordinate`n`n"
    if ($NewYCoordinate -lt 0) {
        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0,0
    } else {
        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0,$NewYCoordinate
    }
    [Console]::ForegroundColor = [System.ConsoleColor]::White ; [Console]::BackgroundColor = [System.ConsoleColor]::Black  ; [Console]::Write("$("EventId".PadRight($PaddingEventId))$("Time".PadRight($PaddingTime))$("Operation".PadRight($PaddingOperation))$("PID".PadRight($PaddingPID))$("Application".PadRight($PaddingApplication))$("Direction".PadRight($PaddingDirection))$("SrcAddr".PadRight($PaddingSrcAddr))$("SrcPort".PadRight($PaddingSrcPort))$("DestAddr".PadRight($PaddingDestAddr))$("DestPort".PadRight($PaddingDestPort))$("Protocol".PadRight($PaddingProtocol))")
    $host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0,$($CurrentCursorYPosition)
    [Console]::ForegroundColor = [System.ConsoleColor]::White ; [Console]::BackgroundColor = [System.ConsoleColor]::Black  ; [Console]::Write("  Press `"Enter`" to stop tailing the firewall")
    $host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0,$($CurrentCursorYPosition)
    [Console]::ForegroundColor = [System.ConsoleColor]::$CurrentForegroundColor
    [Console]::BackgroundColor = [System.ConsoleColor]::$currentBackgroundColor 
}

function Get-Protocol ([int]$protocolNumber) {
    switch ($protocolNumber) {
        "1" { Write-Output "ICMP" }
        "3" { Write-Output "GGP" }
        "6" { Write-Output "TCP" }
        "8" { Write-Output "EGP" }
        "12" { Write-Output "PUP" }
        "17" { Write-Output "UDP" }
        "20" { Write-Output "HMP" }
        "27" { Write-Output "RDP" }
        "46" { Write-Output "RSVP" }
        "47" { Write-Output "PPTP data over GRE" }
        "50" { Write-Output "AH" }
        "66" { Write-Output "RVD" }
        "88" { Write-Output "IGMP" }
        "89" { Write-Output "OSPF" }
        default { Write-Output "$protocolNumber" }
    }
}

# Get first event:
Function TailWinEvents ([string]$ComputerName, [string]$IncludeApplications, [string]$ExcludeApplications){
    $firstEvent = Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{logname='security';id=5154,5156,5158,5031,5150,5151,5152,5153,5155,5157,5159} -MaxEvents 1
    $firstEventTime = $firstEvent.TimeCreated
    $newEvents = Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{logname='security';id=5154,5156,5158,5031,5150,5151,5152,5153,5155,5157,5159;StartTime=$($firstEvent.TimeCreated)}
    $lastEvent = $newEvents[0]
    $latestRecordId = 0
    $latestTime = $lastEvent.TimeCreated
    while ($true) {
        #break out of the loop if a key is pressed.
        if ([Console]::KeyAvailable)
        {
            # read the key, and consume it so it won't
            # be echoed to the console:
            $keyInfo = [Console]::ReadKey($true)
            # exit loop
            if ($keyInfo.Key -eq "Enter") { break }
        }
        Write-Verbose "Latest Record ID: $latestRecordId  -- $($newEvents[0].RecordId) NewestEventRecordID "
        if ($latestRecordId -ne $newEvents[0].RecordId) {
            foreach ($event in ($newEvents | Sort-Object -Property RecordId)) {
            Write-Verbose "IncludeApplications : $($IncludeApplications.Length)"
            if ($IncludeApplications.Length -gt 3) {
                Write-Verbose "Filtering Include"
                #$event.message
                $event = $event| Where-Object -FilterScript $Global:IncludeFilter
            }
            Write-Verbose "event after filter Count: $($($event | Measure-object).count)"

            if ($($($event | Measure-object).count) -eq 0) {
                continue
            }
            
            if ($ExcludeApplications.Length -gt 3) {
                Write-Verbose "Filtering Exclude"
                #$event.message
                $event = $event | Where-Object -FilterScript $Global:ExcludeFilter
                
            }
            Write-Verbose "event after filter Count: $($($event | Measure-object).count)"
            if ($($($event | Measure-object).count) -eq 0) {
                continue
            }

            switch ($event.id){
                "5154" {
                    $eventProperties = [pscustomobject]@{
                
                        Time= $event.TimeCreated.ToString("HH:mm:ss.fff")
                        AllowDrop = "ALLOW"
                        ProcessId= [string]$event.Properties[0].value
                        Application = $event.Properties[1].value.split("\")[-1]
                        Direction = switch ($event.Properties[6].Value)
                        {
                            "%%14609" {Write-Output "Listen"}
                            default   {Write-Output "Unknown"}
                        }
                        #$(if ($event.Properties[2].Value -eq "%%14592") {Write-Output "Inbound"} else {Write-Output "Unknown"} )
                        SourceAddr = [string]$event.Properties[2].value
                        SourcePort = [string]$event.Properties[3].value
                        DestAddr = "NA"
                        DestPort = "NA"
                        Protocol = [string]$event.Properties[4].value
                    }
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow    ; [Console]::Write("$($($event.Id).ToString().PadRight($PaddingEventId))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White     ; [Console]::Write("$($eventProperties.time.PadRight($PaddingTime))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Green     ; [Console]::Write("$($eventProperties.AllowDrop.PadRight($PaddingOperation))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Cyan      ; [Console]::Write("$($eventProperties.ProcessId.ToString().PadRight($PaddingPID))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow    ; [Console]::Write("$($eventProperties.Application.PadRight($PaddingApplication))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Gray      ; [Console]::Write("$($eventProperties.Direction.PadRight($PaddingDirection))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White     ; [Console]::Write("$($eventProperties.SourceAddr.PadRight($PaddingSrcAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Magenta   ; [Console]::Write("$($eventProperties.SourcePort.PadRight($PaddingSrcPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::DarkBlue  ; [Console]::Write("$($eventProperties.DestAddr.PadRight($PaddingDestAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::DarkBlue  ; [Console]::Write("$($eventProperties.DestPort.PadRight($PaddingDestPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Green     ; [Console]::WriteLine($(Get-Protocol -ProtocolNumber "$($eventProperties.Protocol.ToString().PadRight($PaddingProtocol))"))
                    #Write-Output "$($(Format-Table -InputObject $eventProperties -HideTableHeaders -Property $a| Out-String).replace("`n"," "))"
                } 
                "5156" {
                    $eventProperties = [pscustomobject]@{
                
                        Time= $event.TimeCreated.ToString("HH:mm:ss.fff")
                        AllowDrop = "ALLOW"
                        ProcessId= [string]$event.Properties[0].value
                        Application = $event.Properties[1].value.split("\")[-1]
                        Direction = switch ($event.Properties[2].Value)
                        {
                            "%%14592" {Write-Output "Inbound"}
                            "%%14593" {Write-Output "Outbound"}
                            default   {Write-Output "Unknown"}
                        }
                        #$(if ($event.Properties[2].Value -eq "%%14592") {Write-Output "Inbound"} else {Write-Output "Unknown"} )
                        SourceAddr = [string]$event.Properties[3].value
                        SourcePort = [string]$event.Properties[4].value
                        DestAddr = [string]$event.Properties[5].value
                        DestPort = [string]$event.Properties[6].value
                        Protocol = [string]$event.Properties[7].value
                    }
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($($event.Id).ToString().PadRight($PaddingEventId))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.time.PadRight($PaddingTime))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Green    ; [Console]::Write("$($eventProperties.AllowDrop.PadRight($PaddingOperation))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Cyan     ; [Console]::Write("$($eventProperties.ProcessId.ToString().PadRight($PaddingPID))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.Application.PadRight($PaddingApplication))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Gray     ; [Console]::Write("$($eventProperties.Direction.PadRight($PaddingDirection))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.SourceAddr.PadRight($PaddingSrcAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Magenta  ; [Console]::Write("$($eventProperties.SourcePort.PadRight($PaddingSrcPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.DestAddr.PadRight($PaddingDestAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.DestPort.PadRight($PaddingDestPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Green    ; [Console]::WriteLine($(Get-Protocol -ProtocolNumber "$($eventProperties.Protocol.ToString().PadRight($PaddingProtocol))"))
                    #Write-Output "$($(Format-Table -InputObject $eventProperties -HideTableHeaders -Property $a| Out-String).replace("`n"," "))"
                } 
                "5158" {
                    $eventProperties = [pscustomobject]@{
                
                        Time= $event.TimeCreated.ToString("HH:mm:ss.fff")
                        AllowDrop = "ALLOW"
                        ProcessId= [string]$event.Properties[0].value
                        Application = $event.Properties[1].value.split("\")[-1]
                        Direction = "NA"
                        SourceAddr = [string]$event.Properties[2].value
                        SourcePort = [string]$event.Properties[3].value
                        DestAddr = "NA"
                        DestPort = "NA"
                        Protocol = [string]$event.Properties[4].value
                    }
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($($event.Id).ToString().PadRight($PaddingEventId))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.time.PadRight($PaddingTime))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Green    ; [Console]::Write("$($eventProperties.AllowDrop.PadRight($PaddingOperation))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Cyan     ; [Console]::Write("$($eventProperties.ProcessId.ToString().PadRight($PaddingPID))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.Application.PadRight($PaddingApplication))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Gray     ; [Console]::Write("$($eventProperties.Direction.PadRight($PaddingDirection))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.SourceAddr.PadRight($PaddingSrcAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Magenta  ; [Console]::Write("$($eventProperties.SourcePort.PadRight($PaddingSrcPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.DestAddr.PadRight($PaddingDestAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.DestPort.PadRight($PaddingDestPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Green    ; [Console]::WriteLine($(Get-Protocol -ProtocolNumber "$($eventProperties.Protocol.ToString().PadRight($PaddingProtocol))"))
                    }
                "5031" {$eventProperties = [pscustomobject]@{
                
                        Time= $event.TimeCreated.ToString("HH:mm:ss.fff")
                        AllowDrop = "BLOCKED"
                        ProcessId= "-"
                        Application = $event.Properties[1].value.split("\")[-1]
                    }
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow  ; [Console]::Write("$($($event.Id).ToString().PadRight($PaddingEventId))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White   ; [Console]::Write("$($eventProperties.time.PadRight($PaddingTime))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Magenta ; [Console]::Write("$($eventProperties.AllowDrop.PadRight($PaddingOperation))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Cyan    ; [Console]::Write("$($eventProperties.ProcessId.ToString().PadRight($PaddingPID))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow  ; [Console]::WriteLine("$($eventProperties.Application.PadRight($PaddingApplication))")
                    }
                "5150" {$eventProperties = [pscustomobject]@{
                
                        Time= $event.TimeCreated.ToString("HH:mm:ss.fff")
                        AllowDrop = "DROP"
                        ProcessId= "NA"
                        Application = "NA"
                        Direction = switch ($event.Properties[0].Value)
                        {
                            "%%14592" {Write-Output "Inbound"}
                            "%%14593" {Write-Output "Outbound"}
                            default   {Write-Output "Unknown"}
                        }
                        SourceAddr = [string]$event.Properties[1].value
                        SourcePort = "NA"
                        DestAddr = [string]$event.Properties[2].value
                        DestPort = "NA"
                        Protocol = "NA"
                    }
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($($event.Id).ToString().PadRight($PaddingEventId))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.time.PadRight($PaddingTime))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.AllowDrop.PadRight($PaddingOperation))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Cyan     ; [Console]::Write("$($eventProperties.ProcessId.ToString().PadRight($PaddingPID))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.Application.PadRight($PaddingApplication))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Gray     ; [Console]::Write("$($eventProperties.Direction.PadRight($PaddingDirection))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.SourceAddr.PadRight($PaddingSrcAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Magenta  ; [Console]::Write("$($eventProperties.SourcePort.PadRight($PaddingSrcPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.DestAddr.PadRight($PaddingDestAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.DestPort.PadRight($PaddingDestPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Green    ; [Console]::WriteLine($(Get-Protocol -ProtocolNumber "$($eventProperties.Protocol.ToString().PadRight($PaddingProtocol))"))
                    }
                "5151" {$eventProperties = [pscustomobject]@{
                
                        Time= $event.TimeCreated.ToString("HH:mm:ss.fff")
                        AllowDrop = "DROP"
                        ProcessId= "NA"
                        Application = "NA"
                        Direction = switch ($event.Properties[0].Value)
                        {
                            "%%14592" {Write-Output "Inbound"}
                            "%%14593" {Write-Output "Outbound"}
                            default   {Write-Output "Unknown"}
                        }
                        SourceAddr = [string]$event.Properties[1].value
                        SourcePort = "NA"
                        DestAddr = [string]$event.Properties[2].value
                        DestPort = "NA"
                        Protocol = "NA"
                    }
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($($event.Id).ToString().PadRight($PaddingEventId))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.time.PadRight($PaddingTime))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.AllowDrop.PadRight($PaddingOperation))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Cyan     ; [Console]::Write("$($eventProperties.ProcessId.ToString().PadRight($PaddingPID))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.Application.PadRight($PaddingApplication))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Gray     ; [Console]::Write("$($eventProperties.Direction.PadRight($PaddingDirection))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.SourceAddr.PadRight($PaddingSrcAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Magenta  ; [Console]::Write("$($eventProperties.SourcePort.PadRight($PaddingSrcPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.DestAddr.PadRight($PaddingDestAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.DestPort.PadRight($PaddingDestPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Green    ; [Console]::WriteLine($(Get-Protocol -ProtocolNumber "$($eventProperties.Protocol.ToString().PadRight($PaddingProtocol))"))
                    }
                "5152" {$eventProperties = [pscustomobject]@{
                
                        Time= $event.TimeCreated.ToString("HH:mm:ss.fff")
                        AllowDrop = "DROP"
                        ProcessId= [string]$event.Properties[0].value
                        Application = $event.Properties[1].value.split("\")[-1]
                        Direction = switch ($event.Properties[2].Value)
                        {
                            "%%14592" {Write-Output "Inbound"}
                            "%%14593" {Write-Output "Outbound"}
                            default   {Write-Output "Unknown"}
                        }
                        SourceAddr = [string]$event.Properties[3].value
                        SourcePort = [string]$event.Properties[4].value
                        DestAddr = [string]$event.Properties[5].value
                        DestPort = [string]$event.Properties[6].value
                        Protocol = [string]$event.Properties[7].value
                    }
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($($event.Id).ToString().PadRight($PaddingEventId))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.time.PadRight($PaddingTime))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.AllowDrop.PadRight($PaddingOperation))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Cyan     ; [Console]::Write("$($eventProperties.ProcessId.ToString().PadRight($PaddingPID))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.Application.PadRight($PaddingApplication))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Gray     ; [Console]::Write("$($eventProperties.Direction.PadRight($PaddingDirection))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.SourceAddr.PadRight($PaddingSrcAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Magenta  ; [Console]::Write("$($eventProperties.SourcePort.PadRight($PaddingSrcPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.DestAddr.PadRight($PaddingDestAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.DestPort.PadRight($PaddingDestPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Green    ; [Console]::WriteLine($(Get-Protocol -ProtocolNumber "$($eventProperties.Protocol.ToString().PadRight($PaddingProtocol))"))
                    }
                "5153" {$eventProperties = [pscustomobject]@{
                
                        Time= $event.TimeCreated.ToString("HH:mm:ss.fff")
                        AllowDrop = "BLOCKED"
                        ProcessId= [string]$event.Properties[0].value
                        Application = $event.Properties[1].value.split("\")[-1]
                        Direction = "NA"
                        SourceAddr = [string]$event.Properties[2].value
                        SourcePort = [string]$event.Properties[3].value
                        DestAddr = "NA"
                        DestPort = "NA"
                        Protocol = [string]$event.Properties[4].value
                    }
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($($event.Id).ToString().PadRight($PaddingEventId))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.time.PadRight($PaddingTime))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.AllowDrop.PadRight($PaddingOperation))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Cyan     ; [Console]::Write("$($eventProperties.ProcessId.ToString().PadRight($PaddingPID))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.Application.PadRight($PaddingApplication))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Gray     ; [Console]::Write("$($eventProperties.Direction.PadRight($PaddingDirection))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.SourceAddr.PadRight($PaddingSrcAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Magenta  ; [Console]::Write("$($eventProperties.SourcePort.PadRight($PaddingSrcPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.DestAddr.PadRight($PaddingDestAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.DestPort.PadRight($PaddingDestPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Green    ; [Console]::WriteLine($(Get-Protocol -ProtocolNumber "$($eventProperties.Protocol.ToString().PadRight($PaddingProtocol))"))
                    }
                "5155" {$eventProperties = [pscustomobject]@{
                
                        Time= $event.TimeCreated.ToString("HH:mm:ss.fff")
                        AllowDrop = "BLOCKED"
                        ProcessId= [string]$event.Properties[0].value
                        Application = $event.Properties[1].value.split("\")[-1]
                        Direction = "NA"
                        SourceAddr = [string]$event.Properties[2].value
                        SourcePort = [string]$event.Properties[3].value
                        DestAddr = "NA"
                        DestPort = "NA"
                        Protocol = [string]$event.Properties[4].value
                    }
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($($event.Id).ToString().PadRight($PaddingEventId))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.time.PadRight($PaddingTime))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.AllowDrop.PadRight($PaddingOperation))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Cyan     ; [Console]::Write("$($eventProperties.ProcessId.ToString().PadRight($PaddingPID))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.Application.PadRight($PaddingApplication))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Gray     ; [Console]::Write("$($eventProperties.Direction.PadRight($PaddingDirection))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.SourceAddr.PadRight($PaddingSrcAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Magenta  ; [Console]::Write("$($eventProperties.SourcePort.PadRight($PaddingSrcPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.DestAddr.PadRight($PaddingDestAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.DestPort.PadRight($PaddingDestPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Green    ; [Console]::WriteLine($(Get-Protocol -ProtocolNumber "$($eventProperties.Protocol.ToString().PadRight($PaddingProtocol))"))
                    }
                "5157" {$eventProperties = [pscustomobject]@{
                
                        Time= $event.TimeCreated.ToString("HH:mm:ss.fff")
                        AllowDrop = "DROP"
                        ProcessId= [string]$event.Properties[0].value
                        Application = $event.Properties[1].value.split("\")[-1]
                        Direction = switch ($event.Properties[2].Value)
                        {
                            "%%14592" {Write-Output "Inbound"}
                            "%%14593" {Write-Output "Outbound"}
                            default   {Write-Output "Unknown"}
                        }
                        SourceAddr = [string]$event.Properties[3].value
                        SourcePort = [string]$event.Properties[4].value
                        DestAddr = [string]$event.Properties[5].value
                        DestPort = [string]$event.Properties[6].value
                        Protocol = [string]$event.Properties[7].value
                    }
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($($event.Id).ToString().PadRight($PaddingEventId))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.time.PadRight($PaddingTime))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.AllowDrop.PadRight($PaddingOperation))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Cyan     ; [Console]::Write("$($eventProperties.ProcessId.ToString().PadRight($PaddingPID))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.Application.PadRight($PaddingApplication))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Gray     ; [Console]::Write("$($eventProperties.Direction.PadRight($PaddingDirection))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.SourceAddr.PadRight($PaddingSrcAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Magenta  ; [Console]::Write("$($eventProperties.SourcePort.PadRight($PaddingSrcPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.DestAddr.PadRight($PaddingDestAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.DestPort.PadRight($PaddingDestPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Green    ; [Console]::WriteLine($(Get-Protocol -ProtocolNumber "$($eventProperties.Protocol.ToString().PadRight($PaddingProtocol))"))
                    }
                "5159" {$eventProperties = [pscustomobject]@{
                
                        Time= $event.TimeCreated.ToString("HH:mm:ss.fff")
                        AllowDrop = "BLOCKED"
                        ProcessId= [string]$event.Properties[0].value
                        Application = $event.Properties[1].value.split("\")[-1]
                        Direction = "NA"
                        SourceAddr = [string]$event.Properties[2].value
                        SourcePort = [string]$event.Properties[3].value
                        DestAddr = "NA"
                        DestPort = "NA"
                        Protocol = [string]$event.Properties[4].value
                    }
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($($event.Id).ToString().PadRight($PaddingEventId))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.time.PadRight($PaddingTime))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.AllowDrop.PadRight($PaddingOperation))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Cyan     ; [Console]::Write("$($eventProperties.ProcessId.ToString().PadRight($PaddingPID))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.Application.PadRight($PaddingApplication))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Gray     ; [Console]::Write("$($eventProperties.Direction.PadRight($PaddingDirection))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::White    ; [Console]::Write("$($eventProperties.SourceAddr.PadRight($PaddingSrcAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Magenta  ; [Console]::Write("$($eventProperties.SourcePort.PadRight($PaddingSrcPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Yellow   ; [Console]::Write("$($eventProperties.DestAddr.PadRight($PaddingDestAddr))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Red      ; [Console]::Write("$($eventProperties.DestPort.PadRight($PaddingDestPort))")
                    [Console]::ForegroundColor = [System.ConsoleColor]::Green    ; [Console]::WriteLine($(Get-Protocol -ProtocolNumber "$($eventProperties.Protocol.ToString().PadRight($PaddingProtocol))"))
                    }
                }
            }
            Rewrite-Header
        }
        $latestRecordId = $newEvents[0].RecordId
        $latestTime = $newEvents[0].TimeCreated
        $Index2  = Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{logname='security';id=5154,5156,5158,5031,5150,5151,5152,5153,5155,5157,5159;StartTime=$latestTime} -MaxEvents 1
        
        #Write-Verbose "Sleep 1 second"
        Start-Sleep -Seconds 1

        Write-Verbose "LatestRecordId : $latestRecordId"
        Write-Verbose "latestTime     : $latestTime"
        Write-Verbose "Index2recordId : $($index2.RecordId)"
        if ($Index2.RecordId -gt $latestRecordId)
        {
            Write-Verbose "Index2 greater than latestRecordId!"

            Start-Sleep 1
            $newEvents = Get-WinEvent  -ComputerName $ComputerName -LogName Security -FilterXPath "*[System[(EventRecordID>=$latestRecordId) and (EventID=5154 or EventID=5156 or EventID=5158 or EventID=5031 or EventID=5150 or EventID=5151 or EventID=5152 or EventID=5153 or EventID=5155 or EventID=5157 or EventID=5159)]]"
            
            Write-Verbose "latestRecordId : $latestRecordId"
            Write-Verbose "Newevents Count: $($($NewEvents | Measure-object).count)"
        }
    }
}

TailWinEvents -ComputerName $ComputerName -includeApplications $IncludeApplications -excludeApplications $ExcludeApplications

# Disable Auditing for Firewall
write-Host "`n`nDisabling firewall logging..."
auditpol.exe /set /subcategory:"Filtering Platform Packet Drop" /success:disable /failure:disable
auditpol.exe /set /subcategory:"Filtering Platform Connection" /success:disable /failure:disable
Sleep 3
'@

$Script | Out-file "$env:temp\TailEventLog.ps1" -Force
Start-Process -filepath powershell.exe -argumentList @("-noprofile","-executionpolicy ByPass","-file `"$env:temp\TailEventLog.ps1`" $computerName `"$IncludeApplications`" `"$ExcludeApplications`"")
