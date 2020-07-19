<# RDP session - connection properties   #>

$ErrorActionPreference = "SilentlyContinue"         # prevent error display

# get the rdp-tcp#
$session=$args[0]
#$session="RDP-Tcp#9" # Example
$rdpSession="Connection $session created"    # Build the string to search to find the ActivityID

#Get the ActivityID for that RDP session
$CorrelationId=(Get-WinEvent Microsoft-Windows-RemoteDesktopServices-RdpCoreTS/Operational `
             | Where-Object -Property Message -Match $rdpSession )[0].ActivityID

if($error.count -ge 1) {                                        # if previous command errored...
     Write-host "Required RDP log not found"
     exit                                                                   # Bailing out if the RDP 8+ log not found
     }

Get-WinEvent Microsoft-Windows-RemoteDesktopServices-RdpCoreTS/Operational `
             | ?{$_.ActivityID -eq $CorrelationId} `
             | Where-Object -Property Message -Match "client operating system type|Microsoft::Windows::RDS::Graphics|client supports version|Client not supported" `
             |Sort-Object -Property id -Descending   | ft id, Message -AutoSize


<#  Inspired by:
       https://dille.name/blog/2014/11/07/displaying-rds-event-log-messages-with-powershell/
       https://4sysops.com/archives/search-the-event-log-with-the-get-winevent-powershell-cmdlet/
#>

