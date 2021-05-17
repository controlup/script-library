<#  
.SYNOPSIS
      Script to show an ICA session's ICA round trip time and network latency.
.DESCRIPTION
      The script runs for 20 seconds and measures (once every 2 seconds) the ICA RTT and network latency of the relevant session.
      The output shows the session info (username, device name/IP, session name/ID) and 10 reads (once every 2 seconds) of the session's ICA RTT and network latency in seconds.
.PARAMETER 
      This script has 2 parameters:
      SessionID - The ID of the session from the ControlUp Console (e.g. 7, 34, 21)
      SessionName - The name of the session from the ControlUp Console (e.g. ICA-TCP#1, ICA-CGP#2)
.EXAMPLE
        ./ICARTT.ps1 "ICA-TCP#2" "7"
.OUTPUTS
        Session info (username, device name/IP, session name/ID) and 10 reads (once every 2 seconds) of session's ICA RTT and network latency.
.LINK
        See https://www.controlup.com
#>

$session_name = $args[0]
$sessionID = $args[1]
if ($session_name -Match 'RDP') {
    Write-Host "There is no ICA RTT data on an RDP session. Please choose an ICA session."
    exit
}
if (-Not (Get-WmiObject -Namespace root\Citrix\euem Citrix_Euem_ClientConnect)) {
    Write-host "Couldn't find EUEM data. It's available on XenApp\XenDesktop 7.x and higher versions."
    exit}
Else {
    $final_obj = @()
    $obj = New-Object PSObject
    $session_info = Get-WmiObject -Namespace root\Citrix\euem Citrix_Euem_ClientConnect | where {$_.WinstationName -eq $session_name} | select username, ClientMachineIP, ClientMachineName, PSComputerName
    $obj | Add-Member Username $session_info.username
    $obj | Add-Member "Device Name" $session_info.ClientMachineName
    $obj | Add-Member "Device IP" $session_info.ClientMachineIP
    $obj | Add-Member "Connected To" $session_info.PSComputerName
    $obj | Add-Member "Session Name" $session_name
    $obj | Add-Member "Session ID" $sessionID
    For ($i=0; $i -le 9; $i++) {
        $temp_obj = New-Object PSObject
        $temp = Get-WmiObject -Namespace root\Citrix\euem citrix_euem_RoundTrip | where {$_.SessionID -eq $sessionID} | select NetworkLatency, RoundtripTime
        $time = Get-Date -Format T
        $temp_obj | Add-Member "Time" $time
        $temp_obj | Add-Member "ICA RTT" $temp.RoundtripTime
        $temp_obj | Add-Member "Network Latency" $temp.NetworkLatency
        $final_obj += $temp_obj
        sleep 2  
    }
    Write-Output $obj
    Write-Output $final_obj | Format-Table
}

