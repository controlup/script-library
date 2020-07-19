$ErrorActionPreference = "Stop"

$regkeyICA = "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\ICA-TCP"
Try { $LAValue= (Get-ItemProperty $regkeyICA).LanAdapter }
Catch { Write-Host "The ICA protocol is not installed on this computer or bound to any adapter."; Exit }

if ($LAValue -eq 0) {
    Write-Host "The ICA Listener is bound to: All network adapters configured with this protocol"
} Else {
    $regkeyLinkage = "HKLM:\SYSTEM\CurrentControlSet\Services\TCPIP\Linkage"
    $BindList = (Get-ItemProperty $regkeyLinkage).Bind
    $NicList = @()

    #replace GUID with user friendly name
    foreach ($nic in $BindList) {  
        $nic = $nic -replace "\\device\\", "" 
        $regkeyCxn = "HKLM:\SYSTEM\CurrentControlSet\Control\Network\{4D36E972-E325-11CE-BFC1-08002be10318}\" + $nic + "\Connection"
        $NicName = (Get-ItemProperty $regkeyCxn).Name
        $NicList += $NicName 
    }
    #Array starts counting at '0' so subtract 1 from the LanAdapter value because it starts counting at '1'
    Write-Host "The ICA Listener is bound to: $NicList[($LAValue) - 1]"
}

