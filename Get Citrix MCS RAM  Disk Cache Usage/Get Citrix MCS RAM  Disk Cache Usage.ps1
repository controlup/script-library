#Requires -Version 3.0

$availablemem = [math]::Round((Get-counter '\Citrix MCS Storage Driver\Cache memory target size').countersamples[0].cookedvalue / 1mb)
$memused=[math]::Round((Get-counter '\Citrix MCS Storage Driver\Cache memory used').countersamples[0].cookedvalue / 1mb)

$usedmemperc=[math]::Round(($memused / $availablemem) * 100)

$availabledisk=[math]::Round((Get-counter '\Citrix MCS Storage Driver\Cache disk size').countersamples[0].cookedvalue / 1gb,2)
$diskused=[math]::Round((Get-counter '\Citrix MCS Storage Driver\Cache disk used').countersamples[0].cookedvalue / 1gb,2)

$useddiskperc=[math]::Round(($diskused / $availabledisk) * 100)

$Object = New-Object PSObject -Property @{
    "Total Ram Cache (mb)" = $availablemem
    "Used Ram Cache (mb)" = $memused
    "Used Ram Cache (%)" = $usedmemperc
    "Total Disk Cache (gb)" = $availabledisk
    "Used Disk Cache (gb)" = $diskused
    "Used Disk Cache (%)" = $useddiskperc
}

$object | select  "Used Ram Cache (%)","Used Disk Cache (%)","Total Ram Cache (mb)", "Used Ram Cache (mb)","Total Disk Cache (gb)","Used Disk Cache (gb)"
