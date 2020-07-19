Write-Host "List all Printers"
$RedirectedFolders = Get-WmiObject -Class Win32_Printer | select -Property Name,Sharename
if ($RedirectedFolders -eq $null) {
    Write-Host "No Printers"
} else {
    $RedirectedFolders | Format-Table -Autosize
    Write-Host "----------------------------------------------------------------"
    Write-Host "Default Printer is:"
    (Get-WmiObject -Class Win32_Printer -Filter "Default = $true").Name
}

