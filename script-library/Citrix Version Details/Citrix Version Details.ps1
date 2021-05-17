# This section prevents the table from getting truncated on the right side
$pshost = get-host
$pswindow = $pshost.ui.rawui
$newsize = $pswindow.buffersize
$newsize.height = 300
$newsize.width = 200
$pswindow.buffersize = $newsize
$newsize = $pswindow.windowsize
$newsize.height = 50
$newsize.width = 125
$pswindow.windowsize = $newsize

$regkey = "hklm:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
$AllSubKey= Get-ChildItem $regkey | ForEach {Get-ItemProperty $_.PSPath} | where {$_.publisher -match "citrix"}
If ($AllSubKey  -eq $null) {
    Write-Host "Citrix components are not installed on this computer."
    Exit 0
}
$AllDetails = @()
ForEach ($SubKey in $AllSubKey) {
        $Details = New-Object PSObject
        $Details| add-member -MemberType NoteProperty -Name "Display Name" -Value $SubKey.DisplayName
        $Details| add-member -MemberType NoteProperty -Name "Version" -Value $SubKey.DisplayVersion
        $Details| Add-Member -MemberType NoteProperty -Name Identifier -Value $SubKey.PSChildName
        $AllDetails += $Details
}

$AllDetails | Sort-Object "Display Name" | ft -auto

