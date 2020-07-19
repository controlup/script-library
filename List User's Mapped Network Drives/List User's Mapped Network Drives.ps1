$drives = Get-WmiObject -Class Win32_MappedLogicalDisk | select @{Name="Drive";Expression={$_.Name}}, @{Name="UNC Share";Expression={$_.ProviderName}}

if ($drives -ne $null) {Write-Output $drives | ft -AutoSize}
if ($drives -eq $null) {Write-Output "No mapped drives present in this user's session."}
