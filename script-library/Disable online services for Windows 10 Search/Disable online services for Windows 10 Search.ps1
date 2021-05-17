$searchkey = $null
try {
$searchkey = Get-Item -path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search' -ErrorAction Stop
} catch {
Write-Output "The Search registry key does not exist, so this user session is probably not running on Windows 10"
}

$valueexists = $false
if ($searchkey -ne $null) {
    try {
    $bingvalue = Get-Itemproperty -path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search' -Name 'BingSearchEnabled' -ErrorAction Stop
    if ($bingvalue.BingSearchEnabled -eq 0) { 
        Write-Output "No change performed. BingSearchEnabled = 0"
     } else {
        Set-Itemproperty -path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search' -Name 'BingSearchEnabled' -value '0' -Force -ErrorAction Stop
        if ($?) { Write-Output "BingSearchEnabled value set to 0" } else { Write-Output "Error setting registry value" }
     }
    } catch {    
    New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled" -Value 0 -PropertyType DWORD
    if ($?) { Write-Output "BingSearchEnabled value created, value set to 0" } else { Write-Output "Error creating registry value" }
    }
}

