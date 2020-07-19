#requires -Version 3.0

$ErrorActionPreference = "Stop"

$Office = Get-ItemProperty HKLM:\software\WoW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | where displayname -match 'microsoft office' | where {$_.displayversion} | Select-Object displayname, displayversion | Group-Object displayversion
If ($Office -eq $null) {
    $Office = Get-ItemProperty HKLM:\software\Microsoft\Windows\CurrentVersion\Uninstall\* | where displayname -match 'microsoft office' | where {$_.displayversion} | Select-Object displayname, displayversion | Group-Object displayversion
}
If ($Office -ne $null) {
    $version = $Office.name.split(".")[0] + ".0"
} Else {
    Write-Host "Office is not installed on this computer."
    Exit
}

Try {
    $excel = Get-ItemProperty "hkcu:\Software\Policies\Microsoft\office\$version\excel\options" -Name autorecoverpath | Select-Object -ExpandProperty autorecoverpath
    Write-Host "Excel GPO Autosave path = $excel"
} Catch { }
Try {
    $excel = Get-ItemProperty "hkcu:\Software\Microsoft\office\$version\excel\options" -Name autorecoverpath | Select-Object -ExpandProperty autorecoverpath
    Write-Host "Excel Autosave path = $excel"
} Catch { }
If ($excel -eq $null) {
    Write-Host "Excel AutoSave path is the default."
}

Try {
    $word = Get-ItemProperty "hkcu:\Software\Policies\Microsoft\office\$version\word\options" -Name autosave-path | Select-Object -ExpandProperty autosave-path
    Write-Host "Word GPO Autosave path = $word"
} Catch { }
Try {
    $word = Get-ItemProperty "hkcu:\Software\Microsoft\office\$version\word\options" -Name autosave-path | Select-Object -ExpandProperty autosave-path
    Write-Host "Word Autosave path = $word"
} Catch { }
If ($word -eq $null) {
    Write-Host "Word AutoSave path is the default."
}

