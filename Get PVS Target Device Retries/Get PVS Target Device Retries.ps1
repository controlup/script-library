$devicename = $args[0].Split(".")[0]

$ErrorActionPreference = "Stop"

If ( (Get-PSSnapin -Name McliPSSnapIn -ErrorAction SilentlyContinue) -eq $null )
{
    Try {
        Import-Module "$env:programfiles\Citrix\Provisioning Services Console\McliPSSnapIn.dll"  -WarningAction SilentlyContinue
    } Catch {
        Write-Host "There is a problem loading the Powershell SnapIn. It is not possible to continue."
        Exit 1
    }
}

Try {
    Mcli-help | Out-Null
}
Catch {
    Write-Host "This is not a PVS server. Exiting."
    Exit 1
}

Try {
    $retrycount = (mcli-get devicestatus -p devicename=$devicename -f status | select-string "status:").line.split(":")[1].trim()
}
Catch {
    Write-Host "This is not a PVS target device."
    Exit 1
}

Write-Host "Target Device Retries = $retrycount"

