#File locations:
$installutilroot = "$env:systemroot\Microsoft.NET\Framework64\v4.0.30319\installutil.exe"
$pvsconsoleinstallroot = "C:\Program Files\Citrix\Provisioning Services Console"

#Script:
If (!(Test-Path "$pvsconsoleinstallroot\Citrix.PVS.SnapIn.dll")) {
    Write-Warning "This server is not running provisioning services 7.6 or higher, or the console is not installed."
}
Else {
    $exit=$false
    Try {
        Get-PSSnapin -Registered Citrix.PVS.SnapIn -ErrorAction stop | Out-Null
    }
    Catch {
        Write-Warning "PowerShell Snapin for Provisioning Services is not registered. Attempting to register it now..."
        Try {
            cd "$pvsconsoleinstallroot\"
            Start-Process -FilePath $installutilroot -ArgumentList "Citrix.PVS.SnapIn.dll" -WorkingDirectory "$pvsconsoleinstallroot\" -NoNewWindow -wait
            }
        Catch {
            Write-Warning "An error occurred while registering the snapin for Provisioning services."
        }
    }

    Try {
        Add-PSSnapin citrix.pvs.snapin -ErrorAction Stop
    }
    Catch {
        Write-Warning "Powershell snapin not registered for provisioning services, the script will now close."
        $exit=$true
    }

    If (!$exit) { 
       
       Write-Host "Site Report:"
       Get-PvsSite | Select name, @{Name = 'Servers'; Expression = {(Get-PvsServer -sitename $_.sitename).count}},@{Name = 'Online Devices'; Expression = {(Get-PvsDeviceInfo -sitename $_.sitename -OnlyActive).count}} -wait | ft       
       
       Write-Host "Collection Report:"
       Get-PvsCollection | Select name, sitename, Enabled,  @{N='Devices'; E={$_.devicecount}},  @{N='Active Devices'; E={$_.activedevicecount}} | ft

       Write-Host "Disk Report:"
       $disks = Get-PvsDiskInfo
       $disks | Select name, enabled, active,  @{N='Site'; E={$_.sitename}},  @{N='Store'; E={$_.storename}},  @{N='Server'; E={$_.servername}}, locked, @{N='Devices'; E={$_.devicecount}} | ft
       ForEach ($disk in $disks) {
            If ($disk.ServerName.length -gt 0) {
                Write-Warning "Disk image $($disk.name) is currently bound to a single server: $($disk.servername)"
            }
        }

       Write-Host "Server Report:"
       $servers = Get-PvsServerInfo
       $servers | Select name, sitename, @{Name = 'Active'; Expression ={[System.Convert]::ToBoolean($_.active)}}, devicecount | ft
       ForEach ($server in $servers | Where {$_.active -eq 0}) {
            Write-Warning ("$($server.name) is currently offline!")
       }

        $stores = @()
        Write-Host "Store Report:"
        ForEach ($server in $servers | Where {$_.active -eq 1}) {
            ForEach ($store in Get-PvsStore -ServerName $server.name) {
                $properties = @{
                'ServerName'=$server.name;
                'Store'=$store.name;
                'Free Space (MB)'=(Get-PvsStoreFreeSpace -StoreName $store.name -ServerName $server.name)}
                $object = New-Object -TypeName PSObject -Property $properties
                $stores+=$object
            }
        }
       $stores | Select store, servername, 'Free Space (MB)' | ft

        Write-Host "Device Retry Report:"
        $devices=Get-PvsDeviceInfo | Where {$_.active -and $_.status -gt 0}
        If ($devices.count -gt 0) {
            $devices | Select devicename,Collectionname,sitename,@{Name = 'Server'; Expression = {(Get-PvsDeviceInfo -devicename $_.devicename).servername}},@{Name = 'Retries'; Expression = {$_.status}} | sort -Property Retries -Descending | ft
        }
        else {
            Write-Host -ForegroundColor Green "No Devices with retries! - Devices online: $($devicesonline.count)"
        }
    }
}

