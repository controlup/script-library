$devicename = $args[0].Split(".")[0]

Import-Module "C:\Program Files\Citrix\Provisioning Services Console\McliPSSnapIn.dll"  -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

# If Import-Module was successful...
if ($?) {
    Mcli-Run Reboot -p DeviceName=$devicename
    Mcli-Run MarkDown -p DeviceName=$devicename
    Mcli-Run ResetDeviceForDomain -p DeviceName=$devicename
    # if ResetDevice was successful...
     if ($?) {
      Write-Host "AD account reset. Rebooting device..."
     }
}

