## win32_useraccount WMI class found to be very slow
## wmic useraccount where name='%username%' get sid

(New-Object System.Security.Principal.NTAccount("$env:userdomain\$env:username")).Translate([System.Security.Principal.SecurityIdentifier]).value

