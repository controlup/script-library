#requires -Version 3.0
<#
.SYNOPSIS
Gets the public IP of the machine

.DESCRIPTION
The script uses a webrequest to see the which public IP the machine is using.

.AUTHOR
Ton de Vreede, based on community script by user rsheth
#>

$ErrorActionPreference = 'Stop'
try {
    # Use basic parsing in case Internet Explorer has never been run and there is no first configuration of IE.
    (Invoke-WebRequest -Uri "http://ifconfig.me/ip" -UseBasicParsing).Content
    Exit 0
}
catch {
    # Webrequest failed. There can be many reasons for this. So as not to polute output the address 0.0.0.0 as this is clearly in error but still in IP address format.
    Write-Output '0.0.0.0'
    Exit 1
}
