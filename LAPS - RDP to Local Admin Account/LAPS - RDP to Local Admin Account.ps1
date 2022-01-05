<#
    .SYNOPSIS
        Connects via RDP to the Local Administrator account on the targeted machines

    .DESCRIPTION
        Retrieves the password for a machine protected by the Local Administrator Password Solution, generates an RDP file then connects to the machine.

    .EXAMPLE
        . .\Connect-ToLocalAdminAccount.ps1 -ComputerName W2019-001
        Retrieves the password for machine W2019-001 protected by the Local Administrator Password Solution, generates an RDP file then connects to the machine.

    .NOTES
        Designed to run as the CONSOLE context on the target machine so the user running the script requires rights to get the password

    .CONTEXT
        CONSOLE

    .MODIFICATION_HISTORY
        Created TTYE : 2020-10-13


    AUTHOR: Trentent Tye
#>
[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='Enter the SamAccountName of the machine')][ValidateNotNullOrEmpty()]  [string]$ComputerName
)

function Encrypt-RdpPassword {
    param (
        [String]$Password
    )
    Try {
        Add-Type -AssemblyName System.Security
	
        # use unicode (UTF-16LE) instead of UTF-8 in order to work with .rdp files ("password 51:b:")
        $EncryptArray = [System.Security.Cryptography.ProtectedData]::Protect($([System.Text.Encoding]::Unicode.GetBytes($Password)), $Null, "LocalMachine")
	
        Return (@($EncryptArray | ForEach-Object -Process { "{0:X2}" -f $_ }) -join "")
    } Catch {
        Write-Error "Failed to encrypt the password"
    }
}


#Use native ADSI queries to avoid using ActiveDirectory powershell modules (which might not be installed on the target machines)
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry
$objSearcher.Filter = "(&(objectCategory=Computer)(SamAccountname=$($COMPUTERNAME)`$))"
$objSearcher.SearchScope = "Subtree"
$ComputerObj = $objSearcher.FindOne()
$password = $ComputerObj.Properties["ms-Mcs-AdmPwd"]

#find local administrator account
$account = Get-WmiObject -ComputerName $ComputerName -Class Win32_UserAccount -Filter "LocalAccount='True' And Sid like '%-500'"

$RdpPassword = Encrypt-RdpPassword -Password $password

Write-Output @"
screen mode id:i:1
use multimon:i:0
desktopwidth:i:1200
desktopheight:i:860
session bpp:i:32
winposstr:s:0,1,949,371,3375,1906
compression:i:1
keyboardhook:i:2
audiocapturemode:i:1
videoplaybackmode:i:1
connection type:i:7
networkautodetect:i:1
bandwidthautodetect:i:1
displayconnectionbar:i:1
enableworkspacereconnect:i:0
disable wallpaper:i:0
allow font smoothing:i:0
allow desktop composition:i:0
disable full window drag:i:1
disable menu anims:i:1
disable themes:i:0
disable cursor setting:i:0
bitmapcachepersistenable:i:1
full address:s:$($ComputerName)
username:s:$($account.caption)
::domain:s:$($ComputerName)
password 51:b:$($RdpPassword)
audiomode:i:0
redirectprinters:i:0
redirectcomports:i:0
redirectsmartcards:i:0
redirectclipboard:i:1
redirectposdevices:i:0
autoreconnection enabled:i:1
authentication level:i:0
prompt for credentials:i:0
negotiate security layer:i:1
remoteapplicationmode:i:0
alternate shell:s:
shell working directory:s:
gatewayhostname:s:
gatewayusagemethod:i:4
gatewaycredentialssource:i:4
gatewayprofileusagemethod:i:0
promptcredentialonce:i:0
gatewaybrokeringtype:i:0
use redirection server name:i:0
rdgiskdcproxy:i:0
kdcproxyname:s:
drivestoredirect:s:
administrative session:i:1
"@ | Out-File "$env:temp\LapsRDP.rdp" -Force

Start-Process -FilePath mstsc.exe -ArgumentList "$env:temp\LapsRDP.rdp"

