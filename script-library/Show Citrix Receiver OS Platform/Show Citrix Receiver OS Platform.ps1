<#            Citrix Receiver Client OS Platform
Show the client device Operating System type form for a specific user session.
Use this script without requesting the display of the Receiver version to get the results 
grouped by client OS (available only as a ControlUp Script Based Action). 

Categorization of the Client OS is accomplished by querying the Citrix VDA or XA65 worker 
for the ClientPlatformId registry value in the appropriate Citrix ICA hive for that user session 
and follow the conversion described in this document:
https://www.citrix.com/mobilitysdk/docs/clientdetection.html


     Author:   Marcel Calef
     Date:      2018-12-16
     Version:  2.9

Parameters:
     Include Version - true or false 
     Session ID
     Receiver Version
Sources:
     https://www.citrix.com/mobilitysdk/docs/clientdetection.html
#>

$ErrorActionPreference = "Continue"        # Ignore PoSh errors, a proper message will be displayed later

$outputDesired=$args[0]
$sessionID=$args[1]
$rxVer=$args[2]

if ([string]::IsNullOrEmpty($rxVer))  {Write-Host "No Citrix session connected"; exit 1}

#$sessionID=6
#(Get-ItemProperty HKLM:\software\Citrix\ICA\Session\$sessionID\Connection -name ClientProductId).ClientproductId

$rxOS="NotCitrix"
# Get the ClientProductID from the Citrix Session
$rxOS=(Get-ItemProperty HKLM:\software\Citrix\ICA\Session\$sessionID\Connection -name ClientProductId).ClientproductId

if ($rxOS -eq 1) {$platform="Windows"}
if ($rxOS -eq 81) {$platform="Linux"}
if ($rxOS -eq 82) {$platform="Mac"}
if ($rxOS -eq 83) {$platform="iOS"}
if ($rxOS -eq 84) {$platform="Android"}
if ($rxOS -eq 85) {$platform="Blackberry"}
if ($rxOS -eq 86) {$platform="Windows Phone 8/WinRT"}
if ($rxOS -eq 87) {$platform="Windows Mobile"}
if ($rxOS -eq 88) {$platform="Blackberry Playbook"}
if ($rxOS -eq 257) {$platform="HTML5"}
if ($rxOS -match "NotCitrix") {$platform="N/A"}  # report N/A for not Citrix

if ($outputDesired -match "true") {Write-Host "version: $rxVer      " -noNewLine}
Write-Host " Client OS: $platform         OSid:$rxOS "
