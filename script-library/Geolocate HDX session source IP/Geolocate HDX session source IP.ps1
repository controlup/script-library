#requires -Version 3.0
<#  
.SYNOPSIS     HDX Session geolocate - IP-API sample
.DESCRIPTION  ControlUp can display the Source IP of the HDX Sessions connected to the NetScaler. 
              Sample script action to use ip-api to obtain the geolocation.

              * For non-Commercial use only! For more info see https://members.ip-api.com/ * 
.CONTEXT      HDX Sessions

.TAGS         $HDX
.HISTORY      Marcel Calef     - 2020-11-30 - initial Sample
              Ton de Vreede    - 2020-12-02 - cleanup, changed output to object instead of string. Added optional time zone field. Added error handling for private IP address and lookup failure
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory = $true, HelpMessage = 'Source IP')]
    [string]$sourceIP,
    [Parameter(Mandatory = $true, HelpMessage = 'Show or hide IP')]
    [string]$showIP,
    [Parameter(Mandatory = $true, HelpMessage = 'Country, name or code')]
    [string]$country,
    [Parameter(Mandatory = $true, HelpMessage = 'Region, name or code')]
    [string]$region,
    [Parameter(Mandatory = $true, HelpMessage = 'Show or hide city')]
    [string]$city,
    [Parameter(Mandatory = $true, HelpMessage = 'Show or hide ISP')]
    [string]$isp,
    [Parameter(Mandatory = $true, HelpMessage = 'Show or hide coordinates')]
    [string]$coords,
    [Parameter(Mandatory = $true, HelpMessage = 'Show or hide AS Name')]
    [string]$asname,
    [Parameter(Mandatory = $true, HelpMessage = 'Show or hide Time Zone')]
    [string]$timezone
)

[string]$ErrorActionPreference = 'Stop'
# Remove the comment in the next line to enable Verbose outut
#$VerbosePreference = 'Continue'

Write-Verbose "Variables:"
Write-Verbose "    sourceIP :  $sourceIP"
Write-Verbose "      showIP :  $showIP"
Write-Verbose "     country :  $country"
Write-Verbose "      region :  $region"
Write-Verbose "        city :  $city"
Write-Verbose "         isp :  $isp"
Write-Verbose "      coords :  $coords"
Write-Verbose "      asname :  $asname"
Write-Verbose "    timezone :  $timezone"

# Test the IP number to see even if it is in a private IP range (these ranges are meant for internal use only, so are not publicly registerred)
If ($sourceIP.StartsWith('10.')) {
    Write-Output -InputObject 'The HDX client IP address is in the 10.x.x.x range which is a private IP range and cannot be resolved by ip-api.com.'
    Exit 1
}
elseif ($sourceIP.StartsWith('172.') -and ([int]$sourceIP.Split('.')[1] -ge 16) -and ([int]$sourceIP.Split('.')[1] -le 31)) {
    Write-Output -InputObject 'The HDX client IP address is in the 172.16-31.x.x range which is a private IP range and cannot be resolved by ip-api.com.'
    Exit 1
}
elseif ($sourceIP.StartsWith('192.168.')) {
    Write-Output -InputObject 'The HDX client IP address is in the 192.168.x.x range which is a private IP range and cannot be resolved by ip-api.com.'
    Exit 1
}

# Set URL
[string]$url = "http://ip-api.com/json/$sourceIP"

# Get the Geolocation data
try {
    $geolocationSample = Invoke-RestMethod -Method Get -Uri $url
} 
catch {
    $_.Exception.Response.Headers.ToString();
    Write-Output -InputObject "Query failed. Have you exceeded the 45 per minute limit on the non-commercial service?`nPlease look into the commercial plans offered by https://ip-api.com"
    Exit 1
}

# Test if lookup was successful
If ($geolocationSample.status -ne 'success') {
    Write-Output -InputObject "The IP lookup failed. The status returned by ip-api.com is: $($geolocationSample.status)"
    Exit 1
}

# Declare array for desired properties and add them
[array]$arrProperties = @()
if ($showIP -eq 'true') { $arrProperties += 'Query' }
if ($country -ne 'true') { $arrProperties += 'CountryCode' } else { $arrProperties += 'country' }
if ($region -ne 'true') { $arrProperties += 'Region' } else { $arrProperties += 'RegionName' }
if ($city -eq 'true') { $arrProperties += 'City' }
if ($isp -eq 'true') { $arrProperties += 'ISP' }
if ($coords -eq 'true') { $arrProperties += 'Lat', 'Lon' }
if ($asname -eq 'true') { $arrProperties += 'AS' }
if ($timezone -eq 'true') { $arrProperties += 'TimeZone' }

# Display results
$geolocationsample | Select-Object -Property $arrProperties | Format-Table -AutoSize

