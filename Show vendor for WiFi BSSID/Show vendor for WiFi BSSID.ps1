<#
    .SYNOPSIS
    Show details of the passed WiFi BSSID

    .DESCRIPTION
    Use with ControlUp by creating a session context script script with the client metric (CU 8.5+ required) "WiFi BSSID"

    .PARAMETER bssid
    The BSSID of the WiFi network to look up

    .NOTES
    This script uses the API located at: https://macvendors.co/api

    Author: Guy Leech
    Version: 1.1
    Ton de Vreede: Added some error handling.
#>

[CmdletBinding()]
Param
(
    [Parameter(Mandatory = $true, HelpMessage = 'WiFi BSSid')]
    $bssid
)

$VerbosePreference = $(if ( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if ( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if ( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[string]$uri = "https://macvendors.co/api/$bssid/xml"

try {
    [XML]$wifiVendor = Invoke-WebRequest -URI $uri -UseBasicParsing -TimeoutSec 60 | Select-Object -ExpandProperty Content
}
catch {
    Throw "There was an error invoking the webreques to $uri. Exception:`n $_"
}

if ( $wifiVendor -and $wifiVendor.result ) {
    if ( $wifiVendor.result.PSObject.Properties[ 'error' ] ) {
        Write-Error -Message "Error `"$($wifiVendor.result.error)`" returned from $uri"
    }
    else {
        $wifiVendor.result | Select-Object -Property company, mac_prefix, address
    }
}
else {
    Write-Error -Message "No result returned from $uri"
}
