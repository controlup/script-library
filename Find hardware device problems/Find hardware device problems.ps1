#require -version 3.0
<#
.SYNOPSIS
    Find hardware device problems

.DESCRIPTION
    This script gets all the computer devices and checks the status.

.NOTES
    Version:        1.0
    Author:         Joel Stocker
    Creation Date:  2022-05-17
    Updated:        2022-05-23  Ton de Vreede	Error handling, refactored
#>
$ErrorActionPreference = 'Stop'

# Set output encoding to ensure non-ASCII characters are captured
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Build an array with error codes and description, based on https://support.microsoft.com/en-us/topic/error-codes-in-device-manager-in-windows-524e9e89-4dee-8883-0afa-6bca0456324e
[hashtable]$hshErrorMapping = @{
	1  = "This device is not configured correctly"
	3  = "The driver for this device might be corrupted"
	9  = "Windows cannot identify this hardware"
	10 = "This device cannot start"
	12 = "This device cannot find enough free resources that it can use"
	14 = "This device cannot work properly until you restart your computer"
	16 = "Windows cannot identify all the resources this device uses"
	18 = "Reinstall the drivers for this device"
	19 = "Windows cannot start this hardware device"
	21 = "Windows is removing this device"
	22 = "This device is disabled"
	24 = "This device is not present, is not working properly"
	28 = "The drivers for this device are not installed"
	29 = "This device is disabled"
	31 = "This device is not working properly"
	32 = "A driver (service) for this device has been disabled"
	33 = "Windows cannot determine which resources are required for this device"
	34 = "Windows cannot determine the settings for this device"
	35 = "Your computer's system firmware does not"
	36 = "This device is requesting a PCI interrupt"
	37 = "Windows cannot initialize the device driver for this hardware"
	38 = "Windows cannot load the device driver"
	39 = "Windows cannot load the device driver for this hardware"
	40 = "Windows cannot access this hardware"
	41 = "Windows successfully loaded the device driver"
	42 = "Windows cannot load the device driver"
	43 = "Windows has stopped this device because it has reported problems"
	44 = "An application or service has shut down this hardware device"
	45 = "Currently, this hardware device is not connected to the computer"
	46 = "Windows cannot gain access to this hardware device"
	47 = "Windows cannot use this hardware device"
	48 = "The software for this device has been blocked"
	49 = "Windows cannot start new hardware devices"
	50 = "Windows cannot apply all of the properties for this device"
	51 = "This device is currently waiting on another device"
	52 = "Windows cannot verify the digital signature for the drivers required for this device"
	53 = "This device has been reserved for use by the Windows kernel debugger"
	54 = "This device has failed and is undergoing a reset"
}

# Gather the data
try {
	$objDevices = Get-CimInstance -ClassName Win32_PNPEntity | Where-Object { $_.ConfigManagerErrorCode -ne 0 } `
	| Select-Object @{Expression={$_.Name};Label="Device Name"},`
	@{Expression={$_.DeviceID};Label="Device ID"},`
	@{Expression={$_.ConfigManagerErrorCode};Label="Error Code"},`
	@{Expression={$hshErrorMapping.[int]$_.ConfigManagerErrorCode};Label="Error Description"}
}
catch {
	Write-Output -InputObject "There was an issue getting the device(s) status:`n$_"
	Exit 1
}

# Only output the devices that have a problem
If ($objDevices.Count -eq 0) {
	Write-Output -InputObject "No hardware device issues found on machine."
}
Else {
	$objDevices | Format-Table
}
