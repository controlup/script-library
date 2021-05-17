$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Gets the attached USB devices

    .DESCRIPTION
    Gets the attached USB devices using WMI with the Win32_USBControllerDevice class, then displays a summary of the information.

    .EXAMPLE
    This script can be used to find USB devices that may be malfunctioning or misconfigured (Status field not OK). It can also be used to compare the attached USB devices of several machines.
    
    .NOTES
    Based on example code from James Brundage
    
    Context: This script can be triggered from: User session
    Modification history: 29=06-2019 - Ton de Vreede - Added simple error handling, selection and sorting to original community script submission
    
    .LINK
    https://devblogs.microsoft.com/powershell/get-usb-using-wmi-association-classes-in-powershell/

#>

# Retreive, select and sort USB Controller details
try {
    Get-WmiObject Win32_USBControllerDevice | Foreach-Object { [Wmi]$_.Dependent } | Select-Object Name,Manufacturer,PnPClass,DeviceID,Status | Sort-Object Name | Format-Table -AutoSize
}
catch {
    Write-Host "The USB devices could not be retreived. Exception detail:`n$_"
}
