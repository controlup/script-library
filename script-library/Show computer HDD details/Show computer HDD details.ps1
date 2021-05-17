#requires -version 3
$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Report on the target machine NIC(s) hard disk drives

    .DESCRIPTION
    Retreives the HDD(s) of the target machine and outputs the details, verbose if required.

    .PARAMETER bolVerbose
    Switch for verbose reporting.

    .NOTES
    The codes for Availability, ConfigManagerErrorCode and StatusInfo are translated to their meaning
    in the Win32_DiskDrive class.
#>

# Is Verbose output required? Default is false
If ($args[0] -eq 'True') {[bool]$bolVerbose = $true} Else {[bool]$bolVerbose = $false}

# Array for what needs to be displayed
[array]$arrWin32DisplaySet = @(
    'Caption',
    'MediaType',
    'Availability',
    'InterfaceType',
    'Size',
    'FirmwareRevision',
    'DriverProviderName',
    'DriverVersion'
)

# If verbose output is required, add the verbose properties
If ($bolVerbose){
    $arrWin32DisplaySet += @(
        'Model',
        'DeviceID',
        'StatusInfo',
        'Partitions',
        'BytesPerSector',
        'DefaultBlockSize',
        'Index',
        'SectorsPerTrack',
        'TotalCylinders',
        'TotalHeads',
        'TotalSectors',
        'TotalTracks',
        'TracksPerCylinder',
        'CapabilityDescriptions',
        'SCSIBus',
        'SCSILogicalUnit',
        'SCSIPort',
        'SCSITargetId'
    )
}

Function Feedback ($strFeedbackString)
{
  # This function provides feedback in the console on errors or progress, and aborts if error has occured.
  If ($error.count -eq 0)
  {
    # Write content of feedback string
    Write-Host -Object $strFeedbackString -ForegroundColor 'Green'
  }
  
  # If an error occured report it, and exit the script with ErrorLevel 1
  Else
  {
    # Write content of feedback string but in red
    Write-Host -Object $strFeedbackString -ForegroundColor 'Red'
    
    # Display error details
    Write-Host 'Details: ' $error[0].Exception.Message -ForegroundColor 'Red'
    Exit 1
  }
}

# Try to import the CimCmdlets module
Try {
    Import-Module CimCmdlets
  }
  Catch {
    Feedback "There was an error loading the CimCmdlets module."
  }

# Create hashtables with lookup values
$hshAvailability = @{
    '1' = 'Other'
    '2' = 'Unknown'
    '3' = 'Running/Full Power'
    '4' = 'Warning'
    '5' = 'In Test'
    '6' = 'Not Applicable'
    '7' = 'Power Off'
    '8' = 'Off Line'
    '9' = 'Off Duty'
    '10' = 'Degraded'
    '11' = 'Not Installed'
    '12' = 'Install Error'
    '13' = 'Power Save - Unknown'
    '14' = 'Power Save - Low Power Mode'
    '15' = 'Power Save - Standby'
    '16' = 'Power Cycle'
    '17' = 'Power Save - Warning'
    '18' = 'Paused'
    '19' = 'Not Ready'
    '20' = 'Not Configured'
    '21' = 'Quiesced'
    }

$hshConfigManagerErrorCode = @{
    '0' = 'This device is working properly.'
    '1' = 'This device is not configured correctly.'
    '2' = 'Windows cannot load the driver for this device.'
    '3' = 'The driver for this device might be corrupted, or your system may be running low on memory or other resources.'
    '4' = 'This device is not working properly. One of its drivers or your registry might be corrupted.'
    '5' = 'The driver for this device needs a resource that Windows cannot manage.'
    '6' = 'The boot configuration for this device conflicts with other devices.'
    '7' = 'Cannot filter.'
    '8' = 'The driver loader for the device is missing.'
    '9' = 'This device is not working properly because the controlling firmware is reporting the resources for the device incorrectly.'
    '10' = 'This device cannot start.'
    '11' = 'This device failed.'
    '12' = 'This device cannot find enough free resources that it can use.'
    '13' = "Windows cannot verify this device`'s resources."
    '14' = 'This device cannot work properly until you restart your computer.'
    '15' = 'This device is not working properly because there is probably a re-enumeration problem.'
    '16' = 'Windows cannot identify all the resources this device uses.'
    '17' = 'This device is asking for an unknown resource type.'
    '18' = 'Reinstall the drivers for this device.'
    '19' = 'Failure using the VxD loader.'
    '20' = 'Your registry might be corrupted.'
    '21' = 'System failure: Try changing the driver for this device. If that does not work, see your hardware documentation. Windows is removing this device.'
    '22' = 'This device is disabled.'
    '23' = "System failure: Try changing the driver for this device. If that doesn`'t work, see your hardware documentation."
    '24' = 'This device is not present, is not working properly, or does not have all its drivers installed.'
    '25' = 'Windows is still setting up this device.'
    '26' = 'Windows is still setting up this device.'
    '27' = 'This device does not have valid log configuration.'
    '28' = 'The drivers for this device are not installed.'
    '29' = 'This device is disabled because the firmware of the device did not give it the required resources.'
    '30' = 'This device is using an Interrupt Request (IRQ resource that another device is using.'
    '31' = 'This device is not working properly because Windows cannot load the drivers required for this device.'
}

$hshStatusInfo = @{
    '1' = 'Other'
    '2' = 'Unknown'
    '3' = 'Enabled'
    '4' = 'Disabled'
    '5' = 'Not Applicable'
}

# Declare the array of properties to be returned
[array]$arrWin32PropertySet = @(
@{Name='Availability';Expression={$hshAvailability[($_.Availability.ToString())]}},
'BytesPerSector',
'CapabilityDescriptions',
'Caption',
@{Name='ConfigManagerErrorCode';Expression={$hshConfigManagerErrorCode[($_.ConfigManagerErrorCode.ToString())]}},
'DefaultBlockSize',
'DeviceID',
'FirmwareRevision',
'Index',
'InterfaceType',
'LastErrorCode',
'MediaType',
'Model',
'NeedsCleaning',
'Partitions',
'PNPDeviceID',
'SCSIBus',
'SCSILogicalUnit',
'SCSIPort',
'SCSITargetId'
'SectorsPerTrack',
'Size',
@{Name='StatusInfo';Expression={$hshStatusInfo[($_.StatusInfo.ToString())]}},
'TotalCylinders',
'TotalHeads',
'TotalSectors',
'TotalTracks',
'TracksPerCylinder'
)

# Get the HDD details
try{
    # Get the NIC information
    $objHDDs = Get-CIMInstance -ClassName Win32_DiskDrive  | Select-Object -Property $arrWin32PropertySet
    # Add the driver details
    Foreach ($HDD in $objHDDs){
        $DriverDetails = Get-CIMInstance -ClassName Win32_PnPSignedDriver| Where-Object {$_.DeviceID -eq $HDD.PNPDeviceID}
        Add-Member -InputObject $HDD -NotePropertyName DriverProvidername -NotePropertyValue $DriverDetails.DriverProviderName
        Add-Member -InputObject $HDD -NotePropertyName DriverVersion -NotePropertyValue $DriverDetails.DriverVersion
    }
}
catch {
    Feedback "There was an error retrieving the hard drive(s)details."
}

# Display the resulting details
$objHDDs | Select-Object -Property $arrWin32DisplaySet | Sort-Object -Property ProductName | Format-List

