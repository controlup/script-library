#requires -Version 4.0

<#
.SYNOPSIS
    Get Office 365 Details
.DESCRIPTION
    Gets the name, version and install path of all found Office 365 installations.
.AUTHOR
	Ton de Vreede
.NOTES
If multiple installations are detected with different cultures but the same install path it means there could be several languages installed though they share the same code base.
#>

# Set up some defaults
$ErrorActionPreference = 'Stop'
# Configure a larger output width for the ControlUp PowerShell console
[int]$outputWidth = 400
# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

# Get the basic installation information
$OfficeInstalls = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | Where-Object { $_.Name.Split('\')[-1].StartsWith('O365') }

# Create expressions for output
[hashtable]$hshNameExp = @{L = 'Name'; E = { $_.DisplayName } }
[hashtable]$hshVersionExp = @{l = 'Version'; E = { $_.DisplayVersion } }
[hashtable]$hshCultureExp = @{l = 'Culture'; E = { $_.DisplayName.Split(' ')[-1].Trim() } }

# Output the result
If ($null -ne $OfficeInstalls) {
	Write-Output -InputObject 'The following Office 365 installations were found:'
	Foreach ($Office in $OfficeInstalls) {
		$office | Get-ItemProperty | Select-Object $hshNameExp, $hshVersionExp, $hshCultureExp, InstallLocation
	}
}
Else {
	Write-Output -InputObject 'No Office 365 installations were found.'
}

Exit 0
