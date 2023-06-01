#requires -Version 5.0
#requires -Modules Cloudpaging

<#
.SYNOPSIS
 Checks the real-time status of the Cloudpaging Player.
.DESCRIPTION
 Returns the real-time status of the Cloudpaging Player.
  If the Cloudpaging PowerShell module it will notify that the command does not exist.
.EXAMPLE
 .\CheckCloudpagingPlayerStatus.ps1
.NOTES
 To return the status of the Cloudpaging Player, the Player must be installed on the machine and the PowerShell module and cmdlets must exist.
#>

Import-Module Cloudpaging

$ErrorActionPreference = 'stop'
try {
    Get-CloudpagingClient | Select-Object Status
}
Catch {
    Write-Error "There was an error retrieving the Cloudpaging Player Status"
}

