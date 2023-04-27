#requires -Version 5.0
#requires -Modules Cloudpaging

<#
.SYNOPSIS
 Checks version of Cloudpaging Player installed on a given machine.
.DESCRIPTION
 Checks version of Cloudpaging Player installed on a given machine. It also contains some error handling,
  If the Cloudpaging PowerShell module it will notify that the command does not exist.
.EXAMPLE
 .\CheckCloudpagingPlayerVersion.ps1
.NOTES
 To return the version of the Cloudpaging Player, the Player must be installed on the machine and the PowerShell module and cmdlets must exist.
#>

Import-Module Cloudpaging

$ErrorActionPreference = 'stop'
try {
    Get-CloudpagingClient | Select-Object Version
}
Catch {
    Write-Error -InputObject "There was an error retrieving the Cloudpaging Player version"
}

