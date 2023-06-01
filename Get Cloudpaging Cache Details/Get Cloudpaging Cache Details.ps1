#requires -Version 5.0
#requires -Modules Cloudpaging

<#
.SYNOPSIS
 Returns details of the Cloudpaging cache such as cache storage location, percentage used and more.
.DESCRIPTION
  Returns details of the Cloudpaging cache including cache storage location, size, maximum size,
  minimum size and percentage used.
.EXAMPLE
 .\CheckCloudpagingPlayerCacheDetails.ps1
.NOTES
 To return Cloudpaging Player cache details, the Player must be installed on the machine and the PowerShell module and cmdlets must exist.
#>

Import-Module Cloudpaging

$ErrorActionPreference = 'stop'
try {
    Get-CloudpagingCache
}
Catch{
    Write-Error "There was an error retrieving the client cache details."
}

