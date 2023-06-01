#requires -Version 5.0
#requires -Modules Cloudpaging

<#
.SYNOPSIS
 Clears the Cloudpaging cache.
.DESCRIPTION
 This script will purge the current cache in use for cloudpaging application containers by the selected user.
.EXAMPLE
 .\ClearCloudpagingCache.ps1
.NOTES
 To clear the cache, the Player must be installed on the machine and the PowerShell module and cmdlets must exist.
#>

Import-Module Cloudpaging

$ErrorActionPreference = 'stop'
try {
    Clear-CloudpagingCache
}
Catch {
    Write-Error "There was an error clearing the Cloudpaging cache"
}

