# requires -Version 5.0
# requires -Modules Cloudpaging

<#
.SYNOPSIS
 Lists all Cloudpaging application containers available to a user.
.DESCRIPTION
 Returns all Cloudpaging application containers available to a user.
  If the Cloudpaging PowerShell module it will notify that the command does not exist.
.EXAMPLE
 .\CheckCloudpagingApps.ps1
.NOTES
 To return the applications, the Player must be installed on the machine and the PowerShell module and cmdlets must exist.
#>

$ErrorActionPreference = 'stop'
try {Get-CloudpagingApp | Select-Object Name}
Catch{write-output -InputObject "There was an error retrieving the Cloudpaging application containers"; Exit 1}


