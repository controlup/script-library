#requires -Version 3.0
<#
.SYNOPSIS
Shows files that clients have open on SMB shares.

.DESCRIPTION
This script runs Get-SmbOpenFile, which retrieves basic information about the files that are open on behalf of the clients of the Server Message Block (SMB) server.

.NOTES
The required SMBShare module is only available on Windows 8/2012 and later, so this script will not run on older Windows versions.
Based on an idea by Dennis Geerlings

.AUTHOR
Ton de Vreede
#>

[CmdletBinding()]

# Set error handling
$ErrorActionPreference = 'Stop'

# Test Windows version
If ([Environment]::OSVersion.Version -le [version]'6.1') {
    Write-Output -InputObject 'This version of Windows is too low to run this script. The required module SmbShare is only available on Windows 2012/8 and later.'
    Exit 1
}

# Import required module
try {
    Import-Module -Name SmbShare
}
catch {
    Write-Output -InputObject "There was a problem importing the SmbShare module: $_"
    Exit 1
}

# Get the open file{s)
Try {
    $OpenFiles = Get-SmbOpenFile
    If ($null -ne $OpenFiles) {
        $OpenFiles | Format-Table -AutoSize
    }
    else {
        Write-Output -InputObject 'There are no files open on SMB shares.'
    }
}
Catch {
    Write-Output -InputObject "There was a problem executing the Get-SmbOpenFIle cmdlet: $_"
    Exit 1
}
