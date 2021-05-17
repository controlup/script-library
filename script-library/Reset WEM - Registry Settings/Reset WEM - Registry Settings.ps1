<# 
.SYNOPSIS
    Remove ControlUp Registry settings in the cache in the logged on user's registry
.DESCRIPTION
    The registry cache occassionally becomes corrupted. This SBA will remove the WEM cache for only
    the registry settings. A WEM refresh on the agent / VDA is required to complete the refresh. You
    may also log the user off / back on again.
.PARAMETER RegPath
    Path to the registry cache
.NOTES
    Created by Tim Riegler
#>

$RegPath = 'HKCU:\Software\VirtuAll Solutions\VirtuAll User Environment Manager\Agent\Tasks Exec Cache\RegistryValues\*'

# Check if the item exists and delete all sub-keys.
Try {
    If (Test-Path (get-childitem $RegPath)) {
        Write-Host "Key Exists! Deleting exisitng key..."
        # Remove-Item -Path HKCU:\CurrentVersion\* -Recurse
        Remove-Item -Path $RegPath -Recurse
    } 
}
Catch {
    Throw "Key does NOT exist. Exiting."
    Exit 1
}


Write-Host 'Registry Key reset. Please use WEM Admin Console to refresh the workspace agent on this host:' $env:computername
