$ErrorActionPreference = 'Stop'
#requires -version 5.1

<#
    .SYNOPSIS
    Lists the local adminsitrators

    .DESCRIPTION
    Gets the local administrators by querying the Local Adminsitrator group on the target machine

    .EXAMPLE
    - To check is somebody has administrator rights on a machine
    - To check there are not too many local administrators
    
    .NOTES
    Based on a community script submission from Ben Bonnette

    Requires PS 5.1 to be available on the target machine.
        
    Context: This script can be triggered from: Computer
    Modification history: 29062019 - Ton de Vreede - Added simple error handling and changed the output replace the name of the local machine to 'LOCAL', so this script can also be used to compare local admins on machines (otherwise when running on multiple targets each machine gets its won output pane).

#>

# Get the local machine name
[string]$strComputerName = "$([System.Environment]::ExpandEnvironmentVariables('%COMPUTERNAME%'))\"

# Try to get the local administrator group members. REPLACING THE LOCAL COMPUTER NAME WITH 'LOCAL'!
try {
    Get-LocalGroupMember -Group Administrators | Select-Object ObjectClass, @{Label = 'Name'; Expression = { ($_.Name).replace($strComputerName, 'LOCAL\') } }, PrincipalSource
}
catch {
    Write-Host "The local Administrators group members could not be retreived. Exception detail:`n$_"
}

