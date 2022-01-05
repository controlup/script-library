#Requires -version 3.0
<#
    .SYNOPSIS
    Gets available PowerShell modules

    .DESCRIPTION
    Gets the PowerShell modules available to the account running the script; with type, version and description. If specified extra details will be displayed.

    .EXAMPLE
    & '.\List installed Powershell modules.ps1' -Detailed false -ModuleName *zure*
    Outputs a table of all modules whose name contains 'zure'; with module type, version and description.

    & '.\List installed Powershell modules.ps1' -Detailed true
    Outputs a table of ALL modules; with module type, version and description. Then outputs a list of the modules with extra datails such as Company, Tags etc.

    .NOTES
    This script only shows the modules that are available to the account running the script. Modules that have been installed under a different account with the Scope CurrentUser are not displayed.

    Version:        1.0
    Author:         Ton de Vreede
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, HelpMessage = 'Display extra details.')]
    [ValidateSet('true', 'false')]
    [string]$Detailed,
    [Parameter(Mandatory = $false, HelpMessage = 'Name or name with wildcards of modules to get')]
    [string]$ModuleName
)

# Get the module list first
If ($PSBoundParameters.ContainsKey('ModuleName')) {
    $objModules = Get-Module -Name $ModuleName -ListAvailable
}
Else {
    $objModules = Get-Module -ListAvailable
}

# Test if any modules were found
If ($null -eq $objModules) {
    Write-Output -InputObject "No modules matching the name $ModuleName were found."
    Exit 0
}
Else {
    # Set the the size of the PS Buffer to 90% of the max window size
    $PSWindow = (Get-Host).UI.RawUI
    $WideDimensions = $PSWindow.BufferSize
    $WideDimensions.Width = [math]::Round($PSWindow.MaxPhysicalWindowSize.Width * .9)
    $PSWindow.BufferSize = $WideDimensions

    # Output basic info in a table
    Write-Output -InputObject $objModules | Select-Object Name, ModuleType, Version, Description | Sort-Object Name | Format-Table

    # Add list with details if required
    If ([System.Convert]::ToBoolean($Detailed)) {
        Write-Output -InputObject $objModules | Select-Object Name, ModuleType, Version, Description, CompanyName, Author, Path, Tags, CompatiblePSEditions, ExportedCommands, Prefix, `
            RepositorySourceLocation, HelpInfoUri, ClrVersion  | Sort-Object Name | Format-List
    }
}

Exit 0
