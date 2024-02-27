
$ErrorActionPreference = 'Stop'

# Check if App-V client is installed
Function Get-AppVClient{

    <#
    .SYNOPSIS
    This function determines if the App-V client is installed
    .DESCRIPTION
    This function determines if the App-V client is installed
    .EXAMPLE
    Get-AppVClient
    Returns true or false.
    .NOTES
    NAME: Get-AppVClient
    AUTHOR: Nathan Sperry, Virtual Engine
    LASTEDIT: 05/06/2015
    WEBSITE: http://www.virtualengine.co.uk
    KEYWORDS: App-V,App-V ,VirtualEngine,AppV5
    .LINK
    http://virtualengine.co.uk
    #>

    ## TTYE - check if this is the built-in AppV
    if ([boolean](Get-Command -Name Get-AppvStatus -ErrorAction SilentlyContinue)) {
        if ((Get-AppvStatus).AppVClientEnabled -eq $true) {
            return $true
        }
    }

    $Installed = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | where Displayname -match 'Microsoft Application Virtualization' | Select-Object Displayname

    if ($Installed -ne $null) {
        return $true
    } else {
        return $false
    }
}




if (Get-AppVClient -eq $true) {
    If ( (Get-Module -Name AppvClient -ErrorAction SilentlyContinue) -eq $null ) {
            # using try/catch can stop the script completely if needed with "Exit with error" - 'Exit 1' (or some other non-zero exit code)
            # and avoid a long string of errors because the first statement was not successful.
            Try {
                    Import-Module AppvClient
            } Catch {
                    Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
                    Exit 1
            }
    }
}

Sync-AppvPublishingServer -ServerId 1 -global

