#Requires -Version 3.0

<#
    .SYNOPSIS
    This script will return any App-V packages present

    .DESCRIPTION
    This script will return any App-V packages present

    .LINK
    http://virtualengine.co.uk

    NAME: N/A
    AUTHOR: Nathan Sperry, Virtual Engine
    LASTEDIT: 05/06/2015
    VERSI0N : 1.0
    WEBSITE: http://www.virtualengine.co.uk

#>


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

if ($Installed -ne $null) {return $true} else {return $false}

}


if (Get-AppVClient -eq $true)
{

    # Import App-V PoSH Module to make sure its loaded
    If ( ! (Get-module AppVClient ))
    { 

        # Find Installation Path
        $strAppVClient = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\AppV\Client
        $strInstallPath = $strAppVClient.InstallPath
        Import-Module ($strInstallPath + "AppvClient\AppvClient.psd1")
    }

    $result = Get-AppvClientPackage -All

    If ($result -ne $null) {$result} else {Write-Output ('No App-V Packages are present')}

}
Else
{
    Write-Warning ('App-V 5.x client is not installed')
}
