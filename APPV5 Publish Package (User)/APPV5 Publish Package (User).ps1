#Requires -Version 3.0

<#
    .SYNOPSIS
    This script will publish an App-V application to the user

    .DESCRIPTION
    This script will publish an App-V application to the user

    .LINK
    http://virtualengine.co.uk

    NAME: N/A
    AUTHOR: Nathan Sperry, Virtual Engine
    LASTEDIT: 05/06/2015
    VERSI0N : 1.0
    WEBSITE: http://www.virtualengine.co.uk

#>

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

if ($Installed -ne $null) {return $true} else {return $false}

}

$appvname = $args[0]

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

    try
    {
        
        $allpackages = Get-AppvClientPackage -All
        $packages = $allpackages | where {$_.Name -like $appvname}

        if ($packages.count -ge 1)
        {

                foreach ($package in $packages)
                {
                    If ($package.IsPublishedToUser -eq $false)
                    {
                         $result = Publish-AppvClientPackage -Name $package.Name
                         Write-host $package.Name 'has been published to the user'
                    }
                    else
                    {
                        Write-Output ($package.Name + ' is already published to the user')
                    }

                }

        }
        else
        {
            Write-Warning "No App-V packages that match '$appvname' are present on this device"
        }
    }
    catch
    {
     
        $ErrorMessage = $_.Exception.Message
        Write-Output $ErrorMessage

    }
}
Else
{
    Write-Warning 'App-V 5.x client is not installed'
}
