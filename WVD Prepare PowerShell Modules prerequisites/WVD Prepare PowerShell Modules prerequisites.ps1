<#
.SYNOPSIS
    Install the required Azure PowerShell Modules for WVD Script Actions.
.DESCRIPTION
    Install the required Azure PowerShell Modules for WVD Script Actions.
.EXAMPLE
    Get-WVDHostPool -Name <HostPoolName>
.CONTEXT
    Windows Virtual Desktops
.MODIFICATION_HISTORY
    Esther Barthel, MSc - 22/03/20 - Original code
    Esther Barthel, MSc - 22/03/20 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
    Esther Barthel, MSc - 10/05/20 - Changed the script to support WVD 2020 Spring Release (ARM Architecture update)
.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/import-clixml?view=powershell-7
.COMPONENT
    Set-AzSPCredentials - The required Azure Service Principal (Subcription level) and tenantID information need to be securely stored in a Credentials File. The Set-AzSPCredentials Script Action will ensure the file is created according to ControlUp standards
    Az.Desktopvirtualization PowerShell Module - The Az.Desktopvirtualization PowerShell Module must be installed on the machine running this Script Action
.NOTES
    Version:        1.0
    Author:         Esther Barthel, MSc
    Creation Date:  2020-03-22
    Updated:        2020-03-22
                    Standardized the function, based on the ControlUp Standards (v0.2)
    Updated:        2020-05-10
                    Changed the script to support WVD 2020 Spring Release (ARM Architecture update)
    Updated:        2020-07-22
                    Added an extra check to the Invoke-CheckInstallAndImportPSModuleRepreq function for the required NuGet packageprovider. If minimumversion 2.8.5.201 is not found it will install the NuGet PackageProvider
    Purpose:        Script Action, created for ControlUp WVD Monitoring
        
    Copyright (c) cognition IT. All rights reserved.
#>
#Requires -Version 5.1

[CmdletBinding()]
Param()    

# dot sourcing WVD Functions
function Invoke-CheckInstallAndImportPSModulePrereq() {
    <#
    .SYNOPSIS
        Check, Install (if allowed) and Import the given PSModule prerequisite.
    .DESCRIPTION
        Check, Install (if allowed) and Import the given PSModule prerequisite.
    .EXAMPLE
        Invoke-CheckInstallAndImportPSModulePrereq -ModuleName Az.DesktopVirtualization
    .CONTEXT
        Windows Virtual Desktops
    .MODIFICATION_HISTORY
        Esther Barthel, MSc - 23/03/20 - Original code
        Esther Barthel, MSc - 23/03/20 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
    .COMPONENT
    .NOTES
        Version:        1.0
        Author:         Esther Barthel, MSc
        Creation Date:  2020-03-23
        Updated:        2020-03-23
                        Standardized the function, based on the ControlUp Standards (v0.2)
        Updated:        2020-07-22
                        Added an extra check to the Invoke-CheckInstallAndImportPSModuleRepreq function for the required NuGet packageprovider. If minimumversion 2.8.5.201 is not found it will install the NuGet PackageProvider
        Purpose:        Script Action, created for ControlUp WVD Monitoring
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the PowerShell Module name that needs to be installed and imported'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $ModuleName
    )

    #region Check if the given PowerShell Module is installed (and if not, install it with elevated right)
        # Check if the Module is loaded
        If (-not((Get-Module -Name $($ModuleName)).Name))
        # Module is not loaded
        {
            Write-Verbose "* CheckAndInstallPSModulePrereq: PowerShell Module $($ModuleName) is not loaded in current session"
            # Check if the Module is installed
            If (-not((Get-Module -Name $($ModuleName) -ListAvailable).Name))
            # Module is not installed on the system
            {
                # Check if session is evelated
                [bool]$isElevated = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                if ($isElevated)
                {
                    Write-Host ("The PowerShell Module $($ModuleName) is installed from an elevated session, Scope is set to AllUsers.") -ForegroundColor Yellow
                    $psScope = "AllUsers"
                }
                else
                {
                    Write-Warning "The PowerShell Module $($ModuleName) is NOT installed from an elevated session, Scope is set to CurrentUser."
                    $psScope = "CurrentUser"
                }

                # Check the version of the installed NuGet Provider and install if the version is lower than 2.8.5.201 or the provider is missing
                try
                {
                    If (!(([string](Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue).Version) -ge "2.8.5.201"))
                    {
                        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope $psScope -Force -Confirm:$false -WarningAction SilentlyContinue 
                    }
                }
                catch
                {
                    Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
                    Exit
                }
                Write-Verbose "* CheckAndInstallPSModulePrereq: The prerequired NuGet PackageProvider is installed."

                # Install the Module from the PSGallery
                try
                {
                    # Install the Module from the PSGallery
                    Write-Verbose "* CheckAndInstallPSModulePrereq: Installing the $($ModuleName) PowerShell Module from the PSGallery."
                    PowerShellGet\Install-Module -Name $($ModuleName) -Confirm:$false -Force -AllowClobber -Scope $psScope -WarningAction SilentlyContinue
                }
                catch
                {
                    Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
                    Exit
                }
                Write-Verbose "* CheckAndInstallPSModulePrereq: The $($ModuleName) PowerShell Module is installed from the PSGallery."
            }
        }
        # Import the Module
        try 
        {
            Import-Module -Name $($ModuleName) -Force
        }
        catch 
        {
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        Write-Verbose "* CheckAndInstallPSModulePrereq: The $($ModuleName) PowerShell Module is imported in the current session."
    #endregion Check PS Module status
}


#------------------------#
# Script Action Workflow #
#------------------------#
Write-Host ""

#region Retrieve input parameters
## Check if the required PowerShell Modules are installed and can be imported
Invoke-CheckInstallAndImportPSModulePrereq -ModuleName "Az.DesktopVirtualization"
Invoke-CheckInstallAndImportPSModulePrereq -ModuleName "Az.Resources"
Invoke-CheckInstallAndImportPSModulePrereq -ModuleName "Az.Accounts"
Invoke-CheckInstallAndImportPSModulePrereq -ModuleName "AzureAD"

