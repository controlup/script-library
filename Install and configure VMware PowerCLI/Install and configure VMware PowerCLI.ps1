$ErrorActionPreference = 'Stop'
<#
Installs Nuget packageprovider and VMware PowerCLI module for working with VMWare and Powershell for All Users
- Used module Install-Module requires Powershell 5.0 minimum!
- The script will overwrite any existing PowerCLI modules
- MSI based modules must be uninstalled first

Changelog
- Wouter Kursten - June 2020
    - added Nuget installation
    - Added CEIP configuration
    - Added Invalid Certificate Action configuration
#>

# CEIP setting true to send data, false to not send data
[string]$CEIPSetting = $args[0]
# Invalid Certificate handling options are fail,ignore and warn
[string]$InvalidCertificateAction = $args[1]

$CEIPbool = [System.Convert]::ToBoolean($CEIPSetting)

Function Out-CUConsole {
    <# This function provides feedback in the console on errors or progress, and aborts if error has occured.
      If only Message is passed this message is displayed
      If Warning is specified the message is displayed in the warning stream (Message must be included)
      If Stop is specified the stop message is displayed in the warning stream and an exception with the Stop message is thrown (Message must be included)
      If an Exception is passed a warning is displayed and the exception is thrown
      If an Exception AND Message is passed the Message message is displayed in the warning stream and the exception is thrown
    #>

    Param (
        [Parameter(Mandatory = $false)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [switch]$Warning,
        [Parameter(Mandatory = $false)]
        [switch]$Stop,
        [Parameter(Mandatory = $false)]
        $Exception
    )

    # Throw error, include $Exception details if they exist
    if ($Exception) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        If ($Message) {
            Write-Warning -Message "$Message`n$($Exception.CategoryInfo.Category)`nPlease see the Error tab for the exception details."
            Write-Error "$Message`n$($Exception.Exception.Message)`n$($Exception.CategoryInfo)`n$($Exception.Exception.ErrorRecord)" -ErrorAction Stop
        }
        Else {
            Write-Warning "There was an unexpected error: $($Exception.CategoryInfo.Category)`nPlease see the Error tab for details."
            Throw $Exception
        }
    }
    elseif ($Stop) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        Write-Warning -Message "There was an error.`n$Message"
        Throw $Message
    }
    elseif ($Warning) {
        # Write the warning to Warning stream, thats it. It's a warning.
        Write-Warning -Message $Message
    }
    else {
        # Not an exception or a warning, output the message
        Write-Output -InputObject $Message
    }
}

# Check if Powershell 5 is used, minimum version for using Install-Module
If ($psversiontable.psversion.Major -ge 5){
    # Temporarily change security protocol to the only one accepted by the Powershell Gallery
    $oldprot=[Net.ServicePointManager]::SecurityProtocol
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try {
        Install-PackageProvider -Name NuGet -force
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
    }
    catch {
        try{
            # this mostly fails because there was no repository configured
            Register-PSRepository -default
            Install-PackageProvider -Name NuGet -force
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
        }
        catch{
            out-cuconsole -message "NuGet Provider could not be installed or updated." -Exception $_
            exit 0
        }
    }
    try {
        Install-Module VMWare.PowerCLI -SkipPublisherCheck -AllowCLobber -Scope AllUsers -confirm:$false
    }
    catch {
        out-cuconsole -message "PowerCLI could not be installed." -Exception $_
    }
    # Put back old security protocol setting
    [Net.ServicePointManager]::SecurityProtocol = $oldprot
}
else
{
    out-cuconsole -message "This machine has PowerShell version $($PSversionTable.PSVersion)`nThis version is too low to use this script, please install PowerShell 5.1 or higher or use a different machine for PowerCLI." -Warning
}


# This changes the CEIP settings
try{
    Set-PowerCLIConfiguration -ParticipateInCeip $CEIPbool -scope allusers -confirm:$false
    }
catch{
    out-cuconsole -message "Error changing CEIP setting to $CEIPSetting" -Exception $_
}

# this changes the invalid certificate action default = Fail

try{
    Set-PowerCLIConfiguration -InvalidCertificateAction $InvalidCertificateAction -scope allusers -confirm:$false
}
catch{
    Out-CUConsole -message "Error settings the invalid certificate action to $InvalidCertificateAction" -Exception $_
}
