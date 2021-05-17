#requires -version 3
$ErrorActionPreference = 'Stop'

<#
    .SYNOPSIS
    This script will place the provided URLs in the system HOSTS file, redirecting them to 127.0.0.1, and flush the DNS cache.

    .DESCRIPTION
    This script can be used as a quick way to block access to a URL. The provided URLs (comma separated) are placed in a 'ControlUp' section of the HOSTS file, where they are directed to 127.0.0.1. As a result DNS lookup of these URLs always point to home, essentially preventing access to a website unless you know the IP address.
    After this the command IPCONFIG /FLUSHDNS is run to clear the DNS cache.

    .NOTES
    -   Some browsers and other programs cache DNS lookups internally. As a result users may still be able to access a website until the browser/program has been restarted.
    -   The povided URL list overwrites anything else in the ControlUp section of the HOSTS file
    -   If no URLs are provided as input, the script will remove the entire ControlUp section from the HOSTS file.

    .PARAMETER $args[0]
    The URLs to be redicted to 127.0.0.1 in comma separated format (for example youtube.com,twitch.tv,netflix.com)
#>

# Parse the URL input
if ($args) {
    [array]$arrURLs = $args[0].Split(',')
}
else {
    [array]$arrURLs = @()
}

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
        Write-Warning -Message "There was an problem.`n$Message"
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

# Set up some strings
[string]$strBeginMarker = '## ControlUp injected entries start, do not edit this section manually! ##'
[string]$strEndMarker = '## ControlUp injected entries end ##'
[string]$strHostsPath = "$([Environment]::GetFolderPath([System.Environment+SpecialFolder]::System))\Drivers\etc\hosts"

# Back up the HOSTS file
try {
    [string]$strDateShort = (Get-Date).ToString("ddMMyyyy-HHmm")
    Copy-Item -Path $strHostsPath -Destination "$($strHostsPath)-$($strDateShort).bak"
}
catch {
    Out-CUConsole -Message "The hosts file could not be backed up." -Exception $_
}

# Get the file
try {
    [string[]]$arrHosts = Get-Content -Path $strHostsPath
}
catch {
    Out-CUConsole -Message "The hosts file could not be read." -Exception $_
}

# If entries have already been added, get the start and end of the 'ControlUp' section
For ($i = 0; $i -lt $arrHosts.Count; $i ++) {
    if ($arrHosts[$i].Trim() -eq $strBeginMarker) {
        [int]$intCUSectionStart = $i
    }
    elseif ($arrHosts[$i].Trim() -eq $strEndMarker ) {
        [int]$intCUSectionEnd = $i
    }
}

# Create array containing everything EXCEPT the ControlUp section
# If the BeginMarker was found, there should be an end marker, so just check for BeginMarker
if (Test-Path Variable:\intCUSectionStart) {
    [array]$arrOut = $arrHosts[0..($intCUSectionStart - 1)]
    if (($intCUSectionEnd + 1 ) -le ($arrHosts.Count - 1)) {
        $arrOut += $arrHosts[($intCUSectionEnd + 1 )..($arrHosts.Count - 1)]
    }
}
else {
    # BeginMarker not found, looks like there is no ControlUp section in the HOSTS file. Just use the full HOSTS input as source for the output array.
    [array]$arrOut = $arrHosts
}

# Add the entries to create the ControlUp section. If the input was empty, there is no need for the ControlUp section
If ($arrURLS.Count -ne 0) {
    $arrOut += $strBeginMarker
    Foreach ($URL in $arrUrls) {
        if (!([string]::IsNullOrEmpty($URL.Trim()))) {
            $arrOut += "127.0.0.1`t$($URL.Trim())"
        }
    }
    $arrOut += $strEndMarker
}

# Write the new HOSTS file
try {
    $arrOut | Out-File -FilePath $strHostsPath -Encoding UTF8 -Force
    Out-CUConsole -Message "Hosts file written."
}
catch {
    Out-CUConsole -Message 'The HOSTS file could not be written.' -Exception $_
}

# Flush DNS cache
try {
    ipconfig /flushdns
}
catch {
    Out-CUConsole -Message 'Flushing the DNS cache failed' -Exception $_
}
