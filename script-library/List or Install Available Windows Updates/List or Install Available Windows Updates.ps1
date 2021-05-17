<#
    .SYNOPSIS
        Run a Windows Update Report or Installs Windows Updates

    .DESCRIPTION
        Run a Windows Update Report or Installs Windows Updates

    .PARAMETER  <Update <switch>>
        Runs Windows Update to install any missing updates
		
	.PARAMETER	<List <switch>>
		Lists any detected Windows Updates


    .EXAMPLE
        . .\WindowsUpdate.ps1 -List
        Lists all Windows Update detected as needed by this system.

    .EXAMPLE
        . .\WindowsUpdate.ps1 -Update
        Installs any detected updates

    .CONTEXT
        Machine

    .MODIFICATION_HISTORY
        Created TTYE : 2020-06-22


    AUTHOR: Trentent Tye
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$false,HelpMessage='"-list" will display any available updates')] [switch]$list,
    [Parameter(Mandatory=$false,HelpMessage='"-update" will apply any available updates')] [switch]$update
)

###$verbosePreference = "continue"
#Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Start or restart the service
Write-Verbose "Starting or restarting the Windows Update service"
if ((Get-Service -Name wuauserv).Status -eq "Running") {
        Write-Verbose "Windows Update service found running"
}

if ((Get-Service -Name wuauserv).Status -ne "Running") {
    Write-Verbose "Windows Update service was NOT running"
    try {
        Set-Service -Name wuauserv -StartupType Automatic -ErrorAction Stop
    }
    catch {
        Write-Host "Unable to set wuauserv service to Automatic"
    }
    $service = Get-Service -Name wuauserv
    if ($service.Status -ne "Running") {
        try {
        $service.Start()
        $service.WaitForStatus('Running')
        }
        catch {
            Write-Error "Unable to start wuauserv service"
            exit
        }
    }
}

$UpdateCollection = New-Object -ComObject Microsoft.Update.UpdateColl
$Searcher = New-Object -ComObject Microsoft.Update.Searcher
$Session = New-Object -ComObject Microsoft.Update.Session
    
Write-Verbose "Initialising and Checking for Applicable Updates. Please wait ..."
$searchTime = Measure-Command {$Result = $Searcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")}
if ($searchTime.TotalSeconds -gt 60) {
    Write-Verbose "Search took $($searchTime.Hours) Hours(s) $($searchTime.Minutes) minute(s) $($searchTime.Seconds) second(s)" 
}
else
{
    Write-Verbose "Search took $($searchTime.TotalSeconds) seconds" 
}
    
If ($Result.Updates.Count -EQ 0) {
	Write-Host "There are no applicable updates for this computer."
}
Else {
	Write-Verbose  "=============================================================================="
	Write-Host "List of Applicable Updates:"
	For ($Counter = 0; $Counter -LT $Result.Updates.Count; $Counter++) {
		$DisplayCount = $Counter + 1
    	$FoundUpdate = $Result.Updates.Item($Counter)
		Write-Host  "$DisplayCount -- KB$($FoundUpdate.KBArticleIDs) -- $($FoundUpdate.Title)"
	}
}



### Apply update
if ($update) {
    $Counter = 0
    $DisplayCount = 0
    Write-Verbose "Initialising Download of Applicable Updates ..."
    Write-Verbose  "------------------------------------------------"
    $searchTime = Measure-Command {$Downloader = $Session.CreateUpdateDownloader()}
    Write-Verbose "Download Initialization took $($searchTime.TotalSeconds)"
    $UpdatesList = $Result.Updates
    $searchTime = Measure-Command {
        For ($Counter = 0; $Counter -LT $Result.Updates.Count; $Counter++) {
		    $UpdateCollection.Add($UpdatesList.Item($Counter)) | Out-Null
		    $ShowThis = $UpdatesList.Item($Counter).Title
		    $DisplayCount = $Counter + 1
		    Write-Verbose  "$DisplayCount -- Downloading Update $ShowThis "
		    $Downloader.Updates = $UpdateCollection
		    $Track = $Downloader.Download()
		    If (($Track.HResult -EQ 0) -AND ($Track.ResultCode -EQ 2)) {
			    Write-Verbose  "Download Status: SUCCESS" 
		    }
		    Else {
			    Write-Error  "Download Status: FAILED With Error -- $Error"
			    $Error.Clear()
		    }	
	    }
    }
    if ($searchTime.TotalSeconds -gt 60) {
        Write-Verbose "Download took $($searchTime.Minutes) minute(s) $($searchTime.Seconds) second(s)"
    }
    else
    {
        Write-Verbose "Download took $($searchTime.TotalSeconds) seconds"
    }

        
    $Counter = 0
    $DisplayCount = 0
    Write-Verbose "Starting Installation of Downloaded Updates ..."
    Write-Host  "`nInstallation:"
    Write-Host  "------------------------------------------------"
    $searchTime = Measure-Command {	
        ForEach ($UpdateFound in $UpdateCollection) {
            $Track = $Null
            $DisplayCount = $DisplayCount + 1
            $WriteThis = $UpdateFound.Title
		    Write-Host  "$DisplayCount -- Installing Update: $WriteThis"
            $Installer = New-Object -ComObject Microsoft.Update.Installer
            $UpdateToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
            $UpdateToInstall.Add($UpdateFound) | out-null
		    $Installer.Updates = $UpdateToInstall
		    Try {
			    $Track = $Installer.Install()
			    Write-Host  "Installation Status: SUCCESS" 
		    }
		    Catch {
			    [System.Exception]
			    Write-Error  "Update Installation Status: FAILED With Error -- $Error()"
			    $Error.Clear()
            }
	    }
    }
    if ($searchTime.TotalSeconds -gt 60) {
        Write-Verbose "Install took $($searchTime.Hours) Hour(s) $($searchTime.Minutes) minute(s) $($searchTime.Seconds) second(s)"
    }
    else
    {
        Write-Verbose "Install took $($searchTime.TotalSeconds) seconds"
    }
    if ($DisplayCount -ge 1) {
	    Write-Host "Updates were installed.  Reboot to apply..."
    }
}
