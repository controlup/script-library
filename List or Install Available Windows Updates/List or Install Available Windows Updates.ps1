<#
    .SYNOPSIS
        Run a Windows Update Report or Installs Windows Updates

    .DESCRIPTION
        Run a Windows Update Report or Installs Windows Updates

    .PARAMETER  <Update <switch>>
        Runs Windows Update to install any missing updates
		
	.PARAMETER	<List <switch>>
		Lists any detected Windows Updates

    .PARAMETER	<AllUpdates <switch>>
		Retrieves all updates, including drivers

    .PARAMETER	<PatchesOnly <switch>>
		Retrieves all Windows Patches

	.PARAMETER	<SaveOutputTo <string>>
		Saves the output of the script to a CSV file. If the CSV file isn't specified the file is created with a format %YEAR%-%MONTH%-%DAY%.csv

    .EXAMPLE
        . .\WindowsUpdate.ps1 -List -PatchesOnly
        Lists all Windows Update detected as needed by this system.

    .EXAMPLE
        . .\WindowsUpdate.ps1 -AllUpdates -List
        Lists all updates, including drivers, detected as needed by this system.

    .EXAMPLE
        . .\WindowsUpdate.ps1 -Update
        Installs any detected updates

    .EXAMPLE
        . .\WindowsUpdate.ps1 -List -SaveOutputTo D:\WindowsUpdatesToInstall.csv
        Lists all Windows Update detected as needed by this system and saves the output to the specified file.

    .EXAMPLE
        . .\WindowsUpdate.ps1 -List -SaveOutputTo \\mwss01.jupiterlab.com\fileshare\AvailableWindowsUpdate.csv
        Lists all Windows Update detected as needed by this system and saves the output to the specified file.

    .EXAMPLE
        . .\WindowsUpdate.ps1 -List -SaveOutputTo \\mwss01.jupiterlab.com\fileshare\MissingWindowsUpdate
        Lists all Windows Update detected as needed by this system and saves the output to the specified directory with the default filename.
        The default filename is %YEAR%-%MONTH%-%DAY% from the date the script was run. An example name is 2022-10-13.csv.

    .CONTEXT
        Machine

    .MODIFICATION_HISTORY
        Created TTYE : 2020-06-22
        Updated TTYE : 2022-10-13 - added SaveOutputTo optional parameter.


    AUTHOR: Trentent Tye
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$false,HelpMessage='"-list" will display any available updates')]                [switch]$list,
    [Parameter(Mandatory=$false,HelpMessage='"-update" will apply any available updates')]                [switch]$update,
    [Parameter(Mandatory=$false,HelpMessage='"-AllUpdates" will display all updates, including drivers')] [switch]$AllUpdates,
    [Parameter(Mandatory=$false,HelpMessage='"-PatchesOnly" will display only Windows Updates')]          [switch]$PatchesOnly,
    [Parameter(Mandatory=$false,HelpMessage='Saves the output to a CSV file.')]                           [string]$SaveOutputTo = "$($Env:windir)\Temp"
)

function Write-Header {
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory=$true,HelpMessage='Path to the file to create the headers')] [string]$Path
    )
    $stream = [System.IO.StreamWriter]::new($Path)
    $stream.WriteLine("Machine,Date,KB,Patch Description,URL")
    $stream.close()

    if (-not(Test-Path $Path)) { ## We should have a file created now.
        Write-Error "Unable to write to the target file: $Path"
    }
}

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
if ($AllUpdates) {
    $SearchQuery = "IsInstalled=0 and IsHidden=0"
} else {
    $SearchQuery = "IsInstalled=0 and Type='Software' and IsHidden=0"
}
$searchTime = Measure-Command {$Result = $Searcher.Search($SearchQuery)} ## Original String - "IsInstalled=0 and Type='Software' and IsHidden=0". Removed "Type='Software'" to allow scanning for drivers / firmware too
if ($searchTime.TotalSeconds -gt 60) {
    Write-Verbose "Search took $($searchTime.Hours) Hours(s) $($searchTime.Minutes) minute(s) $($searchTime.Seconds) second(s)" 
}
else
{
    Write-Verbose "Search took $($searchTime.TotalSeconds) seconds" 
}
    
If ($Result.Updates.Count -EQ 0) {
	Write-Host "There are no applicable updates for this computer."
} Else {
	Write-Verbose  "=============================================================================="
	Write-Host "List of Applicable Updates:"
	For ($Counter = 0; $Counter -LT $Result.Updates.Count; $Counter++) {
		$DisplayCount = $Counter + 1
    	$FoundUpdate = $Result.Updates.Item($Counter)
		Write-Host  "$DisplayCount -- KB$($FoundUpdate.KBArticleIDs) -- $($FoundUpdate.Title)"
	}

    if (Get-Variable SaveOutputTo -ErrorAction SilentlyContinue) {

        # Set Default File Name
        $DateTime = [datetime]::Today
        if (($SaveOutputTo.EndsWith("\")) -or (-not($SaveOutputTo.EndsWith(".csv")))) {
            if (-not($SaveOutputTo.EndsWith("\"))) { $SaveOutputTo = $SaveOutputTo + "\" }
            $SaveOutputTo = "$SaveOutputTo" + "$($DateTime.Year)-$($DateTime.Month)-$($DateTime.Day).csv"
            Write-Verbose "Path updated to $SaveOutputTo"
            ## Check to see if file already exists
            if (-not(Test-Path $SaveOutputTo)) {
                Write-Verbose "Creating default CSV file with headers"
                Write-Header -Path $SaveOutputTo
            }
        }

        if ($SaveOutputTo.EndsWith(".csv")) { #ensure CSV headers are present
            if (-not(Test-Path $SaveOutputTo)) {
                Write-Header -Path $SaveOutputTo
            } else {
                ## check if the header exists on the file
                if ((Get-Content $SaveOutputTo -First 1) -eq "Machine,Date,KB,Patch Description,URL") {
                    Write-Verbose "File exists with headers"
                } else {
                    Write-Verbose "File exists without headers."
                    Write-Header -Path $SaveOutputTo
                }
            }
        }

        ## File should be created, we'll dump the patching info into it.
        ## Since this script could be used for automation for reporting on the status of updates and the file might be stored on a file share with latency, it might be locked by another operation when we try and write to it.
        ## To avoid contention I'll test to see if it's locked and retry a few times with some random delays. If it fails after a few retries I'll error out.
        $DateTime = [datetime]::Today
        $ShortDateString = $DateTime.ToShortDateString()
        Write-Verbose "Examining Updates..."
        Foreach ($WinUpdate in $Result.Updates) {
            Write-Verbose "Examining KB$($WinUpdate.KBArticleIDs)"
            [int]$LoopCount = 0
            $ExceptionFound = $false
            Write-Verbose "Examining Update : KB$($WinUpdate.KBArticleIDs)"
            Do {
                $ExceptionFound = $false
                try {
                    $stream = [System.IO.StreamWriter]::new($SaveOutputTo, $true)
                    Write-Verbose "Wrote line: `"`"$($env:COMPUTERNAME)`",`"$ShortDateString`",`"$($WinUpdate.KBArticleIDs)`",`"$($WinUpdate.Description)`",`"$($WinUpdate.MoreInfoUrls)`""
                    $stream.WriteLine("`"$($env:COMPUTERNAME)`",`"$ShortDateString`",`"$($WinUpdate.KBArticleIDs)`",`"$($WinUpdate.Description)`",`"$($WinUpdate.MoreInfoUrls)`"")
                } catch {                                                             # if error encountered whilst setting up StreamReader, try again up to 5 times.
                    $ExceptionFound = $true
                    $LoopCount++
                    $delay = $(Get-Random -Minimum 1 -Maximum 10)
                    Write-Verbose "Failed to write to the file. Delaying $delay seconds and retrying..."
                    Start-Sleep -Seconds $delay
                } finally {
                    $stream.Close()
                    $stream.Dispose()
                }
                if ($LoopCount -ge 5) { Write-Error "Failed to write to $SaveOutputTo" }
            } Until ($LoopCount -ge 5 -or $ExceptionFound -eq $false)
        }
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
