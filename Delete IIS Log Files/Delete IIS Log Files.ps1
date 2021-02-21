#requires -Version 3.0

<#
    .SYNOPSIS
    This script will delete IIS log files over a certain age and size.

    .DESCRIPTION
    The script finds the location of IIS log files and deletes the files older than X amount of days and/or over Y size. Can be run as a report only.

    .PARAMETER MinimumAgeDays
    Minimum days since last modification of the file(s) to be deleted, in days. 0 = all modification dates.

    .PARAMETER MinimumSizeMb
    Minimum size of the file(s) to be deleted, in megabytes. 0 = any size.

    .PARAMETER ReportOnly
    Report on the logs only, do not delete. True or False

    .NOTES
    The script cleans ALL log files of ALL IIS sites on the target machine that meet the criteria (not modified in last X days, over Y megabytes size)

    .COMPONENT
    Web Administration PowerShell module

    .AUTHOR
    Ton de Vreede
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory = $true, HelpMessage = 'Minimum days since last modification of the file(s) to be deleted, in days. 0 = all modification dates.')]
    [int]$MinimumAgeDays,
    [Parameter(Mandatory = $true, HelpMessage = 'Minimum size of the file(s) to be deleted, in megabytes. 0 = any size.')]
    [int]$MinimumSizeMb,
    [Parameter(Mandatory = $true, HelpMessage = 'Report on the logs only, do not delete.')]
    [string]$ReportOnly
)

# Set error handling
$ErrorActionPreference = 'Stop'

# Convert $ReportOnly to Boolean
If ($ReportOnly -eq 'True') {
    [bool]$bolReportOnly = $true
}
Else {
    [bool]$bolReportOnly = $false
}

# Import required module
try {
    Import-Module -Name WebAdministration
}
catch {
    Write-Output -InputObject "There was a problem importing the WebAdministration module: $_.Exception"
    Exit 1
}

# Set Datetime. Files older than this date are eligible for deletion
[datetime]$dtOlderThan = (Get-Date).AddDays(-$MinimumAgeDays)

# Create array for logfile locations
[array]$arrLogFileDirectories = @()

# Get the website logfiles
try {
    Foreach ($Website in Get-Website) {
        $arrLogFileDirectories += $Website.logFile.directory
    }
}
catch {
    Write-Output -InputObject "There was an error retrieving the IIS website logfile locations. The reported error is: $_"
    Exit 1
}

# Create vars for filecount and total size
[int]$intTotalDeletedFileCount = 0
[long]$lngTotalDiskSpaceFreed = 0
$lstDirectories = New-Object -TypeName System.Collections.Generic.List[PSObject]

# Get the log files, filtering on last modified and minimum size
try {
    Foreach ($LogFileLocation in $arrLogFileDirectories | Get-Unique) {
        # Set vars for directory sizes and count
        [long]$lngDirectorySpaceFreed = 0
        [int]$intDirectoryFileCount = 0

        # Go through each of the log file directories
        Foreach ($LogFile in Get-Childitem -Path $([System.Environment]::ExpandEnvironmentVariables($LogFileLocation)) -Recurse -File -Filter '*.log' | Where-Object { ($_.LastWriteTime -lt $dtOlderThan) -and ($_.Length / 1Mb -ge $MinimumSizeMb) }) {
            # If Report is True, only add the numbers for reporting, don't touch the files
            If ($bolReportOnly) {
                $lngDirectorySpaceFreed += $LogFile.Length
                $intDirectoryFileCount ++
            }
            # If report is false, try deleting the files and add the numbers from that to the directory report
            Else {
                try {
                    $null = Remove-Item -LiteralPath $LogFile.FullName -Force
                    $lngDirectorySpaceFreed += $LogFile.Length
                    $intDirectoryFileCount ++
                }
                catch {
                    Write-Output -InputObject "There was a problem deleting logfile $($LogFile.FullName): $_"
                }
            }
        }

        # Add values to directory list 
        $objDirectory = [pscustomobject]@{
            Location            = $LogFile.Directory.FullName
            Files           = $intDirectoryFileCount
            'Size (MB)' = [math]::Round($lngDirectorySpaceFreed / 1Mb, 2)
        }
        $lstDirectories.Add($objDirectory)

        # Add file number and size to totals
        $intTotalDeletedFileCount += $intDirectoryFileCount 
        $lngTotalDiskSpaceFreed += $lngDirectorySpaceFreed
    }
}
catch {
    Write-Output -InputObject "There was an error retrieving the IIS website logfiles. The reported error is: $_"
    Exit 1
}

If ($bolReportOnly) {
    Write-Output -InputObject 'Files found:'
    $lstDirectories | Format-Table -AutoSize
    Write-Output -InputObject "$intTotalDeletedFileCount log file(s) with a total size of $([math]::Round($lngTotalDiskSpaceFreed/1Mb,2)) MB were found (NOT deleted)."
}
else {
    Write-Output -InputObject 'Files deleted:'
    $lstDirectories | Format-Table -AutoSize
    Write-Output -InputObject "$intTotalDeletedFileCount log file(s) with a total size of $([math]::Round($lngTotalDiskSpaceFreed/1Mb,2)) MB were deleted."
}
