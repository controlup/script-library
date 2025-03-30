#Requires -Version 5.1

<#
.SYNOPSIS
    Extracts and processes Disk I/O activity from a log file and exports the top entries to CSV.
.DESCRIPTION
    - Reads the DiskIoActivityTracker log file.
    - Filters lines that start with "Timestamp".
    - Extracts relevant information using regex.
    - Sorts by TransferSizeMB in descending order and selects the top results.
    - Exports the results to a user-specified CSV file.
.PARAMETER CsvFilePath
    Full path to the output CSV file where results will be stored.
.PARAMETER Top
    Maximum number of files shown in the table and CSV.
.EXAMPLE
    .\DiskIoParser.ps1  -CsvFilePath "C:\Output\Top50Results.csv"
.AUTHOR
    Chris Twiest and Guy Leech
.NOTES
    Modification History

    2025/03/05    Chris Twiest  Initial release
    2025/03/14    Guy Leech     Optimised. Added -raw and -daysback
#>

[CmdletBinding(DefaultParameterSetName='csv')]

param (
    [Parameter(Mandatory = $true, ParameterSetName = 'csv' , HelpMessage = "Show amount of files")]
    [int]$Top,

    [Parameter(Mandatory = $true, ParameterSetName = 'csv' , HelpMessage = "Enter the full path to the CSV output file")]
    [string]$CsvFilePath ,

    [Parameter(Mandatory = $false, ParameterSetName = 'raw')]
    [switch]$raw
)

try
{
    [int]$outputWidth = 400
    if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
    {
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions
    }
}
catch
{
    ## nothing much we can do but not fatal, just output possibly wrapping prematurely
}

$LogDirectory = Join-Path -Path ([Environment]::GetFolderPath( [Environment+SpecialFolder]::CommonApplicationData )) -ChildPath 'ControlUp\DiskMonitor'
$LogFilePattern = "ActivitiesSummary*"
$TodayStart = (Get-Date).Date  # Midnight today
$TodayEnd = $TodayStart.AddDays(1).AddSeconds(-1)  # Last second of today
$daysBack = 7 #If there are logs found older then 7 days it will clean them up
$CutoffDate = (Get-Date).AddDays( -$daysBack) 

# Get all matching log files
$LogFiles = Get-ChildItem -Path $LogDirectory -Filter $LogFilePattern -File 

# Filter only logs that were created today
$TodaysLogFiles = $LogFiles | Where-Object { $_.LastWriteTime -ge $TodayStart -and $_.LastWriteTime -le $TodayEnd } | Sort-Object LastWriteTime -Descending

Write-Verbose -Message "Start date is $($CutoffDate.ToString('G'))"
# Filter and delete files older than 7 days
$OldLogFiles = $LogFiles | Where-Object { $_.CreationTime -lt $CutoffDate }
if ($OldLogFiles) {
    Write-Output "INFO: Deleting $(($OldLogFiles).Count) log file(s) older than 7 days..."
    $OldLogFiles | ForEach-Object {
        Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
        Write-Output "Deleted: $($_.FullName)"
    }
}

# Check if there are logs created specifically today
if (-Not $TodaysLogFiles) {
    Write-Error "ERROR: No log files created today ($TodayStart - $TodayEnd) found in $LogDirectory"
    exit 1
}

$LogFilePath = $TodaysLogFiles.FullName

# Get the system drive letter
$systemDrive = $env:SystemDrive

# Get the free space on the system drive
$drive = Get-PSDrive -Name ($systemDrive -replace ':', '')

# Calculate and display the free space in GB
$freeSpaceGB = [math]::Round($drive.Free / 1GB, 2)

# Process log file
try {
    Write-Output "Reading log file: $LogFilePath"

    # Initialize a list to store results efficiently
    $Results = @( [IO.File]::ReadLines( $LogFilePath ) | Where-Object { $_ -match 'Timestamp:(.*?), Interval:(\d+), Process Id:(\d+), ProcessName:(.*?), UserName:(.*?), FileName:(.*?), TransferSize \(MB\):([\d,\.]+)' } | ForEach-Object {
          [PSCustomObject]@{
                Timestamp      = $( if( $timestamp = $matches[1] -as [datetime] ) { $timestamp } else { [datetime]$matches[1] })
                Interval       = [int]$matches[2]
                ProcessId      = [int]$matches[3]
                ProcessName    = $matches[4].Trim()
                UserName       = $matches[5].Trim()
                FileName       = $matches[6].Trim()
                TransferSizeMB = [math]::Round(($matches[7] -replace ",", ".") -as [double], 2)
                FreespaceGB    = $freeSpaceGB
          }
    })

    if( $Results.Count -eq 0 ) {
        Throw "No results read from $LogFilePath"
    }

    if( $raw ) {
        $Results
    }
    else {
        # Sort and select top results
        $TopResults = $Results | Sort-Object -Property TransferSizeMB -Descending | Select-Object -First $Top

        # Output results
        $TopResults |  Where-Object {$_.Filename -notlike "*pagefile.sys"} | Select-Object Timestamp, ProcessName, UserName, TransferSizeMB, FileName | Format-Table

        # Export to CSV
        $TopResults |  Where-Object {$_.Filename -notlike "*pagefile.sys"} | Export-Csv -Path $CsvFilePath -NoTypeInformation -Force

        Write-Output "Top $Top results exported to: $CsvFilePath"
    }

} catch {
    Write-Error "ERROR: An error occurred while processing the log file: $_"
    exit 1
}

