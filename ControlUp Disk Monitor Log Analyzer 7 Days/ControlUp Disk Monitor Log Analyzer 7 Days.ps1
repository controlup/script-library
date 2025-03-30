#Requires -Version 5.1

<#!
.SYNOPSIS
    Extracts and processes Disk I/O activity from multiple log files and exports the top entries to CSV.
.DESCRIPTION
    - Reads all Actvitiy Summary log files matching "ActivitiesSummary*".
    - Filters lines that start with "Timestamp".
    - Extracts relevant information using regex.
    - Sorts by TransferSizeMB in descending order and selects the top results.
    - Exports the results to a user-specified CSV file.
    - FreeSpaceGB is the freespace of the time the script is running
.PARAMETER CsvFilePath
    Full path to the output CSV file where results will be stored.
.PARAMETER Top
    Maximum number of files shown in the table and CSV.
.EXAMPLE
    .\DiskIoParser.ps1 -CsvFilePath "C:\Output\Top50Results.csv" -Top 50
.AUTHOR
    Chris Twiest and Guy Leech
.NOTES
    Modification History

    2025/03/05    Chris Twiest  Initial release
    2025/03/14    Guy Leech     Optimised. Added -raw 
#>

param (
    [Parameter(Mandatory = $true, HelpMessage = "Show amount of files")]
    [int]$Top,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the full path to the CSV output file")]
    [string]$CsvFilePath,

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
$DaysThreshold = 7
$CutoffDate = (Get-Date).AddDays(-$DaysThreshold)

# Get all matching log files
$LogFiles = Get-ChildItem -Path $LogDirectory -Filter $LogFilePattern -File 

# Filter files created within the last 7 days
$RecentLogFiles = $LogFiles | Where-Object { $_.CreationTime -ge $CutoffDate } | Sort-Object LastWriteTime -Descending

# Filter and delete files older than 7 days
$OldLogFiles = $LogFiles | Where-Object { $_.CreationTime -lt $CutoffDate }
if ($OldLogFiles) {
    Write-Output "INFO: Deleting $(($OldLogFiles).Count) log file(s) older than $DaysThreshold days..."
    $OldLogFiles | ForEach-Object {
        Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
        Write-Output "Deleted: $($_.FullName)"
    }
}

# Check if there are recent log files
if (-Not $RecentLogFiles) {
    Write-Output "ERROR: No log files found matching pattern $LogFilePattern in $LogDirectory within the last $DaysThreshold days"
    exit 1
}

# Get all matching log files
$LogFiles = Get-ChildItem -Path $LogDirectory -Filter $LogFilePattern -File | Sort-Object LastWriteTime -Descending

if (-Not $LogFiles) {
    Write-Output "ERROR: No log files found matching pattern $LogFilePattern in $LogDirectory"
    exit 1
}

# Get the system drive letter
$systemDrive = $env:SystemDrive

# Get the free space on the system drive
$drive = Get-PSDrive -Name ($systemDrive -replace ':', '')

# Calculate and display the free space in GB
$freeSpaceGB = [math]::Round($drive.Free / 1GB, 2)

# Initialize a list to store results efficiently
$Results = New-Object System.Collections.Generic.List[PSObject]

try {
    foreach ($LogFile in $LogFiles) {
        Write-Output "Reading log file: $($LogFile.FullName)"
        $LogEntries = Get-Content -Path $LogFile.FullName | Where-Object { $_ -match '^Timestamp:' }
        
        foreach ($Entry in $LogEntries) {
            if ($Entry -match 'Timestamp:(.*?), Interval:(\d+), Process Id:(\d+), ProcessName:(.*?), UserName:(.*?), FileName:(.*?), TransferSize \(MB\):([\d,\.]+)') {
                $Timestamp = $matches[1]
                $Interval = [int]$matches[2]
                $ProcessId = [int]$matches[3]
                $ProcessName = $matches[4].Trim()
                $UserName = $matches[5].Trim()
                $FileName = $matches[6].Trim()
                
                # Convert TransferSize to a double and round to 2 decimal places
                $TransferSize = [math]::Round(($matches[7] -replace ",", ".") -as [double], 2)
                
                # Create PowerShell object
                $Obj = [PSCustomObject]@{
                    Timestamp      = $( if( $timestamp = $matches[1] -as [datetime] ) { $timestamp } else { [datetime]$matches[1] })
                    Interval       = $Interval
                    ProcessId      = $ProcessId
                    ProcessName    = $ProcessName
                    UserName       = $UserName
                    FileName       = $FileName
                    TransferSizeMB = $TransferSize
                    FreeSpaceGB    = $freeSpaceGB
                }

                # Add to list
                $Results.Add($Obj)
            }
        }
    }

    if (-Not $Results) {
        Write-Output "WARNING: No matching log entries found in any log file."
        exit 0
    }

        if( $raw ) {
        $Results
    }
    else {


    # Sort and select top results
    $TopResults = $Results | Sort-Object -Property TransferSizeMB -Descending | Select-Object -First $Top

    # Output results
    $TopResults | Select-Object Timestamp, ProcessName, UserName, TransferSizeMB, FileName | Format-Table

    # Export to CSV
    $TopResults | Export-Csv -Path $CsvFilePath -NoTypeInformation -Force

    Write-Output "Top $Top results exported to: $CsvFilePath"
    }

} catch {
    Write-Output "ERROR: An error occurred while processing the log files: $_"
    exit 1
}
