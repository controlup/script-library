#requires -Version 3.0
<#
    .SYNOPSIS
    This script will parse ControlUp Export files for the session, machine and usernam selected and provide a list of the Active Application (+Title) and Active URL (if the active application is a browser).

    .DESCRIPTION
    The script will parse the records from the last X minutes (enter '0' to get all records from the past year). For all records with matching SessionID, UserName and MachineName the Active Application, Active Application Title (if the active application is a browser) and Active URL are displayed. 

    .PARAMETER strExportPath
    The path where the ControlUp exported Session files are stored.

    .PARAMETER strLastXMinutes
    Amount of time to go back in history in minutes. 0 = parse all available export files in export folder (max one year old).

#>

[CmdLetBinding()]
    Param (
    [Parameter(Mandatory = $true, HelpMessage = 'Export files path.')]
    [string]$strExportPath,
    [Parameter(Mandatory = $true, HelpMessage = 'Amount of time to go back in history in minutes. 0 = parse all available export files in export folder (max one year old).')]
    [string]$strLastXMinutes,
    [Parameter(Mandatory = $true, HelpMessage = 'Machine name.')]
    [string]$strMachineName,
    [Parameter(Mandatory = $true, HelpMessage = 'User SessionID.')]
    [string]$strSessionID,
    [Parameter(Mandatory = $true, HelpMessage = 'User name.')]
    [string]$strUserName
)

$ErrorActionPreference = 'Stop'

# First set the desired oldest file date
If ($strLastXMinutes -ne '0') {
    [datetime]$dtNewerThan = (Get-Date).AddMinutes( - [decimal]$strLastXMinutes)
}
Else {
    [datetime]$dtNewerThan = (Get-Date).AddMonths(-12)
}

# Get the CSV files
try {
    $filInputFiles = (Get-ChildItem -Path $strExportPath -Filter "*Sessions*.csv"  | Where-Object { $_.CreationTime -gt $dtNewerThan } | Sort-Object -Property CreationTime)
}
catch {
    Write-Warning -Message "There was an issues collecing files from export path $strExportPath`. Please check the path and file permissions are correct. The reported error was:`n $($_.Exception.ErrorRecord)"
    Exit 1
}

# See if any CSV files were found
If ($filInputFiles.Count -eq 0) {
    # Check to see if there are any Session export files can be found at all
    $filInputFiles = (Get-ChildItem -Path $strExportPath -Filter "*Sessions*.csv" | Sort-Object -Property CreationTime)
    If ($filInputFiles.Count -eq 0) {
        Write-Warning -Message "No Session export files were found in the path $strExportpath`."
        Exit 1
    }
    Else {
        Write-Warning -Message "Session export files were found in the path $strExportpath but they appear to be older than the maximum number of minutes old specified. The most recent Session export file was written on $($filInputFiles[0].CreationTime.ToString())`.
        Either specify a longer history to search or decrease the time interval Session data exports should be made and wait for this interval to pass before trying again."
        Exit 1
    }

}

# Create object to hold matching records
$objMatchingRecords = New-Object -TypeName System.Collections.Generic.List[PSObject]

# Go through the files, add each record that matches to $objMatchingRecords
Foreach ($File in $filInputFiles) {
    $Records = Get-Content -Path $File.FullName | Select-Object -Skip 1 | ConvertFrom-Csv | Where-Object { ($_.'Machine' -eq $strMachineName) -and ($_.User -eq $strUserName) -and ($_.ID -eq $strSessionID) }
    Foreach ($Record in $Records) {
        # Create PSCustomObject with user details
        $objMatch = [PSCustomObject]@{
            'Date & Time'              = $file.CreationTime
            'Active Application'       = $Record.'Active Application'
            'Active Application Title' = $Record.'Active Application Title'
            'Active URL'               = $Record.'Active URL'
            'State'                    = $Record.State
        }
        $objMatchingRecords.Add($objMatch)
    }
}

# Sort the output into condensed entries
# Create object to hold matching records
$objCondensedOutput = New-Object -TypeName System.Collections.Generic.List[PSObject]

# Create hashtable to store records while searching for app last time seen
$hshCurrentRecord = @{}

# Go through the record. If a record contains application information
Foreach ($Record in $objMatchingRecords | Sort-Object 'Date & Time' ) {
    If ($Record.State -eq 'Active') {
        # The Session is active, so there could be data for app etc.
        If ($hshCurrentRecord.Count -eq 0 -and $Record.'Active Application' -ne '') {
            # The current record is empty and source record does contain app data, so add the data
            $hshCurrentRecord.Add('First seen', $Record.'Date & Time')
            $hshCurrentRecord.Add('Last seen', $Record.'Date & Time')
            $hshCurrentRecord.Add('Active Application', $Record.'Active Application')
            $hshCurrentRecord.Add('Active Application Title', $Record.'Active Application Title')
            $hshCurrentRecord.Add('Active URL', $Record.'Active URL')
        }
        elseif ($hshCurrentRecord.Count -eq 5) {
            If (($hshCurrentRecord.'Active Application' -eq $Record.'Active Application') -and ($hshCurrentRecord.'Active URL' -eq $Record.'Active URL')) {
                # The source record has the same app data as the current record being created, so update the last seen timestamp
                $hshCurrentRecord.'Last Seen' = $Record.'Date & Time'
            }
            else {
                # The source record has different app data as the current record being created, so commit the current record and create a new one containing the new app data
                $objMatch = [PSCustomObject]@{
                    'First Seen'               = $hshCurrentRecord.'First Seen'
                    'Last Seen'                = $hshCurrentRecord.'Last Seen'
                    'Active Application'       = $hshCurrentRecord.'Active Application'
                    'Active Application Title' = $hshCurrentRecord.'Active Application Title'
                    'Active URL'               = $hshCurrentRecord.'Active URL'
                }
                $objCondensedOutput.Add($objMatch)
                $hshCurrentRecord.Clear()
                
                if ($Record.'Active Application' -ne '') {
                    # The source record contains appdata, create a new Current Record
                $hshCurrentRecord.Add('First seen', $Record.'Date & Time')
                $hshCurrentRecord.Add('Last seen', $Record.'Date & Time')
                $hshCurrentRecord.Add('Active Application', $Record.'Active Application')
                $hshCurrentRecord.Add('Active Application Title', $Record.'Active Application Title')
                $hshCurrentRecord.Add('Active URL', $Record.'Active URL')
                }
            }
        }

    }
    elseif ($hshCurrentRecord.Count -eq 5) {
        # The session is not in state Active, so the record needs to be commited as there is no app data
        $objMatch = [PSCustomObject]$hshCurrentRecord
        $objCondensedOutput.Add($objMatch)
        $hshCurrentRecord.Clear()
    }
}

if ($hshCurrentRecord.Count -eq 5) {
    # The last source record has been processed, commit the open current record
    $objMatch = [PSCustomObject]@{
        'First Seen'               = $hshCurrentRecord.'First Seen'
        'Last Seen'                = $hshCurrentRecord.'Last Seen'
        'Active Application'       = $hshCurrentRecord.'Active Application'
        'Active Application Title' = $hshCurrentRecord.'Active Application Title'
        'Active URL'               = $hshCurrentRecord.'Active URL'
    }
    $objCondensedOutput.Add($objMatch)
}

# Output the results
if ($objCondensedOutput.Count -eq 0) {
    Write-Warning -Message "No records matching the criteria were found. Perhaps you should try a larger timeframe?"
}
else {
    $objCondensedOutput | Where-Object {$_.'Active Application' -ne 'wfshell.exe'} | Sort-Object -Property 'First Seen' -Descending | Out-Gridview -Title "Active applications and browser URLs for session $strSessionID ($strUserName) on $strMachineName" -Wait
}





