#requires -version 3
$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    This script will report on packets dropped by the firewall

    .DESCRIPTION
    This script will enable firewall logging for dropped packets if not enabled, log for the specified time in minutes,
    revert to old settings if settings were changed and output the firewall log of the packets that were dropped during the logging interval.
    -   Logging is enabled for each profile (Domain, Private and Public)
    -   If any logging was already enabled for a profile, dropped packet logging is also enabled if not yet so
    -   If logging was not enabled at all, dropped packet logging is enabled with a custom file log
    -   After logging, firewall settings are reverted to old settings and custom log file is removed

    The returned output is a grouping of the IP addresses,ports and actions with the count of each combination 

    .PARAMETER intSleepMinutes
    The amount of time in minutes the script will sleep, waiting for results to be logged in firewall log

    .PARAMETER bolReceive
    Include received packets in the report

    .PARAMETER bolSend
    Include sent packets in the report

    .PARAMETER intMaxGroups
    The maximum amount of results to be returned

    .EXAMPLE
    Example is not relevant as this script will be called through ControlUp Console

    .NOTES
    This script requires the NetSecurity module, which is only available on Server 2012 (or later) Windows 8 or later.

#>

# Logging duration
[int]$intSleepMinutes = $args[0]

# Flags for RECEIVE or SEND
If ($args[1] -eq 'True') {[bool]$bolReceive = $true} Else {[bool]$bolReceive = $false}
If ($args[2] -eq 'True') {[bool]$bolSend = $true} Else {[bool]$bolSend = $false}

# Maximum amount of lines (groups) to be displayed
[int]$intMaxGroups = $args[3]

# Array for the log files to be parsed
[array]$arrLogFileNames = @()

# String for Log file location if custom file is created
[string]$strCustomLogFileName = '%systemroot%\system32\LogFiles\Firewall\_TempControlUpFirewall.log'

# Output array
[array]$ParsedLogContents = @()

# Test if at least Receive or Send have been chosen to log
If (!$bolReceive -and !$bolSend){
    Write-Host 'You have chosen to report neither sent or received packets, so there will be no output. Please enable reporting for at least one of these options.'
    Exit 0
}
Function Feedback ($strFeedbackString)
{
  # This function provides feedback in the console on errors or progress, and aborts if error has occured.
  If ($error.count -eq 0)
  {
    # Write content of feedback string
    Write-Host -Object $strFeedbackString -ForegroundColor 'Green'
  }
  
  # If an error occured report it, and exit the script with ErrorLevel 1
  Else
  {
    # Write content of feedback string but in red
    Write-Host -Object $strFeedbackString -ForegroundColor 'Red'
    
    # Display error details
    Write-Host 'Details: ' $error[0].Exception.Message -ForegroundColor 'Red'

    Exit 1
  }
}

# Try to import the NetSecurityModel module
Try {
    Import-Module NetSecurity
  }
  Catch {
    Feedback "There was an error loading the NetSecurity module. This module is only available on Server 2012 or later and in some versions of Windows 8 or later."
  }


# Get firewall settings if firewall is enabled for that profile
$objStartWindowsFirewallSettings = Get-NetFirewallProfile | Where-Object Enabled

# Check the settings, if necesary change
If ($objStartWindowsFirewallSettings)
{
    ForEach($prf in $objStartWindowsFirewallSettings)
    {
        If (($prf.LogAllowed -eq 'True') -or ($prf.LogBlocked -eq 'True') -or ($prf.LogIgnored -eq 'True'))
        {
            # Firewall log has been (partially) enabled, turn on Dropped packet logging if necessary but leave log location intact
            If ($prf.LogBlocked -ne 'True')
            {
                Set-NetFirewallProfile -Name $prf.Name -LogBlocked True
            }    
        }
        else
        {
            # Firewall log not enabled for this profile, enable Dropped packet logging and use custom log location
            # Set size to 4092 just in case the size has been set too small, will also be reverted at the end of the script
            Set-NetFirewallProfile -Name $prf.Name -LogBlocked True -LogFileName $strCustomLogFileName -LogMaxSizeKilobytes 4092
        }
    }
}
# If none of the firewalls have been enabled, write output and exit
Else
{
    Write-Host 'Firewall logging can not be enabled as the firewall is disabled entirely'
    Exit 0
}

# Record logging start time
[datetime]$dtStartTime = Get-Date

# Wait specified amount of time for logging
Start-Sleep -Seconds ($intSleepMinutes * 60)

# Revert firewall settings if they have changed, and create list of filenames to be parsed
Foreach ($prf in $objStartWindowsFirewallSettings)
{   
    # Get current firewall settings to see what needs to be reverted
    $objCurrentProfileSettings = Get-NetFirewallProfile -Name $prf.Name

    # If the log file name is different, this means that logging has been enabled for dropped packets with custom log file name
    If ($prf.LogFileName -ne $objCurrentProfileSettings.LogFileName)
    {
        Set-NetFirewallProfile -Name $prf.Name -LogBlocked $prf.LogBlocked -LogFileName $prf.LogFileName -LogMaxSizeKilobytes $prf.LogMaxSizeKilobytes
        # Add log name to list of files to be parsed
        $arrLogFileNames += $strCustomLogFileName
    }
    # If only the LogBlocked setting has been changed, only that needs to be reverted
    ElseIf ($prf.LogBlocked -ne $objCurrentProfileSettings.LogBLocked)
    {
        Set-NetFirewallProfile -Name $prf.Name -LogBlocked $prf.LogBlocked
        # Add log name to list of file to be parsed
        $arrLogFileNames += $prf.LogFileName
    }
    # No settings have changed, only add log file name to the list
    Else
    {
        $arrLogFileNames += $prf.LogFileName
    }
}

# Read contents of logfile(s) and add to $ParsedLogContents
ForEach($filLogFile in ($arrLogFileNames | Get-Unique))
{
    # Declare csv headers
    [array]$arrCsvHeader = 'Date','Time','Action','Protocol','SourceIP','DestinationIP','SourcePort','DestinationPort','size','tcpflags','tcpsyn','tcpack','tcpwin','icmptype','icmpcode','info','Direction'
    
    # Create the array with contents
    $arrLogContents += Get-Content -Path ([System.Environment]::ExpandEnvironmentVariables("$filLogFile")) | Select-Object -Skip 6 | ConvertFrom-Csv -Header $arrCsvHeader -Delimiter ' '
    
    # If custom log file was used, no need to check dates
    If ($filLogFile -eq $strCustomLogFileName)
    {
        $ParsedLogContents += $arrLogContents | Where-Object {$_.Action -eq 'DROP'} 
    }
    # Standard log file, could contain old entries so parse date and time
    Else 
    {
        Foreach ($line in $arrLogContents)
        {
            If (($line.Action -eq 'DROP') -and ([datetime]::parseexact("$($line.date) $($line.time)", 'yyyy-MM-dd HH:mm:ss', $null) -gt $dtStartTime))
            {
                $ParsedLogContents += $line 
            } 
        }
    }
}

<# Clean up custom log files
- If logging was already enabled in any way, do not clean up that file
- If logging was turned on for this script, the custom log file was created (and a version with extension .old) so only custom log file and .old version need to be removed
Simply test if these exist and remove
#>
[string]$strFileToBeRemoved = [System.Environment]::ExpandEnvironmentVariables("$strCustomLogFileName")
If (Test-Path $strFileToBeRemoved) {Remove-Item $strFileToBeRemoved}
If (Test-Path "$strFileToBeRemoved.old") {Remove-Item "$strFileToBeRemoved.old"}

# Select Send, Receive or both and sort the output by datetime and output
# Grouped by source and destination
    # Some logic to work with the flags
    If ($bolReceive -and $bolSend) {
        $tmpReport = $ParsedLogContents | Select-Object -Property Action,Direction,Protocol,SourceIP,SourcePort,DestinationIP,DestinationPort | Group-Object -Property Action,Direction,Protocol,SourceIP,SourcePort,DestinationIP,DestinationPort | Sort-Object Count -Descending
        }
    elseif ($bolReceive) {
        $tmpReport = $ParsedLogContents | Where-Object {($_.Direction -eq 'RECEIVE')} | Select-Object -Property Action,Direction,Protocol,SourceIP,SourcePort,DestinationIP,DestinationPort | Group-Object -Property Action,Direction,Protocol,SourceIP,SourcePort,DestinationIP,DestinationPort  | Sort-Object Count -Descending
        }
    else {
        $tmpReport = $ParsedLogContents | Where-Object {($_.Direction -eq 'SEND')} | Select-Object -Property Action,Direction,Protocol,SourceIP,SourcePort,DestinationIP,DestinationPort | Group-Object -Property Action,Direction,Protocol,SourceIP,SourcePort,DestinationIP,DestinationPort  | Sort-Object Count -Descending
        }


# Put it all together in a PSCustomObject
If ($tmpReport.Count -ne 0){
    $FinalReport = @( Foreach ($obj in $tmpReport | Select-Object -First $intMaxGroups) {
        $objEx = $obj | Select-object -expandproperty Group
        [pscustomobject][ordered]@{
            Count                   = $obj.Count
            Action                  = $objEX[0].Action
            Direction               = $objEX[0].Direction
            Protocol                = $objEX[0].Protocol
            SourceIP                = $objEX[0].SourceIP
            SourcePort              = $objEX[0].SourcePort
            DestinationIP           = $objEX[0].DestinationIP
            DestinationPort         = $objEX[0].DestinationPort
        }
    })
    # Output the report
    $FinalReport | Sort-Object Count,SourceIP -Descending | Format-Table -AutoSize
}
Else {
    Feedback 'No dropped packets were logged during the logging interval.'
}


