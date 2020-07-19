#requires -version 3.0

<#
    Get Citrix PVS target boot time events from event log and convert to CSV for reporting or alerting purposes

    Ensure that each PVS server's stream service has event logging enabled

    Guy Leech, 2017

    Modification history:

    13/02/18   GL   Added chart view option
#>

[string[]]$computers = @( 'localhost' ) 
[string]$last = $null

if( $args.Count -ge 1 -and $args[0] )
{
    $last = $args[0]
}

[string]$providerName = 'StreamProcess' 
[int]$eventId = 10 
[string]$eventLog = 'Application' 
[int]$outputWidth = 400

[array]$events = @()
[int]$slowest = 0
[int]$fastest = [int]::MaxValue
[long]$totalTime = 0
[int]$count = 1
[hashtable]$modes = @{}
[dateTime]$startDate = (Get-Date).AddYears( -20 ) ## Should be long enough ago!

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

if( ! [string]::IsNullOrEmpty( $last ) )
{
    ## see what last character is as will tell us what units to work with
    [int]$multiplier = 0
    switch( $last[-1] )
    {
        "s" { $multiplier = 1 }
        "m" { $multiplier = 60 }
        "h" { $multiplier = 3600 }
        "d" { $multiplier = 86400 }
        "w" { $multiplier = 86400 * 7 }
        "y" { $multiplier = 86400 * 365 }
        default { Throw "Unknown time multiplier `"$($last[-1])`"" }
    }
    $endDate = Get-Date
    if( $last.Length -le 1 )
    {
        $startDate = $endDate.AddSeconds( -$multiplier )
    }
    else
    {
        $startDate = $endDate.AddSeconds( - ( ( $last.Substring( 0 ,$last.Length - 1 ) -as [int] ) * $multiplier ) )
    }
}

[hashtable]$targets = @{}

$events = @( ForEach( $computer in $computers )
{
    Write-Verbose "$count / $($computers.Count ) : processing $computer from $startDate"
    @( Get-WinEvent -ComputerName $computer -FilterHashtable @{Logname=$eventLog;ID=$eventId;ProviderName=$providerName;StartTime=$startDate} | Where-Object { $_.Message -match 'boot time'}|select TimeCreated,Message | ForEach-Object `
    {
        ## Message will be "Device xxxxx boot time: 2 minutes 50 seconds."
        if( $_.Message -match '^Device (?<Target>[^\s]+) boot time: (?<minutes>\d+) minutes (?<seconds>\d+) seconds\.$' )
        {
            [int]$boottime = ( $matches[ 'minutes' ] -as [int] ) * 60 + ( $matches[ 'seconds' ] -as [int] )
            [pscustomobject][ordered]@{ 'TimeCreated' = $_.TimeCreated ; 'Target' = $matches[ 'Target' ] ; 'BootTime' = $boottime }
            $totalTime += $boottime
            if( $boottime -gt $slowest )
            {
                $slowest = $boottime
            }
            if( $boottime -lt $fastest )
            {
                $fastest = $boottime
            }
            ## Add to hash table for mode calculation
            try
            {
                $modes.Add( $boottime , 1 )
            }
            catch
            {
                $modes.Set_Item( $boottime , $modes[ $boottime ] + 1 )
            }
            ## Add to targets collection so we can count how many unique ones
            try
            {
                $targets.Add( $matches[ 'Target' ] , $_.TimeCreated )
            }
            catch {} ## already got it
        }
    })
    $count++
} )

if( $events -and $events.Count -gt 0 )
{
    ## Now find median (middle) value
    [array]$sorted = $events | select BootTime | sort BootTime

    ## Now find mode (commonest) value
    [int]$mode = 0
    [int]$lastHighestCount = 0
    [int]$highestCount = 0

    $modes.GetEnumerator() | ForEach-Object `
    {
        if( $_.Value -gt $highestCount )
        {
            $lastHighestCount = $highestCount
            $highestCount = $_.Value
            $mode = $_.Key
        }
    }

    if( $highestCount -eq $lastHighestCount -or ( $highestCount -eq 1 -and $modes.Count -gt 1 ) )
    {
        $mode = 0 ## no single most common boot time
    }

    [int]$median = $sorted[$sorted.Count / 2].BootTime
    [int]$mean = [math]::Round( $totalTime / $events.Count )

    [string]$summary = "Got $($events.Count) boot events for $($targets.Count) different target devices since $(Get-Date $startDate -Format G)`nFastest $fastest s slowest $slowest s mean $mean s median $median s mode $mode s ($highestCount instances)"
    
    Write-Output $summary

    $events | Format-Table -AutoSize
}
else
{
    Write-Output "Found no instances of event with id $eventId in the $eventLog event log from provider $providerName since $(Get-Date $startDate -Format G)"
}

