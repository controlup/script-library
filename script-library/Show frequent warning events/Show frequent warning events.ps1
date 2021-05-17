#Requires -version 3.0

<#
    Find repeated event log entries in a given time window across all event logs and output the most frequent

    ControlUp SBA

    Guy Leech @guyrleech 2018

    Modification history:

    09/07/18   GL  Only display Level in results if more than one level was requested (e.g. Warning,Error,Critical)

    20/11/18   GL  Give message if no events found
#>

## Arguments
##   0  Minutes back from current time
##   1  Event log level(s)
##   2  Regex of event log names to exclude

[int]$ERROR_INVALID_PARAMETER = 87
[string]$multipleUsers = '<Multiple>' ## this is output when more than one user has generated the same event log id entry
[int]$outputWidth = 200 ## could expose this as a parameter

## see if we have been called via ControlUp or by ourself
if( ! $args.Count -or $args.Count -lt 2 )
{
    Write-Error "Incorrect number of arguments passed to script - was expecting minutes back , log levels and optional regex for excluding event log names"
    exit $ERROR_INVALID_PARAMETER
}

# Altering the size of the PS Buffer
if( $PSWindow = (Get-Host).UI.RawUI )
{
    if( $WideDimensions = $PSWindow.BufferSize ) 
    {
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions
    }
}

[int]$minutes = $args[0]
[string]$excludedLognames = $args[2]

[hashtable]$eventLevelConversions = @{
        'LogAlways' = 0 ;
        'Critical' = 1 ;
        'Error' = 2;
        'Warning' = 3 ;
        'Informational' = 4;
        'Verbose' = 5;
}

[string]$logname = '*' ## could expose this as a parameter
[datetime]$end = Get-Date
[datetime]$start = $end.AddMinutes( -$minutes )
[array]$results = @()
[int]$eventlogs = 0
[int[]]$eventLevels = @()

ForEach( $level in ( $args[1] -split ',' ) )
{
    $eventLevels += $eventLevelConversions[ $level ]
}

[hashtable]$events = @{}

Get-WinEvent -ListLog $logname -EA silentlycontinue | Where-Object { $_.RecordCount -gt 0  } |  ForEach-Object `
{ 
    if( [string]::IsNullOrEmpty( $excludedLognames ) -or $_.LogName -notmatch $excludedLogNames )
    {
        Write-Verbose "$($_.LogName) $($_.recordcount) $($_.lastwritetime)"
        [hashtable]$filters = @{'Logname'=$_.LogName;StartTime=$start;EndTime=$end }
        if( $eventLevels.Count )
        {
            $filters.Add( 'Level' , $eventLevels )
        }

        [bool]$first = $true

        Get-WinEvent -FilterHashtable $filters -EA SilentlyContinue | ForEach-Object `
        {
            $event = $_
            [string]$key = "{0}:{1}:{2}" -f $event.id , $event.LogName , $event.ProviderName
            $existing = $events[ $key ]
            if( $existing )
            {
                $existing.Count++
                if( $event.TimeCreated -lt $existing.First )
                {
                    $existing.First = $event.TimeCreated
                }
                if( $event.TimeCreated -gt $existing.Last )
                {
                    $existing.Last = $event.TimeCreated
                }
                ## if different user then change to multiple
                if( $event.UserId ) 
                {
                    try
                    {
                        [string]$thisUser = ([System.Security.Principal.SecurityIdentifier]($event.UserId)).Translate([System.Security.Principal.NTAccount]).Value
                        if( $existing.User )
                        {
                            if( $existing.User -ne $multipleUsers )
                            {
                                if( $existing.user -ne $thisUser )
                                {
                                    $existing.user = $multipleUsers
                                }
                            }
                            ## else already recorded it as a multipl euser event
                        }
                        else
                        {
                            $existing.User = $thisUser
                        }
                    }
                    catch {}
                }
            }
            else
            {
                $events.Add( $key , [pscustomobject]@{ 'Count' = [int]1 ; 'First' = $event.TimeCreated ; 'Last' = $event.TimeCreated ; 'Message' = $event.Message ; 'LogName' = $event.LogName ; 'Id' = $event.id ; 'Level' = $event.LevelDisplayName ;
                    'User' = $( if( $event.Userid ) {([System.Security.Principal.SecurityIdentifier]($event.UserId)).Translate([System.Security.Principal.NTAccount]).Value }); 'Task' = $event.TaskDisplayName } )
            }
            if( $first )
            {
                $eventlogs++
                $first = $false
            }
        }
    }
}

Write-Verbose "Found $($events.Count) $($args[1]) events in $eventlogs event logs on $env:COMPUTERNAME between $(Get-Date $start -Format G) and $(Get-Date $end -Format G)"

[string]$dateFormat = 'T'
if( $start.DayOfYear -ne $end.DayOfYear )
{
    $dateFormat = 'G' ## put date in too since covers more than one day
}

[array]$fields = @( 'Count',@{n='First';e={Get-Date $_.First -Format $dateFormat}},@{n='Last';e={Get-Date $_.Last -Format $dateFormat}},@{n='Log Name';e={($_.LogName -split '/')[0] -replace '^Microsoft-Windows-',''}} )
if( $eventLevels.Count -gt 1 )
{
    $fields += 'Level' ## only put the level in if there is more than one, e.g. error and warning
}
$fields += @( 'Id','User','Message','Task' )

if( $events -and $events.Count )
{
    ## Only show those with at least one repetition
    [array]$repeated = @( $events.GetEnumerator() | Select -ExpandProperty Value | Where-Object { $_.Count -gt 1 }  )
    "Found $($repeated.Count) $($args[1]) repeated events out of $($events.Count) events found in event logs between $(Get-Date $start -Format G) and $(Get-Date $end -Format G)"
    if( $repeated -and $repeated.Count )
    {
        $repeated |Sort Count -Descending | Select $fields | Format-Table -Wrap
    }
}
else
{
    "Found no $($args[1]) events in event logs between $(Get-Date $start -Format G) and $(Get-Date $end -Format G)"
}
