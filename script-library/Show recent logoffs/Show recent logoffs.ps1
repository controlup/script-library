<#
    Show logoffs from machine, sorted on most recent, from User Profile Service operational event log entries

    @guyrleech 2018
#>


[datetime]$endDate = Get-Date
[int]$totalSessions = 0
[int]$hoursBack = 0
[string]$logName = 'Microsoft-Windows-User Profile Service/Operational'
[string]$user = $null
[int]$daysBefore = 7
[int]$outputWidth = 400

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

if( $args.Count )
{
    if( ! [string]::IsNullOrEmpty( $args[0] ) )
    {
        $hoursBack = $args[0]
    }
    if( $args.Count -ge 2 -and ! [string]::IsNullOrEmpty( $args[1] ) )
    {
        $user = $args[1]
    }
}

[datetime]$startDate = if( ! $hoursBack )
{
    $endDate.AddYears( -20 ) ## should be long enough!
}
else
{
    $endDate.AddHours( -$hoursBack )
}

## we look further back so we stand a better chance of finding the corresponding logon
[array]$events = @( Get-WinEvent -FilterHashtable @{ LogName = $logName ; id = 1,4 ; StartTime=(Get-Date $startDate).AddDays( -$daysBefore ) ; EndTime=$endDate } -ErrorAction SilentlyContinue -Oldest ) 

[array]$results = @( For( [int]$index = 0 ; $events -and $index -lt $events.Count ; $index++ )
{
    if( $events[ $index ].Id -eq 4 -and $events[ $index ].TimeCreated -ge $startDate ) ## logoff
    {
        $logoffEvent = $events[ $index ]
        [string]$userName = 
            try
            {
                ([Security.Principal.SecurityIdentifier]($logoffEvent.UserId)).Translate([Security.Principal.NTAccount]).Value
            }
            catch
            {
                ## Write-Error "Failed to get user name for SID $($logonEvent.UserId)"
                $logoffEvent.UserId
            }
        if( [string]::IsNullOrEmpty( $user ) -or $userName -match $user )
        {                
            [int]$sessionId = -1
            if( $logoffEvent.Message -match '(\d+)\.$' ) ## Finished processing user logoff notification on session 2. 
            {
                $sessionId = $Matches[ 1 ]
            }
            if( $sessionId -le 0 )
            {
                Write-Error "Failed to get valid session id from text `"$($logoffEvent.Message)`""
            }

            $logonEvent = $null
            if( $sessionId -gt 0 )
            {
                For( [int]$search = $index + 1 ; $search -ge 0 ; $search-- )
                {
                    if( $events[ $search ].Id -eq 1 -and $events[ $search ].UserId -eq $logoffEvent.UserId )
                    {
                        if( $events[ $search ].Message -match '(\d+)\.$' ) ## Recieved user logon notification on session 2.
                        {
                            [int]$loggedOnSessionId = $Matches[ 1 ]
                            if( $loggedOnSessionId -eq $sessionId )
                            {
                                $logonEvent = $events[ $search ]
                                break
                            }
                        }
                    }
                }
            }
            [pscustomobject][ordered]@{ 'UserName' = $userName ; 'Session Id' = $sessionId ; 'Logon Time' = if( $logonEvent ) { $logonEvent.TimeCreated } ; 'Logoff Time' = $logoffEvent.TimeCreated }
        }
    }
} )

[string]$output = if( $hoursBack )
{
    " in last $hoursBack hours (since $(Get-Date $startDate -Format G))"
}

if( ! [string]::IsNullOrEmpty( $user ) )
{
    $output += " for user $user"
}

if( $results -and $results.Count )
{
    Write-Output ( "$($results.Count) session logoffs found" + $output )

    $results | Sort 'Logoff Time' -Descending | Format-Table -AutoSize
}
else
{
    Write-Output "No logoffs found at all"
}

