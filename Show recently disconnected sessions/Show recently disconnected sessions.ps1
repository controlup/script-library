<#
    Show disconnected sessions sorted on how recently disconnected by parsing quser.exe output

    Guy Leech, 2018

    Modification History

    20/11/18   GRL Added logon time
#>

[datetime]$timeNow = Get-Date
[int]$totalSessions = 0
[int]$hoursBack = 0
[int]$totalDisconnected = 0
[int]$outputWidth = 400

if( $args.Count -and ! [string]::IsNullOrEmpty( $args[0] ) )
{
    $hoursBack = $args[0]
}

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

[array]$disconnectedUsers = @( (quser.exe | select -Skip 1).Trim() | ForEach-Object `
{
    $totalSessions++
    [string[]]$fields = $_ -split '\s+'

    ## Headings are:  USERNAME              SESSIONNAME        ID  STATE   IDLE TIME  LOGON TIME

    ## if logon time has AM/PM indicator as opposed to 24 hour then there will be one more field
    [int]$disconnectedFieldCount = if( $_ -match '\s[AP]M$' ) { 7 } else { 6 }

    if( $fields -and $fields.Count -eq $disconnectedFieldCount ) ## for disconnected session there is no SESSIONNAME so field count is reduced by one (so is 7 for active sessions). Do not check for 'Disc' as may be different in non-English
    {
        $totalDisconnected++

        [int]$disconnectedDurationMinutes = `
            if( $fields[3] -match '^(\d+)\+(\d+):(\d+)$' ) ## 33+22:16
            {
                ( [int]$Matches[1] * 24 + [int]$Matches[ 2 ] ) * 60 + [int]$Matches[ 3 ]
            }
            elseif( $fields[3] -match '^(\d+):(\d+)$' ) ## 19:35
            {
                [int]$Matches[1] * 60 + [int]$Matches[ 2 ]
            }
            else
            {
                [int]$fields[3]
            }

        if( ! $hoursBack -or $disconnectedDurationMinutes -le $hoursBack * 60 )
        {
            [pscustomobject]@{ 
                'Username' = $fields[0]
                'Logon Time' = if( $disconnectedFieldCount -eq 7 ) { ( '{0} {1} {2}' -f $fields[-3] , $fields[-2] , $fields[-1] ) } else { ( '{0} {1}' -f $fields[-2] , $fields[-1] ) }
                'Disconnection Time' = Get-Date ( $timeNow.AddMinutes( -$disconnectedDurationMinutes ) ) -Format G
                'Disconnected Hours' = [math]::Round( $disconnectedDurationMinutes / 60 , 1 ) 
                'Disconnected Minutes' = $disconnectedDurationMinutes }
        }
    }
} )

[string]$output = if( $hoursBack )
{
    " in last $hoursBack hours ($totalDisconnected disconnected in total)"
}

if( $disconnectedUsers -and $disconnectedUsers.Count )
{
    Write-Output ( "$($disconnectedUsers.Count) out of $totalSessions sessions ($([math]::Round( ( $disconnectedUsers.Count / $totalSessions ) * 100 ))%) are disconnected" + $output )

    $disconnectedUsers | Sort 'Disconnected Minutes' | Format-Table -AutoSize
}
elseif( $totalSessions )
{
    Write-Output ( "No disconnected sessions found out of $totalSessions sessions" + $output )
}
else
{
    Write-Output "No sessions found at all"
}

