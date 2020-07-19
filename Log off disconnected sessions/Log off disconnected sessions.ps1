<#
    1) Find sessions idle over a given period and disconnect/logoff

    or

    2) Logoff all disconnected sessions over a given idle period or all if idle period is 0

    @guyrleech 2018
#>

[bool]$logoffDisconnected = $true ## set to false to perform 1) or true to perform 2)

[int]$idleOverMinutes = $args[0]
[bool]$logoff = $false
[int]$totalSessions = 0
[string]$qualifier = $null

if( $args.Count -ge 2 -and $args[1] -and $args[1] -eq 'True' )
{
    $logoff = $true
}

[array]$idleSessions = @( (quser.exe | select -Skip 1).Trim() | ForEach-Object `
{
    $totalSessions++
    [string[]]$fields = $_ -split '\s+'

    ## Headings are:  USERNAME              SESSIONNAME        ID  STATE   IDLE TIME  LOGON TIME

    ## if logon time has AM/PM indicator as opposed to 24 hour then there will be one more field
    [int]$disconnectedFieldCount = if( $_ -match '\s[AP]M$' ) { 7 } else { 6 }
    [int]$fieldOffset = if( $fields.Count -eq $disconnectedFieldCount ) { 0 } else { 1 } ## so we can skip passed sessionname since missing if disconnected
    [int]$idleDurationMinutes = `
        if( $fields[ 3 + $fieldOffset ] -eq '.' )
        {
            if( $fieldOffset )
            {
                -1
            }
            else ## disconnected
            {
                0
            }
        }
        elseif( $fields[3 + $fieldOffset ] -match '^(\d+)\+(\d+):(\d+)$' ) ## 33+22:16
        {
            ( [int]$Matches[1] * 24 + [int]$Matches[ 2 ] ) * 60 + [int]$Matches[ 3 ]
        }
        elseif( $fields[3 + $fieldOffset ] -match '^(\d+):(\d+)$' ) ## 19:35
        {
            [int]$Matches[1] * 60 + [int]$Matches[ 2 ]
        }
        else
        {
            [int]$fields[3 + $fieldOffset ]
        }

        if( $logoffDisconnected -and $fieldOffset )
        {
            ## Don't add to list as isn't disconnected and we are in the mode when only looking for disconnected sessions
        }
        elseif( $idleDurationMinutes -ge 0 -and $idleDurationMinutes -ge $idleOverMinutes )
        {
            [pscustomobject][ordered]@{ 
                'Username' = $fields[0]
                'Session Id' = $fields[1 + $fieldOffset]
                'Logon Time' = if( $disconnectedFieldCount -eq 7 ) { ( '{0} {1} {2}' -f $fields[-3] , $fields[-2] , $fields[-1] ) } else { ( '{0} {1}' -f $fields[-2] , $fields[-1] ) }
                'Idle Minutes' = $idleDurationMinutes
                'Connected' = $fieldOffset
        }
    }
} )

if( $logoffDisconnected )
{
    $logoff = $true
    $qualifier = ' disconnected and'
}

"Found $($idleSessions.Count) sessions$qualifier idle over $idleOverMinutes minutes out of $totalSessions sessions total"

$idleSessions | Sort 'Idle Minutes' -Descending | Format-Table -AutoSize

$affectedUsers = New-Object System.Collections.ArrayList

ForEach( $session in $idleSessions )
{
    [string]$exe = $null
    if( $logoff )
    {
        $exe = 'logoff.exe'
    }
    elseif( $session.Connected ) ## disconnect
    {
        $exe = 'tsdiscon.exe'
    }
    if( $exe )
    {
        $process = Start-Process -FilePath $exe -ArgumentList "$($session.'Session Id')" -PassThru -Wait -WindowStyle Hidden
        if( ! $process -or $process.ExitCode )
        {
            Write-Warning "Failed to run $exe successfully for session id $($session.'Session Id') user $($session.UserName)"
        }
        else
        {
            [void]$affectedUsers.Add( $session.Username )
        }
    }
}

Write-Output ( "{0} {1} sessions:" -f $(if( $logoff ) { 'Logged off' } else { 'Disconnected' }) , $affectedUsers.Count )

$affectedUsers | Sort -Unique | ForEach-Object { "`t$_" }


