<#
    Get connected users by parsing quser.exe output,  send them a message a number of times and then logoff

    @GuyRLeech, 2018
#>

[string]$message = $args[0]
[int]$delayBeforeLogoff = -1
[int]$repeatInterval = -1

if( $args.Count -ge 2 -and $args[1] )
{
    $delayBeforeLogoff = $args[1]
}

if( $args.Count -ge 3 -and $args[2] )
{
    $repeatInterval = $args[2]
}

if( $repeatInterval -gt $delayBeforeLogoff )
{
    Write-Warning "Repeat interval of $repeatInterval is greater than $delayBeforeLogoff so only one message will be delivered"
}

[datetime]$startTime = [datetime]::Now
[datetime]$logoffTime = $startTime.AddMinutes( $delayBeforeLogoff ) ## of negative then we won't loop
$allSessions = New-Object System.Collections.ArrayList
[int]$messageDuration = (New-TimeSpan -End $logoffTime -Start $startTime).TotalSeconds + 60
[string]$pidFile = Join-Path -Path $env:temp -ChildPath "ControlUp.Logoffs.$pid"

## Record our pid to a file so so we can have a cancellation SBA, e.g. if the problem requiring the reboot gets fixed
$pid | Out-File -FilePath $pidFile -Force

do
{
    [int]$totalSessions = 0
    [int]$totalDisconnected = 0
    $allSessions.Clear()
    [array]$connectedUsers = @( (quser.exe | select -Skip 1).Trim() | ForEach-Object `
    {
        $totalSessions++
        [string[]]$fields = $_ -split '\s+'

        ## Headings are:  USERNAME              SESSIONNAME        ID  STATE   IDLE TIME  LOGON TIME

        ## if logon time has AM/PM indicator as opposed to 24 hour then there will be one more field
        [int]$disconnectedFieldCount = if( $_ -match '\s[AP]M$' ) { 7 } else { 6 }

        if( $fields -and $fields.Count -ne $disconnectedFieldCount ) ## for disconnected session there is no SESSIONNAME so field count is reduced by one (so is 7 for active sessions). Do not check for 'Disc' as may be different in non-English
        {
            $fields[0]
            [void]$allSessions.Add( $fields[2] )
        }
        else
        {
            $totalDisconnected++
            [void]$allSessions.Add( $fields[1] )
        }

    } | Sort -Unique )

    if( ! $connectedUsers -or ! $connectedUsers.Count )
    {
        [string]$message = $null
        if( $totalDisconnected )
        {
            $message = "No connected users found, $totalDisconnected were disconnected"
        }
        else
        {
            $message = 'No connected or disconnected users found'
        }
        Write-Warning $message
        return ## could break out of loop if just disconnected so they get logged off since we can't message them
    }

    Write-Output "$(Get-Date -Format G): messaging $($connectedUsers.Count) connected sessions ..."
    $msgProcess = Start-Process -FilePath 'msg.exe' -ArgumentList "* /TIME:$messageDuration $($message -replace '\\n' , "`n")" -PassThru -Wait -ErrorAction Stop -WindowStyle Hidden
    if( ! $msgProcess )
    {
        Throw $error[0]
    }
    Write-Output "$(Get-Date -Format G): messaged $($connectedUsers.Count) users"

    if( $repeatInterval -gt 0 -and $repeatInterval -lt $delayBeforeLogoff )
    {
        ## ensure we don't sleep past the logoff time
        [int]$sleepInterval = [math]::Min( [math]::Round( ( New-TimeSpan -End $logoffTime -Start ([datetime]::Now) ).TotalMinutes , 1 ) , $repeatInterval )
        Write-Output "Sleeping for $sleepInterval minutes before sending message again ..."
        Start-Sleep -Seconds ($sleepInterval * 60)
    }
    elseif( $delayBeforeLogoff -gt 0 )
    {
        [int]$sleepFor = ( New-TimeSpan -End $logoffTime -Start ([datetime]::Now) ).TotalSeconds
        Write-Output "Sleeping for $([math]::Round( $sleepFor / 60 , 1 )) minutes before logging off ..."
        Start-Sleep -Seconds $sleepFor
        break
    }
} while( [datetime]::Now -lt $logoffTime )

## Now logoff all users
if( $delayBeforeLogoff -gt 0 )
{
    [int]$loggedOff = 0

    ## get sessions again since may have changed
    (quser.exe | select -Skip 1).Trim() | ForEach-Object `
    {
        [int]$sessionId = -1
        [string[]]$fields = $_ -split '\s+'
        [int]$disconnectedFieldCount = if( $_ -match '\s[AP]M$' ) { 7 } else { 6 }

        if( $fields -and $fields.Count -ne $disconnectedFieldCount )
        {
            ## see if it's our own session in which case we don't log it off
            if( $fields[0] -notmatch '^>' )
            {
                $sessionId = $fields[2]
            }
        }
        else
        {
            $sessionId = $fields[1]
        }
    
        if( $sessionId -ge 0 )
        {
            Write-Output "$(Get-Date -Format G) : logging off session $sessionId ..."
            $logoffProcess = Start-Process -FilePath 'logoff.exe' -ArgumentList $sessionId -PassThru -Wait -WindowStyle Hidden
            if( $logoffProcess )
            {
                Write-Output "$(Get-Date -Format G) : logged off session $sessionId ..."
                $loggedOff++
            }
            else
            {
                Write-Warning "Failed to run logoff for session $sessionId"
            }
        }
    }
    Write-Output "Logged off $loggedOff users"
}

if( Test-Path -Path $pidFile -PathType Leaf -ErrorAction SilentlyContinue )
{
    Remove-Item -Path $pidFile -Force
}

