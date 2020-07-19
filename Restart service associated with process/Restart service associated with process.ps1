<#
    Given a process name and optional command line, figure out if it is a service and restart it as notification from ControlUp console will be after process exits
    Path passed must be a full path otherwise if there are two services of the same name we cannot differentiate
    Will not start non-auto start services since they may not be running

    @guyrleech 2019
#>

if( ! $args -or ! $args.Count -or ! $args[0] )
{
    Throw 'Must specify an executable and optional command line parameters'
}

$VerbosePreference = 'SilentlyContinue'

[string]$processName = $null
[string]$arguments = $null

if( $args[0] -match '^"([^"]*)"\s*?(.*)?$' -or $args[0] -match '^([^\s]*)\s*?(.*)?$' )
{
    $processName = $Matches[1]
    $arguments = $Matches[2].Trim()
}
else
{
    Throw "Failed to parse command line: $($args[0])"
}

Write-Verbose "Executable `"$processName`" arguments `"$arguments`""

[array]$matchingServices = @( Get-WmiObject -Class win32_service | ForEach-Object `
{
    ## PathName will either have quotes if it has spaces or not if not and then followed by any arguments
    Write-Verbose $_.PathName
    if( $_.PathName -match '^"([^"]*)"\s*?(.*)?$' -or $_.PathName -match '^([^\s]*)\s*?(.*)?$' )
    {
        if( $Matches[1] -eq $processName )
        {
            if( ( [string]::IsNullOrEmpty( $arguments ) -and ! $Matches[2] ) -or ( ! [string]::IsNullOrEmpty( $arguments ) -and $arguments -eq $Matches[2].Trim() ) )
            {
                Add-Member -InputObject $_ -MemberType NoteProperty -Name 'Executable' -Value $Matches[1]
                Add-Member -InputObject $_ -MemberType NoteProperty -Name 'Arguments' -Value $Matches[2]
                $_
            }
        }
    }
    elseif( $_.PathName )
    {
        Write-Warning "$($_.PathName) did not match regex"
    }
})

if( $matchingServices -and $matchingServices.Count )
{
    ## Now find how many aren't running so we can find the one to restart
    [array]$notRunning = @( $matchingServices | Where-Object { ! $_.Started } )

    Write-Verbose "Found $($matchingServices.Count) matching services of which $($notRunning.Count) are not running"

    if( $notRunning.Count -eq 1 )
    {
        if( $notRunning[0].StartMode -eq 'Auto' )
        {
            ## Just one not running so we can start it
            $startError= $null
            Write-Output "Starting service $($notRunning[0].Name) ($($notRunning[0].DisplayName)) ..."
            $started = Start-Service -Name $notRunning[0].Name -PassThru -ErrorVariable startError
            if( $? )
            {
                if( $started -and $started.Status -eq 'Running' )
                {
                    Write-Output "Started ok"
                }
                else
                {
                    Write-Error "Failed to start - status is $($started.Status)"
                }
            }
            ## else will have already output an error
        }
        else
        {
            Write-Warning "Service $($notRunning[0].Name) ($($notRunning[0].DisplayName)) is of start type $($notRunning[0].StartMode) so may not need to run"
        }
    }
    elseif( $notRunning.Count -eq 0 )
    {
        Write-Error "Found no non-running instances of a service with executable `"$processName`" and arguments `"$arguments`" so cannot restart any services.`nRunning services are: $(($matchingServices|select -expandproperty Name) -join ' , ')"
    }
    else
    {
        Write-Error "Found $($notRunning.Count) non-running services with executable `"$processName`" and arguments `"$arguments`" so cannot restart due to ambiguity`nNon-running services are: $(($notRunning|select -expandproperty Name) -join ' , ')"
    }  
}
else
{
    Write-Warning "Failed to find a service with executable `"$processName`" and arguments `"$arguments`""
}
