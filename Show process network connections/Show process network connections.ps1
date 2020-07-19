<#
    ControlUp SBA so summarise netstat TCP data for a given process or all processes

    Cannot use Get-NetTCPConnection as this has to run on 2008R2 which does not have it in PowerShell v3.0

    @guyrleech 2018

    Modification History:

    19/09/18   GRL    Made name and RIPE resolution optional
                      Added port number

    25/09/18   GRL    Column renaming and re-ordering
                      IPv6 addresses processed
#>

## change this depending on whether a session or process context action
[bool]$sessionMode = $false

[int]$top = 10
[int]$processId = -1
[int]$sessionId = -1
[string]$state = 'ESTABLISHED'
[hashtable]$connections = @{}
[hashtable]$otherConnections = @{}
[int]$connectionCount = 0
[int]$outputWidth = 400
[bool]$resolveIPs = $false
[bool]$resolveViaRIPE = $false

## If we are passed 4 arguments then is pid and count otherwise just count as we are a computer action
if( $args.Count -ge 4 )
{
    if( $sessionMode )
    {
        $sessionId = $args[0]
    }
    else
    {
        $processId = $args[0]
    }
    $top = $args[1]
    $resolveIPs = $args[2] -eq "true"
    $resolveViaRIPE = $args[3] -eq "true"
}
elseif( $args.Count -ge 3 )
{
    $top = $args[0]
    $resolveIPs = $args[1] -eq "true"
    $resolveViaRIPE = $args[2] -eq "true"
}
else
{
    Throw "Unexpected number of arguments passed ($($args.Count))"
}

$process = $null
if( $processId -ge 0 )
{
    $process = Get-Process -Id $processId

    if( ! $process )
    {
        Throw "Process with PID $processId does not exist"
    }
}

[hashtable]$processtoSession = @{}

if( $sessionId -ge 0 )
{
    Get-Process | Where-Object { $_.SessionId -eq $sessionId } | ForEach-Object `
    {
        $processtoSession.Add( [int]$_.Id , $true )
    }
}
    
## UDP does not have destinations so filter out
netstat -ano|select -skip 4 | ForEach-Object { $_ -replace '^\s+' , '' -replace '\s+' , ' '} | ConvertFrom-Csv -Deli ' ' -Head @('Protocol','LocalAddress','ForeignAddress','State','PID') `
    | Where-Object { $_.Protocol -notmatch 'UDP' -and ( $_.PID -eq $processId -or ( $sessionId -ge 0 -and $_.PID -and $processtoSession[ [int]$_.PID ] ) -or ( $sessionId -lt 0 -and $processId -lt 0 ) ) } | ForEach-Object `
{
    if( $_.State -match $state )
    {
        $connectionCount++
        ##[string]$address,[string]$port = $_.ForeignAddress -split ':'
        [string]$key = $_.ForeignAddress ## $_.protocol + ' ' + $address
        $connection = $connections[ $key ]
        if( $connection )
        {
            $connection.Count++
            try
            {
                $connection.Pids.Add( $_.PID , $true )
            }
            catch {} ## in case there already
        }
        else
        {
            $connections.Add( $key , [pscustomobject]@{ 'Count' = [int]1 ; Pids = @{ $_.PID = $true } } )
        }
    }
    else
    {
        ## Keep a count of connections in other states
        $otherConnection = $otherConnections[ $_.state ]
        if( $otherConnection )
        {
            $otherConnections.Set_Item( $_.state , [int]$otherConnection + 1 )
        }
        else
        {
            $otherConnections.Add( $_.state , [int]1 )
        }
    }
}

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

if( ! $connections -or ! $connections.Count )
{
    [string]$status = "No netstat output matching state `"$state`" found"
    if( $processId -ge 0 )
    {
        $status += " for process id $processId ($($process.Name))"
    }
    elseif( $sessionId -ge 0 )
    {
        $status += " for session id $sessionId"
    }
    Write-Warning $status
}
else
{
    $properties =[System.Collections.ArrayList] @( 'Destination Address','Port','Count' )
    if( $processId -lt 0 )
    {
        $properties.Insert( 0 , 'Process Name' )
    }
    [string]$heading = $null
    if( $top -ge $connectionCount )
    {
        $heading = 'All'
    }
    else
    {
        $heading = "Top $([math]::Min( $top , $connectionCount)) of"
    }
    $heading += " $connectionCount TCP connections matching state `"$state`"" 
    if( $processId -ge 0 )
    {
        $heading += " for process $($process.Name) ($processId)"
    }
    elseif( $sessionId -ge 0 )
    {
        $heading += " for session $sessionId"
    }

    $heading += ' are:'
    Write-Output $heading

    $connections.GetEnumerator()| ForEach-Object `
    {
        ##[string]$protocol,[string]$address = $_.Name -split '\s'
        [string]$address = $null
        [string]$port = $null 
        [bool]$ipv6 = $false
        [string[]]$splitAddress = $_.Name -split ':'

        if( $splitAddress -and $splitAddress.Count -gt 2 )
        {
            [int]$splitAt = $_.Name.ToString().LastIndexOf(':')
            $address = $_.Name.ToString().SubString( 0 , $splitAt ) -replace '^\[' , '' -replace '\]$' , ''
            $port = $_.Name.ToString().SubString( $splitAt + 1 )
            $ipv6 = $true
        }
        else
        {
            $address = $splitAddress[0]
            $port = $splitAddress[1]
        }

        ## if address is still IP address, we'll look it up via DNS or RIPE REST API if requested via parameters

        if( $resolveIPs -or $resolveViaRIPE )
        {
            try
            {
                $resolved = $null
                [ipaddress]$ipAddress = $address ## exception thrown if not valid IP address
                if( $ipAddress -and ( $ipAddress.Address -or $ipAddress.AddressFamily -eq 'InterNetworkV6' ) ) ## filter out 0.0.0.0
                {
                    if( $resolveIPs )
                    {
                        try
                        {
                            $resolved = [System.Net.Dns]::gethostentry( $ipAddress.IPAddressToString )
                        }
                        catch
                        {
                            $resolved = $null
                        }
                        if( $resolved )
                        {
                            $address = $resolved.HostName
                        }
                    }
                    if( ! $resolved -and $resolveViaRIPE )
                    {
                        [string]$registrant = $null
                        Invoke-WebRequest -Uri "https://rest.db.ripe.net/search.json?type=inetnum&type=organisation&query-string=$address&source=ripe&source=apnic-grs&source=arin-grs&source=afrinic-grs&source=lacnic-grs&source=jpirr-grs&source=radb-grs"|select -ExpandProperty Content|ConvertFrom-Json|select -ExpandProperty objects|select -ExpandProperty object|select -ExpandProperty attributes | ForEach-Object `
                        {
                            if( [string]::IsNullOrEmpty( $registrant ) )
                            {
                                ##Write-Warning (($_.attribute|?{$_.Name -ne 'Remarks'}|%{ "{0}={1}" -f $_.Name , $_.Value }) -join ',' )
                                $org = $_.attribute|Where-Object{ $_.name -eq 'org' }|select -ExpandProperty value
                                if( ! [string]::IsNullOrEmpty( $org ) )
                                {
                                    $registrant = $org
                                }
                            }
                        }    
                        if( ! [string]::IsNullOrEmpty( $registrant ) )
                        {
                            $address = "$address (IANA registrant is $registrant)"
                        }
                    }
                }
            }
            catch{ }
        }

        $result = [pscustomobject][ordered]@{
            'Destination Address' = $address
            'Port' = $port
            'Count' = $_.Value.Count
        }
        if( $processId -lt 0 ) ## computer context
        {
            [hashtable]$processes = @{}
            $_.Value.Pids.GetEnumerator()|Select -ExpandProperty Name|ForEach-Object `
            {
                [string]$thisProcess = Get-Process -Id $_ -ErrorAction SilentlyContinue | Select -ExpandProperty Name -ErrorAction SilentlyContinue
                if( [string]::IsNullOrEmpty( $thisProcess ) )
                {
                    $thisProcess = $_ ## can't get process, probably as terminated
                }
                $existingProcess = $processes[ $thisProcess ]
                if( $existingProcess )
                {
                    $processes.Set_Item( $thisProcess , [int]$existingProcess + 1 )
                }
                else
                {
                    $processes.Add( $thisProcess , [int]1 )
                }
            }
            Add-Member -InputObject $result -MemberType NoteProperty -Name 'Process Name' -Value (( $processes.GetEnumerator() | ForEach-Object { if( $_.Value -gt 1 ) { "$($_.Name)($($_.Value))" } else { "$($_.Name)" } } ) -join ',')
        }
        $result
    } | Sort Count -Descending | Select -First $top -Property $properties | Format-Table -AutoSize

    if( $otherConnections -and $otherConnections.Count )
    {
        Write-Output "Connections in other states:"
        $otherConnections.GetEnumerator() | Select @{n='State';e={$_.Name}},@{n='Count';e={$_.Value}} | sort State
    }
    else
    {
        Write-Output "No other connections in other states found"
    }
}

