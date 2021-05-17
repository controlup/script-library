#Logman capture for Network Diagnostics
#requires -version 3.0

## Last modified 28/09/18 1657 BST @guyrleech

## Change these to repurpose the script - currently only supports one or other of sent or receieved being true, not both
[bool]$received = $false
[bool]$sent = $true
[bool]$perSession = $true
[bool]$perComputer = $false

[int]$top = 10
[int]$caplength = 30

if( $received -and $sent )
{
    Write-Warning 'Cannot run with both $received and $sent set to true'
}

if( $perSession -and $perComputer )
{
    Write-Warning 'Cannot run with both $perSession and $perCOmputer set to true'
}

if( ! $received -and ! $sent )
{
    Write-Warning 'Cannot run with both $received and $sent set to false'
}

[int]$outputWidth = 400

$DebugPreference = "SilentlyContinue"

If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning “You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!”
    Break
}

Function Cleanup{
        $ret=start-process -FilePath "$env:windir\system32\logman.exe" -ArgumentList "stop wdc -ets" -wait -PassThru
        foreach($file in Get-ChildItem "$env:temp\output*.etl"){
            Remove-Item $file.FullName -Force -ea SilentlyContinue
        }
}

[int]$sessionId =-1
[int]$ppid = -1

if( $perSession )
{
    $sessionId = $args[0]
    $theSession = quser.exe $sessionId
    if( ! $theSession -or ! $theSession.Count )
    {
        Throw "Session id $sessionId not found"
    }
}
elseif( ! $perComputer )
{
    $ppid = $args[0]
    if( ! ( Get-Process -Id $ppid -ErrorAction SilentlyContinue ) )
    {
        Throw "Process id $ppid no longer exists"
    }
}

if( $perComputer )
{
    if ( $args.Count -ge 1 -and $args[0]){$caplength = $args[0]}
    if ( $args.Count -ge 2 -and $args[1]){$top = $args[1]}
}
else
{
    if ( $args.Count -ge 2 -and $args[1]){$caplength = $args[1]}
    if ( $args.Count -ge 3 -and $args[2]){$top = $args[2]}
}

#write-debug "Capturing Process id: $ppid, process name $(get-process -id $ppid).processname for length: $caplength"

$connectionlogs=@()
$break=$false
$breakreason=""
$tempfile="$env:temp\output%d.etl"

#check prereqs
if (!(test-path $env:windir\system32\logman.exe)){$break=true;$breakreason="logman.exe could not be found"}

if (!$break){

    cleanup

    $loop=0
    #loop for $caplength seconds
    #Start capture here with logman
    write-host "Capturing ETL logs for $caplength seconds" -nonewline
    $ret=start-process -FilePath "$env:windir\system32\logman.exe" -ArgumentList "start wdc -p Microsoft-Windows-Kernel-Network 0x10 0xff -bs 64 -nb 16 38 -max 16 -mode newfile -o ""$tempfile"" -ets" -Wait -PassThru
    Start-Sleep 1
    if ( ! $ret -or $ret.ExitCode -ne 0){write-warning "could not start log capture process! $($ret.exitcode)";exit}
    Do{
        write-host "." -NoNewline
        $loop++
        start-sleep 1
    } while($loop -lt $caplength)

    write-host "Capturing complete"
    $ret=start-process -FilePath "$env:windir\system32\logman.exe" -ArgumentList "stop wdc -ets" -wait -PassThru
    if (!($ret.ExitCode -eq 0)){write-warning "could not stop log capture process! $($ret.exitcode)";exit}
    $logs=@()

    [hashtable]$connectionTable = @{}
    [long]$lastKey = $null
    $netobject = $null
    [hashtable]$sessionProcesses = @{}

    [string]$xpath = '*[System['
    if( $sent )
    {
        $xpath += 'EventID=10 or EventID=42'
    }
    if( $received )
    {
        if( $sent )
        {
            $xpath += ' or '
        }
        $xpath += 'EventID=11 or EventID=43'
    }
    $xpath += ']'
    if( $ppid -gt 0 )
    {
        $xpath += " and EventData[Data[@Name='PID']='$ppid']" ## if session then make an "or" list
    }
    ## FilterXPath string has a maximum length of around 512 so we can't filter on all session processes
    elseif( $sessionId -ge 0 )
    {
        Get-Process | Where-Object { $_.SessionId -eq $sessionId } | ForEach-Object {
            $sessionProcesses.Add( $_.Id , $true )
        }
        if( ! $sessionProcesses -or ! $sessionProcesses.Count )
        {
            Throw "Found no processes for session $sessionId"
        }
    }
    ## else computer scope so no process filtering
    $xpath += ']'

    $startProcessing = [datetime]::Now
    #loop for all logs
    Write-host "Anaylysing captured traffic"
    @( foreach($file in Get-ChildItem "$env:temp\output*.etl"){
        Get-WinEvent -Path $file.fullname -Oldest -FilterXPath $xpath -ErrorAction SilentlyContinue
    } ) | ForEach-Object {

        if( $sessionId -lt 0 -or $sessionProcesses[ $_.Properties[0].Value -as [int] ] )
        {
            $log = $_
            $udp=$false
            $TCP=$false
            $send=$false
            $receive=$false
            #define type
            Switch ($log.id)
            {
                10 {$tcp=$true;$send=$true}
                42 {$udp=$true;$send=$true}
                11 {$tcp=$true;$receive=$true}
                43 {$UDP=$true;$receive=$true}
            }

            $size=$log.Properties[1].Value

            [string]$mainkey = $null
            if( $sent )
            {
                $saddr = $null
                $sport = $null
                $daddr = [long]$log.Properties[2].Value
                $dport=(($log.Properties[4].Value -band 0xff ) -shl 8 ) -bor ( ( $log.Properties[4].Value -band 0xff00 ) -shr 8)
                $mainkey = "$daddr$dport"
            }
            if( $received )
            {
                $daddr = $null
                $dport = $null 
                $saddr = [long]$log.Properties[2].Value 
                $sport=(($log.Properties[4].Value -band 0xff ) -shl 8 ) -bor ( ( $log.Properties[4].Value -band 0xff00 ) -shr 8)
                $mainkey = "$saddr$sport"
            }
        
            ## if we already have object from previous iteration and it's the same key then use that object rather than looking up again
            if( ! $netobject -or $lastKey -ne $mainkey )
            {
                $netobject = $connectionTable[ $mainkey ]
            }
            
            if (!$netobject){
                $netObject = [pscustomobject]@{            
                    Source           = $saddr
                    Destination      = $daddr                 
                    sourceport       = $sport             
                    destinationport  = $dport           
                    UDPReceived      = 0            
                    UDPSent          = 0            
                    TCPReceived      = 0            
                    TCPSent          = 0            
                    TCPTotal         = 0            
                    UDPTotal         = 0            
                    TotalSent        = 0            
                    TotalReceived    = 0
                    PIDs             = New-Object System.Collections.ArrayList
                }
                $connectionTable.Add( $mainkey , $netobject )
            }
      
            if( $netobject.PIDs -notcontains $log.Properties[0].Value ) ## shortish list so lookup not too bad
            {
                $null = $netobject.PIDs.Add( $log.Properties[0].Value )
            }
            if ($tcp){
                if ($send){
                    $netobject.tcpsent+=$size
                    $netobject.totalsent+=$size
                }
                elseif ($receive){
                    $netobject.tcpreceived+=$size
                    $netobject.totalreceived+=$size
                }
                $netobject.TCPtotal+=$size
            }
            elseif ($udp){
            if ($send){
                    $netobject.udpsent+=$size
                    $netobject.totalsent+=$size
                }
                elseif ($receive){
                    $netobject.udpreceived+=$size
                    $netobject.totalreceived+=$size
                }
                $netobject.udptotal+=$size                
            }
            $lastKey = $mainkey
        }
    }
}
Else{
    write-warning "Script execution has been halted: $breakreason"
}

Write-Debug ( "$(Get-Date) : Data processing took {0} seconds" -f ([datetime]::Now - $startProcessing).TotalSeconds )

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

[string]$heading = "`nShowing top $top network"
if( $received )
{
    $heading += ' sources for '
}
else
{
    $heading += ' destinations for '
}

[string]$scope = $null
if( $sessionId -ge 0 )
{
    $scope = "session $sessionId"
}
elseif( $ppid -gt 0 )
{
    $scope = "process $ppid ($(Get-Process -Id $ppid -ErrorAction SilentlyContinue|Select -ExpandProperty Name))"
}
else
{
    $scope = 'whole computer'
}
$heading += "$scope :"

$SummaryObject = [pscustomobject][ordered]@{ "Total Connections" = ($connectionTable.count) }

if( $connectionTable.count )
{
    if( $sent )
    {
        Add-Member -InputObject $SummaryObject -MemberType NoteProperty -Name "Total Sent (KB)" -Value ( [math]::Round(($connectionTable.GetEnumerator()|Select -ExpandProperty Value|Measure-Object -Property 'totalsent' -Sum).Sum /  1KB,1))
    }
    if( $received )
    {
        Add-Member -InputObject $SummaryObject -MemberType NoteProperty -Name "Total Received (KB)" -Value  ([math]::Round(($connectionTable.GetEnumerator()|Select -ExpandProperty Value|Measure-Object -Property 'totalreceived' -Sum).Sum /  1KB,1))
    }
}
else
{
    "No network traffic captured during $caplength seconds for $scope"
    return
}

$SummaryObject | Format-Table -AutoSize

$heading

[string[]]$rawSorter = $null 
$selector = $null

if( $received )
{
    $rawSorter = @( 'TotalReceived' )
    $selector = [System.Collections.ArrayList]@( 'Source' , @{n='Source DNS Name';e={[system.net.dns]::GetHostByAddress($_.source)|Select -ExpandProperty HostName}} , @{n='Source port';e={$_.SourcePort}} , @{n="Total Received (KB)";e={[math]::Round($_.TotalReceived / 1KB,1)}} )
}
if( $sent )
{
    $rawSorter = @( 'TotalSent' )
    $selector = [System.Collections.ArrayList]@( 'Destination' , @{n='Destination DNS Name';e={[system.net.dns]::GetHostByAddress($_.destination)|Select -ExpandProperty HostName}} , @{n='Destination port';e={$_.DestinationPort}} , @{n="Total Sent (KB)";e={[math]::Round($_.TotalSent / 1KB,1)}} )
}

if( $ppid -lt 0 )
{
    $null = $selector.Add( 'Process Name(s)' )
}

$connectionTable.GetEnumerator()|Select -ExpandProperty Value | Where-Object { $rawSorter } | Sort-Object -Property $rawSorter -descending | Select -First $top | ForEach-Object `
{
    ## reconstruct IP addresses
    if( $_.source )
    {
        $_.source = [ipaddress]$_.source
    }
    if( $_.destination )
    {
        $_.destination = [ipaddress]$_.destination
    }
    $result = $_
    if( $ppid -lt 0 ) ## computer or session context so have pids we need to convert to process names (and counts)
    {
        [hashtable]$processes = @{}
        $_.Pids | ForEach-Object `
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
        Add-Member -InputObject $result -MemberType NoteProperty -Name 'Process Name(s)' -Value (( $processes.GetEnumerator() | ForEach-Object { if( $_.Value -gt 1 ) { "$($_.Name)($($_.Value))" } else { "$($_.Name)" } } ) -join ',')
    }
    $result
} | select -Property $selector | Format-Table -AutoSize

cleanup

