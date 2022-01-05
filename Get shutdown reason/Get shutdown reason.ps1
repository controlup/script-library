<#
.SYNOPSIS

Find and display the event for the most recent shutdown/reboot

.PARAMETER hours

How far back to look in the system event log. If not specified will search the entire event log.

.CONTEXT machine

.NOTES

Modification History:

  2021/06/02  Dennis Geerlings   Initial Version
  2021/06/17  Guy Leech          Details when event not found, add -hours parameter, show if local or remote machine instigated it
#>

[CmdletBinding()]

Param
(
    [double]$hours
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 250
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

[hashtable]$eventFilter =  @{ logname = 'system' ; id = 1074 }

if( $PSBoundParameters[ 'hours' ] ) ## this will ignore 0 thus specify that to search back to start of event log
{
    $eventFilter.Add( 'starttime' , (Get-Date).AddMinutes( -( $hours * 60 ) ) )
}

if( $event = Get-WinEvent -FilterHashtable $eventFilter -MaxEvents 1 -ErrorAction SilentlyContinue )
{
    ## process instigating reboot has (machinename) appended so isolate that in case it isn't local machine doing it
    [string]$process = $null
    [string]$rebootedFrom = $null
    if( $event.Properties[0].Value -match '^(.*)\s*\(([^\)]*)\)$' )
    {
        $process = $matches[1]
        $rebootedFrom = $matches[2]
        [ipaddress]$rebootedByIPAddress = $rebootedFrom -as [ipaddress]
        if( $rebootedByIPAddress -and ( $resolved = [System.Net.Dns]::GetHostByAddress( $rebootedFrom ) ) -and $resolved.HostName)
        {
            $rebootedFrom = $resolved.HostName
        }
    }
    else
    {
        $process = $event.Properties[0].Value
    }
    [hashtable]$result = @{
        Date = $event.TimeCreated
        Reason = $event.Properties[2].Value
        'Reason Code' = $event.Properties[3].Value
        Action = $event.Properties[4].Value
        Comment = $event.Properties[5].Value
        User = $event.Properties[6].Value
        Process = $process
        'Rebooted From' = $(if( ! $rebootedFrom -or $rebootedFrom -eq $env:COMPUTERNAME ) { 'Local' } else { $rebootedFrom })
    }
    New-Object -TypeName psobject -Property $result | Select-Object -Property @{n='Process';e={$_.Process}},* -ExcludeProperty Process | Format-Table -AutoSize
}
else
{
    $lastBootTime = Get-WmiObject -Class Win32_operatingsystem | Select-Object -ExpandProperty LastBootUpTime
    $oldestEvent = Get-WinEvent -LogName System -MaxEvents 1 -Oldest
    [string]$message = "No shutdown event (id 1074) found "
    if( $eventFilter[ 'StartTime' ] )
    {
        $message += "since $(Get-Date -Date $eventFilter.StartTime -Format G) "
    }
    if( $oldestEvent )
    {
        if( $oldestEvent.TaskDisplayName -eq 'Log clear' )
        {
            $message += "- log was cleared at $(Get-Date -Date $oldestEvent.TimeCreated -Format G) by $(([System.Security.Principal.SecurityIdentifier]($oldestEvent.UserId)).Translate([System.Security.Principal.NTAccount]).Value)"
        }
        else
        {
            $message += "- oldest system event log entry is $(Get-Date -Date $($oldestEvent.TimeCreated) -Format G)"
        }
    }
    $message += ", last boot time $(Get-Date -Date ([Management.ManagementDateTimeConverter]::ToDateTime( $lastBootTime )) -Format G)"
    Write-Warning -Message $message
}
