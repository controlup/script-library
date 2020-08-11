
<#
    Look for account lock out events on all selected domain controllers

    @guyrleech 24/07/2020
#>

[CmdletBinding()]


Param
(
    [double]$daysAgo = 1 ,
    [string]$username
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })

[int]$outputWidth = 400
[string]$message = $null

# Altering the size of the PS Buffer
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ($WideDimensions = $PSWindow.BufferSize) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

[datetime]$startDate = (Get-Date).AddDays( -$daysAgo )

## Invoke command will run some in parallel and Get-WinEvent will only take a single machine
[array]$events = @( Get-WinEvent -FilterHashtable @{ LogName = 'Security' ; Id = 4740 ; StartTime = $startDate } -ErrorAction SilentlyContinue )

if( $events -and $events.Count )
{
    Write-Verbose -Message "Found $($events.Count) events"

    <#
        Properties array:

        TargetUserName johndoe 
        TargetDomainName GLS16MCS01 
        TargetSid S-1-5-21-1721611859-3364803896-2099701507-2124 
        SubjectUserSid S-1-5-18 
        SubjectUserName GRL-DC03$ 
        SubjectDomainName GUYRLEECH 
        SubjectLogonId 0x3e7 
    #>
    [array]$filtered = @( $events | Where-Object { [string]::IsNullOrEmpty( $username ) -or $_.properties[0].value -eq $username } | Select-Object -Property TimeCreated,@{n='User name';e={$_.Properties[0].value}},@{n='Computer name';e={$_.Properties[1].value}} | Sort-Object -Property TimeCreated )
    if( $filtered -and $filtered.Count )
    {
        $filtered | Format-Table -AutoSize
    }
    else
    {
        $message = "Found no lock out events in last $daysAgo days"
        if( ! [string]::IsNullOrEmpty( $username ) )
        {
            $message += " for user $username"
        }
    }
}
else
{
    $message = "No lock out events found in last $daysAgo days"
}

if( ! [string]::IsNullOrEmpty( $message ) )
{
    if( $oldestEvent = Get-WinEvent -Oldest -LogName Security -ErrorAction SilentlyContinue -MaxEvents 1 )
    {
        $message += ", oldest event in security event log is from $(Get-Date -Date $oldestEvent.TimeCreated -Format G)"
    }
    Write-Output -InputObject $message
}

