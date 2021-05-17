<#
.SYNOPSIS

ENable or disable a Citrix XenApp published application

.DETAILS


.PARAMETER publishedAppName

The full name or pattern of the published application to enable/disable

.PARAMETER operation

Whether to enable or disable the specified published application(s)

.PARAMETER terminate

Whether to terminate any sessions using the published app after waiting for the number of seconds specified via -secondsGrace and sending a message if -messageText is specified

.PARAMETER messageText

Send a message with the given text to all users of the published app

.PARAMETER messageTitle

Title for the message to be sent via the -messageText option

.PARAMETER secondsGrace

The number of seconds to wait before terminating the sessions if -terminate is specified

.CONTEXT

Computer

.NOTES

Uses Citrix Studio PowerShell snapins so must run on a Delivery Controller or machine where the PowerShell snapins are installed

https://twitter.com/guyrleech/status/1169943274741731329?s=20

.MODIFICATION_HISTORY:

@guyrleech 17/09/19

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage="Published Application Name")]
    [string]$publishedAppName ,
    [Parameter(Mandatory=$true,HelpMessage="Operation to perform - enable or disable")]
    [ValidateSet('Enable','Disable')]
    [string]$operation ,
    [Parameter(Mandatory=$false)]
    [ValidateSet('Yes','No')]
    [string]$terminate = 'No' ,
    [int]$secondsGrace = 60 ,
    [Parameter(Mandatory=$false,HelpMessage="Message text to send to users")]
    [string]$messageText ,
    [Parameter(Mandatory=$false,HelpMessage="Title of message to send to users")]
    [string]$messageTitle = 'Message from Administrator'
)

if( ! ( Get-PSSnapin -Name Citrix.Broker.Admin.* -ErrorAction SilentlyContinue ) )
{
    if( ! ( Add-PSSnapin -Name Citrix.Broker.Admin.* -ErrorAction Continue -PassThru ) )
    {
        Throw 'Unable to find the Citrix PowerShell snapins which are typically found on Delivery Controllers'
    }
}

[array]$apps = @( Get-BrokerApplication -Name $publishedAppName -ErrorAction Continue )

if( ! $apps -or ! $apps.Count )
{
    Throw "No published applications found for `"$publishedAppName`""
}

[bool]$disable = $operation -eq 'disable'
[int]$alreadyInDesiredState = $apps | Where-Object { $_.Enabled -eq ! $disable } | Measure-Object | Select-Object -ExpandProperty Count

## Probably only one app so more efficient to filter on each app than get all sessions then filter for each app
[array]$sessions = @( ForEach( $app in $apps )
{
    Get-BrokerSession -ApplicationUid $app.Uid
})

Write-Output -InputObject "$($apps.Count) apps found for `"$publishedAppName`" with $alreadyInDesiredState already $($operation)d with $($sessions.Count) sessions"

if( $sessions.Count -and $PSBoundParameters[ 'messageText' ] )
{
    $sessions | Send-BrokerSessionMessage -MessageStyle Exclamation -Title $messageTitle -Text $messageText
}

[datetime]$timeStarted = Get-Date

if( $alreadyInDesiredState -lt $apps.Count )
{
    [array]$appsAfter = @( $apps | Set-BrokerApplication -Enabled:(!$disable) -PassThru )

    if( ! $appsAfter -or ! $appsAfter.Count )
    {
        Throw "Failed to get results back from changing published application state"
    }

    [string[]]$failed = @( ForEach( $appAfter in $appsAfter )
    {
        if( $appAfter.Enabled -ne ! $disable )
        {
            $appAfter.Name
        }
    })

    if( $failed -and $failed.Count )
    {
        Write-Warning -Message "Failed to $operation $($failed.Count) apps:"
        $failed | Write-Warning
    }
    else
    {
        if( $appsAfter.Count -ne $apps.Count )
        {
            Write-Warning -Message "Requested to $operation $($apps.Count) apps but $($appsAfter.Count) were returned from the $operation operation"
        }
        Write-Output -InputObject "All $($apps.Count) apps $($operation)d ok"
    }
}
else
{
    Write-Warning -Message "No applications have therefore been $($operation)d"
}

if( $sessions.Count -and $terminate -eq 'Yes' )
{
    [int]$secondsToWait = $secondsGrace - ((Get-Date) - $timeStarted).TotalSeconds
    if( $secondsToWait -gt 0 )
    {
        Write-Output -InputObject "Sleeping for grace period of $secondsToWait seconds before terminating apps for $($sessions.Count) sessions"
        Start-Sleep -Seconds $secondsToWait
    }
    $sessions | Stop-BrokerSession
}
