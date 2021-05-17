#requires -version 3.0
<#

    Restart computer at a specified number of hours and/or minutes in the future by creating a scheduled task that runs once and executes shutdown.exe

    @guyrleech 2018
#>

[int]$minutes = -1
[bool]$forceReboot = $false
[string]$taskName = "Reboot scheduled from ControlUp console" 
[string]$reason = $null

## arg 0 - [hh:]mm to reboot
## arg 1 - force
## arg 2 - reason for reboot (optional)

if( $args.Count -ge 2 )
{
    if( $args[0] -match '^(\d{1,2}:)?(\d{1,2})$' )
    {
        if( ! [int]::TryParse( $matches[2] , [ref]$minutes ) )
        {
            Write-Error "Bad format for minutes part of $($args[0])"
            Exit 1
        }
        if( $matches[ 1 ] )
        {
            [int]$hours = 0
            if( ! [int]::TryParse( ($matches[1] -replace ':' , '') , [ref]$hours ) )
            {
                Write-Error "Bad format for hours part of $($args[0])"
                Exit 1
            }
            $minutes += $hours * 60
        }
        if( $minutes -le 0 )
        {
            Write-Error "Immediate reboot is not supported - please use the Power Management menu instead"
            Exit 1
        }
    }
    else
    {
        Write-Error "Incorrectly specified reboot time `"$($args[0])`" - must be [hh]:mm"
        Exit 2
    }
    $forceReboot = ( $args[1] -and $args[1] -match 'true' )
    if( $args -ge 3 -and $args[2] )
    {
        $reason = $args[2]
    }
}
else
{
    Write-Error 'Unexpected number of parameters - expecting reboot time, force and optional reboot reason'
    Exit 3
}

if( $minutes -gt 0 )
{
    [datetime]$rebootTime = (Get-Date).AddMinutes( $minutes )
    Import-Module ScheduledTasks
    ## Can't just call shutdown.exe as /f won't not reboot if users are logged on
    [string]$poshArguments = '-ExecutionPolicy Bypass -WindowStyle Hidden -Nologo -Noninteractive -Command "'
    [string]$arguments = '/f /r'
    if( $reason )
    {
        $arguments += " /c '$reason'" ## should we use /e ?
    }
    if( $forceReboot )
    {
        $poshArguments += "shutdown.exe $arguments"
    }
    else
    {
        $poshArguments += "if( ! (quser.exe).Length ) { shutdown.exe $arguments } else { Write-Warning 'Not rebooting as users are logged on' }"
    }
    $poshArguments += '"'
    $trigger = New-ScheduledTaskTrigger -At $rebootTime -Once -ErrorAction Stop
    $action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument $poshArguments -ErrorAction Stop
    $principal = New-ScheduledTaskPrincipal -UserID 'NT AUTHORITY\SYSTEM' -LogonType ServiceAccount -RunLevel Highest -ErrorAction Stop
    [string]$description = "Reboot scheduled from the ControlUp console by invoking a Script Based Action (SBA). Created at $(Get-Date -Format G)"
    $task = Register-ScheduledTask -Action $action -Trigger $trigger -Description $description -ErrorAction Stop -TaskName $taskName -Force -Principal $principal
    if( $task )
    {
        Write-Output "Successfuly created scheduled task to reboot system at $(Get-Date $rebootTime -Format G)"
    }
    elseif( $task.State -ne 'Ready' )
    {
        Write-Warning "Scheduled task `"$taskName`" created but task state is $($task.State) rather than `"Ready`""
    }
    else
    {
        Exit 4
    }
}

