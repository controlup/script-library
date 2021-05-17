#requires -version 3
<#
    Add or remove users from a specified local group

    @guyrleech (c) 2018

    Modification history:
#>

[bool]$remove = $false
[string[]]$users = $args[0] -split ','
[string]$group = $null
[int]$errors = 0

$when = $null

if( $args.Count -ge 2 -and $args[1] )
{
    $group = $args[1]
}
else
{
    Throw "Must specify a comma separated list of users and the local group name"
}

if( $args.Count -ge 3 -and $args[2] -and $args[2] -eq 'True' )
{
    $remove = $true
}

if( $args.Count -ge 4 -and $args[3] )
{   
    $result = New-Object DateTime
    if( [datetime]::TryParse( $args[3] , [ref]$result ) )
    {
        $when = $result
    }
    else
    { 
        [string]$last = $args[3]
        [long]$multiplier = 0
        switch( $last[-1] )
        {
            "s" { $multiplier = 1 }
            "m" { $multiplier = 60 }
            "h" { $multiplier = 3600 }
            "d" { $multiplier = 86400 }
            "w" { $multiplier = 86400 * 7 }
            "y" { $multiplier = 86400 * 365 }
            default { Throw "Unknown multiplier `"$($last[-1])`"" }
        }
        if( $last.Length -le 1 )
        {
            $when = (Get-Date).AddSeconds( $multiplier )
        }
        else
        {
            $when = (Get-Date).AddSeconds( ( ( $last.Substring( 0 ,$last.Length - 1 ) -as [long] ) * $multiplier ) )
        }
    }
}

if( $when )
{
    ## Create a scheduled task to run this script at the given time
    ## Write this script to local file system   
    [string]$originalScript =  ( & { $myInvocation.ScriptName } ) 
    [string]$copiedScript = Join-Path (Split-Path $originalScript -Parent) ('Async_' + $(Split-Path $originalScript -Leaf))
    Copy-Item -Path $originalScript -Destination $copiedScript -Force
    if( ! ( Test-Path $copiedScript -PathType Leaf ) )
    {
        Throw "Failed to make a copy of the SBA script to use in a scheduled task"
    }
    Import-Module ScheduledTasks
    if( ! $? )
    {
        Throw "The scheduled tasks module is only available on Windows 8 and Server 2012 and higher"
    }
    [string]$taskname = ( "{0} users {1} {2} local group {3}" -f $(if ( $remove ) { 'remove' } else {'add'} ), $args[0] , $(if ( $remove ) { 'from' } else {'to'} ), $group )
    [string]$poshArguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -Nologo -Noninteractive -File `"$copiedScript`" `"$($args[0])`" `"$group`" $remove"
    $trigger = New-ScheduledTaskTrigger -At $when -Once -ErrorAction Stop
    $action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument $poshArguments -ErrorAction Stop
    $principal = New-ScheduledTaskPrincipal -UserID 'NT AUTHORITY\SYSTEM' -LogonType ServiceAccount -RunLevel Highest -ErrorAction Stop
    [string]$description = "Task created by ControlUp SBA at $(Get-Date -Format G)"

    $taskError = $null
    $task = Register-ScheduledTask -Action $action -Trigger $trigger -Description $description -TaskName ($taskName.Substring(0,1).ToUpper() + $taskname.Substring(1)) -Force -Principal $principal -ErrorVariable taskError
    if( ! $task -or $taskError )
    {
        Exit 4
    }
    elseif( $task.State -ne 'Ready' )
    {
        Write-Warning "Scheduled task `"$taskName`" created but task state is $($task.State) rather than `"Ready`""
    }
    else
    {
        Write-Output "Successfuly created scheduled task to $taskname at $(Get-Date $when -Format G)"
    }
}
else ## do it now
{
    $rootDSE = [ADSI]"LDAP://RootDSE"
    [string]$domain = (($rootDSE.rootDomainNamingContext -split ',')[0] -split 'DC=')[-1]

    [array]$adUsers = 
        @( ForEach( $user in $users )
        {
            if( $user -match '^[\- _a-z0-9]' )
            {
                [string]$domainName,[string]$userName = $user.Trim() -split '\\'

                if( [string]::IsNullOrEmpty( $userName ) )
                {
                    $userName = $domainName
                    $domainName = $domain
                }
                $thisUser = [ADSI]"WinNT://$domainName/$userName,user"
                if( ! $thisUser.Path )
                {
                    Write-Error "Failed to find user $domainName\$userName"
                    $missingUsers++
                }
                else
                {
                    $thisUser
                }
            }
        })

    if( $missingUsers )
    {
        Write-Error "Failed to find $missingUsers user(s) - aborting"
        Exit 2
    }

    if( ! $adUsers.Count )
    {
        Write-Error "No valid users specified"
        Exit 3
    }

    [string]$verb = $null
    [string]$preposition = $null

    if( $remove )
    {
        $verb = 'Removing'
        $preposition = 'from'
    }
    else
    {
        $verb = 'Adding'
        $preposition = 'to'
    }

    $localGroup = [ADSI]"WinNT://$env:COMPUTERNAME/$group,group"

    ForEach( $adUser in $adUsers )
    {
        [string]$operation = "$verb $((($aduser.Path -split ':')[1] -split ',')[0] -replace '//' , '' -replace '/' , '\') $preposition `"$group`""
        "$operation ..."
        try
        {
            if( $remove )
            {
                $localGroup.Remove( $adUser.Path )
            }
            else
            {
                $localGroup.Add( $adUser.Path )
            }
        }
        catch
        {
            Write-Error "Error $($operation.Substring(0,1).ToLower() + $operation.Substring(1))) - $($_.Exception.Message)"
            $errors++
        }
    }
    "Finished with $errors errors"
}

Exit $errors

