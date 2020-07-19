#requires -version 3
<#
    Find and delete the scheduled task that the Scheduled Reboot SBA created

    @guyrleech 2018
#>

[string]$taskName = 'Reboot scheduled from ControlUp console'

[bool]$foundTask = $false
[int]$exitCode = 0

Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue | ForEach-Object `
{
    ## The Date property is empty so we get creation time from the task description
    [string]$createdAt = $null
    if( $_.Description -match '\. Created at (.*)$' )
    {
        $createdAt = ", created at $($Matches[1])"
    }
    $_ | Unregister-ScheduledTask -Confirm:$false
    if( $? )
    {
        Write-Output "Successfuly deleted scheduled task `"$($_.TaskName)`"$createdAt"
    }
    else
    {
        Write-Error "Error deleting scheduled task `"$($_.TaskName)`"$createdAt"
        $exitCode = 1
    }
    $foundTask = $true
}

if( ! $foundTask )
{
    Write-Error "Failed to find a scheduled task called `"$taskName`""
}

Exit $exitCode
