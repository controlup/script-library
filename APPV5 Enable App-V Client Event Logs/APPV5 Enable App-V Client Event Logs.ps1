#Requires -Version 2.0

<#
    .SYNOPSIS
    This script will enable all the App-V client event logs

    .DESCRIPTION
    This script will enable all the App-V client event logs

    .LINK
     http://virtualengine.co.uk

    AUTHOR: Nathan Sperry, Virtual Engine
    LASTEDIT: 05/06/2015
    WEBSITE: http://www.virtualengine.co.uk
    KEYWORDS: App-V,App-V 5,.APPV,VirtualEngine,AppV5
#>

$ErrorActionPreference = 'Stop'

try
{
    $appvlogs = Get-WinEvent -ListLog *AppV* -force | Where-Object {$_.IsEnabled -eq $false}

    if ($appvlogs.Count -gt 0)
    {

     foreach ($logitem in $appvlogs)
        {
             Write-Output ('Log enabled: ' + $logitem.LogName)
             $logitem.IsEnabled = $true
             $logitem.SaveChanges()
        }
        Write-Output ('Number of logs enabled: ' + $appvlogs.Count)
    }
    else
    {
        Write-Output ('Event logs already enabled')
    }
}
Catch
{
    $ErrorMessage = $_.Exception.Message
    Write-Output $ErrorMessage
}
