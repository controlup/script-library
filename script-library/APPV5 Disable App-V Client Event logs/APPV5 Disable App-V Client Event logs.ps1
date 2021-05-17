#Requires -Version 2.0

<#
    .SYNOPSIS
    This script will disable all the App-V client event logs

    .DESCRIPTION
    This script will disable all the App-V client event logs

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
    $appvlogs = Get-WinEvent -ListLog *AppV* -force | Where-Object {$_.IsEnabled -eq $true}

    if ($appvlogs.Count -gt 0)
    {
    $i=0
        foreach ($logitem in $appvlogs)
        {
             ### Don't disable the common event logs
             if ($logitem.OwningProviderName -notlike 'Microsoft-AppV-Client')
             {
                 Write-Output ('Log disabled: ' + $logitem.LogName)
                 $logitem.IsEnabled = $false
                 $logitem.SaveChanges()
                 $i=$i+1
             }
        }
        Write-Output ('Number of logs disabled: ' + $i)
    }
    else
    {
        Write-Output 'Event logs already disabled'
    }
}
Catch
{
    $ErrorMessage = $_.Exception.Message
    Write-Output $ErrorMessage
}
