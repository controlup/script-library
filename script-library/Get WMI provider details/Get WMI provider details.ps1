<#
.SYNOPSIS

Show the WMI provider module (dll) currently loaded into the wmiprvse.exe process whose PID is passed as the only parameter

.DETAILS

Looks for the latest event with id 5857 for the given PID in the Microsoft-Windows-WMI-Activity/Operational event log

.PARAMETER wmiprvseProcess

The PID of the wmiprvse.exe process to look for in the Microsoft-Windows-WMI-Activity/Operational event log

.CONTEXT

Process 

.MODIFICATION_HISTORY:

@guyrleech 31/10/19

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage="Process id of the wmiprvse.exe process")]
    [int]$wmiprvseProcess
)

$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { 'Continue' } else { 'SilentlyContinue' })
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { 'Continue' } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'ErrorAction' ] ) { $ErrorActionPreference } else { 'Stop' })

[string]$eventLog = 'Microsoft-Windows-WMI-Activity/Operational'

$theProcess = Get-Process -Id $wmiprvseProcess -ErrorAction SilentlyContinue

if( ! $theProcess )
{
    Throw "No process found for pid $wmiprvseProcess"
}

if( $theProcess.Name -ne 'wmiprvse' )
{
    Throw "Process must be wmiprvse, this is $($theProcess.Name)"
}

$event = Get-WinEvent -LogName $eventLog -FilterXPath "*[System[EventID=5857] and UserData/*/ProcessID=$wmiprvseProcess]" -MaxEvents 1

if( $event )
{
    if( $event.TimeCreated -lt $theProcess.StartTime )
    {
        Throw "Event found for this process is from $(Get-Date -Date $event.TimeCreated -Format G) but process was started at $(Get-Date -Date $theprocess.StartTime -Format G)"
    }
    else
    {
        $module = $event.Properties[4].Value ## This is providerpath in the event text
        $moduleDetails = $null
        [string]$message = "WMI Provider is $module"
        if( $module )
        {
            if( $moduleDetails = Get-ItemProperty -Path ([System.Environment]::ExpandEnvironmentVariables( $module )) -ErrorAction SilentlyContinue )
            {
                $message += "($($moduleDetails.VersionInfo.FileDescription)) from $($moduleDetails.VersionInfo.CompanyName) version $($moduleDetails.VersionInfo.FileVersion) created $(Get-Date -Date $moduleDetails.CreationTime -Format G)"
            }
        }
        Write-Output -InputObject $message
    }
}
else
{
    Throw "No event found for pid $wmiprvseProcess in event log $eventlog"
}

