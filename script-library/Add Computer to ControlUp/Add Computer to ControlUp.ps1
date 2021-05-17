<#
    Add specified machines to CU monitoring - must be run on machine running CU Monitor.

    @guyrleech 11/05/20
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory,HelpMessage='Computer to add to CU')]
    [string]$computerName ,
    [Parameter(HelpMessage='ControlUp Folder to put computer in')]
    [string]$folderName
)

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'

if( ! ( $cuMonitorService = Get-CimInstance -ClassName win32_service -Filter "Name = 'cuMonitor'" ) )
{
    Throw "Unable to find the ControlUp Monitor service which is required for this script to run"
}

[string]$cudll = Join-Path -Path (Split-Path -Path ($cuMonitorService.PathName -replace '"') -Parent) -ChildPath 'ControlUp.PowerShell.User.dll'

if( ! (Test-Path -Path $cudll -PathType Leaf -ErrorAction SilentlyContinue ) )
{
    Throw "Unable to find `"$cudll`" which should be in the same folder as `"$cuMonitorService`""
}

if( ! ( $imported = Import-Module -Name $cudll -PassThru ) )
{
    Throw "Failed to import the PowerShell module in `"$cudll`""
}

Add-CUComputer -ADComputerName $computerName -DomainName $env:USERDNSDOMAIN -FolderPath $folderName

if( ! $? )
{
    Throw "Problem adding $computername.$domainname to folder `"$folderName`""
}

