<#
.SYNOPSIS
    Create a log file for a ControlUp Monitor by creating a Log4net file.
.DESCRIPTION
    Creates a log file with a user-defined logfile and script duration.

.PARAMETER logPath
    The folder in which to create the log file

.PARAMETER exename
    The ControlUp executable to create the log file for

.PARAMETER logDurationSeconds
    Number of seconds to allow log file to colect data

.PARAMETER loglevel
    The logging level to set in the log file
    
.PARAMETER FinalLogLevel
    The logging level to set in the log file once the logging period has finished. FATAL or OFF are recommended as continual logging can cause slowness

.EXAMPLE
    .\log4net.ps1 -LogPath "C:\temp" -logDuration 5 -logPath "C:\temp" -exeName "cuMonitor.exe"
    Saves a log file as cuMonitor.exe.log in the C:\tmp folder. The log4net file is saved for 5 seconds, after that it gets deleted.
    
.NOTES
    Modification History

    2022/11/07  Marcel Calef    Initial Version
    2022/11/20  Benjamin Skoda  Updates
    2022/12/12  Guy Leech       Parameter changes
    2022/12/13  Guy Leech       Changed log file name to include date, time and computer name. Added -FinalLogLevel, disabled nulling of log file name
    2024/02/28  Guy Leech       Added -supportMode to output log lines in a different format. Trimming quotes off log path
#>

[CmdletBinding()]

param (
    [ValidateNotNullorEmpty()]
    [String]$LogPath = 'c:\temp',

    [ValidateSet('cuMonitor.exe','cuagent.exe')]
    [String]$exeName = 'cuMonitor.exe', 

    [uint16]$logDurationSeconds = 60 ,

    [ValidateSet( 'DEBUG', 'ERROR', 'FATAL', 'ALL', 'VERBOSE', 'INFO', 'OFF' )]
    [string]$LogLevel = 'ERROR' ,
   
    [ValidateSet( 'DEBUG', 'ERROR', 'FATAL', 'ALL', 'VERBOSE', 'INFO', 'OFF' )]
    [string]$FinalLogLevel = 'FATAL' ,
   
    [ValidateSet( 'yes','no' )]
    [string]$supportMode = 'no'
)

$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { 'Continue' } else { 'SilentlyContinue' })
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { 'Continue' } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'ErrorAction' ] ) { $ErrorActionPreference } else { 'Stop' })

[int]$outputWidth = 400
try
{
    if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
    {
        Write-Verbose -Message "Setting output width to $outputWidth"
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions
        Write-Verbose -Message "Set output width to $($WideDimensions.width)"
    }
}
catch
{
    ## not much we can do but will hide the error since it is not fundamental to script functionality, just output
    Write-Warning -Message "Failed to set output width to $($WideDimensions.width) : $_"
}

[string]$conversionPattern = "'%date',%logger,'[%thread]','%level','%message%'%newline"

if( $supportMode -ieq 'yes' ) {
    $conversionPattern = "%date{yyyy-MM-dd HH:mm:ss.fff};[%level];%property{log4net:HostName};%appdomain;%logger;%stacktrace;%message;%newline"
}

[XML]$log4net = @"
<?xml version="1.0" encoding="UTF-8"?>
    <!--
        comments ommited
     -->
    <log4net>
      <appender name="RollingFileAppender" type="log4net.Appender.RollingFileAppender">
    		<file value="LogFile.log"/>
    		<appendToFile value="true"/>
    		<rollingStyle value="Size"/>
    		<maxSizeRollBackups value="2"/>
    		<maximumFileSize value="100MB"/>
    		<staticLogFileName value="true"/>
            <lockingModel type="log4net.Appender.FileAppender+MinimalLock" />
    		<layout type="log4net.Layout.PatternLayout">
    			<conversionPattern value="$conversionPattern"/>
    		</layout>
    	</appender>
      <root>
    		<level value="DEBUG"/>
        <appender-ref ref="RollingFileAppender"/>
    	</root>
    </log4net>
"@

## CU agent can quote paths so we unquote
$logPath = $logPath.Trim( '" ' )

##Write-Host "###########################"
Write-Host "User Inputs":
Write-Host "LogLevel set: $LogLevel"
Write-Host "Name of CU component: $($exeName)" 
##Write-Host "Path of CU components : $($exePath)" 
Write-Host "Run the logger for $logDurationSeconds seconds"
Write-Host "Log path is $LogPath"
##Write-Host "###########################"

if ( -Not ( Test-Path $LogPath -PathType Container)  ){
    if( -Not ( New-Item $logPath -ItemType Directory -Force ) ) {
        Throw "Problem creating log folder `"$logPath`""
    }
}

Function Set-LogLevel {
    param(
        [Parameter(Mandatory = $true)]
        [String]$LogLevel
    )
    
    $log4net.log4net.root.level.value = $LogLevel
}

Function Set-FileNameInLog {
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [String]$LogName
    )

    $log4net.log4net.appender.file.value = $LogName
}

function Create-LogFile {
    param (
        [String]$path ,
        [String]$exename
    )

    $exepath = Join-Path -Path $path -ChildPath $exename
    $global:PathToCheck = "$($exePath).log4net"

    try {
        $log4net.Save( $global:PathToCheck )
    }
    catch {
        Write-Warning "Failed write to log4net file `"$global:PathToCheck`" : $_"
    }
}

function Start-Logging {
    param (
        [Parameter(Mandatory = $true)]
        [Int]$logDurationSeconds ,
        [switch]$removeLog4netFile
    )

    Write-Verbose -Message "$(Get-Date -Format G): sleeping for $logDurationSeconds seconds"

    Start-sleep -Seconds $logDurationSeconds

    if( $removeLog4netFile ) {
        Remove-Item -Path $global:PathToCheck -Force -Confirm:$false
    }
}

function Add-RequiredACL { 
    param ($Object, 
        [System.Security.Principal.NTAccount]$Identity, 
        [System.Security.AccessControl.FileSystemRights]$AccessMask, 
        [System.Security.AccessControl.AccessControlType]$Type) 
    
    $InheritanceFlags = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
    $PropagationFlags = [System.Security.AccessControl.PropagationFlags]'None'
    $SD = $null
    $SD = get-acl -Path $Object
    if( $SD ) {
        $Rule = new-object -Typename System.Security.AccessControl.FileSystemAccessRule -argumentlist @($Identity, $AccessMask, $InheritanceFlags, $PropagationFlags, $Type) 
        $SD.AddAccessRule( $Rule ) 
        set-acl -Path $Object -AclObject $SD 
    } else {
        Write-Warning -Message "Failed to get security descriptor for `"$object`" so may not be able to write to it"
    }
}

function Restart-CUService {
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]$exeName
    )
    
    [string]$ServiceName = (Split-Path -Path $exeName -Leaf) -replace '\.[^\.]+$' ## strip last extension
    if ($ServiceName -imatch 'cuMonitor|cuAgent' ) {
        $serviceHandle = $null
        Write-Verbose -Message "$(Get-Date -Format G): restarting service $ServiceName"
        $serviceHandle = restart-service -Name $ServiceName -PassThru
        if( -Not $serviceHandle ) {
            Write-Warning -Message "Failed restart of service $ServiceName"
        }
        elseif( $serviceHandle.Status -ine 'Running' ) {
            Write-Warning -Message "Problem with restart of service $ServiceName - status is $($serviceHandle.Status)"
        }
    }
    ## not a service 
    Write-Host "$(Get-Date -Format G): Service $ServiceName restarted"
}

Function Get-ExePath {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$exeName
    )
    [string]$exePath = $null
    if( $exeName -match '^(cumonitor|cuagent)' ) {
        [string]$servicename = $Matches[1]
        $service = $null
        $service = Get-WmiObject -Class win32_service -Filter "name = '$servicename'"
        if( -Not $service ) {
            Write-Error "Unable to find definition for service $servicename"
        }
        else {
            ## "C:\Program Files\Smart-X\ControlUpAgent\Version 8.6.5.465\cuAgent.exe" /service
            if( $service.PathName -match '^"([^"]+)"' -or $service.PathName -match '^(.*)\s' ) {
                $exePath = Split-Path -Path $Matches[1] -Parent
            } else {
                Write-Warning -Message "Unable to extract exe folder from `"$($service.PathName)`""
            }
        }
    }
    elseif( $exeName -match 'cuconsole' ) {
        ## TODO get path from current running process
    }
    else {
        Write-Warning "Do not know what to do to get exe path for $exeName"
    }
    $exePath ## return
}
        
try {
    ## use well known SID for "NT AUTHORITY\NETWORK SERVICE" to avoid language issues https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/understand-security-identifiers
    Add-RequiredACL -object $LogPath -Identity  ([System.Security.Principal.SecurityIdentifier]('S-1-5-20')).Translate([System.Security.Principal.NTAccount]).Value -AccessMask Modify, Synchronize -Type Allow
}
catch {
    Write-Error $_.Exception.Message
}

Set-LogLevel -LogLevel $LogLevel
[string]$exePath = Get-ExePath -exeName $exeName
if( $exePath ) {
    [string]$logfileFullPath = Join-Path -Path $LogPath -ChildPath "$($exeName)_$($env:COMPUTERNAME)_$([datetime]::Now.ToString( 'yyyy-MM-dd_HH-mm-ss' )).log"
    Set-FileNameInLog -LogName $logfileFullPath
    Write-Host "Path of logfile: $($logfileFullPath)" 
    Write-Verbose -Message "exe path is `"$exePath`""
    if( Test-Path -Path $logfileFullPath -PathType Leaf ) {
        $moveError = $null
        if( Move-Item -Path $logfileFullPath -Destination "$($logfileFullPath).old" -PassThru -ErrorAction SilentlyContinue -ErrorVariable moveError) {
            Write-Host "Moved previous log file to `"$($logfileFullPath).old`""
        }
        else {
            Write-Warning -Message "Problem backing up existing log file to `"$($logfileFullPath).old`" : $moveError"
        }
    }
    Create-LogFile -path $exePath -exename $exeName
    Start-Logging -LogDuration $logDurationSeconds
    ## Don't restart as will cause console to lose connection if it is the cuagent since the console is communicating with it to run the script
    ## Restart-CUService -exeName $exeName
    Set-LogLevel -LogLevel $FinalLogLevel
    ## Set-FileNameInLog -LogName ''
    $log4net.Save( $global:PathToCheck )
}
## else will already have given error/reason

