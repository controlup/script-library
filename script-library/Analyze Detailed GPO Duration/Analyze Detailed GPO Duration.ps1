    <#
    .SYNOPSIS
        Enables or Disables Group Policy Preferences logging and evaluates how long each GPO took for each Group Policy Preferences Extension.

    .DESCRIPTION
        Enables or Disables Group Policy Preferences logging and evaluates how long each GPO took for each Group Policy Preferences Extension.

    .PARAMETER  <Enable <switch>>
		Enable GPP Logging
		
	.PARAMETER	<Disable <switch>>
		Disables GPP Logging

	.PARAMETER  <SessionId>18
		The Session ID of the user the function reports for. 

    .PARAMETER  <User <string[]>>BOTTHEORY\awellman
		User name to pull information

    .EXAMPLE
        . .\GPPEvaluation.ps1 -Enable
        Enables Group Policy Preferences logging for all installed extensions

    .EXAMPLE
        . .\GPPEvaluation.ps1 -Disable
        Enables Group Policy Preferences logging for all installed extensions

    .EXAMPLE
        . .\GPPEvaluation.ps1 -SessionId 18 -User BOTTHEORY\awellman
        Attempts to evaluate Group Policy Preferences extension durations and sums up total time spent on a group policy object.

    .NOTES
        This script must be run on a machine where the user is currently logged on.

    .CONTEXT
        Session

    .MODIFICATION_HISTORY
        Created TTYE : 2020-06-16


    AUTHOR: Trentent Tye
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$false,HelpMessage='Enter the username in the format DOMAIN\Username')][ValidateNotNullOrEmpty()]       [string]$User,
    [Parameter(Mandatory=$false,HelpMessage='Enter the session ID')][ValidateNotNullOrEmpty()]                                   [int]$SessionId,
    [Parameter(Mandatory=$false,HelpMessage='Enable GPP Logging')]                                                               [switch]$Enable = $false,
    [Parameter(Mandatory=$false,HelpMessage='Disable GPP Logging')]                                                              [switch]$Disable = $false,
    [Parameter(Mandatory=$false,HelpMessage='Ignore Time Delta Check')]                                                          [switch]$IgnoreTimeDeltaCheck = $false
)


#Requires -Version 5.0
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
###$VerbosePreference = "continue"

# Altering the size of the PS Buffer
[int]$outputWidth = 400
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

if ($user) {
    [string]$UserDomain, [string]$Username = $user.split("\")
    Write-Verbose -Message "Username   : $username"
    Write-Verbose -Message "UserDomain : $userDomain"
}

if ($SessionId) {
    Write-Verbose -Message "SessionId  : $SessionId"
}
if ($Enable) {
    Write-Verbose -Message "Enable switch found."
}
if ($Disable) {
    Write-Verbose -Message "Disable switch found."
}

function Toggle-GPPLogging {
    param(
        [Parameter(Mandatory=$false)][switch]$Enable,
        [Parameter(Mandatory=$false)][switch]$Disable
    )

    [array]$GPObjs = @( Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\GPExtensions" | ForEach-Object { @{ $_.PSChildName = $_.GetValue('')} } )

    if ($Enable) {
        #Enable Logging
        Write-Output "Enabling GPP Logging"
        $GPPolicyKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Group Policy"
        if (-not(Test-Path -Path $GPPolicyKey)) {
                New-Item -Path $GPPolicyKey | Out-Null
            }

        foreach ($GPObj in $GPObjs) {
            if (-not(Test-Path -Path $GPPolicyKey\$($GPObj.keys))) {
                New-Item -Path $GPPolicyKey\$($GPObj.keys) | Out-Null
            }
            New-ItemProperty -Path $GPPolicyKey\$($GPObj.keys) -Name LogLevel -Value 3 -PropertyType DWORD -Force | Out-Null   #Informational, Warnings and Errors
            New-ItemProperty -Path $GPPolicyKey\$($GPObj.keys) -Name TraceFileMaxSize -Value 16384 -PropertyType DWORD -Force | Out-Null
            New-ItemProperty -Path $GPPolicyKey\$($GPObj.keys) -Name TraceLevel -Value 2 -PropertyType DWORD -Force | Out-Null
            New-ItemProperty -Path $GPPolicyKey\$($GPObj.keys) -Name TraceFilePathMachine -Value "C:\ProgramData\GroupPolicy\Preference\Trace\Computer.log" -PropertyType ExpandString -Force | Out-Null
            New-ItemProperty -Path $GPPolicyKey\$($GPObj.keys) -Name TraceFilePathPlanning -Value "C:\ProgramData\GroupPolicy\Preference\Trace\Planning.log" -PropertyType ExpandString -Force | Out-Null
            New-ItemProperty -Path $GPPolicyKey\$($GPObj.keys) -Name TraceFilePathUser -Value "C:\ProgramData\GroupPolicy\Preference\Trace\User.log" -PropertyType ExpandString -Force | Out-Null
        }
        Start-Process gpupdate -ArgumentList @("/force") -WindowStyle Minimized
    }
    if ($Disable) {
        #Disable Logging
        Write-Output "Disabling GPP Logging"
        $GPPolicyKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Group Policy"
        if (-not(Test-Path -Path $GPPolicyKey)) {
                New-Item -Path $GPPolicyKey | Out-Null
            }

        foreach ($GPObj in $GPObjs) {
            if (-not(Test-Path -Path $GPPolicyKey\$($GPObj.keys))) {
                New-Item -Path $GPPolicyKey\$($GPObj.keys) | Out-Null
            }
            New-ItemProperty -Path $GPPolicyKey\$($GPObj.keys) -Name TraceLevel -Value 0 -PropertyType DWORD -Force | Out-Null
        }
        Start-Process gpupdate -ArgumentList @("/force") -WindowStyle Minimized
    }
}

function Test-RegistryKeyValue {
    <#
    .SYNOPSIS
    Tests if a registry value exists.

    .DESCRIPTION
    The usual ways for checking if a registry value exists don't handle when a value simply has an empty or null value.  This function actually checks if a key has a value with a given name.

    .EXAMPLE
    Test-RegistryKeyValue -Path 'hklm:\Software\Carbon\Test' -Name 'Title'

    Returns `True` if `hklm:\Software\Carbon\Test` contains a value named 'Title'.  `False` otherwise.

    #TTYE: Thank you Aaron Jenson for this function

    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]
        # The path to the registry key where the value should be set.  Will be created if it doesn't exist.
        $Path,

        [Parameter(Mandatory=$true)]
        [string]
        # The name of the value being set.
        $Name
    )

    if( -not (Test-Path -Path $Path -PathType Container) )
    {
        return $false
    }

    $properties = Get-ItemProperty -Path $Path 
    if( -not $properties )
    {
        return $false
    }

    $member = Get-Member -InputObject $properties -Name $Name
    if( $member )
    {
        return $true
    }
    else
    {
        return $false
    }

}

if ($enable) {
    Toggle-GPPLogging -Enable
    exit
}

if ($disable) {
    Toggle-GPPLogging -Disable
    exit
}


#region LSA to get the definitive logon time
$LSADefinitions = @'
    [DllImport("secur32.dll", SetLastError = false)]
    public static extern uint LsaFreeReturnBuffer(IntPtr buffer);

    [DllImport("Secur32.dll", SetLastError = false)]
    public static extern uint LsaEnumerateLogonSessions
            (out UInt64 LogonSessionCount, out IntPtr LogonSessionList);

    [DllImport("Secur32.dll", SetLastError = false)]
    public static extern uint LsaGetLogonSessionData(IntPtr luid, 
        out IntPtr ppLogonSessionData);

    [StructLayout(LayoutKind.Sequential)]
    public struct LSA_UNICODE_STRING
    {
        public UInt16 Length;
        public UInt16 MaximumLength;
        public IntPtr buffer;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct LUID
    {
        public UInt32 LowPart;
        public UInt32 HighPart;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SECURITY_LOGON_SESSION_DATA
    {
        public UInt32 Size;
        public LUID LoginID;
        public LSA_UNICODE_STRING Username;
        public LSA_UNICODE_STRING LoginDomain;
        public LSA_UNICODE_STRING AuthenticationPackage;
        public UInt32 LogonType;
        public UInt32 Session;
        public IntPtr PSiD;
        public UInt64 LoginTime;
        public LSA_UNICODE_STRING LogonServer;
        public LSA_UNICODE_STRING DnsDomainName;
        public LSA_UNICODE_STRING Upn;
    }

    public enum SECURITY_LOGON_TYPE : uint
    {
        Interactive = 2,        //The security principal is logging on 
                                //interactively.
        Network,                //The security principal is logging using a 
                                //network.
        Batch,                  //The logon is for a batch process.
        Service,                //The logon is for a service account.
        Proxy,                  //Not supported.
        Unlock,                 //The logon is an attempt to unlock a workstation.
        NetworkCleartext,       //The logon is a network logon with cleartext 
                                //credentials.
        NewCredentials,         //Allows the caller to clone its current token and
                                //specify new credentials for outbound connections.
        RemoteInteractive,      //A terminal server session that is both remote 
                                //and interactive.
        CachedInteractive,      //Attempt to use the cached credentials without 
                                //going out across the network.
        CachedRemoteInteractive,// Same as RemoteInteractive, except used 
                                // internally for auditing purposes.
        CachedUnlock            // The logon is an attempt to unlock a workstation.
    }
'@

## we use LSA to get the definitive logon time
if( ! ( ([System.Management.Automation.PSTypeName]'Win32.Secure32').Type ) )
{
    Add-Type -MemberDefinition $LSADefinitions -Name 'Secure32' -Namespace 'Win32' -UsingNamespace System.Text -Debug:$false
}

$count = [UInt64]0
$luidPtr = [IntPtr]::Zero

[uint64]$ntStatus = [Win32.Secure32]::LsaEnumerateLogonSessions( [ref]$count , [ref]$luidPtr )

if( $ntStatus )
{
    Write-Error -Message "LsaEnumerateLogonSessions failed with error $ntStatus"
}
elseif( ! $count )
{
    Write-Error -Message "No sessions returned by LsaEnumerateLogonSessions"
}
elseif( $luidPtr -eq [IntPtr]::Zero )
{
    Write-Error -Message "No buffer returned by LsaEnumerateLogonSessions"
}
else
{   
    Write-Verbose -Message "$count sessions retrieved from LSASS"
    [IntPtr] $iter = $luidPtr
    $earliestSession = $null
    [array]$lsaSessions = @( For ([uint64]$i = 0; $i -lt $count; $i++)
    {
        $sessionData = [IntPtr]::Zero
        $ntStatus = [Win32.Secure32]::LsaGetLogonSessionData( $iter , [ref]$sessionData )

        if( ! $ntStatus -and $sessionData -ne [IntPtr]::Zero )
        {
            $data = [System.Runtime.InteropServices.Marshal]::PtrToStructure( $sessionData , [type][Win32.Secure32+SECURITY_LOGON_SESSION_DATA] )

            if ($data.PSiD -ne [IntPtr]::Zero)
            {
                $sid = New-Object -TypeName System.Security.Principal.SecurityIdentifier -ArgumentList $Data.PSiD

                #extract some useful information from the session data struct
                [datetime]$loginTime = [datetime]::FromFileTime( $data.LoginTime )
                $thisUser = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.Username.buffer) #get the account name
                $thisDomain = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.LoginDomain.buffer) #get the domain name
                try
                { 
                    $secType = [Win32.Secure32+SECURITY_LOGON_TYPE]$data.LogonType
                }
                catch
                {
                    $secType = 'Unknown'
                }

                if( ! $earliestSession -or $loginTime -lt $earliestSession )
                {
                    $earliestSession = $loginTime
                }
                if( $thisUser -eq $Username -and $thisDomain -eq $UserDomain -and $secType -match 'Interactive' )
                {
                    $authPackage = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.AuthenticationPackage.buffer) #get the authentication package
                    $session = $data.Session # get the session number
                    if( $session -eq $SessionId )
                    {
                        $logonServer = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.LogonServer.buffer) #get the logon server
                        $DnsDomainName = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.DnsDomainName.buffer) #get the DNS Domain Name
                        $upn = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.upn.buffer) #get the User Principal Name

                        [pscustomobject]@{
                            'Sid' = $sid
                            'Username' = $thisUser
                            'Domain' = $thisDomain
                            'Session' = $session
                            'LoginId' = [uint64]( $loginID = [Int64]("0x{0:x8}{1:x8}" -f $data.LoginID.HighPart , $data.LoginID.LowPart) )
                            'LogonServer' = $logonServer
                            'DnsDomainName' = $DnsDomainName
                            'UPN' = $upn
                            'AuthPackage' = $authPackage
                            'SecurityType' = $secType
                            'Type' = $data.LogonType
                            'LoginTime' = [datetime]$loginTime
                        }
                    }
                }
            }
            [void][Win32.Secure32]::LsaFreeReturnBuffer( $sessionData )
            $sessionData = [IntPtr]::Zero
        }
        $iter = $iter.ToInt64() + [System.Runtime.InteropServices.Marshal]::SizeOf([type][Win32.Secure32+LUID])  # move to next pointer
    }) | Sort-Object -Descending -Property 'LoginTime'

    [void]([Win32.Secure32]::LsaFreeReturnBuffer( $luidPtr ))
    $luidPtr = [IntPtr]::Zero

    Write-Verbose -Message "Found $(if( $lsaSessions ) { $lsaSessions.Count } else { 0 }) LSA sessions for $UserDomain\$Username, earliest session $(if( $earliestSession ) { Get-Date $earliestSession -Format G } else { 'never' })"
}


if( $lsaSessions -and $lsaSessions.Count )
{

    ## get all logon ids for logons that happened at the same time
    [array]$loginIds = @( $lsaSessions | Where-Object { $_.LoginTime -eq $lsaSessions[0].LoginTime } | Select-Object -ExpandProperty LoginId )
    
    if( ! $loginIds -or ! $loginIds.Count )
    {
        Write-Error -Message "Found no login ids for $username at $(Get-Date -Date $lsaSessions[0].LoginTime -Format G)"
    }
    $Logon = New-Object -TypeName psobject -Property @{
        LogonTime = $lsaSessions[0].LoginTime
        LogonTimeFileTime = $lsaSessions[0].LoginTime.ToFileTime()
        FormatTime = $lsaSessions[0].LoginTime.ToString( 'HH:mm:ss.fff' ) 
        LogonID = $loginIds
        UserSID = $lsaSessions[0].Sid
        Type = $lsaSessions[0].Type
        UserName = $Username
        UserDomain = $UserDomain
        ## No point saving XD details since these cannot be used offline
    }
}
else
{
    Throw "Failed to retrieve logon session for $UserDomain\$Username from LSASS"
}
#endregion

if (-not(Test-Path -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Group Policy")) {
    Write-Error -Message "Group Policy Logging not set for Group Policy Preference items. Try rerunning this script wtih -Enable to enable logging."
}

$UserLogPath = Get-ChildItem -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Group Policy"
$GPPExtensions = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\GPExtensions"

$UserLogPaths = New-Object System.Collections.Generic.List[PSObject]
if ($UserLogPath.count -ne $GPPExtensions.count) {
    Write-Verbose -Message "The number of Group Policy Extensions ($($GPPExtensions.count)) is not equal to the number of extensions configured with logging ($($UserLogPath.count)). `nRun this script with -Enable to enable logging for all extensions."
} 

foreach ($Path in $UserLogPath) {
    $logPath = $path.GetValue("TraceFilePathUser")
    Write-Verbose "LogPath: $logPath"
    if ($logPath -like "%COMMONAPPDATA%*") {
        #We need to replace this variable because it doesn't exist in any available context except for group policy?
        Write-Verbose -Message "Replacing Group Policy Variable"
        $UserLogPaths.add($logPath -replace ("%COMMONAPPDATA%","$env:ProgramData"))
    } else {
        $UserLogPaths.add($logPath)
    }
}


$logFilePath = $UserLogPaths | Sort-Object -Unique
Write-Verbose -Message "LogFilePath = $logFilePath"
if (($logFilePath | Measure-Object).count -ne 1) {
    Write-Error -Message "Multiple user log paths found for Group Policy Extensions Logging. Please consolidate to 1 log file or run this script with -Enable to force this configuration. `nNumber of Paths: $($logFilePath.count). `n$($logFilePath)"
} else {
    Write-Verbose -Message "GPP Extensions Logging Path is set in registry to: $($UserLogPaths | sort -Unique)"
}


$FileName = [System.IO.Path]::GetFileNameWithoutExtension($logFilePath)
$File = "$FileName.log"
$LogDirectory = $logFilePath | Split-Path -Parent
if (Test-Path -Path $(Join-Path -Path "$LogDirectory" -ChildPath "$fileName.Bak")) {
    $BakFiles = ls $LogDirectory -Filter "$fileName.bak"
    Write-Verbose -Message "Bak User Files  : $BakFiles"
} else {
    $BakFiles = $null
}
Write-Verbose -Message "File            : $file"
Write-Verbose -Message "Log Directory   : $LogDirectory"



if (($BakFiles | Measure-Object).count -eq 1) {   ## it appears MS only generates 1 bak file per.  If this changes --> good luck Future Trentent
    $LogFile = Get-Content -Path $(Join-Path -Path "$LogDirectory" -ChildPath "$fileName.Bak")
    Write-Verbose -Message "Bak logfile             : $($($LogFile | Measure-Object).count) lines"
} elseif (($BakFiles | Measure-Object).count -gt 1) {
    Write-Error -Message "Only 1 .bak file expected. Multiple found."
}

Write-Verbose -Message "Searching file: $($logFilePath)"
if (Test-Path -Path Variable:\LogFile) {
    $LogFile += Get-Content -Path $logFilePath
} else {
    $LogFile = Get-Content -Path $logFilePath
}
Write-Verbose -Message "LogFile with Bak and Log: $($($LogFile | Measure-Object).count) lines"

#we need to adjust for time zone as the GPP log file will use the "SYSTEM" time and not adjust for our user time...  *ugh*
try {
    $timeZoneEvent = Get-WinEvent -FilterHashtable @{logname='system'; id=22 ;providerName='Microsoft-Windows-Kernel-General';StartTime=$logon.LogonTime} -MaxEvents 1
} catch {
    Write-Verbose -Message "No Timezone events"
}

$TZOffset = 0
if (Test-Path -Path Variable:\timezoneevent) {
    [xml]$timeZoneEventXml = $timeZoneEvent.ToXml()
    $TZObj = @{}
    if ($timeZoneEventXml.Event.EventData.Data.Count -eq 2) {
        foreach ($tzData in $timeZoneEventXml.Event.EventData.Data) { $TZObj.Add($tzData.Name,$tzData."#text") }
    }
    $TZOffset =  $TZObj.'OldBias' - $TZObj.'NewBias'
    Write-Verbose -Message "Timezone Offset: $TZOffset"
}

$foundUsernameLog = $LogFile | Select-String -SimpleMatch "%LogonUser% = `"$username`""
Write-Verbose -Message "Number of lines in the log file with this username: $($($foundUsernameLog | Measure-Object).count)"
$GPPUsernameLog = New-Object System.Collections.Generic.List[PSObject]
Foreach ($line in $foundUsernameLog) {
    $line.ToString() | Select-String -Pattern "(.*?) \[(pid=)(.*?),(tid=)(.*?)\] (.*?$)" -AllMatches | ForEach-Object{
        if ( $_.Matches.groups.count -eq 7) {
            $DateTime = ([DateTime]$_.Matches.groups[1].value).AddMinutes($TZOffset)
            $processid = $_.Matches.groups[3].value
            $threadId = $_.Matches.groups[5].value
            $message = $_.Matches.groups[6].value
            
            $obj = [PSCustomObject]@{
                Time = $DateTime
                ProcessId = $processId
                ThreadId = $threadID
                Message = $Message
            }
            $null = $GPPUsernameLog.Add($obj)
        }
    }
}

Write-Verbose -Message "Found $($GPPUsernameLog.Count) instances of $username in the log"

if ($GPPUsernameLog.Count -eq 0) {
    Write-Error -Message "Unable to find any events for $Username at $($Logon.logontime)"
}

#find the first event *after* the logon time with this user and grab the pid/tid
foreach ($event in $GPPUsernameLog) {
    if ($event.time -ge $Logon.LogonTime) {
        $FirstEventAfterLogon = $event
        Write-Verbose -Message "Logon Time: $($Logon.LogonTime)"
        Write-Verbose -Message "Found first Group Policy processing event after logon: $($event.time)"
        break
    } 
}

#see if data is within 120 seconds...? (might need to extend...?)

$timeDifference = $(($event.time-$Logon.LogonTime).TotalSeconds)
Write-Verbose -Message "Time difference between logon and first event: $($timeDifference) seconds"

if (-not($IgnoreTimeDeltaCheck)) {
    if ($timeDifference -ge 120) {
        Write-Error -Message "The difference between the session logon time and the group policy preferences event is greater than 120 seconds. The logs might have rolled over, or some other error may have occurred.`nSession Logon Time  : $($logon.LogonTime)`nGPP Event Timestamp : $($event.time)`nTrying running this script shortly after the user has logged on."
    }
}
Write-Verbose -Message "PID: $($FirstEventAfterLogon.processId)"
Write-Verbose -Message "TID: $($FirstEventAfterLogon.threadID)"

Write-Verbose -Message "Filtering log down to the PID/TID combination"
$filteredLogFile = $LogFile | Select-String -SimpleMatch "pid=$($FirstEventAfterLogon.processId),tid=$($FirstEventAfterLogon.threadID)"
Write-Verbose -Message "Found $($filteredLogFile.count) line of activity"


Write-Verbose -Message "Generating object out of text"
$GPPLogObj = New-Object System.Collections.Generic.List[PSObject]
Foreach ($line in $filteredLogFile) {
    $line.ToString() | Select-String -Pattern "(.*?) \[(pid=)(.*?),(tid=)(.*?)\] (.*?$)" -AllMatches | ForEach-Object{
        if ( $_.Matches.groups.count -eq 7) {
            $DateTime = ([DateTime]$_.Matches.groups[1].value).AddMinutes($TZOffset)
            $processid = [Convert]::ToInt64(($_.Matches.groups[3].value),16)
            $threadId = [Convert]::ToInt64(($_.Matches.groups[5].value),16)
            $message = $_.Matches.groups[6].value
            
            
            $obj = [PSCustomObject]@{
                Time = $DateTime
                ProcessId = $processId
                ThreadId = $threadID
                Message = $Message
            }
            $null = $GPPLogObj.Add($obj)
        }
    }
}

Write-Verbose -Message "Generated $($GPPLogObj.count) objects"
$StartStopEvents = $GPPLogObj | Where-Object {$_.Message -like "*ProcessGroupPolicyEx*" -or $_.Message -like "*GPO Display Name*" -or $_.Message -like "Completed apply GPO*"}
Write-Verbose -Message "Found $($StartStopEvents.count) Start-Stop Events"

Write-Verbose -Message "Getting Group Policy Extensions and their function names"

$GPPhases = @{}
foreach ($GPPExtension in $GPPExtensions) {

    if (Test-RegistryKeyValue -Path $GPPExtension.PSPath -Name ProcessGroupPolicyEx) {
        $GPEx = Get-ItemPropertyValue $GPPExtension.PSPath -Name ProcessGroupPolicyEx -ErrorAction SilentlyContinue
    } elseif (Test-RegistryKeyValue -Path $GPPExtension.PSPath -Name ProcessGroupPolicy) {
        $GPEx = Get-ItemPropertyValue $GPPExtension.PSPath -Name ProcessGroupPolicy -ErrorAction SilentlyContinue
    } else {
        Write-Verbose -Message "No GP Processing for $($GPPExtension.PSChildName)"
        continue
    }

    if ($GPPhases.ContainsKey($GPEx)) {
        Write-Verbose -Message "GP Phase already exists for $GPEx"
        continue
    }
    
    try {
        $GPFriendlyName = (Get-ItemProperty $GPPExtension.PSPath).'(default)'
    } catch {
        $GPFriendlyName = "Unknown"
    }

    if ($GPFriendlyName -notlike "*[A-Za-z]*") {
        $GPFriendlyName = $GPEx
    }
    $GPPhases.Add($GPEx,$GPFriendlyName)
}
Write-Verbose -Message "Found $($GPPhases.Count) unique Group Policy extensions"


#Get how long the individual GPO's spent in each phase
$GPPInitialization = New-Object System.Collections.Generic.List[PSObject]
$GPPEndResult = New-Object System.Collections.Generic.List[PSObject]
foreach ($event in $StartStopEvents) {
    if ($event.Message -like "Entering*") {
        $Matches = $event.Message | Select-String -Pattern " (.*?)\(\)" -AllMatches
        $eventPhase = $GPPhases."$($Matches.Matches.groups[1])"  #lookup Phase in GPPhase hashtable
        Write-Verbose -Message "GPPhase: $eventPhase"
        $GPPhaseStartEvent = $event

        $indexOfCurrentEvent = ($StartStopEvents.Time.IndexOf($event.Time))
        Write-Verbose -Message "Index: $indexOfCurrentEvent"
        $currentTime = $event.Time
        $completedGPPInitializationEvent = $StartStopEvents[($StartStopEvents.Time.IndexOf($event.Time))+1]
        #Write-Verbose -Message "Completed GPO Event: $($completedGPOEvent)"
        $completionTime = $completedGPPInitializationEvent.Time


        $obj = [PSCustomObject]@{
                Time = $currentTime
                "Duration (ms)" = ($completionTime-$currentTime).TotalMilliseconds
                GPExtension = $eventPhase
        }
        $null = $GPPInitialization.Add($obj)
    }

    if ($event.Message -like "GPO Display Name*") {
        $GPO = ($event.Message | Select-String -Pattern "( : )(.+)" -AllMatches).Matches.Groups[2].Value
        
        Write-Verbose -Message "GPO: $GPO"
        $indexOfCurrentEvent = ($StartStopEvents.Time.IndexOf($event.Time))
        Write-Verbose -Message "Index: $indexOfCurrentEvent"
        $currentTime = $event.Time
        $completedGPOEvent = $StartStopEvents[($StartStopEvents.Time.IndexOf($event.Time))+1]
        #Write-Verbose -Message "Completed GPO Event: $($completedGPOEvent)"
        $completionTime = $completedGPOEvent.Time  #find the current index of the current message "plus one" to find the completion time. Assumption is the next event is the completion time...  Might want to add a check in the future (says past Trentent)
        Write-Verbose -Message "Current Time: $($currentTime.ToString("s.fff"))"
        Write-Verbose -Message "Completion Time: $($completionTime.ToString("s.fff"))"


        $obj = [PSCustomObject]@{
                Time = $currentTime
                "Duration (ms)" = ($completionTime-$currentTime).TotalMilliseconds
                GPO = $GPO
                GPExtension = $eventPhase
        }
        $null = $GPPEndResult.Add($obj)
        
    }
}

#Get the total time per GPO and the extensions
$GPTotalTime = New-Object System.Collections.Generic.List[PSObject]
foreach ($GPO in ($GPPEndResult.GPO | Sort -Unique)) {
    $count = 0
    $totalTime = 0
    $extensions = ""
    $GPPEndResult.Where({$_.GPO -like $GPO}) | foreach {
        $count = $count+1
        $totaltime = ($_."Duration (ms)" +$totaltime)
        if ($count -ge 2) {
            $extensions = $_.GPExtension + ",$extensions"
        } else {
            $extensions = $_.GPExtension
        }
    }
    $obj = [PSCustomObject]@{
            "Duration (ms)" = $totaltime
            GPO = $GPO
            GPExtensions = $Extensions
    }
    $null = $GPTotalTime.Add($obj)
}

Write-Output "User                        : $($username)"
Write-Output "Logon Time                  : $($Logon.LogonTime)"
Write-Output "GPP Event Time              : $($FirstEventAfterLogon.Time)`n`n"
Write-Output "GPP Initialization Duration"
Write-Output "$($GPPInitialization | Sort -Property "Time" | Out-String)" 

Write-Output "`nGPO Processing Breakdown per GPExtension"
Write-Output "$($GPPEndResult | Sort -Property "Duration (ms)" -Descending | Out-String)" 

Write-Output "`nGPO Total Processing Duration"
Write-Output "$($GPTotalTime | Sort -Property "Duration (ms)" -Descending | Out-String)" 

