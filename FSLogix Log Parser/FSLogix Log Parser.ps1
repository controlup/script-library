<#
    .SYNOPSIS
        Finds all FSLogix Profile events for a user's logon

    .DESCRIPTION
        Finds all FSLogix Profile events for a user's logon for review and troubleshooting.

    .EXAMPLE
        . .\Get-FSLogixProfileLog.ps1 -User BOTTHEORY\amttye -SessionId 2
        Gets all events for the user "amttye" from domain "bottheory" on this machine

    .EXAMPLE
        . .\Get-FSLogixProfileLog.ps1 -User BOTTHEORY\amttye -SessionId 2
        Gets all events for the user "amttye" from domain "bottheory" on this machine

    .NOTES
        This script must be run on a machine where the user is currently logged on.

    .CONTEXT
        Session

    .MODIFICATION_HISTORY
        Created TTYE : 2019-11-19
        Edit: Ton de Vreede 2022-9-15 - small bugfix for error handling


    AUTHOR: Trentent Tye
#>
[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='Enter the username in the format DOMAIN\Username')][ValidateNotNullOrEmpty()]       [string]$User,
    [Parameter(Mandatory=$true,HelpMessage='Enter the session ID')][ValidateNotNullOrEmpty()]                                   [int]$SessionId
)


$ErrorActionPreference = "Stop"
###$VerbosePreference = "continue"

[int]$outputWidth = 800
# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

$userdomain = ($user -split "\\")[0]
$username = ($user -split "\\")[1]

function Get-FSLogixProfileEvents {
    [CmdletBinding()]
    param(
    [Parameter(Mandatory=$true)]
    [DateTime]
    $Start,

    [Parameter(Mandatory=$true)]
    [String]
    $Username
    )


    Write-Verbose "Username: `"$($Username)`""
    Write-Verbose "StartTime: `"$($Start)`""

    #FSLogix Log path is here: C:\ProgramData\FSLogix\Logs\Profile
    #at the time of this testing version 2.9.7205.27375 of FSLogix provided all the necessary information

    Write-Verbose "Looking for file with `"$($($start).ToString("yyyyMMdd"))`" in the file name"
    try {
        $FSLogixLogDir = Get-ItemPropertyValue -Path HKLM:\SOFTWARE\FSLogix\Logging -Name Logdir
        Write-Verbose "LogDir value configured. LogDir set to $FSLogixLogDir"
        }
    Catch {
        #LogDir registry value not found. Set to default:
            Write-Verbose "LogDir value not set. Setting LogDir to default path"
            $FSLogixLogDir = "C:\ProgramData\FSLogix\Logs"
        }

    $profileLog = ls "$FSLogixLogDir\Profile" | Where {$_.Name -like "*$($($start).ToString("yyyyMMdd"))*"}

    try {
        Test-Path $profileLog.FullName | out-null
    } catch {
        Write-Error "Unable to determine or find FSLogix profile log file."
        break
    }
    Write-Verbose "Found Profile Log file: $($profileLog.FullName)"


    $FSLogixLog = Get-Content "$($profileLog.fullname)"
    $FSLogixLogObject = New-Object System.Collections.ArrayList

    #Create powershell object out of the FSLogix Log.
    Foreach ($line in $FSLogixLog) {
        $line | Select-String -Pattern "\[(.*?)\]|.+" -AllMatches | ForEach-Object{
            if ( $_.Matches.count -eq 4) { #ignore all lines that don't conform to the grid table
                $MMddyyyy = $(($start).ToString("MM/dd/yyyy"))
                $time = $($_.Matches[0].Value -replace ("\[","") -replace ("\]",""))
                $FSLogixTime = [datetime]"$MMddyyyy $time"
                if ($FSLogixTime -ge $start) {
                    $obj = [PSCustomObject]@{
                        Time = $FSLogixTime
                        ThreadId = $_.Matches[1].Value -replace ("\[","") -replace ("\]","")
                        LogLevel = $_.Matches[2].Value -replace ("\[","") -replace ("\]","")
                        Message = $_.Matches[3].Value.Trim()
                    }

                    $FSLogixLogObject.Add($obj)|Out-Null
                }
            }
        }
    }

    
    $SessionEvents = $FSLogixLogObject | Where {$_.Message -like "*LoadProfile: $username*"}
    Write-Verbose "Number of SessionEvents: $($SessionEvents.Count)"
    $FSLogixStartEvent = $SessionEvents[0].time
    $FSLogixEndEvent = $SessionEvents[1].time
    $returnObject = New-Object System.Collections.ArrayList
    foreach ($FSLogixLogLine in $FSLogixLogObject) {
        if (($FSLogixLogLine.Time -le $FSLogixEndEvent) -and ($FSLogixLogLine.Time -ge $FSLogixStartEvent)) {
            $returnObject.Add($FSLogixLogLine)|Out-Null
        }
    }

    return $returnObject

}

#region Login Information Gathering
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

if( ! ( ([System.Management.Automation.PSTypeName]'Win32.Secure32').Type ) )
{
    Add-Type -MemberDefinition $LSADefinitions -Name 'Secure32' -Namespace 'Win32' -UsingNamespace System.Text -Debug:$false
}

$count = [UInt64]0
$luidPtr = [IntPtr]::Zero

[uint64]$ntStatus = [Win32.Secure32]::LsaEnumerateLogonSessions( [ref]$count , [ref]$luidPtr )

$count = [UInt64]0
$luidPtr = [IntPtr]::Zero

[uint64]$ntStatus = [Win32.Secure32]::LsaEnumerateLogonSessions( [ref]$count , [ref]$luidPtr )

if( $ntStatus )
{
    Write-Error "LsaEnumerateLogonSessions failed with error $ntStatus"
}
elseif( ! $count )
{
    Write-Error "No sessions returned by LsaEnumerateLogonSessions"
}
elseif( $luidPtr -eq [IntPtr]::Zero )
{
    Write-Error "No buffer returned by LsaEnumerateLogonSessions"
}
else
{   
    Write-Debug "$count sessions retrieved from LSASS"
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

    Write-Debug "Found $(if( $lsaSessions ) { $lsaSessions.Count } else { 0 }) LSA sessions for $UserDomain\$Username, earliest session $(if( $earliestSession ) { Get-Date $earliestSession -Format G } else { 'never' })"
}

if( $lsaSessions -and $lsaSessions.Count )
{
    ## get all logon ids for logons that happened at the same time
    [array]$loginIds = @( $lsaSessions | Where-Object { $_.LoginTime -eq $lsaSessions[0].LoginTime } | Select-Object -ExpandProperty LoginId )
    if( ! $loginIds -or ! $loginIds.Count )
    {
        Write-Error "Found no login ids for $username at $(Get-Date -Date $lsaSessions[0].LoginTime -Format G)"
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
    }
}
else
{
    Throw "Failed to retrieve logon session for $UserDomain\$Username from LSASS"
}

#endregion

Get-FSLogixProfileEvents -Start $logon.LogonTime -Username $logon.UserName
