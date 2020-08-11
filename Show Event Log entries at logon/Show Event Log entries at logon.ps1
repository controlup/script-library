<#
    Find all event log entries around a user's logon

    @guyrleech 19/06/2020 - based on https://github.com/guyrleech/Microsoft/blob/master/event%20aggregator.ps1
#>

[Cmdletbinding()]

Param
(
    [Parameter(Mandatory)]
    [int]$sessionId ,
    [Parameter(Mandatory)]
    [string]$domainUsername ,
    [int]$secondsBeforeLogon = 15 ,
    [int]$secondsAfterLogon = 120 ,
    [string]$badOnly = 'no' ,
    [int]$horizontalResolution = 1080
)

$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

[int]$outputWidth = $horizontalResolution / 4.1 ## adjust this to adjust the width of the event message text column
##[int]$windowWidth = 260

# Altering the size of the PS Buffer
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    Write-Verbose -Message "Current width is $($WideDimensions.Width)"
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

## Get logon time for this user

## https://www.codeproject.com/Articles/18179/Using-the-Local-Security-Authority-to-Enumerate-Us
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

[string]$userDomain , [string]$username = $domainUsername -split '\\' , 2

if( [string]::IsNullOrEmpty( $username ) )
{
    Throw 'Must specify user $domainUsername as domain\username'
}

## from Analyze Logon Durations Script
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

    Write-Verbose "Found $(if( $lsaSessions ) { $lsaSessions.Count } else { 0 }) LSA sessions for $UserDomain\$Username, earliest session $(if( $earliestSession ) { Get-Date $earliestSession -Format G } else { 'never' })"
}

$userLoginTime = $null

if( $lsaSessions -and $lsaSessions.Count )
{
    $userLoginTime = $lsaSessions[0].LoginTime
}
else
{
    Throw "No logons found for user $domainUsername in session $sessionId"
}

Write-Verbose "Login time is $(Get-Date -Date $userloginTime -Format G)"

[hashtable]$eventFilter =  @{ starttime = $userLoginTime.AddSeconds( -$secondsBeforeLogon ) ; endtime = $userLoginTime.AddSeconds( $secondsAfterLogon ) }

if( $badOnly -imatch '^y' -or $badOnly -imatch 'true' )
{
    $eventFilter.Add( 'Level' , @( 1 , 2 , 3 ) )
}

$getWinEvent = Get-Command -Name Get-WinEvent

[array]$results = @( . $getWinEvent -ListLog * -Verbose:$false | Where-Object RecordCount -gt 0 | . { Process { . $getWinEvent -ErrorAction SilentlyContinue -Verbose:$False -FilterHashtable ( @{ logname = $_.logname } + $eventFilter ) }} `
            | Select-Object -ExcludeProperty ContainerLog,MachineName,ProcessId,ThreadId,UserId,RecordId,ProviderId,*ActivityId,Version,Qualifiers,Level,Task,OpCode,Keywords,Bookmark,*Ids,Properties -Property *,@{n='User';e={if( $_.UserId ) { ([System.Security.Principal.SecurityIdentifier]($_.UserId)).Translate([System.Security.Principal.NTAccount]).Value }}},@{n='Level';e={$_.LevelDisplayName}} | Sort-Object -Property TimeCreated )

if( $results -and $results.Count )
{
    Write-Output -InputObject "Found $($results.Count) events between $(Get-Date -Date $eventFilter.StartTime -Format G) and $(Get-Date -Date $eventFilter.EndTime -Format G)"
    ## get column positions of output so we can figure where we need to truncate the message
    [array]$outputFields = @( @{ Name = 'Time' ; Expression = {Get-Date -Date $_.TimeCreated -Format T}},'Id','Level',@{ Name = 'Event Log'; Expression = {($_.LogName -Split '/')[0]}},'Message' )
<#
    ## Out-String flattens into a single string so we use split to turn back into lines and skip the first as it is blank
    [string]$outputHeadings = ($results | Format-Table -Property $outputFields | Out-String ) -split '\r?\n' | Select-Object -First 1 -Skip 1
    ## Problem with CU SBA output window is that it thinks it is 80 columns wide no mater what you set the output width to
    [int]$messageWidth = $windowWidth - $outputHeadings.IndexOf( 'Message' )
    ## we don't have room for much in the SBA output window so pick what is the most important, get rid of date and truncate the message
    [array]$outputFields = @( @{ Name = 'Time' ; Expression = {Get-Date -Date $_.TimeCreated -Format T}},'Id','Level',@{ Name = 'Event Log'; Expression = {($_.LogName -Split '/')[0]}},@{ Name ='Text'; Expression = { $(if( $_.Message.Length -gt $messageWidth ) { $_.Message.SubString( 0 , $messageWidth ) } else { $_.Message } ) }} )
#>  
    ## -Wrap only works well when run in CU when the window output width is less than or equal to the physical window width otherwise wraps onto next line(s)
    $results | Format-Table -Property $outputFields -Wrap -AutoSize
}
else
{
    Write-Warning -Message "No events found between $(Get-Date -Date $eventFilter.StartTime -Format G) and $(Get-Date -Date $eventFilter.EndTime -Format G)"
}

