#required -version 3
<#
    Health check VMware App Volumes end-point

    @guyrleech 2020

    Modification History:

    @guyrleech 27/08/2020  Added code for edge case where SQL down when user logs in
    @guyrleech 22/10/2020  Added call to /health_check and /images
    @guyrleech 23/10/2020  Made health_check check compatible with App Volumes v2.18 (and v4.2)
#>

[CmdletBinding()]

Param
(
    [double]$lastDays = 0 ,
    [string]$serviceName = 'svservice' ,
    [string]$serviceProcessName = 'svservice' ,
    [string]$fullServiceName = 'App Volumes Service' ,
    [string]$driverName = 'svdriver' ,
    [string]$productName = 'App Volumes' ,
    [string]$configKeyName = 'HKLM:\SOFTWARE\WOW6432Node\CloudVolumes\Agent' ,
    [string]$mountPoint = '\\SnapVolumesTemp\\' ,
    [int]$outputWidth = 400
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

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

Add-Type -ErrorAction Stop -TypeDefinition @'
    using System;
    using System.Runtime.InteropServices;
    public enum WTS_CONNECTSTATE_CLASS
    {
        WTSActive,
        WTSConnected,
        WTSConnectQuery,
        WTSShadow,
        WTSDisconnected,
        WTSIdle,
        WTSListen,
        WTSReset,
        WTSDown,
        WTSInit
    }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct WTSINFOEX_LEVEL1_W {
        public Int32                  SessionId;
        public WTS_CONNECTSTATE_CLASS SessionState;
        public Int32                   SessionFlags; // 0 = locked, 1 = unlocked , ffffffff = unknown
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 33)]
        public string WinStationName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 21)]
        public string UserName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 18)]
        public string DomainName;
        public UInt64           LogonTime;
        public UInt64           ConnectTime;
        public UInt64           DisconnectTime;
        public UInt64           LastInputTime;
        public UInt64           CurrentTime;
        public Int32            IncomingBytes;
        public Int32            OutgoingBytes;
        public Int32            IncomingFrames;
        public Int32            OutgoingFrames;
        public Int32            IncomingCompressedBytes;
        public Int32            OutgoingCompressedBytes;
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct WTS_SESSION_INFO
    {
        public Int32 SessionID;

        [MarshalAs(UnmanagedType.LPStr)]
        public String pWinStationName;

        public WTS_CONNECTSTATE_CLASS State;
    }
    [StructLayout(LayoutKind.Explicit)]
    public struct WTSINFOEX_LEVEL_W
    { //Union
        [FieldOffset(0)]
        public WTSINFOEX_LEVEL1_W WTSInfoExLevel1;
    } 
    [StructLayout(LayoutKind.Sequential)]
    public struct WTSINFOEX
    {
        public Int32 Level ;
        public WTSINFOEX_LEVEL_W Data;
    }
    public enum WTS_INFO_CLASS
    {
        WTSInitialProgram,
        WTSApplicationName,
        WTSWorkingDirectory,
        WTSOEMId,
        WTSSessionId,
        WTSUserName,
        WTSWinStationName,
        WTSDomainName,
        WTSConnectState,
        WTSClientBuildNumber,
        WTSClientName,
        WTSClientDirectory,
        WTSClientProductId,
        WTSClientHardwareId,
        WTSClientAddress,
        WTSClientDisplay,
        WTSClientProtocolType,
        WTSIdleTime,
        WTSLogonTime,
        WTSIncomingBytes,
        WTSOutgoingBytes,
        WTSIncomingFrames,
        WTSOutgoingFrames,
        WTSClientInfo,
        WTSSessionInfo,
        WTSSessionInfoEx,
        WTSConfigInfo,
        WTSValidationInfo,   // Info Class value used to fetch Validation Information through the WTSQuerySessionInformation
        WTSSessionAddressV4,
        WTSIsRemoteSession
    }
    public static class wtsapi
    {
        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern int WTSQuerySessionInformationW(
                 System.IntPtr hServer,
                 int SessionId,
                 int WTSInfoClass ,
                 ref System.IntPtr ppSessionInfo,
                 ref int pBytesReturned );

        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern int WTSEnumerateSessions(
                 System.IntPtr hServer,
                 int Reserved,
                 int Version,
                 ref System.IntPtr ppSessionInfo,
                 ref int pCount);

        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern IntPtr WTSOpenServer(string pServerName);
        
        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern void WTSCloseServer(IntPtr hServer);
        
        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern void WTSFreeMemory(IntPtr pMemory);
    }
'@ 

Function Get-WTSSessionInformation
{
    [cmdletbinding()]

    Param
    (
        [string[]]$computers = @( $null )
    )

    [long]$count = 0
    [IntPtr]$ppSessionInfo = 0
    [IntPtr]$ppQueryInfo = 0
    [long]$ppBytesReturned = 0
    $wtsSessionInfo = New-Object -TypeName 'WTS_SESSION_INFO'
    $wtsInfoEx = New-Object -TypeName 'WTSINFOEX'
    [int]$datasize = [system.runtime.interopservices.marshal]::SizeOf( [Type]$wtsSessionInfo.GetType() )

    ForEach( $computer in $computers )
    {
        [string]$machineName = $(if( $computer ) { $computer } else { $env:COMPUTERNAME })
        [IntPtr]$serverHandle = [wtsapi]::WTSOpenServer( $computer )

        ## If the function fails, it returns a handle that is not valid. You can test the validity of the handle by using it in another function call.

        [long]$retval = [wtsapi]::WTSEnumerateSessions( $serverHandle , 0 , 1 , [ref]$ppSessionInfo , [ref]$count );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

        if ($retval -ne 0)
        {
             for ([int]$index = 0; $index -lt $count; $index++)
             {
                 $element = [system.runtime.interopservices.marshal]::PtrToStructure( [long]$ppSessionInfo + ($datasize * $index), [type]$wtsSessionInfo.GetType())
                 if( $element -and $element.SessionID -ne 0 ) ## session 0 is non-interactive (session zero isolation)
                 {
                     $retval = [wtsapi]::WTSQuerySessionInformationW( $serverHandle , $element.SessionID , [WTS_INFO_CLASS]::WTSSessionInfoEx , [ref]$ppQueryInfo , [ref]$ppBytesReturned );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                     if( $retval -and $ppQueryInfo )
                     {
                        $value = [system.runtime.interopservices.marshal]::PtrToStructure( $ppQueryInfo , [Type]$wtsInfoEx.GetType())
                        if( $value -and $value.Data -and $value.Data.WTSInfoExLevel1.SessionState -ne [WTS_CONNECTSTATE_CLASS]::WTSListen -and $value.Data.WTSInfoExLevel1.SessionState -ne [WTS_CONNECTSTATE_CLASS]::WTSConnected )
                        {
                            $wtsinfo = $value.Data.WTSInfoExLevel1
                            $idleTime = New-TimeSpan -End ([datetime]::FromFileTimeUtc($wtsinfo.CurrentTime)) -Start ([datetime]::FromFileTimeUtc($wtsinfo.LastInputTime))
                            Add-Member -InputObject $wtsinfo -Force -NotePropertyMembers @{
                                'IdleTimeInSeconds' =  $idleTime | Select -ExpandProperty TotalSeconds
                                'IdleTimeInMinutes' =  $idleTime | Select -ExpandProperty TotalMinutes
                                'Computer' = $machineName
                                'LogonTime' = [datetime]::FromFileTime( $wtsinfo.LogonTime )
                                'DisconnectTime' = [datetime]::FromFileTime( $wtsinfo.DisconnectTime )
                                'LastInputTime' = [datetime]::FromFileTime( $wtsinfo.LastInputTime )
                                'ConnectTime' = [datetime]::FromFileTime( $wtsinfo.ConnectTime )
                                'CurrentTime' = [datetime]::FromFileTime( $wtsinfo.CurrentTime )
                            }
                            $wtsinfo
                        }
                        [wtsapi]::WTSFreeMemory( $ppQueryInfo )
                        $ppQueryInfo = [IntPtr]::Zero
                     }
                     else
                     {
                        Write-Error "$($machineName): $LastError"
                     }
                 }
             }
        }
        else
        {
            Write-Error "$($machineName): $LastError"
        }

        if( $ppSessionInfo -ne [IntPtr]::Zero )
        {
            [wtsapi]::WTSFreeMemory( $ppSessionInfo )
            $ppSessionInfo = [IntPtr]::Zero
        }
        [wtsapi]::WTSCloseServer( $serverHandle )
        $serverHandle = [IntPtr]::Zero
    }
}

[datetime]$startDate = Get-Date -Date "01/01/2000"

if( $PSBoundParameters[ 'lastDays' ] )
{
    $startDate = (Get-Date).AddDays( -$lastDays )
    Write-Output "Checking from $(Get-Date -Date $startDate -Format G)"
}

# Altering the size of the PS Buffer
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ($WideDimensions = $PSWindow.BufferSize) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

$warnings = New-Object -TypeName System.Collections.Generic.List[string]
$information = New-Object -TypeName System.Collections.Generic.List[psobject]

[string]$appVolumesVersion = $null
[datetime]$installDate = New-Object -TypeName DateTime

if( ( $appvolumesKey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -Name DisplayName -ErrorAction SilentlyContinue | Where-Object DisplayName -match $productName | Select-Object -ExpandProperty PSPath )  `
    -or ( $appvolumesKey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' -Name DisplayName -ErrorAction SilentlyContinue | Where-Object DisplayName -match $productName | Select-Object -ExpandProperty PSPath ) )
{
    $appVolumesVersion = Get-ItemProperty -Path $appvolumesKey -Name DisplayVersion | Sort-Object -Descending -Property DisplayVersion | Select-object -ExpandProperty DisplayVersion -First 1
    if( ( [string]$installedOn =  Get-ItemProperty -Path $appvolumesKey -Name InstallDate -ErrorAction SilentlyContinue | Select-Object -ExpandProperty InstallDate ) `
        -and [datetime]::TryParseExact( $installedOn , 'yyyyMMdd' , [System.Globalization.CultureInfo]::InvariantCulture , [System.Globalization.DateTimeStyles]::None , [ref]$installdate ) )
    {
        $information.Add( ([pscustomobject]@{ 'Item' = "$productName Installed On" ; 'Description' = "$(Get-Date -Date $installDate -Format d)" } ) )
    }
    else
    {
        $warnings.Add( "Unable to decode installation date of $installedOn" )
    }
}
else
{
    $warnings.Add( "Unable to find installation of $productName in registry" )
}

if( ! [string]::IsNullOrEmpty( $appVolumesVersion ) )
{
    $information.Add( ([pscustomobject]@{ 'Item' = "$productName Installed Version" ; 'Description' = $appVolumesVersion } ) )
}
else
{
    $warnings.Add( "Unable to find installation details for $productName in the registry" )
}

# Check service running and when started wrt boot
if( ! ( $svservice = Get-CimInstance -ClassName win32_service -filter "name = '$serviceName'" -ErrorAction SilentlyContinue ) )
{
    $warnings.Add( "Unable to find $productName service $svservice" )
}
else
{
    if( $svservice.State -ne 'Running' )
    {
        $warnings.Add( "svservice is not running, it is $($svservice.State)" )
    }

    if( $svservice.StartMode -ne 'Auto' )
    {
        $warnings.Add( "svservice is not set to auto start, it is set to $($svservice.StartMode)" )
    }

    ## get service restart options which can't do natively in PowerShell/CIM
    [bool]$seenFailureActionsLabel = $false
    $restartAfterFailures = New-Object -TypeName System.Collections.Generic.List[int]
    [int]$failureNumber = 1

    <#
    [SC] QueryServiceConfig2 SUCCESS

    SERVICE_NAME: svservice
            RESET_PERIOD (in seconds)    : 43200
            REBOOT_MESSAGE               :
            COMMAND_LINE                 :
            FAILURE_ACTIONS              : RESTART -- Delay = 30000 milliseconds.
                                           RESTART -- Delay = 30000 milliseconds.
                                           RESTART -- Delay = 60000 milliseconds.
    #>
    ## technically only need to skip any lines that don't have a : in them like "[SC] QueryServiceConfig2 SUCCESS"
    sc.exe qfailure $serviceName | Select-Object -Skip 6 | ForEach-Object `
    {
        if( $_ -match 'FAILURE_ACTIONS\s*:\s*' )
        {
            if( $_ -cmatch 'RESTART' )
            {
                $restartAfterFailures.Add( $failureNumber )
            }
            $seenFailureActionsLabel = $true
        }
        elseif( $seenFailureActionsLabel -or $_ -match '^[^:]*$' ) ## if first failure action not set then the FAILURE_ACTIONS tag is not present
        {
            $failureNumber++
            $seenFailureActionsLabel = $true
            if( $_ -cmatch 'RESTART' )
            {
                $restartAfterFailures.Add( $failureNumber )
            }
        }
    }

    if( $restartAfterFailures -and $restartAfterFailures.Count )
    {
        Write-Output -InputObject "Service is set to restart after failures $(($restartAfterFailures|Sort-Object) -join ', ')"
        if( ! $restartAfterFailures.Contains( [int]1 ) )
        {
            $warnings.Add( "Restart service recovery option not set for first failure" )
        }
            
    }
    else
    {
        $warnings.Add( "Restart service recovery options are not set" )
    }

    ## can be more than one process, eg. if a dialogue box is being shown
    if( ! ( $svserviceProcess = Get-Process -Name ( $serviceProcessName -replace '\.exe$') -ErrorAction SilentlyContinue | Select-Object -First 1 ))
    {
        $warnings.Add( "Unable to find running svservice process" )
    }
    else
    {
        if( $exeVersion = Get-ItemProperty -ErrorAction SilentlyContinue -Path $svserviceProcess.Path )
        {
            if( $exeVersion.PSObject.Properties[ 'VersionInfo' ] -and $exeVersion.VersionInfo.PSObject.Properties[ 'ProductVersion' ] )
            {
                $information.Add( ([pscustomobject]@{ 'Item' = "$productName running service executable version" ; 'Description' = $exeVersion.VersionInfo.ProductVersion } ) )
            }
            else
            {
                $warnings.Add( "Running service executable $($svServiceProcess.Path) does not contain version information" )
            }
        }
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        ## TODO figure out how we get the real boot time as this comes back as the parent VM not when we were cloned
        <#
            Event Id 1 source Kernel-General

            The system time has changed to ‎2020‎-‎06‎-‎24T15:57:30.511000000Z from ‎2020‎-‎06‎-‎24T12:38:09.648303500Z.

            Change Reason: An application or system component changed the time.
            Process: '\Device\HarddiskVolume4\Program Files\VMware\VMware Tools\vmtoolsd.exe' (PID 4432).
        #>

        ##Write-Output "svservice process pid $($svserviceProcess.Id) started at $(Get-Date -Format G -Date $svserviceProcess.StartTime), $(($svserviceProcess.StartTime - $os.LastBootUpTime).TotalMinutes) minutes after boot at $(Get-Date -Date $os.LastBootUpTime -Format G)"
    }

    ## check event log set correctly because if not event messages display "The description for Event ID 218 from source svservice cannot be found"
    if( $messageFileValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\Application\$serviceName" -Name 'EventMessageFile' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'EventMessageFile' -ErrorAction SilentlyContinue )
    {
        [string]$messageFile = [System.Environment]::ExpandEnvironmentVariables( ( $messageFileValue -replace '"' ) )
        if( $messageFile -ne ( $svservice.PathName -replace '"') )
        {
            # simple size check to see if someone has applied a workaround by copying the real svservice.exe to the bad location
            if( ! ( $messageFileProperties = (Get-ItemProperty -Path $messageFile -ErrorAction SilentlyContinue) ) -or ! ( $goodFileProperties = (Get-ItemProperty -Path ($svservice.PathName -replace '"') -ErrorAction SilentlyContinue) ) -or $messageFileProperties.Length -ne $goodFileProperties.Length )
            {
                $warnings.Add( "Event log message file set to `"$messageFileValue`" not $($svservice.PathName) - event log messages will not display correctly" )
            }
        }
    }
    else
    {
        $warnings.Add( "Failed to find Application event log setting for $serviceName - event log messages will not display correctly" )
    }
}

## in case we have certificate problems - probably should check first for better security !

##https://stackoverflow.com/questions/41897114/unexpected-error-occurred-running-a-simple-unauthorized-rest-query?rq=1
Add-Type -TypeDefinition @'
public class SSLHandler
{
    public static System.Net.Security.RemoteCertificateValidationCallback GetSSLHandler()
    {
        return new System.Net.Security.RemoteCertificateValidationCallback((sender, certificate, chain, policyErrors) => { return true; });
    }
}
'@

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = [SSLHandler]::GetSSLHandler()
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Get server details and check connectivity
if( $configKey = Get-ItemProperty -Path $configKeyName -ErrorAction SilentlyContinue )
{
    if( ! $configKey.PSObject.Properties[ 'Manager_Address' ] )
    {
        $warnings.Add( "Registry value Manager_Address does not exist in $configKeyName" )
    }
    elseif( ! $configKey.PSObject.Properties[ 'Manager_Port' ] )
    {
        $warnings.Add( "Registry value Manager_Port does not exist in $configKeyName" )
    }
    else
    {
        $information.Add( ([pscustomobject]@{ 'Item' = "$productName server" ; 'Description' =  "$($configKey.Manager_Address):$($configKey.Manager_Port)" } ) )
        $networkTestResult = Test-NetConnection -ComputerName $configKey.Manager_Address -Port $configKey.Manager_Port -InformationLevel Detailed -ErrorAction SilentlyContinue

        $information.Add( ([pscustomobject]@{ 'Item' = "Test Connection to $productName Server" ; 'Description' = $(if( ! $networkTestResult -or ! $networkTestResult.TcpTestSucceeded ) { 'FAILED' } else { 'OK' } ) } ) )
        
        if( ! $networkTestResult -or ! $networkTestResult.TcpTestSucceeded )
        {
           [string]$message = "Failed to connect to $($configKey.Manager_Address) on port $($configKey.Manager_Port)"
            if( ! $networkTestResult.NameResolutionSucceeded )
            {
                $message += ' (name resolution failed)'
            }
            $warnings.Add( $message )
        }

        ## even though we failed to connect we'll try the health check URLs

        $newKey = $null
        [string]$ieKey = 'HKCU:\SOFTWARE\Microsoft\Internet Explorer\Main'

        ## if we don't set the first run registry value, we get this error but using -UseBasicParsing doesn't give us structured HTML "The response content cannot be parsed because the Internet Explorer engine is not available, or Internet Explorer's first-launch configuration is not complete."
        if( ! ( $existingFirstRun = Get-ItemProperty -Path $ieKey -Name DisableFirstRunCustomize -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DisableFirstRunCustomize ) )
        {
            if( ! ( Test-Path -Path $ieKey -ErrorAction SilentlyContinue ) )
            {
                $newKey = New-Item -Path $ieKey -Force
            }
            Set-ItemProperty -Path $ieKey -Name DisableFirstRunCustomize -Value 1
        }

        [string]$healthCheckResult = 'OK'
        $exception = $null
        [string]$healthURL = ('http{0}://{1}:{2}/health_check' -f $(if( $configKey.Manager_Port -ne 80 ) { 's' } ) , $configKey.Manager_Address , $configKey.Manager_Port )
        try
        {
            $health = Invoke-WebRequest -Uri $healthURL
        }
        catch
        {
            $exception = $_
            $health = $null
        }

        if( $health )
        {
            [string]$goodHealthRegex = $(if( $appVolumesVersion -match '^2\.' )
            {
                ## Process: 12316 - IP: 192.168.0.124 - Threads: 7 - Requests: 3 - Uptime: 652.298692s - Objects: 798890 - Browser: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36 Edg/86.0.622.43
                'Process:.*IP:.*Threads:.*Requests:.*Uptime:.*Objects:'
            }
            else
            {
                ## OK <GUID>
                '^OK\s'
            })
            if( $health.StatusCode -ne 200 -or $health.Content -notmatch $goodHealthRegex )
            {
                [string]$errorText = $null
                [hashtable]$tags = @{}
                ForEach( $element in $health.AllElements )
                {
                    try
                    {
                        $tags.Add( $element.tagname , $element.innerText )
                    }
                    catch
                    {
                        ## only expecting 1 element of each so ignore for now
                    }
                }
                    
                <#
                    <h2>Startup Failure</h2>
                    <h3>Unable to start App Volumes Manager</h3>
                    <p>
                        42000 (4060) [Microsoft][ODBC SQL Server Driver][SQL Server]Cannot open database "App_Volumes" requested by the login. The login failed.
                    </p>
                #>
                ForEach( $tagname in @( 'h2' , 'h3' , 'p' ) )
                {
                    if( ! [string]::IsNullOrEmpty( ( $tag = $tags[ $tagname ] ) ) )
                    {
                        if( $errorText )
                        {
                            $errorText += ": $tag"
                        }
                        else
                        {
                            $errorText = $tag
                        }
                    }
                }
                [string]$warningText = $null
                if( $health.StatusCode -eq 200 )
                {
                    $warningText = "Unexpected response"
                }
                else
                {
                    $warningText = "Error $($health.StatusCode)"
                }
                $warnings.Add( "$warningText from health check URL $healthURL - content `"$errorText`"" )
                $healthCheckResult = $warningText
            }
            else
            {
                ## https://childebrandt42.wordpress.com/2020/10/15/appvolumes-monitoring-with-http-health-monitor/
                ## now check /images
                [string]$imagesURL = ('http{0}://{1}:{2}/images' -f $(if( $configKey.Manager_Port -ne 80 ) { 's' } ) , $configKey.Manager_Address , $configKey.Manager_Port )
                try
                {
                    $images = Invoke-WebRequest -UseBasicParsing -Uri $imagesURL -ErrorAction SilentlyContinue
                }
                catch
                {
                    if( $_.Exception.message -match '\b404\b' )
                    {
                        $information.Add( ([pscustomobject]@{ 'Item' = "Calls to health check URLs" ; 'Description' = 'OK' } ) )
                    }
                    else
                    {
                        $warnings.Add( "Failed call to second health check URL $imagesURL - error $($_.Exception|Select-Object -ExpandProperty message)" )
                    }
                    $images = $null
                }

                if( $images ) ## this is actually bad
                {
                    if( $images.StatusCode -ne 404 ) ## expected
                    {
                        $warnings.Add( "Unexpected status code $($images.StatusCode) from /images URL $imagesURL - `"$($images.Content)`"" )
                    }
                }
            }
        }
        else
        {
            $warnings.Add( "Failed call to health check URL $healthURL - $exception" )
            $healthCheckResult = 'FAILED'
        }
            
        $information.Add( ([pscustomobject]@{ 'Item' = "Response from $productName health check URL" ; 'Description' = $healthCheckResult } ) )
        if( $null -eq $existingFirstRun )
        {
            if( $newKey ) ## we created the key so delete it
            {
                Remove-Item -Path $ieKey -Recurse -Force
            }
            else
            {
                Remove-ItemProperty -Path $ieKey -Name DisableFirstRunCustomize -Force
            }
        }
        elseif( $existingFirstRun -eq 0 )
        {
            Set-ItemProperty -Path $ieKey -Name DisableFirstRunCustomize -Value 0
        }
    }
}
else
{
    $warnings.Add( "Failed to read configuration key $configKeyName" )
}

if( ! [string]::IsNullOrEmpty( $driverName ) )
{
    if( ! ($driver = Get-CimInstance -ClassName Win32_SystemDriver -Filter "Name = '$driverName'" -ErrorAction SilentlyContinue ) )
    {
        $warnings.Add( "Unable to find instance of driver $drivername" )
    }
    elseif( $driver.State -ne 'Running' )
    {
        $warnings.Add( "Driver $drivername is $($driver.State) but should be running" )
    }
    elseif( $driver.Started -ne 'True' )
    {
        $warnings.Add( "Driver $drivername is not started" )
    }
    elseif( $driver.Status -ne 'OK' )
    {
        $warnings.Add( "Driver $drivername  status is $($driver.Status) but should be OK" )
    }
    else
    {
        $information.Add( ([pscustomobject]@{ 'Item' = "Driver $driverName" ; 'Description' = 'OK' } ) )
    }
}

# Check event logs for service failures & app crashes
if( ( [array]$crashes = @( Get-WinEvent -FilterHashtable @{ ProviderName = 'Service Control Manager' ; id = 7034 ; Data = $fullServiceName ; StartTime = $startDate } -ErrorAction SilentlyContinue) ) -and $crashes.Count )
{
    [string]$message = "There have been $($crashes.Count) service crashes. Latest @ $(Get-Date -Date $crashes[0].TimeCreated -Format G)"
    if( $crashes.Count -gt 1 )
    {
        $message += ". Oldest @ $(Get-Date -Date $crashes[-1].TimeCreated -Format G)"
    }
    if( $firstEvent = Get-WinEvent -LogName $crashes[0].ContainerLog -Oldest -MaxEvents 1 )
    {
        $message += ". First event in $($crashes[0].ContainerLog) event log @ $(Get-Date -Date $firstEvent.TimeCreated -Format G) ($([math]::Round(([datetime]::Now - $firstEvent.TimeCreated).TotalDays / 7 , 1 )) weeks ago)"
    }
    $warnings.Add(  $message )
}

$oldestEvent = $null

## find oldest event in log containing events so we can report this
if( ! ( $provider = Get-WinEvent -ListProvider $serviceName -ErrorAction SilentlyContinue ) )
{
    $warnings.Add(  "No event provider $serviceName found" )
}
else
{
    ## Find when the first event log entry in this log is
    ForEach( $eventLog in $provider.LogLinks )
    {
        if( $thisOldestEvent = Get-WinEvent -LogName $eventLog.LogName -MaxEvents 1 -Oldest -ErrorAction SilentlyContinue )
        {
            if( ! $oldestEvent -or $oldestEvent -lt $thisOldestEvent )
            {
                $oldestEvent = $thisOldestEvent ## slightly simplistic as event logs could've been cleared at different times but unlikely that writes to more than one event log anyway
            }
        }
        else
        {
            $warnings.Add(  "There are no events at all in the $($eventLog.DisplayName) event log which $serviceName writes to" )
        }
    }
}

$badEvents = New-Object -TypeName System.Collections.Generic.List[psobject]

# Check for svservice specific issues in event log - get relevant events into an array, oldest first so we can grab the first event and not have to work back when searching
[array]$svserviceEventLogEntries = $null
try
{
    $svserviceEventLogEntries = @( Get-WinEvent -FilterHashtable @{ ProviderName = $serviceName ; StartTime = $startDate } -Oldest -ErrorAction SilentlyContinue )
}
catch
{
}

if( ! $svserviceEventLogEntries -or ! $svserviceEventLogEntries.Count )
{
    $warnings.Add(  "No eventlog entries found for the event log provider $serviceName - this is unusual" )
    if( $oldestEvent )
    {
        $warnings.Add(  "Oldest event in the $($eventLog.DisplayName) event log which $serviceName writes to is $(Get-Date -Date $oldestEvent.TimeCreated -Format G) ($([math]::Round(([datetime]::Now - $oldestEvent.TimeCreated).TotalDays / 7 , 1 )) weeks ago)" )
    }
}
else ## look for bad event log entries
{
    ## TODO what errors are there? Should we just check for "error"?
    [int[]]$badLevels = @( 1 , 2, 3 )
    [array]$avErrors = @( ($svserviceEventLogEntries).Where( { ( $_.id -eq '240' -and $_.Properties[1].Value -match 'Connection Error' -and $_.Properties[1].Value -notmatch 'Sending user popup' ) -or $_.Level -in $badLevels } )|select TimeCreated,@{n='Text';e={$_.Properties[1].value}},Id ,Message)
    if( $avErrors -and $avErrors.Count )
    {
        #$warnings.Add( $message )
        ForEach( $avError in $avErrors )
        {
            if( ( [string]::IsNullOrEmpty( ( [string]$errorText = ($avError.Text -replace '\r?\n' , ' ' -replace '\s+' , ' ').Trim() ) ) ) `
                -and ! [string]::IsNullOrEmpty( $avError.Message ) )
            {
                $errorText = $avError.Message -replace '\r?\n' , ' ' -replace 'Details:' -replace '\s+' , ' '
            }
            $badEvents.Add( ( [pscustomobject][ordered]@{
                'Time' = $avError.TimeCreated
                'Id' = $avError.Id
                'Error' = $errorText
            } ) )
        }
    }
}

# Check mount points
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

        if( ! $ntStatus -and $sessionData -ne [IntPtr]::Zero `
            -and ($data = [System.Runtime.InteropServices.Marshal]::PtrToStructure( $sessionData , [type][Win32.Secure32+SECURITY_LOGON_SESSION_DATA] ) ) `
                -and $data.PSiD -ne [IntPtr]::Zero `
                    -and ( $sid = New-Object -TypeName System.Security.Principal.SecurityIdentifier -ArgumentList $Data.PSiD ) )
        {
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
            if( $secType -match 'Interactive' -and $thisDomain -ne 'Window Manager' -and $thisDomain -ne 'Font Driver Host' )
            {
                $authPackage = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.AuthenticationPackage.buffer) #get the authentication package
                $session = $data.Session # get the session number
                $logonServer = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.LogonServer.buffer) #get the logon server
                $DnsDomainName = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.DnsDomainName.buffer) #get the DNS Domain Name
                $upn = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.upn.buffer) #get the User Principal Name

                [pscustomobject]@{
                    'Sid' = $sid
                    'Username' = $thisUser
                    'Domain' = $thisDomain
                    'SessionId' = $session
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
            [void][Win32.Secure32]::LsaFreeReturnBuffer( $sessionData )
            $sessionData = [IntPtr]::Zero
        }
        $iter = $iter.ToInt64() + [System.Runtime.InteropServices.Marshal]::SizeOf([type][Win32.Secure32+LUID])  # move to next pointer
    }) | Sort-Object -Descending -Property 'LoginTime'

    [void]([Win32.Secure32]::LsaFreeReturnBuffer( $luidPtr ))
    $luidPtr = [IntPtr]::Zero

    Write-Debug "Found $(if( $lsaSessions ) { $lsaSessions.Count } else { 0 }) LSA sessions, earliest session $(if( $earliestSession ) { Get-Date $earliestSession -Format G } else { 'never' })"
}

## Borrowed from Analyze Logon Durations - so we can map the various GUIDs to app names

$AppList = New-Object -TypeName System.Collections.Generic.List[psobject]
[string]$username = $null
[string]$domainname = $null

if( $svserviceEventLogEntries -and $svserviceEventLogEntries.Count )
{
    $svserviceEventLogEntries.Where( { $_.Id -eq 218 } ) | . { Process { $username = $domainname = $null ; $event = $_ ; $_.properties[1].Value -split  "`r`n" } } | . { Process {
        ## multiple lines of MOUNTED-READ;External SSD\appvolumes\packages\PuTTY.vmdk;{b2a70e3f-90ef-45b5-87a0-b7d07402a977}
        if( $_  -match '\bLOGIN\s*(.*)\\(.*)$' ) {
            $username   = $matches[2]
            $domainname = $Matches[1]
        }
        elseif( $_ -match '(MOUNTED.*);(.+);({.+})' ) {
            if( ! $username -or ! $domainname )
            {
                Write-Debug -Message "Failed to get username or domain for $_"
            }
            $AppList.Add( [pscustomobject]@{
                'Time'       = $event.TimeCreated
                'UserName'   = $username
                'DomainName' = $domainname
                'MountType'  = $matches[1]
                'AppPath'    = $matches[2]
                'AppGUID'    = $matches[3] 
                'AppName'    = $matches[2].Split( '\' )[-1].Replace( '.vmdk' , '' ).Replace( '!20!' , ' ').Replace( '!2B!' , '+' ) 
                'AppId'      = $null } )
        }
        elseif( $_ -match 'ENABLE-APP;({.+});({.+})' ) ## ENABLE-APP;{65950e61-0304-45a6-b354-f3b26ced3f64};{b2a70e3f-90ef-45b5-87a0-b7d07402a977}
        {
            if( $appObject = $AppList | Where-Object -Property AppGuid -eq $Matches[2] | Select-Object -First 1 )
            {
                $appObject.AppId = $Matches[1]
            }
            else
            {
                Write-Debug "Couldn't find appguid $($matches[2]) in AppList"
            }
        }
        elseif( $_ -match '\bResponse\b.*!FAILURE!' )
        {
            $badEvents.Add( ([pscustomobject]@{ 'Time' = $event.TimeCreated ; 'Id' = $event.Id ; Error = "Found FAILURE response event" }) )
        }
        elseif( $_ -match '\bManager:.*\bStatus:\s+([45]\d\d)' ) ## Manager: grl-appvolv2.guyrleech.local, Status: 500
        {
            [int]$errorCode = $Matches[1]
            [string]$errorText = $null
            if( ( $event.Properties[1].Value  -replace  "`r`n|\<[a-z]+\/\>", ' ' -replace '\s+' , ' ' ) -match '\bResponse\b\s+\(\d+ bytes\):\s+(.*)' ) ## Response (88 bytes): This request is taking too long.<br/>... 
            {
                ## we split the string and are reading in chunks so get all lines here, join and isolate the response text
                if( $Matches[1] -like '<html>*' )
                {
                    ## it's html and probably not complete so unlikely we can pull anything useful out of it
                    $errorText = 'html response from server'
                }
                else
                {
                    $errorText = $Matches[1]
                }
            }
            $badEvents.Add( ([pscustomobject]@{ 'Time' = $event.TimeCreated ; 'Id' = $event.Id ; Error = "Error code $errorCode `"$($errorText.Trim( ' .'))`"" }) )
        }
    }}
}

[array]$userSessions = @( Get-WTSSessionInformation | Sort-Object -Property LogonTime )

[string]$appVolumesLogFile = "${env:ProgramFiles(x86)}\CloudVolumes\Agent\Logs\svservice.log"
$AppVolumesLogonEvents = $null
$parsedLogFile = $false

$information.Add( ([pscustomobject]@{ 'Item' = "Logged On or Disconnected Sessions" ; 'Description' = $userSessions.Count } ) )

# Get disk mounts - may not be any if no users
if( ( [array]$AVmounts = @( Get-Partition | Where-Object { $_.AccessPaths.Where( { $_ -match $mountPoint } , 1 ) } ) ) -and $AVmounts.Count -gt 0 )
{
    $information.Add( ([pscustomobject]@{ 'Item' = "Mounted $productName" ; 'Description' = $AVmounts.Count } ) )
}
## labels are CVApps or CVWritables
elseif( ( [array]$AVVolumes = @( Get-CimInstance -ClassName win32_volume -Filter "Label like 'CV%' and BootVolume = 'false' and systemvolume = 'false'" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DeviceID )) -and $AVVolumes.Count -gt 0 )
{
    $information.Add( ([pscustomobject]@{ 'Item' = "Mounted $productName" ; 'Description' = $AVVolumes.Count } ) )
}
elseif( $userSessions.Count )
{
   $warnings.Add( "No disk partititons mounted on $mountPoint - may be correct depending on assignments" )
}

if( $userSessions -and $userSessions.Count )
{
    [int]$userSessionCounter = 0
    ForEach( $userSession in $userSessions )
    {
        $userSessionCounter++
        ## WTS API logon time is not accurate so we need to find the LSASS one
        $lsaSession = $lsaSessions.Where( { $_.SessionId -eq $userSession.SessionId -and $_.Domain -eq $userSession.DomainName -and $_.Username -eq $userSession.UserName } , 1 )
        [datetime]$logonTime = $(if( $lsaSession ) { $lsaSession.LoginTime } else { $userSession.LogonTime })

        if( $logonTime -ge $startDate -and $svserviceEventLogEntries -and ( $logonRecord = $svserviceEventLogEntries.Where( { $_.Id -eq '210' -and $_.TimeCreated -ge $logonTime -and $_.Properties[1].Value -match "Session ID:\s*$($userSession.SessionId)$" } , 1 ) ) )
        {
            ## if we haven't parsed the service log file yet we'll do it now so that we only get events after the logon of the first session chronologically
            ## need to parse the log file for App Volumes v2 to get GUIDS so we can find disk mount times
            ## we also need the log file for 4.x to see if it has completely failed because of SQL issues which don't go into the event log
            if( ! $parsedLogFile -and (Test-Path -Path $appVolumesLogFile -ErrorAction SilentlyContinue) )
            {
                $timeZone = Get-TimeZone
                
                $AppVolumesLogonEvents = @( Get-Content -Path $appVolumesLogFile | . { Process {
                    ## [2020-01-10 13:43:23.383 UTC] [svservice:P7776:T1108] Service path: C:\Program Files (x86)\CloudVolumes\Agent\svservice.exe
                    ## Split into 3 - the first two [ ] delimited sections and then the rest
                    if( $_ -match '^\[([^\]]+)\] \[([^\]]+)\] (.+)$' -and ( $time = [datetime]::ParseExact( $Matches[1] , 'yyyy-MM-dd HH:mm:ss.fff UTC' , $null ) ) `
                        -and ( $adjustedForTimeZone = [System.TimeZoneInfo]::ConvertTimeFromUtc( $time, $timeZone ) ) `
                            -and $adjustedForTimeZone -ge $logonTime )
                    {
                        [pscustomobject]@{
                            'Time' = $adjustedForTimeZone
                            'ProcessInfo' = $Matches[2]
                            'Message' = $Matches[3]
                        }
                    }
                    
                }})
                $parsedLogFile = $true
            }

            if( $appVolumesVersion -match '^2\.' )
            {
                ## now try and marry up GUIDs between log file and event log so we can map mount points to apps
                ##[array]$AppVolGUIDMappings = $( @( 
                Foreach ($App in $AppList) {
                    $result = $null
                    Write-Verbose "AppGUID   : $($App.AppGUID)"
                    $AppVolGUIDEvents = $AppVolumesLogonEvents.Where( { $_.Message -like "*$($App.AppGUID)*" } )
                    Write-Debug "AppVolGUIDEvents : $($AppVolGuidEvents | Out-String)"
                    $AppVolumesLogonEvents.Where( { $_.Message -like "*$($App.AppGUID)*" } ) | . { Process {
                        $GUIDEvent = $_.Message
                        if( ! $result -and $GUIDEvent -match '(\\Device\\\w+)' ) {
                            [string]$appDevice = $Matches[1]
                            Write-Verbose "AppDevice : $appDevice"
                            $AppVolumesLogonEvents.Where( {$_.Message -like "*$appDevice*"} ) | . { Process {
                            ## Do we assume VMWare will always keep the messages to a specific format? Or filter out the AppGUID
                            ## and assume what's left is the device GUID?  I'm leaning towards the latter... Let me know if this fails future Trentent
                                $AppVolDeviceEvent = $_.Message
                                if (($AppVolDeviceEvent -like "*$appDevice*") -and ($AppVolDeviceEvent -notlike "*$($App.AppGUID)*")) {
                                    Write-Debug "AppVolDeviceEvent: $AppVolDeviceEvent"
                                    if( $appDevice -and $App.AppGUID -and ( $GUIDs = [regex]::Match( $AppVolDeviceEvent , "({[0-9A-Fa-f]{8}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{12}})" ) ) -and $GUIDS.Success )
                                    {
                                        $DiskGUID =  $(if( ! ( $GUIDS -is [array] ) -or $GUIDS.count -eq 1 ) {
                                            $GUIDS.Value
                                        } elseif ($GUIDS.count -eq 2) {
                                            $GUIDS.Value.Where( { $_ -ne $App.AppGUID } )
                                        } else {
                                            Write-Debug "AppVolumes: ERROR -- Unable to determine DiskGUD from `"$AppVolDeviceEvent`""
                                        })
                                        ## only create object if all fields are present
                                        if( $DiskGUID -and ($appName = $AppList.Where( { $_.AppGUID -eq $App.AppGUID } , 1 ) | Select-Object -ExpandProperty AppName ) )
                                        {
                                            Write-Verbose "$appName DiskGUID : $DiskGUID"
                                            ## add to existing app object if not already present
                                            if( $app.PSObject.Properties[ 'DiskGUID' ] )
                                            {
                                                if( $app.DiskGUID -ne $DiskGUID )
                                                {
                                                    $warnings.Add( "Different disk GUIDs $($app.DiskGUID) and $DiskGUID found for app $appName" )
                                                }
                                            }
                                            else
                                            {
                                                Add-Member -InputObject $app -MemberType NoteProperty -Name DiskGUID -Value $DiskGUID
                                            }
                                            if( $app.PSObject.Properties[ 'AppName' ] )
                                            {
                                                if( $app.AppName -ne $AppName )
                                                {
                                                    $warnings.Add( "Different app names $($app.AppName) and $appName for same app via GUID $($app.AppGUID)" )
                                                }
                                            }
                                            else
                                            {
                                                Add-Member -InputObject $app -MemberType NoteProperty -Name AppName -Value $AppName
                                            }
                                            if( $app.PSObject.Properties[ 'AppDevice' ] )
                                            {
                                                if( $app.AppDevice -ne $AppDevice )
                                                {
                                                    $warnings.Add( "Different app devices $($app.Device) and $AppDevice for app $appName" )
                                                }
                                            }
                                            else
                                            {
                                                Add-Member -InputObject $app -MemberType NoteProperty -Name AppDevice -Value $AppDevice
                                            }
                                        }
                                    }
                                }
                            }}
                        }
                    } }
                }
            }

            ## There will be a pair of 226 (Detected new volume, processing...) and 227 (New volume finished processing) events for each disk mounted for the user so time each
            [array]$thisUsersApps = @( $AppList.Where( { $_.username -eq $userSession.UserName -and $_.domainname -eq $userSession.DomainName -and $_.Time -ge $logonTime } ) )
            Write-Verbose -Message "Looking for $($thisUsersApps.Count) apps for user $($userSession.username)"
            ForEach( $userApp in $thisUsersApps )
            {
                [string]$message = "disk mount for $(if( $userApp.MountType -eq 'MOUNTED-WRITE' )
                    {
                        "writable volume"
                    }
                    else
                    {
                        "app $($userapp.AppName)"
                    })"
                $message += " from $($userapp.AppPath) for user $($userSession.UserName) in session $($userSession.SessionId), logged on at $(Get-Date -Date $logonTime -Format G)"

                if( ! ( $startMount = $svserviceEventLogEntries.Where( { $_.id -eq '226' -and $_.TimeCreated -ge $logonRecord.TimeCreated -and ( $_.Properties[1].value -imatch "\bGUID: $($userapp.DiskGUID)" -or $_.Properties[1].value -imatch "VolumeId: $($userapp.AppGUID)" ) } , 1 ) ))
                {
                    $message = "Failed to find start event for " + $message
                    $warnings.Add( $message )
                }
                elseif( ! ( $endMount = $svserviceEventLogEntries.Where( { $_.id -eq '227' -and $_.TimeCreated -ge $logonRecord.TimeCreated -and ( $_.Properties[1].value -imatch "\bGUID: $($userapp.DiskGUID)" -or $_.Properties[1].value -imatch "VolumeId: $($userapp.AppGUID)" ) } , 1 ) ))
                {
                    $message = "Failed to find end event for " + $message
                    $warnings.Add( $message )
                }
                else
                {
                    ## first character needs to be made upper case
                    $information.Add( ([pscustomobject]@{ 'Item' = ( '{0}{1}' -f $message.Substring(0,1).ToUpper() , $message.Substring( 1 ) ) ; 'Description' = "$([math]::Round( ($endMount.TimeCreated - $startMount.TimeCreated).TotalSeconds , 2 )) seconds" } ) )

                    ## check we can find the mounted partition
                    if( ! ( [string]$GUID = $(if( $endMount.properties[1].value -match 'GUID:\s*({.*})' ) { $Matches[1 ] }) ) `
                        -and ! ( [string]$GUID = $(if( $startMount.properties[1].value -match 'GUID:\s*({.*})' ) { $Matches[1 ] }) ) )
                    {
                        $warnings.Add( "Unable to find GUID for partition access path for $($userApp.AppName) for $($userSession.UserName)" )
                    }
                    elseif( ! ( $thisMount = $AVmounts.Where( { $_.AccessPaths.Where( { $_ -match $GUID } , 1 ) } ) ) `
                        -and ! $AVVolumes.Where( { $_ -match $GUID } , 1 ) )
                    {
                        $warnings.Add( "Unable to find a mounted volume for GUID $GUID for user $($userSession.UserName) $($userApp.AppName)" )
                    }
                }
            }

            ## Check for catastrophic failures via log file
            if( $AppVolumesLogonEvents -and $AppVolumesLogonEvents.Count )
            {
                ## get logon time of next user session to ensure we only look before that
                $nextlogontime = $null
                if( $userSessionCounter -lt $userSessions.Count )
                {
                    $nextlsaSession = $lsaSessions.Where( { $_.SessionId -eq $userSessions[$userSessionCounter].SessionId -and $_.Domain -eq $userSessions[$userSessionCounter].DomainName -and $_.Username -eq $userSessions[$userSessionCounter].UserName } , 1 )
                    $nextlogonTime = $(if( $nextlsaSession ) { $nextlsaSession.LoginTime } else { $userSessions[$userSessionCounter].LogonTime })
                }
                if( ( $theEvent = $AppVolumesLogonEvents.Where( { $_.Time -ge $logonTime -and ( ! $nextlogontime -or $_.Time -lt $nextlogontime ) -and $_.Message -match 'HttpUserLogin: (succeeded|failed) \(user login\)' } , 1 ) ) )
                {
                    if( $Matches[1] -eq 'failed' )
                    {
                        $badEvents.Add( ([pscustomobject]@{ 'Time' = $theEvent.Time ; 'Id' = $null ; Error = "Found failed logon to $productName server in log file, probably for user $($userSession.UserName) $($userApp.AppName)" }) )
                    }
                }
                else
                {
                    $warnings.Add( "Unable to find successful logon to $productName server for user $($userSession.UserName) $($userApp.AppName)" )
                }
            }
        }
        else
        {
            $warnings.Add( "Unable to find $productName logon event for $($userSession.username) in session $($userSession.SessionId)" )
        }
    }
}

$information | Format-Table -AutoSize

if( $warnings -and $warnings.Count )
{
    Write-Output -InputObject "$($warnings.Count) warnings:`n"
    $warnings | Format-Table -AutoSize
}
else
{
    Write-Output -InputObject "No warnings to report"
}

if( $badEvents -and $badEvents.Count )
{
    [string]$message = "Found $($badEvents.Count) events indicating $productName issues"
    if( $oldestEvent )
    {
        $message += ", oldest event in event log is from $(Get-Date -Date ($oldestEvent.TimeCreated) -Format G)"
    }

    Write-Output -InputObject "`n$message"

    $badEvents | Sort-Object -Property Time | Format-Table -AutoSize
}

