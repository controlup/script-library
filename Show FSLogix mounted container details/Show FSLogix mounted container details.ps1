#requires -version 3

<#
.SYNOPSIS
Show FSLogix currently mounted volume details & cross reference to FSLogix session information in the registry

.DESCRIPTION
Gets Windows disks, volumes and partitions information and correlates with HKEY_LOCAL_MACHINE\SOFTWARE\FSLogix\Profiles\Sessions to show disk sizes, capacities and free space
Cross references to file share to check space and vhd size

.PARAMETER label
Only include partitions whose label matches this regular expression. They are typically labelled "Profile-%username%"

.PARAMETER searchWindowMinutes
How many minutes after logon to search the event logs for FSlogix events for a specific session

.NOTES
    Based on https://github.com/guyrleech/Microsoft/blob/master/Show%20FSlogix%20volumes.ps1
    
    Modification History:

    2022/08/18   GRL   Initial public release
    2022/08/26   GRL   Only add share info if available, added summary before details
#>

[CmdletBinding()]

Param
(
    [string]$label ,
    [decimal]$searchWindowMinutes = 10
    ##[switch]$noUsedSpace 
)

## TODO disconnects
## TODO AVD
## TODO .metadata (separate script)

$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { 'Continue' } else { 'SilentlyContinue' })
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { 'Continue' } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'ErrorAction' ] ) { $ErrorActionPreference } else { 'Stop' })

## for Office containers only - https://docs.microsoft.com/en-us/fslogix/office-container-configuration-reference
[hashtable]$vhdAccessModes = @{
    0 = 'Direct Access'
    1 = 'Difference disk stored on network'
    2 = 'Difference disk stored on local machine'
    3 = 'Unique disk per session'
}

## https://docs.microsoft.com/en-us/fslogix/fslogix-error-codes-reference
[hashtable]$fslogixReasonCodes = @{
    4	= 'The FSLogix system will not handle profiles for special users'
    2	= 'The user is a member of the FSLogix Exclude group, and should therefore not receive a FSLogix Profile'
    3	= 'A local profile for the user already exists'
    1	= 'The user is not a member of the FSLogix Include group, and should therefore not receive a FSLogix Profile'
    0	= 'The FSLogix Profile has been attached and is working'
}

[hashtable]$fslogixErrorCodes = @{
    0	= 'The system is working as expected. Check Reason to see the state of the Profile'
    1	= 'The system is in an error state'
    2	= 'The DLL that provides the Virtual Disk API ("virtdisk.dll") cannot be found'
    3	= 'Unable to get the user SID from the user token'
    5	= 'A security API failed'
    6	= 'There was an error determining the path to the VHD/X file'
    7	= 'There was an error creating a directory'
    8	= 'There was an error impersonating the user'
    9	= 'There was an error creating the VHD/X file'
    10	= 'There was an error closing a handle'
    11	= 'There was an error opening the VHD/X file'
    12	= 'There was an error attaching the VHD/X'
    13	= 'There was an error getting the physical path of the virtual disk'
    14	= 'There was an error opening the device'
    15	= 'There was an error initializing the disk'
    16	= 'There was an error retrieving the volume GUID'
    17	= 'There was an error formatting the volume'
    18	= 'Unable to determine the user''s profile directory'
    19	= 'There was an error creating a junction in the file system'
    20	= 'There was an error importing registry data'
    21	= 'There was an error checking group membership for the user'
    22	= 'There was an error trying to determine the profile type'
    23	= 'There was an error processing the redirections.xml file'
    100	= 'The VHD/X is attached and ready. The system is waiting for the Windows Profile Service to begin creation of the user''s profile'
    200	= 'The FSLogix Profile system is currently working on setting up the profile'
    300	= 'The FSLogix Profile was already attached for the user logging on. This only happens on a machine that has been configured to allow multiple, concurrent logons for the same user'
}

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

if( $null -eq ($fslogixInstalls = Get-ItemProperty -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.PSObject.Properties[ 'displayname' ] -and $_.DisplayName -match 'fslogix' -and $_.Publisher -match 'fslogix' } ))
{
    Write-Warning -Message "FSlogix does not appear to be installed"
}

if( $null -eq ($fslogixServices = @( Get-Service -DisplayName 'FSlogix*' -ErrorAction SilentlyContinue ) ) )
{
    Write-Warning -Message "No FSlogix services found"
}
else
{
    ForEach( $service in $fslogixServices )
    {
        if( $service.Status -ine 'running' )
        {
            Write-Warning -Message "`"$($service.displayname)`" service is not running, it is $($service.status)"
        }
    }
}

if( $null -eq ( $fslogixDrivers = Get-CimInstance -ClassName win32_systemdriver -Filter "Caption like 'FSlogix%'" -ErrorAction SilentlyContinue ) )
{
    Write-Warning -Message 'No FSlogix device drivers found'
}
else
{
    ForEach( $driver in $fslogixDrivers )
    {
        if( $driver.State -ine 'running' )
        {
            Write-Warning -Message "`"$($driver.displayname)`" driver is not running, it is $($driver.State)"
        }
    }
}

# TODO check fslogix enabled

Function Get-FolderSize( [string]$folderName )
{
    $items = @( $folderName )
    [array]$files = While( $items )
    {
        $newitems = $items | Get-ChildItem -Force -ErrorAction SilentlyContinue | Where-Object { ! ( $_.Attributes -band [System.IO.FileAttributes]::ReparsePoint ) }
        $newitems
        $items = $newitems | Where-Object { $_.Attributes -band [System.IO.FileAttributes]::Directory }
    }
    if( $files -and $files.Count )
    {
        [long]($files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue | Select -ExpandProperty Sum)
    }
    else
    {
        [long]0
    }
}

[array]$partitions = @( Get-Partition | Where-Object { $_.DiskId -match '&ven_msft&prod_virtual_disk' -and ! $_.DriveLetter -and $_.Type -eq 'Basic' } )

if( ! $partitions -or ! $partitions.Count )
{
    Write-Warning "No partitions found mounted off virtual disks"
}

Write-Verbose "Found $($partitions.Count) virtual disk partitions"

[array]$fixedVolumes = @( Get-Volume | Where-Object { $_.DriveType -eq 'Fixed' } )

if( ! $fixedVolumes -or ! $fixedVolumes.Count )
{
    Write-Warning "Unable to find any fixed volumes"
}
else
{
    Write-Verbose "Found $($fixedVolumes.Count) fixed volumes"
}

[array]$virtualDisks = @( Get-Disk | Where-Object { $_.BusType -eq 'File Backed Virtual' } )

if( ! $virtualDisks -or ! $virtualDisks.Count )
{
    Write-Warning "Unable to find any file backed virtual disks"
}
else
{
    Write-Verbose "Found $($virtualDisks.Count) file backed virtual disks"
}

[int]$counter = 0

## from ALD - so we know where to search in the event log of volume attach start and complete events to time the mounting

#region LSASS

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


## Can't use WMI/CIM since servers could be non-Windows
Add-Type @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PInvoke.Win32
{
    public static class Disk
    {
        // Thanks to https://www.pinvoke.net/default.aspx/kernel32.getdiskfreespaceex
        [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool GetDiskFreeSpaceEx(
                string lpDirectoryName, 
                out ulong lpFreeBytesAvailable, 
                out ulong lpTotalNumberOfBytes, 
                out ulong lpTotalNumberOfFreeBytes);
    }
}
'@

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

                if( $secType -ieq 'RemoteInteractive' )
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
            [void][Win32.Secure32]::LsaFreeReturnBuffer( $sessionData )
            $sessionData = [IntPtr]::Zero
        }
        $iter = $iter.ToInt64() + [System.Runtime.InteropServices.Marshal]::SizeOf([type][Win32.Secure32+LUID])  # move to next pointer
    }) | Sort-Object -Descending -Property 'LoginTime'

    [void]([Win32.Secure32]::LsaFreeReturnBuffer( $luidPtr ))
    $luidPtr = [IntPtr]::Zero

    Write-Verbose "Found $(if( $lsaSessions ) { $lsaSessions.Count } else { 0 }) LSA sessions, earliest session $(if( $earliestSession ) { Get-Date $earliestSession -Format G } else { 'never' })"
}
#endregion LSASS

#region WTSAPI

# from https://github.com/guyrleech/Microsoft/blob/master/WTSApi.ps1

[string]$WTSApi = @'
    using System;
    using System.Text;
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
    
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct WTS_PROCESS_INFO_W {
        public uint SessionId;
        public uint ProcessId;
        [MarshalAs(UnmanagedType.LPTStr)]
        public String pProcessName;
        public IntPtr pUserSid;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
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
    
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct WTSCONFIGINFOW {
        public UInt32 version;
        public UInt32 fConnectClientDrivesAtLogon;
        public UInt32 fConnectPrinterAtLogon;
        public UInt32 fDisablePrinterRedirection;
        public UInt32 fDisableDefaultMainClientPrinter;
        public UInt32 ShadowSettings;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 21)]   
        public string  LogonUserName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 18)]   
        public string  LogonDomain;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 261)]   
        public string  WorkDirectory;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 261)]   
        public string  InitialProgram;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 261)]   
        public string  ApplicationName;  
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct WTSCLIENTW {
      [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 21)]    
      public string   ClientName;
      [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 18)]
      public string   Domain;
      [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 21)]
      public string   UserName;
      [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 261)]
      public string   WorkDirectory;
      [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 261)]
      public string   InitialProgram;
      public byte   EncryptionLevel;
      public UInt32  ClientAddressFamily;
      [MarshalAs(UnmanagedType.ByValArray, SizeConst = 31)]
      public UInt16[] ClientAddress;
      public UInt16 HRes;
      public UInt16 VRes;
      public UInt16 ColorDepth;
      [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 261)]
      public string   ClientDirectory;
      public UInt32  ClientBuildNumber;
      public UInt32  ClientHardwareId;
      public UInt16 ClientProductId;
      public UInt16 OutBufCountHost;
      public UInt16 OutBufCountClient;
      public UInt16 OutBufLength;
      [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 261)]
      public string   DeviceId;
    }
        
    [StructLayout(LayoutKind.Sequential)]
    public struct WTS_CLIENT_DISPLAY
    {
        public uint HorizontalResolution;
        public uint VerticalResolution;
        public uint ColorDepth;
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
        public static extern int WTSEnumerateProcessesW(
                 System.IntPtr hServer,
                 uint Reserved,
                 uint Version,
                 ref System.IntPtr ppProcessInfo,
                 ref int pCount);
                 
        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern int WTSWaitSystemEvent(
                 System.IntPtr hServer,
                 int EventMask,
                 ref System.IntPtr pEventFlags );

        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern IntPtr WTSOpenServer(string pServerName);
        
        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern void WTSCloseServer(IntPtr hServer);
        
        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern void WTSFreeMemory(IntPtr pMemory);

        [DllImport("advapi32.dll", SetLastError=true)]
        public static extern bool ConvertSidToStringSidA(IntPtr pSid , ref StringBuilder stringSid );

        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern IntPtr LocalFree( IntPtr hMem );
    }
'@

Function Get-WTSSessionInformation
{
    [cmdletbinding()]

    Param
    (
        [string[]]$computers = @( $null ) ,
        [int]$waitForLogonTimeInMilliseconds
    )

    [long]$count = 0
    [IntPtr]$ppSessionInfo = 0
    [IntPtr]$ppQueryInfo = 0
    [long]$ppBytesReturned = 0
    $wtsSessionInfo = New-Object -TypeName 'WTS_SESSION_INFO'
    $wtsInfoEx = New-Object -TypeName 'WTSINFOEX'
    $wtsClientInfo = New-Object -TypeName 'WTSCLIENTW'
    $wtsConfigInfo = New-Object -TypeName 'WTSCONFIGINFOW'
    [int]$datasize = [system.runtime.interopservices.marshal]::SizeOf( [Type]$wtsSessionInfo.GetType() )

    ForEach( $computer in $computers )
    {
        $wtsinfo = $null
        [string]$machineName = $(if( $computer ) { $computer } else { $env:COMPUTERNAME })
        [IntPtr]$serverHandle = [wtsapi]::WTSOpenServer( $computer )

        ## If the function fails, it returns a handle that is not valid. You can test the validity of the handle by using it in another function call.

        [long]$retval = [wtsapi]::WTSEnumerateSessions( $serverHandle , 0 , 1 , [ref]$ppSessionInfo , [ref]$count );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

        if ($retval -ne 0)
        {
            Write-Verbose -Message "Got $count sessions for $machineName"
             for ([int]$index = 0; $index -lt $count; $index++)
             {
                 ## session 0 is non-interactive (session zero isolation)
                 if( ( $element = [system.runtime.interopservices.marshal]::PtrToStructure( [long]$ppSessionInfo + ($datasize * $index), [type]$wtsSessionInfo.GetType()) ) -and $element.SessionID -ne 0 )
                 {
                    #$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                    [bool]$continueChecking = $true
                    do
                    {
                         if( ( $retval = [wtsapi]::WTSQuerySessionInformationW( $serverHandle , $element.SessionID , [WTS_INFO_CLASS]::WTSSessionInfoEx , [ref]$ppQueryInfo , [ref]$ppBytesReturned ) -and $ppQueryInfo ) -and $ppQueryInfo )
                         {
                            if( ( $value = [system.runtime.interopservices.marshal]::PtrToStructure( $ppQueryInfo , [Type]$wtsInfoEx.GetType())) -and $value.Data `
                                -and $value.Data.WTSInfoExLevel1.SessionState -ne [WTS_CONNECTSTATE_CLASS]::WTSListen -and $value.Data.WTSInfoExLevel1.SessionState -ne [WTS_CONNECTSTATE_CLASS]::WTSConnected `
                                    -and $value.Data.WTSInfoExLevel1.SessionState -ne [WTS_CONNECTSTATE_CLASS]::WTSConnectQuery)
                            {
                                if( $wtsinfo = $value.Data.WTSInfoExLevel1 )
                                {
                                    if( $wtsinfo.LogonTime -gt 0 )
                                    {
                                        $idleTime = New-TimeSpan -End ([datetime]::FromFileTimeUtc($wtsinfo.CurrentTime)) -Start ([datetime]::FromFileTimeUtc($wtsinfo.LastInputTime))
                                        Add-Member -InputObject $wtsinfo -Force -NotePropertyMembers @{
                                            'IdleTimeInSeconds' =  [math]::Round( ( $idleTime | Select -ExpandProperty TotalSeconds ) , 1 )
                                            'IdleTimeInMinutes' =  [math]::Round( ( $idleTime | Select -ExpandProperty TotalMinutes ) , 2 )
                                            'Computer' = $machineName
                                            'LogonTime' = [datetime]::FromFileTime( $wtsinfo.LogonTime )
                                            'DisconnectTime' = $( $time = [datetime]::FromFileTime( $wtsinfo.DisconnectTime ) ; if( $time.Year -lt 1900 ) { $null } else { $time })
                                            'LastInputTime' = [datetime]::FromFileTime( $wtsinfo.LastInputTime )
                                            'SessionState' = $wtsinfo.SessionState
                                            'ConnectTime' = [datetime]::FromFileTime( $wtsinfo.ConnectTime )
                                            'CurrentTime' = [datetime]::FromFileTime( $wtsinfo.CurrentTime ) }
                                        $continueChecking = $false
                                    }
                                    elseif( $PSBoundParameters[ 'waitForLogonTimeInMilliseconds' ] )
                                    {
                                        Write-Warning -Message "$(Get-Date -Format G): zero logon time"
                                        Start-Sleep -Milliseconds 200
                                    }
                                    else ## not got logon time but not asked to wait so don't loop
                                    {
                                        $continueChecking = $false
                                    }
                                }
                                else ## no WTSInfoExLevel1 data
                                {
                                    $continueChecking = $false
                                }
                            }
                            else ## no data or not in a state we are interested in
                            {
                                $continueChecking = $false
                            }
                            [wtsapi]::WTSFreeMemory( $ppQueryInfo )
                            $ppQueryInfo = [IntPtr]::Zero
                         }
                         else
                         {
                            Write-Error "$($machineName): $LastError"
                            $continueChecking = $false
                         }
                     } while( $continueChecking )

                     if( $wtsinfo )
                     {
                        ## WTSClientInfo
                        $ppQueryInfo = [IntPtr]::Zero
                        if( ( $retval = [wtsapi]::WTSQuerySessionInformationW( $serverHandle , $element.SessionID , [WTS_INFO_CLASS]::WTSClientInfo , [ref]$ppQueryInfo , [ref]$ppBytesReturned ) ) `
                            -and $ppQueryInfo -and ( $wtsClientInfo = [system.runtime.interopservices.marshal]::PtrToStructure( $ppQueryInfo , [Type]$wtsClientInfo.GetType()) ) )
                        {
                            ForEach( $property in $wtsClientInfo.PSObject.Properties )
                            {
                                Add-Member -InputObject $wtsinfo -MemberType NoteProperty -Name $property.Name -Value $property.Value -Force
                            }
                           [wtsapi]::WTSFreeMemory( $ppQueryInfo )
                           $ppQueryInfo = [IntPtr]::Zero
                        }
                        else
                        {
                            $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                            Write-Warning -Message "Failed to get WTSClientInfo for session id $($element.SessionID)"
                        }
                        
                        ## WTSConfigInfo
                        $ppQueryInfo = [IntPtr]::Zero
                        if( ( $retval = [wtsapi]::WTSQuerySessionInformationW( $serverHandle , $element.SessionID , [WTS_INFO_CLASS]::WTSConfigInfo , [ref]$ppQueryInfo , [ref]$ppBytesReturned ) ) `
                            -and $ppQueryInfo -and ( $wtsConfigInfo = [system.runtime.interopservices.marshal]::PtrToStructure( $ppQueryInfo , [Type]$wtsConfigInfo.GetType()) ) )
                        {
                            ForEach( $property in $wtsConfigInfo.PSObject.Properties )
                            {
                                ## WorkDirectory and InitialProgram don't seem to work and we have no new strings here so don't add string type properties
                                if( $property.TypeNameOfValue -ne 'System.String' )
                                {
                                    Add-Member -InputObject $wtsinfo -MemberType NoteProperty -Name $property.Name -Value $property.Value -Force
                                }
                            }
                           [wtsapi]::WTSFreeMemory( $ppQueryInfo )
                           $ppQueryInfo = [IntPtr]::Zero
                        }
                        else
                        {
                            $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                            Write-Warning -Message "Failed to get WTSConfigInfo for session id $($element.SessionID)"
                        }

                        [UInt16]$clientProtocolType = ([UInt16]::MaxValue)
                        
                        if( ( $retval = [wtsapi]::WTSQuerySessionInformationW( $serverHandle , $element.SessionID , [WTS_INFO_CLASS]::WTSClientProtocolType , [ref]$ppQueryInfo , [ref]$ppBytesReturned ) ) -and $ppQueryInfo )
                        {
                            $clientProtocolType = [system.runtime.interopservices.marshal]::PtrToStructure( $ppQueryInfo , [Type]$clientProtocolType.GetType())
                            Add-Member -InputObject $wtsinfo -MemberType NoteProperty -Name ClientProtocolType -Value $clientProtocolType
                            [wtsapi]::WTSFreeMemory( $ppQueryInfo )
                            $ppQueryInfo = [IntPtr]::Zero
                        }
                        $wtsinfo
                        $wtsinfo = $null
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

try
{
    Add-Type -TypeDefinition $WTSApi
}
catch
{
    ## hopefully because already loaded otherwise we are doomed
}

[array]$WTSsessions = @( Get-WTSSessionInformation )

Write-Verbose -Message "Got $($WTSsessions.Count) WTS sessions"

if( $WTSsessions -and $WTSsessions.Count -eq 0 )
{
    Write-Warning -Message "Found no logged on sessions so should not be any mounted FSlogix volumes"
}

#endregion WTSAPI

## "feature" in PS ISE means we have to translate paths
## see https://twitter.com/matbg/status/155777500454004736
[string]$pathFixup = '^$' ## won't match any non-empty string
if( $host -and $host.Name -match '\bISE\b' )
{
    $pathFixup = '\\\\\?\\'
}

## cache shares so we ony interrogate once for capacity & free space
[hashtable]$shares = @{}
[hashtable]$uniqueNetworkVHDs = @{}
[hashtable]$uniqueShares = @{}
[hashtable]$uniqueShareHosts = @{}
[long]$VHDsizeTotalMB = 0
[int]$vhdxMeasured = 0

[array]$results = @( ForEach( $partition in $partitions )
{
    $counter++
    Write-Verbose "$counter / $($partitions.Count) : Partition GUID $($partition.Guid)"

    $volume = $fixedVolumes | Where-Object { $_.UniqueId -match $partition.Guid }
    if( -Not $volume )
    {
        Write-Warning "Unable to find fixed volume with GUID $($partition.Guid)"
    }
    if( -Not $PSBoundParameters[ 'label' ] -or ($volume -and $volume.FileSystemLabel -match $label ))
    {
        [string]$uniqueId = ($partition.UniqueId -split '[{}]')[-1]
        $disk = $virtualDisks | Where-Object { $_.UniqueId -eq $uniqueId }
        if( -Not $disk )
        {
            Write-Warning "Unable to find disk with unique id $uniqueId"
        }
        $result = [pscustomobject][ordered]@{
            'Label' = $volume | Select-Object -ExpandProperty FileSystemLabel
            'Operational Status' = $volume | Select-Object -ExpandProperty OperationalStatus
            'Health Status' = $volume | Select-Object -ExpandProperty HealthStatus
            'Provisioning Type' = $disk | Select-Object -ExpandProperty ProvisioningType
            ##'Disk Size (GB)' = [math]::Round( ( $disk | Select-Object -ExpandProperty Size ) / 1GB , 2 )
            ##'Volume Size (GB)' = [math]::Round( ( $volume | Select-Object -ExpandProperty Size ) / 1GB , 2 )
            'Volume Capacity (GB)' = [math]::Round(  ( $volume | Select-Object -ExpandProperty SizeRemaining ) / 1GB , 2 )
            ## avoid divide by zero
            'Volume Free Capacity %' = $(if( $volume -and $volume.PSObject.Properties[ 'size' ] -and $volume.size -gt 0 ) { [math]::Round( ( $volume | Select-Object -ExpandProperty SizeRemaining ) / $volume.Size * 100 , 2 ) })
        }
        <#  ## Commennted out because get too many folders when ODFC so need a better way of showing data usage
        Write-Verbose -Message "Partition is `"$($partition.AccessPaths)`""
        ## \\?\Volume{451ca07e-00c3-40bb-a5f9-75a559033bb8}\ 
        [array]$paths = Get-ChildItem -LiteralPath (($partition | Select-Object -ExpandProperty AccessPaths) -replace $pathFixup , '\\.\') | . { Process `
        {
            [string]$folder = $_.FullName
            [string]$childFolder = $_.Name
        
            if( -Not $noUsedSpace )
            {
                Add-Member -InputObject $result -MemberType NoteProperty -Name "`"$childFolder`" Folder Size (GB)" -Value ([math]::Round( (Get-FolderSize -folderName $folder) / 1GB , 2 ))
            }

            ##Add-Member -InputObject $result -MemberType NoteProperty -Name "`"$childFolder`" Folder Permissions" -Value ((Get-Acl -LiteralPath $folder | Select -ExpandProperty AccessToString) -replace "[`n`r]" , ' , ')
        }}
        #>
        [bool]$profileDisk = $false
        [bool]$officeDisk = $false
        [bool]$gotShareInfo = $false


        $fslogixRegValue = Get-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles\Sessions\*" -ErrorAction SilentlyContinue | Where-Object { $_.Volume -eq $volume.Path }
        [string]$userSID = $null
        if( -Not ($profileDisk = ($null -ne $fslogixRegValue )))
        {
            if( $fslogixRegValue = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\FSLogix\ODFC\Sessions\*" -ErrorAction SilentlyContinue | Where-Object { $_.Volume -eq $volume.Path } )
            {
                $officeDisk = $true
            }
        }
        if( $fslogixRegValue )
        {
            Add-Member -Force -InputObject $result -NotePropertyMembers @{
                'Username' = ([System.Security.Principal.SecurityIdentifier]( $userSID = $fslogixRegValue.PSChildName )).Translate([System.Security.Principal.NTAccount]).Value
                'Profile Path' = $(if( $fslogixRegValue.PSObject.Properties[ 'UserProfilePath' ] ) { $fslogixRegValue.UserProfilePath } else {  $fslogixRegValue | Select-Object -ErrorAction SilentlyContinue -ExpandProperty ProfilePath } )
                'Local Profile Path' = $fslogixRegValue | Select-Object -ErrorAction SilentlyContinue -ExpandProperty LocalProfilePath
                'Session Id' = $fslogixRegValue | Select-Object -ErrorAction SilentlyContinue -ExpandProperty WindowsSessionID
                'Last Profile Load Time (s)' = ( $fslogixRegValue | Select-Object -ErrorAction SilentlyContinue -ExpandProperty LastProfileLoadTimeMS ) / 1000
            }
        }
        else
        {
            Write-Warning "Couldn't find FSlogix registry key for volume $($volume.Path)"
        }

        $logontime = $null
        $mountStartTime = $null
        $mountEndTime = $null
        $profileLoadEnd = $null
        $vhdxSize = $null
        $cachedShare = $null
        $shareCapacityGB = $null
        $shareFreeSpaceGB = $null
        [string]$shareName = $null
        [string]$sourceFolder = $null

        ## Get Start of mount for this user and disk
        if( $logon = $lsaSessions | Where-Object { $_.username -eq $result.Username.Split( '\' )[-1] -and $_.domain -eq $result.Username.Split( '\' )[0 ] } | Sort-Object -Descending -Property LoginTime | Select-Object -First 1 )
        {
            [string]$volumeGUID = $null
            ## \\?\Volume{451ca07e-00c3-40bb-a5f9-75a559033bb8}\
            if(  ($partition | Select-Object -ExpandProperty AccessPaths) | Where-Object { $_ -match '({[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}})' }| Select-Object -First 1)
            {
                $volumeGUID = $Matches[ 1 ]
            }
            ## \\?\UNC\grl-nas02\Software\FSLogix\S-1-5-21-1721611859-3364803896-2099701507-1109_billybob\Profile_billybob.VHDX
            $mountStartTime = Get-WinEvent -Oldest -ErrorAction SilentlyContinue -FilterHashtable @{ LogName = 'Microsoft-Windows-VHDMP-Operational' ; Id = 22 ; Starttime = $logon.LoginTime ; EndTime = $logon.LoginTime.AddMinutes( $searchWindowMinutes )}    | Where-Object { $_.Properties[0].value -match $userSID }   | Select-Object -first 1 -ExpandProperty TimeCreated
            $mountEndTime   = Get-WinEvent -Oldest -ErrorAction SilentlyContinue -FilterHashtable @{ LogName = 'Microsoft-Windows-Kernel-IO/Operational' ; Id = 2 ; Starttime = $logon.LoginTime ; EndTime = $logon.LoginTime.AddMinutes( $searchWindowMinutes )} | Where-Object { $_.Properties[0].value -ieq $volumeGUID  } | Select-Object -first 1 -ExpandProperty TimeCreated
            $profileLoadEnd = Get-WinEvent -Oldest -ErrorAction SilentlyContinue -FilterHashtable @{ LogName = 'Microsoft-FSLogix-Apps/Operational' ; Id = 25 ; Starttime = $logon.LoginTime ; EndTime = $logon.LoginTime.AddMinutes( $searchWindowMinutes )} | Where-Object { $_.Properties[4].value -ieq $userSID }
        }

        [string]$location = $disk | Select-Object -ExpandProperty Location
        if( -Not [string]::IsNullOrEmpty( $location ) )
        {
            ## account running the script may not have permissions for share/file
            if( $vhdxProperties = Get-ItemProperty -Path $location -ErrorAction SilentlyContinue )
            {
                $vhdxSize = [math]::Round( ($vhdxProperties.Length ) / 1MB , 1 )
                $vhdxMeasured++
            }

            ## \\grl-nas02\Software\FSLogix\S-1-5-21-1721611859-3364803896-2099701507-2441_admingle\Profile_admingle.VHDX
            
            if( $location -match '^\\\\([^\\]+)\\([^\\]+)\\' )
            {
                $shareName = '\\{0}\{1}' -f $Matches[ 1 ] , $Matches[ 2 ]
                $sourceFolder = Split-Path -Path $location -Parent
                
                $VHDsizeTotalMB += $vhdxSize

                try
                {
                    $uniqueNetworkVHDs.Add( $location , $location )
                }
                catch {}  ## already got, doesn't matter only used for counting

                try
                {
                    $uniqueShares.Add( $sharename , $location )
                }
                catch {}  ## already got, doesn't matter only used for counting

                try
                {
                    $uniqueShareHosts.Add( $Matches[ 1 ] , $location )
                }
                catch {}  ## already got, doesn't matter only used for counting
            }
            elseif( $officeDisk -and $fslogixRegValue -and $fslogixRegValue.PSobject.Properties[ 'VHDRODiffDiskFilePath' ] -and $fslogixRegValue.VHDRODiffDiskFilePath -eq $location )
            {
                if( $fslogixRegValue.PSObject.Properties[ 'VHDRootFilePath' ] -and -Not [string]::IsNullOrEmpty( $fslogixRegValue.VHDRootFilePath ) )
                {
                    if( $fslogixRegValue.VHDRootFilePath -match '^(\\\\[^\\]+\\[^\\]+)\\' )
                    {
                        $shareName = $Matches[ 1 ]
                        $sourceFolder = Split-Path -Path $fslogixRegValue.VHDRootFilePath -Parent
                    }
                    else
                    {
                        Write-Warning -Message "VHDRootFilePath `"$($fslogixRegValue.VHDRootFilePath)`" in $(($fslogixRegValue.PSParentPath -split '\\Registry::')[-1]) does not appear to be a share"
                    }
                }
                else
                {
                    Write-Warning -Message "No VHDRootFilePath value in $(($fslogixRegValue.PSParentPath -split '\\Registry::')[-1])"
                }
            }
            elseif( $location -match '^[A-Z]:\\.*?(?<SID>S-1-5-((32-\d*)|(21-\d*-\d*-\d*-\d*))).*?(?<name>[A-Z]+).*?\.vhd' )## local disk so we need to try and find out from where it came
            {
                Write-Warning -Message "Disk is local at $location but cannot find FSlogix registry entry for the session for SID $($matches['SID'])"
            }

            if( $shareName )
            {
                if( -Not ( $cachedShare = $shares[ $sharename ] ))
                {
                    ## get share info if we can
            
                    [uint64]$userFreeSpace = 0
                    [uint64]$totalSize = 0
                    [uint64]$totalFreeSpace = 0

                    $gotShareInfo = [PInvoke.Win32.Disk]::GetDiskFreeSpaceEx( $shareName , [ref]$userFreeSpace , [ref]$totalSize , [ref]$totalFreeSpace ) ; $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                    if( $gotShareInfo )
                    {
                        $cachedShare = [pscustomobject]@{
                            'Size' = [math]::Round( $totalSize / 1GB , 1 )
                            'FreeSpace' = [math]::Round( $totalFreeSpace / 1GB )
                        }
                        $shares.Add( $shareName , $cachedShare )
                    }
                    else
                    {
                        $thisProcess = Get-Process -Id $pid
                        [int]$parentProcessId = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = '$pid'" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ParentProcessId -ErrorAction SilentlyContinue
                        $parentProcess = $(if( $parentProcessId -gt 0 ) { Get-Process -Id $parentProcessId -ErrorAction SilentlyContinue } )
                        if( $thisProcess.Name -ine 'cuagent' -and $parentProcess.Name -ine 'cuagent' )
                        {
                            Write-Warning -Message "Problem querying share $shareName : $LastError"
                        }
                        else
                        {
                            ## keep quiet as a CU limitation as of the time of script release
                            Write-Verbose -Message "No share access from process $pid $(($process.Name)) , parent $parentProcessId ($($parentProcess|Select-Object -ExpandProperty Name))"
                        }
                    }
                }
                if( $cachedShare )
                {
                    $shareCapacityGB = $cachedShare.Size
                    $shareFreeSpaceGB = $cachedShare.FreeSpace
                }
            }
        }

        Add-Member -Force -InputObject $result -NotePropertyMembers @{
            'OfficeDisk' = $officeDisk
            'Source Folder' = $sourceFolder
            'Paths' = ($partition | Select-Object -ExpandProperty AccessPaths) -join ' , '
            'VHD' = $location
            'VHD Actual Size (MB)' = $vhdxSize
            'VHD Access Mode' = $(if( $officeDisk -and $fslogixRegValue -and $fslogixRegValue.PSobject.Properties[ 'VhdAccessMode' ] ) { $vhdAccessModes[ $fslogixRegValue.VhdAccessMode ] } )
            ##'Physical Sector Size' = $disk | Select-Object -ExpandProperty PhysicalSectorSize
            'Logon Time' = $logon | Select-Object -ExpandProperty LoginTime
            'Mount Start Time' =  $mountStartTime
            'Mount End Time' =  $mountEndTime
            'Mount Duration (s)' = $(if( $mountStartTime -and $mountEndTime ) { [math]::Round( ($mountEndTime - $mountStartTime).TotalSeconds , 2 ) } )
            'Profile Load Time (s)' = $(if( -Not $officeDisk -and $mountStartTime -and $profileLoadEnd ) { [math]::Round( ($profileLoadEnd.TimeCreated - $mountStartTime).TotalSeconds , 2 ) } )
            'Profile Status' = $(if( $profileDisk ) { ( $profileLoadEnd | Select-Object -ExpandProperty Message ) -replace '^Profile load:\s*' -replace '\s*Username:\s*\S.*$' })
        }

        if( $gotShareInfo )
        {
            Add-Member -Force -InputObject $result -NotePropertyMembers @{
                'Share Capacity (GB)' = $shareCapacityGB
                'Share Free Space %' = $(if( $shareCapacityGB -gt 0 ) { [int](($shareFreeSpaceGB / $shareCapacityGB ) * 100) } )
            }
        }

        $result
    }
    else
    {
        Write-Verbose "Excluding $($volume.FileSystemLabel)"
    }
})

if( $null -ne $results -and $results.Count -gt 0 )
{
    [datetime]$lastBootTime = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -ExpandProperty LastBootupTime

    $profileLoadTimeStatistics = $results | Where-Object officeDisk  -eq $false | Measure-Object -Property 'Profile Load Time (s)' -Sum -Average -Maximum -Minimum
    $diskMountTimeStatistics = $results | Measure-Object -Property 'Mount Duration (s)' -Sum -Average -Maximum -Minimum

    Write-Output -InputObject "Results for $($WTSsessions.count) user sessions with $($uniqueNetworkVHDs.Count) network mounted disks from $($uniqueShares.Count) shares on $($uniqueShareHosts.Count) hosts in total"
    if( $vhdxMeasured -gt 0 )
    {
        Write-Output -InputObject "Network VHD disks are consuming $([math]::Round( $VHDsizeTotalMB / 1024 , 1 ))GB, average size is $([Math]::Round( $VHDsizeTotalMB / 1024 / $vhdxMeasured , 1 ))GB"
    }
    Write-Output -InputObject "Last boot time was $(Get-Date -Date $lastBootTime -Format G), up $([math]::Round( ([datetime]::Now - $lastbootTime).TotalHours , 1 )) hours"
    Write-Output -InputObject "Slowest profile load time was $($profileLoadTimeStatistics.Maximum)s, average $($profileLoadTimeStatistics.Average)s"
    Write-Output -InputObject "Slowest VHD mount time was $($diskMountTimeStatistics.Maximum)s, average $($diskMountTimeStatistics.Average)s"

    $sortedPropertyNames = $results[0].psobject.Properties | Select-Object -ExpandProperty Name | Sort-Object

    $results | Select-Object -Property $sortedPropertyNames -ExcludeProperty OfficeDisk
}
else
{
    Write-Warning -Message "No FSlogix volumes found"
}

## Check that each current user session has a result and if not go looking for errors
ForEach( $WTSsession in $WTSsessions )
{
    if( -Not $results -or $results.Count -eq 0 -or -Not $results.Where( { $_.'Session Id' -eq $WTSsession.SessionId } ) )
    {
        Write-Verbose -Message "No FSlogix result for session $($WTSsession.SessionId) for $($WTSsession.Username)"
        $logonTime = $WTSsession.LogonTime
        [string]$username = $WTSsession.username -replace '@.*$' ## can be truncated since only 20 characters
        [string]$userSid = (New-Object -TypeName System.Security.Principal.NTAccount( "$($WTSsession.DomainName)\$Username" )).Translate([System.Security.Principal.SecurityIdentifier]).value
        if( $lsaSession = $lsaSessions | Where-Object {  $_.Session -eq $WTSsession.SessionId -and $_.Domain -ieq $WTSsession.DomainName -and $_.Username -ieq $username } )
        {
            if( $lsaSession -is [array] -and $lsaSession.Count -gt 1 )
            {
                $logontime = $lsaSession[0].LoginTime
            }
            else
            {
                $logontime = $lsaSession.LoginTime
            }
        }
        else
        {
            Write-Warning -Message "Unable to find LSA session for user $($WTSsession.DomainName)\$Username in session id $($WTSsession.SessionId)"
        }

        ## look for profile event to see why it didn't load the profile
        if( $profileLoadEnd = Get-WinEvent -Oldest -ErrorAction SilentlyContinue -FilterHashtable @{ LogName = 'Microsoft-FSLogix-Apps/Operational' ; Id = 25 ; Starttime = $logontime ; EndTime = $logonTime.AddMinutes( $searchWindowMinutes )} | Where-Object { $_.Properties[4].value -ieq $userSID } )
        {
            [int32]$status = $profileLoadEnd.Properties[0].value
            [int32]$reason = $profileLoadEnd.Properties[1].value
            [int32]$error  = $profileLoadEnd.Properties[2].value
            Write-Output -InputObject "FSlogix profile for $($profileLoadEnd.Properties[3].value) failed at $(Get-Date -Format G -Date $profileLoadEnd.TimeCreated)"
            Write-Output -InputObject "`tFSlogix error $status ($($fslogixErrorCodes[ $status ])), reason $reason ($($fslogixReasonCodes[ $reason ])), windows error $error"
        }
        else
        {
            Write-Warning -Message "Unable to find FSlogix profile load event in event log for $($WTSsession.DomainName)\$Username, session id $($WTSsession.SessionId), logged in at $(Get-Date -Format G -Date $lsasession.LoginTime)"
        }
    }
}
