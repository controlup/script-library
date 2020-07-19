<#
.SYNOPSIS

Find processes which have all threads suspended, which aren't UWP apps as they can be validly suspended, or have no open handles and optionally kill them

.DETAILS

Does not look as session 0 processes.

.PARAMETER killParameter

If true will kill all zombie processes found otherwise will not

.MODIFICATION_HISTORY

    19/11/18   @guyrleech   Check session 0 processes except system process (pid=4) and "memory compression" process

    05/12/18   @guyrleech   Added logic to find suspended processes (all threads suspended)

    23/01/20   @guyrleech   Added kill option, code to cater for wwahost.exe and backgroundTaskHost.exe and updated to adhere to ControlUp scripting standards

    30/01/20   @guyrleech   Processes expected to be suspended now found via IsFrozen flag, replaced quser.exe parsing with calling WTS APIs via P/Invoke
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$false,HelpMessage='Whether to kill zombies found or not')]
    [string]$killParameter = 'False'
)

$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { 'Continue' } else { 'SilentlyContinue' })
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { 'Continue' } else { 'SilentlyContinue' })
##$ErrorActionPreference = $(if( $PSBoundParameters[ 'ErrorAction' ] ) { $ErrorActionPreference } else { 'Stop' })

[bool]$kill = [System.Convert]::ToBoolean( $killParameter )
[int]$recentlyCreated = 30
[int]$outputWidth = 400
[string]$logname = 'Microsoft-Windows-User Profile Service/Operational'
[string[]]$excludedProcesses = @( 'System' , 'Idle' , 'Memory Compression' , 'Secure System' , 'Registry' )

function Get-AllProcesses( [bool]$withThreads )
{
    [int]$processCount = 0
    [IntPtr]$processHandle = [IntPtr]::Zero
    [int]$nameBufferSize = 1KB
    [IntPtr]$nameBuffer = [Runtime.InteropServices.Marshal]::AllocHGlobal( $nameBufferSize )
    $path = New-Object UNICODE_STRING

    ## Get all device paths so can convert native paths
    [hashtable]$drivePaths = @{}

    ForEach( $logicalDrive in [Environment]::GetLogicalDrives())
    {
        $targetPath = New-Object System.Text.StringBuilder 256
        if([Kernel32]::QueryDosDevice( $logicalDrive.Substring(0, 2) , $targetPath, 256) -gt 0)
        {
            $targetPathString = $targetPath.ToString()
            $drivePaths.Add( $targetPathString , $logicalDrive )
        }
    }

    do
    {
        ## can't be QueryLimitedInformation as that gives access denied when enumerating threads
        [NT_STATUS]$result = [NtDll]::NtGetNextProcess( $processHandle , [ProcessAccessRights]::QueryInformation , [AttributeFlags]::None , 0 , [ref]$processHandle )
        if( $result -eq [NT_STATUS]::STATUS_SUCCESS -and $processHandle )
		{   
            ## Call extended version so we get handle information
	        $processInfo = New-Object -TypeName ProcessExtendedBasicInformation
            $size = [System.Runtime.InteropServices.Marshal]::SizeOf($processInfo)
            [IntPtr]$ptr = [Runtime.InteropServices.Marshal]::AllocHGlobal( $size )
            #need to set size field to size of extended structure
            $processInfo.Size = $size
            $marshalResult = [System.Runtime.InteropServices.Marshal]::StructureToPtr( $processInfo , $ptr , $false )
            [int]$returnedLength = 0
            $processQueryResult = [NtDll]::NtQueryInformationProcess( $processHandle , [PROCESS_INFO_CLASS]::ProcessBasicInformation , $ptr , $size , [ref]$returnedLength )
            if( $processQueryResult -eq [NT_STATUS]::STATUS_SUCCESS )
            {
                $processInfo = [System.Runtime.InteropServices.Marshal]::PtrToStructure( $ptr , [Type]$processInfo.GetType() )
                ## Not interested in self
                if( $processInfo -and $processInfo.BasicInfo.UniqueProcessId -ne $pid )
                {
                    ## Now get process name
                    $returnedLength = 0
                    [string]$processName = $null
                    ## Have to get native image file names here as win32 ones are missing for zombies
                    [NT_STATUS]$imageResult = [NtDll]::NtQueryInformationProcess( $processHandle , [PROCESS_INFO_CLASS]::ProcessImageFileName , $nameBuffer , $nameBufferSize , [ref]$returnedLength )
                    if( $imageResult -eq [NT_STATUS]::STATUS_SUCCESS -and $returnedLength )
                    {
                        $unicodeString = [System.Runtime.InteropServices.Marshal]::PtrToStructure( $nameBuffer , [Type]$path.GetType() )
                        if( $unicodeString -and ! [string]::IsNullOrEmpty( $unicodeString.buffer ) )
                        {
                            $processName = $unicodeString.Buffer
                            if( ! [string]::IsNullOrEmpty( $processName ) )
                            {
                                [string]$translated = $null
                                ## iterate ovoer local device paths to find this one
                                $drivePaths.GetEnumerator() | ForEach-Object `
                                {
                                    $regex = "^$([regex]::Escape( $_.Key ))\\"
                                    if( ! $translated -and $processName -match $regex )
                                    {
                                        $translated = $processName -replace $regex, $_.Value
                                    }
                                }
                                if( $translated )
                                {
                                    $processName = $translated
                                }
                            }
                        }
                    }
                    $sessionId = $null
                    $sessionInfo = New-Object -TypeName uint32
                    [int]$sessionSize = [System.Runtime.InteropServices.Marshal]::SizeOf( $sessionInfo )
                    [IntPtr]$sessionPtr = [Runtime.InteropServices.Marshal]::AllocHGlobal( $sessionSize )
                    $returnedLength = 0
                    $processQueryResult = [NtDll]::NtQueryInformationProcess( $processHandle , [PROCESS_INFO_CLASS]::ProcessSessionInformation , $sessionPtr , $sessionSize , [ref]$returnedLength )
                    if( $processQueryResult -eq [NT_STATUS]::STATUS_SUCCESS -and $returnedLength )
                    {
                        $sessionId = [System.Runtime.InteropServices.Marshal]::ReadInt32( $sessionPtr ) ## [System.Runtime.InteropServices.Marshal]::PtrToStructure( $sessionPtr , [Type]$sessionInfo.GetType() )
                    }
                    Add-Member -InputObject $processInfo -NotePropertyMembers @{
                        'ProcessHandle' = $processHandle
                        'ProcessName' = $processName
                        'SessionId' = $sessionId
                    }
                    [Runtime.InteropServices.Marshal]::FreeHGlobal( $sessionSize )
                    $sessionSize = [IntPtr]::Zero
                    $processInfo
                    $processCount++
                }
            }
            [Runtime.InteropServices.Marshal]::FreeHGlobal( $ptr )
            $ptr = [IntPtr]::Zero
		}
    }
    while( $result -eq [NT_STATUS]::STATUS_SUCCESS )
    
    [Runtime.InteropServices.Marshal]::FreeHGlobal( $nameBuffer )
    $nameBuffer = [IntPtr]::Zero
}

Function Get-ProcessInfo
{
    Param
    (
        $process ,
        $logonTime ## from quser, if null then they have no current session
    )

    Add-Member -InputObject $process -MemberType NoteProperty -Name 'Logged On Now' -Value $(if ($logonTime -ne $null) { 'Yes' } else { 'No' } )

    [string]$loggedOnUser = $null
    $logon = $logons[ $process.SessionId ]
    $logoff = $logoffs[ $process.SessionId ]
    if( ! $logon )
    {
        ## we want a more precise time than quser gives
        $logonevent = Get-WinEvent -FilterHashtable @{ LogName = $logname ; id = 1 } -ErrorAction SilentlyContinue | Where-Object { $_.Message -match "user logon notification on session $($process.SessionId)\." }  | Select -First 1
        if( ! $logonevent )
        {
            if( ! $logonTime )
            {
                [string]$Warning = "Unable to find logon event for session $($process.SessionId) in event log `"$logname`""
                if( ! $oldestEvent )
                {
                    $oldestEvent = Get-WinEvent -LogName $logname -Oldest -MaxEvents 1 -ErrorAction Continue
                }
                if( $oldestEvent )
                {
                    $Warning += ". Oldest event is @ $(Get-Date -Date $oldestEvent.TimeCreated -Format G)"
                }
                else
                {
                    $warning += ". This event log is empty or access is denied"
                }
                Write-Warning $warning
            }
            else
            {
                $logon = $logonTime
            }
        }
        else
        {
            $logon = $logonevent.TimeCreated
            $loggedOnUser = ([System.Security.Principal.SecurityIdentifier]($logonevent.UserId)).Translate([System.Security.Principal.NTAccount]).Value
            if( $loggedOnUser )
            {
                try
                {
                   $sessionOwners.Add( $process.SessionId , $loggedOnUser )
                }
                catch{} ## already got it
            }
        }

        $logons.Add( $process.SessionId , $logon )

        if( ! $logonTime -and  ! $logoff ) ## only look for logoff time if the user is not currently logged on
        {
            $logoffevent = Get-WinEvent -FilterHashtable @{ LogName = $logname ; id = 4 } -ErrorAction SilentlyContinue | Where-Object { $_.Message -match "user logoff notification on session $($process.SessionId)\." }  | Select -First 1
            if( ! $logoffevent )
            {
                [string]$Warning = "Unable to find logoff event for session $($process.SessionId) in event log `"$logname`""
                if( ! $oldestEvent )
                {
                    $oldestEvent = Get-WinEvent -LogName $logname -Oldest -MaxEvents 1 -ErrorAction Continue
                }
                if( $oldestEvent )
                {
                    $Warning += ". Oldest event is @ $(Get-Date -Date $oldestEvent.TimeCreated -Format G)"
                }
                else
                {
                    $warning += ". This event log is empty or access is denied"
                }
                Write-Warning $warning
            }
            else
            {
                $logoff = $logoffevent.TimeCreated
                $logoffs.Add( $process.SessionId , $logoff )
            }
        }
    }
    else
    {
        $loggedOnUser = $sessionOwners[ $process.SessionId ]
    }

    if( $logon )
    {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name 'Logon Time' -Value $logon
    }
    if( $logoff )
    {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name 'Logoff Time' -Value $logoff
    }
    if( $loggedOnUser -and $loggedOnUser -ne $process.UserName )
    {
        Add-Member -InputObject $_ -MemberType NoteProperty -Name 'Logged on User' -Value $loggedOnUser
    }
    $process
}

$compilerParameters = New-Object System.CodeDom.Compiler.CompilerParameters
$compilerParameters.CompilerOptions = '-unsafe'

Add-Type -ErrorAction Stop -CompilerParameters $compilerParameters -TypeDefinition @'
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
    
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    unsafe public struct WTSCLIENTA {
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
      public ushort  ClientAddressFamily;
      [MarshalAs(UnmanagedType.ByValArray, SizeConst = 31 , ArraySubType = UnmanagedType.U2)]
      public ushort[] ClientAddress;
      public ushort HRes;
      public ushort VRes;
      public ushort ColorDepth;
      [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 261)]
      public string   ClientDirectory;
      public ulong  ClientBuildNumber;
      public ulong  ClientHardwareId;
      public ushort ClientProductId;
      public ushort OutBufCountHost;
      public ushort OutBufCountClient;
      public ushort OutBufLength;
      [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 261)]
      public string   DeviceId;
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
    $wtsClientInfo = New-Object -TypeName 'WTSCLIENTA'
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
                        }
                        [wtsapi]::WTSFreeMemory( $ppQueryInfo )
                        $ppQueryInfo = [IntPtr]::Zero
                     }
                     else
                     {
                        Write-Error "$($machineName): $LastError"
                     }
                     $retval = [wtsapi]::WTSQuerySessionInformationW( $serverHandle , $element.SessionID , [WTS_INFO_CLASS]::WTSClientInfo , [ref]$ppQueryInfo , [ref]$ppBytesReturned );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                     if( $retval -and $ppQueryInfo )
                     {
                        $value = [system.runtime.interopservices.marshal]::PtrToStructure( $ppQueryInfo , [Type]$wtsClientInfo.GetType())
                        if( $value -and $wtsinfo )
                        {
                        }
                        [wtsapi]::WTSFreeMemory( $ppQueryInfo )
                        $ppQueryInfo = [IntPtr]::Zero
                     }
                     if( $wtsInfo )
                     {
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

if( $PSVersionTable.PSVersion.Major -lt 3 )
{
    Throw "This script must be run using PowerShell version 3.0 or higher, this is $($PSVersionTable.PSVersion.ToString())"
}

##    https://undocumented.ntinternals.net/
##    https://www.geoffchappell.com/studies/windows/km/ntoskrnl/api/ps/psquery/class.htm

Add-Type "
using System;
using System.Runtime.InteropServices;

    public static class NtDll
    {
        [DllImport(`"ntdll.dll`")]
        public static extern NT_STATUS NtQuerySystemInformation(
            [In] SYSTEM_INFORMATION_CLASS SystemInformationClass,
            [In] IntPtr SystemInformation,
            [In] int SystemInformationLength,
            [Out] out int ReturnLength);

        [DllImport(`"ntdll.dll`")]
        public static extern NT_STATUS NtQueryInformationProcess(
            [In] IntPtr ProcessHandle,
            [In] PROCESS_INFO_CLASS ProcessInformationClass,
            [In] IntPtr ProcessInformation,
            [In] int ProcessInformationLength,
            [Out] out int ReturnLength);
            
        [DllImport(`"ntdll.dll`")]
        public static extern NT_STATUS NtQueryInformationThread(
            [In] IntPtr ThreadHandle,
            [In] THREAD_INFO_CLASS ThreadInformationClass,
            [In] IntPtr ThreadInformation,
            [In] int ThreadInformationLength,
            [Out] out int ReturnLength);

        [DllImport(`"ntdll.dll`")]
        public static extern NT_STATUS NtGetNextProcess(
          [In] IntPtr ProcessHandle,
          [In] ProcessAccessRights DesiredAccess,
          [In] AttributeFlags HandleAttributes,
          [In] int Flags,
          [Out] out IntPtr NewProcessHandle);

        [DllImport(`"ntdll.dll`")]
        public static extern NT_STATUS NtGetNextThread(
          [In] IntPtr ProcessHandle,
          [In] IntPtr ThreadHandle,
          [In] ThreadAccessRights DesiredAccess,
          [In] AttributeFlags HandleAttributes,
          [In] int Flags,
          [Out] out IntPtr NewThreadHandle);
    }

	public static class Kernel32
	{
        [DllImport(`"kernel32.dll`", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CloseHandle(
            [In] IntPtr hHandle );
            
		[DllImport(`"kernel32.dll`", SetLastError = true)]
			public static extern uint QueryDosDevice(string lpDeviceName, System.Text.StringBuilder lpTargetPath, int ucchMax);
	}

    public static class Msvcrt
    {
		    [DllImport(`"msvcrt.dll`", SetLastError = true)]
			public static extern IntPtr memset(
                [In] IntPtr dest ,
                [In] int c ,
                [In] int count );
    }
    
    [StructLayout(LayoutKind.Sequential)]
    public struct UNICODE_STRING
    {
        public UInt16 Length;
        public UInt16 MaximumLength;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string buffer;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    public struct PROCESS_SESSION_INFORMATION
    {
        public ulong SessionId ;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct ProcessBasicInformation
    {
        public int ExitStatus;
        public IntPtr PebBaseAddress;
        public IntPtr AffinityMask;
        public int BasePriority;
        public IntPtr UniqueProcessId;
        public IntPtr InheritedFromUniqueProcessId;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    public struct CLIENT_ID
    {
        public IntPtr UniqueProcess;
        public IntPtr UniqueThread;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct ThreadBasicInformation
    {
        public int ExitStatus;
        public IntPtr TebBaseAddress;
        public CLIENT_ID ClientId;
        public IntPtr AffinityMask;
        public int Priority;
        public int BasePriority;
    }

    [StructLayout(LayoutKind.Sequential)]
    public class ProcessExtendedBasicInformation
    {
        public IntPtr Size;
        public ProcessBasicInformation BasicInfo;
        public ProcessExtendedBasicInformationFlags Flags;
    }

	[StructLayout(LayoutKind.Sequential)]
    public struct SystemHandleEntry
    {
        public int OwnerProcessId;
        public byte ObjectTypeNumber;
        public byte Flags;
        public ushort Handle;
        public IntPtr Object;
        public int GrantedAccess;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    public struct SystemHandleTableInfoEntryEx
    {
        public UIntPtr Object;
        public IntPtr UniqueProcessId;
        public IntPtr HandleValue;
        public uint GrantedAccess;
        public ushort CreatorBackTraceIndex;
        public ushort ObjectTypeIndex;
        public uint HandleAttributes;
        public uint Reserved;
    }

	public enum SYSTEM_INFORMATION_CLASS
    {
        SystemBasicInformation = 0,
        SystemPerformanceInformation = 2,
        SystemTimeOfDayInformation = 3,
        SystemProcessInformation = 5,
        SystemProcessorPerformanceInformation = 8,
        SystemHandleInformation = 16,
        SystemInterruptInformation = 23,
        SystemExceptionInformation = 33,
        SystemRegistryQuotaInformation = 37,
        SystemLookasideInformation = 45 ,
        SystemExtendedHandleInformation = 64
    }
    
    public enum THREAD_INFO_CLASS
    {
        ThreadBasicInformation
    }

    public enum PROCESS_INFO_CLASS
    {
        ProcessBasicInformation = 0 ,
        ProcessDebugPort = 7 ,
        ProcessSessionInformation = 24 ,
        ProcessWow64Information = 26 ,
        ProcessImageFileName = 27 ,
        ProcessBreakOnTermination = 29 ,
        ProcessImageFileNameWin32 = 43 ,
        ProcessHandleTable = 58 ,
        ProcessSubsystemInformation = 75 
    }

	public enum OBJECT_INFORMATION_CLASS
    {
        ObjectBasicInformation = 0,
        ObjectNameInformation = 1,
        ObjectTypeInformation = 2,
        ObjectAllTypesInformation = 3,
        ObjectHandleInformation = 4
    }
    
    public enum ProcessExtendedBasicInformationFlags
    {
        None = 0,
        IsProtectedProcess = 0x00000001,
        IsWow64Process = 0x00000002,
        IsProcessDeleting = 0x00000004,
        IsCrossSessionCreate = 0x00000008,
        IsFrozen = 0x00000010,
        IsBackground = 0x00000020,
        IsStronglyNamed = 0x00000040,
        IsSecureProcess = 0x00000080,
        IsSubsystemProcess = 0x00000100,
    }
    
    public enum GenericAccessRights : uint
    {
        None = 0,
        Access0 = 0x00000001,
        Access1 = 0x00000002,
        Access2 = 0x00000004,
        Access3 = 0x00000008,
        Access4 = 0x00000010,
        Access5 = 0x00000020,
        Access6 = 0x00000040,
        Access7 = 0x00000080,
        Access8 = 0x00000100,
        Access9 = 0x00000200,
        Access10 = 0x00000400,
        Access11 = 0x00000800,
        Access12 = 0x00001000,
        Access13 = 0x00002000,
        Access14 = 0x00004000,
        Access15 = 0x00008000,
        Delete = 0x00010000,
        ReadControl = 0x00020000,
        WriteDac = 0x00040000,
        WriteOwner = 0x00080000,
        Synchronize = 0x00100000,
        AccessSystemSecurity = 0x01000000,
        MaximumAllowed = 0x02000000,
        GenericAll = 0x10000000,
        GenericExecute = 0x20000000,
        GenericWrite = 0x40000000,
        GenericRead = 0x80000000,
    }

    public enum ProcessAccessRights : uint
    {
        None = 0,
        Terminate = 0x0001,
        CreateThread = 0x0002,
        SetSessionId = 0x0004,
        VmOperation = 0x0008,
        VmRead = 0x0010,
        VmWrite = 0x0020,
        DupHandle = 0x0040,
        CreateProcess = 0x0080,
        SetQuota = 0x0100,
        SetInformation = 0x0200,
        QueryInformation = 0x0400,
        SuspendResume = 0x0800,
        QueryLimitedInformation = 0x1000,
        SetLimitedInformation = 0x2000,
        AllAccess = 0x1FFFFF,
        GenericRead = GenericAccessRights.GenericRead,
        GenericWrite = GenericAccessRights.GenericWrite,
        GenericExecute = GenericAccessRights.GenericExecute,
        GenericAll = GenericAccessRights.GenericAll,
        Delete = GenericAccessRights.Delete,
        ReadControl = GenericAccessRights.ReadControl,
        WriteDac = GenericAccessRights.WriteDac,
        WriteOwner = GenericAccessRights.WriteOwner,
        Synchronize = GenericAccessRights.Synchronize,
        MaximumAllowed = GenericAccessRights.MaximumAllowed,
        AccessSystemSecurity = GenericAccessRights.AccessSystemSecurity
    }
    
    public enum ThreadAccessRights : uint
    {
        Terminate = 0x0001 ,
        SuspendResume = 0x0002 ,
        GetContext = 0x0008 ,
        SetContext = 0x0010 ,
        QueryInformation = 0x0040 ,
        SetInformation = 0x0020 ,
        SetThreadToken = 0x0080 ,
        Impersonate = 0x0100 ,
        DirectImpersonation = 0x0200 ,
        SetLimitedInformation = 0x0400 ,
        QueryLimitedInformation = 0x0800 ,
        Resume = 0x1000
    }

    public enum AttributeFlags : uint
    {
        None = 0,
        Inherit = 0x00000002,
        Permanent = 0x00000010,
        Exclusive = 0x00000020,
        CaseInsensitive = 0x00000040,
        OpenIf = 0x00000080,
        OpenLink = 0x00000100,
        KernelHandle = 0x00000200,
        ForceAccessCheck = 0x00000400,
        IgnoreImpersonatedDevicemap = 0x00000800,
        DontReparse = 0x00001000,
    }
	public enum NT_STATUS
    {
        STATUS_SUCCESS = 0x00000000,
        STATUS_BUFFER_OVERFLOW = unchecked((int)0x80000005L),
        STATUS_INFO_LENGTH_MISMATCH = unchecked((int)0xC0000004L)
    }
" -ErrorAction Stop

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

## Get all current user sessions so we can find processes in non-existent sessions 
[hashtable]$sessions = @{}
[hashtable]$sessionOwners = @{}
$oldestEvent = $null

[array]$wtssessions = @( Get-WTSSessionInformation )

ForEach( $wtssession in $wtssessions )
{
    $sessions.Add( $wtssession.SessionId , $wtssession.LogonTime )
}

Write-Verbose "Got $($sessions.Count) sessions"

[hashtable]$logons = @{}
[hashtable]$logoffs = @{}

$zombies = New-Object -TypeName System.Collections.Generic.List[PSObject]
[bool]$suspended = $false
[hashtable]$sessionsWithZombies = @{}

## Grab all processes now so we can get the flags for UWP processes to see if they are frozen which is normal behaviour
[array]$allProcesses = Get-AllProcesses
[hashtable]$frozenUWPProcesses = @{}

ForEach( $item in $allProcesses )
{
    if( ( [int]$item.Flags -band [int][ProcessExtendedBasicInformationFlags]::IsFrozen ) -eq [int][ProcessExtendedBasicInformationFlags]::IsFrozen )
    {
        $frozenUWPProcesses.Add( [int]$item.BasicInfo.UniqueProcessId , $item )
    }
}

## ignore some session 0 processes as they're special. Also ignore anything newly created in case it is in the process of being created.
$zombies += @( Get-Process -IncludeUserName | Where-Object { $_.StartTime -lt (Get-Date).AddSeconds( -$recentlyCreated ) `
    -and ( $suspended = ( $_.Threads | Where-Object { $_.ThreadState -eq 'Wait' -and $_.WaitReason -eq 'Suspended' } | Measure-Object | Select -ExpandProperty Count) -eq $_.Threads.Count )`
        -or ! $_.HandleCount -or ( $_.SessionId -gt 1 -and ! $sessions[ $_.SessionId ] ) } | ForEach-Object `
{
    if( $_.SessionId -or ( ! $_.SessionId -and $_.Name -notin $excludedProcesses))
    {
        [bool]$probablyExpected = $false
        if( $suspended )
        {
            if( $_.Path )
            {
                $probablyExpected = ( $frozenUWPProcesses[ [int]$_.Id ] -ne $null )
            }
            else
            {
                Write-Verbose "No path for $($_.Name)"
            }
        }
        [string]$vendor = '-'
        if( $_.Path )
        {
            $vendor = Get-ItemProperty -Path $_.Path -ErrorAction SilentlyContinue | Select -ExpandProperty VersionInfo -ErrorAction SilentlyContinue | Select -ExpandProperty CompanyName -ErrorAction SilentlyContinue
        }
        if( ! $probablyExpected )
        {
            [string]$processName = $_.Name + '.exe'
            [string]$loadedModules = if( $_.Modules ) { $_.Modules | Where-Object { $_.ModuleName -ne $processName } | Measure-Object | Select -ExpandProperty Count } else {'-'}
            Add-Member -InputObject $_ -NotePropertyMembers @{
                'Suspended' = $( if( $suspended ) { 'Yes' } else { 'No' } )
                'Vendor' = $vendor
                'Loaded Modules' = $loadedModules }

            if( $_.SessionId )
            {
                Get-ProcessInfo -process $_ -logonTime $sessions[ $_.SessionId ]
            }
            else  ## don't try and get session logon/logoff details for session 0 since cannot be logged on to
            {
                Add-Member -InputObject $_ -MemberType NoteProperty -Name 'Logged On Now' -Value '-'
                $_
            }
            try
            {
                $sessionsWithZombies.Add( $_.SessionId , $(if( $sessions[ $_.SessionId ] ) { 'Yes' } else { 'No' } ) )
            }
            catch{}
        }
    }
})

If( $zombies -and $zombies.Count )
{
    $outputFields = New-Object System.Collections.ArrayList
    $outputFields += @{n='Process';e={$_.Name}},@{n='Pid';e={$_.Id}},@{n='Session Id';e={$_.SessionId}},'Vendor','Username','Suspended'
    if( ( $zombies | Where { $_.PsObject.properties[ 'Logged on User' ] -and $_.'Logged on User' } | Measure-Object | Select -ExpandProperty Count ) -gt 0 )
    {
        $outputFields += 'Logged on User'
    }
    $outputFields += 'Logged On Now',@{n='Process Start Time';e={$_.StartTime}},'Loaded Modules','Handles',@{n='Threads';e={$_.Threads.Count}},@{n='Working Set (KB)';e={[math]::Round( $_.WS/1KB)}},@{n='Total CPU (s)';e={[math]::Round( $_.TotalProcessorTime.TotalSeconds , 2 )}},'Logon Time','Logoff Time'
    "Found $($zombies.Count) user mode zombie processes in $($sessionsWithZombies.Count) sessions of which $($sessionsWithZombies.GetEnumerator()|Where-Object { $_.Value -eq 'no' }|Measure-Object|Select -ExpandProperty Count) no longer exist"
    $zombies | Sort -Property 'SessionId' | Format-Table -Property $outputFields -AutoSize
    If( $kill )
    {
        [int]$killed = 0
        ForEach( $zombie in $zombies )
        {
            $killedProcess = Stop-Process -Id $zombie.Id -Force -PassThru
            if( $? -and $killedProcess -and $killedProcess.HasExited )
            {
                $killed++
            }
        }
        "Killed $killed processes"
    }
}
Else
{
    Write-Output 'No zombie processes found'
}

