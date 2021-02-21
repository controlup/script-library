<#
    Use low level Windows APIs to find processes which are exiting and see which processes have handles open to them

    @guyrleech 2019

    With ideas from:

    https://undocumented.ntinternals.net/
    https://www.geoffchappell.com/studies/windows/km/ntoskrnl/api/ps/psquery/class.htm

    Modification History:

    18/01/19   GRL  Added thread handle analysis
    22/12/20   GRL  Added code to detect known leaky processes and show details of services hosted in any leaking svchost.exe processes
    23/12/20   GRL  Code optimisation, formatting changes
    24/12/20   GRL  Check for User Input Delay if CUAgent leaks as known MS issue, sort leaky service names
    27/12/20   GRL  Change message for 8.2+ CUAgent non-fault leak
    29/12/20   GRL  Replaced quser.exe parsing with calling WTS APIs via P/Invoke, added handle closing code from "Show Zombie Processes"
    04/01/21   GRL  Updated kill code
    05/01/21   GRL  Kill switch disabled
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$false,HelpMessage='Whether to kill zombies found or not')]
    [string]$killParameter = 'False'
)

[bool]$kill = [System.Convert]::ToBoolean( $killParameter )
[int]$outputWidth = 250 ## we want some of the columns to wrap as we use newlines as delimiters not commas. This number works nicely on an HD (1920x1080) number with default CU console font sizes
[int]$recentlyCreated = 30 ## seconds within which we ignore processes created

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

if( $PSVersionTable.PSVersion.Major -lt 3 )
{
    Write-Warning "This script must be run using PowerShell version 3.0 or higher, this is $($PSVersionTable.PSVersion.ToString())"
    Exit
}

$warnings = New-Object -TypeName System.Collections.Generic.List[string]
$dyingProcesses = New-Object -TypeName System.Collections.Generic.List[object]
$failedToDieProcesses = New-Object -TypeName System.Collections.Generic.List[object]

## Now use low level APIs to find processes which are trying to exit but are unable to do so because other processes have handles still open to them
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

function Get-AllHandles( )
{
    Param
    (
         [int]$onlyPid = -1 ,
         [int[]]$handleType = @()
    )
    $length = 0x10000
    $ptr = [IntPtr]::Zero

    while ($true)
    {
        $ptr = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($length)
        $wantedLength = 0
        $result = [NtDll]::NtQuerySystemInformation([SYSTEM_INFORMATION_CLASS]::SystemExtendedHandleInformation, $ptr, $length, [ref] $wantedLength)
        if ($result -eq [NT_STATUS]::STATUS_INFO_LENGTH_MISMATCH)
        {
            $length = [Math]::Max($length, $wantedLength)
            [System.Runtime.InteropServices.Marshal]::FreeHGlobal($ptr)
            $ptr = [IntPtr]::Zero
        }
        elseif ($result -eq [NT_STATUS]::STATUS_SUCCESS)
		{
            break
		}
        else
		{
            throw (New-Object System.ComponentModel.Win32Exception)
		}
    }

	if ([IntPtr]::Size -eq 4)
	{
		$handleCount = [System.Runtime.InteropServices.Marshal]::ReadInt32($ptr)
	}
	else
	{
		$handleCount = [System.Runtime.InteropServices.Marshal]::ReadInt64($ptr)
	}

	$offset = [IntPtr]::Size * 2 ## Reserver value after count
	$She = New-Object -TypeName SystemHandleTableInfoEntryEx
    $size = [System.Runtime.InteropServices.Marshal]::SizeOf($She)

    for ($i = 0; $i -lt $handleCount; $i++)
    {
        $thisHandle = [SystemHandleTableInfoEntryEx][System.Runtime.InteropServices.Marshal]::PtrToStructure([IntPtr]([long]$ptr + $offset),[Type]$She.GetType())
            
        if( ( ! $handleType -or ! $handleType.Count -or $handleType -contains $thisHandle.ObjectTypeIndex ) `
            -and ( $onlyPid -lt 0 -or $thisHandle.UniqueProcessId -eq $onlyPid ) )
        {
            $thisHandle
        }
        $offset += $size
    }

    if ($ptr -ne [IntPtr]::Zero)
	{
        [System.Runtime.InteropServices.Marshal]::FreeHGlobal($ptr)
        $ptr = [IntPtr]::Zero
	}
}

function Get-AllThreadsForProcess( [IntPtr]$processHandle )
{
    [int]$threadCount = 0
    [IntPtr]$threadHandle = [IntPtr]::Zero
    do
    {
        [NT_STATUS]$result = [NtDll]::NtGetNextThread( $processHandle , $threadHandle , [ThreadAccessRights]::QueryLimitedInformation , [AttributeFlags]::None , 0 , [ref]$threadHandle )
        if( $result -eq [NT_STATUS]::STATUS_SUCCESS -and $threadHandle )
		{
            $threadCount++
            ## call NtQueryInformationThread and get THREAD_BASIC_INFORMATION structure
            ## once we have open handles too all threads, we can iterate over the ones for this process later and match object ids to ones in all handle list to find pid like we do for processes
            
            $threadBasicInfo = New-Object 'ThreadBasicInformation'
            [int]$size = [System.Runtime.InteropServices.Marshal]::SizeOf($threadBasicInfo)
            [int]$returnedLength = 0
            [IntPtr]$ptr = [Runtime.InteropServices.Marshal]::AllocHGlobal( $size )
            $threadQueryResult = [NtDll]::NtQueryInformationThread( $threadHandle , [THREAD_INFO_CLASS]::ThreadBasicInformation , $ptr , $size , [ref]$returnedLength )
            if( $threadQueryResult -eq [NT_STATUS]::STATUS_SUCCESS )
            {
                $threadBasicInfo = [System.Runtime.InteropServices.Marshal]::PtrToStructure( $ptr , [Type]$threadBasicInfo.GetType() )
                Add-Member -InputObject $threadBasicInfo -NotePropertyMembers @{
                    'ThreadHandle' = $threadHandle
                }
                $threadBasicInfo
            }
        }
    } while( $result -eq [NT_STATUS]::STATUS_SUCCESS )
}


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
[string]$logname = 'Microsoft-Windows-User Profile Service/Operational'
[string[]]$excludedProcesses = @( 'System' , 'Idle' , 'Memory Compression' , 'Secure System' , 'Registry' )

Function Get-ProcessInfo
{
    Param
    (
        $process ,
        $logonTime ## from WTS session info, if null then they have no current session
    )

    Add-Member -InputObject $process -MemberType NoteProperty -Name 'Logged On Now' -Value $(if ($logonTime -ne $null) { 'Yes' } else { 'No' } )

    [string]$loggedOnUser = $null
    $logon = $logons[ $process.SessionId ]
    $logoff = $logoffs[ $process.SessionId ]
    if( ! $logon )
    {
        ## we want a more precise time than WTS Session info gives
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

        if( $logon )
        {
            $logons.Add( $process.SessionId , $logon )
        }

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

## Get all current user sessions so we can find processes in non-existent sessions 
[hashtable]$sessions = @{}
[hashtable]$sessionOwners = @{}
$oldestEvent = $null

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

[array]$wtssessions = @( Get-WTSSessionInformation )

ForEach( $wtssession in $wtssessions )
{
    $sessions.Add( $wtssession.SessionId , $wtssession.LogonTime )
}

Write-Verbose "Got $($sessions.Count) sessions"

[hashtable]$logons = @{}
[hashtable]$logoffs = @{}

$zombies = New-Object -typename System.Collections.Generic.List[psobject]
[bool]$suspended = $false
[hashtable]$sessionsWithZombies = @{}
[hashtable]$allGetProcessProcesses = @{}
[hashtable]$allCIMProcesses = @{}

## "snapshot" processes 
Get-Process -IncludeUserName | . { Process { $allGetProcessProcesses.Add( [int]$_.Id , $_ ) } }
Get-CimInstance -ClassName win32_process | . { Process { $allCIMProcesses.Add( [int]$_.ProcessId , $_ ) } }

## Grab all processes now so we can get the flags for UWP processes to see if they are frozen which is normal behaviour
[array]$allProcesses = Get-AllProcesses
[hashtable]$frozenUWPProcesses = @{}
$allThreads = New-Object System.Collections.ArrayList

ForEach( $item in $allProcesses )
{
    $allThreads += Get-AllThreadsForProcess -processHandle $item.ProcessHandle -processPid $item.BasicInfo.UniqueProcessId

    if( ( [int]$item.Flags -band [int][ProcessExtendedBasicInformationFlags]::IsFrozen ) -eq [int][ProcessExtendedBasicInformationFlags]::IsFrozen )
    {
        $frozenUWPProcesses.Add( [int]$item.BasicInfo.UniqueProcessId , $item )
    }
}

[array]$allProcessAndThreadHandles = Get-AllHandles -handleType 7,8 ## Only Process and Thread handles
[int]$uwpFrozenApps = 0

## ignore some session 0 processes as they're special. Also ignore anything newly created in case it is in the process of being created or was deliberately created suspended.
$zombies += @( Get-Process -IncludeUserName | Where-Object { $_.StartTime -lt (Get-Date).AddSeconds( -$recentlyCreated ) `
    -and ( $suspended = ( $_.Threads.Where( { $_.ThreadState -eq 'Wait' -and $_.WaitReason -eq 'Suspended' } )).Count -eq $_.Threads.Count )`
        -or ! $_.HandleCount -or ( $_.SessionId -gt 1 -and ! $sessions[ $_.SessionId ] ) } | . { Process `
{
    if( $_.SessionId -or ( $_.SessionId -eq 0 -and $excludedProcesses -notcontains $_.Name ))
    {
        [bool]$probablyExpected = $false
        if( $suspended )
        {
            if( $_.Path )
            {
                if( $probablyExpected = ( $frozenUWPProcesses[ [int]$_.Id ] -ne $null ) )
                {
                    $uwpFrozenApps++
                }
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
        else
        {
            Write-Verbose "Excluding $($_.Name) ($($_.Id)) as frozen flag set so probably UWP"
        }
    }}
})

[datetime]$startedKilling = [datetime]::Now

if( $zombies -and $zombies.Count )
{
    $outputFields = New-Object -TypeName System.Collections.ArrayList
    $outputFields += @{n='Process';e={$_.Name}},@{n='Pid';e={$_.Id}},@{n='Session Id';e={$_.SessionId}},'Vendor','Username','Suspended'
    if( ( $zombies | Where-Object { $_.PsObject.properties[ 'Logged on User' ] -and $_.'Logged on User' } | Measure-Object | Select -ExpandProperty Count ) -gt 0 )
    {
        $outputFields += 'Logged on User'
    }
    $outputFields += 'Logged On Now',@{n='Process Start Time';e={$_.StartTime}},'Loaded Modules','Handles',@{n='Threads';e={$_.Threads.Count}},@{n='Working Set (KB)';e={[math]::Round( $_.WS/1KB)}},@{n='Total CPU (s)';e={[math]::Round( $_.TotalProcessorTime.TotalSeconds , 2 )}},'Logon Time','Logoff Time'
    "Found $($zombies.Count) user mode zombie processes in $($sessionsWithZombies.Count) sessions of which $($sessionsWithZombies.GetEnumerator() | Where-Object Value -eq 'no' | Measure-Object | Select-Object -ExpandProperty Count) no longer exist"
    $zombies | Sort-Object -Property 'SessionId' | Format-Table -Property $outputFields -AutoSize

    if( $kill )
    {
        ForEach( $zombie in $zombies )
        {
            if( $zombie.hasExited )
            {
                Add-Member -InputObject $zombie -MemberType NoteProperty -Name 'WasDying' -Value ([datetime]::Now)
            }
            ## ignore errors as we explicitly check for process exit
            Stop-Process -Id $zombie.Id -Force
        }
    }
}
else
{
    Write-Output "No processes found with all threads suspended (except $uwpFrozenApps UWP apps which is normal), no open handles or in a non-existent session"
}

## We have a handle open to all processes as a result of calling Get-AllProcesses so we get all handles for this PowerShell process and for each process from Get-AllProcesses we look for Flags of
## IsProcessDeleting and for each of those we find the handle we saved for it in the handle list for this process and then the Object property is looked up against all process handles which gives us
## the pid (UniqueProcessId) of the pid that has the open handle

Write-Verbose -Message "$(Get-Date -Format G): getting all process and thread handles"

[hashtable]$handles = @{}
$allProcessAndThreadHandles.Where( { $_.UniqueProcessId -eq $pid } ) | . { Process `
{
    $handles.Add( $_.HandleValue , $_ )
}}

[hashtable]$leakers = @{}

[int]$exitingProcesses = 0
[int]$zero = 0
[uint16]$seven = 7
[uint16]$eight = 8
[int]$frozenFlag = [ProcessExtendedBasicInformationFlags]::IsFrozen

Write-Verbose -Message "$(Get-Date -Format G): getting results"
[array]$results = @( $allProcesses.Where( { ( ( [int]$_.Flags -band [int][ProcessExtendedBasicInformationFlags]::IsProcessDeleting ) -eq [int][ProcessExtendedBasicInformationFlags]::IsProcessDeleting ) } ) | . { Process `
{
    $deadProcess = $_
    $exitingProcesses++
    Write-Verbose ( "{0} ({1}) flags {2:x}" -f $deadProcess.ProcessName , $deadProcess.BasicInfo.UniqueProcessId , $deadProcess.Flags )
    if( ( [int]$_.Flags -band $frozenFlag ) -eq $zero )
    {
        if( $ourHandleToDeadProcess = $handles[ $deadProcess.ProcessHandle ] )
        {
            $object = $ourHandleToDeadProcess.Object
            [System.Collections.Generic.List[object]]$openHandles = @( $allProcessAndThreadHandles.Where( { $_.ObjectTypeIndex -eq $seven -and $_.Object -eq $object -and $_.UniqueProcessId -ne $pid } ) | Group-Object -Property UniqueProcessId )
  
            [int]$processHandleCount = $openHandles.Count
            [int]$threadHandleCount = 0
            
            [array]$openThreadHandles = @( $allThreads.Where( { $_.ClientId.UniqueProcess -eq $deadProcess.BasicInfo.UniqueProcessId } ) | . { Process `
            {
                $deadThread = $_
                ## Group by UniqueProcessId since nay existing elements in $openHandles will be so we must add the same type objects here
                [array]$leakedThreadHandles = @( $allProcessAndThreadHandles.Where( { $_.ObjectTypeIndex -eq $eight -and $_.Object -eq $handles[ $deadThread.ThreadHandle ].Object -and $_.UniqueProcessId -ne $pid } ) )
                if( $leakedThreadHandles.Count )
                {
                    $threadHandleCount += $leakedThreadHandles | Group-Object -Property Object | Measure-Object -Property Count -Sum | Select-Object -ExpandProperty Sum
                    $leakedThreadHandles
                }
            }} | Group-Object -Property UniqueProcessId)
            
            if( $openThreadHandles -and $openThreadHandles.Count )
            {
                $openHandles += $openThreadHandles
            }

            [string]$separator = $null
            [hashtable]$thisProcessesLeakers = @{}
        
            $openHandles | . { Process `
            {
                if( $_.Name )
                {
                    if( $thisLeaker = $thisProcessesLeakers[ $_.Name ] )
                    {
                        $thisProcessesLeakers.Set_Item( $_.Name , $thisLeaker + [int]$_.Count )
                    }
                    else
                    {
                        $thisProcessesLeakers.Add( $_.Name , $_.Count )
                    }
                    if( $leaker = $leakers[ $_.Name ] )
                    {
                        $leakers.Set_Item( $_.Name , $leaker + [int]$_.Count )
                    }
                    else
                    {
                        $leakers.Add( $_.Name , $_.Count )
                    }
                }
                else
                {
                    Write-Verbose "Name (pid) is null for $($deadProcess.Name)"
                }
            }}
            if( $thisProcessesLeakers -and $thisProcessesLeakers.Count )
            {
                Add-Member -InputObject $deadProcess -NotePropertyMembers @{
                    'Leakers' = $thisProcessesLeakers
                    'ProcessHandleCount' = $processHandleCount
                    'ThreadHandleCount' = $threadHandleCount }
                $deadProcess
            }
            else ## found no handles so may not be zombie per se
            {
                Write-Verbose "`tFound no handles for zombie $($deadProcess.ProcessName) ($($deadProcess.BasicInfo.UniqueProcessId))"
            }
        }
        else
        {
            Write-Warning "No open handle in pid $pid for handle $($deadProcess.ProcessHandle) in dead process $($deadProcess.ProcessName)"
        }
    }
    else
    {
        Write-Verbose "Excluding $($deadProcess.ProcessName) ($($deadProcess.BasicInfo.UniqueProcessId)) as frozen flag set so probably UWP"
    }
}})

Write-Verbose -Message "$(Get-Date -Format G): got results from $exitingProcesses terminating processes"

[int]$killed = 0
$sinceKilling = New-TimeSpan -Start $startedKilling -End ([datetime]::Now)

if( $kill -and $zombies -and $zombies.Count )
{
    ## check on state of previously killed processes

    ForEach( $zombie in $zombies )
    {
        if( ! ( Get-Process -InputObject $zombie -ErrorAction SilentlyContinue | Where-Object StartTime -lt $startedKilling ) )
        {
            $killed++
        }
        elseif( $zombie.PSObject.Properties[ 'WasDying' ] )
        {
            $dyingProcesses.Add( $zombie )
        }
        else
        {
            $failedToDieProcesses.Add( $zombie )
        }                
    }

    "Successfully killed $killed processes"
}

if( $results -and $results.Count )
{
    "`nSummary of the $($results.Count) dead processes which other processes still have open process or thread handles to:"
    [array]$grouped = @( $results | Group-Object -Property ProcessName | . { Process `
    {
        ## grouping has flattened the leakers so we need to summarise per grouped item which is a pid
        $deadProcess = $_
        [string]$separator = $null
        [string]$handleTypes = if( $deadProcess.Group.ProcessHandleCount )
        {
            'Process'
        }
        $handleTypes += if( $deadProcess.Group.ThreadHandleCount )
        {
            if( $deadProcess.Group.ProcessHandleCount )
            {
                ' &'
            }
            'Thread'
        }
        ## Need to amalgamate the pid, session and leakers data for each of these dead processes so we can display on a single output line
        [string]$offenders = ($deadProcess.group.Leakers | . { Process `
        { 
            $_.GetEnumerator()|select name,value 
        }} `
            | Group-Object -Property Name -AsHashTable).GetEnumerator() | Select -Property @{n='Pid';e={$_.Name}},@{n='TotalHandles';e={$_.value|Measure-Object -Property value -Sum|select -ExpandProperty Sum}} | . { Process `
        {
            "{0}{1} ({2}) {3} {4} handles" -f $separator , ( $allCIMProcesses[ [int]$_.Pid ] | Select -ExpandProperty Name ) , $_.Pid , $_.TotalHandles , $handleTypes
            $separator = "`n"
        }}
        
        ## See how many different sessions these are in
        [hashtable]$exisingSessions = @{}
        [hashtable]$nonexistentSessions = @{}
        $deadProcess.group | . { Process `
        {
            try
            {
                if( ! $_.SessionId -or $sessions[ $_.SessionId ] )
                {
                    $exisingSessions.Add( $_.SessionId , $_ )
                }
                else
                {
                    $nonexistentSessions.Add( $_.SessionId , $_ )
                }
            }
            catch {}
        }}
        Add-Member -InputObject $deadProcess -NotePropertyMembers @{
            'Live Sessions' = $exisingSessions.Count
            'Dead Sessions' = $nonexistentSessions.Count
            'Instances' = $deadProcess.Group.Count
            'Offenders' = $offenders
        }
        $deadProcess
    }})
    
    $grouped | Sort -Property Instances -Descending | Format-Table -Wrap -AutoSize -Property 'Name' , 'Instances' , 'Live Sessions','Dead Sessions',@{n='Processes with handles to this process';e={$_.Offenders -join "`n"}}

    if( $leakers -and $leakers.Count )
    {
        "The $($leakers.Count) processes causing these zombies are:"

        $leakers.GetEnumerator() | Select @{n='Executable';e={$allCIMProcesses[ [int]$_.Key ]|select -ExpandProperty ExecutablePath}},
            @{n='Company';e={$allGetProcessProcesses[ [int]$_.Key ]|Select -ExpandProperty Company}},
            @{n='Version';e={$allGetProcessProcesses[ [int]$_.Key ]|Select -ExpandProperty ProductVersion}},
            @{n='Pid';e={$_.Key}},
            @{n='Session Id';e={$allGetProcessProcesses[ [int]$_.Key ]|Select -ExpandProperty SessionId}},
            @{n='Username';e={$allGetProcessProcesses[ [int]$_.Key ]|Select -ExpandProperty Username}},
            @{n='Start Time';e={$allGetProcessProcesses[ [int]$_.Key ]|Select -ExpandProperty StartTime}},
            @{n='Handles to Zombies';e={$_.Value}} | Sort -Property 'Handles to Zombies' -Descending | Format-Table -AutoSize

        ## See if any are ControlUp processes and check versions as there was an issue, fixed in 8.2 RTM
        [string[]]$knownLeakyProcesses = @( 'cuagent' , 'AppLoadTimeTracer' )
        [hashtable]$svchostProcesses = @{}
        [hashtable]$services = @{}
        $leakyServices = New-Object -TypeName System.Collections.Generic.List[object]

        ## build hashtable where svchost pid is key and has array of service names(s) so can be sorted for display
        Get-WmiObject -Class win32_service -Filter "State = 'running'" | Where-Object { $_.PathName -match '\\svchost\.exe\b' } | ForEach-Object `
        {
            $service = $_
            if( $existingEntry = $services[ $service.ProcessId ] )
            {
                $existingEntry.Add( $service.DisplayName )
            }
            else
            {
                $services.Add( $service.ProcessId , [System.Collections.Generic.List[string]]$service.DisplayName )
            }
        }

        ForEach( $leaker in $leakers.GetEnumerator() )
        {
            if( ( $processDetails = $allGetProcessProcesses[ [int]$leaker.Key ] ) -and $null -ne $processDetails.PSObject.Properties[ 'Name' ] )
            {
                if( $processDetails.PSObject.Properties[ 'Company' ] -and $processDetails.Company -cmatch '^ControlUp' -and  $processDetails.Name -in $knownLeakyProcesses -and $processDetails.PSObject.Properties[ 'FileVersion' ] )
                {
                    if( [version]$processDetails.FileVersion -lt [version]'8.2' )
                    {
                        $warnings.Add( "ControlUp process $($processDetails.Name) version $($processDetails.FileVersion) has a handle leak, fixed in 8.2" )
                    }
                    elseif( Get-Counter -Counter "\User Input Delay per Session($(Get-Process -Id $pid|Select-Object -ExpandProperty SessionId))\Max Input Delay"  -ErrorAction SilentlyContinue )
                    {
                        $warnings.Add( "cuAgent.exe appears due to a known Microsoft leak with the User Input Delay performance metrics, until Microsoft fixes this issue, the workaround is to disable the User Input Delay feature via the registry. For more details please contact ControlUp Support" )
                    }
                }
                ## if services then we will list the service names for that pid
                if( $processDetails.Name -eq 'svchost' -and ($serviceNames = $services[ [uint32]$processDetails.Id ] ))
                {
                    $leakyServices.Add( (New-Object -TypeName PSCustomObject -Property (@{ 'Pid' = $processDetails.Id ; 'Services' = $serviceNames }) ) )
                }
            }
        }

        if( $leakyServices -and $leakyServices.Count )
        {
            'Leaking svchost process services are:'
            $leakyServices | Sort-Object -Property Pid | Select-Object -Property Pid,@{n='Services';e={( $_.Services | Sort-Object ) -join "`n"}} | Format-Table -Wrap -AutoSize
        }
    }
}
else
{
    Write-Output "`nFound no processes with handles open to dead processes"
}

## Close all handles otherwise parent PowerShell process will leak handles until exit which makes this script run slower if nothing else
$allProcesses | . { Process `
{
    [void][Kernel32]::CloseHandle( $_.ProcessHandle )
}}

$allThreads | . { Process `
{
    [void][Kernel32]::CloseHandle( $_.ThreadHandle )
}}

if( $warnings -and $warnings.Count )
{
    ''
    $warnings | Write-Warning
}

if( $dyingProcesses -and $dyingProcesses.Count )
{
    ''
    Write-Warning -Message "Got $($dyingProcesses.Count) processes which were already exiting but are still alive after being terminated (pids $(($dyingProcesses | Select-Object -ExpandProperty Id | Sort-Object) -join ','))"
}

if( $failedToDieProcesses -and $failedToDieProcesses.Count )
{
    ''
    Write-Warning -Message "Got $($failedToDieProcesses.Count) processes which didn't exit within $([math]::Round( $sinceKilling.TotalSeconds , 1 )) seconds of being terminated (pids $(($failedToDieProcesses | Select-Object -ExpandProperty Id | Sort-Object) -join ','))"
}

