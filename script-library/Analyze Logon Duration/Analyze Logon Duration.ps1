#Requires -version 3

<#
 	.SYNOPSIS
        An advanced function that gives you a break-down analysis of a user's most recent logon on the machine.
		
    .DESCRIPTION
        This function gives a detailed report on the logon process and its phases.
        Each phase documented have a column for duration in seconds, start time, end time
        and interim delay which is the time that passed between the end of one phase
        and the start of the one that comes after.
		
	.PARAMETER  <UserName <string[]>
		The user name the function reports for. The default is the user who runs the script.
		
	.PARAMETER	<UserDomain <string[]>
		The user domain name the function reports for. The default is the domain name of the user who runs the script.

	.PARAMETER  <HDXSessionId>560
		The Session ID of the user the function reports for. 
        Required for the "HDX Connection" phase,
        The machine the script runs on has to be part of the Citrix Site.

	.PARAMETER  <XDUsername>
        A User with administrative permissions to the Citrix XenApp/XenDesktop Site, at least Read-Only
        privileges, the machine the script runs on has to be part of the Citrix Site.
	
    .PARAMETER  <XDPassword>
		Password for the Citrix Site user provided.

	.PARAMETER  <CUDesktopLoadTime>
		Specifies the duration of the Shell phase, can be used with ControlUp as passed argument.

	.PARAMETER  <ClientName>
		Specifies the client name of the Citrix session.
    
    .NOTES
        The HDX duration is a new metric that requires changes to the ICA protocol. 
        This means that, if the new version of the client is not being used, the metrics returned are NULL.
        It may take a few seconds until the HDX duration is reported and available at the Delivery Controller.
		
    .LINK
        For more information refer to:
            http://www.controlup.com

    .LINK
        Stay in touch:
        http://twitter.com/nironkoren

    .EXAMPLE
        C:\PS> Get-LogonDurationAnalysis -UserName Rick
		
		Gets analysis of the logon process for the user 'Rick' in the current domain.
#>

## Last modified 1216 GMT 03/09/20 @guyrleech

## A mechanism to allow script use offline with saved event logs
[hashtable]$global:terminalServicesParams = @{ 'ProviderName' = 'Microsoft-Windows-TerminalServices-LocalSessionManager' }
[hashtable]$global:securityParams = @{ 'ProviderName' = 'Microsoft-Windows-Security-Auditing' }
[hashtable]$global:userProfileParams = @{ 'ProviderName' = 'Microsoft-Windows-User Profile Service' }
[hashtable]$global:groupPolicyParams = @{ 'ProviderName' = 'Microsoft-Windows-GroupPolicy' }
[hashtable]$global:scheduledTasksParams = @{ 'ProviderName' = 'Microsoft-Windows-TaskScheduler' }
[hashtable]$global:citrixUPMParams = @{ 'ProviderName' = 'Citrix Profile Management' }
[hashtable]$global:printServiceParams = @{ 'ProviderName' = 'Microsoft-Windows-PrintService' }
[hashtable]$global:AppVolumesParams = @{ 'ProviderName' = 'svservice' }
[hashtable]$global:windowsShellCoreParams = @{ 'ProviderName' = 'Microsoft-Windows-Shell-Core' }
[hashtable]$global:appReadinessParams = @{ 'ProviderName' = 'Microsoft-Windows-AppReadiness' }
[hashtable]$global:winlogonParams = @{ 'ProviderName' = 'Microsoft-Windows-Winlogon' }
[hashtable]$global:appReadinessParams = @{ 'ProviderName' = 'Microsoft-Windows-AppReadiness' }
[int]$global:windowsMajorVersion = [System.Environment]::OSVersion.Version.Major
[bool]$offline = $false
[int]$suggestedSecurityEventLogSizeMB = 100
[int]$outputWidth = 400
$script:warnings = New-Object -TypeName System.Collections.Generic.List[string]
[string]$global:appVolumesLogFile = "${env:ProgramFiles(x86)}\CloudVolumes\Agent\Logs\svservice.log"
[version]$global:appVolumesVersion = $null
[bool]$global:WaitForFirstVolumeOnly = $true
$script:ivantiEMNonBlockingPhases = New-Object -TypeName System.Collections.Generic.List[psobject]
$script:vmwareDEMNonBlockingPhases = New-Object -TypeName System.Collections.Generic.List[psobject]

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

$AuditDefinitions = @'
    /// The AuditFree function frees the memory allocated by audit functions for the specified buffer.
    /// https://msdn.microsoft.com/en-us/library/windows/desktop/aa375654(v=vs.85).aspx
    [DllImport("advapi32.dll")]
    public static extern void AuditFree(IntPtr buffer);

    /// The AuditQuerySystemPolicy function retrieves system audit policy for one or more audit-policy subcategories.
    /// https://msdn.microsoft.com/en-us/library/windows/desktop/aa375702(v=vs.85).aspx
    [DllImport("advapi32.dll", SetLastError = true)]
    public static extern bool AuditQuerySystemPolicy(Guid pSubCategoryGuids, uint PolicyCount, out IntPtr ppAuditPolicy);
        
    /// The AuditQuerySystemPolicy function retrieves system audit policy for one or more audit-policy subcategories.
    /// https://msdn.microsoft.com/en-us/library/windows/desktop/aa375702(v=vs.85).aspx</returns>
    [DllImport("advapi32.dll", SetLastError = true)]
    public static extern bool AuditSetSystemPolicy( IntPtr ppAuditPolicy , uint PolicyCount);

    /// The AUDIT_POLICY_INFORMATION structure specifies a security event type and when to audit that type.
    /// https://msdn.microsoft.com/en-us/library/windows/desktop/aa965467(v=vs.85).aspx
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct AUDIT_POLICY_INFORMATION
    {
        /// A GUID structure that specifies an audit subcategory.
        public Guid AuditSubCategoryGuid;
        /// A set of bit flags that specify the conditions under which the security event type specified by the AuditSubCategoryGuid and AuditCategoryGuid members are audited.
        public AUDIT_POLICY_INFORMATION_TYPE AuditingInformation;
        /// A GUID structure that specifies an audit-policy category.
        public Guid AuditCategoryGuid;
    }

    [Flags]
    public enum AUDIT_POLICY_INFORMATION_TYPE
    {
        None = 0,
        Success = 1,
        Failure = 2,
    }

    // from https://gallery.technet.microsoft.com/scriptcenter/Grant-Revoke-Query-user-26e259b0
    public enum Rights
    {
        SeTrustedCredManAccessPrivilege,             // Access Credential Manager as a trusted caller
        SeNetworkLogonRight,                         // Access this computer from the network
        SeTcbPrivilege,                              // Act as part of the operating system
        SeMachineAccountPrivilege,                   // Add workstations to domain
        SeIncreaseQuotaPrivilege,                    // Adjust memory quotas for a process
        SeInteractiveLogonRight,                     // Allow log on locally
        SeRemoteInteractiveLogonRight,               // Allow log on through Remote Desktop Services
        SeBackupPrivilege,                           // Back up files and directories
        SeChangeNotifyPrivilege,                     // Bypass traverse checking
        SeSystemtimePrivilege,                       // Change the system time
        SeTimeZonePrivilege,                         // Change the time zone
        SeCreatePagefilePrivilege,                   // Create a pagefile
        SeCreateTokenPrivilege,                      // Create a token object
        SeCreateGlobalPrivilege,                     // Create global objects
        SeCreatePermanentPrivilege,                  // Create permanent shared objects
        SeCreateSymbolicLinkPrivilege,               // Create symbolic links
        SeDebugPrivilege,                            // Debug programs
        SeDenyNetworkLogonRight,                     // Deny access this computer from the network
        SeDenyBatchLogonRight,                       // Deny log on as a batch job
        SeDenyServiceLogonRight,                     // Deny log on as a service
        SeDenyInteractiveLogonRight,                 // Deny log on locally
        SeDenyRemoteInteractiveLogonRight,           // Deny log on through Remote Desktop Services
        SeEnableDelegationPrivilege,                 // Enable computer and user accounts to be trusted for delegation
        SeRemoteShutdownPrivilege,                   // Force shutdown from a remote system
        SeAuditPrivilege,                            // Generate security audits
        SeImpersonatePrivilege,                      // Impersonate a client after authentication
        SeIncreaseWorkingSetPrivilege,               // Increase a process working set
        SeIncreaseBasePriorityPrivilege,             // Increase scheduling priority
        SeLoadDriverPrivilege,                       // Load and unload device drivers
        SeLockMemoryPrivilege,                       // Lock pages in memory
        SeBatchLogonRight,                           // Log on as a batch job
        SeServiceLogonRight,                         // Log on as a service
        SeSecurityPrivilege,                         // Manage auditing and security log
        SeRelabelPrivilege,                          // Modify an object label
        SeSystemEnvironmentPrivilege,                // Modify firmware environment values
        SeDelegateSessionUserImpersonatePrivilege,   // Obtain an impersonation token for another user in the same session
        SeManageVolumePrivilege,                     // Perform volume maintenance tasks
        SeProfileSingleProcessPrivilege,             // Profile single process
        SeSystemProfilePrivilege,                    // Profile system performance
        SeUnsolicitedInputPrivilege,                 // "Read unsolicited input from a terminal device"
        SeUndockPrivilege,                           // Remove computer from docking station
        SeAssignPrimaryTokenPrivilege,               // Replace a process level token
        SeRestorePrivilege,                          // Restore files and directories
        SeShutdownPrivilege,                         // Shut down the system
        SeSyncAgentPrivilege,                        // Synchronize directory service data
        SeTakeOwnershipPrivilege                     // Take ownership of files or other objects
    }
    public sealed class TokenManipulator
    {
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        internal struct TokPriv1Luid
        {
            public int Count;
            public long Luid;
            public int Attr;
        }

        internal const int SE_PRIVILEGE_DISABLED = 0x00000000;
        internal const int SE_PRIVILEGE_ENABLED = 0x00000002;
        internal const int TOKEN_QUERY = 0x00000008;
        internal const int TOKEN_ADJUST_PRIVILEGES = 0x00000020;

        internal sealed class Win32Token
        {
            [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
            internal static extern bool AdjustTokenPrivileges(
                IntPtr htok,
                bool disall,
                ref TokPriv1Luid newst,
                int len,
                IntPtr prev,
                IntPtr relen
            );

            [DllImport("kernel32.dll", ExactSpelling = true)]
            internal static extern IntPtr GetCurrentProcess();

            [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
            internal static extern bool OpenProcessToken(
                IntPtr h,
                int acc,
                ref IntPtr phtok
            );

            [DllImport("advapi32.dll", SetLastError = true)]
            internal static extern bool LookupPrivilegeValue(
                string host,
                string name,
                ref long pluid
            );

            [DllImport("kernel32.dll", ExactSpelling = true)]
            internal static extern bool CloseHandle(
                IntPtr phtok
            );
        }

        public static int AddPrivilege(Rights privilege)
        {
            bool retVal;
            TokPriv1Luid tp;
            IntPtr hproc = Win32Token.GetCurrentProcess();
            IntPtr htok = IntPtr.Zero;
            retVal = Win32Token.OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok);
            tp.Count = 1;
            tp.Luid = 0;
            tp.Attr = SE_PRIVILEGE_ENABLED;
            retVal = Win32Token.LookupPrivilegeValue(null, privilege.ToString(), ref tp.Luid);
            retVal = Win32Token.AdjustTokenPrivileges(htok, false, ref tp, Marshal.SizeOf(tp), IntPtr.Zero, IntPtr.Zero);
            Win32Token.CloseHandle(htok);
            return Marshal.GetLastWin32Error();
        }

        public static int RemovePrivilege(Rights privilege)
        {
            bool retVal;
            TokPriv1Luid tp;
            IntPtr hproc = Win32Token.GetCurrentProcess();
            IntPtr htok = IntPtr.Zero;
            retVal = Win32Token.OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok);
            tp.Count = 1;
            tp.Luid = 0;
            tp.Attr = SE_PRIVILEGE_DISABLED;
            retVal = Win32Token.LookupPrivilegeValue(null, privilege.ToString(), ref tp.Luid);
            retVal = Win32Token.AdjustTokenPrivileges(htok, false, ref tp, Marshal.SizeOf(tp), IntPtr.Zero, IntPtr.Zero);
            Win32Token.CloseHandle(htok);
            return Marshal.GetLastWin32Error();
        }
    }
'@

Function Get-SystemPolicy( [Guid]$subCategoryGuid)
{
    $buffer = [IntPtr]::Zero
    if ([Win32.Advapi32]::AuditQuerySystemPolicy( $subCategoryGuid , 1 , [ref]$buffer) -and $buffer -ne [IntPtr]::Zero )
    {
        [System.Runtime.InteropServices.Marshal]::PtrToStructure( [System.IntPtr]$buffer , [type][Win32.Advapi32+AUDIT_POLICY_INFORMATION] ) ## return
        [Win32.Advapi32]::AuditFree($buffer)
        $buffer = [IntPtr]::Zero
    }
}
        
Function Set-SystemPolicy( [Guid]$subCategoryGuid , [Guid]$categoryGuid  )
{
    [bool]$result = $false
    $policy = New-Object -TypeName 'Win32.Advapi32+AUDIT_POLICY_INFORMATION'
    [IntPtr]$buffer = [System.Runtime.InteropServices.Marshal]::AllocHGlobal( [System.Runtime.InteropServices.Marshal]::SizeOf( [type]$policy.GetType() ) )
    if( $buffer -ne [IntPtr]::Zero )
    {
        $policy.AuditSubCategoryGuid = $subCategoryGuid
        $policy.AuditCategoryGuid = $categoryGuid
        $policy.AuditingInformation = [Win32.Advapi32+AUDIT_POLICY_INFORMATION_TYPE]::Success
        [System.Runtime.InteropServices.Marshal]::StructureToPtr( $policy , $buffer , $false )
        [uint64]$number = 1
        $result = [Win32.Advapi32]::AuditSetSystemPolicy( $buffer , $number ); $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
        if( ! $result )
        {
            $script:warnings.Add( "AuditSetSystemPolicy failed - $LastError" )
        }
        [System.Runtime.InteropServices.Marshal]::FreeHGlobal( $buffer )
        $buffer = [IntPtr]::Zero
    }
    else
    {
        $script:warnings.Add( "Failed to allocate memory for audit buffer" )
    }
    $result ## return
}

Function Test-AuditSetting( [string]$GUID , [string]$name , [ref]$setting )
{
    $auditEvent = Get-SystemPolicy -subCategoryGuid $GUID
    if( $auditEvent )
    {
        $setting.Value = $auditEvent.AuditingInformation.ToString()
        ( $auditEvent.AuditingInformation -band [Win32.Advapi32+AUDIT_POLICY_INFORMATION_TYPE]::Success ) -eq [Win32.Advapi32+AUDIT_POLICY_INFORMATION_TYPE]::Success 
    }
    else
    {
        $script:warnings.Add( "Could not get setting for `"$name`" with GUID $GUID" )
    }
}

Function Test-AuditSettings
{
    [CmdletBinding()]
              
    [hashtable]$requiredAuditEvents = @{
        'Process Creation'    = '0cce922b-69ae-11d9-bed3-505054503030'
        'Process Termination' = '0cce922c-69ae-11d9-bed3-505054503030'
        ##'Logon'               = '0cce9215-69ae-11d9-bed3-505054503030'
    }

    [string]$resultString = $null

    if( ! ( ([System.Management.Automation.PSTypeName]'Win32.Advapi32').Type ) )
    {
        [void](Add-Type -MemberDefinition $AuditDefinitions -Name 'Advapi32' -Namespace 'Win32' -UsingNamespace System.Text -Debug:$false)
    }
    [string]$newline = $null
    [string]$setting = $null
    ForEach( $requiredAuditEvent in ($requiredAuditEvents.GetEnumerator() ))
    {
        $result = Test-AuditSetting -GUID $requiredAuditEvent.Value -name $requiredAuditEvent.Name -setting ([ref]$setting)
        if( $result -eq $null -or $result -eq $false )
        {
            $resultString += "$($newline)Auditing of `"$($requiredAuditEvent.Name)`" is not set to at least `"Success`" as required, it is set to `"$setting`""
            $newline = "`n"
        }
    }
    $resultString
}


Function Get-JSONProperty
{
    [CmdletBinding()]

    Param
    (
        [Parameter(ValueFromPipeline,Mandatory=$true,HelpMessage='JSON object to search')]
        $inputObject ,
        [Parameter(Mandatory=$true,HelpMessage='JSON property name to search for')]
        [string]$name ,
        [switch]$multiple ,
        [switch]$regex
    )

    $foundIt = $null

    If( $inputObject -and ! [string]::IsNullOrEmpty( $name ) )
    {
        ForEach( $property in $inputObject.PSObject.Properties )
        {
            If( $property.MemberType.ToString() -eq 'NoteProperty' )
            {
                If( ( ! $regex -and $property.Name -eq $name ) -or ( $regex -and $property.Name -match $name ) )
                {
                    Return $property
                }
                Elseif( $property.Value -is [PSCustomObject] )
                {
                    If( ( $multiple -or ! $foundIt ) -and ( $result = Get-JSONProperty -name $name -inputObject $property.value -multiple:$multiple -regex:$regex ))
                    {
                        $foundIt = $result
                        $result
                    }
                }
            }
        }
    }
}

function Get-LogonDurationAnalysis {
    [CmdletBinding(DefaultParameterSetName="None")]
    param (
        [Parameter(Position=0,
                   Mandatory=$false)]
        [Alias('User')]
        [String]
        $Username = $env:USERNAME,
        
        [Parameter(Position=1,
                   Mandatory=$false)]
        [Alias('Domain')]
        [String]
        $UserDomain = $env:USERDOMAIN,
        
        [Parameter(Mandatory=$false)]
        [Alias('HDX')]
        [int]
        $HDXSessionId,
        
        [Parameter(Mandatory=$false)]
        [String]
        $XDUsername,
        
        [Parameter(Mandatory=$false)]
        [System.Security.SecureString]
        $XDPassword,
        
        [Parameter(Mandatory=$false)]
        [decimal]
        $CUDesktopLoadTime,

        [Parameter(Mandatory=$false)]
        [String]
        $ClientName

    )
    begin {
        $Script:Output = New-Object -TypeName System.Collections.Generic.List[psobject]
        $Script:AppVolumesOutput = New-Object -TypeName System.Collections.Generic.List[psobject]
        $Script:LogonStartDate = $null
        $Script:UseFSLogixWinLogonEvents = $false

        ## array indexes for event log property fields to make retrieval more meaningful
        ## Event id 4688 (process start)
  
        Set-Variable -Name SubjectUserName   -Value 1  -Option ReadOnly
        Set-Variable -Name SubjectDomainName -Value 2  -Option ReadOnly
        Set-Variable -Name SubjectLogonId    -Value 3  -Option ReadOnly
        Set-Variable -Name ProcessIdNew      -Value 4  -Option ReadOnly
        Set-Variable -Name NewProcessName    -Value 5  -Option ReadOnly
        Set-Variable -Name ProcessIdStart    -Value 7  -Option ReadOnly
        Set-Variable -Name NewProcessCmdLine -Value 8  -Option ReadOnly
        Set-Variable -Name TargetUserName    -Value 10 -Option ReadOnly
        Set-Variable -Name TargetDomainName  -Value 11 -Option ReadOnly
        Set-Variable -Name TargetLogonId     -Value 12 -Option ReadOnly
        Set-Variable -Name ParentProcessName -Value 13 -Option ReadOnly
        
        [string]$auditingWarning = $null
        if( ! $offline )
        {
            Test-AuditSettings
        }
        [bool]$SearchCommandLine = $false
        if ([version](Get-CimInstance Win32_OperatingSystem).version -gt ([version]6.1)) { # are we using a version of Windows newer than Windows 2008R2/Windows 7 as not implemented prior to that?
            if (Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Audit' -ErrorAction SilentlyContinue) {
                $commandLinePolicy = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Audit' -Name 'ProcessCreationIncludeCmdLine_Enabled' -ErrorAction SilentlyContinue
                if ($commandLinePolicy -and $commandLinePolicy.ProcessCreationIncludeCmdLine_Enabled -eq 1) {
                    if (-not($auditingWarning -like "*Process Termination*")) { #need process termination auditing enabled or else we can't find when the process finishes
                        Set-Variable -Name CommandLine -Value 8 -Option ReadOnly
                        $SearchCommandLine = $true
                    }
                }
            }
        }
        Write-Debug "Process command line auditing enabled is $SearchCommandLine"
        
        ## Event id 4689 (process stop)
        Set-Variable -Name ProcessStopSid -Value 0 -Option ReadOnly
        Set-Variable -Name ProcessIdStop  -Value 5 -Option ReadOnly
        Set-Variable -Name ProcessName    -Value 6 -Option ReadOnly

        # Generate a new XPath string
        function New-XPath {
            [CmdletBinding(DefaultParameterSetName="None")]  
            param(
                [ValidateNotNullOrEmpty()]
                [array]
                $EventId,
        
                [Parameter(ParameterSetName='DateTime',Mandatory=$true)]
                [DateTime]
                $FromDate,
        
                [Parameter(ParameterSetName='DateTime')]
                [DateTime]
                $ToDate,
        
                [hashtable]
                $SecurityData,
                [Alias('Data')]
        
                $EventData,
        
                [hashtable]
                $UserData ,

                [switch]
                $encode
            )
            [string]$lessThan = if( $encode ) { '&lt;' } else { '<' }
            [string]$greaterThan = if( $encode ) { '&gt;' } else { '>' }
            [System.Text.StringBuilder]$sb = "*[System[("
            $ecounter = 0
            foreach ($eid in $EventId) {
                if ($ecounter -gt 0) {
                    [void]$sb.Append(" or EventID='$eid'")
                }
                else {
                    [void]$sb.Append("EventID='$eid'")
                }
                $ecounter++
            }
            if ($ToDate) {
                [void]$sb.Append(") and TimeCreated[@SystemTime$($greaterThan)='$($FromDate.ToUniversalTime().ToString("s")).$($FromDate.ToUniversalTime().ToString("fff"))Z'")
                [void]$sb.Append(" and @SystemTime$($lessThan)='$($ToDate.ToUniversalTime().ToString("s")).$($FromDate.ToUniversalTime().ToString("fff"))Z']")
                if (!$SecurityData) {
                    [void]$sb.Append("]]")
                }
            }
            elseif ($FromDate) {
                [void]$sb.Append(") and TimeCreated[@SystemTime$($greaterThan)='$($FromDate.ToUniversalTime().ToString("s")).$($FromDate.ToUniversalTime().ToString("fff"))Z']")
                if (!$SecurityData) {
                    [void]$sb.Append("]]")
                }
            }
            else {
                [void]$sb.Append(")]]")
            }
            if ($SecurityData) {
                    [void]$sb.Append(" and Security[@$($SecurityData.Keys[0])='$($SecurityData.Values[0])']]]")
            }
            if ($EventData -and $EventData.GetType() -eq [hashtable]) {
                foreach ($i in $EventData.Keys) {
                    $counter = 0
                    [void]$sb.Append(" and *[EventData[Data[@Name='$i']")
                    foreach ($x in $($EventData.$i)) {
                        if ($counter -gt 0) {
                            [void]$sb.Append(" or Data=`"$($x)`"")
                        }
                        else {
                            [void]$sb.Append(" and (Data=`"$($x)`"")
                        }
                        $counter++
                    }
                    [void]$sb.Append(")]]")
                }
            }
            elseif ($EventData) {
                [void]$sb.Append(" and *[EventData[Data and (Data='$EventData')]]")
            }
            if ($UserData) {
                [void]$sb.Append(" and *[UserData[EventXML[($($UserData.Keys[0])=`'$($UserData.Values[0])`')]]]")
            }
            $sb.ToString()
        }
        
        # Get an event from the Windows Eventlog using specified parameters
        function Get-PhaseEventFromCache {
            [CmdletBinding(DefaultParameterSetName="None")]
            param (
                [ValidateNotNullOrEmpty()]
                $startEvent ,

                $endEvent ,

                [String]
                $PhaseName ,

                [decimal]
                $CUAddition ,

                [string]$source = 'Windows'
            )
            
            if( ! $startEvent )
            {
                Write-Error "Get-PhaseEventFromCache - no start event"
            }
            if( ! $endEvent )
            {
                if($CUAddition -gt 0 -and $startEvent ) {
                    [DateTime]$EndEvent = $StartEvent.TimeCreated.AddMilliseconds($CUAddition*1000)
                }
                else {
                    Write-Error "Get-PhaseEventFromCache - no end event"
                }
            }
            $EventInfo = @{}
            if ($EndEvent) {
                if ((($EndEvent).GetType()).Name -eq 'DateTime') {
                    $Duration = New-TimeSpan -Start $StartEvent.TimeCreated -End $EndEvent
                    $EventInfo.EndTime = $EndEvent
                }
                else {
                    $Duration = New-TimeSpan -Start $StartEvent.TimeCreated -End $EndEvent.TimeCreated
                    $EventInfo.EndTime = $EndEvent.TimeCreated 
                }
            }
            $EventInfo.Source = $source
            $EventInfo.PhaseName = $PhaseName
            $EventInfo.StartTime = $StartEvent.TimeCreated
            $EventInfo.Duration = $Duration.TotalSeconds
            $PSObject = New-Object -TypeName PSObject -Property $EventInfo
            if ($EventInfo.Duration -and $PhaseName -eq 'GP Scripts' -and ($StartEvent.Properties[3]).Value) {
                $PSObject
            }
            elseif ($EventInfo.Duration -and $PhaseName -eq 'GP Scripts') {
                $sharedVars.Add( 'GPASync' , [math]::Round( $PSObject.Duration , 1 ) )
            }
            elseif ($EventInfo.Duration) {
                $PSObject
            }
        }
        
        function Get-EventLogEnabledStatus {
            [CmdletBinding(DefaultParameterSetName="None")]
            param (
                [string]$eventLog
            )

            [string]$status = $null
            if( ! [string]::IsNullOrEmpty( $eventLog ) )
            {
                $eventlogProperties = wevtutil.exe get-log $eventLog
                if( ! $? -or ! $eventlogProperties )
                {
                    $status = "Unable to find event log `"$eventLog`""
                }
                elseif( $eventlogProperties | Where-Object { $_ -match '^enabled: (.*$)' -and $Matches.Count -ge 2 -and $Matches[1] } )
                {
                    if( $Matches[1] -ne 'true' )
                    {
                        $status = "Event log `"$eventLog`" is not enabled so it cannot accept events"
                    }
                }
                else
                {
                    $status = "Unable to determine if event log `"$eventLog`" is enabled"
                }                        
            }
            $status
        }
            
        # Get an event from the Windows Eventlog using specified parameters
        function Get-PhaseEvent {
            [CmdletBinding(DefaultParameterSetName="None")]
            param (
                [AllowNull()]
                [String]
                $StartEventFile ,
                
                [AllowNull()]
                [String]
                $EndEventFile ,

                [ValidateNotNullOrEmpty()]
                [String]
                $PhaseName,
            
                [ValidateNotNullOrEmpty()]
                [String]
                $StartProvider,
            
                [ValidateNotNullOrEmpty()]
                [String]
                $EndProvider,
            
                [ValidateNotNullOrEmpty()]
                [String]
                $StartXPath,
            
                [ValidateNotNullOrEmpty()]
                [String]
                $EndXPath,
            
                [string]
                $eventLog ,

                [System.Diagnostics.Eventing.Reader.EventLogRecord]
                $StartEvent,
            
                [System.Diagnostics.Eventing.Reader.EventLogRecord]
                $EndEvent,
            
                [int]
                $CUAddition ,

                [string]
                $source = 'Windows' ,

                [hashtable]$sharedVars
            )
            [datetime]$started = Get-Date

            [hashtable]$startParams = if( $PSBoundParameters[ 'StartEventFile' ] ) { @{ 'Path' = $StartEventFile } } else { @{ 'ProviderName' = $StartProvider } }
            [hashtable]$endParams = if( $PSBoundParameters[ 'EndEventFile' ] ) { @{ 'Path' = $EndEventFile } } else { @{ 'ProviderName' = $EndProvider } }

            try {
                $PSCmdlet.WriteVerbose("Looking $PhaseName Events")
                if(!$StartEvent) {
                    $StartEvent = Get-WinEvent -Oldest -MaxEvents 1 @startParams -FilterXPath $StartXPath -ErrorAction Stop -Verbose:$False
                }
                if (!$EndEvent) {
                    if ($StartProvider -eq 'Microsoft-Windows-Security-Auditing' -and $EndProvider -eq 'Microsoft-Windows-Security-Auditing') {
                        $EndEvent = Get-WinEvent -MaxEvents 1 @endParams -FilterXPath ("{0}{1}" -f $EndXPath,(
                            "and *[EventData[Data[@Name='ProcessId']" +
                            "and (Data=`'$($StartEvent.Properties[4].Value)`')]]")
                            ) -ErrorAction Stop # Responsible to match the process termination event to the exact process
                    }
                    elseif ($CUAddition) {
                        [DateTime]$EndEvent = $StartEvent.TimeCreated.AddSeconds($CUAddition)
                    }
                    else {
                        $EndEvent = Get-WinEvent -Oldest -MaxEvents 1 @endParams -FilterXPath $EndXPath 
                    }
                }
            }
            catch {
                [string]$eventLogStatus = Get-EventLogEnabledStatus -eventLog $eventLog
                if( ! [string]::IsNullOrEmpty( $eventLogStatus ) )
                {
                    $warnings.Add( $eventLogStatus )
                }
                if ($PhaseName -ne 'Citrix Profile Mgmt' -and $PhaseName -ne 'GP Scripts') {
                    if ($StartProvider -eq 'Microsoft-Windows-Security-Auditing' -or $EndProvider -eq 'Microsoft-Windows-Security-Auditing' ) {
                        $warnings.Add("Could not find $PhaseName events (requires audit process tracking)")
                    }
                    else {
                        $warnings.Add( "Could not find $PhaseName events")
                    }
                }
            }
            finally {
                $EventInfo = @{}
                if ($EndEvent) {
                    if ((($EndEvent).GetType()).Name -eq 'DateTime') {
                        $Duration = New-TimeSpan -Start $StartEvent.TimeCreated -End $EndEvent
                        $EventInfo.EndTime = $EndEvent
                    }
                    else {
                        $Duration = New-TimeSpan -Start $StartEvent.TimeCreated -End $EndEvent.TimeCreated
                        $EventInfo.EndTime = $EndEvent.TimeCreated 
                    }
                }
                $EventInfo.Source = $source
                $EventInfo.PhaseName = $PhaseName
                $EventInfo.StartTime = $StartEvent.TimeCreated
                $EventInfo.Duration = $Duration.TotalSeconds
                $PSObject = New-Object -TypeName PSObject -Property $EventInfo
                if ($EventInfo.Duration -and $PhaseName -eq 'GP Scripts' -and ($StartEvent.Properties[3]).Value) {
                    $PSObject
                }
                elseif ($EventInfo.Duration -and $PhaseName -eq 'GP Scripts') {
                    $sharedVars.Add( 'GPASync' , [math]::Round( $PSObject.Duration , 1 ) )
                    ##$Script:GPAsync = "{0:N1}" -f $PSObject.Duration
                }
                elseif ($EventInfo.Duration) {
                    $PSObject
                }
            }
        }

        function Get-CitrixData {
            [OutputType([System.Collections.Generic.List[psobject]])]
            [CmdletBinding()]
            Param (
                [int]$sessionId
            )
            if( ! ( $clientStartup = Get-CimInstance -Namespace root\Citrix\EUEM -ClassName Citrix_Euem_ClientStartup | Where-Object SessionId -eq $sessionId ) )
            {
                $warnings.Add( "Failed to get Citrix information via CIM for session $sessionId" )
            }
            elseif( $clientStartup.WfIcaTimestamp.Year -lt 2020 )
            {
                $warnings.Add( "Bad date $(Get-Date -Date $clientStartup.WfIcaTimestamp -Format G) returned from root\Citrix\EUEM\Citrix_Euem_ClientStartup" )
            }
            ## check if this data is for a reconnection
            elseif( $clientStartup.WfIcaTimestamp -gt $logon.LogonTime )
            {
                [hashtable]$onlineOfflineTS = @{}
                if( $global:terminalServicesParams[ 'Path' ] )
                {
                    $onlineOfflineTS.Add( 'Path' , $global:terminalServicesParams[ 'Path' ] )
                }
                ## Look for disconnect and reconnect events for this session and user between these two times
                [array]$connectionEvents = @( Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable ( @{ StartTime = $logon.LogonTime ; EndTime = $clientStartup.WfIcaTimestamp.AddSeconds( 120 ) ; Id = @( 24 , 25) ; ProviderName = 'Microsoft-Windows-TerminalServices-LocalSessionManager' } + $onlineOfflineTS ) | Where-Object { $_.Properties[0].Value -eq "$($Logon.Userdomain)\$($logon.username)" -and $_.Properties[1].Value -eq $sessionId } )
                if( $connectionEvents -and $connectionEvents.Count )
                {
                    [string]$warningMessage = "Session "
                    if( $disconnectedEvent = $connectionEvents | Where-Object { $_.Id -eq 24 }  | Select-Object -First 1 )
                    {
                        $warningMessage += "disconnected at $(Get-Date -Date $disconnectedEvent.TimeCreated -Format G) "
                    }
                    if( $reconnectedEvent = $connectionEvents | Where-Object { $_.Id -eq 25 }  | Select-Object -First 1 )
                    {
                        if( $disconnectedEvent )
                        {
                            $warningMessage += 'and '
                        }
                        $warningMessage += "reconnected at $(Get-Date -Date $reconnectedEvent.TimeCreated -Format G) "
                    }
                    $warningMessage += 'so ignoring Citrix WMI event data which is for the reconnection'
                    $script:warnings.Add( $warningMessage )
                }
                else
                {
                    $script:warnings.Add( "Citrix WMI ICA event is $([math]::Round( ($clientStartup.WfIcaTimestamp - $logon.LogonTime).TotalMinutes , 1 ) ) minutes after logon but unable to find evidence of disconnect & reconnect in event log" )
                }
            }
            else
            {
                <#
                https://support.citrix.com/article/CTX114495

                SCCD - STARTUP_CLIENT
                This is the high-level client connection startup metric. It starts as close as possible to the time of the request (mouse click) and ends when the ICA connection between the client device and server running Presentation Server has been established.
                In the case of a shared session, this duration will normally be much smaller, as many of the setup costs associated with the creation of a new connection to the server are not incurred.

                SCD - SESSION_CREATION_CLIENT
                New session creation time, from the moment wfica32.exe is launched to when the connection is established.
                #>

                [datetime]$clicktime = $clientstartup.WfIcaTimestamp.AddMilliseconds($ClientStartup.SCCD).AddMilliseconds(-$ClientStartup.SCD)
                $returning = New-Object -TypeName System.Collections.Generic.List[psobject]
                $returning.Add( 
                    [pscustomobject]@{
                            Source = 'Citrix'
                            PhaseName = 'App/Desktop Icon Clicked until ICA File Downloaded'
                            StartTime = $clickTime
                            EndTime   = $clientStartup.WfIcaTimestamp
                            Duration  = ($ClientStartup.SCD - $ClientStartup.SCCD) / 1000 } )
                $returning.Add( 
                    [pscustomobject]@{
                            Source = 'Citrix'
                            PhaseName = 'ICA File Opened until Remote Session Commences'
                            StartTime = $clientStartup.WfIcaTimestamp
                            EndTime   = $clientstartup.WfIcaTimestamp.AddMilliseconds($ClientStartup.SCCD)
                            Duration  = $ClientStartup.SCCD / 1000 } )
                $returning
            }
        }

        # Connects to the Citrix Broker Monitor Service to get information about a session
        function Get-ODataPhase {
            [CmdletBinding()]
            param (
                [string]
                $SessionKeyPath = 'HKLM:\SOFTWARE\Citrix\Ica\Session\CtxSessions',
                
                [string]
                $DDCPath = 'HKLM:\SOFTWARE\Citrix\VirtualDesktopAgent\State'
            )

            try {
                if ($PSBoundParameters[ 'Verbose' ]) {
                    $PSCmdlet.WriteVerbose("Querying registry for `"SessionKey`" in {0}" -f $SessionKeyPath)
                }
                $CtxSessionsKey = Get-ItemProperty $SessionKeyPath
                if ($PSBoundParameters[ 'Verbose' ]) {
                    $PSCmdlet.WriteVerbose("Querying registry for `"DDC`" in {0}" -f $DDCPath)
                }
                $DDC = Get-ItemProperty $DDCPath | Select-Object -ExpandProperty 'RegisteredDdcFqdn'
	            }
	        catch {
		        $warnings.Add( "Could not access registry: {0}" -f ($Error[0].Exception))
	        }
	        finally {
		        $SessionsIdList = ($CtxSessionsKey | Get-Member -MemberType NoteProperty).Name | Where-Object {$_ -notmatch "PS*"}
	        }
	        if ((($SessionsIdList.GetType()).BaseType).Name -eq "Array") {
		        foreach ($i in $SessionsIdList) {
			        if ($CtxSessionsKey.$i -eq $HDXSessionId) {
				        $SessionKey = $i.Replace('({|})','')
			        }
		        }
	        }
	        else {
		        $SessionKey = $SessionsIdList.Replace('({|})','')
	        }

            $HDXStartTime = $null
            $HDXEndTime = $null

            try {
                Write-Debug "Checking session $sessionKey on DDC $DDC as user $XDUsername"
                $XDCreds = New-Object System.Management.Automation.PSCredential ($XDUsername, $XDPassword)
	            $ODataData = (Invoke-RestMethod -Uri "http://$DDC/Citrix/Monitor/OData/v1/Data/Sessions(guid'$SessionKey')/CurrentConnection" `
                   -Credential $XDCreds ).entry.content.properties
	            try {
		            [DateTime]$HDXStartTime = $ODataData.HdxStartDate.'#text'
		            [DateTime]$HDXEndTime = $ODataData.HdxEndDate.'#text'
	            }
                catch [System.Management.Automation.PropertyNotFoundException] {
                    $warnings.Add( "HDX duration records were null.")
                }
	            catch {
		            $warnings.Add( "No records for this session found on DDC $DDC.")
	            }
                finally {
                    if (($HDXStartTime) -and ($HDXEndTime)) {
		                $HDXSessionDuration = (New-TimeSpan -Start $HDXStartTime -End $HDXEndTime).TotalSeconds
                        [pscustomobject]@{
                            PhaseName = 'HDX Connection'
                            StartTime = $HDXStartTime.ToLocalTime()
                            EndTime = $HDXEndTime.ToLocalTime()
                            Duration = $HDXSessionDuration
                        }
                    }
                }
            }
            catch {
	            $warnings.Add( (("Could not initiate a connection to {0},`n {1}`nMake sure the user {2} has at least the `"Read-Only Administrator`" role") -f $DDC, $Error[0].Exception.Message , $XDUsername ))
	        }
        }

        function Get-UserLogonDetails {
            [CmdletBinding()]

            Param(
                [Parameter(Mandatory=$true)]
                [string]
                $UserName
            )
                [string[]]$sess = (quser.exe "$username" | Select -Skip 1 | Select -Last 1) -split '\s+'
                [string]$info = $null

                if( $sess -and $sess.Count )
                {
                    if( $sess[-1] -match '^[AP]M$' )
                    {
                        $info = " - logon was $($sess[-3..-1] -join ' ')"
                    }
                    else
                    {
                        $info = " - logon was $($sess[-2..-1] -join ' ')"
                    }
                }
                else
                {
                    $info = " - user $username not currently logged on"
                }
                $info
        }

        function Get-LogonTask {
            [CmdletBinding()]
            param(
            [Parameter(Mandatory=$true)]
            [string]
            $UserName,
            
            [Parameter(Mandatory=$true)]
            [string]
            $UserDomain,
            
            [Parameter(Mandatory=$true)]
            [DateTime]
            $Start,
            
            [Parameter(Mandatory=$true)]
            [DateTime]
            $End
            )

            [hashtable]$logonTaskParams = $global:scheduledTasksParams.Clone()
            $logonTaskParams.Add( 'StartTime' , $start )
            $logonTaskParams.Add( 'Id' , @(119,201) )
            [array]$logontaskEvents = @( Get-WinEvent -FilterHashtable $logonTaskParams -ErrorAction SilentlyContinue)

            $logontaskEvents | Where-Object { $_.Id -eq 119 -and $_.TimeCreated -le $end -and $_.Properties[1].Value -eq "$UserDomain\$UserName" } | ForEach-Object `
            {
                $taskStart = $_
                $taskEnd = $logontaskEvents | Where-Object { $_.Id -eq 201 -and $taskStart.Properties[2].Value -eq $_.Properties[1].Value }  ## Correlate task instance id
                if( $taskEnd )
                {
                    New-Object -TypeName psobject -Property @{ 
                            'TaskName'="$($TaskEnd.Properties[0].Value)"
                            'ActionName'="$($TaskEnd.Properties[2].Value)"
                            'Duration'=$taskEnd.TimeCreated - $taskStart.TimeCreated
                        }
                } 
            }
        }

        function Get-PrinterEvents {
            [CmdletBinding()]
            param(
            [Parameter(Mandatory=$true)]
            [DateTime]
            $Start,
            
            [Parameter(Mandatory=$true)]
            [DateTime]
            $End,

            [Parameter(Mandatory=$false)]
            [String]
            $ClientName
            )

            Write-Verbose "Get-PrinterEvents Start Time: $start"
            Write-Verbose "Get-PrinterEvents End Time: $end"
            Write-Verbose "Get-PrinterEvents ClientName: $ClientName"

            if( ! $offline )
            {
                [string]$eventLogStatus = Get-EventLogEnabledStatus -eventLog 'Microsoft-Windows-PrintService/Operational'
                if( ! [string]::IsNullOrEmpty( $eventLogStatus ) )
                {
                    $warnings.Add( $eventLogStatus )
                    return
                }
            }

            if( [string]::IsNullOrEmpty( $End ) )
            {
                $warnings.Add( "No logon end event was found.  Please wait and try again once logon has completed.  Printer information will not be displayed." )
                return
            }

            if (-not(Test-Path HKU:\)) {
                New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS | out-null
            }
            
            $UserPrinterGUIDs = [System.Collections.ArrayList]@()
            [array]$PrinterClientSidePortGUIDs = @()

            if (-not(Test-Path HKU:\$($Logon.UserSID)\Printers\Connections\ -ErrorAction SilentlyContinue)) {
                Write-Verbose "Unable to find mapped printers in the user session."  #we'll do our best though with what's available
            } else {
                $UserPrinterGUIDs += Get-ItemProperty -Path HKU:\$($Logon.UserSID)\Printers\Connections\* -Name GuidPrinter -ErrorAction SilentlyContinue
                $PrintServers = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\Servers" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PSChildName
                Write-Verbose "Found the following print servers:"
                Write-Verbose "$printServers"
                $PrinterClientSidePortGUIDs = @( foreach ($printServer in $printServers) {
                   Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\Servers\$printServer\Monitors\Client Side Port\*" -Name PrinterPath -ErrorAction SilentlyContinue
                })
                if ($DebugPreference -eq "continue") {
                    Write-Debug "Printer GUIDS:"
                    foreach ($ClientSidePortGUID in $PrinterClientSidePortGUIDs) {
                        Write-Debug "$($ClientSidePortGUID.printerPath)"
                    }
                }
    
            }
            [hashtable]$printerParams = $global:printServiceParams.Clone() +  @{ StartTime = $start ; EndTime = $end ; Id = 300,306}
            [array]$printerTaskEvents = @( Get-WinEvent -FilterHashtable $printerParams -ErrorAction SilentlyContinue )
            if ($printerTaskEvents.count -eq 0) {
                #no printer events found.  This may occur if the application is not set to wait for printers (totally normal!) so just return without a message
                Write-Verbose "No Printer Events Found."
                return
            }
            #get list of printers:
            $listOfPrinters = [System.Collections.ArrayList]@()
            $AllPrinterEvents = [System.Collections.ArrayList]@()
            foreach ($printerEvent in $printerTaskEvents) {
                if ($printerEvent.Id -eq "300") { #look for event ID 300 -- "Add printer".  Should be unique for each printer
                    #check if this is a GUID
                    if ($printerEvent.Properties.Value -match("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$")) {
                        foreach ($printerGUID in $UserPrinterGUIDs) {
                            Write-Debug "Searching for User Printer GUID: $($printerGUID.GuidPrinter)"
                            if ($printerGUID.GuidPrinter -eq $printerEvent.Properties.Value) {
                                #if printer has a GUID than it's a direct connection printer.  Capture its properties here
                                $printerName = $printerGUID.PSChildName -replace (",","\")
                                $printerGUIDValue = $printerEvent.Properties.Value
                                $printer = New-Object PSObject -property @{Name="$printerName";Value="$printerGUIDValue";Type="Direct Connection"}
                                Write-Verbose "Found Direct Connection Printer: $($printer)"
                                Write-Verbose "GUID: $($printerGUIDValue)"
                                $listOfPrinters += $printer

                                if ($SearchCommandLine) {
                                    Write-Verbose "We can search the command line for the print driver install events"
                                    #pull driver installation time -- requires 2012R2+ and command line capture policy enabled.
                                    $printDriverInstallationStartEvent = ($securityEvents|Where-Object { $_.Id -eq 4688 -and $_.properties[$NewProcessName].Value -eq 'C:\Windows\System32\drvinst.exe' -and $_.properties[$CommandLine].Value -like "*$printerGUIDValue*" }) 
                                    $printDriverInstallationEndEvent = ($securityEvents|Where-Object { $_.Id -eq 4689 -and $_.properties[$ProcessIdStop].Value -eq $printDriverInstallationStartEvent.Properties[$ProcessIdNew].Value -and $_.properties[$processName].Value -eq 'C:\Windows\System32\drvinst.exe'})
                                    Write-Verbose "New-TimeSpan -Start $($printDriverInstallationStartEvent.TimeCreated) -End $($printDriverInstallationEndEvent.TimeCreated)"
                                    $Duration = New-TimeSpan -Start $($printDriverInstallationStartEvent.TimeCreated) -End $($printDriverInstallationEndEvent.TimeCreated)
                                    $EventInfo = @{}
                                    $EventInfo.Source = 'Printers'
                                    $EventInfo.PhaseName = "    Driver : $printerName "
                                    $EventInfo.Duration = $Duration.TotalSeconds
                                    $EventInfo.EndTime = $printDriverInstallationEndEvent.TimeCreated
                                    $EventInfo.StartTime = $printDriverInstallationStartEvent.TimeCreated
                                    $AllPrinterEvents +=  New-Object -TypeName PSObject -Property $EventInfo
                                    Write-Debug "Post event creation"
                                    $PSObject = New-Object -TypeName PSObject -Property $EventInfo
        
                                    if ($EventInfo.Duration) {
                                        Write-Verbose "Adding driver phase to Output"
                                        $Script:Output.Add( $PSObject )
                                    }
                                }
                            }
                        }
                        foreach ($PrinterClientSidePortGUID in $PrinterClientSidePortGUIDs) {
                            Write-Debug "Searching for Printer Client Side Port GUID: $($PrinterClientSidePortGUID.PSChildName)"
                            if ($PrinterClientSidePortGUID.PSChildName -eq $printerEvent.Properties.Value) {
                                #we've found a printer client side port match.  This maybe due to a user reconnecting and the GUID's change on reconnect.
                                #the client side port registry keys contain the path to the real printer key
                                Write-Verbose "Client side port printer path: $($PrinterClientSidePortGUID.PrinterPath)"
                                $printerPath = ($PrinterClientSidePortGUID.PrinterPath -replace "\\Users\\$($Logon.UserSID)\\Printers\\","" -replace "\^","")
                                foreach ($printerGUID in $UserPrinterGUIDs) {
                                    $printerName = $printerGUID.PSChildName -replace (",","\")
                                    Write-Debug "Searching for Printer Name Match: $printerName"
                                    if ($printerName -eq $printerPath) {
                                        Write-Verbose "Found a Match: $printerName"
                                        #check to see if we captured this previously
                                        if (-not($listOfPrinters.Name -contains $printerPath)) {
                                            #if printer has a GUID than it's a direct connection printer.  Capture its properties here
                                            $printerGUIDValue = $printerEvent.Properties.Value
                                            $printer = New-Object PSObject -property @{Name="$printerName";Value="$printerGUIDValue";Type="Direct Connection"}
                                            Write-Verbose "Found Direct Connection Printer: $($printer)"
                                            Write-Verbose "GUID: $($printerGUIDValue)"
                                            $listOfPrinters += $printer

                                            if ($SearchCommandLine) {
                                                Write-Verbose "We can search the command line for the print driver install events"
                                                #pull driver installation time -- requires 2012R2+ and command line capture policy enabled.
                                                $printDriverInstallationStartEvent = ($securityEvents|Where-Object { $_.Id -eq 4688 -and $_.properties[$NewProcessName].Value -eq 'C:\Windows\System32\drvinst.exe' -and $_.properties[$CommandLine].Value -like "*$printerGUIDValue*" }) 
                                                $printDriverInstallationEndEvent = ($securityEvents|Where-Object { $_.Id -eq 4689 -and $_.properties[$ProcessIdStop].Value -eq $printDriverInstallationStartEvent.Properties[$ProcessIdNew].Value -and $_.properties[$processName].Value -eq 'C:\Windows\System32\drvinst.exe'})
                                                Write-Verbose "New-TimeSpan -Start $($printDriverInstallationStartEvent.TimeCreated) -End $($printDriverInstallationEndEvent.TimeCreated)"
                                                $Duration = New-TimeSpan -Start $($printDriverInstallationStartEvent.TimeCreated) -End $($printDriverInstallationEndEvent.TimeCreated)
                                                $EventInfo = @{}
                                                $EventInfo.Source = 'Printers'
                                                $EventInfo.PhaseName = "    Driver : $printerName "
                                                $EventInfo.Duration = $Duration.TotalSeconds
                                                $EventInfo.EndTime = $printDriverInstallationEndEvent.TimeCreated
                                                $EventInfo.StartTime = $printDriverInstallationStartEvent.TimeCreated
                                                $AllPrinterEvents +=  New-Object -TypeName PSObject -Property $EventInfo
                                                Write-Verbose "Post event creation"
                                                $PSObject = New-Object -TypeName PSObject -Property $EventInfo
        
                                                if ($EventInfo.Duration) {
                                                    Write-Verbose "Adding driver phase to Output"
                                                    $Script:Output.Add( $PSObject )
                                                }
                                            }
                                        }
                                    }
                                }
                                ########################################################################

                            }
                    }
                   } else {
                        #printer is a regular mapped printer.  Capture its properties here
                        #check client name in case there were concurrent logons to ensure we're targetting just events from this user
                        if( ! [string]::IsNullOrEmpty( $clientName ) ) {
                            if ($printerEvent.Message -like "*$clientName*") {
                                $printerName = ($printerEvent.Message -split "Printer " -split " on " -split "\(from")[1]
                                $printer = New-Object PSObject -property @{Name="$printerName";Value="N/A";Type="Mapped"}
                                Write-Verbose "Found Mapped Printer           : $($printer)"
                                $listOfPrinters += $printer
                            }
                        }
                    }
                }
            }

            foreach ($printer in $listOfPrinters) {
                $phaseName = "    Printer: $($printer.Name)"
                Write-Verbose "Phase: $phaseName"
                
                #capture each event 300 and 306 for the target printer.  There are further events 312 and 314 (add forms, deleting forms) that
                #occur for direct connection printers that is difficult to capture because the events lack targets, you can only do it via
                #date stamps.  Relying on that would be risky if there were concurrent logons, so we'll rely on the interim delay.
                $Events = New-Object -Typename System.Collections.Generic.List[psobject]

                foreach ($printerEvent in $printerTaskEvents | Where-Object {($_.message -like "*$($printer.Name)*") -or ($_.message -like "*$($printer.Value)*")}) {
                    #$printerEvent
                    $Event = [pscustomobject]@{
                        'TimeCreated' = $printerEvent.TimeCreated
                        'Id' = $printerEvent.Id
                    }
                    $Events.Add( $Event )
                    Write-Verbose "Found $($printer.name)"
                    
                }
                write-Verbose "Events: $($events.count) for $($printer.name)" #this should be more than 1
                if ($events.count -gt 1) {
                    if( $Duration = New-TimeSpan -Start $($Events[-1].TimeCreated) -End $($Events[0].TimeCreated) )
                    {             
                        $eventInfo = [pscustomobject]@{
                            'Source' = 'Printers'
                            'PhaseName' = $PhaseName
                            'Duration' = $Duration.TotalSeconds
                            'EndTime' = $Events[0].TimeCreated
                            'StartTime' = $Events[-1].TimeCreated
                        }
                        $Script:Output.Add( $eventInfo )
                        $AllPrinterEvent.Add( $EventInfo )
                    }
                    Clear-Variable Events
                }
            }

            #capture the totality of the printer mapping sequence.
            if( $AllPrinterEvents -and $AllPrinterEvents.Count )
            {
                if( $Duration = New-TimeSpan -Start $($AllPrinterEvents.StartTime | sort -Descending)[-1] -End $($AllPrinterEvents.EndTime | sort -Descending)[0] )
                {
                    $Script:Output.Add( [pscutomobject]@{
                        'Source' = 'Printers'
                        'PhaseName' = "Connect to Printers"
                        'Duration' = $Duration.TotalSeconds
                        'EndTime' = ($AllPrinterEvents.EndTime | sort -Descending)[0]
                        'StartTime' = (($AllPrinterEvents.StartTime | sort -Descending)[-1]).AddMilliseconds(-10) #we subtract 5 milliseconds so the order sorts correctly
                        } )
                }
            }
            else
            {
                Write-Debug "No printer events found for client $ClientName"
            }
        }
        
        function Get-FSLogixProfileEvents {
            [CmdletBinding()]
            param(
            [Parameter(Mandatory=$true)]
            [DateTime]
            $Start,
            
            [Parameter(Mandatory=$true)]
            [DateTime]
            $End,

            [Parameter(Mandatory=$true)]
            [String]
            $Username
            )

            Write-Verbose "Entered Get-FSLogixProfileEvents function"
            #Default FSLogix Log path is here: C:\ProgramData\FSLogix\Logs\Profile
            #at the time of this testing version 2.9.7205.27375 of FSLogix provided all the necessary information
            $GetFSLogixEvents = $false

            if( $offline )
            {
                if (Test-Path $(Join-Path -Path $global:logsFolder -ChildPath 'FSLogixProfileLog.txt')) {
                    Write-Verbose "Offline FSLogix Logfile found."
                    $profileLog = $(Join-Path -Path $global:logsFolder -ChildPath 'FSLogixProfileLog.txt')
                    $GetFSLogixEvents = $true
                } else {
                    Write-Verbose "Unable to determine or find offline FSLogix profile log file."
                }
            } else {
                [string]$FSLogixLogDir = $null

                try {
                    $FSLogixLogDir = Get-ItemPropertyValue -Path HKLM:\SOFTWARE\FSLogix\Logging -Name Logdir -ErrorAction SilentlyContinue
                }
                Catch {
                    #LogDir registry value not found. Set to default:
                    Write-Verbose "LogDir value not set. Setting LogDir to default path"
                    $FSLogixLogDir = Join-Path -Path ([Environment]::GetFolderPath( [System.Environment+SpecialFolder]::CommonApplicationData )) -ChildPath 'FSLogix\Logs'
                }

                [string]$FSLogixProfileLogDir = $( if( ! [string]::IsNullOrEmpty( $FSLogixLogDir ) ) { Join-Path -Path $FSLogixLogDir -ChildPath 'Profile' } )
                if ( ! [string]::IsNullOrEmpty( $FSLogixLogDir ) -and ( Test-Path -Path $FSLogixProfileLogDir -ErrorAction SilentlyContinue ) ) {
                    Write-Verbose "Found FSLogix Profile Log directory."
                    $profileLog = Get-ChildItem $FSLogixProfileLogDir | Where-Object Name -like "*$($($start).ToString("yyyyMMdd"))*"
                    if ( $profileLog -and ( Test-Path $profileLog.FullName -ErrorAction SilentlyContinue )) {
                        $GetFSLogixEvents = $true
                    } else {
                        Write-Verbose "Unable to determine or find FSLogix profile log file."
                    }
                }
            }

            if ($GetFSLogixEvents) {
                Write-Verbose "Found Profile Log file: $($profileLog.FullName)"

                $FSLogixLogObject = New-Object -TypeName System.Collections.Generic.List[psobject]
                $date = Get-Date -Date $start -Format d

                #Create powershell object out of the FSLogix Log.
                Get-Content -Path "$($profileLog.fullname)" | . { Process {
                    $line = $_
                    ## [14:08:54.654][tid:00000bbc.000007e0][INFO]           ===== Begin Session: Logon 
                    if( $line -match '^\[([^\]]+)\]\[([^\]]+)\]\[([^\]]+)\]\s*(.+)' `
                        -and ( $fslogixTime = "$($matches[1]) $date" -as [datetime] ) `
                        -and $fslogixTime -ge $start -and $fslogixTime -le $end ) {
                            $FSLogixLogObject.Add( [PSCustomObject]@{
                                Time     = $FSLogixTime
                                ThreadId = $Matches[2]
                                LogLevel = $Matches[3]
                                Message  = $Matches[4].Trim()
                            })
                        }
                    }
                }

                ## see if it's running but not configured in which case omit the phase and issue a warning
                if( $FSLogixLogObject.Where( { $_.Message -match 'Profiles feature is not enabled' } , 1 ) )
                {
                    $warnings.Add( 'FSLogix is running but not configured' )
                }
                elseif( $failure = $FSLogixLogObject.Where( { $_.Message -match "LoadProfile failed.*\b$username\b" } , 1 ) )
                {
                    $warnings.Add( "Error $($failure.LogLevel -replace 'ERROR:') loading FSlogix profile" )
                }

                $SessionEvents = $FSLogixLogObject.Where( { $_.Message -like "*LoadProfile: $username*" } )
                Write-Verbose "FSLogix: SessionEvents Count: $($SessionEvents.message.count)"
                    
                if ($SessionEvents.message.count -le 1) {
                    #TTYE - It's been noticed that if TimeZone settings apply during the logon the timestamps in the log file
                    #will be modified to reflect that, this 
                    Write-Verbose "FSLogix: Unable to find start or end event in the log file."
                    Write-Verbose "FSLogix: Will attempt to use WinLogon to track this phase"
                    if ( $global:winlogonParams[ 'Path' ] -or ( Get-WinEvent -ListProvider 'Microsoft-Windows-Winlogon' -ErrorAction SilentlyContinue)) {
                        [scriptblock]$winLogonScriptBlock = $null
                        if( $global:winlogonParams[ 'Path' ] )
                        {
                            $winLogonScriptBlock =
                            {
                                Param( $logon , $username , $WinlogonFile )
                                Get-PhaseEvent -source 'FSLogix' -PhaseName 'LoadProfile*' -StartProvider 'Microsoft-Windows-Winlogon' `
                                    -StartEventFile $WinlogonFile `
                                    -EndEventFile $WinlogonFile `
                                    -EndProvider 'Microsoft-Windows-Winlogon' -StartXPath (
                                    New-XPath -EventId 811 -From (Get-Date -Date $Logon.LogonTime) `
                                        -SecurityData @{
                                            UserID=$Logon.UserSID
                                        } -EventData @{
                                            Event="2"
                                            SubscriberName="frxsvc"
                                            }) -EndXPath (
                                    New-XPath -EventId 812 -From (Get-Date -Date $Logon.LogonTime) `
                                        -SecurityData @{
                                            UserID=$Logon.UserSID
                                        } -EventData @{
                                            Event="2"
                                            SubscriberName="frxsvc"
                                            })
                            }
                        }
                        else ## online
                        {
                            $winLogonScriptBlock =
                            {
                                Param( $logon )
                                Get-PhaseEvent -source 'FSLogix' -PhaseName 'LoadProfile*' -StartProvider 'Microsoft-Windows-Winlogon' `
                                    -EndProvider 'Microsoft-Windows-Winlogon' -StartXPath (
                                    New-XPath -EventId 811 -From (Get-Date -Date $Logon.LogonTime) `
                                        -SecurityData @{
                                            UserID=$Logon.UserSID
                                        } -EventData @{
                                            Event="2"
                                            SubscriberName="frxsvc"
                                            }) -EndXPath (
                                    New-XPath -EventId 812 -From (Get-Date -Date $Logon.LogonTime) `
                                        -SecurityData @{
                                            UserID=$Logon.UserSID
                                        } -EventData @{
                                            Event="2"
                                            SubscriberName="frxsvc"
                                            })
                            }
                        }
                        
                        if( ( $FSLogixWinLogonOutput = Invoke-Command $winLogonScriptBlock -ArgumentList $logon ) `
                            -and ( $Duration = New-TimeSpan -Start $FSLogixWinLogonOutput.StartTime -End $FSLogixWinLogonOutput.EndTime ) )
                        {
                            $Script:Output.Add( [pscustomobject]@{
                                'Source' = 'FSLogix'
                                'PhaseName' = 'LoadProfile*'
                                'Duration' = $Duration.TotalSeconds
                                'EndTime' = $FSLogixWinLogonOutput.EndTime
                                'StartTime' = $FSLogixWinLogonOutput.StartTime
                            } )
                        }
                    }
                } else {
                        Write-Verbose "FSLogix: Using FSLogix log file for calculations"
                        $FSLogixStartEvent = $SessionEvents[0].time
                        $FSLogixEndEvent = $SessionEvents[1].time

                        if( $Duration = New-TimeSpan -Start $SessionEvents[0].time -End $SessionEvents[1].time )
                        {
                            $Script:Output.Add([pscustomobject]@{
                                'Source' = 'FSLogix'
                                'PhaseName' = "LoadProfile"
                                'Duration' = $Duration.TotalSeconds
                                'EndTime' = $SessionEvents[1].time
                                'StartTime' = $SessionEvents[0].time
                            })
                        }
                    }
                }
            }
 

        function Get-AppVolumeEvents {
            [CmdletBinding()]
            param(
            [Parameter(Mandatory=$true)]
            [DateTime]
            $Start,
            
            [Parameter(Mandatory=$true)]
            [DateTime]
            $End
            )
            
            ## event log filters for using online or offline
            [hashtable]$onlineOfflineFilter = @{}

            if( $global:AppVolumesParams[ 'Path' ] )
            {
                $onlineOfflineFilter.Add( 'Path' , $global:AppVolumesParams[ 'Path' ] )
            }
            else
            {
                $onlineOfflineFilter.Add( 'LogName' , 'Application' )
            }

            if( ! $global:appVolumesVersion -and ! $offline -and ( $appvolumesKey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -Name DisplayName -ErrorAction SilentlyContinue | Where-Object DisplayName -match 'App Volumes Agent' | Select-Object -ExpandProperty PSPath ) ) {
                $global:appVolumesVersion = Get-ItemProperty -Path $appvolumesKey -Name DisplayVersion | Sort-Object -Descending -Property DisplayVersion | Select-object -ExpandProperty DisplayVersion -First 1
            }
            
            if ($global:appVolumesVersion -lt [version]'2.12') {
                $script:warnings.Add( "Untested AppVolumes version detected: $global:appVolumesVersion" )
            }

            Write-Debug "AppVolumes: AppVolumes version detected: $global:appVolumesVersion"
            
            <#
            https://docs.vmware.com/en/VMware-App-Volumes/2.18/com.vmware.appvolumes.admin.doc/GUID-8CB3E73C-2392-40A2-A19A-825D8D487D08.html
            WaitForFirstVolumeOnly	
            REG_DWORD	
            Defined in seconds, only hold logon for the first volume. After the first volume is complete, the remaining are handled in the background, 
            and the logon process is allowed to proceed. To wait for all volumes to load before releasing the logon process, 
            set this value to 0. The default is 1.

            From the description this will block the logon process as it will wait for all volumes to attach. We'll track this and change how we measure
            The AppVolumes - VolumeAttach stage.
            #>
            
            if( ! $offline )
            {
                try {
                    if( ( Get-ItemPropertyValue -Path HKLM:\SYSTEM\CurrentControlSet\Services\svservice\Parameters -Name WaitForFirstVolumeOnly ) -eq  0 ) {
                       $global:WaitForFirstVolumeOnly = $false
                    }
                } catch {
                    Write-Debug 'AppVolumes: WaitForFirstVolumeOnly value not found. Using Default'
                }
            }
            ## else we will have set it by parsing the svservice.log file name which contains the version number as well

            Write-Debug "AppVolumes: WaitForFirstVolumeOnly set to $global:WaitForFirstVolumeOnly"
            
            <#
            Determines if logon is ASYNC or synchronously. We'll do that by looking at event 218 and see if the last line in the message is 'ASYNC'. If
            it's not then we'll assume we're running synchonsouly.
            #>
            
            [bool]$Async = $false
            if( ( Get-WinEvent -FilterHashtable ( @{ ProviderName='svservice'; StartTime=$Start; Id=218 } + $onlineOfflineFilter ) -ErrorAction SilentlyContinue | Where-Object { $_.properties -and $_.properties[1].value -cmatch 'ASYNC$' } ) ) {
                $Async=$true
            }
            Write-Debug "AppVolumes: AppVolumes Mount Mode Async? : $Async"

            if (Test-Path -Path $appVolumesLogFile ) {
            
                ## Step 1, Parse the log file to a sortable, searchable object.
                [int]$DEBUGNumberOfLines = 0

                $StreamStartTime = Get-Date
                $timeZone = Get-TimeZone
                
                $svserviceLogObject = @( Get-Content -Path $appVolumesLogFile | . { Process {
                    ## [2020-01-10 13:43:23.383 UTC] [svservice:P7776:T1108] Service path: C:\Program Files (x86)\CloudVolumes\Agent\svservice.exe
                    ## Split into 3 - the first two [ ] delimited sections and then the rest
                    if( $_ -match '^\[([^\]]+)\] \[([^\]]+)\] (.+)$' -and ( $time = [datetime]::ParseExact( $Matches[1] , 'yyyy-MM-dd HH:mm:ss.fff UTC' , $null ) ) `
                        -and ( $adjustedForTimeZone = [System.TimeZoneInfo]::ConvertTimeFromUtc( $time, $timeZone ) ) `
                            -and $adjustedForTimeZone -ge $start )
                    {
                        $DEBUGNumberOfLines++
                        [pscustomobject]@{
                            'Time' = $adjustedForTimeZone
                            'ProcessInfo' = $Matches[2]
                            'Message' = $Matches[3]
                        }
                    }
                    
                }})

                $StreamEndTime = Get-Date
                Write-Verbose "AppVolumes log number of lines parsed : $DEBUGNumberOfLines"
                Write-Verbose "AppVolumes log parsing took : $($(New-TimeSpan -Start $StreamStartTime -End $StreamEndTime).TotalSeconds) seconds"

                ## Step 2, Create an object with the following relationship --> AppName, DiskGUID, AppGUID
                ## Sort log by relevant events
                [System.Collections.Generic.List[psobject]]$AppVolumesLogonEvents = $svserviceLogObject | Where {($_.Time -ge $Start) -and ($_.Time -le $End)}
                
                ## Get Mapped in AppVolumes from the Event Logs
                $AppList = New-Object -TypeName System.Collections.Generic.List[psobject]
                Get-WinEvent -FilterHashtable ( @{ 'ProviderName' = 'svservice' ; Id = 218 ; StartTime = $start ; EndTime = $end } + $onlineOfflineFilter ) -ErrorAction SilentlyContinue | . { Process { $_.properties[1].Value -split  "`r`n" } } | . { Process {
                    ## multiple lines of MOUNTED-READ;External SSD\appvolumes\packages\PuTTY.vmdk;{b2a70e3f-90ef-45b5-87a0-b7d07402a977}
                    if( $_ -match '(MOUNTED.*);(.+);({.+})' ) {
                        $AppList.Add( [pscustomobject]@{
                            'MountType' = $matches[1]
                            'AppPath'   = $matches[2]
                            'AppGUID'   = $matches[3] 
                            'AppName'   = $matches[2].Split( '\' )[-1].Replace( '.vmdk' , '' ).Replace( '!20!' , ' ').Replace( '!2B!' , '+' ).Replace( '!5C!' , '\' )  
                            'AppId'     = $null } )
                    }
                    elseif( $_ -match 'ENABLE-APP;({.+});({.+})' ) ## ENABLE-APP;{65950e61-0304-45a6-b354-f3b26ced3f64};{b2a70e3f-90ef-45b5-87a0-b7d07402a977}
                    {
                        if( $appObject = $AppList | Where-Object AppGuid -eq $Matches[2] | Select-Object -First 1 )
                        {
                            $appObject.AppId = $Matches[1]
                        }
                        else
                        {
                            Write-Debug "Couldn't find appguid $($matches[2]) in AppList"
                        }
                    }
                }}

                Write-Verbose "AppVolumes List:"
                Write-Verbose "$($AppList | Select-Object -Property * | Out-String)"

                [array]$AppVolGUIDMappings = $( @( Foreach ($App in $AppList) {
                    $result = $null
                    Write-Verbose "AppGUID   : $($App.AppGUID)"
                    $AppVolGUIDEvents = $AppVolumesLogonEvents.Message | Where-Object {$_ -like "*$($App.AppGUID)*"}
                    Write-Debug "AppVolGUIDEvents : $($AppVolGuidEvents | Out-String)"
                    foreach ($GUIDEvent in $AppVolGUIDEvents) {
                        if( ! $result -and $GUIDEvent -match '(\\Device\\\w+)' ) {
                            [string]$appDevice = $Matches[1]
                            Write-Verbose "AppDevice : $appDevice"
                            $AppVolDeviceEvents = $AppVolumesLogonEvents.Message | Where-Object {$_ -like "*$appDevice*"}
                            ## Do we assume VMWare will always keep the messages to a specific format? Or filter out the AppGUID
                            ## and assume what's left is the device GUID?  I'm leaning towards the latter... Let me know if this fails future Trentent
                            foreach ($AppVolDeviceEvent in $AppVolDeviceEvents) {
                                if (($AppVolDeviceEvent -like "*$appDevice*") -and ($AppVolDeviceEvent -notlike "*$($App.AppGUID)*")) {
                                    Write-Debug "AppVolDeviceEvent: $AppVolDeviceEvent"
                                    if( $appDevice -and $App.AppGUID -and ( $GUIDs = [regex]::Match( $AppVolDeviceEvent , "({[0-9A-Fa-f]{8}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{12}})" ) ) -and $GUIDS.Success )
                                    {
                                        $DiskGUID =  $(if ( ! ( $GUIDS -is [array] ) -or $GUIDS.count -eq 1) {
                                            $GUIDS.Value
                                        } elseif ($GUIDS.count -eq 2) {
                                            $GUIDS.Value | Where-Object { $_ -ne $App.AppGUID }
                                        } else {
                                            Write-Debug "AppVolumes: ERROR -- Unable to determine DiskGUD from `"$AppVolDeviceEvent`""
                                        })
                                        ## only create object if all fields are present
                                        if( $DiskGUID -and ($appName = $AppList | Where-Object AppGUID -eq $App.AppGUID | Select-Object -ExpandProperty AppName ) )
                                        {
                                            Write-Verbose "$appName DiskGUID : $DiskGUID"
                                            $result = [pscustomobject]@{
                                                'AppGUID' = $App.AppGUID
                                                'AppDevice' = $appDevice
                                                'DiskGUID' = $DiskGUID
                                                'AppName' = $appName
                                                'AppId' = $app.AppId }
                                            $result
                                        }
                                    }
                                }
                            }
                        }
                    }
                } ) | Sort-Object -Property AppGUID -Unique ) ## ensure only have 1 entry per app

                ## Now that we have the full relationships for each App we can trace how long they took for each stage
                $appVolArgs = @{}
                $appVolArgs.Add("StartTime", $($Start))
                $appVolArgs.Add("EndTime", $($End))
                $appVolArgs.Add("ProviderName","svservice")
                $AppVolWinEvents = @( Get-WinEvent -FilterHashtable ( $appVolArgs + $onlineOfflineFilter ) -ErrorAction SilentlyContinue )

                Write-Debug "Number of AppVolume events: $($AppVolWinEvents.count)"

                #We are going to aggregate all the results. We're going to check for the 3 properties for each AppVolume (DiskGUID, AppGUID, AppDevice)
                #for each event in the Windows Application Event log, Security Event Log and the svservice.log file. If there is a match we'll add it to a sortable object with those
                #properties.
                $perAppTimes = New-Object -TypeName "System.Collections.Generic.List[psobject]"
                for ($i=0; $i -lt $AppVolGUIDMappings.Count; $i++) {
                    #get all Windows Application events for that specific AppVol Object
                    foreach ($event in $AppVolWinEvents) {
                        if (($event.message -like "*$($AppVolGUIDMappings[$i].DiskGUID)*") -or ($event.message -like "*$($AppVolGUIDMappings[$i].AppDevice)*") -or ($event.message -like "*$($AppVolGUIDMappings[$i].AppGUID)*") -or ($event.Message -like "*$($AppVolGUIDMappings[$i].AppId)*")){
                            $perAppTimes.Add( [pscustomobject]@{ 
                                'Time' = $event.TimeCreated
                                'Message' = $event.Message
                                'ID' = $event.Id
                                'DiskGUID' = $AppVolGUIDMappings[$i].DiskGUID
                                'AppDevice' = $AppVolGUIDMappings[$i].AppDevice
                                'AppGUID' = $AppVolGUIDMappings[$i].AppGUID
                                'AppName' = $AppVolGUIDMappings[$i].AppName
                                'AppId' = $AppVolGUIDMappings[$i].AppId
                                'EventRecord' = $event
                            })
                            #Write-Host "$($event | out-string)"
                        }
                    }
                    #Get svservice.log events
                    $svserviceLogObject | Where-Object { $_.Time -le $End -and $_.Time -ge $Start -and ( $_.message -like "*$($AppVolGUIDMappings[$i].DiskGUID)*" -or $_.message -like "*$($AppVolGUIDMappings[$i].AppDevice)*" -or $_.message -like "*$($AppVolGUIDMappings[$i].AppGUID)*" -or ( $AppVolGUIDMappings[$i].AppId -and $_.Message -like "*$($AppVolGUIDMappings[$i].AppId)*")) } | . { Process {
                        $perAppTimes.Add( [pscustomobject]@{ 
                            'Time' = $_.Time
                            'ID' = $_.ProcessInfo
                            'Message' = $_.Message
                            'DiskGUID' = $AppVolGUIDMappings[$i].DiskGUID
                            'AppDevice' = $AppVolGUIDMappings[$i].AppDevice
                            'AppGUID' = $AppVolGUIDMappings[$i].AppGUID
                            'AppName' = $AppVolGUIDMappings[$i].AppName
                            'AppId' = $AppVolGUIDMappings[$i].AppId
                            'EventRecord' =$null })
                    }}

                    #Get Process start events
                    $securityEvents | Where-Object { $_.Id -eq 4688 -and $_.TimeCreated -le $End -and $_.TimeCreated -ge $Start `
                            -and ($_.Properties[8].Value -like "*$($AppVolGUIDMappings[$i].DiskGUID)*" -or $_.Properties[8].Value -like "*$($AppVolGUIDMappings[$i].AppDevice)*" -or $_.Properties[8].Value -like "*$($AppVolGUIDMappings[$i].AppGUID)*" -or ( $AppVolGUIDMappings[$i].AppId -and $_.Properties[8].Value -like "*$($AppVolGUIDMappings[$i].AppID)*" )) } | . { Process {
                        $perAppTimes.Add( [pscustomobject]@{
                            'Time' = $_.TimeCreated
                            'ID' = $_.Id
                            'Message' = $_.Message
                            'DiskGUID' = $AppVolGUIDMappings[$i].DiskGUID
                            'AppDevice' = $AppVolGUIDMappings[$i].AppDevice
                            'AppGUID' = $AppVolGUIDMappings[$i].AppGUID
                            'AppName' = $AppVolGUIDMappings[$i].AppName
                            'AppId' = $AppVolGUIDMappings[$i].AppId
                            'EventRecord' = $_
                        })
                    }}
                }

                #Get Process End Events
                 ## -and $_.EventRecord.properties[$processName].Value -like '*\cmd.exe*'
                $perAppTimes += @( $perAppTimes | Where-Object { $_.Id -eq 4688 -and $_.EventRecord.properties[5].Value -like '*\cmd.exe' } | . { Process {
                    $event = $_
                    $BatchStartEvent = $_.EventRecord
                    ## $_.TimeCreated -eq $event.Time -and $_.Message -like $event.message } ) `
                    ##if( ( $BatchStartEvent = $securityEvents | Where-Object { $_.Id -eq 4688 -and $event.EventRecord.RecordId -eq $_.RecordId } ) `
                    if ( $BatchEndEvent = $securityEvents | Where-Object { $_.Id -eq 4689 -and $_.TimeCreated -ge $BatchStartEvent.TimeCreated -and $_.properties[$ProcessIdStop].Value -eq $BatchStartEvent.Properties[$ProcessIdNew].Value -and $_.properties[$processName].Value -like '*\cmd.exe*'} | Select-Object -First 1 ) {
                        Write-Debug "Found process end event: $(($BatchEndEvent).TimeCreated) : $(($BatchEndEvent).Message.Substring(0,20))"
                        #This finds process start events.  Need to find process end events (see further below)
                        [pscustomobject] @{
                            'Time' =  $BatchEndEvent.TimeCreated
                            'ID' =  $BatchEndEvent.Id
                            'Message' =  $BatchEndEvent.Message
                            'DiskGUID' =  $event.DiskGUID
                            'AppDevice' =  $event.AppDevice
                            'AppGUID' =  $event.AppGUID
                            'AppName' =  $event.AppName
                            'AppId' = $event.AppId
                            'EventRecord' =  $event }
                    }
                }})
                
                #measure appvolumes logon blocking event (Wait for volume(s) to mount) -- start (eventID 210) and end event "227" in - last 227 event if WaitForVolumes is false/first if true
                if( $startRecord = $AppVolWinEvents.Where( { $_.Id -eq "210" -and $_.Properties[1].Value -match "Session ID:\s*$SessionId$" } , 1 ) ) {
                    if ($global:WaitForFirstVolumeOnly) {
                        $endRecord = $AppVolWinEvents.Where( { $_.id -eq "226" -and $_.TimeCreated -ge $startRecord.TimeCreated } ) | Sort-Object -Property TimeCreated | Select-Object -First 1
                    } else {
                        $endRecord = $AppVolWinEvents.Where( { $_.id -eq "227" -and $_.TimeCreated -ge $startRecord.TimeCreated } ) | Sort-Object -Property TimeCreated | Select-Object -Last 1
                    }
                    if( $endRecord )
                    {
                        if( $Duration = New-TimeSpan -Start $StartRecord.TimeCreated -End $EndRecord.TimeCreated -ErrorAction SilentlyContinue )
                        {
                            ## as this is a blocking phase we add to main output
                            $Script:output.Add( [pscustomobject]@{ 
                                'Source' = 'App Volumes'
                                'PhaseName' = 'Wait For Volume Attach'
                                'Duration' = $Duration.TotalSeconds
                                'EndTime' = $EndRecord.TimeCreated
                                'StartTime' = $StartRecord.TimeCreated
                                })
                        }
                        else
                        {
                            Write-Debug -Message "Unable to get duration for measure app volumes logon blocking event"
                        }
                    }
                    else
                    {
                        ## Don't have disk info so can't check if mounted. Will be errors in the log file
                        $Warnings.Add( "Unable to find VMware App Volumes disk mount event for apps $((($applist | Select-Object -ExpandProperty AppName) -replace '!2B!' , '+' -replace '!20!' , ' ' ) -join ', ')" )
                    }
                }

                #OK, we finally have all of our information we need to measure everything. We will measure each app for each stage. For each stage
                #we are going to find the process start and end times and use the process end times to determine when the app completed it's work.
                #this is because the events AppVolumes generates are ambiguous for when a script finishes or when a stage was complete. If we can't find the event
                #we're going to have to rely on the events within the capture itself
                #Find first Logon Event
                $stages = @( "Prestartup" , "Startup_Postsvc" , "Startup" , "Logon" , "AllVolAttached" , "ShellStart" )

                foreach ($App in $AppVolGUIDMappings) {
                    if( $global:appVolumesVersion -ge '4.0' ) {
                        $perAppTimes | Where-Object { $_.AppGuid -eq $app.AppGUID -and $_.Message -like "*RunScript_VolumeScripts: user script*" }| Sort-Object -Property Time | . { Process {
                            $scriptStartEvent = $_
                            [int]$index = ($AppVolumesLogonEvents.FindIndex( {$args[0].Message -eq $ScriptStartEvent.Message } ))

                            if( $index -ge 0 -and ( $stage = $([regex]::Matches($ScriptStartEvent.Message, "(\[.*?\])") | Select-Object -First 1 -ErrorAction SilentlyContinue ) ))
                            {
                                Write-Verbose "AppVolumes: Phase Event: $($app.AppName) - $stage"
                                #Get launch event that should occur immediately after RunScript_
                                #Look for event 4688 and 4689 pairs
                                Write-Verbose "AppVolumes: Message to key in on: `"$($ScriptStartEvent.Message)`""

                                $StartEvent = $null
                                $EndEvent = $null
                                $result = $null
                                for ($a = $index ; $a -lt ($index + 8) -and ! $result ; $a++) {  #script launch events should occur within the next 4 lines (for AppVolumes 4) from where we found RunScript_Volume
                                    #sample line: CreateProcessCheckResult: Successfully launched: "C:\PROGRA~2\CLOUDV~1\Agent\Config\Default\app\SHELLS~1.BAT". WaitMilliseconds -1 ms, pid=6244 tid=2780
                                    #if ( $AppVolumesLogonEvents[$a].Message -like '*launch*BAT*pid=*' -and $AppVolumesLogonEvents[$a].ProcessInfo -eq $ScriptStartEvent.Id ) {  
                                    if ( $AppVolumesLogonEvents[$a].Message -match 'launch.*\.BAT\b.*\bpid=(\d+)' -and $AppVolumesLogonEvents[$a].ProcessInfo -eq $ScriptStartEvent.Id ) {  
                                        [int]$processId = $Matches[1]

                                        Write-Verbose "AppVolumes: Found PID : $processId"

                                       if( ! $result -and ( $startRecord = $securityEvents.Where( { $_.Id -eq 4688 -and $_.TimeCreated -ge $Start -and $_.Properties[$ProcessIdNew].Value -eq $processId -and $_.Properties[13].value.EndsWith( '\svservice.exe' )  -and $_.Properties[5].Value.EndsWith( '\cmd.exe' ) } )| Select-Object -Last 1 ) `
                                            -and ( $endRecord = $securityEvents.Where( { $_.Id -eq 4689 -and $_.TimeCreated -ge $Start -and $_.properties[$ProcessIdStop].Value -eq $startRecord.Properties[$ProcessIdNew].Value -and $_.properties[$processName].Value.EndsWith( '\cmd.exe' ) } )| Select-Object -Last 1 ) ) {
                                                 $result = [pscustomobject]@{
                                                        Source = 'App Volumes'
                                                        PhaseName = "$($stage -replace '[\[\]]') `"$($App.AppName)`""
                                                        Duration  = ($endRecord.TimeCreated - $startRecord.TimeCreated).TotalSeconds
                                                        EndTime   = $endRecord.TimeCreated
                                                        StartTime = $startRecord.TimeCreated }
                                        }
                                    }
                                }
                                if( $result )
                                {
                                    $Script:AppVolumesOutput.Add( $result )
                                }
                                else
                                {
                                    Write-Debug -Message "Failed to find launch event for app $($App.Name) $($scriptStartEvent.Message)"
                                }
                            }
                        }}
                    }
                    else ## not v4.0 or higher
                    {
                        foreach ($stage in $stages) {
                            [bool]$foundStage = $false
                            try {
                                if( ! $foundStage -and ( $startRecord = $perAppTimes | Where-Object { $_.AppName -eq $App.AppName -and $_.PSObject.Properties[ 'EventRecord' ] -and $_.EventRecord.Id -eq 4688 -and $_.EventRecord.Properties[8].value -like "*\$stage.bat*" } | Select-Object -First 1 -ExpandProperty EventRecord ) `
                                    -and ( $EndRecord = ($securityEvents | Where-Object { $_.Id -eq 4689 -and $_.TimeCreated -ge $StartRecord.TimeCreated -and $_.properties[$ProcessIdStop].Value -eq $StartRecord.Properties[$ProcessIdNew].Value -and $_.properties[$processName].Value -like '*\cmd.exe'} | Select-Object -First 1 ) ) )
                                {
                                    $result = [pscustomobject]@{
                                        'Source' = 'App Volumes'
                                        'PhaseName' = "$stage `"$($App.AppName)`""
                                        'Duration' = ($EndRecord.TimeCreated - $startRecord.TimeCreated).TotalSeconds
                                        'EndTime' = $EndRecord.TimeCreated
                                        'StartTime' = $StartRecord.TimeCreated
                                    }
                                    if( $appVolumesVersion -le [version]'2.12.9999.0' )
                                    {
                                        $Script:Output.Add( $result )
                                    }
                                    else ## not blocking so we will list afterwards
                                    {
                                        $Script:AppVolumesOutput.Add( $result )
                                    }
                                    $foundStage = $true ## only want one per phase
                                }
                            }
                            catch {
                                Write-Debug "Found no start or stop events for $($App.AppName) in $($stage)"
                            }
                        }
                    }
                }

                <#

                # GRL Code commented out since AppVolumes events are output as a separate phase so can't intertwine them with the Windows phases

                $AppVolumesLogonPhase = New-Object -TypeName "System.Collections.Generic.List[psobject]"
                $AppVolumesShellStartPhase =  New-Object -TypeName "System.Collections.Generic.List[psobject]"

                foreach ($phase in $Script:AppVolumesOutput ) {
                    if (($phase.phaseName -like "*prestartup*") -or ($phase.phaseName -like "*startup_postsvc*") -or ($phase.phaseName -like "*startup*") -or ($phase.phaseName -like "*logon*") -or ($phase.phaseName -like "*allvolattached*")) {
                        $AppVolumesLogonPhase.Add($phase)
                    }

                    if (($phase.phaseName -like "*shellstart*")) {
                        $AppVolumesShellStartPhase.Add($phase)
                    }

                }

                if( $AppVolumesLogonPhase -and ( $Duration = New-TimeSpan -Start ($AppVolumesLogonPhase | Sort -Property StartTime)[0].StartTime -End ($AppVolumesLogonPhase | Sort -Property EndTime -Descending)[0].EndTime )) {
                    $Script:AppVolumesOutput.Add( [pscustomobject]@{ 
                        'PhaseName' = "AppVolumes - Logon"
                        'Duration'  = $Duration.TotalSeconds
                        'EndTime'   = $AppVolumesLogonPhase | Sort-Object -Property EndTime -Descending | Select-Object -First 1 -ExpandProperty EndTime
                        'StartTime' = $AppVolumesLogonPhase | Sort-Object -Property StartTime | Select-Object -First 1 -ExpandProperty StartTime
                    } )
                }

                if( $AppVolumesShellStartPhase -and ( $Duration = New-TimeSpan -Start ($AppVolumesShellStartPhase | Sort -Property StartTime)[0].StartTime -End ($AppVolumesShellStartPhase | Sort -Property EndTime -Descending)[0].EndTime ) ) {
                    $ScriptAppVolumesOutput.Add( [pscustomobject]@{
                        'PhaseName' = "AppVolumes - ShellStart"
                        'Duration' = $Duration.TotalSeconds
                        'EndTime' = $AppVolumesShellStartPhase | Sort-Object -Property EndTime -Descending| Select-Object -First 1 -ExpandProperty EndTime
                        'StartTime' = $AppVolumesShellStartPhase | Sort-Object -Property StartTime | Select-Object -First 1 -ExpandProperty StartTime
                    })
                }3
                #>

            } else {
                Write-Verbose "AppVolumes log file `"$appVolumesLogFile`" not found. Skipping AppVolumes enumeration"
            }
        }

        ## Set up runspacepool as we will parallelise some operations
        $SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        
        ## need to import the functions we need from this module
        @( 'New-XPath' , 'Get-PhaseEvent' , 'Get-EventLogEnabledStatus' ) | ForEach-Object `
        {
            $function = $_
            $Definition = Get-Content Function:\$function -ErrorAction Continue
            $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $function , $Definition
            $sessionState.Commands.Add($SessionStateFunction)
        }

        $RunspacePool = [runspacefactory]::CreateRunspacePool(
            1, ## Min Runspaces
            10 , ## Max parallel runspaces ,
            $sessionstate ,
            $host
        )
        
        $sharedVars = [hashtable]::Synchronized(@{})
        $RunspacePool.Open()
        $tsevent = $null
        $logonEvent = $null
        $UserLogon = $null
        $wmiEvent = $null
        $jobs = New-Object System.Collections.ArrayList
        $prelogonData = New-Object -TypeName System.Collections.Generic.List[psobject]
        $odataPhase = $null

        [string]$initialProgram = $null

        if( $offline )
        {
            $logon = Get-Content -Path (Join-Path -Path $global:logsFolder -ChildPath 'logon.json' ) | ConvertFrom-Json
            $logon.LogonTime = [DateTime]::FromFileTime( $logon.LogonTimeFileTime ) ## have to use this absolute figure otherwise is wrong timezone potentially
            $logon.UserSID = New-Object System.Security.Principal.SecurityIdentifier -ArgumentList $logon.UserSID.Value
            $global:windowsMajorVersion = $logon.OSversion
            $ClientName = $logon.ClientName
            $CUDesktopLoadTime = $logon.CUDesktopLoadTime
            $initialProgram = $logon.InitialProgram
        }
        else
        {
            $initialProgram = Get-ItemProperty -Path "HKLM:\SOFTWARE\Citrix\Ica\Session\$SessionId\Connection" -Name InitialProgram -ErrorAction SilentlyContinue | Select-Object -ExpandProperty InitialProgram
            $OS = Get-CimInstance -ClassName win32_operatingsystem -ErrorAction SilentlyContinue
            $CS = Get-CimInstance -ClassName win32_computersystem -ErrorAction SilentlyContinue
            if( $OS )
            {
                Write-Debug "OS is $($OS.Caption) $($OS.Version), last booted $(Get-Date $OS.LastBootupTime -Format G), PowerShell $($PSVersionTable.PSVersion.ToString())"
            }
            if( $CS )
            {
                Write-Debug "Manufacturer $($CS.Manufacturer) model $($CS.Model), name $($CS.Name) domain $($CS.Domain) virtual $($CS.HypervisorPresent)"
            }

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
                    OSversion = $global:windowsMajorVersion
                    ClientName = $ClientName
                    CUDesktopLoadTime = $CUDesktopLoadTime
                    InitialProgram = $initialProgram
                    UserName = $Username
                    UserDomain = $UserDomain
                    ## No point saving XD details since these cannot be used offline
                }
                if( $dumpForOffline )
                {
                    if( $logon )
                    {
                        $logon | ConvertTo-Json | Set-Content -Path (Join-Path -Path $global:logsFolder -ChildPath 'logon.json' )
                    }

                    Write-Debug "Required files dumped to `"$logsFolder`". Please zip and email to support@controlup.com"
                }
            }
            else
            {
                Throw "Failed to retrieve logon session for $UserDomain\$Username from LSASS"
            }
        }

        Write-Debug "Logon data: $Logon Logon Ids $($logon.LogonID -join ' , ')"
    }

    process {          
            [hashtable]$parameters = @{
                'UserName' = $userName
                'UserDomain' = $UserDomain
                'Logon' = $logon
                'SharedVars' = $sharedVars
                'UserProfileEventFile' = $global:userProfileParams[ 'Path' ]
                'GroupPolicyEventFile' = $global:groupPolicyParams[ 'Path' ]
                'CitrixUPMEventFile'   = $global:citrixUPMParams[ 'Path' ]
             }

        # If the machine is a Citrix VDA and a Session ID is provided, look for "HDX Connection" Phase
        $odataPhase = $null
        [string]$profilerDataJsonFile = $(if( $offline -and $global:logsFolder ){ (Join-Path -Path $global:logsFolder -ChildPath 'profilerdata.json' ) })

        if( $offline -and ! ( Test-Path -Path $profilerDataJsonFile -ErrorAction SilentlyContinue ) )
        {
            Write-Debug "Skipping HDX check as in offline mode"
        }
        elseif ((Get-Service -Name BrokerAgent -ErrorAction SilentlyContinue) -and $HDXSessionId  ) {
            [System.Collections.Generic.List[psobject]]$prelogonData = Get-CitrixData -sessionId $HDXSessionId

            if( ( $citrixClient = Get-CimInstance -Namespace root\Citrix\hdx -ClassName Citrix_Client | Where-Object SessionId -eq $sessionId ) `
                -or ( $citrixClient = Get-CimInstance -Namespace root\Citrix\hdx -ClassName Citrix_Client_Enum | Where-Object SessionId -eq $sessionId ) )
            {
                $odataPhase = [pscustomobject]@{
                    'Client Name' = $citrixClient.Name
                    'Client Version' = $citrixClient.Version
                    'Client Address' = $citrixClient.Address }
                if( $vdaVersion = Get-Process -name BrokerAgent -ErrorAction SilentlyContinue|Get-ItemProperty -ErrorAction SilentlyContinue|Select-Object -ExpandProperty versioninfo|Select-Object -ExpandProperty productversion )
                {
                    Add-Member -InputObject $odataPhase -MemberType NoteProperty -Name 'VDA Version' -Value $vdaVersion
                }
            }
        }
        elseif( (Get-Service -Name WSNM -ErrorAction SilentlyContinue) -or ( ! [string]::IsNullOrEmpty(  $profilerDataJsonFile ) -and ( Test-Path -Path $profilerDataJsonFile -ErrorAction SilentlyContinue ) ) )
        {
            ## VMware Horizon View Agent
            [string]$horizonSessionKey = "HKLM:\SOFTWARE\VMware, Inc.\VMware VDM\SessionData\$SessionId"
            [hashtable]$onlineOfflineTS = @{}
            if( $global:terminalServicesParams[ 'Path' ] )
            {
                $onlineOfflineTS.Add( 'Path' , $global:terminalServicesParams[ 'Path' ] )
            }
            $odataPhase = [pscustomobject]@{}
            
            $horizonInfoValues = Get-ItemProperty -Path $horizonSessionKey -ErrorAction SilentlyContinue
            if( ! $offline -and ! $horizonInfoValues )
            {
                $VDMversion = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\VMware, Inc.\VMware VDM' -Name 'ProductVersion' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'ProductVersion') -as [version]
                if( ! $VDMversion -or $VDMversion.Major -lt 7 )
                {
                   $script:warnings.Add( "At least version 7.x of VMware Horizon View is required, detected version was $($VDMversion.ToString())" )
                }
                elseif( $currentSessionState -eq 4 )
                {
                    [string]$warningMessage = "Session is currently disconnected "
                    if( $disconnectionEvent = Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable ( @{ StartTime = $logon.LogonTime ; EndTime = [datetime]::Now ; Id = 24 ; ProviderName = 'Microsoft-Windows-TerminalServices-LocalSessionManager' } + $onlineOfflineTS ) | Where-Object { $_.Properties[0].Value -eq "$($Logon.Userdomain)\$($logon.username)" -and $_.Properties[1].Value -eq $sessionId } | Select-Object -First 1 )
                    {
                        $warningMessage += "(at $(Get-Date -Date $disconnectionEvent.TimeCreated -Format G)) "
                    }
                    $warningMessage += "so VMware Horizon View session key has been deleted meaning that data is not available"
                    $script:warnings.Add( $warningMessage )
                }
                else
                {
                    [string]$message = "Horizon SessionData registry key `"$horizonSessionKey`" not yet present - it can take up to 10 minutes after logon to appear - logon was $([math]::Round( ([datetime]::Now - $logon.LogonTime).Minutes , 1 )) minutes ago"
                    
                    ## see if parent key exists as this has been observed to be missing
                    [string]$parentKey = Split-Path -Path $horizonSessionKey -Parent
                    if( ! ( Test-Path -Path $parentKey -PathType Container -ErrorAction SilentlyContinue ) )
                    {
                        $message += " (parent key `"$(Split-Path -Path $parentKey -Leaf)`" is missing)"
                    }
                    $script:warnings.Add( $message )
                    ## if whole key not there soon after logon then almost certainly RDP as PCOIP and BLAST cause session key to be created
                    Add-Member -InputObject $odataPhase -MemberType NoteProperty -Name 'Display Protocol' -Value 'RDP'
                }
            }
            else
            {
                if( $horizonInfoValues )
                {
                    ## Use Citrix phase to pass back info we want displayed first
                    ## values are missing for RDP protocol
                    [hashtable]$sessionProperties = @{ 'Display Protocol' = 'RDP' }
                    if( $horizonInfoValues.PSObject.Properties[ 'ViewClient_Protocol' ] )
                    {
                        $sessionProperties.'Display Protocol' = $horizonInfoValues.ViewClient_Protocol
                    }
                    if( $horizonInfoValues.PSObject.Properties[ 'ViewClient_Machine_Name' ] )
                    {
                        $sessionProperties.Add( 'Client Name' , $horizonInfoValues.ViewClient_Machine_Name )
                    }
                    if( $horizonInfoValues.PSObject.Properties[ 'ViewClient_Broker_DNS_Name' ] )
                    {
                        $sessionProperties.Add( 'Broker' , $horizonInfoValues.ViewClient_Broker_DNS_Name )
                    }
                    Add-Member -InputObject $odataPhase -NotePropertyMembers $sessionProperties
                }

                ## See if profilerdata exists yet - it can take up to 10 minues to appear!
                if( ! $offline -and ! $horizonInfoValues.PSObject.Properties[ 'ProfilerData' ] )
                {
                    $script:warnings.Add( "Horizon ProfilerData registry value not yet present in VMware Horizon View session key `"$horizonSessionKey`" - it can take up to 10 minutes after logon to appear - logon was $([math]::Round( ([datetime]::Now - $logon.LogonTime).Minutes , 1 )) minutes ago" )
                }
                else
                {
                    if( $dumpForOffline )
                    {
                        $horizonInfoValues.ProfilerData | Out-File -FilePath $profilerDataJsonFile
                    }
                    if( $offline -and ( Test-Path -Path $profilerDataJsonFile -ErrorAction SilentlyContinue ) )
                    {
                        $profilerData = Get-Content -Path $profilerDataJsonFile | ConvertFrom-Json
                    }
                    else
                    {
                        $profilerData = $horizonInfoValues.ProfilerData | ConvertFrom-Json
                    }

                    if( $profilerData )
                    {
                        ## If session has reconnected then profilerdata is for the reconnection not the original logon so no point showing it
                        if( ( $brokerTime = Get-JSONProperty -inputObject $profilerData -name 'broker' ) `
                         -and ($StartTime = (Get-Date -Date $brokertime.value.s).ToLocalTime()) `
                            -and $StartTime -gt $logon.LogonTime )
                        {
                            ## Look for disconnect and reconnect events for this session and user between these two times
                            [array]$connectionEvents = @( Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable ( @{ StartTime = $logon.LogonTime ; EndTime = $StartTime.AddSeconds( 120 ) ; Id = @( 24 , 25) ; ProviderName = 'Microsoft-Windows-TerminalServices-LocalSessionManager' } + $onlineOfflineTS ) | Where-Object { $_.Properties[0].Value -eq "$($Logon.Userdomain)\$($logon.username)" -and $_.Properties[1].Value -eq $sessionId } )
                            if( $connectionEvents -and $connectionEvents.Count )
                            {
                                [string]$warningMessage = "Session "
                                if( $disconnectedEvent = $connectionEvents | Where-Object { $_.Id -eq 24 }  | Select-Object -First 1 )
                                {
                                    $warningMessage += "disconnected at $(Get-Date -Date $disconnectedEvent.TimeCreated -Format G) "
                                }
                                if( $reconnectedEvent = $connectionEvents | Where-Object { $_.Id -eq 25 }  | Select-Object -First 1 )
                                {
                                    if( $disconnectedEvent )
                                    {
                                        $warningMessage += 'and '
                                    }
                                    $warningMessage += "reconnected at $(Get-Date -Date $reconnectedEvent.TimeCreated -Format G) "
                                }
                                $warningMessage += 'so ignoring VMware brokering data which is for the reconnection'
                                $script:warnings.Add( $warningMessage )
                            }
                            else
                            {
                                $script:warnings.Add( "VMware brokering event is $([math]::Round( ($StartTime - $logon.LogonTime).TotalMinutes , 1 ) ) minutes after logon but unable to find evidence of disconnect & reconnect in event log" )
                            }
                        }
                        else ## profilerdata data is for the logon so get the VMware phases
                        {
                            ForEach( $jsonProperty in @( 'broker' , 'authentication' , 'protocol-connection' , 'clientConnectWait' , 'appLaunch' , 'agentPrepare' , 'protocolStartup' ))
                            {
                                If( $property = Get-JSONProperty -inputObject $profilerData -name $jsonProperty )
                                {
                                    Try
                                    {
                                        $object = [pscustomobject]@{
                                            Source = 'Horizon'
                                            PhaseName =  (Get-Culture).TextInfo.ToTitleCase( ($jsonProperty -creplace '([A-Z])' , ' $1' -replace '(\-)' , ' ') )
                                            StartTime = (Get-Date -Date $property.value.s).ToLocalTime()
                                            EndTime   = (Get-Date -Date $property.value.e).ToLocalTime()
                                            Duration  = ($property.value.d -as [int]) / 1000 }
                                    }
                                    Catch
                                    {
                                        $object = $null
                                    }

                                    if( $object -and $object.Duration -gt 0 )
                                    {
                                        $prelogonData.Add( $object )
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        $script:warnings.Add( "Failed to translate JSON session data information in VMware Horizon View session key `"$horizonSessionKey`"" )
                    }
                }
            }
        }
        
        [hashtable]$securityFilter = @{StartTime=$logon.LogonTime;EndTime=($logon.LogonTime.AddMinutes( 60 ));Id=4018,5018,4688,4689}
        if( $securityParams[ 'Path' ] )
        {
            $securityFilter.Add( 'Path' , $securityParams[ 'Path' ] )
        }
        else
        {
            $securityFilter.Add( 'LogName' , 'Security' )
        }
        [array]$securityEvents = @( Get-WinEvent -FilterHashtable $securityFilter -ErrorAction SilentlyContinue)
        if( ! $securityEvents -or ! $securityEvents.Count )
        {
            Write-Error "Failed to cache any relevant security event logs from $(Get-Date $logon.LogonTime -Format G) for 60 minutes"
        }
        
        ## Get CSE finishes as we may need them for VMware DEM but if we don't they are useful/interesting anyway
        
        ## Find event id 4001 from GP log so we can get activity id to cross ref to 5016 event for finishing of GPO processing
        ## TODO make work offline
        [string]$query = "*[EventData[Data[@Name='PrincipalSamName'] and (Data='$($logon.UserDomain)\$($logon.Username)')]] and *[System[(EventID='4001')]]"
        $CSEArray = $null
        [hashtable]$CSE2GPO = @{}

        if( $startProcessingEvent = Get-WinEvent -ProviderName Microsoft-Windows-GroupPolicy -FilterXPath $query -MaxEvents 1 -ErrorAction SilentlyContinue )
        {
            $query = "*[System[(EventID='4016' or EventID='5016' or EventID='6016' or EventID='7016') and TimeCreated[@SystemTime>='$($startProcessingEvent.TimeCreated.ToUniversalTime().ToString("s")).$($startProcessingEvent.TimeCreated.ToUniversalTime().ToString("fff"))Z'] and Correlation[@ActivityID='{$($startProcessingEvent.ActivityID.Guid)}']]]"
            if( ! ( $CSEarray = @( Get-WinEvent -ProviderName Microsoft-Windows-GroupPolicy -FilterXPath $query -ErrorAction SilentlyContinue ) ) -or ! $CSEArray.Count )
            {
                $warnings.Add( "Failed to find any group policy event id 5016 instances for CSE finishes" )
            }
            else
            {
                ## build hash table of cse id and GPO names so we can output when we iterate over finish events later
                $CSEArray.Where( { $_.Id -eq 4016 } ).ForEach( `
                {
                    $CSE2GPO.Add( $_.Properties[0].Value , $_.Properties[5].Value )
                })
            }
        }
        else
        {
            $warnings.Add( "Failed to find group policy processing starting event id 4001" )
        }

        ## TODO make work offline - difficult given we are looking at registry
        if( Get-Service -Name ImmidioFlexProfiles -ErrorAction SilentlyContinue )
        {
            ## Need to see if we are being run via GPO or logon script          
            if ( ! (Test-Path -Path HKU:\ -ErrorAction SilentlyContinue)) 
            {
                if( ! ( New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS ) )
                {
                    $warnings.Add( "Unable to map HKEY_USERS" )
                }
            }
            [string]$productName = $null

            ## get all flexeengine processes first as we use several. Need to check source and target subjectlogonid and domain\user as will use different ones depending on if run via CSE or logon script
            [array]$flexEngineStarts = @( $securityEvents.Where( { $_.Id -eq 4688 -and (($_.Properties[$TargetLogonId].value -in $Logon.LogonId `
                -and $_.Properties[$TargetUserName ].value -eq $Username -and $_.Properties[$TargetDomainName ].value -eq $UserDomain ) `
                    -or ($_.Properties[$SubjectLogonId].value -in $Logon.LogonId `
                        -and $_.Properties[$SubjectUserName ].value -eq $Username -and $_.Properties[$SubjectDomainName ].value -eq $UserDomain ))`
                            -and $_.properties[$NewProcessName].Value -match '\\flexengine\.exe$' -and $_.TimeCreated -ge $logon.LogonTime } ) )

            if( ! ( $immidioKey = Get-Item -Path "HKU:\$($Logon.UserSID.Value)\Software\Policies\Immidio\Flex Profiles\Arguments" -ErrorAction SilentlyContinue ) )
            {
                $warnings.Add( "Unable to get VMware DEM GPO settings from HKCU" )
            }
            elseif( $immidioKey.GetValue('GPClientSideExtension') -eq [int]1 ) ## GPO
            {
                Write-Verbose -Message "VMware DEM running as CSE"
                ## reported later in CSEs
                <#
                if( $CSEarray -and $CSEarray.Count )
                {
                    # TODO what about older versions, eg. UEM?
                    if( ! ( $DEMEvent = $CSEarray.Where( { $_.Properties[2].Value -match 'VMware\b.*\bEnvironment Manager' -or $_.Properties[2].Value -match 'VMware\b.*\bUEM\b' } , 1 ) ) )
                    {
                        $warnings.Add( "Failed to find group policy event log id 5016 for VMware DEM" )
                    }
                    else
                    {
                        $Script:output.Add( ([pscustomobject]@{
                                        Source    = 'VMware DEM'
                                        PhaseName = 'Logon'
                                        StartTime = $DEMEvent.TimeCreated
                                        EndTime   = $DEMEvent.TimeCreated.AddMilliseconds( -$DEMEvent.Properties[0].Value )
                                        Duration  = $DEMEvent.Properties[0].Value / 1000 }))
                    }
                }
                #>
            }
            else ## run by logon script so we look for the flexengine processes with specific command lines for this user after logon
            {              
                Write-Verbose -Message "VMware DEM not running as CSE"
                ## look for flexengine.exe -r but don't insist as parent of gpscript.exe (logon script) in case launched some other way
                if( ! $flexEngineStarts -or ! $flexEngineStarts.Count )
                {
                    $warnings.Add( "Failed to find any flexengine.exe process start events for VMware DEM" )
                }
                else
                {
                    ## find the flexengine start with -r argument - could be quoted or unquoted path to flexengine.exe. Need to exclude -ra and ::Async -r calls
                    ## Don't use .Where() as doesn't work with $Matches
                    ## "C:\Program Files\Immidio\Flex Profiles\FlexEngine.exe" -r
                    if( $flexengineMinusRStart = $flexEngineStarts | Where-Object { ( $_.Properties[$NewProcessCmdLine].Value -match '^"([^"]+)"\s+(.*)$' -or $_.Properties[$NewProcessCmdLine].Value -match '^([^\s]+)\s+(.*)$' ) `
                        -and ($theMatch = $Matches[2]) -like '*-r*' -and $theMatch -notlike '*-ra*' -and $theMatch -notlike '*::*' } | Select-Object -Last 1)
                    {
                        [string]$executable = $matches[1]
                        ## find stop event
                        if( $flexengineMinusRStop = $securityEvents.Where( { $_.Id -eq 4689 -and $_.TimeCreated -ge $flexengineMinusRStart.TimeCreated -and $_.Properties[$ProcessIdStop].value -eq $flexengineMinusRStart.Properties[$ProcessIdNew].value `
                            -and $_.Properties[$SubjectLogonId].value -eq $flexengineMinusRStart.Properties[$SubjectLogonId].value } ) | Select -Last 1 )
                        {
                            ## try and pull the product name from the flexengine.exe file so we can report if DEM, UEM, etc
                            $productName = $(if( ! [string]::IsNullOrEmpty( $executable ) -and ($properties = Get-ItemProperty -Path $executable -ErrorAction SilentlyContinue) ) { $properties | Select-Object -ExpandProperty VersionInfo | Select-Object -ExpandProperty ProductName })
                            if( [string]::IsNullOrEmpty( $productName ) )
                            {
                                $productName = 'VMware DEM'
                            }
                            $Script:output.Add( ([pscustomobject]@{
                                Source    = $productName
                                PhaseName = 'Logon'
                                StartTime = $flexengineMinusRStart.TimeCreated
                                EndTime   = $flexengineMinusRStop.TimeCreated
                                Duration  = ($flexengineMinusRStop.TimeCreated - $flexengineMinusRStart.TimeCreated).TotalSeconds }))
                        }
                        else
                        {
                            $warnings.Add( "Failed to find process terminated event for '$($flexengineMinusRStart[$NewProcessCmdLine].Value)' started at $(Get-Date -Date $flexengineMinusRStart.TimeCreated -Format G)" )
                        }
                    }
                    else
                    {
                        $warnings.Add( "Unable to find VMware DEM flexengine.exe process start with -r argument" )
                    }
                }
            }
            ## see if any async flexengine runs and add those to a separate list for displaying in the non-blocking section
            $flexEngineStarts | Where-Object { ( $_.Properties[$NewProcessCmdLine].Value -match '^"([^"]+)"\s+(.*)$' -or $_.Properties[$NewProcessCmdLine].Value -match '^([^\s]+)\s+(.*)$' ) -and (( $arguments = $Matches[2] ) -like '*::Async*' -or $arguments -like '*::DefaultApplications*' ) } | ForEach-Object `
            {
                $asyncFlexengineStart = $_
                ## find stop event - may not be same subjectlogonid as may be launched in generic 999 but end up in the user's session
                if( ! ( $asyncFlexengineStop = $securityEvents.Where( { $_.Id -eq 4689 -and $_.TimeCreated -ge $asyncFlexengineStart.TimeCreated -and $_.Properties[$ProcessIdStop].value -eq $asyncFlexengineStart.Properties[$ProcessIdNew].value `
                    -and ( $_.Properties[$SubjectLogonId].value -eq $asyncFlexengineStart.Properties[$SubjectLogonId].value -or $_.Properties[$ProcessStopSid].Value -eq $logon.UserSID.Value ) } ) | Select -Last 1 ) )
                {
                    $warnings.Add( "Failed to find process terminated event for '$($asyncFlexengineStart.Properties[$NewProcessCmdLine].Value)' started at $(Get-Date -Date $asyncFlexengineStart.TimeCreated -Format G)" )
                }
                if( [string]::IsNullOrEmpty( $productName ) )
                {
                    if( [string]::IsNullOrEmpty( ( $productName = $(if( ! [string]::IsNullOrEmpty( $Matches[1] ) -and ($properties = Get-ItemProperty -Path $Matches[1] -ErrorAction SilentlyContinue) ) { $properties | Select-Object -ExpandProperty VersionInfo | Select-Object -ExpandProperty ProductName }) )))
                    {
                        $productName = 'VMware DEM'
                    }
                }
                $script:vmwareDEMNonBlockingPhases.Add( ([pscustomobject]@{
                    Source    = $productName
                    PhaseName = $arguments -replace '.*::(\w+).*$' , '$1' -creplace '([a-z])([A-Z])' , '$1 $2' ## turn "DefaultApplications" into "Default Applications"
                    StartTime = $asyncFlexengineStart.TimeCreated
                    EndTime   = $asyncFlexengineStop | Select-Object -ExpandProperty TimeCreated
                    Duration  = $(if( $asyncFlexengineStop ) { ($asyncFlexengineStop.TimeCreated - $asyncFlexengineStart.TimeCreated).TotalSeconds } )}))
            }
        }

        ## 14/05/19 GRL - if published app then logon finished when icast.exe exits, for published desktop it's explorer.exe start
        [bool]$isPublishedApp = $false
        [bool]$isScript = $false
        $logonFinishedEvent = $null
        [int]$shellPid = -1
        [string]$shellProgram = $null
        [string]$publishedApp = $null
        [string]$publishedAppParameters = $null

        ## Grab the first exe, which is usually icast.exe, as that's the process we look for. If published desktop then value won't exist or will be empty
        if( ! [string]::IsNullOrEmpty( $initialProgram ) )
        {
            if( $initialProgram -match '^"([^"]*)"\s*"([^"]*)"(\s*.*)?' -or $initialProgram -match '^([^\s]*)\s*"([^"]*)"(\s*.*)?' ) ## if icast.exe used then published app will always be "quoted"
            {
                ## look for the published app/script - if a script then figure out what the process would be that launches it
                $shellProgram = $Matches[ 1 ]
                $publishedApp = $Matches[ 2 ]
                $publishedAppParameters = $( if( $Matches[3] ) { $Matches[ 3 ].Trim() } )
                $isPublishedApp = $true
                Write-Debug "Published app detected for session $sessionId (`"$initialProgram`") shell `"$shellProgram`" published app `"$publishedApp`" with parameters `"$publishedAppParameters`""

                ## Executable for published app may have been specified without a full path but events will have path so get the full path
                $publishedApp = [System.IO.Path]::GetFullPath( $( switch ( [System.IO.Path]::GetExtension( $publishedApp ) )
                {
                    ## seems that .vbs scripts must be specified via wscript or cscript as the executable
                    '.cmd'  { Join-Path -Path ([environment]::GetFolderPath('System')) -ChildPath 'cmd.exe' ; $isScript = $true } 
                    default { [System.Environment]::ExpandEnvironmentVariables( $publishedApp ) }
                }))
            }
            else
            {
                Write-Error "Unable to retrieve published app from `"$initialProgram`""
            }
        }
        else ## published desktop so logon finished is when explorer starts 
        {
            Write-Debug "Published desktop detected for session $sessionId"
            $publishedApp = $shellProgram = (Join-Path -Path $env:SystemRoot -ChildPath 'explorer.exe' )
        }
        
        $userinitStartEvent = $null

        if( $global:windowsMajorVersion -ge 10 )
        {
            $userinitStartEvent = ($securityEvents | Where-Object { $_.Id -eq 4688 -and $_.Properties[$TargetLogonId].value -in $Logon.LogonId `
                -and $_.Properties[$TargetUserName ].value -eq $Username -and $_.Properties[$TargetDomainName ].value -eq $UserDomain -and $_.properties[$NewProcessName].Value -eq (Join-Path -Path ([environment]::GetFolderPath('System')) -ChildPath 'userinit.exe' ) } | Select -Last 1 )
        }
        ## else older OS where we don't have enough properties in the process started events to get what we need so will have to look up later

        if( ! [string]::IsNullOrEmpty( $publishedApp ) )
        {
            ## look for the process start event for the shell (explorer.exe) or pubished app by finding the process start after logon for this user with the same logonid. Select last one in case manually restarted in the session
            if( $userinitStartEvent ) ## we have userinit pid so get published app which isn't a child of this process (e.g. if cmd.exe then don't grab logon scripts but if explorer then must be child of userinit.exe)
            {
                if( $isPublishedApp )
                {
                    $logonFinishedEvent = ($securityEvents | Where-Object { $_.Id -eq 4688 -and $_.Properties[$SubjectLogonId].value -in $Logon.LogonId `
                        -and $_.Properties[$SubjectUserName ].value -eq $Username -and $_.Properties[$SubjectDomainName ].value -eq $UserDomain `
                            -and $_.properties[$NewProcessName].Value -eq $publishedApp `
                                 -and $_.Properties[$ProcessIdStart].value -ne $userinitStartEvent.Properties[$ProcessIdNew].value} ) | Select -Last 1
                }
                else
                {
                    $logonFinishedEvent = ($securityEvents | Where-Object { $_.Id -eq 4688 -and $_.Properties[$SubjectLogonId].value -in $Logon.LogonId `
                        -and $_.Properties[$SubjectUserName ].value -eq $Username -and $_.Properties[$SubjectDomainName ].value -eq $UserDomain `
                            -and $_.properties[$NewProcessName].Value -eq $publishedApp `
                                 -and $_.Properties[$ProcessIdStart].value -eq $userinitStartEvent.Properties[$ProcessIdNew].value} ) | Select -Last 1
                }
            }
            if( ! $logonFinishedEvent -and $SearchCommandLine -and  ! [string]::IsNullOrEmpty( $publishedAppParameters ) ) ## we have parameters so look for those in process invocation
            {
                $logonFinishedEvent = ($securityEvents | Where-Object { $_.Id -eq 4688 -and $_.Properties[$SubjectLogonId].value -in $Logon.LogonId `
                    -and $_.Properties[$SubjectUserName ].value -eq $Username -and $_.Properties[$SubjectDomainName ].value -eq $UserDomain `
                        -and $_.properties[$NewProcessName].Value -eq $publishedApp -and $_.Properties[$NewProcessCmdLine].Value -match [regex]::Escape( $publishedAppParameters ) } ) | Select -Last 1 
            }
            if( ! $logonFinishedEvent ) ## probably older OS so we don't have userinit pid yet
            {
                $logonFinishedEvent = ($securityEvents | Where-Object { $_.Id -eq 4688 -and $_.Properties[$SubjectLogonId].value -in $Logon.LogonId `
                    -and $_.Properties[$SubjectUserName ].value -eq $Username -and $_.Properties[$SubjectDomainName ].value -eq $UserDomain `
                        -and $_.properties[$NewProcessName].Value -eq $publishedApp } ) | Select -Last 1 
            }
            if( $logonFinishedEvent )
            {
                $shellPid = $logonFinishedEvent.Properties[$ProcessIdNew].Value
            }
        }
        if( $logonFinishedEvent )
        {
            Write-Debug "Got logon finished time of $((Get-Date -Date $logonFinishedEvent.TimeCreated).ToString( 'hh:mm:ss.fff' )), shell pid $shellPid"
        }
        else
        {
            Write-Debug "Failed to get logon finished time"
        }

        ## This doesn't work on Win7 & 2008R2/2012R2 as target username and domainname don't exist - event only have first 9 properties
        if( ! $userinitStartEvent )
        {
            if( $logonFinishedEvent )
            {
                ## shell will have been spawned by userinit.exe whose pid is $ProcessIdStart of $shellStart so now we can find that
                $userinitStartEvent = ($securityEvents  |Where-Object { $_.Id -eq 4688 -and $_.Properties[$ProcessIdNew].Value -eq $logonFinishedEvent.Properties[$ProcessIdStart].Value `
                    -and $_.properties[$NewProcessName].Value -eq (Join-Path -Path ([environment]::GetFolderPath('System')) -ChildPath 'userinit.exe' ) } | Select -Last 1 )
            }
            else
            {
                Write-Debug "Couldn't find a shell process event for user"
            }
        }
        
        ## Now that we don't user the security logon event 4624, we need another way to get the Winlogon PID to be able to find the mpnotify event but it's only required for that
        ## From the userinit start event, we have the pid of winlogon as that is the parent process

        ## only seem to get this when RDS role is installed (Multi-user Win10 or Server OS)
        if( ! $offline -and (Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Terminal Server\Compatibility' -ErrorAction SilentlyContinue) )
        {
            $networkStartEvent = $null
            if( $userinitStartEvent ) ## need userinitstartevent as it contains Winlogon PID which we used to get from logon event
            {
                $networkStartEvent = ($securityEvents|Where-Object { $_.Id -eq 4688 -and $_.Properties[$ProcessIdStart].value -eq $userinitStartEvent.Properties[$ProcessIdStart].value -and $_.properties[$NewProcessName].Value -eq 'C:\Windows\System32\mpnotify.exe' } | Select -Last 1 )
            }
            if( $networkStartEvent )
            {
                $Script:Output.Add( ( Get-PhaseEventFromCache -source 'Windows' -PhaseName 'Network Providers' `
                    -startEvent $networkStartEvent `
                    -endEvent ($securityEvents|Where-Object { $_.Id -eq 4689 -and $_.TimeCreated -ge $networkStartEvent.TimeCreated -and $_.Properties[$ProcessIdStop].value -eq $networkStartEvent.Properties[$ProcessIdNew].Value -and $_.properties[$ProcessName].Value -eq 'C:\Windows\System32\mpnotify.exe' } | Select -Last 1) ) )
            }
            else
            {
                [string]$warning = "Unable to find network providers start event"
                if( $auditingWarning )
                {
                    $warning += "`n$auditingWarning"
                    $auditingWarning = $null ## stop multiple occurrences
                }
                $script:warnings.Add( $warning )
            }
        }

        if ( $global:citrixUPMParams[ 'Path' ] -or ( Get-WinEvent -ListProvider 'Citrix Profile management' -ErrorAction SilentlyContinue)) {
            ($PowerShell = [PowerShell]::Create()).RunspacePool = $RunspacePool

            [scriptblock]$citrixScriptBlock = $null
            if( $global:citrixUPMParams[ 'Path' ] )
            {
                $citrixScriptBlock =
                {
                    Param( $logon , $username , $CitrixUPMEventFile , $UserProfileEventFile )
                    Get-PhaseEvent -PhaseName 'Citrix Profile Mgmt' -StartProvider 'Citrix Profile management' `
                        -StartEventFile $CitrixUPMEventFile `
                        -EndEventFile $UserProfileEventFile `
                        -EndProvider 'Microsoft-Windows-User Profiles Service' -StartXPath (
                        New-XPath -EventId 10 -From (Get-Date -Date $Logon.LogonTime) `
                         -EventData $UserName) -EndXPath (
                        New-XPath -EventId 1 -From (Get-Date -Date $Logon.LogonTime) `
                         -SecurityData @{
                             UserID=$Logon.UserSID
                    })
                }
            }
            else ## online
            {
                $citrixScriptBlock =
                {
                    Param( $logon , $username )
                    Get-PhaseEvent -PhaseName 'Citrix Profile Mgmt' -StartProvider 'Citrix Profile management' `
                        -EndProvider 'Microsoft-Windows-User Profiles Service' -StartXPath (
                        New-XPath -EventId 10 -From (Get-Date -Date $Logon.LogonTime) `
                            -EventData $UserName) -EndXPath (
                        New-XPath -EventId 1 -From (Get-Date -Date $Logon.LogonTime) `
                         -SecurityData @{
                             UserID=$Logon.UserSID
                    })
                }
            }
            [void]$PowerShell.AddScript( $citrixScriptBlock )
            [void]$PowerShell.AddParameters( $Parameters )
            [void]$jobs.Add( [pscustomobject]@{ 'PowerShell' = $PowerShell ; 'Handle' = $PowerShell.BeginInvoke() } )
        }
                
#region TTYE ActiveSetup time

         if ( $global:windowsShellCoreParams[ 'Path' ] -or ( Get-WinEvent -ListProvider 'Microsoft-Windows-Shell-Core' -ErrorAction SilentlyContinue)) {
            ($PowerShell = [PowerShell]::Create()).RunspacePool = $RunspacePool

            [scriptblock]$windowsShellCoreScriptBlock = $null
            if( $global:windowsShellCoreParams[ 'Path' ] )
            {
                $windowsShellCoreScriptBlock =
                {
                    Param( $logon , $username , $WindowsShellCoreFile )
                    Get-PhaseEvent -source 'Shell' -PhaseName 'ActiveSetup' -StartProvider 'Microsoft-Windows-Shell-Core' `
                        -StartEventFile $WindowsShellCoreFile `
                        -EndEventFile $WindowsShellCoreFile `
                        -EndProvider 'Microsoft-Windows-Shell-Core' -StartXPath (
                        New-XPath -EventId 62170 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                TaskName="ActiveSetup"
                                }) -EndXPath (
                        New-XPath -EventId 62171 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                TaskName="ActiveSetup"
                                })
                }
            }
            else ## online
            {
                $windowsShellCoreScriptBlock =
                {
                    Param( $logon )
                    Get-PhaseEvent -source 'Shell' -PhaseName 'ActiveSetup' -StartProvider 'Microsoft-Windows-Shell-Core' `
                        -EndProvider 'Microsoft-Windows-Shell-Core' -StartXPath (
                        New-XPath -EventId 62170 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                TaskName="ActiveSetup"
                                }) -EndXPath (
                        New-XPath -EventId 62171 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                TaskName="ActiveSetup"
                                })
                }
            }
            [void]$PowerShell.AddScript( $windowsShellCoreScriptBlock )
            [void]$PowerShell.AddParameters( $Parameters )
            [void]$jobs.Add( [pscustomobject]@{ 'PowerShell' = $PowerShell ; 'Handle' = $PowerShell.BeginInvoke() } )
        }
#endregion

#region TTYE AppVolumes ShellStart --> event can be tracked in the Winlogon log

         if ( $global:winlogonParams[ 'Path' ] -or ( Get-WinEvent -ListProvider 'Microsoft-Windows-Winlogon' -ErrorAction SilentlyContinue)) {
            ($PowerShell = [PowerShell]::Create()).RunspacePool = $RunspacePool

            [scriptblock]$winlogonScriptBlock = $null
            if( $global:winlogonParams[ 'Path' ] )
            {
                $winlogonScriptBlock =
                {
                    Param( $logon , $username , $WinlogonFile )
                    Get-PhaseEvent -source 'App Volumes' -PhaseName 'ShellStart' -StartProvider 'Microsoft-Windows-Winlogon' `
                        -StartEventFile $WinlogonFile `
                        -EndEventFile $WinlogonFile `
                        -EndProvider 'Microsoft-Windows-Winlogon' -StartXPath (
                        New-XPath -EventId 811 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                Event="12"
                                SubscriberName="svservice"
                                }) -EndXPath (
                        New-XPath -EventId 812 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                Event="12"
                                SubscriberName="svservice"
                                })
                }
            }
            else ## online
            {
                $winlogonScriptBlock =
                {
                    Param( $logon )
                    Get-PhaseEvent -source 'App Volumes' -PhaseName 'ShellStart' -StartProvider 'Microsoft-Windows-Winlogon' `
                        -EndProvider 'Microsoft-Windows-Winlogon' -StartXPath (
                        New-XPath -EventId 811 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                Event="12"
                                SubscriberName="svservice"
                                }) -EndXPath (
                        New-XPath -EventId 812 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                Event="12"
                                SubscriberName="svservice"
                                })
                }
            }
            [void]$PowerShell.AddScript( $winlogonScriptBlock )
            [void]$PowerShell.AddParameters( $Parameters )
            [void]$jobs.Add( [pscustomobject]@{ 'PowerShell' = $PowerShell ; 'Handle' = $PowerShell.BeginInvoke() } )
        }
#endregion

#region TTYE FSLogix ShellStart --> event can be tracked in the Winlogon log

         if ( $global:winlogonParams[ 'Path' ] -or ( Get-WinEvent -ListProvider 'Microsoft-Windows-Winlogon' -ErrorAction SilentlyContinue)) {
            ($PowerShell = [PowerShell]::Create()).RunspacePool = $RunspacePool

            [scriptblock]$winlogonScriptBlock = $null
            if( $global:winlogonParams[ 'Path' ] )
            {
                $winlogonScriptBlock =
                {
                    Param( $logon , $username , $WinlogonFile )
                    Get-PhaseEvent -source 'FSLogix' -PhaseName 'ShellStart' -StartProvider 'Microsoft-Windows-Winlogon' `
                        -StartEventFile $WinlogonFile `
                        -EndEventFile $WinlogonFile `
                        -EndProvider 'Microsoft-Windows-Winlogon' -StartXPath (
                        New-XPath -EventId 811 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                Event="12"
                                SubscriberName="frxsvc"
                                }) -EndXPath (
                        New-XPath -EventId 812 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                Event="12"
                                SubscriberName="frxsvc"
                                })
                }
            }
            else ## online
            {
                $winlogonScriptBlock =
                {
                    Param( $logon )
                    Get-PhaseEvent -source 'FSLogix' -PhaseName 'ShellStart' -StartProvider 'Microsoft-Windows-Winlogon' `
                        -EndProvider 'Microsoft-Windows-Winlogon' -StartXPath (
                        New-XPath -EventId 811 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                Event="12"
                                SubscriberName="frxsvc"
                                }) -EndXPath (
                        New-XPath -EventId 812 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                Event="12"
                                SubscriberName="frxsvc"
                                })
                }
            }
            [void]$PowerShell.AddScript( $winlogonScriptBlock )
            [void]$PowerShell.AddParameters( $Parameters )
            [void]$jobs.Add( [pscustomobject]@{ 'PowerShell' = $PowerShell ; 'Handle' = $PowerShell.BeginInvoke() } )
        }
#endregion

        ##TTYE AppX file association load time

         if ( $global:windowsShellCoreParams[ 'Path' ] -or ( Get-WinEvent -ListProvider 'Microsoft-Windows-Shell-Core' -ErrorAction SilentlyContinue)) {
            ($PowerShell = [PowerShell]::Create()).RunspacePool = $RunspacePool

            [scriptblock]$windowsShellCoreScriptBlock = $null
            if( $global:windowsShellCoreParams[ 'Path' ] )
            {
                $windowsShellCoreScriptBlock =
                {
                    Param( $logon , $username , $WindowsShellCoreFile )
                    Get-PhaseEvent -source 'Shell' -PhaseName 'AppX File Associations' -StartProvider 'Microsoft-Windows-Shell-Core' `
                        -StartEventFile $WindowsShellCoreFile `
                        -EndEventFile $WindowsShellCoreFile `
                        -EndProvider 'Microsoft-Windows-Shell-Core' -StartXPath (
                        New-XPath -EventId 62443 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                Info="AppDefaults-Logon-UserProfileCreated"
                                }) -EndXPath (
                        New-XPath -EventId 62443 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                Info="AppDefaults-Logon-UserProfileLoaded"
                                })
                }
            }
            else ## online
            {
                $windowsShellCoreScriptBlock =
                {
                    Param( $logon )
                    Get-PhaseEvent -source 'Shell' -PhaseName 'AppX File Associations' -StartProvider 'Microsoft-Windows-Shell-Core' `
                        -EndProvider 'Microsoft-Windows-Shell-Core' -StartXPath (
                        New-XPath -EventId 62443 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                Info="AppDefaults-Logon-UserProfileCreated"
                                }) -EndXPath (
                        New-XPath -EventId 62443 -From (Get-Date -Date $Logon.LogonTime) `
                            -SecurityData @{
                                UserID=$Logon.UserSID
                            } -EventData @{
                                Info="AppDefaults-Logon-UserProfileLoaded"
                                })
                }
            }
            [void]$PowerShell.AddScript( $windowsShellCoreScriptBlock )
            [void]$PowerShell.AddParameters( $Parameters )
            [void]$jobs.Add( [pscustomobject]@{ 'PowerShell' = $PowerShell ; 'Handle' = $PowerShell.BeginInvoke() } )
        }

        ##TTYE AppX application load time
        if ( $global:appReadinessParams[ 'Path' ] -or ( Get-WinEvent -ListProvider 'Microsoft-Windows-AppReadiness' -ErrorAction SilentlyContinue)) {
            ($PowerShell = [PowerShell]::Create()).RunspacePool = $RunspacePool

            [scriptblock]$appReadinessCoreScriptBlock = $null
            if( $global:appReadinessParams[ 'Path' ] )
            {
                $appReadinessCoreScriptBlock =
                {
                    Param( $logon , $username , $appReadinessFile )
                    Get-PhaseEvent -source 'Shell' -PhaseName 'AppX - Load Packages' -StartProvider 'Microsoft-Windows-AppReadiness' `
                        -StartEventFile $appReadinessFile `
                        -EndEventFile $appReadinessFile `
                        -EndProvider 'Microsoft-Windows-AppReadiness' -StartXPath (
                        New-XPath -EventId 209 -From (Get-Date -Date $Logon.LogonTime) `
                            -EventData @{
                                User=$Logon.UserSID
                                From=2
                                To=0
                                }) -EndXPath (
                        New-XPath -EventId 209 -From (Get-Date -Date $Logon.LogonTime) `
                            -EventData @{
                                User=$Logon.UserSID
                                From=1
                                To=2
                                })
                }
            }
            else ## online
            {
                $appReadinessCoreScriptBlock =
                {
                    Param( $logon )
                    Get-PhaseEvent -source 'Shell' -PhaseName 'AppX - Load Packages' -StartProvider 'Microsoft-Windows-AppReadiness' `
                        -EndProvider 'Microsoft-Windows-AppReadiness' -StartXPath (
                        New-XPath -EventId 209 -From (Get-Date -Date $Logon.LogonTime) `
                            -EventData @{
                                User=$Logon.UserSID
                                From=2
                                To=0
                                }) -EndXPath (
                        New-XPath -EventId 209 -From (Get-Date -Date $Logon.LogonTime) `
                            -EventData @{
                                User=$Logon.UserSID
                                From=1
                                To=2
                                })
                }
            }
            [void]$PowerShell.AddScript( $appReadinessCoreScriptBlock )
            [void]$PowerShell.AddParameters( $Parameters )
            [void]$jobs.Add( [pscustomobject]@{ 'PowerShell' = $PowerShell ; 'Handle' = $PowerShell.BeginInvoke() } )
        }

        ($PowerShell = [PowerShell]::Create()).RunspacePool = $RunspacePool

        [scriptblock]$scriptBlock = $null
        if( $global:userProfileParams[ 'Path' ] )
        {
            $scriptBlock = `
            {
                Param( $logon , $UserProfileEventFile )
                Get-PhaseEvent -PhaseName 'User Profile' `
                    -StartEventFile $UserProfileEventFile `
                    -EndEventFile $UserProfileEventFile `
                    -eventLog 'Microsoft-Windows-User Profile Service/Operational' `
                    -StartProvider 'Microsoft-Windows-User Profiles Service' `
                    -EndProvider 'Microsoft-Windows-User Profiles Service' `
                    -StartXPath (New-XPath -EventId 1 -From (Get-Date -Date $Logon.LogonTime) `
                    -SecurityData @{UserID=$Logon.UserSID}) `
                    -EndXPath (New-XPath -EventId 2 -From (Get-Date -Date $Logon.LogonTime) `
                    -SecurityData @{
                        UserID=$Logon.UserSID
                    })
            }
        }
        else ## online
        {
            $scriptBlock = `
            {
                Param( $logon )
                Get-PhaseEvent -PhaseName 'User Profile' `
                    -eventLog 'Microsoft-Windows-User Profile Service/Operational' `
                    -StartProvider 'Microsoft-Windows-User Profiles Service' `
                    -EndProvider 'Microsoft-Windows-User Profiles Service' `
                    -StartXPath (New-XPath -EventId 1 -From (Get-Date -Date $Logon.LogonTime) `
                    -SecurityData @{UserID=$Logon.UserSID}) `
                    -EndXPath (New-XPath -EventId 2 -From (Get-Date -Date $Logon.LogonTime) `
                    -SecurityData @{
                        UserID=$Logon.UserSID
                    })
            }
        }
        [void]$PowerShell.AddScript( $scriptBlock )
        [void]$PowerShell.AddParameters( $Parameters )
        [void]$jobs.Add( [pscustomobject]@{ 'PowerShell' = $PowerShell ; 'Handle' = $PowerShell.BeginInvoke() } )
        
        ($PowerShell = [PowerShell]::Create()).RunspacePool = $RunspacePool

        [scriptblock]$groupPolicyScriptBlock = $null
        if( $global:groupPolicyParams[ 'Path' ] )
        {
            $groupPolicyScriptBlock = {
                Param( $logon , $Username , $UserDomain , $groupPolicyEventFile )
                Get-PhaseEvent -PhaseName 'Group Policy' `
                    -StartEventFile $groupPolicyEventFile `
                    -EndEventFile $groupPolicyEventFile `
                    -eventLog 'Microsoft-Windows-GroupPolicy/Operational' `
                    -StartProvider 'Microsoft-Windows-GroupPolicy' `
                    -EndProvider 'Microsoft-Windows-GroupPolicy' `
                    -StartXPath (
                    New-XPath -EventId 4001 -From (Get-Date -Date $Logon.LogonTime) `
                    -EventData @{
                        PrincipalSamName="$UserDomain\$UserName"
                    }) -EndXPath (
                    New-XPath -EventId 8001 -From (Get-Date -Date $Logon.LogonTime) `
                    -EventData @{
                        PrincipalSamName="$UserDomain\$UserName"
                    })
             }
        }
        else
        {
            $groupPolicyScriptBlock = {
                Param( $logon , $Username , $UserDomain )
                Get-PhaseEvent -PhaseName 'Group Policy' `
                    -eventLog 'Microsoft-Windows-GroupPolicy/Operational' `
                    -StartProvider 'Microsoft-Windows-GroupPolicy' `
                    -EndProvider 'Microsoft-Windows-GroupPolicy' `
                    -StartXPath (
                    New-XPath -EventId 4001 -From (Get-Date -Date $Logon.LogonTime) `
                    -EventData @{
                        PrincipalSamName="$UserDomain\$UserName"
                    }) -EndXPath (
                    New-XPath -EventId 8001 -From (Get-Date -Date $Logon.LogonTime) `
                    -EventData @{
                        PrincipalSamName="$UserDomain\$UserName"
                    })
            }
        }
         
        [void]$PowerShell.AddScript( $groupPolicyScriptBlock )
        [void]$PowerShell.AddParameters( $Parameters )
        [void]$jobs.Add( [pscustomobject]@{ 'PowerShell' = $PowerShell ; 'Handle' = $PowerShell.BeginInvoke() } )
        
        ($PowerShell = [PowerShell]::Create()).RunspacePool = $RunspacePool

        [scriptblock]$gpScriptBlock = $null
        if( $global:groupPolicyParams[ 'Path' ] )
        {
            $gpScriptBlock = 
            {
                Param( $logon , $UserDomain , $Username , $sharedVars , $groupPolicyEventFile )
                Get-PhaseEvent -PhaseName 'GP Scripts' -StartProvider 'Microsoft-Windows-GroupPolicy' -SharedVars $sharedVars `
                    -StartEventFile $groupPolicyEventFile `
                    -EndEventFile $groupPolicyEventFile `
                    -EndProvider 'Microsoft-Windows-GroupPolicy' `
                    -StartXPath (
                    New-XPath -EventId 4018 -From (Get-Date -Date $Logon.LogonTime) `
                    -EventData @{PrincipalSamName="$UserDomain\$UserName";ScriptType=1}) `
                    -EndXPath (
                    New-XPath -EventId 5018 -From (Get-Date -Date $Logon.LogonTime) `
                    -EventData @{
                        PrincipalSamName="$UserDomain\$UserName"
                        ScriptType=1
                    })
             }
        }
        else
        {
            $gpScriptBlock = 
            {
                Param( $logon , $UserDomain , $Username , $sharedVars )
                Get-PhaseEvent -PhaseName 'GP Scripts' -StartProvider 'Microsoft-Windows-GroupPolicy' -SharedVars $sharedVars `
                    -EndProvider 'Microsoft-Windows-GroupPolicy' `
                    -StartXPath (
                    New-XPath -EventId 4018 -From (Get-Date -Date $Logon.LogonTime) `
                    -EventData @{PrincipalSamName="$UserDomain\$UserName";ScriptType=1}) `
                    -EndXPath (
                    New-XPath -EventId 5018 -From (Get-Date -Date $Logon.LogonTime) `
                    -EventData @{
                        PrincipalSamName="$UserDomain\$UserName"
                        ScriptType=1
                    })
             }
        }
        [void]$PowerShell.AddScript( $gpScriptBlock )    
        [void]$PowerShell.AddParameters( $Parameters )
        [void]$jobs.Add( [pscustomobject]@{ 'PowerShell' = $PowerShell ; 'Handle' = $PowerShell.BeginInvoke() } )
        
        ($PowerShell = [PowerShell]::Create()).RunspacePool = $RunspacePool

        if( $userinitStartEvent )
        {
            $endevent = $null
            if( $isPublishedApp )
            {
                [string]$shell = $shellProgram
                if( [string]::IsNullOrEmpty( $shell ) )
                {
                    $shell = Join-Path -Path $env:SystemRoot -ChildPath 'icast.exe'
                }
                ## we already have process end of this but not process start
                $endevent = ($securityEvents|Where-Object { $_.Id -eq 4688 -and $_.TimeCreated -ge $logon.LogonTime -and $_.Properties[$SubjectLogonId].value -in $Logon.LogonID -and $_.Properties[$NewProcessName].value -eq $shell } | Select -Last 1)
            }
            else
            {
                $endevent = $logonFinishedEvent ## this is explorer starting
            }
            if( $endEvent )
            {
                $Script:Output.Add( ( Get-PhaseEventFromCache -source 'Windows' -PhaseName 'Pre-Shell (Userinit)' -startEvent $userinitStartEvent -endEvent $endEvent ) )
            }
            else
            {
                Write-Debug "Unable to find userinit end event"
            }
        }
        else
        {
            [string]$info = "Unable to find Pre-Shell (Userinit) start event"
            if( $auditingWarning )
            {
                $info += "`n$auditingWarning"
                $auditingWarning = $null ## stop multiple occurrences
            }
            $script:warnings.Add( $info )
        }

        ## See if user has a login script in AD and if so look for start and end in process start/stop events
        $ADuser = ([ADSI]"WinNT://$UserDomain/$Username,user")
        if( $ADUser -and $ADuser.LoginScript )
        {
            if( $searchCommandLine -or $offline )
            {
                ## could be more than one since usrlogon.cmd may also be launched so need to check we have the right one although not checking down to which server as can't. Don't check for actual process as could be cmd, wscript, etc
                ## can't check for parent of userinit.exe as that doesn't exist on Win7/2008R2 but we could check for its PID as parent if we have $userinitStartEvent
                [string]$escapedLogonScript = [regex]::Escape( ( Join-Path -Path '\NETLOGON' -ChildPath ($ADuser.LoginScript.ToString()) ) )
                $logonScriptStartEvent = ($securityEvents|Where-Object { $_.Id -eq 4688 -and $_.Properties[$SubjectUserName].value -eq $userName -and $_.Properties[$SubjectDomainName].value -eq $UserDomain `
                     -and $_.Properties[$SubjectLogonId].value -in $Logon.LogonId -and $_.Properties[$CommandLine].value -match "[^\\\""]$($escapedLogonScript)[^a-z0-9_]" } ) | Select -Last 1
                if( $logonScriptStartEvent )
                {
                    $Script:Output.Add( ( Get-PhaseEventFromCache -source 'Windows' -PhaseName 'User logon script' `
                        -startEvent $logonScriptStartEvent `
                        -endEvent ($securityEvents|Where-Object { $_.Id -eq 4689 -and $_.TimeCreated -ge $logonScriptStartEvent.TimeCreated -and $_.Properties[$ProcessIdStop].value -eq $logonScriptStartEvent.Properties[$ProcessIdNew].value -and $_.Properties[$SubjectLogonId].value -eq $logonScriptStartEvent.Properties[$SubjectLogonId].value } | Select -Last 1) ) )
                }
            }
            else
            {
                $logonScriptStartEvent = $null
            }

            if( ! $logonScriptStartEvent )
            {
                [string]$warning = "Unable to find user logon script ($($ADUser.LoginScript)) start event"
                if( $auditingWarning )
                {
                    $warning += "`n$auditingWarning"
                    $auditingWarning = $null ## stop multiple occurrences
                }
                if( $commandLinePolicy -and $commandLinePolicy.ProcessCreationIncludeCmdLine_Enabled -ne 1 )
                {
                    $warning += ', "Command line process auditing" is not enabled'
                }
                $script:warnings.Add( $warning )
            }
        }

        if ($CUDesktopLoadTime -gt 0 ) {
            $shellStartEvent = ($securityEvents|Where-Object { $_.Id -eq 4688 -and $_.Properties[$SubjectLogonId].value -in $Logon.LogonID -and $_.properties[$NewProcessName].Value -eq 'C:\Windows\explorer.exe' } | Select -Last 1 ) 
            if( $shellStartEvent )
            {
                $Script:Output.Add( ( Get-PhaseEventFromCache -source 'Windows' -PhaseName 'Shell' -startEvent $shellStartEvent -CUAddition $CUDesktopLoadTime ) )
            }
            else
            {
                [string]$warning = "Unable to find Shell start event"
                if( $auditingWarning )
                {
                    $warning += "`n$auditingWarning"
                    $auditingWarning = $null ## stop multiple occurrences
                }
                $script:warnings.Add( $warning )
            }
        }
        
        $jobs | ForEach-Object `
        {
            $_.powershell.EndInvoke( $_.handle ) | ForEach-Object `
            {
                $script:output.Add( $_ )
            }
            $_.PowerShell.Dispose()
        }
        $jobs.clear()

        $Script:GPAsync = $sharedVars[ 'GPASync' ]
        if( $userinitStartEvent )
        {
            $end = ($Script:Output | Where-Object PhaseName -eq 'Pre-Shell (Userinit)' ) | Select-Object -ExpandProperty EndTime
            Write-Debug "Get-PrinterEvents -Start $($userinitStartEvent.TimeCreated) -End $end -ClientName $ClientName"
            if( $end )
            {
                Get-PrinterEvents -Start $userinitStartEvent.TimeCreated -End $end -ClientName $ClientName
            }
        }
        
        if( $userProfileEndTime = $Script:Output | Where-Object PhaseName -eq 'User Profile' | Select-Object -First 1 -ExpandProperty StartTime -ErrorAction SilentlyContinue )
        {
            Get-FSLogixProfileEvents -Username $Username -Start $Logon.LogonTime -End $userProfileEndTime
        }
        else
        {
            $warnings.Add( "Unable to find user profile stage" )
        }

        if( (Get-Variable -Name end -ErrorAction SilentlyContinue ) -and ( Get-ItemProperty 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -Name DisplayName -ErrorAction SilentlyContinue|Where DisplayName -match 'App Volumes Agent' ) `
            -or ( $offline -and (Test-Path -Path $appVolumesLogFile -ErrorAction SilentlyContinue)))
        {
            Get-AppVolumeEvents -Start $Logon.LogonTime -End $end.AddSeconds( 120 )
        }

        if ( $Script:Output.Count -lt 2 ) {
            $PSCmdlet.WriteWarning("Not enough data for that session, Aborting function...")
            Throw 'Could not find more than a single phase, script is aborted'
        }
    }
    end {
        $LogonTimeReal = $Logon.FormatTime
        [System.Collections.Generic.List[psobject]]$Script:Output = $Script:Output | Sort-Object -Property StartTime
        $TotalDur = 'N/A'
        if ( $Script:LogonStartDate) { ## Not set any more, used to be via OData function
            $Script:LogonStartDate = $Script:LogonStartDate.ToLocalTime()
            ForEach( $phase in $script:output ) {
                if ($phase.PhaseName -eq 'Shell' -or $phase.PhaseName -eq 'Pre-Shell (Userinit)' ) {
                    [decimal]$thisDuration = New-TimeSpan -Start $Script:LogonStartDate -End $Script:Output[-1].EndTime | Select-Object -ExpandProperty TotalSeconds
                    if( $TotalDur -eq 'N/A' -or $TotalDur -as [decimal] -lt $thisDuration ) {
                        $TotalDur = $thisDuration
                    }
                }
            }
            $Deltas = New-TimeSpan -Start $Script:LogonStartDate -End $Script:Output[0].StartTime
            $Script:Output[0] | Add-Member -MemberType NoteProperty -Name TimeDelta -Value $Deltas -Force
            $LogonTimeReal =  (Get-Date -Date $Script:LogonStartDate).ToString( 'HH:mm:ss.ff' )
        }
        else {
            $TotalDur = 'N/A'
            ForEach( $phase in $script:output ) {
                if ($phase.PhaseName -eq 'Shell' -or $phase.PhaseName -eq 'Pre-Shell (Userinit)' ) {
                    [decimal]$thisDuration = New-TimeSpan -Start $Logon.LogonTime -End $phase.EndTime | Select-Object -ExpandProperty TotalSeconds
                    if( $TotalDur -eq 'N/A' -or $TotalDur -as [decimal] -lt $thisDuration ) {
                        $TotalDur = $thisDuration
                    }
                }
            }
            $Deltas = New-TimeSpan -Start $Logon.LogonTime -End $Script:Output[0].StartTime
            $Script:Output[0] | Add-Member -MemberType NoteProperty -Name TimeDelta -Value $Deltas -Force
        }
        
#region Ivanti EM
        if( ($emservice = Get-Service -name 'AppSense EmCoreService' -ErrorAction SilentlyContinue ) )
        {
            # 9659 is for personalisation success
            # 9661 is personalisation server problem
            # 9662 is a trigger summary
            if( ( [array]$appSenseEvents = @( Get-WinEvent -Oldest -FilterHashtable @{ StartTime = $logon.LogonTime ; UserID = $logon.UserSid ; Id = 9662 , 9659 , 9661 ; ProviderName = 'AppSense Environment Manager.' } -ErrorAction SilentlyContinue ).Where( 
                { ($_.Id -eq 9662 -and $_.Properties[4].Value -match "SessionID:$sessionID`$") -or ($_.Id -eq 9659 -and $_.Properties[1].Value -match "SessionID:$sessionID`$") -or ($_.Id -eq 9661 -and $_.Properties[0].Value -match "SessionID:$sessionID`$")} )) -and $appsenseEvents.Count )
            {
                ## Times are in UTC so convert to local time - https://devblogs.microsoft.com/scripting/powertip-convert-from-utc-to-my-local-time-zone/
                $currentTimeZone = Get-CimInstance -ClassName Win32_TimeZone | Select-Object -ExpandProperty StandardName
                $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById( $currentTimeZone )
                $emuserProcess = $null
                [bool]$foundPSError = $false
                [bool]$foundPSGood = $false

                ForEach( $appsenseEvent in $appSenseEvents )
                {
                    if( $appsenseEvent.Id -eq 9659 ) ## User personalization settings for Dsktp updated from personalization server.
                    {
                        if( $appsenseEvent.Properties[0].Value -eq 'Dsktp' -and ! $emuserProcess )
                        {
                            ## The emuser is launched in SubjectLogonId 999 as system so we have to check that it is the right one (e.g. two overlapping logons) so look for current running process rather than in event logs - restarted process will confused things though
                            ## could also check that parent is emcoreservice.exe
                            if( $emuserProcess = Get-Process -Name EmUser -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $SessionId -and $_.StartTime -ge $logon.LogonTime  -and $_.StartTime -le $appsenseEvent.TimeCreated -and $_.Path -match '\\Environment Manager\\Agent\\EmUser\.exe$' }  | Sort-Object -Property StartTime | Select-Object -First 1 )
                            {
                                $Script:Output.Add( ( [pscustomobject]@{ 
                                    'Source' = 'Ivanti EM'
                                    'PhaseName' = 'Personalization Loading'
                                    'StartTime' = $emuserProcess.StartTime
                                    'EndTime'   = $appsenseEvent.TimeCreated
                                    'Duration'  = ($appsenseEvent.TimeCreated - $emuserProcess.StartTime).TotalSeconds }))
                                $foundPSGood = $true
                            }
                            else
                            {
                                $warnings.Add( "Unable to find running Ivanti EM emuser.exe process for this session so cannot determine personalisation load time" )
                            }
                        }
                        else ## ignore non Dsktp phase or if we have already had it since starting at oldest event
                        {
                            Write-Debug -Message "Discarding Ivanti event $($appsenseEvent.Id) from $(Get-Date -Date $appsenseEvent.TimeCreated) `"$($appsenseEvent.Message)`""
                        }
                    }
                    elseif( $appsenseEvent.Id -eq 9661 )
                    {
                        if( ! $foundPSError )
                        {
                            $warnings.Add( "Ivanti error: $($appsenseEvent.Message)" )
                            $foundPSError = $true
                        }
                    }
                    else
                    {
                         $emphase = [pscustomobject]@{ 
                            'Source' = 'Ivanti EM'
                            'PhaseName' = (Get-Culture).TextInfo.ToTitleCase( ( $appsenseEvent.Properties[0].Value -replace '_' , ' ' ).ToLower()) ## LOGON_PRE_DESKTOP
                            'StartTime' = [System.TimeZoneInfo]::ConvertTimeFromUtc( [datetime]$appsenseEvent.Properties[1].Value , $TZ )
                            'EndTime'   = [System.TimeZoneInfo]::ConvertTimeFromUtc( [datetime]$appsenseEvent.Properties[2].Value , $TZ )
                            'Duration'  = [int]$appsenseEvent.Properties[3].Value / 1000 }
                        if( $emphase.Phasename -match 'Desktop Created' )
                        {
                            $script:ivantiEMNonBlockingPhases.Add( $emphase )
                        }
                        else
                        {
                            $Script:Output.Add( $emphase )
                        }
                    }
                }
                if( ! $foundPSGood )
                {
                    [string]$message = "Found no evidence of Ivanti personalisation for this session but it may not be enabled or configured for this user" 
                    if( $ivantiPSservers = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\AppSense\Environment Manager\Personalization' -Name 'ServerList' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'ServerList' )
                    {
                        $message += " (found $ivantiPSservers in HKLM policies key)"
                    }
                    $warnings.Add( $message )
                }
            }
            else ## we could look to see what auditing is enable in the XML config - need to check for value putting config in non-default location
            {
                [string]$status = $(if( $emservice.Status -ne 'Running' ) { 'not ' })
                ## see if we have a config file
                if( [string]::IsNullOrEmpty( ( [string]$configPath = Get-ItemProperty -Path "HKLM:\SOFTWARE\AppSense Technologies\Communications Agent" -Name 'native config path' -ErrorAction SilentlyContinue | Select -ExpandProperty 'native config path' ) ) )
                {
                    $configPath = Join-Path -Path ([Environment]::GetFolderPath( [System.Environment+SpecialFolder]::CommonApplicationData )) -ChildPath 'AppSense'
                }
                [string]$emConfigFile = [System.IO.Path]::Combine( $configPath , 'Environment Manager' , 'configuration.aemp' )
                if( ! ( Test-Path -Path $emConfigFile -PathType Leaf -ErrorAction SilentlyContinue ) )
                {
                    $warnings.Add( "No Ivanti EM configuration file found at `"$emConfigFile`"" )
                }
                $warnings.Add( "Ivanti EM service present and $($status)running but found no relevant local events - are event ids 9662 & 9659 enabled in the configuration?" )
            }
        }
#endregion Ivanti EM

        $LogonTaskList = Get-LogonTask -UserName $Username -UserDomain $UserDomain -Start $Logon.LogonTime -End $Script:Output[-1].EndTime

        $outputObject = [pscustomobject][ordered]@{ 'User name ' = $username }

        if( $odataPhase -and $odataPhase.PSObject.Properties )
        {   
            ## odataphase no longer built from OData so just copy
            ForEach( $property in ( $odataPhase.PSObject.Properties | Sort-Object -Property Name ))
            {
                Add-Member -InputObject $outputObject -MemberType NoteProperty -Name $property.Name -Value $property.Value
            }
        }

        ($outputObject | Format-List | Out-String).Trim()
        ''
        
        $earliest = $null
        $latest = $null
        [double]$totalDuration = 0 
        [double]$duration = 0
        [string]$indent = ''
        [string]$prelogonVendor = 'VMware'

        if( $prelogonData -and $prelogonData.Count )
        {
            ForEach( $item in $prelogonData )
            {
                if( ! $earliest -or $item.StartTime -lt $earliest )
                {
                    $earliest = $item.StartTime
                }
                if( ! $latest -or $item.EndTime -gt $latest )
                {
                    $latest = $item.EndTime
                }
            }

            $duration = ($latest - $earliest).TotalSeconds

            ## No start/end for this as a total of the phase and sorting on start puts it at the end where we want it
            $prelogonData.Add( ( [pscustomobject]@{ 'PhaseName' = 'Pre-Windows Duration' ; Duration = $duration } ) )

            ## calculate delay between latest action finish and Windows logon commencing
            [double]$phaseDelay = 0
            [string]$delayBetweenPhases = $null
            if( $latest )
            {
                $phaseDelay = [math]::Round( ($Logon.LogonTime - $latest).TotalSeconds, 1)
                if( $phaseDelay -lt 0 )
                {
                    $phaseDelay = 0
                }
                $delayBetweenPhases = "Delay between $prelogonVendor and Windows phases: $phaseDelay seconds`n"
            }
        }

        $totalDuration = $duration

        $earliestOverall = $earliest
        $latestOverall = $latest
        $earliest = $null
        $latest = $null

        if( $Script:Output -and $Script:Output.Count )
        {
            ForEach( $item in $Script:Output )
            {
                if( ! $earliest -or $item.StartTime -lt $earliest )
                {
                    $earliest = $item.StartTime
                }
                if( ! $latest -or $item.EndTime -gt $latest )
                {
                    $latest = $item.EndTime
                }
            }
            $duration = ($latest - $earliest).TotalSeconds
        }

        $totalDuration += $duration

        [datetime]$start = $(if( ! $earliestOverall -or $earliest -lt $earliestOverall ) { $earliest } else { $earliestOverall })
        [datetime]$end = $(if( ! $latestOverall -or $latest -gt $latestOverall ) { $latest } else { $latestOverall })

        ([pscustomobject]@{
            'Logon start' = '{0} {1}' -f (Get-Date -Date $start -Format d), (Get-Date -Date $start -Format 'HH:mm:ss' )
            'Logon end'   = '{0} {1}' -f (Get-Date -Date $end -Format d), (Get-Date -Date $end -Format 'HH:mm:ss')
            'Duration'    = "$([math]::Round( ($end - $start).TotalSeconds , 1 )) seconds" } | Format-List | Out-String).Trim()
        ''
        
        $Script:Output.Add( ( [pscustomobject]@{ 'Source' = 'Windows' ; 'PhaseName' = 'Windows Logon Time' ; 'StartTime' = $logon.LogonTime ; 'EndTime' = $logon.LogonTime ; 'Duration' = 0.0 } ) )
        $Script:Output.Add( ( [pscustomobject]@{ 'PhaseName' = 'Windows Duration' ; Duration = $duration } ) )

        ## find the longest source and phasenames so we can pad them all to the same width so the two tables have the same dimensions
        [int]$longestSource = 0
        [int]$longestPhasename = 0
        
        ForEach( $horizonItem in $prelogonData )
        {
            if( $horizonItem.PSObject.Properties[ 'Source' ] -and $horizonItem.Source.Length -gt $longestSource )
            {
                $longestSource = $horizonItem.Source.Length
            }
            if( $horizonItem.PhaseName.Length -gt $longestPhasename )
            {
                $longestPhasename = $horizonItem.PhaseName.Length
            }
        }
        
        ForEach( $outputItem in $Script:Output )
        {
            if( $outputItem.PSObject.Properties[ 'Source' ] -and $outputItem.Source.Length -gt $longestSource )
            {
                $longestSource = $outputItem.Source.Length
            }
            if( $outputItem.PhaseName.Length -gt $longestPhasename )
            {
                $longestPhasename = $outputItem.PhaseName.Length
            }
        }

        ## have to do it this way as PS v4 errors if you declare and assign in the same statement
        $format = New-Object System.Collections.Generic.List[psobject]]
        
        $format.Add( (@{Expression={$_.Source.PadRight($longestSource,' ')};Label="Source"} ) )
        $format.Add( (@{Expression={$_.PhaseName.PadRight($longestPhasename,' ')};Label="Phase"} ) )
        $format.Add( (@{Expression={'{0:N1}' -f $_.Duration};Label="Duration (s)"} ) )
        $format.Add( (@{Expression={'{0:HH:mm:ss.f}' -f $_.StartTime};Label="Start Time"} ) )
        $format.Add( ( @{Expression={'{0:HH:mm:ss.f}' -f $_.EndTime};Label="End Time"} ) )
                 
        if( $prelogonData -and $prelogonData.Count )
        {
            ($prelogonData | Sort-Object -Property 'StartTime' | Format-Table -Property $Format -AutoSize | Out-String).Trim() -split "`r`n" | ForEach-Object { "$indent$_" }
            ''
            ## $delayBetweenPhases
        }
        
        ## sort on start time so we can calculate gap with previous phase
        $Script:Output = $Script:Output | Sort-Object -Property StartTime

        ## Calculate gaps between sorted phases now that we have all components and are sorted in ascending start order . This assumes they are in some way synchronous
        for( $i=1 ; $i -le $Script:Output.Count - 1 ; $i++ ) {
            if( $Script:Output[$i].PSObject.Properties[ 'StartTime' ] ) {
                if( ( $Deltas = New-TimeSpan -Start $Script:Output[$i-1].EndTime -End $Script:Output[$i].StartTime -ErrorAction SilentlyContinue ) -lt 0 ) {
                    #if tasks are run asynchronously, then deltas may not be timed correctly.  Setting the value as blank to avoid confusion.
                    $Deltas = ""
                }
                Add-Member -InputObject $Script:Output[$i] -MemberType NoteProperty -Name TimeDelta -Value $Deltas -Force
            }
        }

        $Format.Add(  @{Expression={'{0:N1}' -f ($_.TimeDelta | Select-Object -ExpandProperty TotalSeconds)};Label="Gap (s)"} )

        ( $Script:Output | Format-Table -Property $Format -AutoSize | Out-String).Trim() -split "`r`n" | ForEach-Object { "$indent$_" }
        
        ''
        'Non blocking logon tasks'
        '------------------------'

        if ($Script:GPAsync)
        {
            "`nGroup Policy asynchronous scripts were processed for $Script:GPAsync seconds"
        }

        $LogonTaskList | Format-Table @{Expression={$_.TaskName};Label="Logon Scheduled Task"},@{Expression={'{0:s\.ff}' -f $_.Duration};Label="Duration (s)"},@{Expression={$_.ActionName};Label="Action Name"} -AutoSize
        
        $format.RemoveAt( $format.Count - 1 ) ## Remove "Gap (s)" as not relevant now
        
        if( $script:vmwareDEMNonBlockingPhases -and $script:vmwareDEMNonBlockingPhases.Count )
        {
            ##"$productName Phases"
            ( $script:vmwareDEMNonBlockingPhases | Sort-Object -Property StartTime | Format-Table -Property $Format -AutoSize | Out-String).Trim() -split "`r`n" | ForEach-Object { "$indent$_" }
            ''
        }

        if( $Script:AppVolumesOutput -and $Script:AppVolumesOutput.Count )
        {
            'App Volumes Phase'
            ''
            ( $Script:AppVolumesOutput | Sort-Object -Property StartTime | Format-Table -Property $Format -AutoSize | Out-String).Trim() -split "`r`n" | ForEach-Object { "$indent$_" }
            ''
        }
        
        if( $script:ivantiEMNonBlockingPhases -and $script:ivantiEMNonBlockingPhases.Count )
        {
            ''
            ( $script:ivantiEMNonBlockingPhases | Sort-Object -Property StartTime | Format-Table -Property $Format -AutoSize | Out-String).Trim() -split "`r`n" | ForEach-Object { "$indent$_" }
            ''
        }

        if( $CSEArray -and $CSEArray.Count )
        {         
             $Format.Add( @{Expression={ $_.GPOs };Label="GPO(s)"} )
            'Group Policy Client Side Extension Processing'
            ''
            $lastToFinish = $null
            [hashtable]$GPOTotalTimes = @{}

            [array]$CSEtimings = @( $CSEArray.Where( { $_.Id -ne '4016' } ).ForEach( 
            {
                $CSE = $_
                [double]$duration = $CSE.Properties[0].Value / 1000
                if( ! $lastToFinish -or $CSE.TimeCreated -gt $lastToFinish )
                {
                    $lastToFinish = $CSE.TimeCreated
                }

                ## look up the list of GPOs via the CSE extension id from 4016 event we built earlier
                [string[]]$GPOs = @( $CSE2GPO[ $CSE.Properties[3].Value ] -split "`n" )
                ForEach( $GPO in $GPOs )
                {
                    try
                    {
                        if( ! [string]::IsNullOrEmpty( $GPO.Trim() ) )
                        {
                            $GPOTotalTimes.Add( $GPO , $duration )
                        }
                        ## else empty string so ignore
                    }
                    catch
                    {
                        ## already have it so we add to the time
                        [double]$alreadyGot = $GPOTotalTimes.Get_Item( $GPO )
                        $GPOTotalTimes.Set_Item( $GPO , $alreadyGot + $duration )
                    }
                }

                [pscustomobject]@{
                    Source    = 'CSE'
                    PhaseName = $CSE.Properties[2].Value
                    StartTime = $CSE.TimeCreated.AddMilliseconds( -$CSE.Properties[0].Value )
                    EndTime   = $CSE.TimeCreated
                    Duration  = $duration 
                    GPOs      = ($GPOs -join ', ').Trim( '[, ]') }
            } ) )
            
            if( $lastToFinish )
            {
                "Overall Group Policy Processing Duration:`t" + ( "{0:N2}" -f ( $lastToFinish - $startProcessingEvent.TimeCreated ).TotalSeconds ) + " Seconds"
                ''
            }
            ($CSEtimings | Sort-Object -Property StartTime | Format-Table -Property $Format -AutoSize | Out-String).Trim() -split "`r`n" | ForEach-Object { "$indent$_" }
            
            if( $GPOTotalTimes -and $GPOTotalTimes.Count )
            {
                ''
                "$($GPOTotalTimes.Count) processed GPO CSEs sorted by the most time spent processing them (seconds)"
                $GPOTotalTimes.GetEnumerator() | Where-Object Name | Sort-Object -Property Value -Descending | Format-Table -AutoSize -Property @{n='GPO';e={$_.Name}},@{n='Time Spent (s)';e={$_.Value}}
            }
        }

        ## wraps text for reasons unknown - bug
        if( $warnings -and $warnings.Count )
        {
            ''
            $warnings | Write-Warning
        }
    }
}

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

$windowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())

[bool]$global:dumpForOffline = $false
[string]$global:logsFolder = $null
[string]$username = $null
[string]$UserDomain = $null

## if we have extra parameters then let's go into debug mode - must be used with XenDesktop credentials even if dummy until support for null parameters arrives
if( $args.Count -gt 7 -or $env:CONTROLUP_SUPPORT )
{
    $global:logsFolder = $(if( $args.Count -gt 7 ) { $args[ 7 ] } else { $env:CONTROLUP_SUPPORT } )
    $DebugPreference = 'Continue'

    if( $global:logsFolder -match '^Prep:(\d+)$' )
    {
        if( ! ( $windowsPrincipal.IsInRole( [System.Security.Principal.WindowsBuiltInRole]::Administrator )))
        {
           Throw 'This script must be run with administrative privilege'
        }
        [int]$logSize = $Matches[1]
        $securityEventLog = Get-WinEvent -ListLog 'Security'
        [string]$size = $null

        if( $logSize -lt 1 )
        {
            Throw "$logSize cannot be less than 1MB"
        }
        if( $logSize -lt $suggestedSecurityEventLogSizeMB )
        {
            Write-Warning "Log size of $($logSize)MB is less than the recommended $($suggestedSecurityEventLogSizeMB)MB"
        }
        elseif( $logSize -lt $securityEventLog.MaximumSizeInBytes / 1MB )
        {
            Write-Warning "New Security event log size of $($logSize)MB is less than the current $([int]($securityEventLog.MaximumSizeInBytes / 1MB))MB"
        }
        elseif( $logSize -gt $securityEventLog.MaximumSizeInBytes / 1MB )
        {
            Write-Debug "Increasing security event log maximum size to $($logSize)MB from $([int]($securityEventLog.MaximumSizeInBytes / 1MB))MB"
            $size = "/maxsize:$($logSize * 1MB)"
        }
        else
        {
            Write-Warning "Security event log already has max size of $($logSize)MB so not changing"
        }
        
        if( $securityEventLog.LogMode -ne 'Circular' )
        {
            Write-Warning "Security event log was previousy not set to overwrite (was $($securityEventLog.LogMode))"
        }
        
        wevtutil.exe set-log Security /retention:false /autobackup:false $size 
        
        $null = New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Audit" -Name 'ProcessCreationIncludeCmdLine_Enabled' -Value 1 -PropertyType 'Dword' -Force
        
        [string[]]$eventLogs = @( 'Microsoft-Windows-PrintService/Operational' , 'Microsoft-Windows-GroupPolicy/Operational' , 'Microsoft-Windows-TaskScheduler/Operational' , 'Microsoft-Windows-User Profile Service/Operational' , 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' )
        [int]$newEventLogSize = 10MB
        ForEach( $eventLog in $eventLogs )
        {
            $eventLogProperties = Get-WinEvent -ListLog $eventLog
            if( $eventLogProperties )
            {
                $commandLine =  "`"$eventLog`" /retention:false /autobackup:false /enabled:true"
                if( $eventLogProperties.MaximumSizeInBytes -ge $newEventLogSize )
                {
                    Write-Warning "Event log `"$eventLog`" already has max size of $([int]($eventLogProperties.MaximumSizeInBytes / 1MB))MB so not changing"
                }
                else
                {
                    $commandLine += " /maxsize:$newEventLogSize"
                }
                Start-Process -FilePath "wevtutil.exe" -ArgumentList "set-log $commandLine" -Wait -WindowStyle Hidden
            }
        }

        [array]$requiredAuditEvents = @(
            [pscustomobject]@{ 'Policy' = 'Process Creation'     ; 'CategoryGuid' = '6997984C-797A-11D9-BED3-505054503030' ; 'SubCategoryGuid' = '0cce922b-69ae-11d9-bed3-505054503030' }
            [pscustomobject]@{ 'Policy' = 'Process Termination'  ; 'CategoryGuid' = '6997984C-797A-11D9-BED3-505054503030' ; 'SubCategoryGuid' = '0cce922c-69ae-11d9-bed3-505054503030' }
        )
        if( ! ( ([System.Management.Automation.PSTypeName]'Win32.Advapi32').Type ) )
        {
            [void](Add-Type -MemberDefinition $AuditDefinitions -Name 'Advapi32' -Namespace 'Win32' -UsingNamespace System.Text,System.ComponentModel,System.Security,System.Security.Principal -Debug:$false)
        }
        [int]$privReturn = [Win32.Advapi32+TokenManipulator]::AddPrivilege( [Win32.Advapi32+Rights]::SeSecurityPrivilege )
        if( $privReturn )
        {
            Write-Warning "Failed to enable SeSecurityPrivilege"
        }
        ForEach( $requiredAuditEvent in $requiredAuditEvents )
        {
            if( ! ( Set-SystemPolicy -categoryGuid $requiredAuditEvent.CategoryGuid -subCategoryGuid $requiredAuditEvent.SubCategoryGuid  ) )
            {
                Write-Warning "Unable to set $($requiredAuditEvent.Policy)"
            }
        }
        Exit 0
    }
    elseif( $global:logsFolder[0] -eq '+' )
    {
        if( ! ( $windowsPrincipal.IsInRole( [System.Security.Principal.WindowsBuiltInRole]::Administrator )))
        {
           Throw 'This script must be run with administrative privilege'
        }
        ## we are dumping the logs
        $global:logsFolder = $global:logsFolder.Substring(1)
        if( ! ( Test-Path -Path $global:logsFolder -PathType Container -ErrorAction SilentlyContinue ) )
        {
            $dumpDir = New-Item -Path $global:logsFolder -ItemType Directory -Force -ErrorAction Stop
        }
        wevtutil.exe export-log "Application" $(Join-Path -Path $global:logsFolder -ChildPath 'Application.evtx')
        wevtutil.exe export-log "Security" $(Join-Path -Path $global:logsFolder -ChildPath 'Security.evtx')
        wevtutil.exe export-log "Microsoft-Windows-GroupPolicy/Operational" $(Join-Path -Path $global:logsFolder -ChildPath 'Group Policy.evtx')
        wevtutil.exe export-log "Microsoft-Windows-PrintService/Operational" $(Join-Path -Path $global:logsFolder -ChildPath 'Print Service.evtx')
        wevtutil.exe export-log "Microsoft-Windows-TaskScheduler/Operational" $(Join-Path -Path $global:logsFolder -ChildPath 'Task Scheduler.evtx')
        wevtutil.exe export-log "Microsoft-Windows-User Profile Service/Operational" $(Join-Path -Path $global:logsFolder -ChildPath 'User Profile Service.evtx')
        wevtutil.exe export-log "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational" $(Join-Path -Path $global:logsFolder -ChildPath 'Terminal Services LSM.evtx')
        $global:dumpForOffline = $true
        
        #region FSLogix Offline Dump

        [string]$FSLogixLogDir = $null

        try {
            $FSLogixLogDir = Get-ItemPropertyValue -Path HKLM:\SOFTWARE\FSLogix\Logging -Name Logdir -ErrorAction SilentlyContinue
        }
        Catch {
            #LogDir registry value not found. Set to default:
            Write-Verbose "LogDir value not set. Setting LogDir to default path"
        }

        if ($FSLogixLogDir -eq $null) {
            $FSLogixLogDir = Join-Path -Path ([Environment]::GetFolderPath( [System.Environment+SpecialFolder]::CommonApplicationData )) -ChildPath 'FSLogix\Logs'
        }
        
        [string]$FSLogixProfileLogDir = Join-Path -Path $FSLogixLogDir -ChildPath 'Profile'
        if (Test-Path -Path $FSLogixProfileLogDir -ErrorAction SilentlyContinue ) {
            Write-Verbose "Found FSLogix Profile Log directory."
            $profileLog = Get-ChildItem -Path $FSLogixProfileLogDir | Where-Object Name -like "*$($($start).ToString("yyyyMMdd"))*"
            if ( Test-Path $profileLog.FullName -ErrorAction SilentlyContinue ) {
                Copy-Item -Path $profileLog.FullName -Destination $(Join-Path -Path $global:logsFolder -ChildPath 'FSLogixProfileLog.txt')
            } else {
                Write-Verbose "Unable to determine or find FSLogix profile log file."
            }
        }
         #endregion   

        if( Test-Path -Path $appVolumesLogFile -ErrorAction SilentlyContinue )
        {
            if( ( $appvolumesKey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -Name DisplayName -ErrorAction SilentlyContinue | Where-Object DisplayName -match 'App Volumes Agent' | Select-Object -ExpandProperty PSPath ) )
            {
                $global:appVolumesVersion = Get-ItemProperty -Path $appvolumesKey -Name DisplayVersion | Sort-Object -Property DisplayVersion -Descending | Select-Object -ExpandProperty DisplayVersion -First 1
                ## insert version number in file name so we can pull it out when processing since we do different things for different versions
                ## TODO also need to put WaitForFirstVolumeOnly in file name so can pull out when running offline
                try {
                    if( ( Get-ItemPropertyValue -Path HKLM:\SYSTEM\CurrentControlSet\Services\svservice\Parameters -Name WaitForFirstVolumeOnly ) -eq  0 ) {
                        $global:WaitForFirstVolumeOnly = $false
                    }
                } catch {
                    Write-Debug 'AppVolumes: WaitForFirstVolumeOnly value not found. Using Default'
                }
                Copy-Item -Path $global:appVolumesLogFile -Destination ( Join-Path -Path $global:logsFolder -ChildPath ( ( Split-Path -Path $global:appVolumesLogFile -Leaf ) -replace '(.*)(\.\w{3})' , ( '$1' + ".$global:appVolumesVersion" + ".$global:WaitForFirstVolumeOnly" + '$2' )))
            }
            else
            {
                Copy-Item -Path $global:appVolumesLogFile -Destination $global:logsFolder
            }
        }
    }
    elseif( Test-Path -LiteralPath $global:logsFolder -PathType Container -ErrorAction SilentlyContinue )
    {
        $offline = $true

        ## look for event log files so we can use instead of live logs
        Get-ChildItem -Path $global:logsFolder -Filter '*.evtx' -ErrorAction SilentlyContinue | ForEach-Object `
        {
            $file = $_
            switch -Regex( $file.BaseName )
            {
                'sec'         { $global:securityParams = @{ 'Path' = $file.FullName } ; break }
                'group|gpo'   { $global:groupPolicyParams = @{ 'Path' = $file.FullName } ; break }
                'ts|terminal' { $global:terminalServicesParams = @{ 'Path' = $file.FullName } ; break }
                'prof'        { $global:userProfileParams = @{ 'Path' = $file.FullName  } ; break }
                'app'         { $global:citrixUPMParams = @{ 'Path' = $file.FullName } ; $global:AppVolumesParams = @{ 'Path' = $file.FullName } ; break }
                'sched'       { $global:scheduledTasksParams = @{ 'Path' = $file.FullName } ; break }
                'print'       { $global:printServiceParams = @{ 'Path' = $file.FullName } ; break }
                'appdefaults'  { $global:windowsShellCoreParams = @{ 'Path' = $file.FullName } ; break }
                'appreadiness' { $global:appReadinessParams = @{ 'Path' = $file.FullName } ; break }
            }
        }
        if( ! $global:securityParams[ 'Path' ] )
        {
            Write-Warning "Could not find Security event log file in `"$global:logsFolder`""
        }
        if( ! $global:groupPolicyParams[ 'Path' ] )
        {
            Write-Warning "Could not find Group Policy operational event log file in `"$global:logsFolder`""
        }
        if( ! $global:terminalServicesParams[ 'Path' ] )
        {
            Write-Warning "Could not find Terminal Services-Local Session Manager operational event log file in `"$global:logsFolder`""
        }
        if( ! $global:userProfileParams['Path' ] )
        {
            Write-Warning "Could not find User Profile Service operational event log file in `"$global:logsFolder`""
        }
        if( ! $global:scheduledTasksParams[ 'Path' ] )
        {
            Write-Warning "Could not find User Task Scheduler operational event log file in `"$global:logsFolder`""
        }
        if( ! $global:citrixUPMParams[ 'Path' ] )
        {
            Write-Warning "Could not find Application event log (for Citrix Profile Management) file in `"$global:logsFolder`""
        }
        if( ! $global:AppVolumesParams[ 'Path' ] )
        {
            Write-Warning "Could not find Application event log (for App Volumes) file in `"$global:logsFolder`""
        }
        if( ! $global:printServiceParams[ 'Path' ] )
        {
            Write-Warning "Could not find User Print Service operational event log file in `"$global:logsFolder`""
        }
        if( ! $global:windowsShellCoreParams[ 'Path' ] )
        {
            Write-Warning "Could not find Windows-Shell-Core AppDefaults event log file in `"$global:logsFolder`""
        }
        if( ! $global:appReadinessParams[ 'Path' ] )
        {
            Write-Warning "Could not find App Readiness Admin event log file in `"$global:logsFolder`""
        }
        Set-Variable -Name CommandLine -Value 8 -Option ReadOnly -ErrorAction SilentlyContinue

        ## Appvolumes log file has had the version number put in it so we can extract that too
        $svserviceLogFile = Get-ChildItem -Path $global:logsFolder -Filter "svservice.*.log"
        if( $svserviceLogFile )
        {
            if( $svserviceLogFile -is [array] )
            {
                Write-Warning "$($svserviceLogFile.Count) app volumes log files found in `"$global:logsFolder`""
            }
            elseif( $svserviceLogFile.BaseName -match '(\d{1,4}\.\d{1,4}\.\d{1,4}\.\d{1,4})\.(true|false)$' )
            {
                $global:appVolumesVersion = $Matches[1]
                $global:WaitForFirstVolumeOnly = [bool]::Parse( $Matches[2] )
            }
            else
            {
                Write-Warning "Unable to find version number in `"$($svserviceLogFile.BaseName)`""
            }
            $global:appVolumesLogFile = $svserviceLogFile | Select-Object -First 1 -ExpandProperty FullName
        }
        <#
        [string]$svserviceLogfie = Join-Path -Path $global:logsFolder -ChildPath 'svservice.log'
        if( Test-Path -Path $svserviceLogfie -ErrorAction SilentlyContinue )
        {
            $appVolumesLogFile = $svserviceLogfie
        }
        #>

        [string]$jsonFile = Join-Path -Path $global:logsFolder -ChildPath 'logon.json'
        if( ! ( Test-Path -Path $jsonFile -PathType Leaf -ErrorAction SilentlyContinue ) )
        {
            Throw "Unable to find JSON file `"$jsonFile`" containing previosuly saved logon information"
        }
        $logonDetails = Get-Content -Path $jsonFile -ErrorAction SilentlyContinue | ConvertFrom-Json        ## Read username and domain for now as the rest will be retrieved from the JSON later
        if( $logonDetails )
        {
            $UserName = $logonDetails.UserName
            $UserDomain = $logonDetails.UserDomain
            if( [string]::IsNullOrEmpty( $UserName ) -or [string]::IsNullOrEmpty( $UserDomain ) )
            {
                Throw "Failed to get user name and/or domain details from JSON file `"$jsonFile`" containing previosuly saved logon information"
            }
        }
        else
        {
            Throw "Unable to get details from JSON file `"$jsonFile`" containing previosuly saved logon information"
        }

        if (Test-Path "${env:ProgramFiles(x86)}\CloudVolumes\Agent\Logs\svservice.log") {
            Write-Verbose "Found AppVolumes log file."
            Copy-Item -Path "${env:ProgramFiles(x86)}\CloudVolumes\Agent\Logs\svservice.log" -Destination $(Join-Path -Path $global:logsFolder -ChildPath 'svservice.log')
        } else {
            Write-Verbose "Unable to determine or find AppVolumes log file."
        }
    }
    Write-Debug "Running script as Windows version $global:windowsMajorVersion"
}
else ## online
{
    if( ! ( $windowsPrincipal.IsInRole( [System.Security.Principal.WindowsBuiltInRole]::Administrator )))
    {
       Throw 'This script must be run with administrative privilege'
    }
}

#region Get local session information
$TSSessions = @'
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
public class RDPInfo
{
    [DllImport("wtsapi32.dll")]
    static extern IntPtr WTSOpenServer([MarshalAs(UnmanagedType.LPStr)] String pServerName);

    [DllImport("wtsapi32.dll")]
    static extern void WTSCloseServer(IntPtr hServer);

    [DllImport("wtsapi32.dll")]
    static extern Int32 WTSEnumerateSessions(
        IntPtr hServer,
        [MarshalAs(UnmanagedType.U4)] Int32 Reserved,
        [MarshalAs(UnmanagedType.U4)] Int32 Version,
        ref IntPtr ppSessionInfo,
        [MarshalAs(UnmanagedType.U4)] ref Int32 pCount);

    [DllImport("wtsapi32.dll")]
    static extern void WTSFreeMemory(IntPtr pMemory);

    [DllImport("Wtsapi32.dll")]
    static extern bool WTSQuerySessionInformation(System.IntPtr hServer, int sessionId, WTS_INFO_CLASS wtsInfoClass, out System.IntPtr ppBuffer, out uint pBytesReturned);

    [StructLayout(LayoutKind.Sequential)]
    private struct WTS_SESSION_INFO
    {
        public Int32 SessionID;
        [MarshalAs(UnmanagedType.LPStr)]
        public String pWinStationName;
        public WTS_CONNECTSTATE_CLASS State;
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
        WTSClientProtocolType
    }

    public enum WTS_CONNECTSTATE_CLASS
    {
        WTSActive,       // 0
        WTSConnected,    // 1
        WTSConnectQuery, // 2
        WTSShadow,       // 3
        WTSDisconnected, // 4
        WTSIdle,         // 5
        WTSListen,       // 6
        WTSReset,        // 7
        WTSDown,         // 8
        WTSInit          // 9
    }

    public static IntPtr OpenServer(String Name)
    {
        IntPtr server = WTSOpenServer(Name);
        return server;
    }

    public static void CloseServer(IntPtr ServerHandle)
    {
        WTSCloseServer(ServerHandle);
    }

    public static List<string> ListUsers(String ServerName)
    {
        IntPtr serverHandle = IntPtr.Zero;
        List<String> resultList = new List<string>();
        serverHandle = OpenServer(ServerName);

        try
        {
            IntPtr SessionInfoPtr = IntPtr.Zero;
            IntPtr userPtr = IntPtr.Zero;
            IntPtr domainPtr = IntPtr.Zero;
            IntPtr clientNamePtr = IntPtr.Zero;
            IntPtr winStationNamePtr = IntPtr.Zero;
            IntPtr sessionStatePtr = IntPtr.Zero;
            Int32  sessionCount = 0;
            Int32  retVal = WTSEnumerateSessions(serverHandle, 0, 1, ref SessionInfoPtr, ref sessionCount);
            Int32  dataSize = Marshal.SizeOf(typeof(WTS_SESSION_INFO));
            IntPtr currentSession = (IntPtr)SessionInfoPtr;
            uint bytes = 0;

            if (retVal != 0)
            {
                for (int i = 0; i < sessionCount; i++)
                {
                    WTS_SESSION_INFO si = (WTS_SESSION_INFO)Marshal.PtrToStructure((System.IntPtr)currentSession, typeof(WTS_SESSION_INFO));
                    currentSession += dataSize;
                    

                    WTSQuerySessionInformation(serverHandle, si.SessionID, WTS_INFO_CLASS.WTSUserName, out userPtr, out bytes);
                    WTSQuerySessionInformation(serverHandle, si.SessionID, WTS_INFO_CLASS.WTSDomainName, out domainPtr, out bytes);
                    WTSQuerySessionInformation(serverHandle, si.SessionID, WTS_INFO_CLASS.WTSClientName, out clientNamePtr, out bytes);
                    WTSQuerySessionInformation(serverHandle, si.SessionID, WTS_INFO_CLASS.WTSWinStationName, out winStationNamePtr, out bytes);
                    WTSQuerySessionInformation(serverHandle, si.SessionID, WTS_INFO_CLASS.WTSConnectState, out sessionStatePtr, out bytes);

                    if(Marshal.PtrToStringAnsi(domainPtr).Length > 0 && Marshal.PtrToStringAnsi(userPtr).Length > 0)
                    {
                        if(Marshal.PtrToStringAnsi(clientNamePtr).Length < 1)                       
                            resultList.Add("UserName:" + Marshal.PtrToStringAnsi(domainPtr) + "\\" + Marshal.PtrToStringAnsi(userPtr) + "\tSessionID:" + si.SessionID + "\tClientName:N/A" + "\tSessionName:N/A" + "\tSessionState:" + Marshal.ReadInt16( sessionStatePtr ) );
                        else
                            resultList.Add("UserName:" + Marshal.PtrToStringAnsi(domainPtr) + "\\" + Marshal.PtrToStringAnsi(userPtr) + "\tSessionID:" + si.SessionID + "\tClientName:" + Marshal.PtrToStringAnsi(clientNamePtr) + "\tSessionName:" + Marshal.PtrToStringAnsi(winStationNamePtr) + "\tSessionState:" + Marshal.ReadInt16( sessionStatePtr ) );
                    }
                    WTSFreeMemory(clientNamePtr);
                    WTSFreeMemory(userPtr);
                    WTSFreeMemory(domainPtr);
                    WTSFreeMemory(winStationNamePtr);
                    WTSFreeMemory(sessionStatePtr);
                }
                WTSFreeMemory(SessionInfoPtr);
            }
        }
        catch(Exception ex)
        {
            Console.WriteLine("Exception: " + ex.Message);
            resultList.Add("Exception: " + ex.Message);
        }
        finally
        {
            CloseServer(serverHandle);
            
        }
        return resultList;
    }
}
'@

#here we sort out the parameters.  There is an issue with some parameters not being passed so we need to run some checks and validate them.
$SessionId = $(if( $args.Count -ge 3) { $args[2] })
$XDUsername = $null
$XDPassword = $null

if( [string]::IsNullOrEmpty( $UserName ) -or [string]::IsNullOrEmpty( $UserDomain ) -and $args.Count )
{
    $args_fix = ($args[0] -split '\\')
    if( ! $args_fix -or $args_fix.Count -ne 2 )
    {
        Throw 'Must be run with at least the domain\username of the user to report on'
    }
    $UserName = $args_fix[1]
    $UserDomain = $args_fix[0]
}

if( [string]::IsNullOrEmpty( $UserName ) -or [string]::IsNullOrEmpty( $UserDomain ) )
{
    Throw 'Must be run with at least the domain\username of the user to report on'
}

$foundAllParameters = $false
[int]$currentSessionState = -1

if( ! $offline )
{
    Add-Type $TSSessions -Debug:$false

    $sessionInfo = [RDPInfo]::listUsers("localhost")
    $sessionArray = @()

    #converts Output from pInvoke to PowerShell Object
    foreach ($line in $sessionInfo) {
        $sessionInfoObject = New-Object System.Object
        foreach ($object in ($line -split "\t")) {
    
            if ($object -like "*UserName*") { Write-Debug "Username: $object"
                $sessionInfoObject | Add-Member -type NoteProperty -name UserName -value ($object -split ":")[1] }
            if ($object -like "*SessionID*") { Write-Debug "SessionID: $object"
                $sessionInfoObject | Add-Member -type NoteProperty -name SessionID -value ($object -split ":")[1] }
            if ($object -like "*ClientName*") { Write-Debug "ClientName: $object"
                $sessionInfoObject | Add-Member -type NoteProperty -name ClientName -value ($object -split ":")[1] }
            if ($object -like "*SessionName*") { Write-Debug "SessionName: $object"
                $sessionInfoObject | Add-Member -type NoteProperty -name SessionName -value ($object -split ":")[1] }
            if ($object -like "*SessionState*") { Write-Debug "SessionState: $object"
                $sessionInfoObject | Add-Member -type NoteProperty -name SessionState -value ($object -split ":")[1] }
        }
        $sessionArray += $sessionInfoObject
    
    }
    #endregion

    foreach ($session in $sessionArray) {
        try {
            if ($session.Username -eq $args[0] -and $session.SessionId -eq $args[2] -and $session.ClientName -eq $args[4] -and $session.SessionName -eq $args[3] ) {
                Write-Verbose "All session parameters found"
                $SessionName = $args[3]
                $ClientName = $args[4]
                $currentSessionState = $session.SessionState
                $foundAllParameters = $true
            }
        }
        catch {
            ## not all parameters are required when run manually
        }
    }
}

if (-not($foundAllParameters)) {
    Write-Verbose "Only partial parameters found"
    $UserTest = $args[0]
    $matchingSession = $sessionArray | Where-Object -FilterScript { $_.UserName -eq $UserTest }
    
    if ($matchingSession -and -not ( $matchingSession -is [array]) ) {
        $ClientName = $matchingSession.ClientName
        $SessionName = $matchingSession.SessionName
        $SessionId = $matchingSession.SessionId
        $currentSessionState = $matchingSession.SessionState
    } else {
        $SessionName = $null
        $clientName = $null
        if( ! $offline ) {
            Write-Warning "User $UserTest appears not to be logged on currently so data may be incorrect"
        }
    }
}

Write-Debug "$($args.Count) arguments passed"

if( ! $ClientName -and $args.Count -ge 5 )
{
    $ClientName = $args[4]
}

if( ! $SessionName -and $args.Count -ge 4 )
{
    $SessionName = $args[3]
}

if ( $args.Count -ge 7 -and $args[5] -and $args[6]) {
    $XDUsername = $args[5]
    $XDPassword = $args[6]
}

if ($SessionName -eq $null -and $ClientName -eq $null -and $args.count -eq 5) {
    $XDUsername = $args[3]
    $XDPassword = $args[4]
}

Write-Debug "Logon Parameters discovered:"
Write-Debug "Username:     $Username"
Write-Debug "UserDomain:   $userDomain"
Write-Debug "ClientName:   $ClientName"
Write-Debug "SessionName:  $SessionName"
Write-Debug "SessionState: $currentSessionState"
Write-Debug "SessionId:    $SessionID"
Write-Debug "XDUsername:   $XDUserName"

[hashtable]$params = @{
    'Username' = $Username
    'UserDomain' =  $UserDomain
    'ClientName' = $clientName
}

if( $args.Count -ge 2 -and ![string]::IsNullOrEmpty( $args[1] ) )
{
    $params.Add( 'CUDesktopLoadTime' , ( $args[1] -replace ',' , '.' ) ) ## if passed as 1,234 change to 1.234
    Write-Debug "CUDesktopLoadTime: $($params[ 'CUDesktopLoadTime' ])"
}

if ($SessionName -imatch "RDP") {
        Get-LogonDurationAnalysis @params ## TODO what if this is a Horizon View session where we'll need sessionid
    }
else {
    $params.Add( 'HDXSessionId' , $SessionId )

    if ($XDUsername -and $XDPassword ) {
        Get-LogonDurationAnalysis @params -XDUsername $XDUsername -XDPassword (ConvertTo-SecureString -String $XDPassword -AsPlainText -Force)
    } else {
        Get-LogonDurationAnalysis @params
    }
}
