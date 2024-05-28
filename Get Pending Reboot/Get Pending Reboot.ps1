﻿<# 
	.SYNOPSIS 
 		Check various locations to determine if a reboot is pending

	.DESCRIPTION 
 		Check various registry locations to see if a computer rename, domain join/disjoin, Component Based Servicing, Application updates or Windows Update has registered the need for a reboot
        Also check if the SCCM client is installed and if so query that for a reboot requirement
    
    .EXAMPLE
        & '.\Get pending reboot.ps1'

 	    Check if a reboot is required and write to the standard output whether it is or not and if one is required it will list which components have requested it

     CONTEXT

        Computer

     MODIFICATION HISTORY 
        
        28/06/19   GRL   Iinitial version
        23/04/24   GRL   Fix for double exception when SSCM not installed
 
    .LINK 
        Based on code from Brian Wilhite at https://github.com/bcwilhite/PendingReboot

    .NOTES
        The script must be run with administrative privileges

    This script is provided with absolutely no warranty - use is at the user's own risk. ControlUp Technologies Ltd and their associates cannot be held responsible for any unexpected or undesired behaviour of this script.
#>

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'

$thisWindowsPrincipal = New-Object -TypeName System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())

# Check to see if we are currently running "as Administrator"
if ( ! $thisWindowsPrincipal  -or ! ($thisWindowsPrincipal.IsInRole( [System.Security.Principal.WindowsBuiltInRole]::Administrator ))  )
{
    Throw 'The script must be run with administrative privileges'
}

<#
.SYNOPSIS
    Test the pending reboot status on a local and/or remote computer.

.DESCRIPTION
    This function will query the registry on a local and/or remote computer and determine if the
    system is pending a reboot, from Microsoft/Windows updates, Configuration Manager Client SDK, Pending
    Computer Rename, Domain Join, Pending File Rename Operations and Component Based Servicing.

    ComponentBasedServicing = Component Based Servicing
    WindowsUpdate = Windows Update / Auto Update
    CCMClientSDK = SCCM 2012 Clients only (DetermineifRebootPending method) otherwise $null value
    PendingComputerRenameDomainJoin = Detects a pending computer rename and/or pending domain join
    PendingFileRenameOperations = PendingFileRenameOperations, when this property returns true,
                                  it can be a false positive
    PendingFileRenameOperationsValue = PendingFilerenameOperations registry value; used to filter if need be,
                                       Anti-Virus will leverage this key property for def/dat removal,
                                       giving a false positive

.PARAMETER ComputerName
    A single computer name or an array of computer names.  The default is localhost ($env:COMPUTERNAME).

.PARAMETER Credential
    Specifies a user account that has permission to perform this action. The default is the current user.
    Type a username, such as User01, Domain01\User01, or User@Contoso.com. Or, enter a PSCredential object,
    such as an object that is returned by the Get-Credential cmdlet. When you type a user name, you are
    prompted for a password.

.PARAMETER Detailed
    Indicates that this function returns a detailed result of pending reboot information, why the system is
    pending a reboot, not just a true/false response.

.PARAMETER SkipConfigurationManagerClientCheck
    Indicates that this function will not test the Client SDK WMI class that is provided by the System
    Center Configuration Manager Client.  This parameter is useful when SCCM is not used/installed on
    the targeted systems.

.PARAMETER SkipPendingFileRenameOperationsCheck
    Indicates that this function will not test the PendingFileRenameOperations MultiValue String property
    of the Session Manager registry key.  This parameter is useful for eliminating possible false positives.
    Many Anti-Virus packages will use the PendingFileRenameOperations MultiString Value in order to remove
    stale definitions and/or .dat files.

.EXAMPLE
    PS C:\> Test-PendingReboot

    ComputerName IsRebootPending
    ------------ ---------------
    WKS01                   True

    This example returns the ComputerName and IsRebootPending properties.

.EXAMPLE
    PS C:\> (Test-PendingReboot).IsRebootPending
    True

    This example will return a bool value based on the pending reboot test for the local computer.

.EXAMPLE
    PS C:\> Test-PendingReboot -ComputerName DC01 -Detailed

    ComputerName                     : dc01
    ComponentBasedServicing          : True
    PendingComputerRenameDomainJoin  : False
    PendingFileRenameOperations      : False
    PendingFileRenameOperationsValue :
    SystemCenterConfigManager        : False
    WindowsUpdateAutoUpdate          : True
    IsRebootPending                  : True

    This example will test the pending reboot status for dc01, providing detailed information

.EXAMPLE
    PS C:\> Test-PendingReboot -ComputerName DC01 -SkipConfigurationManagerClientCheck -SkipPendingFileRenameOperationsCheck -Detailed

    CommputerName                    : dc01
    ComponentBasedServicing          : True
    PendingComputerRenameDomainJoin  : False
    PendingFileRenameOperations      : False
    PendingFileRenameOperationsValue :
    SystemCenterConfigManager        :
    WindowsUpdateAutoUpdate          : True
    IsRebootPending                  : True

.LINK
    Background:
    https://blogs.technet.microsoft.com/heyscriptingguy/2013/06/10/determine-pending-reboot-statuspowershell-style-part-1/
    https://blogs.technet.microsoft.com/heyscriptingguy/2013/06/11/determine-pending-reboot-statuspowershell-style-part-2/

    Component-Based Servicing:
    http://technet.microsoft.com/en-us/library/cc756291(v=WS.10).aspx

    PendingFileRename/Auto Update:
    http://support.microsoft.com/kb/2723674
    http://technet.microsoft.com/en-us/library/cc960241.aspx
    http://blogs.msdn.com/b/hansr/archive/2006/02/17/patchreboot.aspx

    CCM_ClientSDK:
    http://msdn.microsoft.com/en-us/library/jj902723.aspx

.NOTES
    Author:  Brian Wilhite
    Email:   bcwilhite (at) live.com
#>

function Test-PendingReboot
{
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias("CN", "Computer")]
        [String[]]
        $ComputerName = $env:COMPUTERNAME,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.CredentialAttribute()]
        $Credential,

        [Parameter()]
        [Switch]
        $Detailed,

        [Parameter()]
        [Switch]
        $SkipConfigurationManagerClientCheck,

        [Parameter()]
        [Switch]
        $SkipPendingFileRenameOperationsCheck
    )

    process
    {
        foreach ($computer in $ComputerName)
        {
            try
            {
                $invokeWmiMethodParameters = @{
                    Namespace    = 'root/default'
                    Class        = 'StdRegProv'
                    Name         = 'EnumKey'
                    ComputerName = $computer
                    ErrorAction  = 'Stop'
                }

                $hklm = [UInt32] "0x80000002"

                if ($PSBoundParameters.ContainsKey('Credential'))
                {
                    $invokeWmiMethodParameters.Credential = $Credential
                }

                ## Query the Component Based Servicing Reg Key
                $invokeWmiMethodParameters.ArgumentList = @($hklm, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\')
                $registryComponentBasedServicing = (Invoke-WmiMethod @invokeWmiMethodParameters).sNames -contains 'RebootPending'

                ## Query WUAU from the registry
                $invokeWmiMethodParameters.ArgumentList = @($hklm, 'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\')
                $registryWindowsUpdateAutoUpdate = (Invoke-WmiMethod @invokeWmiMethodParameters).sNames -contains 'RebootRequired'

                ## Query JoinDomain key from the registry - These keys are present if pending a reboot from a domain join operation
                $invokeWmiMethodParameters.ArgumentList = @($hklm, 'SYSTEM\CurrentControlSet\Services\Netlogon')
                $registryNetlogon = (Invoke-WmiMethod @invokeWmiMethodParameters).sNames
                $pendingDomainJoin = ($registryNetlogon -contains 'JoinDomain') -or ($registryNetlogon -contains 'AvoidSpnSet')

                ## Query ComputerName and ActiveComputerName from the registry and setting the MethodName to GetMultiStringValue
                $invokeWmiMethodParameters.Name = 'GetMultiStringValue'
                $invokeWmiMethodParameters.ArgumentList = @($hklm, 'SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName\', 'ComputerName')
                $registryActiveComputerName = Invoke-WmiMethod @invokeWmiMethodParameters

                $invokeWmiMethodParameters.ArgumentList = @($hklm, 'SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\', 'ComputerName')
                $registryComputerName = Invoke-WmiMethod @invokeWmiMethodParameters

                $pendingComputerRename = $registryActiveComputerName -ne $registryComputerName -or $pendingDomainJoin

                ## Query PendingFileRenameOperations from the registry
                if (-not $PSBoundParameters.ContainsKey('SkipPendingFileRenameOperationsCheck'))
                {
                    $invokeWmiMethodParameters.ArgumentList = @($hklm, 'SYSTEM\CurrentControlSet\Control\Session Manager\', 'PendingFileRenameOperations')
                    $registryPendingFileRenameOperations = (Invoke-WmiMethod @invokeWmiMethodParameters).sValue
                    $registryPendingFileRenameOperationsBool = [bool]$registryPendingFileRenameOperations
                }

                ## Query ClientSDK for pending reboot status, unless SkipConfigurationManagerClientCheck is present
                $systemCenterConfigManager = 'N/A'
                if (-not $PSBoundParameters.ContainsKey('SkipConfigurationManagerClientCheck'))
                {
                    $invokeWmiMethodParameters.NameSpace = 'ROOT\ccm\ClientSDK'
                    $invokeWmiMethodParameters.Class = 'CCM_ClientUtilities'
                    $invokeWmiMethodParameters.Name = 'DetermineifRebootPending'
                    $invokeWmiMethodParameters.Remove('ArgumentList')

                    try
                    {
                        $sccmClientSDK = Invoke-WmiMethod @invokeWmiMethodParameters
                        $systemCenterConfigManager = $sccmClientSDK.ReturnValue -eq 0 -and ($sccmClientSDK.IsHardRebootPending -or $sccmClientSDK.RebootPending)
                    }
                    catch
                    {
                        Write-Warning -Message "SCCM query error - $_"
                    }
                }

                $isRebootPending = $registryComponentBasedServicing -or `
                    $pendingComputerRename -or `
                    $pendingDomainJoin -or `
                    $registryPendingFileRenameOperationsBool -or `
                    ($systemCenterConfigManager -and $systemCenterConfigManager -ne 'Not found') -or `
                    $registryWindowsUpdateAutoUpdate
                
                
                $result = New-Object -TypeName PSObject
                Add-Member -InputObject $result -MemberType NoteProperty -Name IsRebootPending -Value $isRebootPending
                if ($PSBoundParameters.ContainsKey('Detailed'))
                {
                    Add-Member -InputObject $result -MemberType NoteProperty -Name 'Component Based Servicing (CBS)' -Value $registryComponentBasedServicing
                    Add-Member -InputObject $result -MemberType NoteProperty -Name 'Pending Computer Rename or Domain Join'  -Value  $pendingComputerRename
                    Add-Member -InputObject $result -MemberType NoteProperty -Name 'Pending File Rename/Delete Operations'  -Value $registryPendingFileRenameOperationsBool
                    Add-Member -InputObject $result -MemberType NoteProperty -Name 'SCCM'  -Value  $systemCenterConfigManager     
                    Add-Member -InputObject $result -MemberType NoteProperty -Name 'Windows Update' -Value $registryWindowsUpdateAutoUpdate  
                }
                $result
            }
            catch
            {
                Write-Verbose "$Computer`: $_"
            }
        }
    }
}

$results = Test-PendingReboot -Detailed

if( $results )
{
    if( $results.IsRebootPending )
    {
        Write-Output -InputObject "A reboot is pending"
        $results | Select-Object -Property * -ExcludeProperty 'IsRebootPending' | Format-List
    }
    else
    {
        Write-Output -InputObject "A reboot is not pending"
    }
}
else
{
    Write-Warning -Message "Failed to determine if a reboot is pending or not"
}

