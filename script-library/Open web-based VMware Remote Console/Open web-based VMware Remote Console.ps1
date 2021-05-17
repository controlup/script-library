<#
    .SYNOPSIS
        This script launches a remote console session in your default browser for a vSphere/ESXi VM.
    .DESCRIPTION
        This requires PowerCLI v5.0 or greater, Adobe Flash, and the VMware client integration/Remote 
        console plug-in installed (vmware-vmrc-win32-x86.exe) in order to function. Adapted from 
        William Lam and Dylan Thompson.
    .PARAMETER VMName
        The name of the VM as known by vCenter.
    .PARAMETER VCName
        The name of the vCenter. Needed to perform all actions related to the VM.
    .PARAMETER VCUser
        User that is authorized to login to vCenter
    .PARAMETER VCUserPwd
        The password for the vCenter user
    .EXAMPLE
        .\LaunchVMConsole.ps1 DC01 vcenter.company.org company\admin P@ssw0rd
    .LINK
        http://www.virtuallyghetto.com/2011/10/how-to-generate-vm-remote-console-url.html
    .LINK
        https://dthomo.wordpress.com/2012/12/10/generate-list-of-remote-console-urls-for-vcenter-5-0-through-powercli/
#>

<#
$args[0] = VM Name
$args[1] = VC Name server.domain.com
$args[2] = VC User domain\username
$args[3] = VCUserPassword
$args[4] = Hypervisor Platform
#>

#region Setup the script

$ErrorActionPreference = "Stop"

# variables
if ($args.count -ne 5) {
    Write-Host "ControlUp must be connected to the VM's hypervisor for this command to work. Please connect to the appropriate hypervisor and try again."
    Exit 1
}
$VMName = $args[0]
$VCName = $args[1]
$VCUser = $args[2]
$VCUserPwd = $args[3]
$Hyper = $args[4]

#endregion

#region Functions

Function Load-PowerCLI ()
{
    if (Get-Command "*find-module")
    {
	$PCLIver = 	(Find-Module "*vmware.powercli").Version.major
    }
    If ($PCLIver -eq $null) 
	{
        $PCLIver = (Get-ItemProperty HKLM:\software\Microsoft\Windows\CurrentVersion\Uninstall\* | where displayname -match 'PowerCLI').DisplayVersion
		If ($PCLIver -eq $null) 
		{
			$PCLIver = (Get-ItemProperty HKLM:\software\WoW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | where displayname -match 'PowerCLI').DisplayVersion
		}
        If ($PCLIver -eq $null) 
		{
            Write-Host "PowerCLI is not installed on this computer. Please download and install the PowerCLI package from VMware at https://www.vmware.com/support/developer/PowerCLI/"
            Exit
        }
    }
	($PCLIver| Out-String).Split(".")[0]

	If ($PCLIver -ge 10) 
	{
		$PCLI = "vmware.powercli"
            Try 
			{
				Import-Module -Name $PCLI
            } 
			Catch 
			{
                  Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
                  Exit 1
            }
    }
    elseIf ($PCLIver -ge "6") 
	{
		$PCLI = "VMware.VimAutomation.Core"
        If ((Get-Module -Name $PCLI -ErrorAction SilentlyContinue) -eq $null) 
		{
            Try 
			{
                  Import-Module $PCLI
            } 
			Catch 
			{
                  Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
                  Exit 1
            }
        }
    }
	ElseIf ($PCLIver -ge "5") 
	{
		$PCLI = "VMware.VimAutomation.Core"
        If ((Get-PSSnapin $PCLI -ErrorAction "SilentlyContinue") -eq $null) 
		{
            Try 
			{
                Add-PSSnapin $PCLI
            } Catch 
			{
                Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
                Exit 1
            }
        }
    } 
	Else 
	{
        Write-Host "This version of PowerCLI seems to be unsupported. Please upgrade to the latest version of PowerCLI and try again."
    }
}


#endregion Functions

#region  ----Main----

If ($Hyper -ne "VMware") {
    Write-Host "This script only supports VMs on vSphere/ESXi, not other hypervisors. Please pick another computer and try again."
    Exit 1
}

Load-PowerCLI

Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Scope Session -Confirm:$false | Out-Null

Try {
    $VCObj = (Connect-VIServer -Server $VCName  -User $VCUser -Password $VCUserPwd -Force -WarningAction SilentlyContinue)
}
Catch {
    Write-Host "There was an error connecting to the vCenter. Please check your entries and try again."
    Exit 1
}

Write-Host "Connecting to $VCName ..."
$VCVersion = $VCObj.Version

If (!($VCVersion -ge '5.0')) {
    Write-Host "vSphere 4.x is not supported, this script cannot be run."
    Disconnect-VIServer -Confirm:$false
    Exit 1
}

If ((Get-PowerCLIVersion).build -ge '1295336') {
    # this command only works with PowerCLI 5.5R1 or higher.
    Write-Host "Connecting to $VMName ... (5.5 method)"
    Get-VM $VMName | Open-VMConsoleWindow -Confirm:$false
}
Else {
    # get vCenter UUID
    $UUID = $DefaultVIServer.InstanceUuid.toUpper()

    # get MoRef ID of the VM
    $MoRef = (Get-VM -Name $VMName).ExtensionData.MoRef.Value
        
    # assemble URL
    If ($VCVersion -eq '5.0') {
        $ConsoleLink = "https://" + $vcName + ":9443/vsphere-client/vmrc/vmrc.jsp?vm=" + $UUID + ":VirtualMachine:" + $MoRef
    }
    ElseIf ($VCVersion -eq '5.1' -or $VCVersion -eq '5.5') {
        $ConsoleLink = "https://" + $vcName + ":9443/vsphere-client/vmrc/vmrc.jsp?vm=urn:vmomi:VirtualMachine:" + $MoRef + ":" + $UUID
    }
    Else {
        Write-Host "This is an unknown version of vCenter. Exiting script."
        Disconnect-VIServer -Confirm:$false
        Exit 1
    }
    Write-Host "Connecting to $VMName ...(pre-5.5 method)"
    Start-Process $ConsoleLink
}

#endregion Main

