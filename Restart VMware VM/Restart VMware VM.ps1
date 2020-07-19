<#
.SYNOPSIS
   Restart a VMware VM
.DESCRIPTION
   Gracefully restart a VMware VM
.PARAMETER machineName
   The name of the VM being restarted  - automatically supplied by CU
.PARAMETER VIUser
   The name of the vCenter user - manually entered by user
.PARAMETER VIPwd
   The password of the vCenter user - manually entered by user
#>

$ErrorActionPreference = "Stop"

$machineName = $args[0]
$VIUser = $args[1]
$VIPwd = $args[2]
$vCenter = $args[3]

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
Load-PowerCLI

# Because this is the main function of the script it is put into a try/catch frame so that any errors will be 
# handled in a ControlUp-friendly way.

Try {
    Connect-VIServer $vCenter -User $VIUser -Password $VIPwd -Force | Out-Null
} Catch {
    # capture any failure and display it in the error section, then end the script with a return
    # code of 1 so that CU sees that it was not successful.
    Write-Error "Unable to connect to the vCenter server. Please correct and re-run the script." -ErrorAction Continue
    Write-Error $Error[1] -ErrorAction Continue
    Exit 1
}

Try {
    $machine = Get-VM | Where {$_.PowerState -eq "PoweredOn" -and $_.ExtensionData.Guest.ToolsStatus -eq "toolsOK" -and ($_.ExtensionData.Guest.Hostname).split(".")[0] -match $machineName} 
} Catch {
    # capture any failure and display it in the error section, then end the script with a return
    # code of 1 so that CU sees that it was not successful.
    Write-Error "There is a running VMs which is running the VMTools, but the vCenter does not have its hostname. Please investigate this problem and correct." -ErrorAction Continue
    Write-Error $Error[1] -ErrorAction Continue
    Exit 1
}

If ($machine -ne $null) {
    Restart-VMGuest $machine -Confirm:$false | Out-Null
    Write-Host "Restarting $machineName"
} else {
    # capture any failure and display it in the error section, then end the script with a return
    # code of 1 so that CU sees that it was not successful.
    Write-Error "The computer you want to restart does not seem to be managed by this vCenter. Please check and try again." -ErrorAction Continue
    Write-Error $Error[1] -ErrorAction Continue
    Exit 1
}

Disconnect-VIServer $vCenter -Confirm:$false

Write-Host "The VM was restarted successfully."

