<#  
.SYNOPSIS
      This script outputs a list of all the VMs from a XenDesktop site, that are un-registered and powered on.
.DESCRIPTION
      This script will list all the VMs in a specific XenDesktop site that are powered on in the hypervisor, and not registered to any DDC.
      It must be run on a XenDesktop broker. The script prompts for and requires at least a XenDesktop read-only admin account.
.PARAMETER 
       No parameters.
.EXAMPLE
        Get-PoweredOnAndUnregisteredVMs.ps1
.OUTPUTS
        A text list with the powered on VMs in a given XenDesktop site that are not registered. 
.LINK
        See http://www.ControlUp.com
#>

$ErrorActionPreference = "Stop"

If ( (Get-PSSnapin -Name Citrix.Broker.Admin.* -ErrorAction SilentlyContinue) -eq $null )
{
    Try {
        Add-PSSnapin Citrix.Broker.Admin.* | Out-Null
    } Catch {
        Write-Host "Unable to load Citrix Snapin. It is not possible to continue."
        Exit 1
    }
}

$VMs = Get-BrokerDesktop -RegistrationState Unregistered -PowerState On

If ($VMs) {
    Write-Host "Powered-on Unregistered VMs:"
    ForEach ($vm in $VMs) {
        $vm.MachineName
    }
} Else {
    Write-Host "All the powered-on VMs are registered succcessfully!"
}

