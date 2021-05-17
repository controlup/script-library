<#  
.SYNOPSIS
      This script outputs a list of all the VMs from a XenDesktop site, that are in maintenance mode.
.DESCRIPTION
      This script will list all the VMs in a specific XenDesktop site that are in maintenance mode.
      It must be run on a XenDesktop broker.
.PARAMETER 
       No parameters.
.EXAMPLE
        GetMaintenanceModeMachines.ps1
.OUTPUTS
        A text list with the VMs in a given XenDesktop site that are in maintenance mode.
.LINK
        See http://www.ControlUp.com
#>

$ErrorActionPreference = "Stop"

$VMs = @()

If ( (Get-PSSnapin -Name Citrix.Broker.Admin.* -ErrorAction SilentlyContinue) -eq $null )
{
    Try {
        Add-PSSnapin Citrix.Broker.Admin.* | Out-Null
    } Catch {
        Write-Host "Unable to load Citrix Snapin. It is not possible to continue."
        Exit 1
    }
}

$VMs = Get-BrokerDesktop -InMaintenanceMode $true

If ($VMs) {
    Write-Host "Machines In Maintenance Mode:"
    ForEach ($vm in $VMs) {
        $vm.MachineName
      }
      Write-Host "Total:"$vms.Count
} Else {
    Write-Host "There are no machines in maintenace mode!"
}


