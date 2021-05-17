<#
.SYNOPSIS
   Disable XenDesktop maintenance mode for the selected computer(s).
.PARAMETER Identity
   The name of the server being taken out of maintenance mode - automatically supplied by CU
#>

$ErrorActionPreference = "Stop"

If ( (Get-PSSnapin -Name Citrix.Broker.Admin.* -ErrorAction SilentlyContinue) -eq $null )
{
    Try {
        Add-PsSnapin Citrix.Broker.Admin.*
    } Catch {
        # capture any failure and display it in the error section, then end the script with a return
        # code of 1 so that CU sees that it was not successful.
        Write-Error "Unable to load the snapin" -ErrorAction Continue
        Write-Error $Error[1] -ErrorAction Continue
        Exit 1
    }
}

$machineName = $args[0]

# Because this is the main function of the script it is put into a try/catch frame so that any errors will be 
# handled in a ControlUp-friendly way.

Try {
    $TargetMachine = Get-BrokerSharedDesktop "*\$machineName"
}
Catch {
    Write-Error "Unable to determine desktop status - possibly insufficient administrative privileges" -ErrorAction Continue
    Write-Error $Error[1] -ErrorAction Continue
    Exit 1
}

If ($TargetMachine -ne $null) {
    $TargetMachine | Set-BrokerSharedDesktop -InMaintenanceMode $false
    Write-Host "$machineName is no longer in Maintenance Mode"
} else {
    $TargetMachine = Get-BrokerPrivateDesktop "*\$machineName"
    if ($TargetMachine -ne $null) {
        $TargetMachine | Set-BrokerPrivateDesktop -InMaintenanceMode $false
        Write-Host "$machineName is no longer in Maintenance Mode"
    } else {
        Write-Error "Could not find desktop $machineName" -ErrorAction Continue
        Exit 1
    }
} 

