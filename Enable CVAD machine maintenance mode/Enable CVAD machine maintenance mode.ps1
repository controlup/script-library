<#
.SYNOPSIS
   Enable/Disable XenDesktop maintenance mode for the selected computer(s).

.PARAMETER machineName
   The name of the server being enabled - automatically supplied by CU
   
.PARAMETER maintenanceModeOperation
    Whether to enable or disable maintenance mode

.NOTES
    Modification History:

    2013/11/24 Zeev Esienberg  Original Version
    2023/02/16 Guy Leech       Replaced deprecated cmdlets, update to use scripting standards
    2023/02/23 Guy Leech       Check result of action, not perform action if already in desired state. Changes to make enable/disable same script
#>

[CmdletBinding()]

Param
(
    [string]$machineName ,
    [ValidateSet('enable','disable')]
    [string]$maintenanceModeOperation = 'enable'
)

$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { 'Continue' } else { 'SilentlyContinue' })
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { 'Continue' } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'ErrorAction' ] ) { $ErrorActionPreference } else { 'Stop' })

If ( $null -eq (Get-PSSnapin -Name Citrix.Broker.Admin.* -ErrorAction SilentlyContinue) )
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

# Because this is the main function of the script it is put into a try/catch frame so that any errors will be 
# handled in a ControlUp-friendly way.

[bool]$maintenanceModeToSet = $maintenanceModeOperation -ieq 'enable'

$TargetMachine = $null

Try {
    $TargetMachine = Get-BrokerMachine -machineName "*\$machineName"
}
Catch {
    Write-Error "Unable to get machine status - possibly insufficient administrative privileges" -ErrorAction Continue
    Write-Error $Error[1] -ErrorAction Continue
    Exit 1
}

[string]$not = ''
if( -Not $maintenanceModeToSet ) {
    $not = 'not '
}

If ($TargetMachine -ne $null) {
    if( $TargetMachine.InMaintenanceMode -eq $maintenanceModeToSet ) {
        Write-Warning "Machine is already $($not)in maintenance mode"
    } else {
        Set-BrokerMachineMaintenanceMode -InputObject $TargetMachine -MaintenanceMode:$maintenanceModeToSet
        if( $? ) {
            Write-Host "$machineName is now $($not)in maintenance Mode"
        } else {
            Write-Error "Problem setting maintenance mode on $($TargetMachine.MachineName)"
        }
    }
} else {
    Write-Error "Unable to find machine $machineName" -ErrorAction Continue
    Exit 1
}

