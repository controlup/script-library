#requires -version 3

<#
.SYNOPSIS

Show Citrix Delivery Groups not used in the last n days and optionally disable them

.DETAILS

Delivery groups themselves do not have a last used property, this comes from the machines in a delivery group so a delivery group could have been used more recently than reported but the machine(s) used been deleted or moved to another deployment group

.PARAMETER daysNotAccessed

The number of days that the delivery group has not had a session launched from it

.PARAMETER disable

Disable delivery groups that have not had a session launched in the number of days specified by -daysNotAccessed

.PARAMETER ddc

The delivery controller to connect to. If not specified the local machine will be used.

.CONTEXT

Computer (but only on a Citrix Delivery Controller)

.NOTES


.MODIFICATION_HISTORY:

    @guyrleech 02/10/2020  Initial release
    @guyrleech 06/10/2020  Changes after feedback. Added -disable argument
#>

[CmdletBinding()]

Param
(
    [double]$daysNotAccessed = 30 ,
    [ValidateSet('true','false')]
    [string]$disable = 'false' ,
    [string]$ddc 
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputwidth = 400

Function Get-DeliveryGroupMachineDetails
{
    Param
    (
        [Parameter(Mandatory)]
        [array]$machinesByDeliveryGroup ,
        [datetime]$lastUsedBefore ,
        [AllowNull()]
        $deliveryGroupDetails
    )
    
    $lastUsed = $null
    [int]$machines = 0
    [int]$inMaintenanceMode = 0
    [int]$registered = 0
    [int]$sessionCount = 0
    [int]$logonsDisabled = 0
    [int]$available = 0
    [string]$provisioning = $null
    [string]$lastUsedBy = $null

    if( $deliveryGroupByMachine = $machinesByDeliveryGroup.Where( { ( ! $deliveryGroupDetails -and ! $_.Name ) -or ( $deliveryGroupDetails -and $_.Name -eq $deliveryGroupDetails.Name ) } ) )
    {
        ForEach( $machine in $deliveryGroupByMachine.Group )
        {
            $machines++
            if( ! $lastUsed -or $machine.LastConnectionTime -gt $lastUsed )
            {
                $lastUsed = $machine.LastConnectionTime
                $lastUsedBy = $machine.LastConnectionUser
            }
            if( $machine.InMaintenanceMode )
            {
                $inMaintenanceMode++
            }
            elseif( $machine.RegistrationState -eq 'Registered' -and $machine.WindowsConnectionSetting -eq 'LogonEnabled' )
            {
                $available++
            }
            if( $machine.WindowsConnectionSetting -ne 'LogonEnabled' )
            {
                $logonsDisabled++
            }
            if( $machine.RegistrationState -eq 'Registered' )
            {
                $registered++
            }
            $sessionCount += $machine.SessionCount
            if( ! [string]::IsNullOrEmpty( $provisioning ) )
            {
                if( $provisioning -ne $machine.ProvisioningType )
                {
                    $provisioning = $provisioning + $machine.ProvisioningType
                }
            }
            else
            {
                $provisioning = $machine.ProvisioningType
            }
        }
    }
    if( ! $lastUsed -or $lastUsed -le $lastUsedBefore )
    {
        ## don't report machines not in delivery groups if no machines found
        if( $deliveryGroupDetails -or $machines -gt 0 )
        {
            [pscustomobject]@{
                'Delivery Group' = $deliveryGroupDetails | Select-Object -ExpandProperty Name
                'Maintenance Mode' = $deliveryGroupDetails | Select-Object -ExpandProperty InMaintenanceMode
                'Enabled' = $deliveryGroupDetails | Select-Object -ExpandProperty Enabled
                'Description' = $deliveryGroupDetails | Select-Object -ExpandProperty Description
                'Last Used (d)' = $(if( $lastused ) { [math]::Round( ([datetime]::Now - $lastUsed).TotalDays , 0 ) } else { 'No record' } )
                'Last User' = ($lastUsedBy -split '\\')[-1]
                'Provisioning' = $provisioning
                'Sessions' = $sessionCount
                'Machines' = $machines ## $deliveryGroupByMachine | Select-Object -ExpandProperty Group | Measure-Object | Select-Object -ExpandProperty Count
                'Available' = $available
                'Maintenance' = $inMaintenanceMode
                'Registered' = $registered
                'Logons Disabled' = $logonsDisabled
            }
        }
    }
}
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

[hashtable]$brokerParameters = @{}

if( $PSBoundParameters[ 'ddc' ] )
{
    $brokerParameters.Add( 'AdminAddress' , $ddc )
}

## new CVAD have modules so use these in preference to snapins which are there for backward compatibility
if( ! (  Import-Module -Name Citrix.DelegatedAdmin.Commands -ErrorAction SilentlyContinue -PassThru ) `
    -and ! ( Add-PSSnapin -Name Citrix.Broker.Admin.* -ErrorAction SilentlyContinue -PassThru ) )
{
    Throw 'Failed to load Citrix PowerShell cmdlets - is this a Delivery Controller or have Studio or the PowerShell SDK installed ?'
}

[array]$machines = @( Get-BrokerMachine @brokerParameters )
[array]$machinesByDeliveryGroup = @( $machines | Group-Object -Property DesktopGroupName )

if( ! $machinesByDeliveryGroup -or ! $machinesByDeliveryGroup.Count )
{
    Throw "Got no machines"
}

[array]$deliveryGroups = @( Get-BrokerDesktopGroup @brokerParameters )

if( ! $deliveryGroups -or ! $deliveryGroups.Count )
{
    Throw "Got no delivery groups"
}

Write-Verbose -Message "Got $($deliveryGroups.Count) delivery groups and $($machines.Count) machines"

[datetime]$lastUsedDate = (Get-Date).AddDays( -$daysNotAccessed )

## look at all delivery groups but also looks for machines not in delivery groups
[array]$results = @( ForEach( $deliveryGroupDetails in $deliveryGroups )
{
    Get-DeliveryGroupMachineDetails -deliveryGroupDetails $deliveryGroupDetails -machinesByDeliveryGroup $machinesByDeliveryGroup -lastUsedBefore $lastUsedDate
}
    ## Report machines which aren't in a Delivery Group
    Get-DeliveryGroupMachineDetails -machinesByDeliveryGroup $machinesByDeliveryGroup -lastUsedBefore $lastUsedDate
)

Write-Output -InputObject "Found $($deliveryGroups.Count) delivery groups and $($machines.Count) machines in total, including $($results.Where( { -not $_.'Delivery Group' } )|Select-Object -ExpandProperty 'Machines') machines not assigned to any delivery group"
Write-Output -InputObject "Found $($results.Where( { $_.'Machines' -eq 0 } ).Count) delivery groups containing no machines"

if( $results -and $results.Count )
{
    Write-Output -InputObject "Found $($results.Where( { $_.'Delivery Group' } ).Count) delivery groups not used in last $daysNotAccessed days" ## already filtered in the function that builds the results

    $results | Sort-Object -Property 'Last Used (d)' | Format-Table -AutoSize -Property *

    if( $disable -eq 'true' )
    {
        [int]$disabled = 0
        [int]$actuallyDisabled = 0

        $results.Where( { ! [string]::IsNullOrEmpty( $_.'Delivery Group' ) -and $_.Enabled } ).ForEach(
        {
            $disabled++
            Write-Verbose "Disabling `"$($_.'Delivery Group')`""
            if( ! ( $result = Set-BrokerDesktopGroup -Name $_.'Delivery Group' -Enabled $false @brokerParameters  -PassThru ) -or $result.Enabled -ne $false )
            {
                Write-Warning -Message "Failed to disable delivery group `"$($_.'Delivery Group')`""
            }
            else
            {
                $actuallyDisabled++
            }
        })
        if( $disabled -eq 0 )
        {
            Write-Warning -Message "Found no delivery groups not used in the last $daysNotAccessed days not already disabled"
        }
        else
        {
            Write-Output -InputObject "Successfully disabled $actuallyDisabled delivery groups out of $disabled not used in the last $daysNotAccessed days"
        }
    }
}
else
{
    Write-Output -InputObject "Found no delivery groups not used in the last $daysNotAccessed days"
}

