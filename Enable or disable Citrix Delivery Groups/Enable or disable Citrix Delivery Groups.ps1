#requires -version 3

<#
.SYNOPSIS

Enable or disable Citrix Delivery Group(s) specified by name or pattern

.DETAILS

Citrix Studio shows disabled delivery groups but offers no mechanism to change the enabled state

.PARAMETER deliveryGroup

The name or patternn of the delivery group(s) to enable or disable

.PARAMETER disable

Disable delivery groups that have not had a session launched in the number of days specified by -daysNotAccessed

.PARAMETER ddc

The delivery controller to connect to. If not specified the local machine will be used.

.CONTEXT

Computer (but only on a Citrix Delivery Controller)

.NOTES


.MODIFICATION_HISTORY:

    @guyrleech 06/10/2020  Initial release
    @guyrleech 07/10/2020  Added test to stop all delivery groups being disabled
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory,HelpMessage='Name or pattern of the Citrix delivery group(s) to enable/disable')]
    [string]$deliveryGroup ,
    [ValidateSet('true','false')]
    [string]$disable = 'false' ,
    [string]$ddc 
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputwidth = 400

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

## new CVAD versions have modules so use these in preference to snapins which are there for backward compatibility
if( ! (  Import-Module -Name Citrix.DelegatedAdmin.Commands -ErrorAction SilentlyContinue -PassThru -Verbose:$false) `
    -and ! ( Add-PSSnapin -Name Citrix.Broker.Admin.* -ErrorAction SilentlyContinue -PassThru -Verbose:$false) )
{
    Throw 'Failed to load Citrix PowerShell cmdlets - is this a Delivery Controller or have Studio or the PowerShell SDK installed ?'
}

[array]$allDeliveryGroups = @( Get-BrokerDesktopGroup @brokerParameters -ErrorAction SilentlyContinue )
if( ! $allDeliveryGroups -or ! $allDeliveryGroups.Count )
{
    Throw "Retrieved no delivery groups at all"
}

[array]$deliveryGroups = @( $allDeliveryGroups.Where( { $_.Name -like $deliveryGroup } ) )

if( ! $deliveryGroups -or ! $deliveryGroups.Count )
{
    Throw "Found no delivery groups matching `"$deliveryGroup`""
}

[string]$desiredState = $(if( $disable -eq 'true' ) { 'disabled' } else { 'enabled' } )
[bool]$newState = ($disable -eq 'false')

Write-Verbose -Message "Matched $($deliveryGroups.Count) delivery groups out of $($allDeliveryGroups.Count) in total"

[int]$numberAlreadyInDesiredState = $deliveryGroups.Where( { $_.Enabled -eq $newState } ).Count

Write-Verbose -Message "Got $numberAlreadyInDesiredState delivery groups already $desiredState, new enabled state is $newState"

if( $numberAlreadyInDesiredState -ge $deliveryGroups.Count )
{
    Write-Warning -Message "All $numberAlreadyInDesiredState delivery groups matching `"$deliveryGroup`" are already $desiredState"
}
elseif( $desiredState -eq 'disabled' -and $deliveryGroups.Count -eq $allDeliveryGroups.Count )
{
    Throw "This script will not disable all delivery groups which is what is being requested here with all $($deliveryGroups.Count) delivery groups targeted"
}
else
{
    [int]$disabled = 0
    [int]$actuallyDisabled = 0

    $deliveryGroups.Where( { $_.Enabled -ne $newState } ).ForEach(
    {
        $disabled++
        Write-Verbose "$($desiredState -replace 'ed$' , 'ing') `"$($_.Name)`""
        if( ! ( $result = Set-BrokerDesktopGroup -InputObject $_ -Enabled $newState @brokerParameters -PassThru ) -or $result.Enabled -ne $newState )
        {
            Write-Warning -Message "Failed to $desiredState delivery group `"$($_.Name)`""
        }
        else
        {
            $actuallyDisabled++
        }
    })
    if( $disabled -eq 0 )
    {
        Write-Warning -Message "Found no delivery groups not already $desiredState"
    }
    else
    {
        Write-Output -InputObject "Successfully $desiredState $actuallyDisabled delivery groups matching `"$deliveryGroup`" ($numberAlreadyInDesiredState were already $desiredState)"
    }
}
