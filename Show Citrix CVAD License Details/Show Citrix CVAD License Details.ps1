<#
.SYNOPSIS
    Get Citrix licence information

.NOTES
    Modification History:

    2023/04/12  Guy Leech  Scripted started
    2023/04/14  Guy Leech  Initial release
#>

[CmdletBinding()]

Param
(
    [int]$daysToExpiry = 60 ,
    [switch]$raw
)

[hashtable]$licenceType = @{
    'CCS' = 'Concurrent'
    'UD'  = 'User/Device'
}

#region ControlUpScriptingStandards
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 400
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}
#endregion ControlUpScriptingStandards

Function Out-PassThru
{
    Process
    {
        $_
    }
}

Import-Module -Name Citrix.Licensing.Admin.V1,Citrix.Configuration.Commands

$site = $null
$site = Get-ConfigSite
if( $null -eq $site )
{
    Throw "Failed to get Citrix site information"
}

If( -Not ( Get-Command -Name Get-LicInventory -ErrorAction SilentlyContinue) )
{
    Throw "Cannot find Citrix cmdlet Get-LicInventory"
}

if( -not $raw )
{
    Write-Output -InputObject "Site `"$($site.SiteName)`" licence server is $($site.LicenseServerName):$($site.LicenseServerPort), licensing model is $($site.LicensingModel)"
}

$licenceInventory = $null
$licenceInventory = Get-LicInventory -AdminAddress $site.LicenseServerUri

if( $null -eq $licenceInventory )
{
    Throw "Failed to get licence inventory from $($site.LicenseServerUri)"
}

[array]$licences = @( $licenceInventory | Where-Object LicenseType -ine 'SYS' | Select-Object -Property * -ExcludeProperty *localized* )

if( $null -eq $licences -or $licences.Count -eq 0 )
{
    Write-Warning -Message "No licences found"
}
else
{
    [datetime]$warningDate = [datetime]::Now.AddDays( $daysToExpiry )
    ## Remove "License" from all property names
    $processedLicences = @( ForEach( $licence in $licences )
    {
        [hashtable]$newLicense = @{}
        ForEach( $property in $licence.PSObject.Properties )
        {
            $newLicense.Add( ( $property.Name -replace '^Licenses?' ) , $property.Value )
        }
        [int]$alerts = 0
        if( $licence.LicensesInUse -ge $licence.LicensesAvailable -or $licence.LicenseOverdraft -gt 0 )
        {
            $alerts++
        }
        if( $licence.LicenseExpirationDate -le $warningDate )
        {
            $alerts++
        }
        $newLicense.Add( 'Alerts' , $alerts )
        [pscustomobject]$newLicense
    })
    
    $outputCommand = Get-Command -Name Format-Table
    [hashtable]$outputArguments = @{ AutoSize  = $true ; Wrap = $true }
    if( $raw )
    {
        $outputCommand = Get-Command -Name Out-PassThru
        $outputArguments = @{}
    }

    $processedLicences | Sort-Object -property ExpirationDate | Select-Object -Property `
        @{n='Alert';e={ '*' * $_.alerts }},
        @{n='Expires';e={$_.ExpirationDate.ToString('D')}},
        @{n='Subscription Advantage';e={$_.ubscriptionAdvantageDate.ToString('D')}}, ## regex has removed the 's' of 'ubscriptionAdvantageDate'
        @{n='Model';e={if( $type = $licenceType[ $_.Model ] ) { $type } else { $_.Model }}},
        Type,ProductName,Edition,InUse,Available,Overdraft -ExcludeProperty Model,Alerts | . $outputCommand @outputArguments
}

