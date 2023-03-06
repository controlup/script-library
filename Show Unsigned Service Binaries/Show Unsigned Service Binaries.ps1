#requires -version 3

<#
.SYNOPSIS
    Get the exe or dll used in every service, show details and check signature

.PARAMETER showAllSigning
    Show services where the executable has any signing state.
    The default is to only show those where signing is not valid

.PARAMETER showNonMicrosoftOnly
    Show services not from Microsoft.
    The default is to show those from any vendor including Microsoft.
    Setting this to 'yes' and passing -showAllSigning 'yes' which show all services not from Microsoft

.NOTES
    Modification History:

    2023/19/01  @guyrleech  Initial release
#>

[CmdletBinding()]

Param
(
    [ValidateSet('Yes','No')]
    [string]$showAllSigning = 'No' ,
    [ValidateSet('Yes','No')]
    [string]$showNonMicrosoftOnly = 'No'
)

#Region ControlUp_Standards
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 400
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    try
    {
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions
    }
    catch
    {
        ## Nothing we can do but shouldn't cause script to end
    }
}
#EndRegion ControlUp_Standards

[array]$services = @( Get-Service )
[string]$baseServicesKey = 'HKLM:SYSTEM\CurrentControlSet\Services'

[array]$results = @( Get-CimInstance -ClassName win32_service | ForEach-Object `
{ 
    $service = $_
    $executableFile = $null
    [string]$serviceType = 'OwnProcess'
    [string]$serviceName = $service.Name
    ## win32_service doesn't have the service type that tells us whether it is a user service or not
    if( $userService = $services | Where-Object { $_.Name -eq $service.Name -and ($_.ServiceType -as [int]) -band 0xC0 }  )
    {
        $serviceName = $userService.name -replace '_([0-9a-f]+)$'
    }

    if( $service.pathname -match '\\svchost\.exe\b' )
    {
        $serviceType = 'SharedProcess'
        if( -Not ( $executableFile = Get-ItemProperty -Path "$BaseServicesKey\$($servicename)\Parameters" -name ServiceDll -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Servicedll ) `
            -and -Not ( $executableFile = Get-ItemProperty -Path "$BaseServicesKey\$($servicename)" -name ServiceDll -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Servicedll ))
        {
            Write-Warning -Message "Could not find a ServiceDll registry value for svchost based service $($service.Name) ($($service.DisplayName))"
        }
        elseif( $userservice )
        {
            $serviceType = 'UserSharedProcess'
        }
    }
    ## get exe in quotes, eg has spaces in path and then arguments, or no spaces in path but possibly space following if arguments
    elseif( $service.PathName -match  '^"([^"]+)"' -or $service.PathName -match '^([^\s]+)' )
    {
        $executableFile = $Matches[1]
        if( $userservice )
        {
            $serviceType = 'UserOwnProcess'
        }
    }
    else
    {
        Write-Warning -Message "Couldn't get executable from $($_.Pathname) for service ($_.Name) ($($_.DisplayName))"
    }

    $versionInfo = $null

    if( $executableFile )
    {
        $versionInfo = Get-ItemProperty -Path $executableFile -ErrorAction SilentlyContinue | Select-Object -ExpandProperty versioninfo -ErrorAction SilentlyContinue
    }

    Select-Object -InputObject $service -Property @(
        @{ n='Name'      ; e={$serviceName}} , 
        'Displayname' ,
        'StartMode' ,
        @{ n='Account'   ; e={$_.StartName }} ,
        'State' ,
        ##'Pathname' ,
        @{ n = 'Type'    ; e = { $serviceType }} ,
        @{ n = 'Binary'  ; e = { $executableFile }},
        ## some VMware executables don't populate companyname :-(
        @{ n = 'Vendor'  ; e = { if( $versionInfo -and [string]::IsNullOrEmpty( $versionInfo.CompanyName )) { $versionInfo.LegalCopyright } else { Select-Object -InputObject $versionInfo -ExpandProperty companyname -ErrorAction SilentlyContinue }}},
        @{ n = 'Signing' ; e = { Get-AuthenticodeSignature -FilePath $executableFile -ErrorAction SilentlyContinue | Select -ExpandProperty Status }}  ,
        @{ n = 'Delete Flag' ; e = { Get-ItemProperty -Path "$BaseServicesKey\$($servicename)" -name DeleteFlag -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DeleteFlag }})
})

$results | Where-Object { ( $showAllSigning -ieq 'yes' -or $_.Signing -ine 'valid' ) -and ($showNonMicrosoftOnly -ieq 'no' -or ( $showNonMicrosoftOnly -ieq 'yes' -and ( -Not $_.Vendor -or$_.Vendor -notmatch 'Microsoft' ))) } | Sort-Object -property DisplayName | Format-Table -AutoSize

