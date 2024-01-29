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
    @guyrleech 11/12/2023  Added Citrix Cloud support
    @guyrleech 13/12/2023  Added parameter alias -cloudCustomerId
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory,HelpMessage='Name or pattern of the Citrix delivery group(s) to enable/disable')]
    [string]$deliveryGroup ,
    [ValidateSet('true','false','yes','no')]
    [string]$disable = 'false' ,
    [Alias('cloudCustomerId')]
    [string]$ddc
)

#region Controlup_Script_Standards
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

#endregion Controlup_Script_Standards

#region Functions

function Get-StoredCredentials {
    <#
    .SYNOPSIS
        Retrieve the Citrix Cloud Credentials
    .EXAMPLE
        Get-AzSPStoredCredentials
    .CONTEXT
        Azure
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-08-03
        Purpose:        WVD Administration, through REST API calls
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$system ,
        [string]$username ,
        [Alias('customerId')]
        [string]$tenantId
    )

    [string]$credentialsFile = $null
    $strSPCredFolder = [System.IO.Path]::Combine( [environment]::GetFolderPath('CommonApplicationData') , 'ControlUp' , 'ScriptSupport' )
    $credentials = $null

    ## might pass a username for AD domain credentials since could be running as system and may need access to different AD credentials on the same machine
    if( [string]::IsNullOrEmpty( $username ) )
    {
        $username = $env:USERNAME
    }

    Write-Verbose -Message "Get-StoredCredentials $system for user $username"

    if( $system -match 'citrixcloud' )
    {
        [string]$filePattern = [System.IO.Path]::Combine( $strSPCredFolder , "$($username)_*$($System)_Cred.xml" ) ## this will also capture files without customer id in file name
        [string[]]$matchingFiles = @( Get-ChildItem -Path $filePattern -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name )
        
        if( $null -eq $matchingFiles -or $matchingFiles.Count -le 0 )
        {
            Write-Error -Message "No files found matching $filePattern found in $strSPCredFolder"
        }
        else
        {
            [string]$tenantIdRegex = "\b$($username)_(.+)_$($System)_Cred\.xml$"
            [string]$legacyFileName = "$($username)_$($System)_Cred.xml"
            ## see if there are any files with customer id in the name and/or with a specific customer id passed as an argument
            if( [string]::IsNullOrEmpty( $tenantId ) )
            {
                if( $matchingFiles.Count -eq 1 )
                {
                    $credentialsFile = [System.IO.Path]::Combine( $strSPCredFolder , $matchingFiles[ 0 ] )
                    if( $matchingFiles[ 0 ] -match $tenantIdRegex )
                    {
                        $tenantId = $Matches[ 1 ]
                    }
                    ## else hopefully a file name without tenant id in as in a legacy file
                }
                else
                {
                    Write-Error -Message "No Customer Id specified but there are $($matchingFiles.Count) Citrix Cloud credential files for user $username"
                }
            }
            else ## tenant id passed so look for that file specifically if there are more than one otherwise use the single file anyway
            {
                $specificFile = $null
                if( $matchingFiles.Count -gt 1 )
                {
                    $specificFile = "$($username)_$($tenantId)_$($System)_Cred.xml" 
                }
                elseif( $matchingFiles.Count -eq 1 -and $matchingFiles[0] -ieq $legacyFileName )
                {
                    $specificFile = $legacyFileName
                }
                elseif( $matchingFiles.Count -eq 1 -and $matchingFiles[0] -ieq "$($username)_$($tenantId)_$($System)_Cred.xml" )
                {
                    $specificFile = $matchingFiles[0]
                }
                if( $null -ne $specificFile -and $matchingFiles -contains $specificFile )
                {
                    $credentialsFile = [System.IO.Path]::Combine( $strSPCredFolder , $specificFile )
                }
                else
                {
                    Write-Error -Message "No $system credential files for user $username and Citrix Customer Id $tenantId found but there are $($matchingFiles.Count) credential files"
                }
            }
        }
    }
    else ## not Citrix
    {
        $credentialsFile = $(if( -Not [string]::IsNullOrEmpty( $tenantId ) )
        {
            [System.IO.Path]::Combine( $strSPCredFolder , "$($username)_$($tenantId)_$($System)_Cred.xml" )
        }
        else
        {
            [System.IO.Path]::Combine( $strSPCredFolder , "$($username)_$($System)_Cred.xml" )
        })
    }

    Write-Verbose -Message "`tCredentials file is $credentialsFile"

    If ( -Not [string]::IsNullOrEmpty( $credentialsFile ) -and ( Test-Path -Path $credentialsFile) )
    {
        try
        {
            if( ( $credentials = Import-Clixml -Path $credentialsFile ) -and -Not [string]::IsNullOrEmpty( $tenantId ) )
            {
                ## this will also add Citrix customer id if was passed or gleaned from file name
                Add-Member -InputObject $credentials -MemberType NoteProperty -Name TenantId -Value $tenantId -Force
            }
        }
        catch
        {
            Write-Error -Message "The required PSCredential object could not be loaded from $credentialsFile : $_"
        }
    }
    Elseif( $system -eq 'azure' )
    {
        ## try old azure file name 
        $credentials = Get-SPStoredCredentials -system 'az' -tenantId $tenantId 
    }

    if( -not $credentials )
    {
        Write-Error -Message "The $system Credentials file stored for this user ($($env:USERNAME)) cannot be found at $credentialsFile.`nCreate the file with the Set-credentials script action (prerequisite)."
    }
    return $credentials
}

function Get-BearerToken {
    ## https://www.mycugc.org/blogs/eltjo-van-gulik/2019/01/16/blog-monitoring-citrix-cloud-with-odata-and-powers
    ## https://developer.cloud.com/citrix-cloud/citrix-cloud-api-overview/docs/get-started-with-citrix-cloud-apis#bearer_token_tab_oauth_2.0_flow
    param (
        [Parameter(Mandatory=$true)][string]
        $clientId,
        [Parameter(Mandatory=$true)][string]
        $clientSecret
    )
    [string]$bearerToken = $null
    [hashtable]$body = @{
        'grant_type' = 'client_credentials'
        'client_id' = $clientId
        'client_secret' = $clientSecret
    }
    
    $response = $null
    try
    {
        $response = Invoke-RestMethod -Uri 'https://api-us.cloud.com/cctrustoauth2/root/tokens/clients' -Method POST -Body $body
    }
    catch
    {
        Write-Verbose -Message "Get-BearerToken: exception $_"
        if( $_.Exception.Message -imatch 'Unable to connect to the remote server' -or  $_.Exception.Message -imatch 'The operation has timed out' -and -not [string]::IsNullOrEmpty( $script:proxyServer )  )
        {
            $response = Invoke-RestMethod -Uri 'https://api-us.cloud.com/cctrustoauth2/root/tokens/clients' -Method POST -Body $body -Proxy $script:proxyServer -ProxyUseDefaultCredentials
        }
        else
        {
            Throw $_
        }
    }

    if( $null -ne $response )
    {
        $bearerToken = "CwsAuth Bearer=$($response | Select-Object -expandproperty access_token)" 
    }
    ## else will have output error
    $bearerToken ## return    
}

#endregion Functions

[bool]$cloud = $false
[hashtable]$brokerParameters = @{}
[string]$none = 'NONE'
$citrixCloudCredentials = $null
[string[]]$uninstallKeys = @(
    'HKLM:\Software\Wow6432node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' )

[array]$remoteSDKPackages = @( Get-ItemProperty -Path $uninstallKeys -ErrorAction SilentlyContinue | Where-Object { $_.PSObject.Properties[ 'DisplayName' ] -and $_.DisplayName -match '^Citrix .* Remote PowerShell SDK$' -and $_.publisher -match '^Citrix' } )

Write-Verbose -Message "Got $($remoteSDKPackages.Count) remote SDK packages"

## new CVAD versions have modules so use these in preference to snapins which are there for backward compatibility
if( ! (  Import-Module -Name Citrix.DelegatedAdmin.Commands -ErrorAction SilentlyContinue -PassThru -Verbose:$false) `
    -and ! ( Add-PSSnapin -Name Citrix.Broker.Admin.* -ErrorAction SilentlyContinue -PassThru -Verbose:$false) )
{
    Throw 'Failed to load Citrix PowerShell cmdlets - have the Citrix on-prem or Remote PowerShell SDK been installed ? - '
}

$invocationNow = $MyInvocation ## so we can see if -cloudCustomerId alias alias used when calling script

## if Citrix Remote SDK present then we do Cloud or if the -cloudCustomerId parameter alias used
if( ( $null -ne $remoteSDKPackages -and $remoteSDKPackages.Count -gt 0 ) -or ( $invocationNow.Line -match '\s-cloudCustomerId\s' -and -not [string]::IsNullOrEmpty( $ddc ) ) )
{
    $cloud = $true
    $citrixCloudCredentials = $null
    [string]$cloudCustomerId = $ddc
    $authtoken = $null
    ## pass cloud customer id - if not present then get customer id from citrix credential files if there is only 1
    $citrixCloudCredentials = Get-StoredCredentials -system 'CitrixCloud' -tenantId $cloudCustomerId ## if $ddc is null then it will try and find the cloud customer id in the credentials file name
    if( -Not $citrixCloudCredentials )
    {
        exit 1 ## already output error
    }
    if( [string]::IsNullOrEmpty( $ddc ) )
    {
        if( $citrixCloudCredentials.PSObject.Properties[ 'tenantId' ] )
        {
            $cloudCustomerId = $citrixCloudCredentials.tenantId
        }
        else
        {
            Throw "No Citrix Customer Id passed and could not glean from credentials file"
        }
    }
    [Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13
    $authtoken = Get-BearerToken -clientId $citrixCloudCredentials.username -clientSecret $citrixCloudCredentials.GetNetworkCredential().password
    if( [string]::IsNullOrEmpty( $authtoken ) )
    {
        Throw "Authenticattion to Citrix cloud failed"
    }

    Get-XDAuthentication -BearerToken ($authtoken -replace "^CwsAuth Bearer=") -CustomerId $cloudCustomerId

    if( -Not $? )
    {
        Throw "Failed to authenticate to Citrix Cloud for customer $cloudCustomerId"
    }
}
elseif( -Not [string]::IsNullOrEmpty( $ddc ) )
{
    $brokerParameters.Add( 'AdminAddress' , $ddc )
}

[array]$allDeliveryGroups = @( Get-BrokerDesktopGroup @brokerParameters -MaxRecordCount 999999 -ErrorAction SilentlyContinue )
if( ! $allDeliveryGroups -or ! $allDeliveryGroups.Count )
{
    Throw "Retrieved no delivery groups at all"
}

[array]$deliveryGroups = @( $allDeliveryGroups.Where( { $_.Name -like $deliveryGroup } ) )

if( $null -eq $deliveryGroups -or $deliveryGroups.Count -eq 0 )
{
    Throw "Found no delivery groups matching `"$deliveryGroup`""
}

[string]$desiredState = $(if( $disable -ieq 'true' -or $disable -ieq 'yes' ) { 'disabled' } else { 'enabled' } )
[bool]$newState = ($disable -ieq 'false' -or $disable -ieq 'no' )

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
