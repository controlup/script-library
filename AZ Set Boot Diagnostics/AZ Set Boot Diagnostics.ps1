#require -version 3.0

<#
.SYNOPSIS
    Change the state of bootdiagnostics for an Azure VM - when disabled screenshots are not generated

.DESCRIPTION
    Using REST API calls

.PARAMETER azid
    The relative URI of the Azure VM
    
.PARAMETER AZtenantId
    Optional Azure tenant id. Specify when there is a need to access multiple tenants with different credentials.

.PARAMETER  enable
    Whether to enable ("yes") or disable ("no") the boot diagnostics

.NOTES
    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2021-10-13
    Updated:        2022-02-22 Guy Leech  Look for _AZ_ credentials file if _Azure_ not found. Checks on AZ id and tenant id validity
#>

[CmdletBinding()]

Param
(
    [string]$AZid , ## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId ,
    [ValidateSet('Yes','No')]
    [string]$enable = 'Yes'
)    

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 250
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

[string]$apiversion = '2021-04-01'
[string]$computeApiVersion = '2021-07-01'
[string]$baseURL = 'https://management.azure.com'
[string]$credentialType = 'Azure'

Write-Verbose -Message "AZid is $AZid"

#region AzureFunctions

function Get-AzSPStoredCredentials {
    <#
    .SYNOPSIS
        Retrieve the Azure Service Principal Stored Credentials.
    .DESCRIPTION
        Retrieve the Azure Service Principal Stored Credentials from a stored credentials file.
    .EXAMPLE
        Get-AzSPStoredCredentials
    .CONTEXT
        Azure
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-08-03
        Purpose:        WVD Administration, through REST API calls
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$system ,
        [string]$tenantId
    )

    $strAzSPCredFolder = [System.IO.Path]::Combine( [environment]::GetFolderPath('CommonApplicationData') , 'ControlUp' , 'ScriptSupport' )
    $AzSPCredentials = $null

    Write-Verbose -Message "Get-AzSPStoredCredentials $system"

    [string]$credentialsFile = $(if( -Not [string]::IsNullOrEmpty( $tenantId ) )
        {
            [System.IO.Path]::Combine( $strAzSPCredFolder , "$($env:USERNAME)_$($tenantId)_$($System)_Cred.xml" )
        }
        else
        {
            [System.IO.Path]::Combine( $strAzSPCredFolder , "$($env:USERNAME)_$($System)_Cred.xml" )
        })

    Write-Verbose -Message "`tCredentials file is $credentialsFile"

    If (Test-Path -Path $credentialsFile)
    {
        try 
        {
            if( ( $AzSPCredentials = Import-Clixml -Path $credentialsFile ) -and -Not [string]::IsNullOrEmpty( $tenantId ) -and -Not $AzSPCredentials.ContainsKey( 'tenantid' ) )
            {
                $AzSPCredentials.Add(  'tenantID' , $tenantId )
            }
        }
        catch 
        {
            Write-Error -Message "The required PSCredential object could not be loaded from $credentialsFile : $_"
        }
    }
    Elseif( $system -eq 'Azure' )
    {
        ## try old Azure file name 
        $azSPCredentials = Get-AzSPStoredCredentials -system 'AZ' -tenantId $tenantId 
    }

    if( -not $AzSPCredentials )
    {
        Write-Error -Message "The Azure Service Principal Credentials file stored for this user ($($env:USERNAME)) cannot be found at $credentialsFile.`nCreate the file with the Set-AzSPCredentials script action (prerequisite)."
    }
    return $AzSPCredentials
}

function Get-AzBearerToken {
    <#
    .SYNOPSIS
        Retrieve the Azure Bearer Token for an authentication session.
    .DESCRIPTION
        Retrieve the Azure Bearer Token for an authentication session, using a REST API call.
    .EXAMPLE
        Get-AzBearerToken -SPCredentials <PSCredentialObject> -TenantID <string>
    .CONTEXT
        Azure
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-03-22
        Updated:        2020-05-08
                        Created a separate Azure Credentials function to support ARM architecture and REST API scripted actions
        Purpose:        WVD Administration, through REST API calls
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='Azure Service Principal credentials' )]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCredential] $SPCredentials,

        [Parameter(Mandatory=$true, HelpMessage='Azure Tenant ID' )]
        [ValidateNotNullOrEmpty()]
        [string] $TenantID
    )

    # URL for REST API call to authenticate with Azure (using the TenantID parameter)
    [string]$uri = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
        
    # Create the Invoke-RestMethod Body (using the SPCredentials parameter)
    [hashtable]$body = @{
        grant_type    = 'client_credentials'
        client_Id     = $SPCredentials.UserName
        client_Secret = $SPCredentials.GetNetworkCredential().Password
        scope         = "$baseURL/.default"
    }

    # Create the Invoke-RestMethod parameters
    [hashtable]$invokeRestMethodParams = @{
        Uri             = $uri
        Body            = $body
        Method          = 'POST'
        ContentType     = 'application/x-www-form-urlencoded'
    }

    Invoke-RestMethod @invokeRestMethodParams | Select-Object -ExpandProperty access_token -ErrorAction SilentlyContinue
}

function Invoke-AzureRestMethod {
   
    [CmdletBinding()]
    Param(
        [Parameter( Mandatory=$true, HelpMessage='A valid Azure bearer token' )]
        [ValidateNotNullOrEmpty()]
        [string]$BearerToken ,
        [string]$uri ,
        [ValidateSet('GET','POST','PUT','PATCH')] ## add others as necessary
        [string]$method = 'GET' ,
        [hashtable]$body ,
        [string]$property = 'value'
    )

    [hashtable]$header = @{
        'Authorization' = "Bearer $BearerToken"
        'Content-Type'  = 'application/json'
    }

    [hashtable]$invokeRestMethodParams = @{
        Uri             = $uri
        Method          = $method
        Headers         = $header
    }

    if( $PSBoundParameters[ 'body' ] )
    {
        $invokeRestMethodParams.Add( 'Body' , ( $body | ConvertTo-Json -Depth 10 )) ## -Depth is to ensure that all nested hashtables are converted
    }
    
    if( -not [String]::IsNullOrEmpty( $property ) )
    {
        Invoke-RestMethod @invokeRestMethodParams | Select-Object -ErrorAction SilentlyContinue -ExpandProperty $property
    }
    else
    {
        Invoke-RestMethod @invokeRestMethodParams ## don't pipe through select as will slow script down for large result sets if processed again after rreturn
    }
}

#endregion AzureFunctions

[string]$vmName = ($AZid -split '/')[-1]
    
if( [string]::IsNullOrEmpty( $vmName ) )
{
    Throw "Azure id `"$AZid`" does not appear valid - failed to find VM name"
}

if( -Not [string]::IsNullOrEmpty( $AZtenantId ) -and -Not ( $AZtenantId -as [guid] ) )
{
    Throw "Azure tenant id `"$AZtenantId`" is invalid"
}

If ($azSPCredentials = Get-AzSPStoredCredentials -system $credentialType -tenantId $AZtenantId )
{
    # Sign in to Azure with a Service Principal with Contributor Role at Subscription level and retrieve the bearer token
    Write-Verbose -Message "Authenticating to tenant $($azSPCredentials.tenantID) as $($azSPCredentials.spCreds.Username)"
    if( -Not ( $azBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID ) )
    {
        Throw "Failed to get Azure bearer token"
    }
        
    ## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/retrieve-boot-diagnostics-data
    try
    {
        $bootDiagnostics = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid/retrieveBootDiagnosticsData?api-version=$computeApiVersion" -property $null -method POST
    }
    catch
    {
        $bootDiagnostics = $null
    }

    ## if $bootDiagnostics is null then they are not enabled
    $state = $null
    [string]$stateText = $null

    if( $enable -eq 'yes' )
    {
        if( $bootDiagnostics )
        {
            Write-Warning -Message "Boot diagnostics are already enabled in VM $vmName"
        }
        else
        {
            $state = $true
            $stateText = 'enable'
        }
    }
    else ## disable
    {
        if( -Not $bootDiagnostics )
        {
            Write-Warning -Message "Boot diagnostics are already disabled in VM $vmName"
        }
        else
        {
            $state = $false
            $stateText = 'disable'
        }
    }

    if( $state -ne $null )
    {
        ## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/create-or-update
        [hashtable]$body = @{
                'properties' = @{
                    'diagnosticsProfile' = @{
                        'bootDiagnostics' = @{
                            'enabled' = $state
                        }
                 }
            }
        }
        if( -Not ( $reconfigured = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid`?api-version=$computeApiVersion" -property $null -body $body -method PATCH ) )
        {
            Write-Error -Message "Error when trying to $stateText boot diagnostics"
        }
        elseif( $reconfigured.properties.diagnosticsProfile.bootDiagnostics.enabled -eq $state )
        {
            Write-Output -InputObject "Operation to $stateText boot diagnostics succeeded"
        }
        else
        {
            Write-Warning -Message "No error from call to $stateText boot diagnostics but state appears not to have not changed"
        }
    }
}

