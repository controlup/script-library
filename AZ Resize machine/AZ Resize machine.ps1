#require -version 3.0

<#
.SYNOPSIS
    Change the size of the Azure VM passed

.DESCRIPTION
    Using REST API calls

.PARAMETER azid
    The relative URI of the Azure VM
    
.PARAMETER newSize
    The SKU to change the VM to use

.NOTES
    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2021-11-09
    Updated:
#>

[CmdletBinding()]

Param
(
    [string]$AZid ,## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId ,
    [ValidateSet('Yes','No','True','False')]
    [string]$confirmAction = 'No' , 
    [string]$resizeTo
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

[string]$computeApiVersion = '2021-07-01'
[string]$insightsApiVersion = '2015-04-01'
[string]$baseURL = 'https://management.azure.com'
[string]$credentialType = 'Azure'

Write-Verbose -Message "AZid is $AZid"

#region AzureFunctions
function Get-AzSPStoredCredentials {
    <#
    .SYNOPSIS
        Retrieve the Azure Service Principal Stored Credentials
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
            if( ( $AzSPCredentials = Import-Clixml -Path $credentialsFile ) -and -Not [string]::IsNullOrEmpty( $tenantId )  -and -Not $AzSPCredentials.ContainsKey( 'tenantid' ) )
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
        $azSPCredentials = Get-AzSPStoredCredentials -system 'AZ' -tenantId $AZtenantId 
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
    
    ## https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
    [string]$uri = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"

    [hashtable]$body = @{
        grant_type    = 'client_credentials'
        client_Id     = $SPCredentials.UserName
        client_Secret = $SPCredentials.GetNetworkCredential().Password
        scope         = "$baseURL/.default"
    }

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
        [ValidateSet('GET','POST','PUT','DELETE')] ## add others as necessary
        [string]$method = 'GET' ,
        $body , ## not typed because could be hashtable or pscustomobject
        [string]$property = 'value' ,
        [string]$contentType = 'application/json'
    )

    [hashtable]$header = @{
        'Authorization' = "Bearer $BearerToken"
    }

    if( ! [string]::IsNullOrEmpty( $contentType ) )
    {
        $header.Add( 'Content-Type'  , $contentType )
    }

    [hashtable]$invokeRestMethodParams = @{
        Uri             = $uri
        Method          = $method
        Headers         = $header
    }

    if( $PSBoundParameters[ 'body' ] )
    {
        $invokeRestMethodParams.Add( 'Body' , ( $body | ConvertTo-Json -Depth 20))
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

If ($azSPCredentials = Get-AzSPStoredCredentials -system $credentialType -tenantId $AZtenantId )
{
    # Sign in to Azure with a Service Principal with Contributor Role at Subscription level and retrieve the bearer token
    Write-Verbose -Message "Authenticating to tenant $($azSPCredentials.tenantID) as $($azSPCredentials.spCreds.Username)"
    if( -Not ( $azBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID ) )
    {
        Throw "Failed to get Azure bearer token"
    }

    [string]$vmName = ($AZid -split '/')[-1]
    
    [string]$subscriptionId = $null
    [string]$resourceGroupName = $null

    ## subscriptions/58ffa3cb-2f63-4f2e-a06d-369c1fcebbf5/resourceGroups/WVD/providers/Microsoft.Compute/virtualMachines/GLMW10WVD-0
    if( $AZid -match '\bsubscriptions/([^/]+)/resourceGroups/([^/]+)/providers/Microsoft\.' )
    {
        $subscriptionId = $Matches[1]
        $resourceGroupName = $Matches[2]
    }
    else
    {
        Throw "Failed to parse subscription id and resource group from $AZid"
    }

    if( -Not ( $virtualMachine = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid/?api-version=$computeApiVersion" -property $null ) )
    {
        Throw "Failed to get VM for $azid"
    }

    if( $availableSizes = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid/vmSizes?api-version=$computeApiVersion" )
    {
        Write-Verbose -Message "$($availableSizes.Count) available VM sizes for $($virtualMachine.Name), current size is $($virtualMachine.properties.hardwareProfile.vmSize)"
        if( -Not ( $newSize = $availableSizes.Where( { $_.name -match $resizeTo } ) ) )
        {
            Throw "Unable to find VM size $resizeTo"
        }
        elseif( $newSize.Count -gt 1 )
        {
            Throw -Message "Found $($newSize.Count) VM sizes matching $resizeTo - $(($newSize | Select-Object -ExpandProperty name) -join ',')"
        }
        elseif( $virtualMachine.properties.hardwareProfile.vmSize -eq $newSize.name )
        {
            Throw "VM $($virtualMachine.name) is already size $($virtualMachine.properties.hardwareProfile.vmSize)"
        }
        else
        {
            if( $confirmAction -ieq 'yes' -or $confirmAction -ieq 'true' )
            {
                ## if we are in session 0 then do not prompt
                [int]$thisSession = Get-Process -Id $pid -ErrorAction SilentlyContinue | Select-Object -ExpandProperty SessionId
                if( $thisSession -eq 0 )
                {
                    Write-Warning -Message "Action not confirmed by user - runnning in session 0 so unable to prompt"
                    Exit 2
                }
                else
                {
                    Add-Type -AssemblyName PresentationCore,PresentationFramework
                    ## script must be running on console for this to work otherwise will hang indefinitely
                    $answer = [Windows.MessageBox]::Show( "Are you sure you want to change $vmName from $($virtualMachine.properties.hardwareProfile.vmSize) to $($newSize.Name)" , "Confirm Resize Operation" , 'YesNo' ,'Question' )
                    if( $answer -ine 'Yes' )
                    {
                        Write-Warning -Message "Action not confirmed by user - aborting"
                        Exit 1
                    }
                }
            }
            ## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/create-or-update
            [hashtable]$body = @{
                'location' = $virtualMachine.location
                'properties' = @{
                    'hardwareprofile' = @{
                        'vmSize' = $newSize.Name
                    }
                }
            }
            if( -Not ( $resized = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid/?api-version=$computeApiVersion" -property $null -body $body -method PUT ) )
            {
                Write-Error -Message "Error when changing $($virtualMachine.Name) from $($virtualMachine.properties.hardwareProfile.vmSize) to $($newSize.name)"
            }
            elseif( $resized.properties.hardwareProfile.vmSize -ne $newSize.name )
            {
                Write-Error -Message "Failed to change $($virtualMachine.Name) from $($virtualMachine.properties.hardwareProfile.vmSize) to $($newSize.name), it is now $($resized.properties.hardwareProfile.vmSize)"
            }
            else
            {
                Write-Output -InputObject "Changed from $($virtualMachine.properties.hardwareProfile.vmSize) to $($resized.properties.hardwareProfile.vmSize) ok"
            }
        }
    }
    else
    {
        Write-Warning -Message "Failed to get available VM sizes for $vmname"
    }
}
