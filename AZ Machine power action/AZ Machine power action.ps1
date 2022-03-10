#require -version 3.0

<#
.SYNOPSIS
    Perform an action on the Azure VM passed

.DESCRIPTION
    Using REST API calls
    
.PARAMETER azid
    The Azure id of the VM to perform the action on

.PARAMETER action
    The action to perform on the Azure VM
    
.PARAMETER maxWaitTimeSeconds
    The maximum number of seconds to wait for the action to complete. If not specified or less than or equal to zero, no waiting will be done

.PARAMETER sleepMilliseconds
    The period to sleep for in milliseconds between calls to get the status of the operation

.PARAMETER AZtenantId
    The azure tenant ID

.EXAMPLE
    & '.\AZ VM action.ps1' -AZid /subscriptions/58ffa3cb-4242-4f2e-a06d-deadbeefdead/resourceGroups/WVD/providers/Microsoft.Compute/virtualMachines/GLW10WVD-0 -action deallocate -maxWaitTimeSeconds 60

    Stop and deallocate the VM GLW10WVD-0 in resource group WVD using the saved credentials for the user running the script

.NOTES
    Saved credentials for the user running the script must be available in the file "C:\ProgramData\ControlUp\ScriptSupport\%username%_AZ_Cred.xml" - there is a ControlUp script to create them

    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2021-11-25
    Updated:        2022-01-17  Guy Leech    Fix for tenant id handling
                    2022-01-19  Guy Leech    Change to OAuth v2. Added confirmation option
#>

[CmdletBinding()]

Param
(
    [string]$AZid , ## passed by CU as the URL to the VM minus the FQDN ,
    [string]$AZtenantId ,
    [ValidateSet('start','stop','shutdown','turnoff','restart','deallocate','hibernate','redeploy','delete')]
    [string]$action ,
    [ValidateSet('Yes','No','True','False')]
    [string]$confirmAction = 'No' ,
    [int]$maxWaitTimeSeconds = 0 ,
    [int]$sleepMilliseconds = 2500
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

## mandatory parameters best avoided in CU scripts as can cause scripts to hang if missing since willbe promoptinng, siliently, for missing parameters
if( [string]::IsNullOrEmpty( $AZid ) )
{
    Throw "Missing Azure id parameter"
}

if( [string]::IsNullOrEmpty( $action ) )
{
    Throw "Missing action parameter"
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
        Retrieve the Azure Bearer Token for an authentication session
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
        [ValidateSet('GET','POST','PUT','DELETE','PATCH')] ## add others as necessary
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

    $responseHeaders = $null

    if( $PSVersionTable.PSVersion -ge [version]'7.0.0.0' )
    {
        $invokeRestMethodParams.Add( 'ResponseHeadersVariable' , 'responseHeaders' )
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

[hashtable]$operationMappings = @{
    'shutdown' = 'poweroff'
    'stop' = 'deallocate'
    'turnoff' = 'poweroff'
    'hibernate' = 'deallocate'
}

[hashtable]$parameterMappings = @{
    'turnoff' = 'skipShutdown=true&'
    'hibernate' = 'hibernate=true&'
}

If ($azSPCredentials = Get-AzSPStoredCredentials -system $credentialType -tenantId $AZtenantId )
{
    # Sign in to Azure with a Service Principal with Contributor Role at Subscription level and retrieve the bearer token
    Write-Verbose -Message "Authenticating to tenant $($azSPCredentials.tenantID) as $($azSPCredentials.spCreds.Username)"
    if( -Not ( $azBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID ) )
    {
        Throw "Failed to get Azure bearer token"
    }

    [string]$vmName = ($AZid -split '/')[-1]
    
    ## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/instance-view
    [string]$instanceViewURI = "$baseURL/$azid/instanceView`?api-version=$computeApiVersion"
    if( $null -eq ( [array]$virtualMachineStatuses = @( Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $instanceViewURI -property 'statuses' ) ) `
        -or $virtualMachineStatuses.Count -eq 0 )
    {
        Throw "Failed to get VM instance view via $instanceViewURI : $_"
    }

    ## code property will be an array so cannot use $matches (https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comparison_operators?view=powershell-5.1)
    ##   ProvisioningState/updating
    ##   PowerState/running

    if( ( $line = ( $virtualMachineStatuses.code -match 'ProvisioningState/' )) -and ( $status = ($line -split '/' , 2 )[-1] ) )
    {
        ## check not performing an operation already
        Write-Verbose -Message "Current VM provisioning state is $status"
        if( $status -eq 'Updating' )
        {
            Throw "VM is already performing an operation : $($virtualMachineStatuses | Format-Table | Out-String)"
        }
    }   

    if( ( $line = ( $virtualMachineStatuses.code -match 'PowerState/' )) -and ( $powerstate = ($line -split '/' , 2 )[-1] ))
    {
        ## check not already in requested or similar state
        Write-Verbose -Message "Current VM status is $powerstate"

        if( ( $powerstate -eq 'deallocated' -and $action -in @( 'deallocate' , 'stop' , 'turnoff' , 'restart' , 'hibernate' ) ) `
            -or ( $powerstate -eq 'running' -and $action -eq 'start' ) `
            -or ( $powerstate -eq 'stopped' -and $action -in @( 'stop' , 'shutdown' , 'turnoff' , 'hibernate' )))
        {
            Throw "VM is already $powerstate so cannot $action"
        }
    }

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
            $answer = [Windows.MessageBox]::Show( "Are you sure you want to $action $vmName" , "Confirm Run Operation" , 'YesNo' ,'Question' )
            if( $answer -ine 'Yes' )
            {
                Write-Warning -Message "Action not confirmed by user - aborting"
                Exit 1
            }
        }
    }

    if( -Not ( $operation = $operationMappings[ $action ] ) )
    {
        $operation = $action
    }

    $parameters = $parameterMappings[ $action ]

    ## doesn't return anything
    try
    {
        [string]$operationURI = "$baseURL/$azid/$operation`?$($parameters)api-version=$computeApiVersion"
        [string]$status = 'unknown'

        Write-Verbose -Message "Performing action $operationURI"

        Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $operationURI -property $null -method POST

        if( $maxWaitTimeSeconds -gt 0 )
        {
            [datetime]$start = [datetime]::Now
            [datetime]$end = $start.AddSeconds( $maxWaitTimeSeconds )

            do
            {
                if( ( $state = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri $instanceViewURI -property 'statuses' ) )
                {
                    if( ( $line = ( $state.code -match 'ProvisioningState/' )) -and ( $status = ($line -split '/')[-1]) )
                    {
                        Write-Verbose -Message "`tCurrent VM provisioning state is $status"
                        if( $status -eq 'Succeeded' )
                        {
                            break
                        }
                        elseif( $status -eq 'Failed' )
                        {
                            Write-Error -Message "Provisioning failed for $vmName"
                            break
                        }
                        elseif( $status -ne 'Updating' -and $status -ne 'True' )
                        {
                            Write-Warning -Message "Unexpected provisioning state `"$status`""
                        }
                    } 
                    else
                    {
                        Write-Warning -Message "Failed call to get provisoning state of VM $vmName : $($state.code)"
                    }
                }
                else
                {
                    Write-Warning -Message "Failed call to get state of VM $vmName"
                }
                Write-Verbose -Message "$(Get-Date -Format G) : provisioning state of $vmName is $($state.code) so waiting $sleepMilliseconds ms"
                Start-Sleep -Milliseconds $sleepMilliseconds
            } while( [datetime]::Now -le $end )

            [datetime]$timeNow = [datetime]::Now

            if( -Not $state -or $status -ne 'Succeeded' )
            {
                Write-Warning -Message "VM provisioning status still $status after $([int](($timeNow - $start).TotalSeconds)) seconds"
            }
            else
            {
                Write-Output -InputObject "$action on $vmName succeeded after $([int](($timeNow - $start).TotalSeconds)) seconds"
            }
        }
        else
        {
            Write-Output -InputObject "$action on $vmName submitted ok"
        }
    }
    catch
    {
        Write-Error -Message "Error from request : $_"
    }
}

