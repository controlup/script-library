<#
.SYNOPSIS
    Get and show details of RAS sessions

.DESCRIPTION
    AD credentials for a RAS admin must previously have been created on the machine where this script runs as the account running the script

.NOTES
    Modification History:

    2023/11/13  Guy Leech  Script born
#>

[CmdletBinding()]

Param
(
    [string]$sessionState = 'All' ,
    [string]$sortBy = 'User' ,
    [int]$authRetrties = 2 ,
    [string]$computer
)

#region ControlUp_Script_Standards
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
#endregion ControlUp_Script_Standards

[array]$sessionProperties = @( 'User' , @{ name = 'Machine' ; expression = { if( [string]::IsNullOrEmpty( $_.SessionHostName )) { $_.vmname } else { $_.SessionHostName } }} , 'PoolName' , 'Source' , 'LogonTime' , 'State'  )
$agents = $null
$ADcredential = $null

#region Functions
Function Get-StoredCredentials {
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
        [string]$tenantId
    )

    $strSPCredFolder = [System.IO.Path]::Combine( [environment]::GetFolderPath('CommonApplicationData') , 'ControlUp' , 'ScriptSupport' )
    $credentials = $null

    ## might pass a username for AD domain credentials since could be running as system and may need access to different AD credentials on the same machine
    if( [string]::IsNullOrEmpty( $username ) )
    {
        $username = $env:USERNAME
    }

    Write-Verbose -Message "Get-StoredCredentials $system for user $username"

    [string]$credentialsFile = $(if( -Not [string]::IsNullOrEmpty( $tenantId ) )
        {
            [System.IO.Path]::Combine( $strSPCredFolder , "$($username)_$($tenantId)_$($System)_Cred.xml" )
        }
        else
        {
            [System.IO.Path]::Combine( $strSPCredFolder , "$($username)_$($System)_Cred.xml" )
        })

    Write-Verbose -Message "`tCredentials file is $credentialsFile"

    If (Test-Path -Path $credentialsFile)
    {
        try
        {
            if( ( $credentials = Import-Clixml -Path $credentialsFile ) -and -Not [string]::IsNullOrEmpty( $tenantId ) -and -Not $credentials.ContainsKey( 'tenantid' ) )
            {
                $credentials.Add(  'tenantID' , $tenantId )
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
#endregion Functions

## RAS auth fails if domain name passed
$ADcredential = Get-StoredCredentials -system ADDomain -username ( $env:username -replace '^.*\\' ) ## currently only support a single domain but could put domain name into file name although file contains domain\user

if( $null -eq $ADcredential )
{
    exit 1 ## already given error message
}

Import-Module -Name RASAdmin -Verbose:$false -Debug:$false

[hashtable]$RASoptions = @{  
}

if( -Not [string]::IsNullOrEmpty( $computer ) -and $computer -ine 'localhost' -and $computer -ne '.' )
{
    $RASoptions.Add( 'Server' , $computer )
}

New-RASSession -Username ($ADCredential.UserName -replace '^.*\\') -Password $ADCredential.Password -Retries $authRetrties -Force @RASoptions

if( -Not $? )
{
    Throw "Failed to connect to RAS on local machine as $($ADCredential.username)"
}

[array]$sessions = @( Get-RASRDSession -State $sessionState )

if( $null -eq $sessions -or $sessions.Count -eq 0 )
{
    Write-Output "No RAS sessions found in state `"$sessionState`""
}
else
{
    Write-Output "Found $($sessions.Count) sessions in state `"$sessionState`""

    $sessions | Sort-Object -Property $sortBy | Format-Table -AutoSize -Property $sessionProperties
}

Remove-RASSession

