<#
.SYNOPSIS
    Get and show details of RAS agents

.DESCRIPTION
    AD credentials for a RAS admin must previously have been created on the machine where this script runs as the account running the script

.NOTES
    Modification History:

    2023/11/06  Guy Leech  Script born
#>

[CmdletBinding()]

Param
(
    [string]$serverType = 'All' ,
    [ValidateSet('Yes','No')]
    [string]$nonOKonly = 'no' ,
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

## TODO do we put Get-RASBroker in and exit if not primary as subsequent RAS cmdlets fail?

$agentError = $null
$agents = @( Get-RASAgent -ServerType $serverType -ErrorAction SilentlyContinue -ErrorVariable agentError | Where-Object { $nonOKonly -ieq 'no' -or  $_.AgentState -ne 'OK' } )

if( $null -eq $agents -or $agents.Count -eq 0 )
{
    Throw "No RAS agents found of type `"$serverType`" - $agentError"
}

Write-Verbose -Message "Got $($agents.Count) RAS agents"

$agentStatuses | Format-Table -AutoSize 

$agents | Select-Object -Property @{name = 'Name';expression={ if( -Not [string]::IsNullOrEmpty( $_.Name )) { $_.Name } else { $_.Server} }},@{name = 'Version';expression={ $_.AgentVer}},@{name = 'Type';expression={ $_.servertype}},AgentState | Sort-Object -Property type | Format-Table -AutoSize

[array]$versions = @( $agents | Where-Object { $_.PSObject.Properties[ 'agentver' ] -and -Not [string]::IsNullOrEmpty( $_.agentver ) } | Group-Object -Property agentver )

if( $null -ne $versions -and $versions.Count -gt 1 )
{
    Write-Warning -Message "There are $($versions.Count) different RAS agent versions in use"
}

Remove-RASSession

