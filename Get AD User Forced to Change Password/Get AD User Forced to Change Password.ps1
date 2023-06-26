<#
.SYNOPSIS
    Lists all users in Active Directory with "User must change password at next logon" selected.

.DESCRIPTION
    Lists all users in Active Directory with "User must change password at next logon" selected.
    Not getting any results means there are no user accounts with the setting active within the given parameters.

    This script is intended to be used within ControlUp as an action. If the .NET engine is used all optional parameter are truly optional.
    The classic engine cannot handle empty optional parameters. If the .NET engine is not used pass "none" to leave an optional parameter empty.

.EXAMPLE
    Parameters:
        DaysBeforeAccountExpiration = 30
        SearchBases = none
        Users = none
        IncludeDisabledUserAccounts = none

    Running the ControlUp Action using the above parameters retrieves all active user accounts that are due to expire within the next 30 days.

.PARAMETER IncludeDisabledUserAccounts
    If True, disabled user accounts will be included in the results.

.PARAMETER Searchbases
    The distinguished names of the OUs to search in. Ignored if Users is specified. Multiple OU's in string format are supported when split with a semicolon (;).

.NOTES
    Author: 
        Rein Leen
    Contributor(s):
        Bill Powell
        Gillian Stravers
    Context: 
        Machine
    Modification_history:
        Rein Leen       23-05-2023      Version ready for release
#>

#region [parameters]
[CmdletBinding()]
Param (
    [Parameter(Position = 0, Mandatory = $false, HelpMessage = 'Include disabled user accounts (True), or exclude them (False)')]
    [string]$IncludeDisabledUserAccounts,
    # A valid Distinguished name always contains two domainComponents.
    [Parameter(Position = 1, Mandatory = $false, HelpMessage = "The distinguished names of the OUs to search in. Multiple OU's in string format are supported when split with a semicolon.")]
    [string]$Searchbases    
)
#endregion [parameters]

#region [prerequisites]
# Required dependencies
#Requires -Version 5.1
#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

# Import modules (required in .NET engine)
Import-Module -Name ActiveDirectory

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

#region [functions]
# Function to get the ControlUp engine under which the script is running.
function Get-ControlUpEngine {
    $runtimeEngine = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $PID"
    switch ($runtimeEngine.ProcessName) {
        'cuAgent.exe' {
                return '.NET'
            }
        'powershell.exe' {
                return 'Classic'
            }
    }
}

# Function to assert the parameters are correct
function Assert-ControlUpParameter {
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [object]$Parameter,
        [Parameter(Position = 1, Mandatory = $true)]
        [boolean]$Mandatory,
        [Parameter(Position = 2, Mandatory = $true)]
        [ValidateSet('.NET','Classic')]
        [string]$Engine
    )

    # If a parameter is optional passing using a hyphen (-) or none is required when using the Classic engine. If this is the case return $null.
    if (($Mandatory -eq $false) -and (($Parameter -eq '-') -or ($Parameter -eq 'none'))) {
        return $null
    }

    # If a parameter is optional when using the .NET engine it should be empty. if this is the case return $null.
    if (($Engine -eq '.NET') -and ($Mandatory -eq $false) -and ([string]::IsNullOrWhiteSpace($Parameter))) {
        return $null
    }

    # Check if a mandatory parameter isn't null
    if (($Mandatory -eq $true) -and ([string]::IsNullOrWhiteSpace($Parameter))) {
        throw [System.ArgumentException] 'This parameter cannot be empty'
    }

    # ControlUp can add double quotes when using the .NET engine when a parameter value contains spaces. Remove these.
    if ($Engine -eq '.NET') {
        # Regex used to match double quotes
        $possiblyQuotedStringRegex = '^(?<op>"{0,1})\b(?<text>[^"]*)\1$'
        $Parameter -match $possiblyQuotedStringRegex | Out-Null
        return $Matches.text
    } else {
        return $Parameter
    }
}
#endregion [functions]

#region [variables]
$controlUpEngine = Get-ControlUpEngine

# Validate $IncludeDisabledUserAccounts
$IncludeDisabledUserAccounts = Assert-ControlUpParameter -Parameter $IncludeDisabledUserAccounts -Mandatory $false -Engine $controlUpEngine
# Convert string to boolean or set default
if (-not [string]::IsNullOrWhiteSpace($IncludeDisabledUserAccounts)) {
    [bool]$includeDisabled = [System.Convert]::ToBoolean($IncludeDisabledUserAccounts)
} else {
    [bool]$includeDisabled = $false
}

# Validate $Searchbases
$Searchbases = Assert-ControlUpParameter -Parameter $Searchbases -Mandatory $false -Engine $controlUpEngine
# Split $SeachBases on ";"
if (-not [string]::IsNullOrWhiteSpace($SearchBases)){
    $splitSearchBases = $SearchBases.Split(';').Trim()
}
#endregion [variables]

#region [actions]
Write-Verbose ('Starting actions')
if ([string]::IsNullOrWhiteSpace($Searchbases)) {
    $splitSearchBases = @((Get-ADDomain).DistinguishedName)
    Write-Verbose ('No searchbase included, querying entire directory on "{0}"' -f $Searchbases[0])
}

$ADPropertyList = "Name,UserPrincipalName,DistinguishedName,CN,Path" -split ','

$splitSearchBases | ForEach-Object {
    $searchbase = $_            
    try {
        if ($includeDisabled -eq $true) {
            $users = Get-ADUser -Filter * -SearchBase $searchbase -Properties cn,pwdlastset | Where-Object { ($_.pwdlastset -eq 0) -and (-not [string]::IsNullOrWhiteSpace($_.UserPrincipalName)) } | Select-Object  $ADPropertyList           
        } else {
            $users = Get-ADUser -Filter 'enabled -eq $true' -SearchBase $searchbase -Properties cn,pwdlastset | Where-Object { ($_.pwdlastset -eq 0) -and (-not [string]::IsNullOrWhiteSpace($_.UserPrincipalName)) } | Select-Object  $ADPropertyList 
        }
        # Return found users
        Write-Verbose ('Found {0} users' -f $users.Count)
        $users | ForEach-Object {
            $user = $_
            $cn = "CN=" + $user.CN
            if ($user.DistinguishedName.StartsWith($cn)) {
                $user.CN = $cn
                $TrimLength = $cn.Length + 1
                $user.Path = $user.DistinguishedName.Substring($TrimLength, $user.DistinguishedName.Length - $TrimLength)
            }
            $user
        }
    } catch [System.ArgumentException], [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Warning ('Searchbase {0} not found' -f $searchbase)
    } catch {
        Write-Warning ('Unknown error on searchbase {0}' -f $searchbase)
    }
} | Sort-Object -Property Path,CN | Format-Table -Property Name,UserPrincipalName,CN,Path -AutoSize

Write-Verbose ('Finished actions')

#endregion [actions]

