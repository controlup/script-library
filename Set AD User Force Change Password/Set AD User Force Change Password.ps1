<#
.SYNOPSIS
    Set "ChangePasswordAtLogon" to true.

.DESCRIPTION
    Set "ChangePasswordAtLogon" to true on the selected users.

    This script is intended to be used within ControlUp as an action.

.EXAMPLE
    Parameters:
        Users = User1, User2

    Running the ControlUp Action using the above parameter will force the users User1 and User2 to change their passwords at next logon.

.EXAMPLE
    Parameters:
        Users = User1, User2@controlup.local

    Running the ControlUp Action using the above parameter will force the users User1 and User2 to change their passwords at next logon.
    Users can be passed as SamAccountName or UPN and be split with a comma (,) or semicolon (;).

.PARAMETER Users
    Enter the users you want to set the 'User must change password at next logon' to true. 
    Lookup of users will happen based on their Active Directory Identity (SAMAccountName or UPN).
    Multiple users in string format are supported when split with a comma or semicolon.
.NOTES
    Author: 
        Rein Leen
    Contributor(s):
        Bill Powell
        Gillian Stravers
    Context: 
        Machine
    Modification_history:
        Rein Leen       25-05-2023      Version ready for release
#>

#region [parameters]
[CmdletBinding()]
Param (
    [Parameter(Position = 0, Mandatory = $true, HelpMessage = 'Input an username, an array of usernames, or a delimited string of usernames')]
    [string]$Users
)
#endregion [parameters]

#region [prerequisites]
# Required dependencies
#Requires -Version 5.1
#Requires -Modules ActiveDirectory

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

# Validate $Users
$Users = Assert-ControlUpParameter -Parameter $Users -Mandatory $true -Engine $controlUpEngine
# Split $Users on common delimiters
$splitUsers = $Users.Split(@(',',';')).Trim()
#endregion [variables]

#region [actions]
foreach ($user in $splitUsers) {
    Write-Verbose ('Starting actions on {0}' -f $user)
    # Initialize hashtable
    $userProperties = [hashtable]@{
        'Name' = $null
        'UserPrincipalName' = $null
        'TransformationSuccessful' = $null
    }
    try {
        # Make sure user exists        
        $foundUser = Get-ADUser -LDAPFilter ('(|(UserPrincipalName={0})(SamAccountName={0}))' -f $user)
        Write-Verbose ('Performing action on {0}' -f $foundUser.DistinguishedName)
        # Set ChangePasswordAtLogon
        Set-ADUser -Identity $foundUser.DistinguishedName -ChangePasswordAtLogon $true | Out-Null
        # Add properties to hashtable
        $userProperties['Name'] = $foundUser.SamAccountName
        $userProperties['UserPrincipalName'] = $foundUser.UserPrincipalName
        $userProperties['TransformationSuccessful'] = $true
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{
        # Expected error if user cannot be found
        Write-Warning ('User {0} not found' -f $user)
        # Add properties to hashtable
        $userProperties['Name'] = $user
        $userProperties['TransformationSuccessful'] = $false
    } catch {
        # Unexpected error(s)
        Write-Warning ('Unknown error on user {0}' -f $user)
        # Add properties to hashtable
        $userProperties['Name'] = $user
        $userProperties['TransformationSuccessful'] = $false
    }
    # Return hashtable
    [PSCustomObject]$userProperties
    Write-Verbose ('Finished actions on user {0}' -f $user)
}
#endregion [actions]

