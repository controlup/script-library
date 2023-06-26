<#
.SYNOPSIS
    Gets the Active Directory expiration date of users within the given parameters.

.DESCRIPTION
    Gets the Active Directory expiration date of users within the given parameters. Only user accounts which expire are shown in the results.
    Not getting any results means there are no user accounts expiring within the given parameters.

    This script is intended to be used within ControlUp as an action. If the .NET engine is used all optional parameter are truly optional.
    The classic engine cannot handle empty optional parameters. If the .NET engine is not used pass "none" to leave an optional parameter empty.

.EXAMPLE
    Parameters:
        DaysBeforeAccountExpiration = 30
        SearchBases = none
        Users = none
        IncludeDisabledUserAccounts = none

    Running the ControlUp Action using the above parameters retrieves all active user accounts that are due to expire within the next 30 days.

.EXAMPLE
    Parameters:
        DaysBeforeAccountExpiration = 0
        SearchBases = OU=UserAccounts,OU=All Accounts,DC=controlup,DC=local
        Users = none
        IncludeDisabledUserAccounts = none

    Running the ControlUp Action using the above parameters retrieves all active user accounts within the given OU and any nested OU's.

.EXAMPLE
    Parameters:
        DaysBeforeAccountExpiration = none
        SearchBases = none
        Users = user1, user2@controlup.local; user3@controlup.local
        IncludeDisabledUserAccounts = none

    Running the ControlUp Action using the above parameters retrieves the account expiration data for users user1, user2 and user3.
    Users can be passed as SamAccountName or UPN and be split with a comma (,) or semicolon (;).

.PARAMETER DaysBeforeAccountExpiration
    Number of days before account expiration. 0 days returns the account expiration date of all users within the given parameters.

.PARAMETER SearchBases
    The distinguished names of the OUs to search in. Ignored if Users is specified. Multiple OU's in string format are supported when split with a semicolon (;).

.PARAMETER Users
    User(s) to get the AccountExpirationDate of based on UPN or SamAccountName. Using this parameter ignores all other parameters.
    Multiple users in string format are supported when split with a comma (,) or semicolon (;).

.PARAMETER IncludeDisabledUserAccounts
    If True, disabled user accounts will be included in the results.

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
    [Parameter(Position = 0, Mandatory = $false, HelpMessage = 'Number of days before account expiration. 0 days returns the account expiration date of all users within the given parameters.')]
    [string]$DaysBeforeAccountExpiration,
    [Parameter(Position = 1, Mandatory = $false, HelpMessage = "The distinguished names of the OUs to search in. Ignored if Users is specified. Multiple OU's in string format are supported when split with a semicolon.")]
    [string]$SearchBases,
    [Parameter(Position = 2, Mandatory = $false, HelpMessage = 'User(s) to get the AccountExpirationDate of based on UPN or SamAccountName. Using this parameter ignores Searchbases. Multiple users in string format are supported when split with a comma or semicolon')]
    [string]$Users,
    [Parameter(Position = 3, Mandatory = $false, HelpMessage = 'If True, disabled user accounts will be included in the results.')]
    [string]$IncludeDisabledUserAccounts = $false  
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

# Function to return the expiration details
$UserExpirationProperties = "Name,UserPrincipalName,AccountExpirationDate,DaysUntilAccountExpiration,DistinguishedName" -split ','
function Get-ADUserAccountExpirationDetails {
    Param (
        [object]$userObject,
        [datetime]$currentDate
    )

    $userProperties = $userObject | Select-Object -Property $UserExpirationProperties

    # Exlude system accounts
    if ([string]::IsNullOrWhiteSpace($userObject.UserPrincipalName)) {
        continue
    }
    
    # Exclude accounts where account expiration is maxed ([int64]::MaxValue) which is an AD default value
    if (($userObject.accountExpires -ne 0) -and ($userObject.accountExpires -ne [int64]::MaxValue)) {
        $userAccountExpires = [datetime]::FromFileTime($userObject.accountExpires)
        # Add properties to hashtable
        $userProperties.AccountExpirationDate = [string]::Concat($userAccountExpires.ToString("s"), "Z")
        $userProperties.DaysUntilAccountExpiration = ($userAccountExpires.Date - $currentDate).Days
        $userProperties
    }
}
#endregion [functions]

#region [variables]
$controlUpEngine = Get-ControlUpEngine

# Validate $DaysBeforeAccountExpiration
$DaysBeforeAccountExpiration = Assert-ControlUpParameter -Parameter $DaysBeforeAccountExpiration -Mandatory $false -Engine $controlUpEngine

# Validate $SearchBases
$SearchBases = Assert-ControlUpParameter -Parameter $SearchBases -Mandatory $false -Engine $controlUpEngine
if ([string]::IsNullOrWhiteSpace($Searchbases)) {
    # If $Searchbases is not specified use the root of the domain
    $splitSearchbases = @((Get-ADDomain).DistinguishedName)
} elseif ($Searchbases -match 'DC=[a-z]*,DC=[a-z]*') {
    # Split $SeachBases on semicolon.
    $splitSearchBases = $SearchBases.Split(';').Trim()
} else {
    throw [System.ArgumentException] 'The Searchbases parameter does not follow the distinguished name format.'
}

# Validate $Users
$Users = Assert-ControlUpParameter -Parameter $Users -Mandatory $false -Engine $controlUpEngine
if (-not [string]::IsNullOrWhiteSpace($Users)) {
    # Split $Users on common delimiters
    $splitUsers = $Users.Split(@(',',';')).Trim()
}

# Validate $IncludeDisabledUserAccounts
$IncludeDisabledUserAccounts = Assert-ControlUpParameter -Parameter $IncludeDisabledUserAccounts -Mandatory $false -Engine $controlUpEngine
if (-not [string]::IsNullOrWhiteSpace($IncludeDisabledUserAccounts)) {
    # Convert $IncludeDisabledUserAccounts to a boolean
    try {
        [bool]$IncludeDisabledUsers = [System.Convert]::ToBoolean($IncludeDisabledUserAccounts)
    } catch {
        throw [System.ArgumentException] 'The IncludeDisabledUsers parameter cannot be converted to a boolean'
    }
}

# Get current date
$currentDate = [datetime]::UtcNow.Date
$currentDateFileTime = $currentDate.ToFileTimeUtc()

# Get reference filetime if required, default to space age if not
if ([string]::IsNullOrWhiteSpace($DaysBeforeAccountExpiration) -or ($DaysBeforeAccountExpiration -eq 0)) {
    [Int64]$referenceFileTime = [int64]::MaxValue
}
else {
    [Int64]$referenceFileTime = $currentDate.AddDays($DaysBeforeAccountExpiration).ToFileTime()
}

# Define the filter to retrieve all users
if ($IncludeDisabledUsers -eq $true) {
    $filter = ('(accountExpires -gt {0}) -and (accountExpires -lt {1})' -f $currentDateFileTime, $referenceFileTime)
} else {
    $filter = ('(enabled -eq $true) -and (accountExpires -gt {0}) -and (accountExpires -lt {1})' -f $currentDateFileTime, $referenceFileTime)
}    
#endregion [variables]

#region [actions]

# If users are specified, retrieve their user objects from Active Directory
$userObjects = New-Object System.Collections.Generic.List[object]
Write-Verbose ('Starting actions')
if (-not [string]::IsNullOrWhiteSpace($splitUsers)) {
    foreach ($user in $splitUsers) {
        $foundUser = Get-ADUser -LDAPFilter ('(|(UserPrincipalName={0})(SamAccountName={0}))' -f $user) -Properties accountExpires
        if ([string]::IsNullOrWhiteSpace($foundUser)) {
            Write-Warning ('User {0} was not found on either UserPrincipalName or SamAccountName' -f $user)
            continue
        }
        $userObjects.Add($foundUser)
    }
}
else {
    # Retrieve all user objects from Active Directory that match the filter
    foreach ($searchbase in $splitSearchbases) {
        Get-ADUser -Filter $filter -SearchBase $searchbase -Properties accountExpires | ForEach-Object {
            $userObjects.Add($_)
        }
    }
}
Write-Verbose ('Found {0} users with the given parameters' -f $userObjects.Count)

# Loop through each user object and display the expiration date if it is set
$userObjects | Sort-Object -Property accountExpires | ForEach-Object {
    Get-ADUserAccountExpirationDetails -userObject $_ -currentDate $currentDate       
} | Format-Table

Write-Verbose ('Finished actions')
#endregion [actions]

