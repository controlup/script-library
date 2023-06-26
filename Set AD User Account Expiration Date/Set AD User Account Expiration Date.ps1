<#
.SYNOPSIS
    Sets the expiration date of one or multiple user accounts in AD.

.DESCRIPTION
    Sets the expiration date of one or multiple user accounts in AD.

    This script is intended to be used within ControlUp as an action.

.EXAMPLE
    Parameters:
        Users = User1, User2
        Date = 2024-12-31

    Running the ControlUp Action using the above parameters will set the accounts of User1 and User2 to expire on December 31 2024

.EXAMPLE
    Parameters:
        Users = User1, User2@controlup.local
        Date = 2023-06-24 17:00 +3:30

    Running the ControlUp Action using the above parameters will set the accounts of User1 and User2 to expire on June 24 2023 1:30PM UTC (5PM in Tehran, Iran)
    Users can be passed as SamAccountName or UPN and be split with a comma (,) or semicolon (;).

.PARAMETER Users
    User(s) to set the AccountExpirationDate of based on UPN or SamAccountName. 
    Multiple users in string format are supported when split with a (,) or semicolon (;).

.PARAMETER Date
    Date (string) when the account should expire. Use the format yyyy-MM-dd HH:mm K. 
    Time and timeoffset are optional. If time and timeoffset are not set the account will expire at the end of the previous day.

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
    [Parameter(Position = 0, Mandatory = $true, HelpMessage = 'User(s) to perform the action on')]
    [string]$Users,
    [Parameter(Position = 1, Mandatory = $true, HelpMessage = 'Date (string) when the account should expire. Use the format yyyy\MM\dd HH:mm K. Time and Timeoffset are optional')]
    [ValidateScript({
        $acceptedFormats = @(
          'yyyy-M-d'
          'yyyy-M-d H:mm'
          'yyyy-M-d H:mm K'
          'yyyy-M-d H:mm z'
          '\"yyyy-M-d H:mm\"'
          '\"yyyy-M-d H:mm K\"'
          '\"yyyy-M-d H:mm z\"'
        ) -as [string[]]
        [datetime]::ParseExact($_, $acceptedFormats, $null, [System.Globalization.DateTimeStyles]::None)
    })]
    [string]$Date
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

# Validate $Date
$Date = Assert-ControlUpParameter -Parameter $Date -Mandatory $true -Engine $controlUpEngine

# Validate $Users
$Users = Assert-ControlUpParameter -Parameter $Users -Mandatory $true -Engine $controlUpEngine
# Split $Users on common delimiters
$splitUsers = $Users.Split(@(',',';')).Trim()

# Active Directory always uses UTC. Therefor convert the datetime to UTC
$acceptedFormats = @(
    'yyyy-M-d'
    'yyyy-M-d H:mm'
    'yyyy-M-d H:mm K'
    'yyyy-M-d H:mm z'
    '\"yyyy-M-d H:mm\"'
    '\"yyyy-M-d H:mm K\"'
    '\"yyyy-M-d H:mm z\"'
) -as [string[]]
$utcDatetime = ([datetime]::ParseExact($Date, $acceptedFormats, $null, [System.Globalization.DateTimeStyles]::None))
#endregion [variables]

#region [actions]
foreach ($user in $splitUsers) {
    $userProperties = [hashtable]@{
        'Name' =                            $null
        'UserPrincipalName' =               $null
        'OriginalAccountExpirationDate' =   $null
        'NewAccountExpirationDate' =        $null
    }
    $userObject = Get-ADUser -LDAPFilter ('(|(UserPrincipalName={0})(SamAccountName={0}))' -f $user) -Properties accountExpires
    if ((-not [string]::IsNullOrWhiteSpace($userObject)) -and ([string]::IsNullOrWhiteSpace($userObject.Enabled))) {
        Write-Warning ('The account for user {0} is disabled. No action will be performed' -f $userObject.Name)
        $userProperties['Name'] = $userObject.Name
        $userProperties['UserPrincipalName'] = $userObject.UserPrincipalName
        $userProperties['OriginalAccountExpirationDate'] = if ($userObject.accountExpires -eq 0) {'Not set'} else {[datetime]::FromFileTime($userObject.accountExpires).ToShortDateString()}
        $userProperties['NewAccountExpirationDate'] = 'Not modified (account is disabled)'
    } elseif (-not [string]::IsNullOrWhiteSpace($userObject)) {
        Write-Verbose ('Found user {0}, setting account expiration date with to {1}' -f $userObject.Name, $Date)
        Set-ADUser -Identity $userObject.DistinguishedName -AccountExpirationDate $utcDatetime
        $userProperties['Name'] = $userObject.Name
        $userProperties['UserPrincipalName'] = $userObject.UserPrincipalName
        $userProperties['OriginalAccountExpirationDate'] = if ($userObject.accountExpires -eq 0) {'Not set'} else {[datetime]::FromFileTime($userObject.accountExpires).ToShortDateString()}
        $userProperties['NewAccountExpirationDate'] = $utcDatetime.ToShortDateString()
    } else {
        Write-Warning ('User {0} could not be found' -f $user)
        continue
    }
    [PSCustomObject]$userProperties
}
#endregion [actions]

