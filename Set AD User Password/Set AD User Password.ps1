<#
.SYNOPSIS
    Reset the AD Password for a user with a human-readable password.

.DESCRIPTION
    Reset the AD Password for a user with a human-readable password.
    This script creates a password starting with 3 random numbers, several capitalized words (as much as needed for the minimal password length) and a special character.

    To start using this script first create an account or login to an existing account at https://auth.what3words.com/.
    After logging in, generate an API-key at https://what3words.com/select-plan?referrer=%2Fpublic-api&currency=USD. Most likely the Free plan will be more than enough for the intent of this script.

    Note! Your API-key should be visible at https://accounts.what3words.com/overview. While creating this script, this page didn't always show the key. My workaround was to just click around on other pages on the sidebar until it showed up (yes, really).
    The API-key consists of a string of 8 numbers and uppercase letters.

    This script is intended to be used within ControlUp as an action. The password will be shown to the console since there currently is no way to securely send the password directly to the user.

.EXAMPLE
    Parameters:
        APIKey = XX123YYY
        MinimalPasswordLength = 50
        User = user1@controlup.local

    Running the ControlUp Action using the above parameters will set the a new password for the user user1@controlup.local with a human readable password of 50 or more characters.

.EXAMPLE
    Parameters:
        APIKey = XX123YYY
        MinimalPasswordLength = none
        User = user1

    Running the ControlUp Action using the above parameters will set the a new password for the user user1 with a human readable password of 20 (default value) or more characters.

.PARAMETER APIKey
    Your personal API key. The API-key consists of a string of 8 numbers and uppercase letters.

.PARAMETER User
    Enter the user to reset the password for. Works with UPN or SamAccountName.
    
.PARAMETER MinimalPasswordLength
    The minimal length of the password. A generated password can be longer since only whole words are used in a password.

.NOTES
    Author: 
        Rein Leen
    Contributor(s):    
    Context: 
        Machine
    Modification_history:
        Rein Leen       06-05-2023      Version ready for release
#>

#region [parameters]
[CmdletBinding()]
Param (
    [Parameter(Position = 0, Mandatory = $true, HelpMessage = 'Your personal API key.')]
    [string]$APIKey,
    [Parameter(Position = 1, Mandatory = $true, HelpMessage = 'User for which to reset the password. Works with sAMAccountName or UPN.')]
    [string]$User,
    [Parameter(Position = 2, Mandatory = $false, HelpMessage = 'Minimal length of the new password')]
    [int]$MinimalPasswordLength = 20
)
#endregion [parameters]

#region [prerequisites]
# Required dependencies
#Requires -Version 5.1
#Requires -Modules ActiveDirectory

# Import modules (required in .NET engine)
Import-Module -Name ActiveDirectory

# Setting error actions
$ErrorActionPreference = 'Stop'
$DebugPreference = 'SilentlyContinue'
#endregion [prerequisites]

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
$APIKey = Assert-ControlUpParameter -Parameter $APIKey -Mandatory $true -Engine $controlUpEngine

# Validate $Users
$User = Assert-ControlUpParameter -Parameter $User -Mandatory $true -Engine $controlUpEngine

# Validate $Users
$MinimalPasswordLength = Assert-ControlUpParameter -Parameter $MinimalPasswordLength -Mandatory $true -Engine $controlUpEngine
#endregion [variables]

#region [actions]
# Generate password components
[string]$password = Get-Random -Minimum 100 -Maximum 999
$specialCharacter = '!@#$%&*'[(Get-Random -Minimum 0 -Maximum 6)]

try {
    # Get TextInfo (needed for capitalizing words)
    $textInfo = (Get-Culture).TextInfo

    while ($password.Length -lt ($MinimalPasswordLength -1)) {
        # Generate coordinates
        [double]$MinLat = -90.0
        [double]$MaxLat =  90.0
        [double]$MinLong = -180.0
        [double]$MaxLong = 180.0
        $latitude = Get-Random -Minimum $MinLat -Maximum $MaxLat
        $longitude  = Get-Random -Minimum $MinLong -Maximum $MaxLong

        # TODO: The result may return a default in English depending on the coordinates (Oceans and Antarctica). Discuss how we should handle this.
        $result = Invoke-RestMethod -Uri ('https://api.what3words.com/v3/convert-to-3wa?coordinates={0},{1}&key={2}' -f $latitude, $longitude, $APIKey) -UseBasicParsing
    
        # Capitalize words and add them to the password string until the minimum password length is reached
        $i = 0
        $words = $result.words.Split('.')
        while (($password.Length -lt ($MinimalPasswordLength -1)) -and ($i -lt 3)) {        
            $password += $textInfo.ToTitleCase($words[$i])
            $i++
        }
    }
    # Add special characters after the password
    $password += $specialCharacter
} catch {
    Write-Warning 'Could not generate a password. Password has not been reset.'
    Write-Warning ('Ran into the following issue: {0}' -f $PSItem.Exception.Message)
}

$foundUser = Get-ADUser -LDAPFilter ('(|(UserPrincipalName={0})(SamAccountName={0}))' -f $User)
if ($null -eq $foundUser) {
    Write-Warning ('Could not find user {0}. Password has not been reset' -f $user)
} else {
    Set-ADAccountPassword -Identity $foundUser.DistinguishedName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force)
    Set-ADUser -Identity $foundUser.DistinguishedName  -ChangePasswordAtLogon:$True
    Write-Output (@"
Set the following password for {0}:
{1}
"@ -f $foundUser.UserPrincipalName, $password)
}
#endregion [actions]

