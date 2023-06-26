<#
.SYNOPSIS
    List inactive users in Active Directory and optionally disables/enables them.

.DESCRIPTION
    List all Active Directory Accounts which have not been logged into for specified days or more. 
    System account and accounts without any login activity are ignored by this script.
    Reporting on inactive accounts will return the command to disable those accounts.
    Disabling accounts will return the command to re-enable those accounts to counter mistakes.

    This script is intended to be used within ControlUp as an action.

.EXAMPLE
    Parameters:
        Operation = report
        Days = 30
        Searchbases = none
        Users = none

    Running the ControlUp Action using the above parameter will list all enabled user accounts which have not logged in with in the last 30 days.

.EXAMPLE
    Parameters:
        Operation = report 
        Days = none
        Searchbases = OU=UserAccounts,OU=All Accounts,DC=controlup,DC=local
        Users = none

    Running the ControlUp Action using the above parameter will list all enabled user accounts within the given OU and list the last logon datetime.

.EXAMPLE
    Parameters:
        Operation = disable
        Days = none
        Searchbases = none
        Users = User1, User2@controlup.local

    Running the ControlUp Action using the above parameter will disable the user accounts for User1 and User2.
    Users can be passed as SamAccountName or UPN and be split with a comma (,) or semicolon (;).

.PARAMETER Operation
    Valid options: report / enable / disable
    Operation to perform:
    - report = show accounts inactive for X days
    - enable = enable the accounts (requires $Users parameter)
    - disable = disable the accounts (requires $Users parameter)

.PARAMETER Days
    Enter the number of days since last login. Only functional on report operation.
    If 0 or empty all non-system enabled accounts with login activity will be returned.

.PARAMETER SearchBases
    Enter a searchbase or an array of searchbases (Distinguishe name) to query. Only functional on report operation. Multiple OU's in string format are supported when split with a semicolon (;).

.PARAMETER Users
    Enter the users to perform the requested operation on. Works with UPN or SamAccountName.
    Multiple users in string format are supported when split with a comma (,) or semicolon (;). Using this parameter ignores Searchbases.

.NOTES
    Author: 
        Rein Leen
    Contributor(s):
        Bill Powell
        Gillian Stravers
    Context: 
        Machine
    Modification_history:
        Rein Leen       26-05-2023      Version ready for release
#>
#region [parameters]
[CmdletBinding()]
Param (
    [Parameter(Position = 0, Mandatory = $false, HelpMessage = 'Operation to perform. report = show accounts inactive for X days; enable = enable the accounts (requires $Users parameter); disable = disable the accounts (requires $Users parameter)')]
    [string]$Operation,
    [Parameter(Position = 1, Mandatory = $false, HelpMessage = 'Number of days since last logon to filter from the results. 0 will not filter the results.')]
    [string]$Days,
    [Parameter(Position = 2, Mandatory = $false, HelpMessage = 'The distinguished names of the OUs to search in. Ignored if Users is specified. Multiple OUs in string format are supported when split with a semicolon.')]
    [string]$SearchBases,
    [Parameter(Position = 3, Mandatory = $false, HelpMessage = 'User(s) to perform the requested operation on. Works with SID, GUID, sAMAccountName or Distinguished name. Using this parameter ignores Searchbases')]
    [string]$Users    
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

function Get-ADUserLastLogonDate {
    Param (
        [int]$Days,
        [string[]]$Searchbases,
        [string[]]$Users
    )

    $userProperties = [hashtable]@{
        'Name' =                        $null
        'UserPrincipalName' =           $null
        'LastLogon' =                   $null
        'sAMAccountName' =              $null
    }

    # Create the right filter
    $filter = '((enabled -eq $true) -and (LastLogon -gt 0))'

    # Create a list to hold all data
    $userDetailsHashtable = @{}

    # Get all DC's to retrieve LastLogon. This property is propagated with a delay of 9-14 days (default) and can only accurately be retrieved on the DC where a user authenticated with.
    Write-Verbose 'Retrieving Domain Controllers'
    $allDomainControllers = Get-ADDomainController -Filter * | Where-Object { $_.Enabled -eq $true }
    Write-Verbose ('Found {0} Domain Controllers.' -f $allDomainControllers.Length)

    # Get all user details from both DC's and add it to the list, exclude system accounts
    foreach ($domainController in $allDomainControllers) {
        # If users are specified, retrieve their user objects from Active Directory. Filter on $Days if specified.
        Write-Verbose ('Querying Domain Controller {0}' -f $domainController.Name)
        if (-not [string]::IsNullOrWhiteSpace($Users)) {
            $userObjects = foreach ($user in $Users) {
                Get-ADUser -LDAPFilter ('(|(UserPrincipalName={0})(SamAccountName={0}))' -f $user) -Properties LastLogon -Server $domainController
            }
        } else {
            # Retrieve all user objects from Active Directory that match the filter
            $userObjects = foreach ($searchbase in $Searchbases) {
                Get-ADUser -Filter $filter -SearchBase $searchbase -Properties LastLogon -Server $domainController | Where-Object {$_.UserPrincipalName}
            }
        }
        foreach ($user in $userObjects) {
            $currentUserDetails = $userDetailsHashtable[$user.UserPrincipalName]
            if ([string]::IsNullOrWhiteSpace($currentUserDetails)) {
                $userDetailsHashtable[$user.UserPrincipalName] = [PSCustomObject]@{'LastLogon' = $user.LastLogon; 'Name' = $user.Name; 'sAMAccountName' = $user.sAMAccountName}
            } elseif ($user.LastLogon -gt $currentUserDetails.LastLogon ) {
                $userDetailsHashtable[$user.UserPrincipalName].LastLogon = $user.LastLogon
            }
        }
    }
    # Return the user details
    foreach ($user in $userDetailsHashtable.GetEnumerator()) {
        if ($user.Value.LastLogon -lt [datetime]::UtcNow.AddDays(-$Days).ToFileTime()) {
            $userProperties['UserPrincipalName'] = $user.Name
            $userProperties['Name'] = $user.Value.Name
            $userProperties['LastLogon'] = if ($user.Value.LastLogon -eq 0) {'never'} else {[datetime]::FromFileTime($user.Value.LastLogon)}
            $userProperties['sAMAccountName'] = $user.Value.sAMAccountName
            [PSCustomObject]$userProperties
        }
    }
}
#endregion [functions]

#region [variables]
$controlUpEngine = Get-ControlUpEngine

# Validate $Operation
$Operation = Assert-ControlUpParameter -Parameter $Operation -Mandatory $false -Engine $controlUpEngine
# Set default if empty
if (([string]::IsNullOrWhiteSpace($Operation)) -or ($Operation -notin @('report','enable','disable'))) {
    $Operation = 'report'
}

# Validate $Days
$Days = Assert-ControlUpParameter -Parameter $Days -Mandatory $false -Engine $controlUpEngine

# Validate $Users
$Users = Assert-ControlUpParameter -Parameter $Users -Mandatory $false -Engine $controlUpEngine
# Split $Users on common delimiters
if (-not [string]::IsNullOrWhiteSpace($Users)) {
    $splitUsers = $Users.Split(',;').Trim()
}

# Validate $Searchbases
$Searchbases = Assert-ControlUpParameter -Parameter $Searchbases -Mandatory $false -Engine $controlUpEngine
# If $Searchbases is not specified use the root of the domain.
# Overwrite $Searchbases if $Users is specified
if (-not [string]::IsNullOrWhiteSpace($SearchBases)){
    $splitSearchBases = $SearchBases.Split(';').Trim()
} elseif (([string]::IsNullOrWhiteSpace($Searchbases)) -or (-not [string]::IsNullOrWhiteSpace($Users))) {
    $splitSearchbases = @((Get-ADDomain).DistinguishedName)
}
#endregion [variables]

#region [actions]
if ($Operation -eq 'report') {
    $userDetails = (Get-ADUserLastLogonDate -Days $Days -SearchBases $splitSearchbases -Users $splitUsers)
    Write-Verbose ('Found {0} users given the query parameters' -f $userDetails.Length)
    # Output data
    if ($userDetails | Where-Object {$_.LastLogon -ne 'never'}){
        Write-Output ($userDetails | Where-Object {$_.LastLogon -ne 'never'} | Sort-Object -Property LastLogon -Descending)
    }
    if ($userDetails | Where-Object {$_.LastLogon -eq 'never'}){
        Write-Output $userDetails | Where-Object {$_.LastLogon -eq 'never'}
    }   
    # Show command to disable the accounts found by the current query
    foreach ($user in $userDetails.sAMAccountName) {
        $userControlUpString += ('{0},' -f $user)
    }
    
    Write-Output (@'

To disable the accounts found run the action again in ControlUp with the disable parameter and the following users (remove names to exclude them):
{0}
'@ -f $userControlUpString.Substring(0,$userControlUpString.LastIndexOf(',')))

} elseif ($Operation -eq 'enable') {
    foreach ($user in $splitUsers) {
        Try {
            Get-ADUser -LDAPFilter ('(|(UserPrincipalName={0})(SamAccountName={0}))' -f $user) | Enable-ADAccount
            Write-Output ('Enabled user account for user {0}.' -f $user)
        } Catch {
            Write-Warning ('Could not enable the user account for user {0}.' -f $user)
        }
    }
} else {
    foreach ($user in $splitUsers) {
        Try {
            Get-ADUser -LDAPFilter ('(|(UserPrincipalName={0})(SamAccountName={0}))' -f $user) | Disable-ADAccount
            Write-Output ('Disabled user account for user {0}.' -f $user)
            $userControlUpString += ('{0},' -f $user)
        } Catch {
            Write-Warning ('Could not disable the user account for user {0}.' -f $user)
        }
    }
    Write-Output (@'

To enable the disabled accounts again run the action again in ControlUp with the enable parameter and the following users (remove names to exclude them):
{0}
'@ -f $userControlUpString.Substring(0,$userControlUpString.LastIndexOf(',')))
}
#endregion [actions]

