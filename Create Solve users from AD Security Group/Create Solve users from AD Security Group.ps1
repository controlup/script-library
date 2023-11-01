#requires -Modules ActiveDirectory

<#
  .SYNOPSIS
  This script synchronises AD users from an AD Security Group to Solve 
  .DESCRIPTION
  This script extracts AD users from an AD Security Group. It checks whether the User Principal Name is already registered with Solve (using Get-CUUsersList), and if not, performs the registration using Add-CUUser
  .PARAMETER SolveUserSecurityGroup
  The AD Security Group to perform the action on, supplied via a trigger
  .NOTES
   Version:        0.1
   Context:        Computer, executes on Monitor
   Author:         Bill Powell
   Requires:       Realtime DX 8.8, ActiveDirectory module (RSAT)
   Creation Date:  2023-09-06

  .LINK
   https://support.controlup.com/docs/powershell-cmdlets-for-solve-actions
#>


[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, HelpMessage = 'Active Directory Security Group defining Solve users to be created')]
    [ValidateNotNullOrEmpty()]
    [string]$SolveUserSecurityGroup
)

Import-Module ActiveDirectory

# uncomment for CU Internal Testing
#$SolveUserSecurityGroup = 'sg-SolveUsersTest'

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

#region Load the version of the module to match the running monitor and check that it has the new features

function Get-MonitorDLLs {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][string[]]$DLLList
    )
    [int]$DLLsFound = 0
    Get-CimInstance -Query "SELECT * from win32_Service WHERE Name = 'cuMonitor' AND State = 'Running'" | ForEach-Object {
        $MonitorService = $_
        if ($MonitorService.PathName -match '^(?<op>"{0,1})\b(?<text>[^"]*)\1$') {
            $Path = $Matches.text
            $MonitorFolder = Split-Path -Path $Path -Parent
            $DLLList | ForEach-Object {
                $DLLBase = $_
                $DllPath = Join-Path -Path $MonitorFolder -ChildPath $DLLBase
                if (Test-Path -LiteralPath $DllPath) {
                    $DllPath
                    $DLLsFound++
                }
                else {
                    throw "DLL $DllPath not found in running monitor folder"
                }
            }
        }
    }
    if ($DLLsFound -ne $DLLList.Count) {
        throw "cuMonitor is not installed or not running"
    }
}

$AcceptableModules = New-Object System.Collections.Generic.List[object]
try {
    $DllsToLoad = Get-MonitorDLLs -DLLList @('ControlUp.PowerShell.User.dll')
    $DllsToLoad | Import-Module 
    $DllsToLoad -replace "^.*\\",'' -replace "\.dll$",'' | ForEach-Object {$AcceptableModules.Add($_)}
}
catch {
    $exception = $_
    Write-Error "Required DLLs not loaded: $($exception.Exception.Message)"
}

if (-not ((Get-Command -Name 'Invoke-CUAction' -ErrorAction SilentlyContinue).Source) -in $AcceptableModules) {
   Write-Error "ControlUp version 8.8 commands are not available on this system"
   exit 0
}

#endregion

$ADPropertyList = 'accountExpires,SamAccountName,UserPrincipalName,DistinguishedName,mail,GivenName,Surname' -split ','

#region expand group

[hashtable]$Script:DNsProcessed = @{}    # <- every DN goes into here as it is processed, to prevent
                                         # duplication and infinite recursion
$Script:AccountsInGroup = New-Object System.Collections.Generic.List[PSObject]

function Unpack-NestedGroup {
    [CmdletBinding()]
    param (
        [string]$GroupDN
    )
    if ($Script:DNsProcessed[$GroupDN] -eq $true) {
        return
    }
    $Script:DNsProcessed[$GroupDN] = $true
    $ADGroup = Get-ADGroup -Identity $GroupDN -Properties members
    $ADGroup.members | ForEach-Object {
        $memberDN = $_
        $ADObject = Get-ADObject -Identity $memberDN
        switch ($ADObject.ObjectClass) {
            'group' {
                    Unpack-NestedGroup -GroupDN $memberDN
                }
            'user' {
                    if ($Script:DNsProcessed[$memberDN] -ne $true) {
                        $Script:AccountsInGroup.Add($memberDN)
                        $Script:DNsProcessed[$memberDN] = $true
                    }
                }
            default {}
        }
    }
}

#endregion

$GroupToAdd = Get-ADGroup -Filter "name -eq '$SolveUserSecurityGroup'"

if ($null -ne $GroupToAdd) {
    #
    Unpack-NestedGroup -GroupDN $GroupToAdd.DistinguishedName
    # getting the CU Users list is slow - so no point if we haven't found the group
    $CUUsers = Get-CUUsersList

    $Script:AccountsInGroup | ForEach-Object {
        $DistinguishedName = $_
        $ADUser = Get-ADUser -Identity $DistinguishedName -Properties $ADPropertyList
        <#
        # for CU internal test
        if ($ADUser.accountExpires -lt ((Get-Date).ToFileTime())) {
            #
            # account has expired - extend expiration date by 1 month
            $NewExpiry = (Get-Date).AddMonths(1)
            Set-ADUser -Identity $ADUser.SamAccountName -AccountExpirationDate $NewExpiry
        }
        $ADUser | Out-Null
        #>
        $SolveUser = $CUUsers | Where-Object {$_.upn -eq $ADUser.UserPrincipalName}
        if ($SolveUser -eq $null) {
            #
            # not currently a solve user - let's create them as such
            $UserDnsDomain = (($ADUser.DistinguishedName -split ',' | Where-Object {$_ -like "DC=*"}) -replace "^DC=",'') -join '.'
            $AddUserSplat = @{
                Upn = $ADUser.UserPrincipalName
                Email = $ADUser.mail
                SamAccountName = $ADUser.SamAccountName
                UserDnsDomain = $UserDnsDomain
                FirstName = $ADUser.GivenName
                LastName = $ADUser.Surname
            }
            Add-CUUser @AddUserSplat
            Write-Output "Group: $SolveUserSecurityGroup , user $($ADUser.UserPrincipalName) has been added as a solve user"
        }
        else {
            Write-Output "Group: $SolveUserSecurityGroup , user $($ADUser.UserPrincipalName) is already a solve user"
        }
    }
}
else {
    Write-Error "AD group $SolveUserSecurityGroup not found"
    exit 1
}





