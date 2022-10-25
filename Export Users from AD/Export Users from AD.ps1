#requires -version 5.1

<# 
.SYNOPSIS
	Queries Active Directory for Users/Groups and exports a CSV
.DESCRIPTION
	Queries AD Groups/Users for information about each user account and exports it as
	CSV for Solve User Creation
.PARAMETER Users
	User Parameter is comma separated, needs to be samaccountname or upn
		-User "User1,User2,User3" 
.PARAMETER Groups
	Groups Parameter is comma separated, needs to be the group name
		-Group "Group1,Group2,Group3" 
.PARAMETER ExportFile
	ExportFile Parameter is the file path to export the file on the machine where the script is run from.
		-ExportFile "c:\temp\csv.csv" 
.EXAMPLE
	.\ExportGroupsForUserCreating.ps1 -Group "Group1,Group2,Group3" -User "User1,User2,User3" -ExportFile "c:\temp\csv.csv"
.NOTES
	If a user does not have an email address or UPN the user will not be exported.
	Builtin groups such as Domain Users are not supported.
#>

[CmdletBinding()]
Param(
	[Parameter(Mandatory = $false, HelpMessage = 'Separate AD Users with Comma.' )]
	[array] $Users,	
	[Parameter(Mandatory = $false, HelpMessage = 'Separate AD groups with Comma.' )]
	[array] $Groups,
	[Parameter(Mandatory = $false, HelpMessage = 'Export File to current DIR for full path.' )]
	[string]$ExportFile
)

[bool]$IgnoreEmptyEmail = $true

function recurseGroups {
	#This is a recursion loop. If it finds a group inside a group it re-runs the 
	#function until there are no sub-groups for that group
	Param([string]$DN)
	if ($dn) {
		#checks if DN exists in AD, which it should since we queried them already
		if (([adsi]::Exists("LDAP://$DN"))) {
			$group = [adsi]("LDAP://$DN")
			#Grabs ADSI object and searches the group for groups
			($group).member | ForEach-Object {
				$groupObject = [adsisearcher]"(&(distinguishedname=$($_)))"  
				$groupObjectProps = $groupObject.FindAll().Properties
				if ($groupObjectProps.objectcategory -like "CN=group*") { 
					recurseGroups $_
				}
				else {
					#If the object is a user, it will add that user to the user object
					if ($groupObjectProps.distinguishedname) {
						$userenabled = ($groupObjectProps.useraccountcontrol[0] -band 2) -ne 2
						if ($userenabled) {
							$dnc = $groupObjectProps.distinguishedname.replace("DC=", ".")
							$validityCheck = $null
							$validityCheck = $groupObjectProps.userprincipalname -and $groupObjectProps.givenname -and $groupObjectProps.sn -and $groupObjectProps.mail -and $groupObjectProps.samaccountname -and $($dnc.substring($dnc.indexof(",.") + 2).replace(",.", "."))
							if ($validityCheck) { $global:UsersList.Add([UserObject]::new($groupObjectProps.userprincipalname, $groupObjectProps.givenname, $groupObjectProps.sn, $groupObjectProps.mail, $groupObjectProps.samaccountname, $($dnc.substring($dnc.indexof(",.") + 2).replace(",.", ".")))) }
						}              
					}
				}
			}
		}
	}
}

$global:UsersList = New-Object -TypeName System.Collections.Generic.List[PSObject]

#setup object for csv
class UserObject {
	[string]$upn
	[string]$fname
	[string]$lname
	[string]$email
	[string]$samaccountname
	[string]$DNSName
	UserObject ([String]$upn, [string]$fname, [string]$lname, [string]$email, [string]$samaccountname, [string]$DNSName) {
		$this.upn = $upn
		$this.FName = $fname
		$this.LName = $lname
		$this.Email = $email
		$this.SAMAccountName = $samaccountname
		$this.DNSName = $DNSName
	}
}

$global:ignoreEmail = $IgnoreEmptyEmail
$rootgroups = [System.Collections.ArrayList]@()

#Check if at least one parameter exists or throw error.
if (!$Groups -and !$Users) { throw "No AD groups or Users listed" }

if ($groups) {
	$groups = $groups.split(',')
	#process each group
	foreach ($groupname in $groups) {
		$groupname = $groupname.trim()
		#Setup LDAP Query
		try {
			$filter = "(&(objectClass=group)(|(Name=$groupname)(CN=$groupname)))"
			$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().name
			$rootEntry = New-Object System.DirectoryServices.DirectoryEntry
			$searcher = [adsisearcher]([adsi]"LDAP://$($rootentry.distinguishedName)")
			$searcher.Filter = $filter
			$searcher.SearchScope = "Subtree"
			$searcher.PageSize = 100000
			#Populate Array with each group
			$rootGroups.add($searcher.FindOne().properties) | Out-Null
		}
		catch { continue }
	}
}
else { if (!$users) { Write-Error "There was no group listed, please input a group name" -ForegroundColor red -BackgroundColor black } }

if ($Users) {
	$users = $users.split(',')
	#process each user
	foreach ($user in $Users) {
		$user = $user.trim()
		#Setup LDAP Query
		try {
			$filter = "(&(objectClass=user)(objectCategory=user)(|(userprincipalname=$user)(samaccountname=$user)))"
			$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().name
			$rootEntry = New-Object System.DirectoryServices.DirectoryEntry
			$searcher = [adsisearcher]([adsi]"LDAP://$($rootentry.distinguishedName)")
			$searcher.Filter = $filter
			$searcher.SearchScope = "Subtree"
			$searcher.PageSize = 100000
			$userProps = $searcher.FindOne().properties
		}
		catch { continue }
		
		#Populate User Object
		if ($userProps.distinguishedname) {
			$userenabled = ($userProps.useraccountcontrol[0] -band 2) -ne 2
			if ($userenabled) {
				$dnc = $userProps.distinguishedname.replace("DC=", ".")
				$validityCheck = $null
				$mail = $userProps.mail
				$validityCheck = $userProps.userprincipalname -and $userProps.givenname -and $userProps.sn -and $userProps.mail -and $userProps.samaccountname -and $($dnc.substring($dnc.indexof(",.") + 2).replace(",.", "."))
				if ($validityCheck) { $global:UsersList.Add([UserObject]::new($userProps.userprincipalname, $userProps.givenname, $userProps.sn, $userProps.mail, $userProps.samaccountname, $($dnc.substring($dnc.indexof(",.") + 2).replace(",.", ".")))) }
			}
		}
	}
}
else { if (!$groups) { Write-Error "There were no users listed, please input a user name" -ForegroundColor red -BackgroundColor black } }

#for each AD group it will kick off the recursion loop to check nested groups and add members to user object
foreach ($rootGroup in $rootGroups) {
	If ($rootGroup.distinguishedname -like '*CN=Builtin*') {
		Write-Error -Message "Group $($rootGroup.distinguishedname) is a Builtin group, only non-Builtin groups and users are supported by this script."
	}
	Else {
		recurseGroups $($rootGroup.distinguishedname)
	}
}

#Only grabs unique users
$global:UsersList = $global:UsersList | Sort-Object upn -Unique
#If user exists, check if there's an export file path. If so, export the CSV. Then print the CSV on the screen.
if ($global:UsersList) {
	if ($exportFile -like "*.*") {
		try {
			$global:UsersList | Export-Csv $exportFile -NoTypeInformation -Confirm:$false -Force
		}
		catch { if ($_) { Write-Error "Unable to Save csv.`n`tFile Path: $exportFile`n`nError:`n$_" } }
	}
	Clear-Host
	$global:UsersList | ConvertTo-Csv -NoTypeInformation | Write-Output
}
