#requires -version 5.1

<#
.SYNOPSIS
	Queries Active Directory for Users/Groups and exports a CSV
.DESCRIPTION
	Queries AD Groups/Users for information about each user account and exports it as
	CSV for Solve User Creation
.PARAMETER Users
	User Parameter is comma separated, needs to be samaccountname or upn. Use STRING '$null' if there are no usernames.
		-User "User1,User2,User3"
.PARAMETER Groups
	Groups Parameter is comma separated, needs to be the group name. Use STRING '$null' if there are no group names.
		-Group "Group1,Group2,Group3"
.PARAMETER DisplayWarnings
	Please enter 'true' to display warnings, 'false' to hide them in the output (this is useful if you only want to display the found user results without any clutter in the screen output).
.PARAMETER ExportFile
	ExportFile Parameter is the file path to export the file on the machine where the script is run from.
		-ExportFile "c:\temp\csv.csv"
.EXAMPLE
	.\ExportGroupsForUserCreating.ps1 -Groups "Group1,Group2,Group3" -Users "User1,User2,User3" -ExportFile "c:\temp\csv.csv"
.NOTES
	A user account should contain the following data, otherwise it will not be exported:
	userprincipalname (UPN)
	givenname
	sn (surname)
	mail (email address)
	samaccountname
	distinguishedname
	Builtin groups such as Domain Users are not supported.
#>

[CmdletBinding()]
Param(
	[Parameter(Mandatory = $true, HelpMessage = 'Separate AD Users with Comma.' )]
	[string]$Users,
	[Parameter(Mandatory = $true, HelpMessage = 'Separate AD groups with Comma.' )]
	[string]$Groups,
	[Parameter(Mandatory = $true, HelpMessage = 'Display warnings or not.' )]
	[string]$DisplayWarnings,
	[Parameter(Mandatory = $false, HelpMessage = 'Full path for file export' )]
	[string]$ExportFile
)

# Basic setup
$ErrorActionPreference = 'Stop'
If ($DisplayWarnings -eq 'false') {
	$WarningPreference = 'SilentlyContinue'
}

# Output settings
[int]$outputWidth =550
if ( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) ) {
	$WideDimensions.Width = $outputWidth
	$PSWindow.BufferSize = $WideDimensions
}

# Create lists
$rootgroups = [System.Collections.ArrayList]@()
$UsersList = New-Object -TypeName System.Collections.Generic.List[PSObject]

# Setup object for csv
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

function Get-GroupUsersWithRecurse {
	#This is a recursion loop. If it finds a group inside a group it re-runs the
	#function until there are no sub-groups for that group
	Param (
		[string]$DN
	)
	if ($null -ne $dn) {
		#checks if DN exists in AD, which it should since we queried them already
		if (([adsi]::Exists("LDAP://$DN"))) {
			$group = [adsi]("LDAP://$DN")
			If ($group.member.count -ne 0) {
				#Grabs ADSI object and searches the group for groups
				$group.member | ForEach-Object {
					$groupObject = [adsisearcher]"(&(distinguishedname=$($_)))"
					$groupObjectProps = $groupObject.FindAll().Properties
					if ($groupObjectProps.objectcategory -like "CN=Group,CN=Schema,CN=Configuration,DC=*") {
						Get-GroupUsersWithRecurse $groupObjectProps.distinguishedname
					}
					else {
						#If the object is a user, it will add that user to the user object
						if ($groupObjectProps.distinguishedname) {
							[string]$fullDn = $groupObjectProps.distinguishedname
							If ($null -ne $groupObjectProps.useraccountcontrol) {
									if (($groupObjectProps.useraccountcontrol[0] -band 2) -ne 2) {
									[string]$dnc = $fullDn.replace("DC=", ".")
									$validityCheck = $null
									$validityCheck = $groupObjectProps.userprincipalname -and $groupObjectProps.givenname -and $groupObjectProps.sn -and $groupObjectProps.mail -and $groupObjectProps.samaccountname -and $($dnc.substring($dnc.indexof(",.") + 2).replace(",.", "."))
									if ($validityCheck) {
										$UsersList.Add([UserObject]::new($groupObjectProps.userprincipalname, $groupObjectProps.givenname, $groupObjectProps.sn, $groupObjectProps.mail, $groupObjectProps.samaccountname, $($dnc.substring($dnc.indexof(",.") + 2).replace(",.", "."))))
									}
									Else {
										Write-Warning -Message "User $($fullDn.substring(0,$fullDn.indexof(",DC="))) will not be exported as one or more of the following properties are missing in Active Directory:`nuserprincipalname (UPN),givenname, sn (surname), mail (email address), samaccountname, distinguishedname"
									}
								}
								Else {
									Write-Warning -Message "User $($fullDn.substring(0,$fullDn.indexof(",DC="))) will not be exported as the account is Disabled."
								}
							}
						}
					}
				}
			}
			Else {
				[string]$strGroupName = $DN.Split(',')[0].Replace('CN=', '')
				Write-Warning -Message "The members of group $strGroupName could not be retrieved, either because the group is empty or because it is a BUILTIN group such as Domain Users."
				continue
			}
		}
	}
}

#Check if at least one parameter exists or throw error.
if (($Groups -eq '$null') -and ($Users -eq '$null')) {
	throw "No AD groups or users listed, please specify at least one group or user."
}

if ($groups -ne '$null') {
	[array]$arrGroups = $groups.split(',')
	#process each group
	foreach ($groupname in $arrGroups) {
		$groupname = $groupname.trim()
		#Setup LDAP Query
		try {
			$filter = "(&(objectClass=group)(|(Name=$groupname)(CN=$groupname)))"
			$rootEntry = New-Object System.DirectoryServices.DirectoryEntry
			$searcher = [adsisearcher]([adsi]"LDAP://$($rootentry.distinguishedName)")
			$searcher.Filter = $filter
			$searcher.SearchScope = "Subtree"
			$searcher.PageSize = 100000
			#Populate Array with each group
			$groupToAdd = $searcher.FindOne().properties
			If ($null -ne $groupToAdd) {
				$GroupProperties = $searcher.FindOne().properties
				If ($null -ne $GroupProperties.member) {
					$rootGroups.add($GroupProperties) | Out-Null
				}
				Else {
					Write-Warning -Message "The members of group $GroupName could not be retrieved, either because the group is empty or because it is a BUILTIN group such as Domain Users."
				}
			}
			Else {
				Write-Warning -Message "Group $GroupName could not be found."
			}
		}
		catch {
			continue
		}
	}
}

if ($Users -ne '$null') {
	[array]$arrUsers = $users.split(',')
	#process each user
	foreach ($user in $arrUsers) {
		$user = $user.trim()
		#Setup LDAP Query
		try {
			$filter = "(&(objectClass=user)(objectCategory=user)(|(userprincipalname=$user)(samaccountname=$user)))"
			$rootEntry = New-Object System.DirectoryServices.DirectoryEntry
			$searcher = [adsisearcher]([adsi]"LDAP://$($rootentry.distinguishedName)")
			$searcher.Filter = $filter
			$searcher.SearchScope = "Subtree"
			$searcher.PageSize = 100000
			$userProps = $searcher.FindOne().properties
		}
		catch {
			continue
		}

		#Populate User Object
		if ($userProps.distinguishedname) {
			[string]$fullDn = $userProps.distinguishedname
			if (($userProps.useraccountcontrol[0] -band 2) -ne 2) {
				[string]$dnc = $fullDn.replace("DC=", ".")
				$validityCheck = $null
				$validityCheck = $userProps.userprincipalname -and $userProps.givenname -and $userProps.sn -and $userProps.mail -and $userProps.samaccountname -and $($dnc.substring($dnc.indexof(",.") + 2).replace(",.", "."))
				if ($validityCheck) {
					$UsersList.Add([UserObject]::new($userProps.userprincipalname, $userProps.givenname, $userProps.sn, $userProps.mail, $userProps.samaccountname, $($dnc.substring($dnc.indexof(",.") + 2).replace(",.", "."))))
				}
				Else {
					Write-Warning -Message "User $($fullDn.substring(0,$fullDn.indexof(",DC="))) will not be exported as one or more of the following properties are missing in Active Directory:`nuserprincipalname (UPN),givenname, sn (surname), mail (email address), samaccountname, distinguishedname"
				}
			}
			Else {
				Write-Warning -Message "User $($fullDn.substring(0,$fullDn.indexof(",DC="))) will not be exported as the account is Disabled."
			}
		}
		Else {
			Write-Warning -Message "User $User could not be found."
		}
	}
}


#for each AD group it will kick off the recursion loop to check nested groups and add members to user object
foreach ($rootGroup in $rootGroups) {
	Get-GroupUsersWithRecurse $rootGroup.distinguishedname
}

#If user exists, check if there's an export file path. If so, export the CSV. Then print the CSV on the screen.
if ($UsersList.Count -ne 0) {
	[int]$WriteSuccess = 0
	#Only grabs unique users
	$UsersList = $UsersList | Sort-Object upn -Unique
	if ($exportFile -like "*.*") {
		try {
			$UsersList | Export-Csv $exportFile -NoTypeInformation -Confirm:$false -Delimiter ',' -Force
		}
		catch {
			$WriteSuccess = 1
			Write-Error -Message "Unable to Save csv.`n`tFile Path: $exportFile`n`nError:`n$_" -ErrorAction Continue
		}
	}
	$UsersList | ConvertTo-Csv -NoTypeInformation -Delimiter ',' | Write-Output
}
Else {
	Write-Warning -Message 'No users were found in the users or groups provided. Please check your input.'
}
Exit $WriteSuccess
