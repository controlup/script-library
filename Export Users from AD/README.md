# Name: Export Users from AD

Description: Exports users from active directory to a CSV for solve user import.
Users and groups must be comma separated, if one is not used please have '$null' in the field so it does not attempt to be processed.
A user account should contain the following data, otherwise it will not be exported:	userprincipalname (UPN), givenname, sn (surname), mail (emaill address), samaccountname, distinguishedname
Builtin groups such as Domain Users are not supported
		

Version: 1.0.19

Creator: Steve Schneider

Date Created: 10/11/2022 06:31:44

Date Modified: 10/30/2022 13:33:02

Scripting Language: ps1

