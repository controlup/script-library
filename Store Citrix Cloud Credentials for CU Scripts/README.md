# Name: Store Citrix Cloud Credentials for CU Scripts

Description: Create the credential files (locally) required by the Citrix Cloud script actions.
Stored in %ALLUSERSPROFILE%\ControlUp\ScriptingSupport but the client secret stored in the files can only be decrypted by the Windows user that created that file.
The files created contain the tenant id in the file name so that a single Windows user can have credential files for multiple tenants. Original Azure scripts did have this feature and the files contained the tenant id so only a single file can exist - this script creates both credential files so both new and old Azure script actions can be run.
The script willl overwrite any existing credential files for the user and tenant.


Version: 2.1.26

Creator: Guy Leech

Date Created: 02/24/2022 01:42:11

Date Modified: 01/26/2024 14:09:57

Scripting Language: ps1

