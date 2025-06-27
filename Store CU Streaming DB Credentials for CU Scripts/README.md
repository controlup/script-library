# Name: Store CU Streaming DB Credentials for CU Scripts

Description: Create the credential files (on a ControlUp Monitor machine) required by script actions that use the ControlUp Streaming Database.
Stored in %ALLUSERSPROFILE%\ControlUp\ScriptingSupport but the API key stored in the files can only be decrypted by the Windows user that created that file.
This script should be run on the same machine and as the same Windows user that will run the scripts that interact with the ControlUp Streaming DB.
The script will overwrite any existing CU API key for the user.

Version: 2.0.0

Creator: Guy Leech

Date Created: 06/13/2025 14:09:57

Date Modified: 06/13/2025 14:09:57

Scripting Language: ps1

