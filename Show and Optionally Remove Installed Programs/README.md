# Name: Show and Optionally Remove Installed Programs

Description: Show a list of all installed programs for the system and for any logged on users. Search pattern can be used to narrow the list of results (wildcards allowed).
Optionally run the uninstaller for the one package that matches the given string. If more than one package matches the string passed as a parameter, the script will abort because it only allows removal of one application per invocation. (this can be changed in the script)

If the uninstall program is not msiexec there is no guarantee that the unninstall process will run silently and the script may hang because the uninstall process, which cannot show a user interface because it is running in sessionn zero, is trying to show a user interface. If the uninstall process, or any child processes it launches, have not exited by the time that the timeout time, passed as a parameter (non-whole numbers are allowed, e.g. 1.5), is exceeded, the uninstall process and any child processes will be forcibly terminated.

Version: 1.1.26

Creator: guy.leech01

Date Created: 07/26/2021 14:24:07

Date Modified: 02/08/2022 11:43:24

Scripting Language: ps1

