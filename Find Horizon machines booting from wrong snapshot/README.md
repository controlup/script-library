# Name: Find Horizon machines booting from wrong snapshot

Description: The script displays VDI machines and RDS hosts that are not running on the same Golden Image and Snapshot that are configured in the Desktop Pool settings. Uses the Horizon PowerCLI api's to pull all snapshot information for Horizon Linked Clones and Instant Clones Desktops pools and RDS farms.
The script also uses the Horizon api's to poll the Cloud Pod status of the system and connects to other pods if Cloud Pod has been enabled.

This Script Requires a Horizon Credential file for the user running the scipt. This can be created using the Create credentials for Horizon scripts Script Action.
Requires Horizon 7.5 or later
This script requires VMware PowerCLI to be installed on the machine running the script. PowerCLI can be installed using the Install and configure VMware PowerCLI Script Action

Version: 1.6.76

Creator: Wouter Kursten

Date Created: 09/22/2020 09:36:03

Date Modified: 10/13/2020 16:54:47

Scripting Language: ps1

