# Name: Recover Provisioning for Horizon Linked Clone Pool

Description: This script acts when provisioning gets disabled for linked clones desktop pools because the overcommit ratio is set too low. It will calculate the correct ratio and set it to that.
After changing the ratio it will enable provisioning and when set to true it can also force a rebalance of the datastores.
When using iwth a trigger the Connection Server FQDN and Horizon Pool name need to be configured manually.

This script requires VMWare PowerCLI to be installed on the machine running the script.
PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers' Or by using the 'Install VMware PowerCLI' script.
Credentials can be set using the 'Prepare machine for Horizon View scripts' script.

Version: 2.1.4

Creator: Wouter Kursten

Date Created: 04/02/2020 13:03:03

Date Modified: 11/24/2023 11:08:48

Scripting Language: ps1

