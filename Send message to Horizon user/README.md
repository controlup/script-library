# Name: Send message to Horizon user

Description: Sends messages to the selected Horizon user/s. This script can be used to send messages to a single user session using the Horizon SOAP API's. It can also be used as an automated action with a fixed message and severity level. For the severity level these are allowed: INFO,WARNING, ERROR

This script requires VMWare PowerCLI to be installed on the machine running the script.
PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers' Or by using the 'Install VMware PowerCLI' script.
If you get TLS/SSL errors use this command Set-PowerCLIConfiguration -InvalidCertificateAction ignore
   or Set-PowerCLIConfiguration -InvalidCertificateAction warn
To get rid of the CEIP warning use Set-PowerCLIConfiguration -ParticipateInCeip $true 
   or Set-PowerCLIConfiguration -ParticipateInCeip $false
Credentials can be set using the 'Create credentials for Horizon View scripts' script.

Version: 2.4.6

Creator: Wouter Kursten

Date Created: 07/07/2020 10:52:51

Date Modified: 11/24/2023 11:10:38

Scripting Language: ps1

