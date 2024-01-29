# Name: Enable or disable Citrix Delivery Groups - On-prem or Cloud

Description: Run on a Delivery Controller or where the CVAD PowerShell snapins are available, e.g. Studio is installed or on a machine with the DaaS Remote PowerShell SDK (for Cloud). User running the script for on-prem must have sufficient permission to change the enabled state of the selected delivery groups. For Cloud, CU stored credentials must have previously been saved for the local user running the script.

To use this script as an automated action where parameters cannot be passed, copy the script and set the $disable parameter in the Param() block at the top of the script to "true" or "false" depending on whether you are disabling or enabling delivery groups respectively.

The Cloud Customer Id or Delivery Controller is an optional argument which can be used when there are more than 1 credential file for the user running the script so that the correct Cloud customer can be chosen. When used on-prem, the parameter is used to tell the script what delivery controller to connect to when the script is not run on a DDC.

Version: 2.1.13

Creator: Guy Leech

Date Created: 10/06/2020 22:00:59

Date Modified: 01/26/2024 14:09:57

Scripting Language: ps1

