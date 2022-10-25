# Name: AZ Find Potentially Orphaned Azure Resources

Description: Finds resources, either for the resource group the VM the script is run against resides in or all resource groups in the subscription, that are not currently attached to anything.

Note that Citrix MCS machine catalogs may have resources which are not currently assigned but will be assigned when a VM is created in that catalog so should not be removed.

Version: 2.0.7

Creator: Guy Leech

Date Created: 09/17/2022 19:39:53

Date Modified: 09/18/2022 15:12:02

Scripting Language: ps1

