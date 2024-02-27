# Name: AZ Find Potentially Orphaned Azure Resources

Description: Finds Azure resources, either for the resource group the VM the script is run against resides in or all resource groups in the subscription, that are not currently attached to a parent resource.

Note that Citrix MCS machine catalogs may have resources which are not currently assigned but will be assigned when a VM is created in that catalog so should not be removed.

The information returned should be cross referenced & checked to another source, such as the Azure portal, before any resources are deleted just because they feature in the output of this script.

Version: 3.8.21

Creator: Guy Leech

Date Created: 09/17/2022 19:39:53

Date Modified: 02/25/2024 14:09:57

Scripting Language: ps1

