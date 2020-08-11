# Name: Take ESX Console screenshot

Description: This script uses CreateScreenshot_Task of an ESXi virtual machine through vCenter. The screenshot is the moved from the datastore folder of the VM to a location of choice.
    Screenshots are placed in the virtual machine configuration folder by default. The script moves the screenshot to the desired target folder. For these steps to succeed the account running the script needs the following priviliges:
    1. Virtual Machine - Interaction - Create screenshot
    2. Datastore - Browse Datastore
    3. Datastore - Low level file operations

Version: 1.4.7

Creator: Ton de Vreede

Date Created: 02/26/2020 14:15:55

Date Modified: 08/05/2020 13:28:45

Scripting Language: ps1

