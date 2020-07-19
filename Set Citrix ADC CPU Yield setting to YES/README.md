# Name: Set Citrix ADC "CPU Yield" setting to YES

Description: This Script Action checks if your Citrix ADC appliance is a VPX and if so, changes the cpuyield setting to YES to ensure no High CPU usage is reported on the hypervisor (see https://support.citrix.com/article/CTX229555)

Note: This script only changes the setting on the current ADC and most be performed on both nodes of a HA pair as the setting is node specific and not synchronized.

Version: 1.2.3

Creator: ebarthel

Date Created: 01/05/2020 15:29:41

Date Modified: 02/11/2020 13:21:37

Scripting Language: ps1

