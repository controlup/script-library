# Name: Process CPU Usage Limit

Description: Finds threads over consuming CPU in the selected process and reduces their average CPU consumption based on the agressiveness argument. The higher the agressiveness, the more CPU throttling is performed. The number can be between 1 and 10 including decimal places.
A duration can be set, in minutes or parts there of, for how long the selected process will be monitored/adjusted but if set to 0 then the process will be monitored/adjusted until it exits.

WARNING: This may make interactive applications become sluggish for users if they are targeted

Version: 1.2.7

Creator: guy.leech

Date Created: 10/03/2019 08:55:15

Date Modified: 10/29/2019 21:15:31

Scripting Language: ps1

