# Name: Reduce Session Bandwidth Consumption

Description: When connecting to a remote session, the more graphics data sent the more bandwidth is needed.
While modern protocols like HDX and Blast adjust, when task workers have extremely low bandwidth available (less than 1 or 2 Mbps) they might benefit from a simple trick: Reduce the maximum frames per second sent instead of sacrificing image quality and/or responsiveness.

This script modifies the VMware Blast_VDI and Citrix sesison registry keys and therefore it is recommended to implement policies that would revert this to the desired defaults (periodically or upon reconnect)

Version: 4.5.14

Creator: mc

Date Created: 05/13/2019 16:50:03

Date Modified: 10/17/2019 03:21:25

Scripting Language: BAT

