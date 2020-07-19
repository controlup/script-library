# Name: Test clock synchronization between multiple computers

Description: This action is intended for detecting computers with misaligned system time. It is designed to be executed against a group of computers, comparing the difference between the current time and last midnight, by means of the built-in ControlUp SBA output comparison feature. If the clocks on all selected computers are in sync, you should get one output group.In a typical production-scale environment, a handful of computers may produce a non-standard output, which indicates a clock slip that probably causes authentication errors and other issues.
This script accounts for different time zones, so computers that are in sync but in different time zones should produce the same output.

Version: 2.5.7

Creator: ek

Date Created: 03/12/2019 17:17:31

Date Modified: 03/17/2019 08:22:48

Scripting Language: ps1

