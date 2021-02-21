# Name: Show zombie processes

Description: Find processes which are in sessions which no longer exist or where there are no handle or thread objects open by that process which means that it will not be able to run (unless a UWMP process which are excluded).
If the kill parameter is set to true, all of these processes above will be killed.
Also find processes which are trying to exit but are unable to do so because another process has an open process or thread handle to the exiting process which prevents the process from completely exiting.
Zombie processes may be responsible for end user application usability issues and the cause for high session id numbers.
This script should be run a few times, a few minutes apart, in order to rule out processes which were in the process of closing handles they had open.
There is a known handle leak in ControlUp processes prior to release 8.2. With 8.2 there is still a situation that could result in a leak where the ControlUp agent is running on Windows 10 1809 and later or Server 2019 and the User Input Delay feature is enabled by default. See the 8.2 release notes for further information


Version: 3.14.32

Creator: Guy Leech

Date Created: 11/02/2018 20:26:58

Date Modified: 01/05/2021 19:51:35

Scripting Language: ps1

