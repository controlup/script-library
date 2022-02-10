# Name: Warn user of high memory use

Description: This is an example of how the 'Show message dialogue' script can be tailored for a specific use case.
_ClientMetric1 has been set to the Memory (Working Set) from ControlUp and the message includes this metric.
Use WPF to display a modal dialogue in a user's session with a message, optional title, and solid colour background.
Take a copy of the script and either add parameters to use as a regular SBA,  or set the defaults in the Param() block so it can be used in an automated action.  If you have created parameters, setting default values for these will also allow use in an automated action.
Unless -showForSeconds is specified and the value is less than the script timeout, it will timeout and show an eror but if no other errors are shown, the message will have been displayed

Version: 1.11.48

Creator: Guy Leech

Date Created: 05/07/2021 18:37:22

Date Modified: 02/07/2022 15:57:44

Scripting Language: ps1

