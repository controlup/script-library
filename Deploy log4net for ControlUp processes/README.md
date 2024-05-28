# Name: Deploy log4net for ControlUp processes

Description: ControlUp binaries by default generate minimum logging to prevent disk exhaustion. 
To enable DEBUG logs, deploying the log4net file in the same directory as a ControlUp process initiates the capture.
This script will initiate the logging process and allows placing of the log file in a custom location.
If log duration is greater than zero, the log capture with stop at the end of the period.

Support Mode changes the format of the lines in the resultant log file to make it easier to merge multiple log files, where required.

Version: 2.10.56

Creator: Marcel Calef

Date Created: 10/01/2020 17:05:27

Date Modified: 05/26/2024 14:09:57

Scripting Language: ps1

