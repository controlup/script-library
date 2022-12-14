# Name: Deploy log4net for ControlUp processes

Description: ControlUp binaries generate minimum logging to prevent disk exhaustion. 
 To enable DEBUG logs, deploying the log4net in the same directory as a ControlUp process initiates the capture.
 This script will initiate the logging process and allow to place the log file in a custom location.
 If log duration different than zero, the log capture with stop at the end of the period.

Version: 1.10.49

Creator: Marcel Calef

Date Created: 10/01/2020 17:05:27

Date Modified: 12/13/2022 19:48:39

Scripting Language: ps1

