# Name: Show Horizon Error and Warning Events

Description: Uses the Horizon REST api's to pull all Error, Warning and Audit_Fail events from the Horizon Event database for all pods. If there is no cloud pod setup it will only process the local pod. After pulling the events it will translate the id's for the various objects to names to show the proper names where needed. 

Output is displayed in the console but also saved to a default location of c:windows\temp\CU_Horizon_error_log.csv

This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console). The script uses the target machine to determine the connection server address, and is executed on the machine running ControlUp Console.
This script requires Horizon Credentials to be set for the account running the scipt on the target machine, these need to be created using the 'Create credentials for Horizon scripts' Script Action

Version: 4.3.12

Creator: Wouter Kursten

Date Created: 08/27/2020 09:54:22

Date Modified: 10/25/2023 15:59:20

Scripting Language: ps1

