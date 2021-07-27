# Name: Horizon Recover Instant Clone RDS Machine

Description: This script Recovers a Horizon View Instant Clone using the VMware Horizon api's. You can use this to 'rebuild' an Instant Clone if there is an issue with the machine. 
This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console). The script uses the target machine to determine the connection server address, and is executed on the machine running ControlUp Console.

Version: 1.1.3

Creator: wouter.kursten

Date Created: 05/18/2021 12:21:00

Date Modified: 05/18/2021 10:49:34

Scripting Language: ps1

