# Name: Get Horizon UAG Health

Description: This script connects to all pod's in a Cloud Pod Architecture(CPA) or only the local one if CPA hasn't been initialized and pulls all health information for configured UAG's. 
This action should be executed against a Horizon endpoint machine (one which has the HZ Primary Connection Server column populated in ControlUp Console). The script uses the target machine to determine the connection server address, and is executed on the machine running ControlUp Console.

Note: Requires PowerCLI 11.4 or higher and Horizon 7.10 or higher

Version: 4.6.9

Creator: Wouter Kursten

Date Created: 01/09/2020 09:32:59

Date Modified: 09/23/2020 07:55:45

Scripting Language: ps1

