# Name: Show Horizon admin audit trail

Description: Uses the Horizon REST api's to pull all admin related events from the Horizon Event database for all pods. If there is no cloud pod setup it will only process the local pod. After pulling the events it will translate the id's for the various objects to names to show the proper names where needed.

Output is displayed in the console but also saved to a default location of c:windows\temp\CU_Horizon_audit_log.csv

Version: 4.8.21

Creator: Wouter Kursten

Date Created: 04/28/2020 13:00:55

Date Modified: 10/25/2023 14:14:04

Scripting Language: ps1

