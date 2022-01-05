# Name: Analyze Logon Duration

Description: Get a detailed overview of the most recent logon process for a specific user. This script queries the event log for every major event that relates to the logon process. Use this action to track down which phase is responsible for delays during the logon process. Uses WMI to retrieve pre-Windows logon phase data from Citrix so does not use OData and therefore does not need credentials

Version: 10.25.150

Creator: Guy Leech

Date Created: 07/02/2018 18:55:36

Date Modified: 11/30/2021 17:04:45

Scripting Language: ps1

