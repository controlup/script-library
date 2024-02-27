# Name: AZ Show Resources Not Used in Time Period

Description: Uses data from the Azure Monitor API to find resources which have not got an entry in the logs in the given period which could mean that they have not been used in that time and be candidates for removal for simplifcation and cost reduction, depending on the resource type. Resources not directly in the logs, such as network interfaces, will be checked by looking for log entries for their parent VM.

Only resources within the same resource group as the chosen resource can be searched or all rescure groups for the tenant depending on the parameters passed.

The output can be sorted/grouped by any of the output columns.

Parameters to specifically include and/or exclude providers by regular expression are available.

Note that not all resources when used will generate an activity in the log so may be shown as no being used when they have actually been used in the given time period.

Version: 3.2.26

Creator: Guy Leech

Date Created: 08/01/2022 23:40:30

Date Modified: 02/25/2024 14:09:57

Scripting Language: ps1

