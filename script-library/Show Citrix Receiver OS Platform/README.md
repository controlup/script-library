# Name: Show Citrix Receiver OS Platform

Description: Show the client device Operating System type form for a specific user session.
Use this script without requesting the display of the Receiver version to get the results grouped by client OS (available only as a ControlUp Script Based Action). 

Categorization of the Client OS is accomplished by querying the Citrix VDA or XA65 worker for the ClientPlatformId registry value in the appropriate Citrix ICA hive for that user session and follow the conversion described in this document:
https://www.citrix.com/mobilitysdk/docs/clientdetection.html


Version: 2.8.23

Creator: Marcel Calef

Date Created: 12/08/2018 17:27:32

Date Modified: 12/16/2018 13:44:06

Scripting Language: ps1

