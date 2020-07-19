# Name: Redirect URLs to 127.0.0.1 in HOSTS file

Description: This script can be used as a quick way to block access to a URL. The provided URLs (comma separated) are placed in a 'ControlUp' section of the HOSTS file, where they are directed to 127.0.0.1. As a result DNS lookup of these URLs always point to home, essentially preventing access to a website unless you know the IP number.
After this the command IPCONFIG /FLUSHDNS is run to clear the DNS cache.

Version: 1.4.19

Creator: Ton de Vreede

Date Created: 03/30/2020 00:06:26

Date Modified: 03/30/2020 15:50:39

Scripting Language: ps1

