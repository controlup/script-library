# Name: Generate Network Trace

Description: Produces a native Windows network trace in etl format, downloads a signed Microsoft utility from GitHub and runs that to convert the etl file to a pcapng file that Wireshark can open.
After successful conversion the source etl trace will be deleted, leaving the converted file whose name and location will be in the output window.
Specifying output files on network shares may not work as the script needs to run as system so may not have access.
Address filters can be a comma separated list of IP addresses or resolveable names or leave blank to not filter.
Protocol filters can be a comma separated list of TCP, UDP, ICMP, IGMP or leave blank to not filter.
Ether type filters can be a comma separated list of IPv4, IPv6, ARP, etc or leave blank to not filter.
https://github.com/microsoft/etl2pcapng

Version: 1.2.20

Creator: Guy Leech

Date Created: 10/02/2024 20:50:37

Date Modified: 01/12/2025 14:09:57

Scripting Language: ps1

