# Name: Cap Memory of Process

Description: Set the maximum working set size for a process such that it cannot consume more than that amount of memory. Gives the ability to set a memory limit on a process, such as one with a known memory leak, so that it cannot consume more than the maximum memory specified via the parameter. Additional memory allocations will be allowed by the OS paging out some of the existing working set, generally the least recently used.

Version: 1.1.2

Creator: Guy Leech

Date Created: 06/10/2019 13:51:30

Date Modified: 12/12/2019 19:12:48

Scripting Language: ps1

