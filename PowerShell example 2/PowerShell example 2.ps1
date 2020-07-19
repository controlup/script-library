<#
If you are interested in using commands like Write-Verbose that need [CmdletBinding()] 
or using parameter validation techniques, you must use the position arguments in order to 
process any ControlUp-provided arguments properly. See the example below.
#>

[CmdletBinding()]
param(
   [parameter(Position=0)]
   [string]$name
# (this is equivalent to $name = $args[0] if you are not using CmdletBinding.)
)

$drivetype = '3'

Write-Verbose "`$name = $name"

Get-WmiObject -class Win32_LogicalDisk -computername $name -filter "drivetype=$drivetype" |
 Sort-Object -property DeviceID |
 Format-Table -property DeviceID,
     @{l='FreeSpace(MB)';e={$_.FreeSpace / 1MB -as [int]}},
                  @{l='Size(GB)';e={$_.Size / 1GB -as [int]}},
   @{l='%Free';e={$_.FreeSpace / $_.Size * 100 -as [int]}}

