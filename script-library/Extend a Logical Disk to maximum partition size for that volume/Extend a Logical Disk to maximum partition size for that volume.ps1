<#
   Extend a Logical Disk to maximum partition size for that volume
   
 Leverage PowerShell commands as described in :
   https://docs.microsoft.com/en-us/powershell/module/storage/resize-partition?view=win10-ps
 to extend a logical disk to the maximum available size

   Note   :  PowerShell 4.0 apparently needed, for sure 3.0 required

   Parameter: Drive as provided by ControlUp Logical Disk Drive Name  e.g. C:\
#>

$driveLetter,$discard = $args[0].split(':')

Write-Output "-------------------------------------------------------------- "
Write-Output "Document previous size for $driveLetter"

Get-Partition -DriveLetter $driveLetter
Write-Output "-------------------------------------------------------------- "
Write-Output "Calculating maximum size for drive $driveLetter"
$MaxSize = (Get-PartitionSupportedSize -DriveLetter $driveLetter).sizeMax

Write-Output "-------------------------------------------------------------- "
$MaxSizeGB = [math]::Round($MaxSize/1024/1024/1024,2)
Write-Output "    the maximum size is $MaxSizeGB GB"

Write-Output "-------------------------------------------------------------- "

if ((get-partition -driveletter C).size -eq $MaxSize) {
    Write-Output "The drive $driveLetter is already at its maximum drive size"
    exit
    }

Resize-Partition -DriveLetter $driveLetter -Size $MaxSize

Write-Output "-------------------------------------------------------------- "
Write-Output "Document resulting size for $driveLetter"

Get-Partition -DriveLetter $driveLetter
