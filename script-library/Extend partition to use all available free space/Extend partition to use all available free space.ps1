#requires -Version 3.0
$ErrorActionPreference = 'Stop'
Set-Strictmode -Version Latest
<#
    .SYNOPSIS
    This script will increase the size of the named drive letter

    .DESCRIPTION
    This script uses standard PowerShell commands to expand the disk of Windows machine. For safety, this script only works if the following conditions are met:
    - At least 100Mb of free space available (this needs to be directly AFTER the chosen partition)
    - The Disk State must be Healthy
    - You must specify a drive letter of the disk to be expanded

    .NOTES
    Though expanding partition size is a standard operation, it is recommended you make a backup of the disk if it contains any vital information. Also, bear in mind that disk expansion can be a very slow process, try to schedule it for non-business hours and be patient.
    This script only works on Windows 8/Server 2012 or later.
#>

# Drive letter of partition to be expanded
[string]$strDriveLetter = $args[0]

Function Out-CUConsole {
    <# This function provides feedback in the console on errors or progress, and aborts if error has occured.
      If only Message is passed this message is displayed
      If Warning is specified the message is displayed in the warning stream (Message must be included)
      If Stop is specified the stop message is displayed in the warning stream and an exception with the Stop message is thrown (Message must be included)
      If an Exception is passed a warning is displayed and the exception is thrown
      If an Exception AND Message is passed the Message message is displayed in the warning stream and the exception is thrown
    #>

    Param (
        [Parameter(Mandatory = $false)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [switch]$Warning,
        [Parameter(Mandatory = $false)]
        [switch]$Stop,
        [Parameter(Mandatory = $false)]
        $Exception
    )
    # Throw error, include $Exception details if they exist
    if ($Exception) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        If ($Message) {
            Write-Warning -Message "$Message`n$($Exception.CategoryInfo.Category)`nPlease see the Error tab for the exception details."
            Write-Error "$Message`n$($Exception.Exception.Message)`n$($Exception.CategoryInfo)`n$($Exception.Exception.ErrorRecord)" -ErrorAction Stop
        }
        Else {
            Write-Warning "There was an unexpected error: $($Exception.CategoryInfo.Category)`nPlease see the Error tab for details."
            Throw $Exception
        }
    }
    elseif ($Stop) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        Write-Warning -Message "There was an problem.`n$Message"
        Throw $Message
    }
    elseif ($Warning) {
        # Write the warning to Warning stream, thats it. It's a warning.
        Write-Warning -Message $Message
    }
    else {
        # Not an exception or a warning, output the message
        Write-Output -InputObject $Message
    }
}

# Test OS version first
[version]$verMinimumWindows = 6.2
If ($verMinimumWindows -gt [System.Environment]::OSVersion.Version) {
    Out-CUConsole -Message "This script only works on Windows 8/2012 or greater."
    Exit 1
}

# Get the partition
try {
    $objPartition = Get-Partition -DriveLetter $strDriveLetter
}
catch {
    Out-CUConsole -Message "The drive/partition could not be retreived. The drive letter could not exist, or this is a permission issue. It is advised this script is run as either SYSTEM or a Local Administrator account." -Exception $_
}

# Check if file system is NTFS or exFAT, as other file systems do not support resizing with the standard Windows tools.
$FileSystemType = ($objPartition | Get-Volume).FileSystemType
if ($FileSystemType -notin 'NTFS', 'exFAT') {
    Out-CUConsole -Message "Partition uses $FileSystemType. Resizing this file system type is not possible with standard Windows tools." -Stop
}

# Get the disk
try {
    $objDisk = $objPartition | Get-Disk 
}
catch {
    Out-CUConsole -Message "There was an issue retreiving the disk drive $strDriveLetter on the machine. This is likely a permissions issue, it is advised this script is run as either SYSTEM or a Local Administrator account." -Exception
}

# Get the possible supported size for the partition
try {
    $MaxSize = (Get-PartitionSupportedSize -DiskNumber $objDisk.DiskNumber -PartitionNumber $objPartition.PartitionNumber).SizeMax
}
catch {
    Out-CUConsole -Message "There was an issue retreiving the maximum supported size partition $($objPartition.PartitionNumber) on disk $($objDisk.DiskNumber) (drive letter $($objPartition.DriveLetter)`:`). This is likely a permissions issue, it is advised this script is run as either SYSTEM or a Local Administrator account." -Exception $_
}

# Check if the size can be increased. Resizing will only be done if there is a minimum of 100Mb available
[double]$dblFreeSpace = $MaxSize - $objPartition.Size
If ($([math]::round($dblFreespace / 1MB, 2)) -lt 100) {
    Out-CUConsole -Message "This script requires that there is a minimum of 100Mb of free space to increase the partition size, but there is only $([math]::round($dblFreeSpace/1MB, 2)) megabytes of free space available."
    Exit 1
}

# Resize the partition
try {
    Resize-Partition -DiskNumber $objDisk.DiskNumber -PartitionNumber $objPartition.PartitionNumber -Size $MaxSize
    Out-CUConsole -Message "Partition $($objPartition.PartitionNumber) on disk $($objDisk.DiskNumber) (drive letter $($objPartition.DriveLetter)`:`) resized from $([math]::round($objPartition.Size/1GB, 2)) Gb to $([math]::round($MaxSize/1GB, 2)) Gb ($($MaxSize - $objPartition.Size) bytes)."
}
catch {
    Out-CUConsole -Message "There was an unexpected error while attempting to resize partition $($objPartition.PartitionNumber) on disk $($objDisk.DiskNumber) (drive letter $($objPartition.DriveLetter)`:`)." -Exception $_
}
