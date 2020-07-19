#requires -Version 3.0
$ErrorActionPreference = 'Stop'

<#
    .SYNOPSIS
    This script will display howe much a partition can be shrunk or expanded by.

    .DESCRIPTION
    This script checks all Basic partitions WITH DRIVE LETTERS to see if there is space availbale for sjrinking and/or expansion. The following is also checked:
    - OS Version, this must be Windows 8/Server 2012 as a minimum to use the ControlUp partition expansion script
    - The Disk Status
#>

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
    Out-CUConsole -Message "The operating system is not Windows 8/Server 2012 or greater. The ControlUp expansion script needs these minimum OS versions to work." -Warning
}

# Get the disks
try {
    $objDisks = Get-Disk
}
catch {
    Out-CUConsole -Message "The disks could no be retreived. This could be is a permission issue. It is advised this script is run as either SYSTEM or a Local Administrator account." -Exception $_
}

[array]$arrPartitionInfo = @()
# Loop through disks, get partitions meeting requirements
# Get the partition
Foreach ($Disk in $objDisks) {
    try {
        $PartitionList = $Disk | Get-Partition | Where-Object { $_.DriveLetter -match '[a-z]' }
        foreach ($Partition in $PartitionList) {
            $SizeInfo = Get-PartitionSupportedSize -DiskNumber $Disk.DiskNumber -PartitionNumber $Partition.PartitionNumber
            [double]$CurrentSize = $([math]::round($Partition.Size / 1GB, 2))
            [double]$MinSize = $([math]::round($SizeInfo.SizeMin / 1MB, 2))
            [double]$MaxSize = $([math]::round($SizeInfo.SizeMax / 1MB, 2))
            $objInfo = [pscustomobject]@{
                'Disk #'            = $Disk.DiskNumber
                'Disk HealthStatus' = $Disk.HealthStatus
                'Partition #'       = $Partition.PartitionNumber
                'Drive letter'      = $Partition.DriveLetter
                'Size Gb'           = $CurrentSize
                'Minimum size Mb'   = "$MinSize ($([math]::round($CurrentSize - $MinSize)) Mb smaller)"
                'Maximum size Mb'   = "$MaxSize ($([math]::round($MaxSize - $CurrentSize)) Mb larger)"
                'Disk FriendlyName' = $Disk.FriendlyName
            }
            $arrPartitionInfo += $objInfo
            # Out-CUConsole -Message "Partion number $($Partition.PartitionNumber) (drive letter $($Partition.DriveLetter)) has $([math]::round($dblFreeSpace / 1MB, 2)) Mb of space available for expansion."
        }
    }
    catch {
        Out-CUConsole -Message "There was an issue retreiving partition information from disk number $($Disk.DiskNumber)." -Exception $_
    }
}

# Ouput the information
foreach ($Disk in $objDisks | Sort-Object Number) {
    Out-CUConsole -Message "Disk number $($Disk.Number) (HealthStatus: $($Disk.HealthStatus), Friendly Name: $($Disk.FriendlyName))"
    $arrPartitionInfo | Where-Object 'Disk #' -eq $Disk.Number | Select-Object 'Partition #', 'Drive letter', 'Size Gb','Minimum size Mb','Maximum size Mb' | Format-Table 
}
