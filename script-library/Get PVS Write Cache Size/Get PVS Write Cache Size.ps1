<#
    PVS Cache in Ram Size Script 
    Written by Matthew Nichols 2015
    Get me at Twitter @Mattnics and at http://Mattnics.com
    See http://blogs.citrix.com/2014/04/18/turbo-charging-your-iops-with-the-new-pvs-cache-in-ram-with-disk-overflow-feature-part-one
        for more information.
    Get-IniContent() courtesy of https://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91
#>

Function Get-IniContent {

[CmdletBinding()]  
Param(  
    [ValidateNotNullOrEmpty()]  
    [ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -eq ".ini")})]  
    [Parameter(ValueFromPipeline=$True,Mandatory=$True)]  
    [string]$FilePath  
)  

$ini = @{}  
        switch -regex -file $FilePath  
        {  
            "^\[(.+)\]$" # Section  
            {  
                $section = $matches[1]  
                $ini[$section] = @{}  
                $CommentCount = 0  
            }  
            "^(;.*)$" # Comment  
            {  
                if (!($section))  
                {  
                    $section = "No-Section"  
                    $ini[$section] = @{}  
                }  
                $value = $matches[1]  
                $CommentCount = $CommentCount + 1  
                $name = "Comment" + $CommentCount  
                $ini[$section][$name] = $value  
            }   
            "(.+?)\s*=\s*(.*)" # Key  
            {  
                if (!($section))  
                {  
                    $section = "No-Section"  
                    $ini[$section] = @{}  
                }  
                $name,$value = $matches[1..2]  
                $ini[$section][$name] = $value  
            }  
        }  
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Finished Processing file: $FilePath"  
        Return $ini  
}

$server = $args[0]

If (!(Test-Path HKLM:System\CurrentControlSet\Services\bnistack\pvsagent)) {
    Write-Host "This computer does not use PVS. Please pick another computer."
    Exit 1
}

$content = Get-IniContent "C:\Personality.ini"
$CacheType = $content["StringData"]["`$WriteCacheType"]
$CacheDrive = (Get-ItemProperty HKLM:System\CurrentControlSet\Services\bnistack\pvsagent).WriteCacheDrive

# Percent Free Disk space (is the cache drive in danger of being full?)
$Disk = Get-WmiObject -class Win32_LogicalDisk -computername $Server -filter "DeviceID='$CacheDrive'"
$DiskFreePercent = [System.Math]::Round($Disk.freespace / $Disk.size * 100, 1)

If ($CacheType -eq "4") {
    # Cache on hard drive only
    $PvsWriteCache   = "$CacheDrive\.vdiskcache"

    # CacheDiskOverflowSize
    $CacheDiskMB = [Math]::Round((Get-Item $PvsWriteCache -Force).length/1MB)

    Write-Host "PVS Cache type = hard disk only"
    Write-Host "vDisk Cache file: $PvsWriteCache"
    Write-Host "vDisk Cache Drive free space: $DiskFreePercent %"
    Write-Host "vDisk Cache file size: $CacheDiskMB MB"
} ElseIf ($CacheType -eq "9") {
    # RAM Cache with disk overflow
    $PvsWriteCache   = "$CacheDrive\vdiskdif.vhdx"

    # CacheDiskOverflowSize
    $CacheDiskMB = [Math]::Round((Get-Item $PvsWriteCache -Force).length/1MB)

    # NonPaged Pool Memory (RAM Cache in use) adjusted for likely kernel usage
    $NPPM = [math]::Round((Get-WmiObject Win32_PerfFormattedData_PerfOS_Memory -ComputerName $server).PoolNonPagedBytes /1MB)

    Write-Host "PVS Cache type = RAM cache with disk overflow"
    Write-Host "vDisk Cache file: $PvsWriteCache"
    Write-Host "vDisk Cache file size: $CacheDiskMB MB"
    Write-Host "vDisk Cache Drive free space: $DiskFreePercent %"
    Write-Host "vDisk RAM Cache usage: $NPPM MB"
} Else {
    Write-Host "The disk cache type is not supported in this script. Please choose another computer and try again."
}

