<#
    .SYNOPSIS
        Enabled the Optimize Drives service

    .DESCRIPTION
        Ensures the Optimize Drives service is available to ensure FSLogix disk compaction can function.

    .LINK
        For more information refer to:
            https://www.controlup.com

    .LINK
        Stay in touch:
        https://twitter.com/trententtye
#>

#Get Optimize Drives service meta data
$DefragSvc = Get-Service -Name defragsvc
if ($DefragSvc.StartType -eq "Disabled") {
    try {
        Set-Service -Name defragsvc -StartupType Automatic
    } catch {
        Write-Error "Unable to modify the service startup type for the Optimize Drives service"
        exit -1
    }
}

$DefragSvc = Get-Service -Name defragsvc
if ($DefragSvc.StartType -ne "Disabled") {
    Write-Output "Optimize Drives service has been tuned to allow for FSLogix Disk Compaction"
} else {
    Write-Error "Unable to modify the service startup type for the Optimize Drives service"
}
