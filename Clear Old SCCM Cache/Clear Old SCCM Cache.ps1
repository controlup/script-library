<#
 	.SYNOPSIS
        A simple script that clears all items older than 7 days in the SCCM update cache.
		
    .DESCRIPTION
        Since Microsoft changed thier approach to Windows Updates, the size of the patches has increased significantly. This can be
        challenging to manage on persistent machines with limited disk space. This script is designed to clear that cache to free up
        some valuable disk space.
    
    .NOTES
        The scrips perform a check to ensure the require COM object exists. If it does not exist, no action will be taken.
		
    .LINK
        For more information refer to:
            http://www.controlup.com

    .LINK
        Stay in touch:
        http://twitter.com/rorymon

    .EXAMPLE
        C:\PS>\. ClearSCCMCache.ps1
		
		Clears SCCM Update cache items older than 7 days.
#>

## Last modified 13:33 GMT 20/04/21 @rorymon

$resman = new-object -com "UIResource.UIResourceMgr"
$cacheInfo = $resman.GetCacheInfo()

if ($resman) {
$cacheinfo.GetCacheElements()  | 
where-object {$_.LastReferenceTime -lt (get-date).AddDays(-7)} | 
foreach {
$cacheInfo.DeleteCacheElement($_.CacheElementID)
}
} else {
write-host ("Required COM Object Does Not Exist. Ensure SCCM Client is Installed.")
}


