function Invoke-MSPatchCheckForCitrix {
    <#
    .SYNOPSIS
        Check the installed MS Patches on a Windows Server 2012 machine.
    .DESCRIPTION
        Check the installed MS Patches on a Windows Server 2012 machine.
        Remake of Christian Edwards script to make it more flexible. And Niklas Akerlund 2013-06-28.
        Mlodified further by ControlUp.
    .EXAMPLE
        Invoke-MSPatchCheckForCitrix -HostName $args[0] -OnlyDisplayNotInstalledPatches $args[1]
    .CONTEXT
        Machine - Target Machine
    .MODIFICATION_HISTORY
        Zeev Eisenberg      - 13/10/15 - Original code
        Esther Barthel, MSc - 24/09/19 - Adding extra Windows Server 2012 (OS Type) check
        Esther Barthel, MSc - 24/09/19 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
    .LINK
        http://blogs.technet.com/b/cedward/archive/2013/05/31/validating-hyper-v-2012-and-failover-clustering-hotfixes-with-powershell-part-2.aspx
    .COMPONENT
    .NOTES
        Version:        0.2
        Author:         Esther Barthel, MSc
        Creation Date:  2019-09-24
        Updated:        2019-09-24
                        Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
        Purpose:        Script Action, created for ControlUp
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the HostName of the machine to run the script on'
        )]
        [string] $HostName,
      
        [Parameter(
            Position=1, 
            Mandatory=$false, 
            HelpMessage='Specify if only the Not installed hotfixes should be shown'
        )]
        [string] $ShowOnlyNotInstalled
      
#        [Parameter(
#            Position=2, 
#            Mandatory=$false, 
#            HelpMessage='Specify if the hotfixes need to be downloaded'
#        )]
#        [string] $DownloadHotfixes
#        [Parameter(
#            Position=3, 
#            Mandatory=$false, 
#            HelpMessage='Specify the download path'
#        )]
#        [string] $DownloadPath
    )    

    If ($ShowOnlyNotInstalled -match "^(yes|y)$") {
        $OnlyNotInstalled = $true
    }
#    If ($DownloadHotfixes -match "^(yes|y)$") {
#        $Download = $true
#    }

    $HotfixList = @()

    # Check if the OS Type is a server OS (exclude Windows 8 and 8.1)
    $strOS=(((Get-WMIObject Win32_OperatingSystem).name).Split("|")[0]).ToString()
    If ($strOS -like "*Windows Server 2012*") 
    {
        Write-Host "OS matches, MS Patch will check patches for $strOS`!" -ForegroundColor Green
        $Ver = (Get-WmiObject Win32_OperatingSystem).Version
        # $Ver = 6.3.9600 = Windows Server 2012 R2 (or Windows 8.1)
        # $Ver = 6.2.9200 = Windows Server 2012 RTM (or Windows 8)

        If ($Ver -ge "6.3.9600") {
            $OSVer = "R2"
        } ElseIf ($Ver -ge "6.2.9200") {
            $OSVer = "RTM"
        } Else {
            Write-Warning "The OS version $Ver was not recognized."
            Exit 1
        }
    } 
    Else 
    {
        Write-Warning "MS Patch cannot check patches for $strOS" 
        Exit
    }


    #region Hotfixes
    # This list of hotfixes is only for Windows Server 2012/2012 R2
    $HF1 = @{"Name" = "December 2014 update rollup for Windows Server 2012 R2";
	    "ID" = "KB3013769";
	    "URL" = "https://download.microsoft.com/download/D/6/1/D6129EA3-CA55-47C8-9276-1F482A4C3CB3/Windows8.1-KB3013769-x64.msu"
    }
    $HF2 = @{"Name" = "Remote Desktop session freezes when you run an application in the session in Windows Server 2012 R2";
    #	"ID" = "KB2978367";  # This is just for identifying the issue. KB2975719 has the fix.
    #	KB2978367 is fixed by KB2975719, the August 2014 rollup (which requires KB2993651)
	    "ID" = "KB2975719";
	    "URL" = "https://download.microsoft.com/download/4/8/6/4863950A-2FE5-421F-B66B-719554CC463F/Windows8.1-KB2975719-x64.msu"
    }
    switch ($OSVer) {
	    "RTM" {
		    $HF3 = @{"Name" = "A network printer is deleted unexpectedly in Windows";
			    "ID" = "KB2967077"; # for 2012 RTM - private hotfix, must be requested
			    "URL" = "A hotfix must be requested from Microsoft from the KB page."
		    # The actual patch is only for Windows Server 2012 R2. To resolve this issue in Windows Server 2012, a private hotfix must be requested from Microsoft.
		    }
	    }
	    "R2" {
		    $HF3 = @{"Name" = "A network printer is deleted unexpectedly in Windows";
			    "ID" = "KB2975719"; # for 2012 R2
			    "URL" = "https://download.microsoft.com/download/4/8/6/4863950A-2FE5-421F-B66B-719554CC463F/Windows8.1-KB2975719-x64.msu"
		    # This patch is only for Windows Server 2012 R2. To resolve this issue in Windows Server 2012, a private hotfix must be requested from Microsoft.
		    }
	    }
    }
    If ($OSVer -eq "R2") {
	    $HF4 = @{"Name" = "Users who have the remote audio setting enabled cause the RD Session Host servers to freeze intermittently in Windows Server 2012 R2";
		    # "ID" = "KB2895698" # This is just for identifying the issue. KB2919335 has the fix.
		    "ID" = "KB2919355"; # "Windows Server 2012 R2 update: April 2014 - for 2012 R2 only, not 2012 RTM
		    "URL" = "https://download.microsoft.com/download/2/5/6/256CCCFB-5341-4A8D-A277-8A81B21A1E35/Windows8.1-KB2919355-x64.msu"
	    }
    }
    If ($OSVer -eq "RTM") {
	    $HF5 = @{"Name" = "FIX: You are logged on with a temporary profile to a remote desktop session after an unexpected restart of Windows Server 2012";
		    "ID" = "KB2896328"; # for 2012 RTM only - private hotfix, must be requested
		    "URL" = "A hotfix must be requested from Microsoft from the KB page."
	    }
    }
    If ($OSVer -eq "RTM") {
	    $HF6 = @{"Name" = "Memory leak occurs in the Dwm.exe process on a Remote Desktop computer that is running Windows Server 2012";
		    # "ID" = "KB2852483"; # This is just for identifying the issue. for 2012 RTM only, Windows Server 2012 update rollup: July 2013
		    "ID" = "KB2855336";
		    "URL" = "https://download.microsoft.com/download/5/C/6/5C652B4A-F479-4210-A102-832CD9E48AE6/Windows8-RT-KB2855336-x64.msu"
	    }
    }
    If ($OSVer -eq "R2") {
	    $HF7 = @{"Name" = "October 2014 update rollup for Windows Server 2012 R2";
		    "ID" = "KB2995388"; # for 2012 R2 only, requires KB2919355.
		    "URL" = "https://download.microsoft.com/download/E/A/A/EAA9D07C-9C11-480C-BB9E-BDE7385BD541/Windows8.1-KB2995388-x64.msu"
	    }
    }

    #endregion Hotfixes

    $MSHotfixList = $HF1,$HF2,$HF3,$HF4,$HF5,$HF6,$HF7

    #Getting installed Hotfixes from server
    $InstalledHotfixList = Get-HotFix -ComputerName $Hostname | Select HotfixID

    foreach($RecommendedHotfix in $MSHotfixList)
    {
	    If ($RecommendedHotfix.ID -ne $null) {
	        $witness = 0
	        foreach($InstHotfix in $InstalledHotfixList)
	        {
	            If($RecommendedHotfix.ID -eq $InstHotfix.HotfixID)
	            {
	                $obj = [PSCustomObject]@{
	                    CitrixNode = $Hostname
	                    RecommendedHotfix = $RecommendedHotfix.ID
	                    Status = "Installed"
	                    Description = $RecommendedHotfix.Name
	                    DownloadURL =  $RecommendedHotfix.URL
	                } 
				    $HotfixList += $obj
				    $witness = 1
	             }
	        }
		    If ($witness -eq 0)
	        {
	            $obj = [PSCustomObject]@{
	                    CitrixNode = $Hostname
	                    RecommendedHotfix = $RecommendedHotfix.ID
	                    Status = "Not Installed"
	                    Description = $RecommendedHotfix.Name
	                    DownloadURL =  $RecommendedHotfix.URL
	                    DownloadURL2 =  $RecommendedHotfix.URL2
	            }
	            $HotfixList += $obj
	        }
	    }
    }

    <# possible future use?
    If ($Download){
        foreach($RecommendedHotfix in $MSHotfixList){
            if ($RecommendedHotfix.DownloadURL -match "http" -and $RecommendedHotfix.Status -eq "Not Installed"){
                Start-BitsTransfer -Source $RecommendedHotfix.DownloadURL -Destination $DownloadPath 
            }
        }
    }
    #>

    If ($OnlyNotInstalled) {
        $HotfixList | Where {$_.Status -eq "Not Installed"} | Sort RecommendedHotFix | ft RecommendedHotFix,Status,Description -AutoSize
    } Else { $HotfixList | sort RecommendedHotFix | ft RecommendedHotFix,Status,Description -AutoSize }
}

$HostnameInput = $args[0]
$ShowOnlyNotInstalledInput = $args[1]
#$DownloadInput = $args[2]
#$DownloadPathInput = $args[3]

Invoke-MSPatchCheckForCitrix -HostName $HostnameInput -ShowOnlyNotInstalled $ShowOnlyNotInstalledInput
