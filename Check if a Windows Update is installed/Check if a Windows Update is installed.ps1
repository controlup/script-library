#requires -version 2.0

# ????o?,,?o????????o?,,?o????????o?,,?o????????o?,,?o????
#
#    Author: ? Ajdin Aliu
#    Date  : 09.02.2015
#    Goal  : Get installed Windows Update
#
# ????o?,,?o????????o?,,?o????????o?,,?o????????o?,,?o????



<#
Version  | Date        | Description of Change
---------+-------------+------------------------------------------------------
0.1      | 09.02.2015  | Initial Version
0.2      | 21.02.2015  | use Get-WMIObject instead of Get-HotFix - exception if not found
0.3      | 22.03.2015  | alternative source for update information in Win32_ReliabilityRecords, Error-Handling & verbose output

#>

<#
.SYNOPSIS
	Check if a Windows Update is installed on a system
.DESCRIPTION
	The script checks if a Windows Update is installed, using Get-WMIObject commandlet and WMI information under 'Win32_QuickFixEngineering'. This provider is also used by Get-HotFix cmdlet.
	If no records found in Win32_QuickFixEngineering, then it tries to find records in Win32_ReliabilityRecords containing HotFixID.	
.EXAMPLE
	Get-WmiObject -query "select * from win32_quickfixengineering WHERE HotFixID = 'kb2992611'" -Computername server
	gwmi -cl Win32_ReliabilityRecords -Computername Server | where { $_.message -match "kb2992611"}
.INPUTS
	Computername,HotFixID
.OUTPUTS
	Update installed state, shows some Update Details if Update is found on system
.NOTES
	Requirements 
    - PowerShell 2.0 or greater required
	- Limitations: 
		- Win32_QuickFixEngineering: 			
			- returns only the updates supplied by Component Based Servicing (CBS) - See MSDN article
			- prefered source > faster, but does not contain always all details about updates, that may be present in Win32_ReliabilityRecords
		- Win32_ReliabilityRecords:
			- To access information from this WMI class, the group policy Configure Reliability WMI Providers must be enabled (disabled on server OS by default). See See MSDN article
			- can contain records about failed installations and uninstallations
		
.LINK	
	TechNet Library Get-WMIObject https://technet.microsoft.com/en-us/library/hh849824%28v=wps.620%29.aspx
	TechNet Library Get-HotFix https://technet.microsoft.com/en-us/library/hh849836(v=wps.620).aspx
	MSDN Win32_QuickFixEngineering class https://msdn.microsoft.com/en-us/library/aa394391%28v=vs.85%29.aspx
	MSDN Win32_ReliabilityRecords class https://msdn.microsoft.com/en-us/library/ee706630(v=vs.85).aspx
	TechNet Wiki http://social.technet.microsoft.com/wiki/contents/articles/4197.how-to-list-all-of-the-windows-and-software-updates-applied-to-a-computer.aspx
	MS Customer Service and Support Information https://support.microsoft.com/kb/888535
	See Security Bulletins 2014 for important updates https://technet.microsoft.com/en-us/library/security/dn553321.aspx
#>


$ErrorActionPreference = "Stop" #Treating All Errors as Terminating
#$VerbosePreference = "Continue" #Verbose Output
$UpdateDetails = $null
$CmpName = $args[0]
$HotFixID = $args[1]
$bFound = $False

#check for arguments
If (!$HotFixID -OR !$CmpName){
	#Usage
	Write-Host "HotFixID and Computername required!"	
	EXIT 1
}

#using Win32_QuickFixEngineering
try {	
	Write-Host "Looking for Update in Win32_QuickFixEngineering..."
	#$MyUpdate = Get-HotFix -ID $HotFixID -Computername $CmpName
	$MyUpdate = Get-WmiObject -Query "select * from win32_quickfixengineering WHERE HotFixID = '$($HotFixID)'" -Computername $CmpName	
	If ($MyUpdate -ne $null){
		Write-Host ">> Update found in Win32_QuickFixEngineering."
		$bFound = $True
		# QFEObject: 			System.Management.ManagementObject#root\cimv2\Win32_QuickFixEngineering
		$MyUpdateType = $MyUpdate | Get-Member | Select -ExpandProperty TypeName -Unique		
		$UpdateDetails = $MyUpdate | ForEach {
			[PSCustomObject] @{
				Source 		=	$MyUpdateType
				UpdateID  	=	$_.HotFixID
				SourceName 	=	$_.SourceName
				Description	=	$_.Description
				Details		=	$_.FixComments
				InstalledBy	=	$_.InstalledBy
				#InstalledOn	=	$_.InstalledOn
			}
		}			
	} Else {
		Write-Host ">> Update not found in Win32_QuickFixEngineering."
	}
}
catch {	
	Write-Host "WARN: Could not get Updatelist from Win32_QuickFixEngineering. Check if system is online and you have appropriate permissions!"
	Exit 1
}
finally {
	If ($Error[0]){
		Write-Host "Exception Details for QFE query:"
		Write-Host "Exception: $($Error[0].Exception.Message)"
		Write-Host "ExceptionType: $($Error[0].Exception.GetType().Fullname)"
		Exit 1
	}
}

If (!$bFound){
	#using Win32_ReliabilityRecords
	try {
		Write-Host "Looking for Update in Win32_ReliabilityRecords..."		
		#$MyUpdate = gwmi -cl Win32_ReliabilityRecords -Computername $CmpName | where { ($_.message -NOTmatch "fail" -OR $_.message -NOTmatch "fehl") -AND $_.message -match "$($HotFixID)"}	#| select -last 1
		$MyUpdate = Get-WmiObject -Class Win32_ReliabilityRecords -Computername $CmpName | Where { $_.message -match "$($HotFixID)"}	#| select -last 1
		
		If ($MyUpdate -ne $null){
			Write-Host ">> Update found in Win32_ReliabilityRecords."
			$bFound = $True  
			# ReliabilityObject: 	System.Management.ManagementObject#root\cimv2\Win32_ReliabilityRecords
			$MyUpdateType = $MyUpdate | Get-Member | Select -ExpandProperty TypeName -Unique
			$UpdateDetails = $MyUpdate | ForEach {
				[PSCustomObject] @{
					Source 		=	$MyUpdateType
					UpdateID	=	$_.ProductName
					SourceName	=	$_.SourceName
					Description	=	$_.ProductName
					Details		=	$_.Message
					InstalledBy	=	$_.user
					InstalledOn	=	$_.ConvertToDateTime($_.TimeGenerated)
				}
			}		
		} Else {
			Write-Host ">> Update not found in Win32_ReliabilityRecords."
		}
	}
	catch {
		Write-Host "WARN: Could not get Updatelist from Win32_ReliabilityRecords. Check if system is online, [Win32_ReliabilityRecords] provider is loaded and you have appropriate permissions!!"
	}
    finally {
        If ($Error[0]){			
			Write-Host "Exception Details for ReliabilityRecords query:"
			Write-Host "Exception: $($Error[0].Exception.Message)"
			Write-Host "ExceptionType: $($Error[0].Exception.GetType().Fullname)"
		}
    }
}

If ($bFound) {
	foreach ($UpdateDetail in $UpdateDetails){
		Write-Host "Update Details:"
		Write-Host "=================================================="
		$UpdateDetail
		Write-Host ""
		If ($UpdateDetail.Source.Equals("System.Management.ManagementObject#root\cimv2\Win32_ReliabilityRecords")){
			#Win32_ReliabilityRecords can also contain unsuccessfull installations of updates.
			Write-Host "UpdateDetails for [ $HotFixID ] found. Check UpdateDetails if installed successfully!"
			Write-Host ""
			Write-Host ""
		} else{
			Write-Host "Update [ $HotFixID ] is installed."
		}
	}

} else {
	Write-Host ""
	If ($Error[0]){
		Write-Host "Could not query all providers! See Output for more details!"
	}else{
		Write-Host "Update [$HotFixID] not found."
	}
}

