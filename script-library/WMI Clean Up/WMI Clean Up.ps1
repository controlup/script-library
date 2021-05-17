###############################################################################
#
# Checks, fixes and repairs WMI
# Author: Michael Albert info@michlstechblog.info
# changes:
#
# License: GPLv2
#
###############################################################################
# Check for Admin rights
$oIdent= [Security.Principal.WindowsIdentity]::GetCurrent()
$oPrincipal = New-Object Security.Principal.WindowsPrincipal($oIdent)
if(!$oPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator )){
	Write-Warning "Please start script with Administrator Rights! Exit script"
	exit 1
}
# Check PATH
if((! (@(($ENV:PATH).Split(";")) -contains "c:\WINDOWS\System32\Wbem")) -and (! (@(($ENV:PATH).Split(";")) -contains "%systemroot%\System32\Wbem"))){
	Write-Warning "WMI Folder not in search path!."
}
# Stop WMI
# Only if installed
Stop-Service -Force ccmexec -ErrorAction SilentlyContinue 
Stop-Service -Force winmgmt
# WMI Binaries
# [String[]]$aWMIBinaries=@("mofcomp.exe","scrcons.exe","unsecapp.exe","winmgmt.exe","wmiadap.exe","wmiapsrv.exe","wmiprvse.exe")
[String[]]$aWMIBinaries=@("unsecapp.exe","wmiadap.exe","wmiapsrv.exe","wmiprvse.exe","scrcons.exe")
foreach ($sWMIPath in @(($ENV:SystemRoot+"\System32\wbem"),($ENV:SystemRoot+"\SysWOW64\wbem"))){
	if(Test-Path -Path $sWMIPath){
		push-Location $sWMIPath
		foreach($sBin in $aWMIBinaries){
			if(Test-Path -Path $sBin){
				$oCurrentBin=Get-Item -Path  $sBin
				Write-Host " Register $sBin"
				& $oCurrentBin.FullName /RegServer
			}
			else{
				# Warning only for System32
				if($sWMIPath -eq $ENV:SystemRoot+"\System32\wbem"){
					Write-Warning "File $sBin not found!"
				}
			}
		}
		Pop-Location
	}
}
# Reregister Managed Objects
if([System.Environment]::OSVersion.Version.Major -eq 5)
{
	# Windows XP and 2003
   foreach ($sWMIPath in @(($ENV:SystemRoot+"\System32\wbem"),($ENV:SystemRoot+"\SysWOW64\wbem"))){
   		if(Test-Path -Path $sWMIPath){
			push-Location $sWMIPath
			Write-Host " Register WMI Managed Objects"
			$aWMIManagedObjects=Get-ChildItem * -Include @("*.mof","*.mfl")
			foreach($sWMIObject in $aWMIManagedObjects){
				$oWMIObject=Get-Item -Path  $sWMIObject
				& mofcomp $oWMIObject.FullName				
			}
			Pop-Location
		}
   }
   if([System.Environment]::OSVersion.Version.Minor -eq 1){
   		# Windows XP
   		& rundll32 wbemupgd,UpgradeRepository
   }
   else{
		# Windows 2003
		& rundll32 wbemupgd,RepairWMISetup
   }
}
else{
	# Other Windows Vista, Server 2008 or greater
	Write-Host " Reset Repository"
	& ($ENV:SystemRoot+"\system32\wbem\winmgmt.exe") /resetrepository
	& ($ENV:SystemRoot+"\system32\wbem\winmgmt.exe") /salvagerepository
}
Start-Service winmgmt

