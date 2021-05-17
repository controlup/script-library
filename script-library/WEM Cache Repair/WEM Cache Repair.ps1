#This script is used to troubleshoot WEM Agent refresh issues

$LocalDatabaseDir = 'C:\Program Files (x86)\Norskale\Norskale Agent Host\Local Databases'
$LocalDatabases =Get-ChildItem "C:\Program Files (x86)\Norskale\Norskale Agent Host\Local Databases\*.sdf"

#This section is used to kill the VUEMUIAgent.exe process if it is running
$ProcessName = Get-Process -Name VUEMUIAgent -ErrorAction SilentlyContinue
$Services = "Netlogon","Norskale Agent Host Service"
if ($ProcessName)
  {
   Stop-Process -Name $([string]$ProcessName.ProcessName)
   Write-Host "VUEMUIAgent has stopped"
  }
else
   {
    Write-Host "VUEMUIAgent is not running"
   }

#Stop Norskale and Netlogon services
foreach ($service in $Services)
   {
       Stop-Service -Name $service -Force
       Write-Host "$service has stopped."
   }

#Delete the WEM Agent Cache
cd $LocalDatabaseDir
Remove-Item $LocalDatabases
Write-Host "Files have been deleted"


#Start Netlogon Service which will start Norskale
$WEMAgent = 'C:\Program Files (x86)\Norskale\Norskale Agent Host\VUEMUIAgent.exe'
foreach ($service in $Services)
   {
       Start-Service $service
       Write-Host "$service is running."
   }

Start-Process -FilePath $WEMAgent
