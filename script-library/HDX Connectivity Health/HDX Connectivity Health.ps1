<#  
.SYNOPSIS
        The script makes an HDX connection to a specific XenApp server (6.0 and above) in order to check the server's health.
.DESCRIPTION
        The script runs on the ControlUp Console computer and initiates an HDX connection against one or more
        XenApp servers in order to check the server's health. The script does not perform any actual login activity
        to the targeted XenApp server. After the connection is initiated the login connection will disappear after 40 
        seconds.
.PARAMETER ServerName
        This script gets only one parameter which is the server name. The server name is the $arg[0] parameter and is mandatory.
.EXAMPLE
        XenAppHDXConnectivty.ps1 XenAppServer005
.OUTPUTS
        Text describing the connection launch status
.LINK
        See http://www.ControlUp.com
#>

$ServerName = $args[0]
$PF = $env:ProgramFiles
$PF86 = ${env:ProgramFiles(x86)}
#Check if the ICA Client is installed on the console computer
If ((Test-Path "$PF\Citrix\ICA Client\WfIcaLib.dll") -or (Test-Path "$PF86\Citrix\ICA Client\WfIcaLib.dll")) {
    # Check bitness. If this is a 64-bit system, then put the rest of the script in a file and re-launch PoSh in 32-bit mode to run it.
    if ($env:PROCESSOR_ARCHITECTURE -eq "AMD64"){
    $a = @'
param ([string]$ServerName)
$PF86 = ${env:ProgramFiles(x86)}
[System.Reflection.Assembly]::LoadFile("$PF86\Citrix\ICA Client\WfIcaLib.dll")| Out-Null 
write-host "Please Wait... Initiating Connection."
"Do not close this window, it will disapear within 30 seconds"
$ICA = New-Object WFICALib.ICAClientClass
$ICA.Address = $ServerName
$ICA.Application = ""
$ICA.OutputMode = [WFICALib.OutputMode]::OutputModeWindowless
$ICA.Launch = $true
$ICA.TWIMode = $true
$ICA.Connect()
Start-Sleep -Seconds 30
$ICA.GetErrorMessage($ICA.GetLastError()) | Out-File $env:APPDATA\$ServerName-ErrSrv.txt -Encoding UTF8
'@
        $a | Out-File $env:APPDATA\$serverName.ps1
        Start-Sleep -Seconds 2
        Start-Process $env:SystemRoot\syswow64\WindowsPowerShell\v1.0\powershell.exe -ArgumentList " -ExecutionPolicy RemoteSigned -Command $env:APPDATA\$ServerName.ps1 -servername $ServerName" -Wait
        Start-Sleep -Seconds 2
        $b = Get-Content $env:APPDATA\$ServerName-ErrSrv.txt
        if($b -eq "No error"){
              write-host "Connection established successfuly!"
       }
       else{
              write-host "Something went wrong: $b"
              exit 1
       }
        Remove-Item $env:APPDATA\$ServerName-ErrSrv.txt
        Remove-Item $env:APPDATA\$ServerName.ps1
    }
    else{
        # already 32-bit. Just run the script.
        [System.Reflection.Assembly]::LoadFile("$PF\Citrix\ICA Client\WfIcaLib.dll") | Out-Null
        $ICA = New-Object WFICALib.ICAClientClass
        $ICA.Address = $ServerName
        $ICA.Application = ""
        $ICA.OutputMode = [WFICALib.OutputMode]::OutputModeWindowless
        $ICA.Launch = $true
        $ICA.TWIMode = $true
        $ICA.Connect()
        sleep -Seconds 7
        write-host $ICA.GetClientErrorMessage($ICA.GetLastError())
    }
}
else {
    Write-Host "ICA Client is not found. Please run this script on a computer with the ICA Client."
}
