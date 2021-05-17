<#  
.SYNOPSIS
        The script sends Tracert.exe command from the server\VDI to the client IP.The output will be the tracert command results, with the latency.
.DESCRIPTION
        The script runs on the the target VDI\XenApp computer. It will initiate a trace route command from the VDI\XenApp 
        machine to the client device which is the ClientIP.
.PARAMETER ServerName
        This script gets only one parameter which is the client IP
.EXAMPLE
        AnalyzeLatency.ps1 10.10.10.1
.OUTPUTS
        The tracert command output, with all the latency between each and every hop.
.LINK
        See http://www.ControlUp.com
#>

$ClientIP = $args[0]
if($ClientIP -eq "0.0.0.0"){
    Write-Host "Session is disconnected"
    exit 1
}
tracert $ClientIP
