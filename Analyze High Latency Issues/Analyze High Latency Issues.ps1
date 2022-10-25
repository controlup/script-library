#require -version 3.0

<#  
.SYNOPSIS
        The script executes the Tracert.exe command from the server\VDI to the client IP.The output will be the tracert command results, with the latency.
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
.NOTES
		10-7-2022 Ton de Vreede
		- Refactored
		- Better handling of time-out of tracert, set to maximum 5 minutes
#>

[CmdletBinding()]
Param
(
	[Parameter(Mandatory = $true, HelpMessage = 'The IP to trace the route to.')]
	[IPAddress]$ClientIP
)
$ErrorActionPreference = 'Stop'

if ($ClientIP.IPAddressToString -eq '0.0.0.0') {
	Write-Output -InputObject 'Session is disconnected, exiting.'
	Exit 1
}

# Start process for maximum of 5 minutes. Error out if it takes too long (over 5 minutes).
$prcTracert = Start-Process -FilePath 'C:\Windows\System32\TRACERT.EXE' -ArgumentList $ClientIP -PassThru -NoNewWindow
try {
	Wait-Process -Id $prcTracert.Id -Timeout 300
}
catch {
	Stop-Process -Id $prcTracert.Id
	Write-Output -InputObject "`nTRACERT.EXE execution exceeded 5 minute timeout."
	Exit 1
}
