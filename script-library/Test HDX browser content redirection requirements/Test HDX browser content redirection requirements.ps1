<#  .SYNOPSIS       HDX Browser content redirection capability test
    .DESCRIPTION    TBA
    .EXAMPLE        HDXredir_test.ps1 
    .CONTEXT        Process
    .TAGS           $Name="iexplore.exe,chrome.exe",$CPU>"5"
    .REFERENCES
                    https://docs.citrix.com/en-us/xenapp-and-xendesktop/7-15-ltsr/multimedia/browser-content-redirection.html
                    https://docs.citrix.com/en-us/citrix-virtual-apps-desktops/multimedia/browser-content-redirection.html
    .MODIFICATION_HISTORY
        Marcel Calef     - 2020-03-10 - Initial release
		Marcel Calef     - 2020-10-07 - Modified to Session and separated test from report
 #>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true, HelpMessage='Session CPU consumption')]        [string]$sessionCPU,
    [Parameter(Mandatory=$true, HelpMessage='SessionID')]                      [int]$sessionID,
    [Parameter(Mandatory=$false, HelpMessage='Citrix VDA Version')]            [string]$vdaVer,
    [Parameter(Mandatory=$false, HelpMessage='Session protocol')]              [string]$protocol,
    [Parameter(Mandatory=$false, HelpMessage='Citrix Workspace app version')]  [string]$ctxRx,
    [Parameter(Mandatory=$false, HelpMessage='session HDX Bandwidth Avg')]     [string]$bwAvg,
    [Parameter(Mandatory=$false,HelpMessage='session Active Application ')]    [string]$activeApp,
    [Parameter(Mandatory=$false,HelpMessage='session Active URL')]             [string]$activeURL
        )

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$VerbosePreference = "continue"  # Uncomment this line to enable verbose debug output

Write-Verbose "inputs:"
Write-Verbose "         sessionCPU :  $sessionCPU"
Write-Verbose "       sessionID :  $sessionID"
Write-Verbose "          vdaVer :  $vdaVer"
Write-Verbose "        protocol :  $protocol"
Write-Verbose "           ctxRx :  $ctxRx"
Write-Verbose "           bwAvg :  $bwAvg"
Write-Verbose "       activeApp :  $activeApp"
Write-Verbose "       activeURL :  $activeURL"

# If not a Citrix session - no need to continue
if($protocol -ne "HDX" -and $protocol -ne "Console"){Write-Output "Not a Citrix session. Exiting"; exit }

###############################################################################################################
# Citrix components versions and capability validations

# Check if the VDA supports it (7.15 CU3 and 1808 or newer)
$vdaVerTest   = ($vdaVer -match '^7\.15\.[3-8]000\.[0-9]{1,6}$|^7\.1[6-8]\..+|^1808\..+|^1811\..+|^19.{2}\..+|^2.{3}\..+')

# Check if the VDA has the webSockets service running
Try{$WebSocketSvc = (Get-Process -Name WebSocketService)}
	catch {$WebSocketSvc = $false} # return $false if process not found

# Check the WebBrowserRedirection Policy  1=enabled (implicitly enabled if not found in VDA > 7.??)
Try {$redirPol = (get-itemproperty -path HKLM:\SOFTWARE\Policies\Citrix\MultimediaPolicies  -name "WebBrowserRedirection").WebBrowserRedirection}
   catch {$redirPol = "notFound"} # return empty if not found

# Read the WebBrowserRedirectionACL for this user session (from the registry (youtube is implicitly allowed)
Try {$redirACL = (get-itemproperty -path HKLM:\SOFTWARE\Policies\Citrix\$sessionID\User\MultimediaPolicies).WebBrowserRedirectionAcl}
   catch {$redirACL = "https://www.youtube.com/*"}  # return YouTube as it is implicitly allowed

# Check the Citrix Reciever (CWA) version
$ctxRxVerTest = ($ctxRx -match '^14\.10\..+$|^13\.9\..+|^18\.08\..+|^19\..+|^20\..+')

# Check if the session has WebSocketAgent
Try{$WebSocketAgent = (Get-Process -Name WebSocketAgent | Where-Object {$_.SessionID -eq $sessionID})}
	catch {$WebSocketAgent = $false} # return $false if process not found

Write-Verbose "        redirPol :  $redirPol"
Write-Verbose "        redirACL :  $redirACL"
Write-Verbose "=============================="

# Need to do a pattern search of the Active URL in the contents of $redirACL
$ActiveInACL = ($redirACL -like "*$activeURL*")

###############################################################################################################

# Report findings - some tests return $true , other return a value

if ($vdaVerTest)      {Write-Output "PASS:       VDA Version $vdaVer supports HDX Browser Content Redirection" }
         else 	      {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       VDA Version $vdaVer does not support HDX Browser Content Redirection (or not found in the output)"
                       Write-Output "      TRY:  Upgrade VDA to 7.15 CU3, 1808 or newer. see CTX230052"
                      }
					  
if ($ctxRxVerTest)    {Write-Output "PASS:       Receiver/CWA Version $ctxRx supports HDX Browser Content Redirection" }
		 else 	      {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       Receiver/CWA Version $ctxRx does not support HDX Browser Content Redirection "
				       Write-Output "      TRY:  Upgrade the Client device Receiver/CWA to a more recent version"
					   if ($vdaVerTest -eq $false) {exit}
				      }

if ($redirPol -eq 1)  {Write-Output "PASS:       HDX Browser redirection policy explicitly ENABLED from Citrix Studio"}

if ($redirPol -eq 0)  {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       HDX Browser redirection policy DISABLED explicitly via policy"
				       Write-Output "      TRY:  Review Citrix Policies in Citrix Studio & Enable HDX Browser content redirection"
					   exit
				      }
				    
if ($WebSocketSvc)    {Write-Output "PASS:       HDX redirection service (WebSocketService.exe) found running in the VDA" }
		 else 	      {Write-Output "`n======!!======!!"
					   if ($redirPol -eq "notFound") {
					   Write-Output "INFO:       HDX Browser redirection policy not set explicitly via policy"
				       Write-Output "      TRY:  Review Citrix Policies in Citrix Studio & Enable HDX Browser content redirection"
					   Write-Output "            Check if the CtxHdxWebSocketService service is running on the VDA"
													 }
					   Write-Output "FAIL:       HDX Browser redirection service (WebSocketService.exe) not found "
					   Write-Output "      TRY:  Check if the CtxHdxWebSocketService service is running on the VDA"
					   exit
					  }

if ($ActiveInACL)     {Write-Output "PASS:       Active URL: $activeURL is on the redirection ACL"
                      }
                      		    
if ($WebSocketAgent)  {
				       Write-Output "PASS:       HDX redirection Agent (WebSocketAgent.exe) found in the session"
					   
					  }
		 else 	      {Write-Output "INFO:       HDX redirection Agent (WebSocketAgent.exe) not found in the session "
				       Write-Output "INFO:       WebSocketAgent.exe is only present if a browser is engaging the Redirection"
				       Write-Output "INFO:       For Chromium browsers (Chrome and new Edge):"
				       Write-Output "      TRY:  Add the Browser redirection Extension from the Chrome store"
				       Write-Output "      TRY:  https://chrome.google.com/webstore/detail/browser-redirection-exten/hdppkjifljbdpckfajcmlblbchhledln/related"
					   ## add here a test if the Active URL is found in $redirACL
				       Write-Output "RECOMMEND:  Check all previous FAILURES and review the ACL in the Citrix Studio"
				      }		



