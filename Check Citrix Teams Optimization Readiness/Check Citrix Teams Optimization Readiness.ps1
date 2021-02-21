#requires -Version 3.0
<#  
.SYNOPSIS     Check Citrix Teams Optimization Readiness [BETA]
.DESCRIPTION  [BETA] Citrix HDX Optimization can improve the user experience for Teams video/audio use.
              The optimization will also significantly reduce the resource consumption on the VDA.
              

.SOURCES      https://techcommunity.microsoft.com/t5/microsoft-teams/still-connecting-to-remote-devices/m-p/1906370
              https://docs.citrix.com/en-us/citrix-virtual-apps-desktops/multimedia/opt-ms-teams.html
              https://docs.citrix.com/en-us/citrix-virtual-apps-desktops/multimedia/opt-ms-teams.html#known-limitations
              https://docs.microsoft.com/en-us/MicrosoftTeams/msi-deployment#clean-up-and-redeployment-procedure
              https://docs.microsoft.com/en-us/MicrosoftTeams/teams-for-vdi

.EXAMPLE:     \\util01\share\research\MSFT_Teams_Citrix_optimization.ps1 -SessionID '1' -vdaVer '1912.0.0.24265' -protocol 'HDX' -ctxRx '20.12.1.42' -userChanges Discard
.CONTEXT      Session
.TAGS         $HDX, $Citrix, $Teams
.HISTORY      Marcel Calef     - 2021-01-06 - BETA Release 
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='SessionID')]             [string]$SessionID,
    [Parameter(Mandatory=$false, HelpMessage='Citrix VDA ver.')]     [string]$vdaVer,
    [Parameter(Mandatory=$false, HelpMessage='protocol')]            [string]$protocol,
    [Parameter(Mandatory=$false, HelpMessage='Citrix Client ver.')]  [string]$ctxRx,
    [Parameter(Mandatory=$false, HelpMessage='userChanges')]         [string]$userChanges
      )

Set-StrictMode -Version Latest
[string]$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'      # Remove the comment in the begining to enable Verbose outut

Write-Verbose "Input:"
#Write-Verbose "   SessionID :  $SessionID"  ## Not showing to allow better results grouping
Write-Verbose "      vdaVer :  $vdaVer"
Write-Verbose "    protocol :  $protocol"
Write-Verbose "       ctxRx :  $ctxRx"

# If not a Citrix session - no need to continue
if($protocol -ne "HDX" -and $protocol -ne "Console"){Write-Output "INFO :   Not a Citrix session. Exiting"; exit }

###############################################################################################################
# Citrix components versions and capability validations

# Check if the VDA supports it (1906 or newer)
$vdaVerTest   = ($vdaVer -match '^19(06|09|12)\..+|^2.{3}\..+')

# Check if the VDA has the webSockets service running
Try{$CtxTeamsSvc = ((Get-Service CtxTeamsSvc).status -match "Running")}
	catch {$CtxTeamsSvc = $false} # return $false if service not running

# Check the WebBrowserRedirection Policy  1=enabled (implicitly enabled if not found in VDA > 7.??)
Try {$redirPol = (get-itemproperty -path HKLM:\SOFTWARE\Policies\Citrix\$SessionID\User\MultimediaPolicies  -name "TeamsRedirection").TeamsRedirection}
   catch {$redirPol = "notFound"} # return empty if not found

# Check the Citrix Reciever (CWA) version
$ctxRxVerTest = ($ctxRx -match '^20\..+')

# Check if $ctxRx was empty
if ($ctxRx -match 'Discard|On Local'){$ctxRx = 'not Received';Write-Verbose "       ctxRx :  $ctxRx"}


# check HKEY_CURRENT_USER\SOFTWARE\Citrix\HDXMediaStream\\MSTeamsRedirSupport
## Teams will look for this to decide to enable or not redirection
Try {$userMSTeamsSupp = (Get-ItemProperty -Path hkcu:software\Citrix\HDXMediaStream -Name "MSTeamsRedirSupport").MSTeamsRedirSupport}
     Catch {$userMSTeamsSupp = "notFound" }

# Check if the CVAD Persist User Changes was provided, if not set as unknown
if (!($userChanges -match '.+')) {$userChanges = 'Not Provided'}
Write-Verbose " userChanges :  $userChanges"

###############################################################################################################
# Teams install mode, version and session log entries

# Check Teams executable path
Try {$machineInstall = ((Test-Path "${env:ProgramFiles(x86)}\Microsoft\Teams\current\Teams.exe") -or
                        (Test-Path "$env:ProgramFiles\Microsoft\Teams\current\Teams.exe"))
    }
    Catch {$machineInstall = "False"}

Try {$userInstall = (test-path "$env:LOCALAPPDATA\Microsoft\Teams\") }
    Catch {$userInstall = "False"}

# Check if MSI install is allowed
Try{$preventMSIinstall = (Get-Itemproperty HKCU:Software\Microsoft\Office\Teams).PreventInstallationFromMsi}
     Catch{$preventMSIinstall = 'not found'}

###############################################################################################################
# Find the teams.exe root process and get it's PID
    Try {$teams = (Get-Process teams | Where-Object {$_.mainWindowTitle}) # get the teams process with the visible window.
         $teamsPID = $teams.id
         $teamsVer = $teams.FileVersion}
    Catch {$teamsPID = 'n/a'; $teamsVer = 'not running' }

#Write-Verbose "    teamsPID :  $teamsPID"    ## Not showing to allow better results grouping
Write-Verbose "    teamsVer :  $teamsVer"

# Check if Teams.exe received the Citrix HDX Redirection RegKey (would be 1, else 0 or not found)
if ($teamsVer -ne 'not running'){
    Try {$hdxRedirLogEntry = ((Select-String -path "$env:appdata\Microsoft\Teams\logs.txt" -Pattern "<$teamsPID> -- info -- vdiUtility: citrixMsTeamsRedir")[-1]).ToString()
     $hdxRedir = $hdxRedirLogEntry.Substring($hdxRedirLogEntry.Length -2, 2)
     }
    Catch {$hdxRedir = "notFound"}
    }
    else {$hdxRedir = "not running"}


###############################################################################################################
# Report findings - some tests return $true , other return a value

Write-Output "====================== Citrix Readiness ======================"
if ($vdaVerTest)      {Write-Output "PASS:       VDA Version $vdaVer supports HDX Teams Optimization" }
         else 	      {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       VDA Version $vdaVer does not support HDX Teams Optimization (or not found in the output)"
                       Write-Output "      TRY:  Upgrade VDA to 1903 or newer. see https://docs.citrix.com/en-us/tech-zone/learn/poc-guides/microsoft-teams-optimizations.html"
                      }
					  
if ($ctxRxVerTest)    {Write-Output "PASS:       Citrix Workspace App Version $ctxRx supports HDX Teams Optimization" }
		 else 	      {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       Receiver/CWA Version $ctxRx is old and does not support HDX Teams Optimization "
				       Write-Output "      TRY:  Upgrade the Client device to latest version of Citrix Workspace App"
					   if ($vdaVerTest -eq $false) {exit}
				      }

if ($redirPol -eq 1)  {Write-Output "PASS:       HDX Teams Optimization policy explicitly ENABLED from Citrix Studio"}

if ($redirPol -eq 0)  {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       HDX Teams Optimization DISABLED explicitly via policy"
				       Write-Output "      TRY:  Review Citrix Policies in Citrix Studio"
					   exit
				      }
				    
if ($CtxTeamsSvc)    {Write-Output "PASS:       HDX Teams Optimization (CtxTeamsSvc) found running in the VDA" }
		 else 	      {Write-Output "`n======!!======!!"
					   if ($redirPol -eq "notFound") {
					   Write-Output "INFO:       HDX Teams Optimization policy not set explicitly via policy"
				       Write-Output "      TRY:  Review Citrix Policies in Citrix Studio & Enable HDX Browser content redirection"
					   Write-Output "            Check if the HDX Teams Optimization service is running on the VDA"
													 }
					   Write-Output "FAIL:       HDX Teams Optimization service (CtxTeamsSvc) not running "
					   Write-Output "      TRY:  Check if the CtxTeamsSvc service is running on the VDA"
					  
					  }

# HKEY_CURRENT_USER\SOFTWARE\Citrix\HDXMediaStream   MSTeamsRedirSupport  will be 1 if VDA and CWA support it.
Switch ($userMSTeamsSupp)
    { '1'            { Write-Output "PASS:       Citrix reports this HDX session supports Teams Optimization (MSTeamsRedirSupport is 1)"
                       if (!$CtxTeamsSvc) {Write-Output "      INFO:  See warning for CtxTeamsSvc"}
                     }
      '0'            { Write-Output "WARN:       Citrix HDX redirection for Teams not supported on this session (MSTeamsRedirSupport is not 1)" 
 				       Write-Output "      TRY:  Review Citrix VDA and Workspace App versions or DIsconnect and reconnect the session"
                       }
      'notFound'     { Write-Output "WARN:       Citrix HDX redirection for Teams not supported on this session (MSTeamsRedirSupport not found in HKCU)"}
    }



Write-Output "====================== Teams Readiness ======================"

if ($machineInstall)  {Write-Output "PASS:       Teams found in the Program Files directory"}
   else    {Write-Output "`n======!!======!!"
            Write-Output "WARN :      Teams not found in the Program Files directory. "
            if ((test-path "$env:LOCALAPPDATA\Microsoft\Teams\") -and !($userChanges -match 'Local')) 
                { Write-Output "WARN :      Teams found in the User's Local AppData folder and VDA is not persistent."
                  Write-Output "      SEE:  https://docs.microsoft.com/en-us/MicrosoftTeams/teams-for-vdi#non-persistent-setup"
                }
            if ($preventMSIinstall -match '1')         
                { Write-Output "WARN :      PreventInstallationFromMSI variable found at HKCU:Software\Microsoft\Office\Teams"
                  Write-Output "      SEE   https://docs.microsoft.com/en-us/MicrosoftTeams/msi-deployment#clean-up-and-redeployment-procedure"
                }
            }
    
                      		   
Switch -Wildcard ($hdxRedir)
    { '*1*'         { Write-Output "PASS:       Teams reports Citrix HDX Optimized - in the GUI: User-> About->Version"}
      '*0*'         { Write-Output "`n======!!======!!"
                      Write-Output "WARN:       Citrix HDX NOT Optimized - in the GUI: User-> About->Version"}
      'not running' { Write-Output "`n======!!======!!"
                      Write-Output "INFO:       Teams was not detected running in this session"}
      default       { Write-Output "`n======!!======!!"
                      Write-Output "WARN:       Teams did not detect Citrix HDX optimization"}
    }
