#requires -Version 3.0
<#  
.SYNOPSIS     Check Horizon Teams Optimization Readiness [BETA]
.DESCRIPTION  [BETA] VMware Horizon is capable of video/audio Optimization for Teams video/audio use.
              The optimization will also significantly reduce the resource consumption on the DataCenter.
              

.SOURCES      https://techzone.vmware.com/resource/microsoft-teams-optimization-vmware-horizon
              https://docs.vmware.com/en/VMware-Horizon/2006/horizon-remote-desktop-features/GUID-F68FA7BB-B08F-4EFF-9BB1-1F9FC71F8214.html
              https://techcommunity.microsoft.com/t5/microsoft-teams/teams-and-vmware-horizon-vdi-best-practices/m-p/1759816

.EXAMPLE:     \\util01\share\research\MSFT_Teams_VMware_Horizon_optimization.ps1 -SessionID 16 -hznVer 7.12.0 -protocol Blast -cltVer 8.1.0
.CONTEXT      Session
.TAGS         $VMware, $Horizon, $Blast, $Teams
.HISTORY      Marcel Calef     - 2021-03-24 - BETA Release
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='SessionID')]             [string]$SessionID,
    [Parameter(Mandatory=$false, HelpMessage='protocol')]            [string]$protocol,
    [Parameter(Mandatory=$false, HelpMessage='Horizon Agent ver.')]  [string]$hznVer,
    [Parameter(Mandatory=$false, HelpMessage='Horizon Client ver.')] [string]$cltVer,
    [Parameter(Mandatory=$false, HelpMessage='userChanges')]         [string]$userChanges
      )

Set-StrictMode -Version Latest
[string]$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'      # Remove the comment in the begining to enable Verbose outut

Write-Verbose "Input:"
#Write-Verbose "   SessionID :  $SessionID"  ## Not showing to allow better results grouping
Write-Verbose "      hznVer :  $hznVer"
Write-Verbose "    protocol :  $protocol"
Write-Verbose "      cltVer :  $cltVer"

# If not a Horizon session - no need to continue
if($protocol -ne "Blast" -and $protocol -ne "Console"){Write-Output "INFO :   Not a VMware Horizon session. Exiting"; exit }

###############################################################################################################
# Horizon components versions and capability validations

# Check if the Horizon Agent supports it (7.13, 8 or newer)
$hznVerTest   = ($hznVer -match '^7\.13\..+|^8\..+')

# Check if the VDA has the webSockets service running
Try{$CtxTeamsSvc = ((Get-Service CtxTeamsSvc).status -match "Running")}
	catch {$CtxTeamsSvc = $false} # return $false if service not running

# Check Horizon's Microsoft Teams Optimization Feature Policy - must be explicitly enabled
Try {$optimPol = ($gpResult | select-string 'Enable Media Optimization for Microsoft Teams' -Context(1,2))}
   catch {$optimPol = "notFound"} # return empty if not found

# Check Horizon's per user Software Acoustic Echo Cancellation Policy - implicitly enabled
Try {$echoPol = ($gpResult | select-string 'Enable software acoustic echo cancellation for Media Optimization for Microsoft Teams' -Context(1,2))}
   catch {$echoPol = "notFound"} # return empty if not found
# It is possible to configure from the client. No way to check here
#Try {$ClientEchoPol = (get-itemproperty -path 'HKCU:\SOFTWARE\VMware, Inc.\VMware Html5mmr\WebrtcRedir'  -name "enableAEC").enableAEC}
#   catch {$ClientEchoPol = "notFound"} # return empty if not found

# Check the Horizon Client version
$cltVerTest = ($cltVer -match '^5\.5.+|^8\..+')

# Check if $cltVer was empty
if ($cltVer -match 'Discard|On Local'){$cltVer = 'not Received';Write-Verbose "       cltVer :  $cltVer"}


# check Horizon's handover registry key for MSTeamsRedirSupport   ###### Not implemented yet
## Teams will look for this to decide to enable or not redirection
Try {$userMSTeamsSupp = (Get-ItemProperty -Path hkcu:software\VMware -Name "MSTeamsRedirSupport").MSTeamsRedirSupport}
     Catch {$userMSTeamsSupp = "notFound" }

# Check if the Hoziron MAchine Type was provided, if not set as unknown    ###### Not implemented yet
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

# Check if Teams.exe received the Horizon Redirection RegKey (would be 1, else 0 or not found)    ###### Not implemented yet
if ($teamsVer -ne 'not running'){
    Try {$hdxRedirLogEntry = ((Select-String -path "$env:appdata\Microsoft\Teams\logs.txt" -Pattern "<$teamsPID> -- info -- vdiUtility: citrixMsTeamsRedir")[-1]).ToString()
     $hznRedir = $hdxRedirLogEntry.Substring($hdxRedirLogEntry.Length -2, 2)
     }
    Catch {$hznRedir = "notFound"}
    }
    else {$hznRedir = "not running"}


###############################################################################################################
# Report findings - some tests return $true , other return a value

Write-Output "====================== Horizon Readiness ======================"
if ($hznVerTest)      {Write-Output "PASS:       Horizon Agent Version $hznVer supports Teams Optimization" }
         else 	      {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       Horizon Agent Version $hznVer does not support Teams Optimization (or not found in the output)"
                       Write-Output "      TRY:  Upgrade the Horizon Agent to 7.13, 8 or newer. see https://techzone.vmware.com/resource/microsoft-teams-optimization-vmware-horizon"
                      }
					  
if ($cltVerTest)    {Write-Output "PASS:       VMware Horizon Client Version $cltVer supports Teams Optimization" }
		 else 	      {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       VMware Horizon Client Version $cltVer is old and does not support Teams Optimization "
				       Write-Output "      TRY:  Upgrade the Client device to latest version of VMware Horizon Client"
					   if ($hznVerTest -eq $false) {exit}
				      }

if ($optimPol -ne "notFound")  {Write-Output "PASS:       Horizon's Microsoft Teams Optimization Feature Policy FOUND - pending validate enabled"}

if ($optimPol -eq "notFound")  {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       Horizon's Microsoft Teams Optimization Feature Policy NOT FOUND"
				       Write-Output "      TRY:  See 'WebRTC Redirection' in https://techzone.vmware.com/resource/microsoft-teams-optimization-vmware-horizon"
					   exit
				      }

# 
				    
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
    
                      		   
Switch -Wildcard ($hznRedir)
    { '*1*'         { Write-Output "PASS:       Teams reports Horizon Optimized - in the GUI: User-> About->Version"}
      '*0*'         { Write-Output "`n======!!======!!"
                      Write-Output "WARN:       Horizon NOT Optimized - in the GUI: User-> About->Version"}
      'not running' { Write-Output "`n======!!======!!"
                      Write-Output "INFO:       Teams was not detected running in this session"}
      default       { Write-Output "`n======!!======!!"
                      Write-Output "WARN:       Teams did not detect Horizon optimization"}
    }
