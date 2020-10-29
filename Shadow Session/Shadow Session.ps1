<#
    .SYNOPSIS
        Shadow a user session

    .DESCRIPTION
        Shadow a remote user session

    .EXAMPLE
        . .\ShadowSession.ps1 -SessionId 1 -ComputerName W2019-001 -RemoteControlSession true -AllowUserConsent false
        Connects to a session and starts shadowing

    .PARAMETER  <SessionId <Int>>
        Session ID integer of the target session

    .PARAMETER  <ComputerName <string>>
        Remote Computer target user resides on

    .PARAMETER  <RemoteControlSession <boolean>>
        true  == grant view/mouse/keyboard control to the shadower
        false == only view the remote session

    .PARAMETER  <AllowUserConsent <boolean>>
        true  == user must consent to the shadow request
        false == bypass user consent

    .PARAMETER  <PromptForAlternativeCredentials <boolean>>
        true  == require the user to specify credentials to connect and shadow
        false == connect using your current credentials

    .PARAMETER  <Override <boolean>>
        true  == Attempt to enable the Shadow firewall rule (if the firewall is enabled) on the target machine and apply your shadow parameters to the GPO on the target machine
        false == respect the GPO on the target machine and do not touch the firewall

    .CONTEXT
        Console

    .NOTES
        In order to grant some of the more advanced functionality (bypass user consent, force remote control of the session) the initiator of this
        Script Action must have local admin rights of the target machine.  If the initiator does not have local admin rights on the target machine
        then the script will connect with whatever capabilities it can as dictated by policy.

    .MODIFICATION_HISTORY
        Created TTYE : 2020-10-08


    AUTHOR: Trentent Tye/Marcel Calef
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='Session ID of the target user')][ValidateNotNullOrEmpty()]                 [int]$SessionId,
    [Parameter(Mandatory=$true,HelpMessage='Name of the target machine user resides on')][ValidateNotNullOrEmpty()] [string]$ComputerName,
    [Parameter(Mandatory=$false,HelpMessage='Enable Remote Control of the target session?')]                        [string]$RemoteControlSession = "false",
    [Parameter(Mandatory=$false,HelpMessage='Allow the user to consent to the shadow request?')]                    [string]$AllowUserConsent = "false",
    [Parameter(Mandatory=$false,HelpMessage='Prompt to use other credentials for shadowing')]                       [string]$PromptForAlternativeCredentials = "false",
    [Parameter(Mandatory=$false,HelpMessage='Attempts to override GPO and enforce shadowing')]                      [string]$Override = "false"
)


#$ErrorActionPreference = "Stop"
#Set-StrictMode -Version Latest
$verbosePreference = "Continue"
#requires -version 4

#Convert string to boolean variable
[bool]$RemoteControlSession            = [bool]::Parse($RemoteControlSession)
[bool]$AllowUserConsent                = [bool]::Parse($AllowUserConsent)
[bool]$PromptForAlternativeCredentials = [bool]::Parse($PromptForAlternativeCredentials)
[bool]$Override                        = [bool]::Parse($Override)

Write-Verbose -Message "Parameters: `
  SessionId                       = $SessionId `
  ComputerName                    = $computerName `
  RemoteControlSession            = $RemoteControlSession `
  AllowUserConsent                = $AllowUserConsent `
  PromptForAlternativeCredentials = $PromptForAlternativeCredentials `
  Override                        = $Override"

function Check-RemoteFirewallSettings ([string]$ComputerName) {

    $FirewallStateResult = New-Object System.Collections.Generic.List[PSObject]
    $Profile                      = "Unknown"
    $FirewallStateEnabledProperty = "Unknown"
    $FirewallRuleEnabledState     = "Unknown"

    try {
        $CimSession = New-CimSession -ComputerName $computerName -ErrorAction Stop
    } catch {
        Write-Verbose -Message "Unable to connect to $computerName"
        $FirewallStateResult.Add(
            [PSCustomObject]@{
                Profile = $Profile
                FirewallStateEnabledProperty = $FirewallStateEnabledProperty
                FirewallRuleEnabledState = $FirewallRuleEnabledState
            }
        )
        return $FirewallStateResult
    }

    ## Get Active Network Connection Category (Domain, Private, Public)
    try {
        $ConnectedNetwork = Get-NetConnectionProfile -CimSession $CimSession
    } catch {
        Write-Verbose -Message "Failed to get the network connection profile for $computerName"
    }
    switch ($ConnectedNetwork.NetworkCategory) {
        "DomainAuthenticated" { $Profile = "Domain"  ; Break }
        "Private"             { $Profile = "Private" ; Break }
        "Public"              { $Profile = "Public"  ; Break }
        Default               { $Profile = "Unknown" ; Break }
    }

    Write-Verbose -Message "Detected Connected Network : $Profile"
    
    ## See if the firewall is enabled.  If not, we can skip checking
    try {
        $FirewallState = Get-NetFirewallProfile -Name $Profile -CimSession $CimSession
    } catch {
        Write-Verbose -Message "Failed to get the firewall state for $computerName"
    }
    Write-Verbose -Message "Firewall State for Network $Profile : $($FirewallState.Enabled)"

    switch ($FirewallState.Enabled) {
        "False"    { $FirewallStateEnabledProperty = "$False"   ; Break }
        "True"     { $FirewallStateEnabledProperty = "$True"    ; Break }
        Default    { $FirewallStateEnabledProperty = "Unknown"  ; Break }
    }
    if ($FirewallStateEnabledProperty -eq "$False") { Remove-CimSession -CimSession $CimSession }

    ## See if the RDP Shadow rule is enabled
    try {
        $FirewallRuleState = Get-NetFirewallRule -Name "RemoteDesktop-Shadow-In-TCP" -CimSession $CimSession
    } catch {
        Write-Verbose "Failed to get Firewall Rule State"
    }
    Write-Verbose "Firewall Rule State : $($FirewallRuleState.Enabled)"
    switch ($FirewallRuleState.Enabled) {
        "False"    { $FirewallRuleEnabledState = "$False"   ; Break }
        "True"     { $FirewallRuleEnabledState = "$True"    ; Break }
        Default    { $FirewallRuleEnabledState = "Unknown"  ; Break }
    }

    if ($FirewallRuleEnabledState -eq "$False") { Remove-CimSession -CimSession $CimSession }

    $FirewallStateResult.Add(
            [PSCustomObject]@{
                Profile = $Profile
                FirewallStateEnabledProperty = $FirewallStateEnabledProperty
                FirewallRuleEnabledState = $FirewallRuleEnabledState
            }
        )
    return $FirewallStateResult
}

function Check-ForShadowPolicies ([string]$ComputerName) {
    ## Check if policy is set on remote machine by GPO
    Write-Verbose "Check for Shadow Policies"
    try {
        $GPOConfigured = Invoke-Command -ComputerName $ComputerName -ScriptBlock { Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" }
    } catch {
        Write-Verbose "***Unable to connect to $ComputerName***"
        return "Fail"
    }

    
    if ([bool]($GPOConfigured.PSobject.Properties.name -match "Shadow")) {
        Write-Verbose -Message "GPO Shadow Policy Configured."
        return $GPOConfigured.Shadow
    } else {
        Write-Verbose -Message "No Shadow Policies Configured."
        return $false
    }
}

function Create-ShadowPolicy ([string]$ComputerName, [bool]$RemoteControlSession, [bool]$AllowUserConsent, [bool]$RemovePolicy) {
    if ($RemovePolicy) {
        Write-Verbose -Message "Remove Shadow Policy"
        try {
            Invoke-Command -ComputerName $ComputerName -ScriptBlock { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -Name "Shadow" -Force }
        } catch {
            Write-Error "Failed to remove policy"
        }
        break
    }

    $GPOValue = 0
    if ($RemoteControlSession -and $AllowUserConsent) {
        Write-Verbose -Message "Policy : Full Control with user's permission" 
        $GPOValue = 1
    }
    if ($RemoteControlSession -and -not $AllowUserConsent) {
        Write-Verbose -Message "Policy : Full Control without user's permission" 
        $GPOValue = 2
    }

    if (-not $RemoteControlSession -and $AllowUserConsent) {
        Write-Verbose -Message "Policy : View Session with user's permission" 
        $GPOValue = 3
    }

    if (-not $RemoteControlSession -and -not $AllowUserConsent) {
        Write-Verbose -Message "Policy : View Session without user's permission" 
        $GPOValue = 4
    }

    Write-Verbose "GPO Value = $GPOValue"

    if ($GPOValue -eq 0) {
        Write-Output "Unable to detect what the shadow policy may be. This script may fail."
    }

    $DefineGPO = Invoke-Command -ComputerName $ComputerName -ScriptBlock { Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -Name "Shadow" -Type DWord -Value $args[0] -Force -PassThru } -ArgumentList $GPOValue
    Write-Debug -Message "$($DefineGPO | Format-List | Out-String)"
    if ([bool]($DefineGPO.PSobject.Properties.name -match "Shadow")) {
        if ($DefineGPO.Shadow -ne $GPOValue) {
            Write-Error "Failed to set shadow policies."
            return $false
        } else {
            return $true
        }
    } else {
        Write-Error "Failed to set shadow policies."
        return $false
    }
}

function Shadow-Session ([int]$sessionId, [string]$ComputerName, $RemoteControlSession, $AllowUserConsent, $PromptForAlternativeCredentials) {
    $StringBuilder = New-Object -TypeName "System.Text.StringBuilder"
    [void]$StringBuilder.Append("/shadow:$sessionId ")
    [void]$StringBuilder.Append("/v:$ComputerName ")
    if ($RemoteControlSession) {
        [void]$StringBuilder.Append("/control ")
    }
    if (-not($AllowUserConsent)) {
        [void]$StringBuilder.Append("/noConsentPrompt ")
    }
    if ($PromptForAlternativeCredentials) {
        [void]$StringBuilder.Append("/prompt ")
    }

    Write-Verbose "Executing : mstsc.exe $($StringBuilder.ToString())"
    Start-Process -FilePath mstsc.exe -ArgumentList @("$($StringBuilder.ToString())")
}

Write-Verbose "Checking $computerName to see if the Windows firewall is enabled"
$RemoteFirewallSettings = Check-RemoteFirewallSettings -ComputerName $ComputerName

Write-Verbose "Are we overriding the target machine's settings?"
if ($Override) {
    Write-Verbose "Yes"
    ##try to enable the firewall rules and overwrite the GPO settings.
    if ($RemoteFirewallSettings.FirewallStateEnabledProperty -eq $true -and $RemoteFirewallSettings.FirewallRuleEnabledState -eq $false) {
        Write-Output "Override : Attempting to enable the Shadow firewall rule"
        $CimSession = New-CimSession -ComputerName $computerName
        Set-NetFirewallRule -Name "RemoteDesktop-Shadow-In-TCP" -Enabled True -CimSession $CimSession
        $RemoteFirewallSettings.FirewallRuleEnabledState = "True"
        Remove-CimSession -CimSession $CimSession
    }
    ##try to overwrite GPO settings with parameters
    Create-ShadowPolicy -ComputerName $ComputerName -RemoteControlSession $RemoteControlSession -AllowUserConsent $AllowUserConsent
} else {
    Write-Verbose "No"
}

## Is the Firewall Enabled and the RDP TCP Shadow policy disabled?
if ($RemoteFirewallSettings.FirewallStateEnabledProperty -eq $true -and $RemoteFirewallSettings.FirewallRuleEnabledState -eq $false) {
    Write-Error "Windows Firewall is enabled but Rule: `"Remote Desktop - Shadow (TCP-In)`" was found disabled. `n`nPlease enable firewall rule `"Remote Desktop - Shadow (TCP-In)`" in order to enable shadowing."
    break
}

Write-Verbose "Checking $computerName for Shadow Policies"
$ShadowPolicy = Check-ForShadowPolicies -ComputerName $ComputerName

if ($ShadowPolicy -ne "Fail") {
    if ($ShadowPolicy -eq $false) {
        Write-Verbose "Shadow Policies not configured -- Attempting to use defined parameters."
        Create-ShadowPolicy -ComputerName $computerName -RemoteControlSession $RemoteControlSession -AllowUserConsent $AllowUserConsent
    } else {
        Write-Output "Group Policy Shadow Policies found. Overriding less restrictive parameters to match more restrictive policy"  ## if AllowUserConsent is True and $RemoteControlSession is False then we'll just enforce those options
        switch ($ShadowPolicy) {
            0       { Write-Output "Detected Policy : No remote control allowed" }
            1       { Write-Output "Detected Policy : Full Control with user's permission"    ; $AllowUserConsent=$true   ; break }
            2       { Write-Output "Detected Policy : Full Control without user's permission" ; break }
            3       { Write-Output "Detected Policy : View Session with user's permission"    ; $RemoteControlSession=$false ; $AllowUserConsent=$true   ; break }
            4       { Write-Output "Detected Policy : View Session without user's permission" ; $RemoteControlSession=$false ; break }
            Default { Write-Output "Detected Policy : Unknown Policy Value : $($GPOConfigured.Shadow) " }
        }
    }
} else {
    Write-Verbose "Unable to determine shadow policies, will attempt to connect with the most restrictive parameters."
    $RemoteControlSession=$false
    $AllowUserConsent=$true
}

Shadow-Session -sessionId $SessionID -ComputerName $ComputerName -RemoteControlSession $RemoteControlSession -AllowUserConsent $AllowUserConsent -PromptForAlternativeCredentials $PromptForAlternativeCredentials

Sleep 5  #sleep to allow mstsc.exe to connect to remote computer before switching policies back

if ($ShadowPolicy -eq $false) {
    Create-ShadowPolicy -ComputerName $ComputerName -RemovePolicy $true
}
