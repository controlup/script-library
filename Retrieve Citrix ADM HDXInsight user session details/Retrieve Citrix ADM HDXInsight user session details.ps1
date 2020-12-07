<#
.SYNOPSIS
  Connect to Citrix ADM and retrieve HDX Insight information.
.DESCRIPTION
  Connect to Citrix ADM and retrieve HDX Insight information, using REST API and JSON.
.NOTES
  Version:        0.1
  Author:         Esther Barthel, MSc
  Creation Date:  2020-11-12
  Updated:        2020-11-15
                  Standardized the function, based on the ControlUp Standards (v0.2)
  Purpose:        Automating Citrix ADM HDX Insight information with REST APIs

  Copyright (c) cognition IT. All rights reserved.
#>
[CmdletBinding()]
Param
(
    [Parameter(
        Position=0, 
        Mandatory=$true, 
        HelpMessage='Enter the Citrix ADM management IP address'
    )]
    [ValidateScript({$_ -match [IPAddress]$_ })]
    [string] $ADMIP,

    [Parameter(
        Position=1, 
        Mandatory=$true, 
        HelpMessage='Enter the UserName'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $HDXUserName
)    

# -------------
# | Functions |
# -------------

function Invoke-ADMLogin {
    <#
    .SYNOPSIS
        Login to the Citrix ADM and return session information.
    .DESCRIPTION
        Login to the Citrix ADM and return session information, using the Invoke-RestMethod cmdlet for the REST API calls. 
    .EXAMPLE
        Invoke-ADMLogin -ADMIP string
    .EXAMPLE
        Invoke-ADMLogin -ADMIP string -ADMCredentials $PSCredentialsObject
    .EXAMPLE
        Invoke-ADMLogin -ADMIP string -Verbose
    .LINK
        https://developer-docs.citrix.com/projects/citrix-adm-nitro-api-reference/en/12.1/configuration/system/login/
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-11-12
        Purpose:        Script created for Citrix ADM Troubleshooting &amp; Management
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Citrix ADM management IP address'
        )]
        [ValidateScript({$_ -match [IPAddress]$_ })]
        [string] $ADMIP,
      
        [Parameter(
            Position=1, 
            Mandatory=$false, 
            HelpMessage='Enter a PSCredential object, containing username and password for the Citrix ADM'
        )]
        [System.Management.Automation.CredentialAttribute()] $ADMCredentials
    )    

    # Check if a PSCredentials object was provided, if not ask for credentials
    If (!($ADMCredentials))
    {
        # Get ADM Credentials
        [System.Management.Automation.PSCredential]$ADMCredentials = $null
        $ADMCredentials = Get-Credential -Message "Enter your credentials for Citrix ADM $NSIP"
    }

    # Extract username and password from PSCredentials for use with NITRO
    $ADMUserName = $ADMCredentials.UserName
    $ADMPassword = $ADMCredentials.GetNetworkCredential().Password

    #region Login to ADM with NITRO
        #Force PowerShell to bypass validation for (self-signed) certificates and SSL connections
        # source: https://blogs.technet.microsoft.com/bshukla/2010/04/12/ignoring-ssl-trust-in-powershell-system-net-webclient/ 
        Write-Verbose "* ADM-Login: Forcing PowerShell to trust all certificates (incl. self-signed netScaler certificate)"
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

        #JSON payload
        $Login = ConvertTo-Json @{
            "login" = @{
                "username"=$ADMUserName;
                "password"=$ADMPassword;
                "session_timeout"=300
            }
        } -Depth 5

        try
        {
            # Login to the ADM and create a session (stored in $NSSession)
            $invokeRestMethodParams = @{
                Uri             = "http://$ADMIP/nitro/v1/config/login"
                Body            = "object="+$Login
                Method          = "Post"
                SessionVariable = "ADMSession"
                ContentType     = "application/json"
            }
            $loginResponse = Invoke-RestMethod @invokeRestMethodParams
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct management IP-address. " + $_.Exception.Message)
            Break
        }
        catch
        {
            # Debug: 
            Write-Debug $_.Exception | Format-List -Force
            # Error:
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Break
        }
        # Check for REST API errors
        If ($loginResponse.errorcode -eq 0)
        {
            Write-Verbose "ADM-login: login successful"
        }
    #endregion Login to ADM with NITRO

    # return session information
    return $ADMSession
}

function Invoke-ADMLogout {
    <#
    .SYNOPSIS
        Logout the given Citrix ADM session.
    .DESCRIPTION
        Logout the given Citrix ADM session, using the Invoke-RestMethod cmdlet for the REST API calls. 
    .EXAMPLE
        Invoke-ADMLogout -ADMIP string -ADMSession SessionVariable
    .EXAMPLE
        Invoke-ADMLogout -ADMIP string -ADMSession SessionVariable -Verbose
    .LINK
        https://developer-docs.citrix.com/projects/citrix-adm-nitro-api-reference/en/12.1/configuration/system/login/#delete
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-11-12
        Purpose:        Script created for Citrix ADM Troubleshooting &amp; Management
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Citrix ADM management IP address'
        )]
        [ValidateScript({$_ -match [IPAddress]$_ })]
        [string] $ADMIP,
      
        [Parameter(
            Position=1, 
            Mandatory=$true, 
            HelpMessage='Enter variable, containing the Session information for the Citrix ADM'
        )]
        [System.Management.Automation.PSObject] $ADMSession
    )    

    #region Logout of ADM with NITRO
        # retrieve the sessionID from the session variable
        $cookieContainer = $ADMSession.Cookies
        [hashtable] $cookieDetails = $CookieContainer.GetType().InvokeMember("m_domainTable",
                [System.Reflection.BindingFlags]::NonPublic -bor
                [System.Reflection.BindingFlags]::GetField -bor
                [System.Reflection.BindingFlags]::Instance,
                $null,
                $cookieContainer,
                @()
        )
        $sessionID = $cookieDetails.Values.Values.Value

        #JSON payload
        $Logout = ConvertTo-Json @{
            "logout" = @{
                "sessionid"=$sessionID
            }
        } -Depth 5
        #demo output: Write-Host "JSON: $Logout" -ForegroundColor Green

        try
        {
            # Connect to the ADM and logoff a session (stored in $NSSession)
            $invokeRestMethodParams = @{
                Uri             = "http://$ADMIP/nitro/v1/config/logout"
                Method          = "Post"
                Body            = "object="+$Logout
                WebSession      = $ADMSession
                ContentType     = "application/json"
            }
            $logoutResponse = Invoke-RestMethod @invokeRestMethodParams
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct management IP-address. " + $_.Exception.Message)
            Break
        }
        catch
        {
            # Debug: 
            Write-Debug $_.Exception | Format-List -Force
            # Error:
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Break
        }
        # Check for REST API errors
        If ($logoutResponse.errorcode -eq 0)
        {
            Write-Verbose "ADM-logout: logout successful"
        }
    #endregion Logout of ADM with NITRO
}

function Get-ADMHDXInsightICASession {
    <#
    .SYNOPSIS
        Retrieve Citrix ADM HDX Insight information for active ICA sessions.
    .DESCRIPTION
        Retrieve Citrix ADM HDX Insight information for active ICA sessions, using the Invoke-RestMethod cmdlet for the REST API calls. 
    .EXAMPLE
        Get-ADMHDXInsightActiveICASession -ADMIP string -ADMSession SessionVariable
    .EXAMPLE
        Get-ADMHDXInsightActiveICASession -ADMIP string -ADMSession SessionVariable -Verbose
    .LINK
        https://developer-docs.citrix.com/projects/citrix-adm-nitro-api-reference/en/12.1/configuration/analytics/hdx-insight/active_ica_session/
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-11-12
        Purpose:        Script created for Citrix ADM Troubleshooting &amp; Management
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Citrix ADM management IP address'
        )]
        [ValidateScript({$_ -match [IPAddress]$_ })]
        [string] $ADMIP,
      
        [Parameter(
            Position=1, 
            Mandatory=$true, 
            HelpMessage='Enter variable, containing the Session information for the Citrix ADM'
        )]
        [System.Management.Automation.PSObject] $ADMSession,
      
        [Parameter(
            Position=2, 
            Mandatory=$false, 
            HelpMessage='Enter a username'
        )]
        [string] $Username
    )    

    #region Retrieve HDX Insight for ICA sessions with NITRO
#        $uri = "http://$ADMIP/nitro/v1/appflow/ica_session&#x8;"
        $uri = "http://$ADMIP/nitro/v1/config/ica_session"

        if ($Username)
        {
            $uri = $uri + "?args=ica_user_name:" + $([System.Web.HTTPUtility]::UrlEncode($Username)) + "&amp;asc=no&amp;order_by=session_setup_time&amp;pagesize=25&amp;type=session_setup_time&amp;cr_enabled=0&amp;sla_enabled=0&amp;duration=last_1_hour"
        }

        try
        {
            # Connect to the ADM
            $invokeRestMethodParams = @{
                Uri             = $uri
                Method          = "Get"
                WebSession      = $ADMSession
                ContentType     = "application/json"
            }
            $hdxInsightResponse = Invoke-RestMethod @invokeRestMethodParams
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct management IP-address. " + $_.Exception.Message)
            Break
        }
        catch
        {
            # Debug: 
            Write-Debug $_.Exception | Format-List -Force
            # Error:
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Break
        }
        # Check for REST API errors
        If ($hdxInsightResponse.errorcode -eq 0)
        {
            Write-Verbose "ADM-HDXInsight: successful"
        }
        return $hdxInsightResponse
    #endregion Retrieve HDX Insight for active ICA sessions with NITRO
}

function ConvertEpochToDateTime {
    <#
    .SYNOPSIS
        Convert Epoch (or unix) time to local datetime string.
    .DESCRIPTION
        Convert Epoch (or unix) time to local datetime string.
    .EXAMPLE
        ConvertEpochToDateTime -EpochTime double 
    .EXAMPLE
        ConvertEpochToDateTime -EpochTime double -Verbose
    .LINK
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-11-15
        Purpose:        Script created for Citrix ADM Troubleshooting &amp; Management
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the epoch time'
        )]
        [ValidateNotNullOrEmpty()]
        [double] $ePochTime
    )    

#    # Convert epoch/unix time to datetime (and take local timezone into account)
    $baseDateTime = New-Object -Type DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, 0
    $Offset = [TimeZoneInfo]::Local | Select BaseUtcOffset
    $utcTime = $(Get-Date $($baseDateTime.AddSeconds($ePochTime)) -Format r)
    $localTime = Get-Date($($off.BaseUtcOffset) + $(Get-Date $($baseDateTime.AddSeconds($ePochTime)) -Format r)) #-UFormat %c 
    return $localTime
}


#region ControlUp Script Standards - version 0.2
    #Requires -Version 5.1
    # Configure a larger output width for the ControlUp PowerShell console
    [int]$outputWidth = 400
    # Altering the size of the PS Buffer
    $PSWindow = (Get-Host).UI.RawUI
    $WideDimensions = $PSWindow.BufferSize
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions

    # Ensure Debug information is shown, without the confirmation question after each Write-Debug
    If ($PSBoundParameters['Debug']) {$DebugPreference = "Continue"}
    If ($PSBoundParameters['Verbose']) {$VerbosePreference = "Continue"}
    $ErrorActionPreference = "Stop"
#endregion

# ------------
# | Workflow |
# ------------

# Force PowerShell to use TLS 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# ------------
# Step 1 - Create a Citrix ADM session (based on provided NSIP)
# ------------

# Get ADM Credentials (to pass on to the login function)
[System.Management.Automation.PSCredential]$ADMCreds = $null
$ADMCreds = Get-Credential -Message "Enter your management credentials for the Citrix ADM"

# Logon to the ADM
$WebSession = Invoke-ADMLogin -ADMIP $ADMIP -ADMCredentials $ADMCreds #-Verbose

# ------------
# Step 2 - Retrieve HDX Insight data for a specific (current) user session
# ------------

$icaSession = Get-ADMHDXInsightICASession -ADMIP $ADMIP -ADMSession $WebSession -Username $HDXUserName #-Verbose

# Present the information:
Write-Host "Citrix ADM HDX Insight session details for $($HDXUserName): " -ForegroundColor Yellow
#$icaSession.ica_session | Select serverside_packet_retransmits,user_type,is_msi,count_usb_rejected,ip_block_name,launch_duration,latitude,l7_threshold_configure_value,longitude,region_code,region,host_delay,total_bytes,is_active,client_ip_address,pn_agent_version,client_latency,client_hostname,count_usb_accepted,usb_status,client_side_ns_delay,server_side_ns_delay,application_enumeration_duration,session_reconnect,serverside_rto,l7_threshold_max_server_breach,duration_summary_bandwidth,country,session_hop_diagram,clientside_packet_retransmits,server_ip_address,count_usb_stopped,rpt_sample_time,clientside_0_win,session_setup_time,bandwidth,clientside_rto,ica_app_name,client_version,clientside_cb,session_rtt,ica_user_name,country_code,l7_threshold_breach_count,state,client_type,euem,l7_clientside_latency,device_type,server_jitter,id,client_jitter,server_latency,is_multi_hop,serverside_0_win,l7_threshold_max_client_breach,serverside_cb,l7_threshold_avg_client_breach,session_type,acr_count,sr_reconnect_count,ha_failover_count,selected_time_totalte,l7_monitoring_supported,l7_threshold_avg_server_breach,client_srtt,edt_type,ica_device_ip_address,city,receiver_version,up_time,client_tx_bytes,session_setup_time_local,client_rx_bytes,server_srtt,session_end_time,l7_serverside_latency | Format-List
$icaSession.ica_session | Select @{Name='User'; Expression={$_.ica_user_name}}, 
    # User information
    @{Name='User access type'; Expression={$_.user_type}}, 
    @{Name='Session id'; Expression={$_.id}},
    @{Name='Session type'; Expression={if($_.session_type -eq 1){return "Desktop"}elseif($_.session_type -eq 0){return "Application"}else{$_.session_type}}},

    # RTT, latency and bandwidt
    @{Name='ICA RTT'; Expression={"{0:N2} ms" -f ($([math]::Round($_.session_rtt,2)))}},
    @{Name='WAN latency'; Expression={if($_.client_latency -eq -1){"-NA-"}else{"{0:N2} ms" -f ($([math]::Round($_.client_latency,2)))}}},
    @{Name='DC latency'; Expression={if($_.server_latency -eq -1){"-NA-"}else{"{0:N2} ms" -f ($([math]::Round($_.server_latency,2)))}}},
    @{Name='Host delay'; Expression={if($_.host_delay -eq -1){"-NA-"}else{"{0:N2} ms" -f ($([math]::Round($_.host_delay,2)))}}},
    @{Name='Bandwidth per interval'; Expression={"{0:N2} Kbps" -f ($([math]::Round($($_.bandwidth)/1kb,2)))}},
    @{Name='Session bandwidth'; Expression={"{0:N2} Kbps" -f ($([math]::Round($($_.duration_summary_bandwidth)/1kb,2)))}},

    # Transmitted bytes
    @{Name='Total bytes'; Expression={"{0:N2} MB" -f (($_.total_bytes)/1MB)}},
    @{Name='Bytes per interval'; Expression={"{0:N2} MB" -f (($_.selected_time_total_byte)/1MB)}},

    # Session time
    @{Name='Start time'; Expression={ConvertEpochToDateTime($_.session_setup_time)}},
    #@{Name='session end time'; Expression={if($_.session_end_time -eq -1){"-NA-"}else{ConvertEpochToDateTime($_.session_end_time)}}},
    @{Name='Uptime'; Expression={ "$((New-TimeSpan -Seconds $($_.up_time)).Hours) h`: $((New-TimeSpan -Seconds $($_.up_time)).Minutes) m`: $((New-TimeSpan -Seconds $($_.up_time)).Seconds)s" }},

    # IP-addresses
    @{Name='Client hostname'; Expression={$_.client_hostname}},
    @{Name='Client type'; Expression={$_.client_type}},
    @{Name='Client IP address'; Expression={$_.client_ip_address}},
    @{Name='Server IP address'; Expression={$_.server_ip_address}},
    @{Name='Citrix ADC IP address'; Expression={$_.ica_device_ip_address}},

    # L7 latency
    @{Name='L7 client side latency'; Expression={"{0:N2} ms" -f ($([math]::Round($_.l7_clientside_latency,2)))}},
    @{Name='L7 server side latency'; Expression={"{0:N2} ms" -f ($([math]::Round($_.l7_serverside_latency,2)))}},

    # network packets information
    @{Name='server side packet retransmit'; Expression={if($_.serverside_packet_retransmits -eq -1){"-NA-"}else{$_.serverside_packet_retransmits}}},
    @{Name='client side packet retransmit'; Expression={if($_.clientside_packet_retransmits -eq -1){"-NA-"}else{$_.clientside_packet_retransmits}}},
    @{Name='client side RTO'; Expression={if($_.clientside_rto -eq -1){"-NA-"}else{$_.clientside_rto}}},
    @{Name='server side RTO'; Expression={if($_.serverside_rto -eq -1){"-NA-"}else{$_.serverside_rto}}},
    @{Name='client side zero window size event'; Expression={if($_.clientside_0_win -eq -1){"-NA-"}else{$_.clientside_0_win}}},
    @{Name='server side zero window size event'; Expression={if($_.serverside_0_win -eq -1){"-NA-"}else{$_.serverside_0_win}}},

    # Session characteristics
    @{Name='HA failover count'; Expression={$_.ha_failover_count}},
    @{Name='EDT type'; Expression={$_.edt_type}},
    @{Name='Active'; Expression={$_.is_active}},
    @{Name='Multi-hop'; Expression={$_.is_multi_hop}},
    @{Name='End user experience monitoring'; Expression={$_.euem}} | Format-List

## Logoff from the ADM (using the global variable SessionID) (because the other method keeps timing out)
Invoke-ADMLogout -ADMIP $ADMIP -ADMSession $WebSession #-Verbose

