function Get-CUStoredCredential {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The system the credentials will be used for.")]
        [string]$System
    )

    # Get the stored credential object
    $strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
    
    try {
        Import-Clixml $strCUCredFolder\$($env:USERNAME)_$($System)_Cred.xml
    }
    catch {
        Write-Error ("The required PSCredential object could not be loaded. " + $_)
    }
}


function Get-NSHASyncStatus ()
{
    <#
    .SYNOPSIS
        Retrieve NetScaler High Availability (sync) status.
    .DESCRIPTION
        Retrieve NetScaler High Availability (sync) status, using the Invoke-RestMethod cmdlet for the REST API calls. 
        This script checks the high Availability status for errors and any configuration errors that might have occurs, especially in regards to the HA sync process. 
        A message is returned to specify the state and whether or not a HA Failover is supported with the current configuration.
    .EXAMPLE
        Get-HASyncStatus -NSIP 192.168.0.101 -NSCredentials $PSCredentialsObject
    .EXAMPLE
        Get-HASyncStatus -NSIP 192.168.0.101
    .EXAMPLE
        Get-HASyncStatus -NSIP 192.168.0.101 -Verbose -Debug
    .CONTEXT
        NetScalers
    .MODIFICATION_HISTORY
        Esther Barthel, MSc - 22/07/19 - Original code
        Esther Barthel, MSc - 27/07/19 - Adding detailed HA config information to the result message
        Esther Barthel, MSc - 27/07/19 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
        Esther Barthel, MSc - 29/07/19 - Improving error handling and returned information
        Esther Barthel, MSc - 30/07/19 - Added Import-CliXml cmdlet for Automated Action support with stored NSCredentials.
        Esther Barthel, MSc - 04/08/19 - Using Switch to chck for different node master/hasstatus/hasync states
        Esther Barthel, MSc - 04/08/19 - Using Switch to chck for different node master/hasstatus/hasync states
        Esther Barthel, MSc - 23/08/19 - Added the new location for the credentials file, based on the standardized function New-CUStoredCredential
    .LINK
        https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Utility/Invoke-RestMethod?view=powershell-6
        https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Utility/Import-Clixml?view=powershell-6
    .COMPONENT
        Set-ADCCredentialsXML.ps1 - to create the XML Credentials file for non-interactive/automated use of this script
    .NOTES
        Version:        0.7
        Author:         Esther Barthel, MSc
        Creation Date:  2019-07-22
        Updated:        2019-08-23
                        Added the new location for the credentials file, based on the standardized function New-CUStoredCredential
        Purpose:        Script Action, created for ControlUp NetScaler Monitoring
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the NetScaler IP address to run the script on'
        )]
        [ValidateScript({$_ -match [IPAddress]$_ })]
        [string] $NSIP,
      
        [Parameter(
            Position=1, 
            Mandatory=$false, 
            HelpMessage='Enter a PSCredential object, containing the username and password'
        )]
        [System.Management.Automation.CredentialAttribute()] $ADCCredentials
    )    

    #region ControlUp Script Standards - version 0.2
        #Requires -Version 3.0

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

    #region script settings
        # Stored ADC Credentials XML file
        $systemName = "ADC"
        $credTargetFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
        $credTarget = "$credTargetFolder\$($Env:Username)_$($systemName)_Cred.xml"
        # Declare ADC Credentials object
        [System.Management.Automation.PSCredential]$adcCredentials = $null
        # NITRO Constants
        $ContentType = "application/json"
        # turn Verbose mode on
        #$VerbosePreference = "Continue"
        # turn Verbose mode off
        #$VerbosePreference="SilentlyContinue"
    #endregion

    Write-Verbose ""
    Write-Verbose "--------------------------------------------- "
    Write-Verbose "| Run Citrix ADC HA healthcheck with NITRO: | "
    Write-Verbose "--------------------------------------------- "
    Write-Verbose ""

    # Load NSCredentials either trough XML file import or Get-Credentials
    If ($null -eq $ADCCredentials)
    {
        # Check for Stored Credentials
        If (Test-Path -Path $credTarget)
        # Stored credentials found, import credentials
        {
            try
            {
                #$adcCredentials = Import-Clixml -Path $credTarget
                $ADCCredentials = Get-CUStoredCredential -System $systemName
            }
            catch
            {
                Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
                Exit
            }
            Write-Verbose "* PSCredentials: Stored $systemName credentials XML file found. Credentials imported for Automated Action support."
        }
        Else
        # No Stored Credentials Found, ask for credentials
        {
            Write-Verbose "* PSCredentials: NO stored $systemName credentials XML file found, using Get-Credential."
            $ADCCredentials = Get-Credential -Message "Enter your Credentials for Citrix ADC $NSIP"
        }
    }

    # Retieving username and password from PSCredentials for use with NITRO
    $NSUserName = $ADCCredentials.UserName
    $NSUserPW = $ADCCredentials.GetNetworkCredential().Password

    # ----------------------------------------
    # | Method #1: Using the SessionVariable |
    # ----------------------------------------
    #region Start NITRO Session
        #Force PowerShell to bypass validation for (self-signed) certificates and SSL connections
        # source: https://blogs.technet.microsoft.com/bshukla/2010/04/12/ignoring-ssl-trust-in-powershell-system-net-webclient/ 
        Write-Verbose "* Certificate Validation: Forcing PowerShell to trust all certificates (including the self-signed netScaler certificate)"
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

        #JSON payload
        $Login = ConvertTo-Json @{
            "login" = @{
                "username"=$NSUserName;
                "password"=$NSUserPW
            }
        }
        try
        {
            # Login to the NetScaler and create a session (stored in $NSSession)
            $invokeRestMethodParams = @{
                Uri             = "https://$NSIP/nitro/v1/config/login"
                Body            = $Login
                Method          = "Post"
                SessionVariable = "NSSession"
                ContentType     = $ContentType
            }
            $loginresponse = Invoke-RestMethod @invokeRestMethodParams
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct NetScaler IP-address. " + $_.Exception.Message)
            Exit
        }
        catch
        {
            If (($_.Exception.Message -like "*Unauthorized*") -and (Test-Path -Path $XMLFile))
            {
                Write-Warning "XML Stored NSCredentials were used, check if the stored NSCredentials are correct for this Citrix ADC (" + $NSIP + ")."
            }
            # Debug: 
            Write-Debug $_.Exception | Format-List -Force
            # Error:
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        # Check for REST API errors
        If ($loginresponse.errorcode -eq 0)
        {
            Write-Verbose "* NITRO: login successful"
        }
    #endregion Start NITRO Session


    # -----------------------
    # | HA Node information |
    # -----------------------
    #region Get HA Node stat
        # Method #1: Making the REST API call to the NetScaler
        try
        {
            # start with clean response variable
            $haNodeStats = $null
            # Get HA node stats from NetScaler (no payload with GET)
            $invokeRestMethodParams = @{
                Uri         = "https://$NSIP/nitro/v1/stat/hanode"
                Method      = "Get"
                WebSession  = $NSSession
                ContentType = $ContentType
            }
            $haNodeStats = Invoke-RestMethod @invokeRestMethodParams
        }
        catch
        {
            Write-Error $error[0] 
            Exit          
        }

        If ( -not ($haNodeStats.errorcode -eq 0))
        # NITRO errorcode found (or $response is empty)
        {
            Write-Error "NITRO: errorcode (" + $haNodeStats.errorcode + ") " + $haNodeStats.message + "."
            Exit
        }

        Write-Verbose "* NITRO: hanode stat information retrieved successful"
        If ( -not ($haNodeStats.hanode))
        # NITRO did not return hanode information
        {
            Write-Error "NITRO: NO hanode information found."
            Exit
        }
        # Check for HA Sync Errors (hanode stats), generate a warning message if found
        If ($haNodeStats.hanode.haerrsyncfailure -gt 0)
        # HA Sync error found (tested by changing a RPC node password)
        {
            Write-Warning "HA Healthcheck: HA Synchronization failure(s) found since Last Transition, which can result in a mismatched configuration."
            Write-Host "Note: If the Synchronization State of both nodes is SUCCESS or ENABLED, ignore the sync failure as it likely occured in the past."-ForegroundColor Yellow
            Write-Host "      (The HA Sync failure statistic is not reset untill the next master state transition)." -ForegroundColor Yellow 
        }
        # HA Sync errors = 0

        # Check for propagation errors
        If ($haNodeStats.hanode.haerrproptimeout -gt 0)
        # HA Propagation timeouts found - could not find information for test scenario, so no recommendation added
        {
            Write-Warning "HA Healthcheck: HA Propagation timeouts found, which can result in a mismatched configuration."
            Write-Host "Cause: Propagation errors are caused by either: " -ForegroundColor Yellow
            Write-Host "       - network connectivity error -> check interface settings." -ForegroundColor Yellow
            Write-Host "       - missing resource on the secondary node." -ForegroundColor Yellow
            Write-Host "       - authentication failure -> ensure nsroot/RPC node passwords are identical on both nodes." -ForegroundColor Yellow
        }
    #endregion

    #region Get HA Node config
        If ($haNodeStats.hanode.hacurstatus -ne "YES")
        # HA status -ne YES
        {
            Write-Warning "HA healthcheck: High Availability is NOT configured for NetScaler $NSIP"
            Write-Output "`t * HA status = $($haNodeStats.hanode.hacurstatus)"
            Exit 
        }

        #HA status -eq YES
        try
        {
            # start with clean response variable
            $haNodeConfig = $null
            # Get HA node config from NetScaler (no payload with GET)
            $invokeRestMethodParams = @{
                Uri         = "https://$NSIP/nitro/v1/config/hanode"
                Method      = "Get"
                WebSession  = $NSSession
                ContentType = $ContentType
            }
            # Method #1: Making the REST API call to the NetScaler
            $haNodeConfig = Invoke-RestMethod @invokeRestMethodParams
        }
        catch
        {
            Write-Error $error[0]
            Exit 
        }

        If ( -not ($haNodeConfig.errorcode -eq 0))
        # NITRO errorcode found (or $response is empty)
        {
            Write-Error "NITRO: errorcode (" + $haNodeConfig.errorcode + ") " + $haNodeConfig.message + "."
            Exit
        }
        Write-Verbose "* NITRO: hanode config information retrieved successful"
        # errorcode -eq 0

        # Check if required HA node info is returned
        If ( -not ($haNodeConfig.hanode))
        # NITRO did not return hanode information
        {
            Write-Error "NITRO: NO hanode config information found."
            Exit
        }
        # hanode contains information
                
        # Check if config for both nodes is returned
        If ( -not ($haNodeConfig.hanode.Count -eq 2))
        # only one hanode config node was found
        {
            Write-Error "HA healthcheck: Only received information for one HA node, health NOT determined."
            Exit
        }
        # hanode config contains both nodes

        # Check actual config information, splitting information into two nodes
        Write-Verbose "* NITRO: hanode config information for both nodes retrieved successful"
        $ID0 = $haNodeConfig.hanode | Where-Object {$_.id -eq 0}
        $ID1 = $haNodeConfig.hanode | Where-Object {$_.id -eq 1}

        # Check node 0 master state
        switch (($ID0.state).ToUpper())
        {
            "PRIMARY"
            # node 0 master state = PRIMARY 
            {
                Write-Verbose "* HA Healthcheck: node 0 master state = PRIMARY"
                # Check node 1 ha status (Node State)
                switch (($ID1.hastatus).ToUpper())
                {
                    "STAYSECONDARY" 
                    {
                        # Check node 1 ha status
                        Write-Verbose "* HA Healthcheck: node 1 ha status: STAYSECONDARY"
                        # check node 1 ha sync
                        Switch (($ID1.hasync).ToUpper())
                        {
                            "SUCCESS"
                            # all green
                            {
                                Write-Verbose "* HA Healthcheck: node 1 Synchronization State: SUCCESS"
                                Write-Host "HA Healthcheck: No High Availability errors found." -ForegroundColor Green
                                Break
                            }
                            "ENABLED"
                            # first sync not run yet
                            {
                                Write-Verbose "* HA Healthcheck: node 1 Synchronization State: ENABLED"
                                # Check received heartbeats
                                If ($haNodeStats.hanode.hapktrxrate -eq 0)
                                # node 1 heartbeats received = 0
                                {
                                    Write-Warning "HA Healthcheck: NO received heartbeats"
                                    Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state."
                                    Break
                                }
                                # Check sent heartbeats
                                If ($haNodeStats.hanode.hapkttxrate -eq 0)
                                # node 1 heartbeats sent = 0
                                {
                                    Write-Warning "HA Healthcheck: NO sent heartbeats"
                                    Write-Host "Cause: Incorrect interface settings can result in heartbeats not being sent." -ForegroundColor Yellow
                                    Write-Host "       Check the HA monitoring & heartbeat settings of the interfaces." -ForegroundColor Yellow
                                    Break
                                }
                                Write-Warning "HA Healthcheck: Synchronization is not (yet) performed."
                                Write-Host "Cause: the initial synchronization is not (yet) performed. A forced synchronization should change the status to SUCCESS."
                                Break
                            }
                            default
                            # node 1 sync state = unknown
                            {
                                Write-Warning ("HA Healthcheck: node 1 Synchronization state: " + $ID1.hasync)
                                # hasync = FAILED
                                If ($ID1.hasync -eq "FAILED")
                                {
                                    Write-Host "Cause: HA Synchronization failures occur when the nsroot password or RPC node password is not identical for both nodes." -ForegroundColor Yellow
                                    Break
                                }
                                # Check received heartbeats
                                If ($haNodeStats.hanode.hapktrxrate -eq 0)
                                # node 1 heartbeats received = 0
                                {
                                    Write-Warning "HA Healthcheck: NO received heartbeats"
                                    Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state." -ForegroundColor Yellow
                                    Break
                                }
                                Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state." -ForegroundColor Yellow
                                Break
                            }
                        }
                    }
                    "UP" 
                    {
                        # Check node 1 ha status
                        Write-Verbose "* HA Healthcheck: node 1 ha status: UP"
                        # check node 1 ha sync
                        Switch (($ID1.hasync).ToUpper())
                        {
                            "SUCCESS"
                            # all green
                            {
                                Write-Verbose "* HA Healthcheck: node 1 Synchronization State: SUCCESS"
                                Write-Host "HA Healthcheck: No High Availability errors found." -ForegroundColor Green
                                Break
                            }
                            "ENABLED"
                            # ENABLED = expected good state 
                            {
                                Write-Verbose "* HA Healthcheck: node 1 Synchronization State: ENABLED"
                                # Check received heartbeats
                                If ($haNodeStats.hanode.hapktrxrate -eq 0)
                                # node 1 heartbeats received = 0
                                {
                                    Write-Warning "HA Healthcheck: NO received heartbeats"
                                    Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state."
                                    Break
                                }
                                Write-Warning "HA Healthcheck: Synchronization is not (yet) performed."
                                Write-Host "Cause: the initial synchronization is not (yet) performed. A forced synchronization should change the status to SUCCESS." -ForegroundColor Yellow
                                Break
                            }
                            default
                            # node 1 sync state = unknown
                            {
                                Write-Warning ("HA Healthcheck: node 1 Synchronization state: " + $ID1.hasync)
                                # hasync = FAILED
                                If ($ID1.hasync -eq "FAILED")
                                {
                                    Write-Host "Cause: HA Synchronization failures occur when the nsroot password or RPC node password is not identical for both nodes." -ForegroundColor Yellow
                                    Break
                                }
                                If ($ID1.hasync -eq "IN PROGRESS")
                                # node 1 hasync INPROGRESS
                                {
                                    Write-Host "Cause:   When a HA synchronization is active, the INPROGRESS status is shown" -ForegroundColor Yellow
                                    Write-Host "         Rerun the Healthcheck to see if the status changes to SUCCESS" -ForegroundColor Yellow
                                    Break
                                }
                                # Check received heartbeats
                                If ($haNodeStats.hanode.hapktrxrate -eq 0)
                                # node 1 heartbeats received = 0
                                {
                                    Write-Warning "HA Healthcheck: NO received heartbeats"
                                    Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state." -ForegroundColor Yellow
                                    Break
                                }
                                Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state." -ForegroundColor Yellow
                                Break
                            }
                        }
                    }
                    default 
                    {
                        Write-Warning ("HA Healthcheck: node 1 ha status: " + $ID1.hastatus)
                        # Check received heartbeats
                        If ($haNodeStats.hanode.hapktrxrate -eq 0)
                        # node 1 heartbeats received = 0
                        {
                            Write-Warning "HA Healthcheck: NO received heartbeats"
                            Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state." -ForegroundColor Yellow
                            Write-Host "       Check the HA monitoring & heartbeat settings of the interfaces." -ForegroundColor Yellow
                            Break
                        }
                        Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state." -ForegroundColor Yellow
                        Write-Host "       Check the HA monitoring & heartbeat settings of the interfaces." -ForegroundColor Yellow
                        Break
                    }
                }          
            }
            "SECONDARY"
            # node 0 master state = SECONDARY 
            {
                Write-Verbose "* HA Healthcheck: node 0 master state: SECONDARY"
                # Check node 1 ha status (Node State)
                switch (($ID1.hastatus).ToUpper())
                {
                    "STAYPRIMARY" 
                    {
                        # Check node 1 ha status
                        Write-Verbose "* HA Healthcheck: node 1 ha status: STAYPRIMARY"
                        # check node 1 ha sync
                        Switch (($ID1.hasync).ToUpper())
                        {
                            "SUCCESS"
                            # all green
                            {
                                Write-Verbose "* HA Healthcheck: node 1 Synchronization State: SUCCESS"
                                Write-Host "HA Healthcheck: No High Availability errors found." -ForegroundColor Green
                                Break
                            }
                            "ENABLED"
                            # all green
                            {
                                Write-Verbose "* HA Healthcheck: node 1 Synchronization State: ENABLED"
                                If (($ID0.state).ToUpper() -eq "SECONDARY")
                                # node 0 master state is SECONDARY; node 1 is configured to STAYPRIMARY and cannot have a SECONDARY node state! 
                                {
                                    Write-Verbose "* HA Healthcheck: node 0 master state: SECONDARY"
                                    Write-Warning "HA Healthcheck: node 0 master state is SECONDARY, while it is configured STAYPRIMARY."
                                    Write-Host "Cause: Incorrect interface settings can result in a secondary master state for both nodes." -ForegroundColor Yellow
                                    Write-Host "       Check the HA monitoring & heartbeat settings of the interfaces." -ForegroundColor Yellow
                                    Break
                                }
                                # Check sent heartbeats
                                If ($haNodeStats.hanode.hapkttxrate -eq 0)
                                # node 1 heartbeats sent = 0
                                {
                                    Write-Warning "HA Healthcheck: NO sent heartbeats"
                                    Write-Host "Cause: Incorrect interface settings can result in heartbeats not being sent." -ForegroundColor Yellow
                                    Write-Host "       Check the HA monitoring & heartbeat settings of the interfaces." -ForegroundColor Yellow
                                    Break
                                }
                                # Check received heartbeats
                                If ($haNodeStats.hanode.hapktrxrate -eq 0)
                                # node 1 heartbeats received = 0
                                {
                                    Write-Warning "HA Healthcheck: NO received heartbeats"
                                    Write-Host "Cause: Incorrect interface settings can result in heartbeats not being received." -ForegroundColor Yellow
                                    Write-Host "       Check the HA monitoring & heartbeat settings of the interfaces." -ForegroundColor Yellow
                                    Break
                                }
                                Write-Host "HA Healthcheck: No High Availability errors found." -ForegroundColor Green
                                Break
                            }
                            default
                            # node 1 sync state = unknown
                            {
                                Write-Warning ("HA Healthcheck: node 1 Synchronization state: " + $ID1.hasync)
                                # hasync = FAILED
                                If ($ID1.hasync -eq "FAILED")
                                {
                                    Write-Host "Cause: HA Synchronization failures occur when the nsroot password or RPC node password is not identical for both nodes." -ForegroundColor Yellow
                                    Break
                                }
                                # Check received heartbeats
                                If ($haNodeStats.hanode.hapktrxrate -eq 0)
                                # node 1 heartbeats received = 0
                                {
                                    Write-Warning "HA Healthcheck: NO received heartbeats"
                                    Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state."
                                    Break
                                }
                                Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state."
                                Break
                            }
                        }
                    }
                    "UP" 
                    {
                        # Check node 1 ha status
                        Write-Verbose "* HA Healthcheck: node 1 ha status: UP"
                        # check node 1 ha sync
                        Switch (($ID1.hasync).ToUpper())
                        {
                            "SUCCESS"
                            {
                                Write-Verbose "* HA Healthcheck: node 1 Synchronization State: SUCCESS"
                                Break
                            }
                            "ENABLED"
                            {
                                Write-Verbose "* HA Healthcheck: node 1 Synchronization State: ENABLED"
                                # hasync = FAILED
                                If ($ID0.hasync -eq "FAILED")
                                {
                                    Write-Warning "HA Healthcheck: node 0 HA synchronization FAILED."
                                    Write-Host "Cause: HA Synchronization failures occur when the nsroot password or RPC node password is not identical for both nodes." -ForegroundColor Yellow
                                    Break
                                }
                                If (($ID0.hastatus).ToUpper() -eq "DISABLED")
                                # node 0 master state is SECONDARY; node 0 ha status is DISABLED. 
                                {
                                    Write-Warning "HA Healthcheck: node 0 ha status: DISABLED."
                                    Write-Host "Cause: node 0 is manually DISABLED and does not participate in High Availability." -ForegroundColor Yellow
                                    Write-Host "       Check the HA configuration of node 0." -ForegroundColor Yellow
                                    Break
                                }
                                If ((($ID0.state).ToUpper() -eq "SECONDARY") -and (($ID1.state).ToUpper() -eq "SECONDARY"))
                                # node 0 master state is SECONDARY; node 1 cannot have a SECONDARY node state! 
                                {
                                    Write-Verbose "* HA Healthcheck: node 0 master state: SECONDARY"
                                    Write-Warning "HA Healthcheck: node 0 master state is SECONDARY."
                                    Write-Host "Cause: Incorrect interface settings can result in a secondary master state for both nodes." -ForegroundColor Yellow
                                    Write-Host "       Check the HA monitoring & heartbeat settings of the interfaces." -ForegroundColor Yellow
                                    Break
                                }
                                # Check received heartbeats
                                If ($haNodeStats.hanode.hapktrxrate -eq 0)
                                # node 1 heartbeats received = 0
                                {
                                    Write-Warning "HA Healthcheck: NO received heartbeats"
                                    Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state."
                                    Break
                                }
                                # Check sent heartbeats
                                If ($haNodeStats.hanode.hapkttxrate -eq 0)
                                # node 1 heartbeats sent = 0
                                {
                                    Write-Warning "HA Healthcheck: NO sent heartbeats"
                                    Write-Host "Cause: Incorrect interface settings can result in heartbeats not being sent." -ForegroundColor Yellow
                                    Write-Host "       Check the HA monitoring & heartbeat settings of the interfaces." -ForegroundColor Yellow
                                    Break
                                }
                                Write-Host "HA Healthcheck: No High Availability errors found." -ForegroundColor Green
                                Break
                            }
                            default
                            # node 1 sync state = unknown
                            {
                                Write-Warning ("HA Healthcheck: node 1 Synchronization state: " + $ID1.hasync)
                                # hasync = FAILED
                                If ($ID1.hasync -eq "FAILED")
                                {
                                    Write-Host "Cause: HA Synchronization failures occur when the nsroot password or RPC node password is not identical for both nodes." -ForegroundColor Yellow
                                    Break
                                }
                                # Check received heartbeats
                                If ($haNodeStats.hanode.hapktrxrate -eq 0)
                                # node 1 heartbeats received = 0
                                {
                                    Write-Warning "HA Healthcheck: NO received heartbeats"
                                    Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state."
                                    Break
                                }
                                Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state." -ForegroundColor Yellow
                                Break
                            }
                        }
                    }
                    default 
                    {
                        Write-Warning ("* HA Healthcheck: node 1 ha status: " + $ID1.hastatus)
                        # Check received heartbeats
                        If ($haNodeStats.hanode.hapktrxrate -eq 0)
                        # node 1 heartbeats received = 0
                        {
                            Write-Warning "HA Healthcheck: NO received heartbeats"
                            Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state."
                            Break
                        }
                        If ($ID1.hasync -eq "IN PROGRESS")
                        # node 1 hasync INPROGRESS
                        {
                            Write-Warning ("HA Healthcheck: HA synchronization in progress.")
                            Write-Host "Cause:   When a HA synchronization is active the INPROGRESS status is shown" -ForegroundColor Yellow
                            Write-Host "         Rerun the Healthcheck to see if the status changes to SUCCESS" -ForegroundColor Yellow
                        }
                        Write-Host "Cause: Network connectivity errors can cause an UNKNOWN synchronization state." -ForegroundColor Yellow
                        Break
                    }
                }          


            }
            default 
            {
                Write-Warning ("* HA Healthcheck: Unknown node 0 master state: " + $ID0.state)
                Break
            }
        }          
         
        # --------------------------
        # | HA Healthcheck metrics |
        # --------------------------
        # Create (color-coded) overview of different hanode config metrics
        Foreach ($ID in $haNodeConfig.hanode)
        {
            # node 0 (NSIP) master state: $ID.state
            Write-Host ("`t * node $($ID.id) (" + $ID.ipaddress + ") Master State: ") -NoNewline
            If ($ID.state -eq "Primary")
            {
                Write-Host ($ID.state) -ForegroundColor Green
            }
            Else
            {
                Write-Host ($ID.state) -ForegroundColor Yellow
            }

            #   - HA state: $ID.hastatus
            Write-Host ("`t`t - Node State              : ") -NoNewline
            If ($ID.hastatus -eq "UP")
            {
                Write-Host ($ID.hastatus) -ForegroundColor Green
            }
            Elseif ($ID.hastatus -like "*UNKNOWN*")
            {
                Write-Host ($ID.hastatus) -ForegroundColor Red
            }
            Else
            {
                Write-Host ($ID.hastatus) -ForegroundColor Yellow
            }

            #   - HA state sync status: $ID.hasync
            Write-Host ("`t`t - Synchronization State   : ") -NoNewline
            If ($ID.hasync -eq "SUCCESS")
            {
                Write-Host ($ID.hasync) -ForegroundColor Green
            }
            ElseIf ($ID.hasync -eq "FAILED")
            {
                Write-Host ($ID.hasync) -ForegroundColor Red
            }
            Else
            {
                Write-Host ($ID.hasync) -ForegroundColor Yellow
            }

            #   - HAMON interfaces: $ID.hamonifaces 
            Write-Host ("`t`t - Monitoring ON interfaces: ") -NoNewline
            Write-Host  ("$($ID.hamonifaces)") -ForegroundColor Yellow
            #   - HAMON interfaces: $ID.haheartbeatifaces 
            Write-Host ("`t`t - Heatbeat OFF interfaces : ") -NoNewline
            Write-Host  ("$($ID.haheartbeatifaces)") -ForegroundColor Yellow
        }

        #region Create (color-coded) overview of different hanode stat metrics
            # * HA sync errors: $haNodeStats.hanode.haerrsyncfailure
            Write-Host ("`t * Sync failure            : ") -NoNewline
            If ($haNodeStats.hanode.haerrsyncfailure -ne 0)
            {
                Write-Host  ($haNodeStats.hanode.haerrsyncfailure) -ForegroundColor Red
            }
            Else
            {
                Write-Host  ($haNodeStats.hanode.haerrsyncfailure) -ForegroundColor Green
            }

            # * HA prop timeouts errors: $haNodeStats.hanode.haerrproptimeout
            Write-Host ("`t * Propagation timeouts    : ") -NoNewline
            If ($haNodeStats.hanode.haerrproptimeout -ne 0)
            {
                Write-Host  ($haNodeStats.hanode.haerrproptimeout) -ForegroundColor Red
            }
            Else
            {
                Write-Host  ($haNodeStats.hanode.haerrproptimeout) -ForegroundColor Green
            }

            # * HA heartbeats received/s: $haNodeStats.hanode.hapktrxrate
            Write-Host ("`t * Heartbeats received (/s): ") -NoNewline
            If ($haNodeStats.hanode.hapktrxrate -eq 0)
            {
                Write-Host  ($haNodeStats.hanode.hapktrxrate) -ForegroundColor Red
            }
            Else
            {
                Write-Host  ($haNodeStats.hanode.hapktrxrate) -ForegroundColor Yellow
            }

            # * HA heartbeats sent/s: ($haNodeStats.hanode.hapkttxrate
            Write-Host ("`t * Heartbeats sent (/s)    : ") -NoNewline
            If ($haNodeStats.hanode.hapkttxrate -eq 0)
            {
                Write-Host  ($haNodeStats.hanode.hapkttxrate) -ForegroundColor Red
            }
            Else
            {
                Write-Host  ($haNodeStats.hanode.hapkttxrate) -ForegroundColor Yellow
            }

            ## * HA System State: $haNodeStats.hanode.hacurstate
            #Write-Host ("`t * System State            : ") -NoNewline
            #If ($haNodeStats.hanode.hacurstate -eq 0)
            #{
            #    Write-Host  ($haNodeStats.hanode.hacurstate) -ForegroundColor Red
            #}
            #Else
            #{
            #    Write-Host  ($haNodeStats.hanode.hacurstate) -ForegroundColor Yellow
            #}

            # * HALastTransitionTime: $HALastTransitionTime
            Write-Host ("`t * Last Transition time    : ") -NoNewline
            Write-Host  ($haNodeStats.hanode.transtime) -ForegroundColor Yellow
        #endregion
    #endregion


    #region End NetScaler NITRO Session
        #Disconnect from the NetScaler (cleanup session)
        $LogOut = @{
            "logout" = @{}
        } | ConvertTo-Json

        try
        {
            # Loout of the NetScaler and remove the session (stored in $NSSession)
            $invokeRestMethodParams = @{
                Uri             = "https://$NSIP/nitro/v1/config/logout"
                Body            = $LogOut
                Method          = "Post"
                WebSession      = $NSSession
                ContentType     = $ContentType
            }
            $logoutresponse = Invoke-RestMethod @invokeRestMethodParams
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct NetScaler IP-address. " + $_.Exception.Message)
            Exit
        }
        catch
        {
            # Debug: 
            Write-Debug $_.Exception | Format-List -Force
            # Error:
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        # Check for REST API errors
        If ($logoutresponse.errorcode -eq 0)
        {
            Write-Verbose "* NITRO: logout successful"
        }
    #endregion End NetScaler NITRO Session
}

Write-Host ""

# Retrieve NSIP
$NSIP = $args[0]
Write-Host ""
Get-NSHASyncStatus -NSIP $NSIP #-Verbose -Debug


## Note: This section is added to support an automated run of the script, where any output is send to a logfile
## In case you want to log the output to a logfile too
#$logFile = $env:TEMP + "\" + $MyInvocation.MyCommand.Name + ".log"

## HA Healthcheck
#((Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString() + " - Running HA Healthcheck script") > "$logFile"
#Get-NSHASyncStatus -NSIP $NSIP *>&1 | Tee-Object "$logFile" -Append


