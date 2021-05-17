function Get-NSContentSwitchBinding {
    <#
    .SYNOPSIS
      Retrieve NetScaler Content Switching Binding information.
    .DESCRIPTION
      Retrieve NetScaler Content Switching binding information, using the Invoke-RestMethod cmdlet for the REST API calls.
    .NOTES
      Version:        0.1
      Author:         Esther Barthel, MSc
      Creation Date:  2018-05-27
      Purpose:        SBA, created for ControlUp NetScaler Monitoring

      Copyright (c) cognition IT. All rights reserved.
    #>

    [CmdletBinding()]
    Param(
      # Declaring the input parameters, provided for the SBA
      [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
      [string]
      $NSIP,

      [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$True)]
      [string]
      $CSvServer,
      
      [Parameter(Position=2, Mandatory=$true, ValueFromPipeline=$true)]
      [string]
      $NSUserName,
      
      [Parameter(Position=3, Mandatory=$true, ValueFromPipeline=$true)]
      [string]
      $NSUserPW
     )    

    #region NITRO settings
        # NITRO Constants
        $ContentType = "application/json"
    #endregion NITRO settings

    # Ensure Debug information is shown, without the confirmation question after each Write-Debug
    If ($PSBoundParameters['Debug']) {$DebugPreference = 'Continue'}

    Write-Output ""
    Write-Output "----------------------------------------------------------------- "
    Write-Output "| Retrieving CS vserver binding information from the NetScaler: | "
    Write-Output "----------------------------------------------------------------- "
    Write-Output ""

    # ----------------------------------------
    # | Method #1: Using the SessionVariable |
    # ----------------------------------------
    #region Start NetScaler NITRO Session
        #Force PowerShell to bypass the CRL check for certificates and SSL connections
            Write-Verbose "Forcing PowerShell to trust all certificates (including the self-signed netScaler certificate)"
            # source: https://blogs.technet.microsoft.com/bshukla/2010/04/12/ignoring-ssl-trust-in-powershell-system-net-webclient/ 
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

        #Connect to NetScaler VPX/MPX
        $Login = ConvertTo-Json @{"login" = @{"username"=$NSUserName;"password"=$NSUserPW}}
        try
        {
            $loginresponse = Invoke-RestMethod -Uri "https://$NSIP/nitro/v1/config/login" -Body $Login -Method POST -SessionVariable NetScalerSession -ContentType $ContentType -Verbose:$VerbosePreference -ErrorAction SilentlyContinue
        }
        Catch [System.Net.WebException]
        {
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Break
        }
        Catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct NetScaler IP-address. " + $_.Exception.Message)
            Break
        }
        Catch
        {
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
    #        echo $_.Exception | Format-List -Force
            Break
        }
        Finally
        {
            If ($loginresponse.errorcode -eq 0)
            {
                Write-Verbose "REST API call to login to NS: successful"
            }
        }
    #endregion Start NetScaler NITRO Session


    # ---------------
    # | CS bindings |
    # ---------------
    #region Get CS Bindings
        # Base URL 
        $strURI = "https://$NSIP/nitro/v1/config/csvserver_binding"

        # Specify the required full URL, including filters and arguments
        $strArgs = ""
        If ($CSvServer)
        {
            # Add Profile name to URI
            Write-Verbose ("Added the CS vServer """ + $CSvServer + """ to the URI")
            $strArgs = ("/" + $CSvServer)
        }
        Else
        {
            Write-Verbose ("Added the bulkbindings argument to the URI as no CS vServer was specified")
            $strArgs = "?bulkbindings=yes"
        }
        $strURI = $strURI + $strArgs

        # Method #1: Making the REST API call to the NetScaler
        try
        {
            # start with clean response variable
            $response = $null
            $response = Invoke-RestMethod -Method Get -Uri $strURI -ContentType $ContentType -WebSession $NetScalerSession -Verbose:$VerbosePreference -ErrorAction SilentlyContinue
        }
        catch
        {
            Write-Error ("An error (" + $_.Exception.GetType().FullName + ") occurred, with message: " + $_.Exception.Message)
            If ($DebugPreference -eq "Continue")
            {
                Write-Debug "Error full details: "
                echo $_.Exception | Format-List -Force
            }
        }
        Finally
        {
            If ($response.errorcode -eq 0)
            {
                Write-Verbose "REST API call to retrieve information: successful"
                If ($response.csvserver_binding)
                {
                    #$response.csvserver_binding | Format-List

                    If (!($response.csvserver_binding.csvserver_spilloverpolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer Spillover policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_spilloverpolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_auditnslogpolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer Audit NSLog bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_auditnslogpolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_domain_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer Domain bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_domain_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Domain name'; E={$_.domainname}}, 
                                                            @{N='Backup IP'; E={$_.backupip}}, 
                                                            @{N='TTL'; E={$_.ttl}}, 
                                                            @{N='Site domain TTL'; E={$_.sitedomainttl}}, 
                                                            @{N='Cookie domain'; E={$_.cookiedomain}}, 
                                                            @{N='Cookie timeout'; E={$_.cookietimeout}}, 
                                                            @{N='Appflow log'; E={$_.appflowlog}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_filterpolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer Filter policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_filterpolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_cmppolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer Compression policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_cmppolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_lbvserver_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer LB vServer bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_lbvserver_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Target vServer'; E={$_.targetvserver}}, 
                                                            @{N='LB vServer'; E={$_.lbvserver}}, 
                                                            @{N='Hits'; E={$_.hits}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_appflowpolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer Appflow policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_appflowpolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_responderpolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer Responder policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_responderpolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_transformpolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer Transform policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_transformpolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_vpnvserver_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer VPN vServer bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_vpnvserver_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='vServer'; E={$_.vserver}}, 
                                                            @{N='Hits'; E={$_.hits}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_feopolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer FEO policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_feopolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_authorizationpolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer Authorization policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_authorizationpolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_cachepolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer Cache policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_cachepolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_rewritepolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer Rewrite policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_rewritepolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_cspolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer CS policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_cspolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}}, 
                                                            @{N='Rule'; E={$_.rule}}, 
    #                                                        @{N='Hits'; E={$_.hits}}, 
                                                            @{N='Hits'; E={$_.hits}} | Format-Table -AutoSize
    #                                                        @{N='PI policy hits'; E={$_.pipolicyhits}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_gslbvserver_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer GSLB vServer bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_gslbvserver_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='vServer'; E={$_.vserver}}, 
                                                            @{N='Hits'; E={$_.hits}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_appqoepolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer AppQoE policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_appqoepolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_tmtrafficpolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer TM traffic policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_tmtrafficpolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_auditsyslogpolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer Audit syslog policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_auditsyslogpolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                    If (!($response.csvserver_binding.csvserver_appfwpolicy_binding.name -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer AppFW policy bindings:" -ForegroundColor Yellow
                        $response.csvserver_binding.csvserver_appfwpolicy_binding | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Target LB vServer'; E={$_.targetlbvserver}}, 
                                                            @{N='Invoke'; E={$_.invoke}} | Format-Table -AutoSize
                    }

                }
                Else
                {
                    Write-Warning "No CS binding information was found"
                }
            }
            Else
            {
                If ($response -eq $null)
                {
                    Write-Warning "No information was returned by NITRO"
                }
                Else
                {
                    Write-Warning "NITRO returned an error."
                    Write-Debug ("code: """ + $response.errorcode + """, message """ + $response.message + """")
                }
            }
        }
    #endregion


    #region End NetScaler NITRO Session
        #Disconnect from the NetScaler VPX
        $LogOut = @{"logout" = @{}} | ConvertTo-Json
        $dummy = Invoke-RestMethod -Uri "https://$NSIP/nitro/v1/config/logout" -Body $LogOut -Method POST -ContentType $ContentType -WebSession $NetScalerSession -Verbose:$VerbosePreference -ErrorAction SilentlyContinue
    #endregion End NetScaler NITRO Session
}

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = 400
$PSWindow.BufferSize = $WideDimensions

try {
    Get-NSContentSwitchBinding -NSIP $args[0] -NSUserName $args[1] -NSUserPW $args[2]
}
catch [System.Management.Automation.ParameterBindingException] {
    Write-Error "Couldn't bind parameter exception, Please make sure to provide all necessary parameters"
}
