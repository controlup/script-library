function Get-NSStats {

    <#
    .SYNOPSIS
      Retrieve NS statistical information.
    .DESCRIPTION
      Retrieve NS statistical information, using the Invoke-RestMethod cmdlet for the REST API calls.
    .NOTES
      Version:        0.3
      Author:         Esther Barthel, MSc
      Creation Date:  2018-03-26
      Updated:        2018-04-08
                      Added percentage presentation
      Purpose:        SBA, created for ControlUp NetScaler Monitoring

      Copyright (c) cognition IT. All rights reserved.
    #>

    [CmdletBinding()]
    Param(
      # Declaring the input parameters, provided for the SBA
      [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] [string] $NSIP,
      [Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$true)]
      [string]
      $NSUserName,
      [Parameter(Position=2, Mandatory=$true, ValueFromPipeline=$true)]
      [string]
      $NSUserPW
     )    

    #region NITRO settings
        # NITRO Constants
        $ContentType = "application/json"
    #endregion

    # Ensure Debug information is shown, without the confirmation question after each Write-Debug
    If ($PSBoundParameters['Debug']) {$DebugPreference = 'Continue'}

    Write-Output "---------------------------------------------------- "
    Write-Output "| Retrieving System statistics from the NetScaler: | "
    Write-Output "---------------------------------------------------- "
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


    # ------------------
    # | NS statistics |
    # ------------------
    #region Get NS Stats
        # Base URL 
        $strURI = "https://$NSIP/nitro/v1/stat/ns"

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
                Write-Verbose "REST API call to retrieve information: Successful"

                If ($response.ns)
                {
                    Write-Output ""
                    Write-Host "* System overview:" -ForegroundColor Yellow
                    ## debug info
                    #$response.ns | Select-Object pktcpuusagepcnt, mgmtcpuusagepcnt, memuseinmb, memusagepcnt
                    $response.ns | Select-Object @{N='Up Since'; E={$_.starttime}}, 
                                                        @{N='# CPUs'; E={$_.numcpus}}, 
                                                        @{N='# SSL cards'; E={$_.sslcards}}, 
                                                        @{N='# SSL cards UP'; E={$_.sslnumcardsup}}, 
                                                        @{N='Packet CPU usage (%)'; E={"{0:N2}" -F ($_.pktcpuusagepcnt)}}, 
                                                        @{N='Management CPU usage (%)'; E={"{0:N2}" -F ($_.mgmtcpuusagepcnt)}}, 
                                                        @{N='Memory usage (MB)'; E={$_.memuseinmb}}, 
                                                        @{N='InUse Memory (%)'; E={"{0:N2}" -F ($_.memusagepcnt)}} | Format-List

                    Write-Output ""
                    Write-Host "* System Disks:" -ForegroundColor Yellow
                    $response.ns | Select-Object @{N='/flash used (%)'; E={"{0,15:N2}" -F ($_.disk0perusage)}}, 
                                                        @{N='/flash available (MB)'; E={$_.disk0avail}}, 
                                                        @{N='/var used (%)'; E={"{0,13:N2}" -F ($_.disk1perusage)}}, 
                                                        @{N='/var available (MB)'; E={$_.disk1avail}} | Format-Table -AutoSize

                    Write-Output ""
                    Write-Host "* Throughput Statistics:" -ForegroundColor Yellow
                    #$response.ns
                    $response.ns | Select-Object @{N='Total received (MB)'; E={"{0,19}" -F $_.totrxmbits}}, 
                                                        @{N='Total transmitted (MB)'; E={"{0,22}" -F $_.tottxmbits}}, 
                                                        @{N='Received - Rate (/s)'; E={"{0,20}" -F $_.rxmbitsrate}}, 
                                                        @{N='Transmitted - Rate (/s)'; E={"{0,23}" -F $_.txmbitsrate}} | Format-Table -AutoSize

                    Write-Output ""
                    Write-Host "* TCP Connections:" -ForegroundColor Yellow
                    $response.ns | Select-Object @{N='All client conn.'; E={"{0,16}" -F $_.tcpcurclientconn}}, 
                                                        @{N='Established client conn.'; E={"{0,24}" -F $_.tcpcurclientconnestablished}}, 
                                                        @{N='All server conn.'; E={"{0,16}" -F $_.tcpcurserverconn}}, 
                                                        @{N='Established server conn.'; E={"{0,24}" -F $_.tcpcurserverconnestablished}} | Format-Table -AutoSize

                    Write-Output ""
                    Write-Host "* HTTP - Total:" -ForegroundColor Yellow
                    $response.ns | Select-Object @{N='Total requests'; E={"{0,14}" -F $_.httptotrequests}}, 
                                                        @{N='Total responses'; E={"{0,15}" -F $_.httptotresponses}}, 
                                                        @{N='Request bytes received'; E={"{0,22}" -F $_.httptotrxrequestbytes}}, 
                                                        @{N='Response bytes received'; E={"{0,23}" -F $_.httptotrxresponsebytes}} | Format-Table -AutoSize

                    Write-Host "* HTTP - Rate (/s):" -ForegroundColor Yellow
                    $response.ns | Select-Object @{N='Total requests'; E={"{0,14}" -F $_.httprequestsrate}}, 
                                                        @{N='Total responses'; E={"{0,15}" -F $_.httpresponsesrate}}, 
                                                        @{N='Request bytes received'; E={"{0,22}" -F $_.httprxrequestbytesrate}}, 
                                                        @{N='Response bytes received'; E={"{0,23}" -F $_.httprxresponsebytesrate}} | Format-Table -AutoSize

                    Write-Output ""
                    Write-Host "* SSL:" -ForegroundColor Yellow
                    $response.ns | Select-Object @{N='Total Transactions'; E={"{0,18}" -F $_.ssltottransactions}}, 
                                                        @{N='Total Session hits'; E={"{0,18}" -F $_.ssltotsessionhits}}, 
                                                        @{N='Transactions - Rate (/s)'; E={"{0,24}" -F $_.ssltransactionsrate}}, 
                                                        @{N='Session hits - Rate (/s)'; E={"{0,24}" -F $_.sslsessionhitsrate}} | Format-Table -AutoSize

                    Write-Output ""
                    Write-Host "* Integrated Caching - Total:" -ForegroundColor Yellow
                    $response.ns | Select-Object @{N='Hits'; E={"{0,4}" -F $_.cachetothits}}, 
                                                        @{N='Misses'; E={"{0,6}" -F $_.cachetotmisses}}, 
                                                        @{N='Origin bandwidth saved (%)'; E={"{0,26:N2}" -F ($_.cachepercentoriginbandwidthsaved)}}, 
                                                        @{N='Max memory (KB)'; E={"{0,14}" -F $_.cache64maxmemorykb}}, 
                                                        @{N='Max memory active value (KB)'; E={"{0,28}" -F $_.cachemaxmemoryactivekb}}, 
                                                        @{N='Utilized memory (KB)'; E={"{0,20}" -F $_.cacheutilizedmemorykb}} | Format-Table -AutoSize

                    Write-Host "* Integrated Caching - Rate (/s):" -ForegroundColor Yellow
                    $response.ns | Select-Object @{N='Hits'; E={"{0,4}" -F $_.cachetothits}}, 
                                                        @{N='Misses'; E={"{0,6}" -F $_.cachetotmisses}} | Format-Table -AutoSize

                    Write-Output ""
                    Write-Host "* Application Firewall - Total:" -ForegroundColor Yellow
                    $response.ns | Select-Object @{N='Requests'; E={"{0,8}" -F $_.appfirewallrequests}}, 
                                                        @{N='Responses'; E={"{0,9}" -F $_.appfirewallresponses}}, 
                                                        @{N='Aborts'; E={"{0,6}" -F $_.appfirewallaborts}}, 
                                                        @{N='Redirects'; E={"{0,9}" -F $_.appfirewallredirects}} | Format-Table -AutoSize
            
                    Write-Host "* Application Firewall - Rate (/s):" -ForegroundColor Yellow
                    $response.ns | Select-Object @{N='Requests'; E={"{0,8}" -F $_.appfirewallrequestsrate}}, 
                                                        @{N='Responses'; E={"{0,9}" -F $_.appfirewallresponsesrate}}, 
                                                        @{N='Aborts'; E={"{0,6}" -F $_.appfirewallabortsrate}}, 
                                                        @{N='Redirects'; E={"{0,9}" -F $_.appfirewallredirectsrate}} | Format-Table -AutoSize
                    ##debug info
                    #$response.ns
                }
                Else
                {
                    Write-Warning "No NetScaler System information was found"
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
    Get-NSStats -NSIP $args[0] -NSUserName $args[1] -NSUserPW $args[2]
}
catch [System.Management.Automation.ParameterBindingException] {
    Write-Error "Couldn't bind parameter exception, Please make sure to provide all necessary parameters"
}
