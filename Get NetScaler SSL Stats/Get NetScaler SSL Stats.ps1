function Get-NSSSLStats {
    <#
    .SYNOPSIS
      Retrieve SSL statistical information.
    .DESCRIPTION
      Retrieve SSL statistical information, using the Invoke-RestMethod cmdlet for the REST API calls.
    .NOTES
      Version:        0.2
      Author:         Esther Barthel, MSc
      Creation Date:  2018-03-26
      Updated:        2018-04-08
                      Finalizing script
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

    Write-Output "-------------------------------------------------------------- "
    Write-Output "| Retrieving SSL statistical information from the NetScaler: | "
    Write-Output "-------------------------------------------------------------- "
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
    # | SSL statistics |
    # ------------------
    #region Get SSL Stats
        # Base URL 
        $strURI = "https://$NSIP/nitro/v1/stat/ssl"

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
                #$response.ssl

                If ($response.ssl)
                {
                    Write-Output ""
                    Write-Host "* SSL hardware statistics:" -ForegroundColor Yellow
                    $response.ssl | Select-Object @{N='# Cards present'; E={$_.sslcards}}, 
                                                        @{N='# Cards UP'; E={$_.sslnumcardsup}}, 
                                                        @{N='# Secondary cards present'; E={$_.sslcardssecondary}}, 
                                                        @{N='# Secondary cards UP'; E={$_.sslnumcardsupsecondary}}, 
                                                        @{N='Cards status'; E={$_.sslcardstatus}}, 
                                                        @{N='Engine status'; E={@(if($_.sslenginestatus -eq 1){"UP"}else{"DOWN"})}},                    # 1 is UP, 0 is DOWN
                                                        @{N='Total SSL sessions'; E={$_.ssltotsessions}}, 
                                                        @{N='SSL sessions (Rate)'; E={$_.sslsessionsrate}}, 
                                                        @{N='SSL crypto utilization (%)'; E={"{0:N2}" -F($_.sslcryptoutilizationstat)}}, 
                                                        @{N='Second card utilization (%)'; E={"{0:N2}" -F($_.sslcryptoutilizationstat2nd)}} | Format-List

                    Write-Output ""
                    Write-Host "* SSL system transactions statistics - Total:" -ForegroundColor Yellow
                    $response.ssl | Select-Object @{N='       SSL'; E={"{0,10}"-F $_.ssltottransactions}}, 
                                                        @{N='     SSLv2'; E={"{0,10}"-F $_.ssltotsslv2transactions}}, 
                                                        @{N='     SSLv3'; E={"{0,10}"-F $_.ssltotsslv3transactions}}, 
                                                        @{N='     TLSv1'; E={"{0,10}"-F $_.ssltottlsv1transactions}}, 
                                                        @{N='   TLSv1.1'; E={"{0,10}"-F $_.ssltottlsv11transactions}}, 
                                                        @{N='   TLSv1.2'; E={"{0,10}"-F $_.ssltottlsv12transactions}} | Format-Table -AutoSize

                    Write-Host "* SSL system transactions statistics - Rate (/s):" -ForegroundColor Yellow
                    $response.ssl | Select-Object @{N='       SSL'; E={"{0,10}"-F $_.ssltransactionsrate}}, 
                                                        @{N='     SSLv2'; E={"{0,10}"-F $_.sslsslv2transactionsrate}}, 
                                                        @{N='     SSLv3'; E={"{0,10}"-F $_.sslsslv3transactionsrate}}, 
                                                        @{N='     TLSv1'; E={"{0,10}"-F $_.ssltlsv1transactionsrate}}, 
                                                        @{N='   TLSv1.1'; E={"{0,10}"-F $_.ssltlsv11transactionsrate}}, 
                                                        @{N='   TLSv1.2'; E={"{0,10}"-F $_.ssltlsv12transactionsrate}} | Format-Table -AutoSize

                    Write-Output ""
                    Write-Host "* SSL Front End sessions statistics - Total:" -ForegroundColor Yellow
                    $response.ssl | Select-Object @{N='       SSL'; E={"{0,10}"-F $_.ssltotsessions}}, 
                                                        @{N='     SSLv2'; E={"{0,10}"-F $_.ssltotsslv2sessions}}, 
                                                        @{N='     SSLv3'; E={"{0,10}"-F $_.ssltotsslv3sessions}}, 
                                                        @{N='     TLSv1'; E={"{0,10}"-F $_.ssltottlsv1sessions}}, 
                                                        @{N='   TLSv1.1'; E={"{0,10}"-F $_.ssltottlsv11sessions}}, 
                                                        @{N='   TLSv1.2'; E={"{0,10}"-F $_.ssltottlsv12sessions}} | Format-Table -AutoSize

                    Write-Host "* SSL Front End sessions statistics - Rate (/s):" -ForegroundColor Yellow
                    $response.ssl | Select-Object @{N='       SSL'; E={"{0,10}"-F $_.sslsessionsrate}}, 
                                                        @{N='     SSLv2'; E={"{0,10}"-F $_.sslsslv2sessionsrate}}, 
                                                        @{N='     SSLv3'; E={"{0,10}"-F $_.sslsslv3sessionsrate}}, 
                                                        @{N='     TLSv1'; E={"{0,10}"-F $_.ssltlsv1sessionsrate}}, 
                                                        @{N='   TLSv1.1'; E={"{0,10}"-F $_.ssltlsv11sessionsrate}}, 
                                                        @{N='   TLSv1.2'; E={"{0,10}"-F $_.ssltlsv12sessionsrate}} | Format-Table -AutoSize

                    Write-Output ""
                    Write-Host "* SSL Back End sessions statistics - Total:" -ForegroundColor Yellow
                    $response.ssl | Select-Object @{N='    SSL'; E={"{0,7}"-F $_.sslbetotsessions}}, 
                                                        @{N='  SSLv3'; E={"{0,7}"-F $_.sslbetotsslv3sessions}}, 
                                                        @{N='  TLSv1'; E={"{0,7}"-F $_.sslbetottlsv1sessions}}, 
                                                        @{N='TLSv1.1'; E={"{0,7}"-F $_.sslbetottlsv11sessions}}, 
                                                        @{N='TLSv1.2'; E={"{0,7}"-F $_.sslbetottlsv12sessions}}, 
                                                        @{N='Multiplex attempts'; E={"{0,18}"-F $_.sslbetotsessionmultiplexattempts}}, 
                                                        @{N='Multiplex successes'; E={"{0,19}"-F $_.sslbetotsessionmultiplexattemptsuccess}}, 
                                                        @{N='Multiplex failures'; E={"{0,18}"-F $_.sslbetotsessionmultiplexattemptfails}} | Format-Table -AutoSize

                    Write-Host "* SSL Back End sessions statistics - Rate (/s):" -ForegroundColor Yellow
                    $response.ssl | Select-Object @{N='    SSL'; E={"{0,7}"-F $_.sslbesessionsrate}}, 
                                                        @{N='  SSLv3'; E={"{0,7}"-F $_.sslbesslv3sessionsrate}}, 
                                                        @{N='  TLSv1'; E={"{0,7}"-F $_.sslbetlsv1sessionsrate}}, 
                                                        @{N='TLSv1.1'; E={"{0,7}"-F $_.sslbetlsv11sessionsrate}}, 
                                                        @{N='TLSv1.2'; E={"{0,7}"-F $_.sslbetlsv12sessionsrate}}, 
                                                        @{N='Multiplex attempts'; E={"{0,18}"-F $_.sslbesessionmultiplexattemptsrate}}, 
                                                        @{N='Multiplex successes'; E={"{0,19}"-F $_.sslbesessionmultiplexattemptsuccessrate}}, 
                                                        @{N='Multiplex failures'; E={"{0,18}"-F $_.sslbesessionmultiplexattemptfailsrate}} | Format-Table -AutoSize

                    Write-Output ""
                    Write-Host "* SSL Encyption/Decryption statistics:" -ForegroundColor Yellow
                    $response.ssl | Select-Object @{N='Total bytes encrypted'; E={"{0,21}"-F $_.ssltotenc}}, 
                                                        @{N='Total bytes decrypted'; E={"{0,21}"-F $_.ssltotdec}}, 
                                                        @{N='Bytes encrypted - Rate (/s)'; E={"{0,27}"-F $_.sslencrate}}, 
                                                        @{N='Bytes decrypted - Rate (/s)'; E={"{0,27}"-F $_.ssldecrate}} | Format-Table -AutoSize
                
                    ##debug info
                    #$response.ssl
                }
                Else
                {
                    Write-Warning "No SSL information was found"
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
    Get-NSSSLStats -NSIP $args[0] -NSUserName $args[1] -NSUserPW $args[2]
}
catch [System.Management.Automation.ParameterBindingException] {
    Write-Error "Couldn't bind parameter exception, Please make sure to provide all necessary parameters"
}
