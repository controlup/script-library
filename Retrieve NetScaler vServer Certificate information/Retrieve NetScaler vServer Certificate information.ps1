function Get-NSCertInfo {
    <#
    .SYNOPSIS
      Retrieve NetScaler vServer Certificate information.
    .DESCRIPTION
      Retrieve NetScaler vServer Certificate information, using the Invoke-RestMethod cmdlet for the REST API calls.
    .NOTES
      Version:        0.2
      Author:         Esther Barthel, MSc
      Creation Date:  2018-03-25
      Updated:        2018-06-03
                      Added binding information for SSL vServer, Service and Profile
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
      $certKeyName,

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
    Write-Output "----------------------------------------------------------- "
    Write-Output "| Retrieving Certificate information from the NetScaler:  | "
    Write-Output "----------------------------------------------------------- "
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


    # ----------------------
    # | CertKey statistics |
    # ----------------------
    #region Get CertKey information
        # Base URL 
        $strURI = "https://$NSIP/nitro/v1/config/sslcertkey"

        # Specify the required full URL, including filters and arguments
        $strArgs = ""
        If ($certKeyName)
        {
            # Add the certkey name to URI
            Write-Verbose ("Added the CertKey name """ + $certKeyName + """ to the URI")
            $strArgs = ("/" + $certKeyName)
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
                If ($response.sslcertkey)
                {
                    #$response.sslcertkey
                    Write-Output ""
                    Write-Host "* Certificate information:" -ForegroundColor Yellow
                    #$response.sslcertkey
                    $response.sslcertkey   | Select-Object @{N='CertKey Name'; E={$_.certkey}}, 
                                                        @{N=' Type'; E={"{0,5}" -F $_.inform}}, 
                                                        @{N='          Expiration date'; E={"{0,25}" -F $_.clientcertnotafter}}, 
                                                        @{N='Expiration days'; E={$_.daystoexpiration}}, 
                                                        @{N='    Status'; E={"{0,10}" -F $_.status}}, 
                                                        @{N='Cert link'; E={$_.linkcertkeyname}},
                                                        @{N='Expiry Monitor'; E={$_.expirymonitor}},
                                                        @{N='Notification (days)'; E={$_.notificationperiod}} | Sort-Object 'Days to expiration' | Format-Table -AutoSize 
                
                }
                Else
                {
                    Write-Warning "No Certificate information was found"
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


    #region Get CertKey SSL vServer Binding information
        # Base URL 
        $strURI = "https://$NSIP/nitro/v1/config/sslcertkey_sslvserver_binding"

        # Specify the required full URL, including filters and arguments
        $strArgs = ""
        If ($certKeyName)
        {
            # Add the certkey name to URI
            Write-Verbose ("Added the CertKey name """ + $certKeyName + """ to the URI")
            $strArgs = ("/" + $certKeyName)
        }
        Else
        {
            Write-Verbose ("Added the bulkbindings argument to the URI as no CertKey was specified")
            $strArgs = "?bulkbindings=yes"
        }
        $strURI = $strURI + $strArgs

        # Method #1: Making the REST API call to the NetScaler
        try
        {
            # start with clean response variable
            $response = $null
            $response = Invoke-RestMethod -Method Get -Uri $strURI -ContentType $ContentType -WebSession $NetScalerSession -Verbose:$VerbosePreference #-ErrorAction SilentlyContinue
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
                If ($response.sslcertkey_sslvserver_binding)
                {
                    Write-Output ""
                    Write-Host "* Certificate SSL vServer binding information:" -ForegroundColor Yellow
                    #$response.sslcertkey_sslvserver_binding
                    $response.sslcertkey_sslvserver_binding   | Select-Object @{N='CertKey Name'; E={$_.certkey}}, 
                                                        @{N='Priority'; E={"{0,8}" -F$_.data}},
                                                        @{N='vServer Name'; E={$_.servername}}, 
                                                        @{N='Version'; E={"{0,7}" -F $_.version}} | Sort-Object -Property 'vServer Name, Priority' | Format-Table -AutoSize 
    #>                
                }
                Else
                {
                    Write-Verbose "No Certificate SSL vServer binding information was found"
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

    #region Get CertKey SSL Service Binding information
        # Base URL 
        $strURI = "https://$NSIP/nitro/v1/config/sslcertkey_service_binding"

        # Specify the required full URL, including filters and arguments
        $strArgs = ""
        If ($certKeyName)
        {
            # Add the certkey name to URI
            Write-Verbose ("Added the CertKey name """ + $certKeyName + """ to the URI")
            $strArgs = ("/" + $certKeyName)
        }
        Else
        {
            Write-Verbose ("Added the bulkbindings argument to the URI as no CertKey was specified")
            $strArgs = "?bulkbindings=yes"
        }
        $strURI = $strURI + $strArgs

        # Method #1: Making the REST API call to the NetScaler
        try
        {
            # start with clean response variable
            $response = $null
            $response = Invoke-RestMethod -Method Get -Uri $strURI -ContentType $ContentType -WebSession $NetScalerSession -Verbose:$VerbosePreference #-ErrorAction SilentlyContinue
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
                If ($response.sslcertkey_service_binding)
                {
                    Write-Output ""
                    Write-Host "* Certificate SSL Service binding information:" -ForegroundColor Yellow
                    #$response.sslcertkey_service_binding
                    $response.sslcertkey_service_binding   | Select-Object @{N='CertKey Name'; E={$_.certkey}}, 
                                                        @{N='Priority'; E={"{0,8}" -F$_.data}},
                                                        @{N='Service Name'; E={$_.servicename}}, 
                                                        @{N='Version'; E={"{0,7}" -F $_.version}} | Sort-Object -Property 'vServer Name, Priority' | Format-Table -AutoSize 
    #>                
                }
                Else
                {
                    Write-Verbose "No Certificate SSL Service binding information was found"
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

    #region Get CertKey SSL Profile Binding information
        # Base URL 
        $strURI = "https://$NSIP/nitro/v1/config/sslcertkey_sslprofile_binding"

        # Specify the required full URL, including filters and arguments
        $strArgs = ""
        If ($certKeyName)
        {
            # Add the certkey name to URI
            Write-Verbose ("Added the CertKey name """ + $certKeyName + """ to the URI")
            $strArgs = ("/" + $certKeyName)
        }
        Else
        {
            Write-Verbose ("Added the bulkbindings argument to the URI as no CertKey was specified")
            $strArgs = "?bulkbindings=yes"
        }
        $strURI = $strURI + $strArgs

        # Method #1: Making the REST API call to the NetScaler
        try
        {
            # start with clean response variable
            $response = $null
            $response = Invoke-RestMethod -Method Get -Uri $strURI -ContentType $ContentType -WebSession $NetScalerSession -Verbose:$VerbosePreference #-ErrorAction SilentlyContinue
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

                If ($response.sslcertkey_sslprofile_binding)
                {
                    Write-Output ""
                    Write-Host "* Certificate SSL Profile binding information:" -ForegroundColor Yellow
                    #$response.sslcertkey_sslprofile_binding
                    $response.sslcertkey_sslprofile_binding   | Select-Object @{N='CertKey Name'; E={$_.certkey}}, 
                                                        @{N='SSL Profile'; E={"{0,7}" -F $_.sslprofile}} | Format-Table -AutoSize 
    #>                
                }
                Else
                {
                    Write-Verbose "No Certificate SSL Profile binding information was found"
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
    Get-NSCertInfo -NSIP $args[0] -NSUserName $args[1] -NSUserPW $args[2]
}
catch [System.Management.Automation.ParameterBindingException] {
    Write-Error "Couldn't bind parameter exception, Please make sure to provide all necessary parameters"
}
