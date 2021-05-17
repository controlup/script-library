function Get-NSSSLBindingInfo {
    <#
    .SYNOPSIS
      Retrieve NetScaler SSL vServer Bindings information.
    .DESCRIPTION
      Retrieve NetScaler SSL vServer Binding information, using the Invoke-RestMethod cmdlet for the REST API calls.
    .NOTES
      Version:        0.2
      Author:         Esther Barthel, MSc
      Creation Date:  2018-02-19
      Updated:        2018-06-23
                      Updated binding information retrieved
                      Added SSL Profile binding
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
        $vServerName,

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

    Write-Output "------------------------------------------------------------------- "
    Write-Output "| Retrieving SSL vServer binding information from the NetScaler:  | "
    Write-Output "------------------------------------------------------------------- "
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


    # -----------------------------------
    # | SSL vServer Bindings statistics |
    # -----------------------------------
    #region Get SSL vServer Bindings
        # Base URL 
        $strURI = "https://$NSIP/nitro/v1/config/sslvserver_binding"

        # Specify the required full URL, including filters and arguments
        $strArgs = ""
        If ($vServerName)
        {
            # Add vServer name to URI
            Write-Verbose ("Added the vServer name """ + $vServerName + """ to the URI")
            $strArgs = ("/" + $vServerName)
        }
        Else
        {
            Write-Verbose ("Added the bulkbindings argument to the URI as no vServer was specified")
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
                If ($response.sslvserver_binding)
                {
                    #$response.sslvserver_binding | Format-List

                    If (!($response.sslvserver_binding.sslvserver_sslpolicy_binding.vservername -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer SSL policy bindings:" -ForegroundColor Yellow
                        #$response.sslvserver_binding.sslvserver_sslpolicy_binding | Format-List
                        $response.sslvserver_binding.sslvserver_sslpolicy_binding | Select-Object @{N='VS name'; E={$_.vservername}}, 
                                                            @{N='Policy name'; E={$_.policyname}}, 
                                                            @{N='Priority'; E={$_.priority}}, 
                                                            @{N='Type'; E={$_.type}}, 
                                                            @{N='Policy inherit'; E={$_.polinherit}}, 
                                                            @{N='GoToPriority Expression'; E={$_.gotopriorityexpression}}, 
                                                            @{N='Invoke'; E={$_.invoke}}, 
                                                            @{N='Label type'; E={$_.labeltype}}, 
                                                            @{N='Label name'; E={$_.labelname}} | Format-Table -AutoSize
                    }
    #>                
                    If (!($response.sslvserver_binding.sslvserver_sslcertkey_binding.vservername -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer SSL certkey bindings:" -ForegroundColor Yellow
        #                $response.sslvserver_binding.sslvserver_sslcertkey_binding | Format-List
                        $response.sslvserver_binding.sslvserver_sslcertkey_binding | Select-Object @{N='VS name'; E={$_.vservername}}, 
                                                            @{N='Certificate KeyPair name'; E={$_.certkeyname}}, 
                                                            #@{N='CRL check'; E={$_.crlcheck}}, 
                                                            #@{N='OCSP check'; E={$_.ocspcheck}}, 
                                                            #@{N='Cleartext port'; E={$_.cleartextport}}, 
                                                            @{N='CA certificate'; E={$_.ca}}, 
                                                            @{N='SNI certificate'; E={$_.snicert}}, 
                                                            @{N='Skip CA name'; E={$_.skipcaname}} | Format-Table -AutoSize
                    }
    #>                
                    If (!($response.sslvserver_binding.sslvserver_sslciphersuite_binding.vservername -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer SSL Cipher suite bindings:" -ForegroundColor Yellow
                        #$response.sslvserver_binding.sslvserver_sslciphersuite_binding | Format-List
                        $response.sslvserver_binding.sslvserver_sslciphersuite_binding | Select-Object @{N='VS name'; E={$_.vservername}}, 
                                                            @{N='Cipher name'; E={$_.ciphername}}, 
                                                            @{N='Description'; E={$_.description}} | Format-Table -AutoSize
                    }
                
                    If (!($response.sslvserver_binding.sslvserver_ecccurve_binding.vservername -eq $null))
                    {
                        Write-Output ""
                        Write-Host "* vServer SSL ECC Curve bindings:" -ForegroundColor Yellow
                        #$response.sslvserver_binding.sslvserver_ecccurve_binding | Format-List
                        $response.sslvserver_binding.sslvserver_ecccurve_binding | Select-Object @{N='VS name'; E={$_.vservername}}, 
                                                            @{N='ECC curve name'; E={$_.ecccurvename}} | Format-Table -AutoSize
                    }
                }
                Else
                {
                    Write-Warning "No vServer binding information was found"
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
    Get-NSSSLBindingInfo -NSIP $args[0] -NSUserName $args[1] -NSUserPW $args[2]
}
catch [System.Management.Automation.ParameterBindingException] {
    Write-Error "Couldn't bind parameter exception, Please make sure to provide all necessary parameters"
}

