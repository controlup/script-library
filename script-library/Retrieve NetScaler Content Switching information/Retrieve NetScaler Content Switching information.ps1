function Get-NSContentSwitchInfo {
    <#
    .SYNOPSIS
      Retrieve NetScaler Content Switching information.
    .DESCRIPTION
      Retrieve NetScaler Content Switching information, using the Invoke-RestMethod cmdlet for the REST API calls.
    .NOTES
      Version:        0.1
      Author:         Esther Barthel, MSc
      Creation Date:  2018-05-22
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
    Write-Output "--------------------------------------------------------- "
    Write-Output "| Retrieving CS vServer information from the NetScaler: | "
    Write-Output "--------------------------------------------------------- "
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


    # -----------------
    # | CS statistics |
    # -----------------
    #region Get CS Config
        # Base URL 
        $strURI = "https://$NSIP/nitro/v1/config/csvserver"

        # Specify the required full URL, including filters and arguments
        $strArgs = ""
        If ($CSvServer)
        {
            # Add Profile name to URI
            Write-Verbose ("Added the CS vServer """ + $CSvServer + """ to the URI")
            $strArgs = ("/" + $CSvServer)
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
                If ($response.csvserver)
                {
                    Write-Host "CS vServer configuration: " -ForegroundColor Yellow
                    #$response.csvserver | Format-List
                    $response.csvserver | Select-Object @{N='Name'; E={$_.name}}, 
                                                            @{N='     IP-address'; E={$_.ipv46}}, 
                                                            @{N='Port'; E={$_.port}}, 
                                                            @{N='Protocol'; E={$_.servicetype}}, 
                                                            @{N='Type'; E={$_.type}}, 
                                                            @{N='State'; E={$_.curstate}}, 
                                                            @{N='ICMP response'; E={$_.icmpvsrresponse}}, 
                                                            @{N='HTTP profile'; E={$_.httpprofilename}}, 
                                                            @{N='Traffic domain'; E={$_.td}} | Sort-Object 'Name' | Format-Table -AutoSize
                }
                Else
                {
                    Write-Warning "No CS vServer information was found"
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
    Get-NSContentSwitchInfo -NSIP $args[0] -NSUserName $args[1] -NSUserPW $args[2]
}
catch [System.Management.Automation.ParameterBindingException] {
    Write-Error "Couldn't bind parameter exception, Please make sure to provide all necessary parameters"
}
