function Get-NSConnectionInfo {
    <#
    .SYNOPSIS
      Retrieve NetScaler Connection information.
    .DESCRIPTION
      Retrieve NetScaler Connection information, using the Invoke-RestMethod cmdlet for the REST API calls.
    .NOTES
      Version:        0.3
      Author:         Esther Barthel, MSc
      Creation Date:  2018-05-20
      Updated:        2018-06-23
                      Adjusted params to work with the args[x] limitations of the SBA
      Updated:        2018-07-01
                      Added sorting for SourceIP and Svc type.
      Purpose:        SBA, created for ControlUp NetScaler Monitoring
      Based upon:     https://support.citrix.com/article/CTX126853

      Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param(
        # Declaring the input parameters, provided for the SBA
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        [string]
        $NSIP,

        [Parameter(Position=1, Mandatory=$false)]
        [string]
        $SourceIP,

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
    Write-Output "| Retrieving NS connectiontable information from the NetScaler: | "
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


    # --------------------------
    # | Connection information |
    # --------------------------
    #region Get NS Connection table information
        # Base URL 
        $strURI = "https://$NSIP/nitro/v1/config/nsconnectiontable"

        # Specify the required full URL, including filters and arguments
        $strArgs = ""
        If ($SourceIP)
        {
            # Add Source IP to URI
            Write-Verbose ("Added the Source IP """ + $SourceIP + """ to the URI")
            $strArgs = ("?args=filterexpression:" + [System.Web.HttpUtility]::UrlEncode("CONNECTION.SRCIP.EQ(" + $SourceIP + ")"))
        }

        If ($strArgs.Length -gt 0)
        {
            $strURI = $strURI + $strArgs + ",detail:FULL"
        }
        Else
        {
            $strURI = $strURI + "?args=detail:FULL"
        }

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
                If ($response.nsconnectiontable)
                {
                    $FilterExpression = ($response.nsconnectiontable | Select-Object filterexpression -Unique).filterexpression
                    Write-Host "Filter Expression: " -ForegroundColor Yellow -NoNewline
                    Write-Host $FilterExpression

                    #$response.nsconnectiontable | Format-List
                    #$response.nsconnectiontable | Select-Object sourceip, sourceport, destip, destport, svctype, idletime, state, entityname, httprequest | Format-Table -AutoSize
                    $response.nsconnectiontable | Select-Object @{N='      Source IP'; E={"{0,15}" -F $_.sourceip}}, 
                                                        @{N='S Port'; E={$_.sourceport}}, 
                                                        @{N=' Destination IP'; E={"{0,15}" -F $_.destip}}, 
                                                        @{N='D Port'; E={$_.destport}}, 
                                                        @{N='Svc type'; E={$_.svctype}}, 
                                                        @{N='Idle time'; E={$_.idletime}}, 
                                                        @{N='       State'; E={"{0,12}" -F $_.state}}, 
                                                        @{N='Entity name'; E={$_.entityname}}, 
                                                        @{N='HTTP request'; E={$_.httprequest}} | Sort-Object '      Source IP','Svc type'| Format-Table -AutoSize
                }
                Else
                {
                    Write-Warning "No connection information was found"
                    Write-Output ""
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
    Get-NSConnectionInfo -NSIP $args[0] -NSUserName $args[1] -NSUserPW $args[2]
}
catch [System.Management.Automation.ParameterBindingException] {
    Write-Error "Couldn't bind parameter exception, Please make sure to provide all necessary parameters"
}
