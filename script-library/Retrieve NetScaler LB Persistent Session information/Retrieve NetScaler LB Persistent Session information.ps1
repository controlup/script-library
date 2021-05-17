function Get-NSLBPersistentSessionInfo {
    <#
    .SYNOPSIS
      Retrieve NetScaler LB Persistent Session information.
    .DESCRIPTION
      Retrieve NetScaler LB Persistent Session information, using the Invoke-RestMethod cmdlet for the REST API calls.
    .NOTES
      Version:        0.5
      Author:         Esther Barthel, MSc
      Creation Date:  2018-02-04
      Updated:        2018-06-23
                      Split up script due to SBA limitations for input variables
      Updated:        2018-07-01
                      Added sorting
    Purpose:        SBA - Created for ControlUp NetScaler Monitoring

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
        $vServer,

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
        $GetGeoLocation = $false
    #endregion

    Write-Output ""
    Write-Output "--------------------------------------------------------------- " #-ForegroundColor Yellow
    Write-Output "| Retrieving vServer persistence sessions from the NetScaler: | " #-ForegroundColor Yellow
    Write-Output "--------------------------------------------------------------- " #-ForegroundColor Yellow
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

    # -------------------------------------
    # | LB Persistence Sessions statistics |
    # -------------------------------------
    #region Get LB Persistent Session Stats
        # Specifying the correct URL 
        $strURI = "https://$NSIP/nitro/v1/config/lbpersistentsessions"

    #    Specify the correct URL, using filters with the args argument
        If ($vServer)
        {
            $strArgs = "?args=vserver:" + [System.Web.HttpUtility]::UrlEncode($vServer) 
            $strURI = $strURI + $strArgs
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
            Write-Error ("A " + $_.Exception.GetType().FullName + " error occurred, with message: " + $_.Exception.Message)
            Write-Verbose "Error full details: "
            If ($VerbosePreference -eq "Continue")
            {
                echo $_.Exception | Format-List -Force
            }
        }
        Finally
        {
            If ($response.errorcode -eq 0)
            {
                Write-Verbose "REST API call to retrieve stats: successful"
                Write-Output ""

                If ($response.lbpersistentsessions)
                {
                    Write-Host "NetScaler LB Persistence Sessions information:" -ForegroundColor Yellow
                    #$response.lbpersistentsessions
                    $response.lbpersistentsessions | Select-Object @{N='vServer'; E={$_.vserver}}, 
                                                        @{N='Type'; E={$_.typestring}}, 
                                                        @{N='Source IP'; E={$_.srcip}}, 
                                                        @{N='Destination IP'; E={$_.destip}}, 
                                                        @{N='Dest. port'; E={$_.destport}}, 
                                                        @{N='Timeout (sec)'; E={$_.timeout}}, 
                                                        @{N='Persistence parameters'; E={$_.persistenceparam}} | Sort-Object 'vServer','Persistence parameters' -Descending  | Format-Table -AutoSize

                }
                Else
                {
                    If ($response.errorcode -eq 0)
                    {
                        Write-Warning "No LB Persistence Sessions found."
                        Write-Output ""
                    }
                    Else
                    {
                        Write-Error ("errorcode: " + $response.errorcode + ", message: " + $response.message)
                    }
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
    Get-NSLBPersistentSessionInfo -NSIP $args[0] -NSUserName $args[1] -NSUserPW $args[2]
}
catch [System.Management.Automation.ParameterBindingException] {
    Write-Error "Couldn't bind parameter exception, Please make sure to provide all necessary parameters"
}
