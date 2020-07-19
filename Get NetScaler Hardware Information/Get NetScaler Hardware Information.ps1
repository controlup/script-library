function Get-NsHwInfo {
    <#
    .SYNOPSIS
      Retrieve NetScaler Serial Number and Mac Address.
    .DESCRIPTION
      Retrieve NetScaler Serial Number and Mac Address and other desired HW info
      using the Invoke-RestMethod cmdlet for the REST API calls.
    .CREDIT / BASED ON THE WORK OF  Esther Barthel, MSc - cognition IT
                                    Retrieve NetScaler SSL stats
      Version:        0.2
      Author:         Marcel Calef
      Creation Date:  2018-11-30
      Updated:        2018-11-30
      Reviewed by:    Esther Barthel, MSc
      Updated:        2019-06-20 by Esther Barthel, MSc
                      Alligned the Hardware settings output with the NetScaler information, added the Credentials popup, replaced Write-Host with Write-Output.
    #>

    [CmdletBinding()]
    Param(
      # Declaring the input parameters, provided for the SBA
      [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] [string] $NSIP,
  [Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)] [System.Management.Automation.CredentialAttribute()] $NSCredentials
     )    

    #region NITRO settings
        # NITRO Constants
        $ContentType = "application/json"
    #endregion

    If ($NSCredentials -eq $null)
    {
        #NS Credentials
        $NSCredentials = Get-Credential -Message "Enter your NetScaler Credentials for $NSIP"
    }

    # Retieving Username and Password from the Credentials to use with NITRO
    $NSUserName = $NSCredentials.UserName
    $NSUserPW = $NSCredentials.GetNetworkCredential().Password


    # Ensure Debug information is shown, without the confirmation question after each Write-Debug
    If ($PSBoundParameters['Debug']) {$DebugPreference = 'Continue'}

    Write-Output "---------------------------------------------------------------- "
    Write-Output "| Retrieving NetScaler Serial Number and MAC Address (Host Id) | "
    Write-Output "---------------------------------------------------------------- "
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
            $loginresponse = Invoke-RestMethod -Uri "https://$NSIP/nitro/v1/config/login" -Body $Login -Method POST `
            -SessionVariable NetScalerSession -ContentType $ContentType -Verbose:$VerbosePreference -ErrorAction SilentlyContinue
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


     #region Get Hardware Information
        # Base URL 
        $strURI = "https://$NSIP/nitro/v1/config/nshardware"

        # Method #1: Making the REST API call to the NetScaler
        try
        {
            # start with clean response variable
            $response = $null
            $response = Invoke-RestMethod -Method Get -Uri $strURI -ContentType $ContentType `
                     -WebSession $NetScalerSession -Verbose:$VerbosePreference -ErrorAction SilentlyContinue
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
            Write-Verbose "REST API call to retrieve information: successful"
			If ($response.errorcode -eq 0)
            {
                Write-Verbose "REST API call to retrieve information: successful"

                If ($response.nshardware)
                {
                    
                    Write-Output ""
                    Write-Output "* Hardware information:"
                    $response.nshardware | Select-Object @{N='Platform'; E={(($_.hwdescription).ToString() + " " + ($_.sysid).ToString())}}, 
                                                        @{N='Manufactured on'; E={( ($_.manufacturemonth).ToString() + "/" + ($_.manufactureday).ToString() + "/" + ($_.manufactureyear).ToString() )}}, 
                                                        @{N='CPU'; E={( ($_.cpufrequncy).ToString() + " MHZ" )}}, 
                                                        @{N='Host Id'; E={$_.host}}, 
                                                        @{N='Serial no'; E={$_.serialno}}, 
                                                        @{N='Encoded serial no'; E={$_.encodedserialno}}  | Format-List
                }
                Else
                {
                    Write-Warning "No Hardware information was found"
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
        $dummy = Invoke-RestMethod -Uri "https://$NSIP/nitro/v1/config/logout" -Body $LogOut -Method POST -ContentType $ContentType `
                     -WebSession $NetScalerSession -Verbose:$VerbosePreference -ErrorAction SilentlyContinue
    #endregion End NetScaler NITRO Session
}
# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = 400
$PSWindow.BufferSize = $WideDimensions

#NS Credentials
$NSCredentials = Get-Credential -Message ("Enter your NetScaler Credentials for " + $args[0])

try {
    Get-NsHwInfo -NSIP $args[0] -NSCredentials $NSCredentials
}
catch [System.Management.Automation.ParameterBindingException] {
    Write-Error "Couldn't bind parameter exception, Please make sure to provide all necessary parameters"
}
catch
{
    Write-Error ("An error (" + $_.Exception.GetType().FullName + ") occurred, with message: " + $_.Exception.Message)
}

