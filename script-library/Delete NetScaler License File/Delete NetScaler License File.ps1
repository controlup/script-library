function Remove-NSLicenseFile {
    <#
    .SYNOPSIS
      Delete local NetScaler License File from the appliance.
    .DESCRIPTION
      Delete local NetScaler License File from the appliance, using the Invoke-RestMethod cmdlet for the REST API calls.
    .NOTES
      Version:        0.2
      Author:         Esther Barthel, MSc
      Creation Date:  2018-02-25
      Updated:        2018-04-08
                      Improving error handling
      Purpose:        SBA - Created for ControlUp NetScaler Monitoring

      Copyright (c) cognition IT. All rights reserved.
    #>

    [CmdletBinding()]
    Param(
      # Declaring the input parameters, provided for the SBA
      [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] [string] $NSIP="192.168.0.132",
      [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] [string] $LicenseFile, 
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
    #endregion NITRO settings

    Clear-Host

    Write-Output "-------------------------------------------------------------- " #-ForegroundColor Yellow
    Write-Output "| Delete a local license file from the NetScaler appliance:  | " #-ForegroundColor Yellow
    Write-Output "-------------------------------------------------------------- " #-ForegroundColor Yellow
    Write-Output ""

    #Pre check on License File
    if ([string]::IsNullOrEmpty($LicenseFile))
    {
        Write-Warning "No license file specified! Action canceled."
        Write-Output ""
        Exit
    }

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


    # ---------------------------------------
    # | Delete local NetScaler License file |
    # ---------------------------------------

    #region delete local NetScaler license file
        # Specifying the correct URL 
        $strURI = ("https://$NSIP/nitro/v1/config/systemfile/$LicenseFile" + "?args=filelocation:" + [System.Web.HttpUtility]::UrlEncode("/nsconfig/license"))

        # based on the Gallery script YesNoPrompt.ps1, submitted by Kent Finkle (source: https://gallery.technet.microsoft.com/scriptcenter/1a386b01-b1b8-4ac2-926c-a4986ac94fed)
        $RUSure = new-object -comobject wscript.shell 
        $intAnswer = $RUSure.popup("Do you really want to delete the ""$LicenseFile"" license file?", 0,"Delete Files",4) 
        If ($intAnswer -eq 6) 
        { 
            # Method #1: Making the REST API call to the NetScaler
            try
            {
                $response = $null
                    # Make the REST API call to delete the license file
                    $response = Invoke-RestMethod -Method Delete -Uri $strURI -ContentType $ContentType -WebSession $NetScalerSession -Verbose:$VerbosePreference -ErrorAction SilentlyContinue
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
                    Write-Verbose "REST API call to delete license file: Successful"
                    Write-Output ""
                    Write-Output "License file succesfully deleted."
                }
                Else
                {
                    Write-Error ("An error occured! Errorcode: " + $response.errorcode + ", message: " + $response.message)
                }
            }
        }
        Else
        { 
            Write-Output ""
            Write-Warning "The Delete action was canceled."
            Exit
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
    Remove-NSLicenseFile -NSIP $args[0] -NSUserName $args[1] -NSUserPW $args[2] -LicenseFile $args[3]
}
catch [System.Management.Automation.ParameterBindingException] {
    Write-Error "Couldn't bind parameter exception, Please make sure to provide all necessary parameters"
}
