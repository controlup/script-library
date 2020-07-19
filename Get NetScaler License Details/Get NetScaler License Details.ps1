function Get-NSLicenseDetails {
    <#
    .SYNOPSIS
      Retrieve NetScaler detailed License file information.
    .DESCRIPTION
      Retrieve NetScaler detailed License file informatio, using the Invoke-RestMethod cmdlet for the REST API calls.
    .NOTES
      Version:        0.1
      Author:         Esther Barthel, MSc
      Creation Date:  2018-02-22
      Purpose:        SBA - Created for ControlUp NetScaler Monitoring
      Credits:        A big shoutout to Ryan Butler (Citrix CTA) for showing me how to retrieve data from the license files of the NetScaler

      Copyright (c) cognition IT. All rights reserved.
      Copyright (c) ControlUp All rights reserved.
    #>

    [CmdletBinding()]
    Param(
      # Declaring the input parameters, provided for the SBA
      [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
      [string]
      $NSIP,
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

    #region Credentials
    #    $NSCreds = Get-Credential
    #    $NSUserName = $NSCreds.UserName
    #    $NSUserPW = $NSCreds.GetNetworkCredential().Password
    #endregion Credentials

    Write-Output "------------------------------------------------------------ " #-ForegroundColor Yellow
    Write-Output "| Retrieving License files information from the NetScaler: | " #-ForegroundColor Yellow
    Write-Output "------------------------------------------------------------ " #-ForegroundColor Yellow
    Write-Output ""

    # ----------------------------------------
    # | Method #1: Using the SessionVariable |
    # ----------------------------------------
    #region Start NetScaler NITRO Session
        #Force PowerShell to bypass the CRL check for certificates and SSL connections
            Write-Verbose "Forcing PowerShell to trust all certificates (including the self-signed NetScaler certificate)"
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

    # ----------------------------------------
    # | Getting NetScaler Time for reference |
    # ----------------------------------------
    #region Get NetScaler System Time
        # Specifying the correct URL 
        $strURI = "https://$NSIP/nitro/v1/config/nsconfig"

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
                Write-Verbose "REST API call to retrieve system time: successful"
                If ($response.nsconfig)
                {
                    $referenceDate = Get-Date "1/1/1970"
                    $nsDate = $referenceDate.AddSeconds($response.nsconfig.systemtime)

                    Write-Verbose ("Calculated NS Date: " + $nsDate)
                }
                Else
                {
                    If ($response.errorcode -eq 0)
                    {
                        Write-Output ""
                        Write-Output "No system time was found."
                    }
                    Else
                    {
                        Write-Error ("errorcode: " + $response.errorcode + ", message: " + $response.message)
                    }
                }
            }
        }
    #endregion

    #region Get NetScaler License Server Type
        # Specifying the correct URL 
        $strURI = "https://$NSIP/nitro/v1/config/nslicense"

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
                Write-Verbose "REST API call to retrieve NS license information: successful"
                If ($response.nslicense)
                {
                    $licensingMode = $response.nslicense.licensingmode
                    #Write-Host ("licensing mode: " + $licensingMode) -ForegroundColor Yellow
                    If (($licensingMode -eq "Pooled") -or ($licensingMode -eq "CICO"))
                    {
                        # retrieve license server information
                        Write-Output "This NetScaler is using a license server"

                        #region Retrieve License Server information
                            # Specifying the correct URL 
                            $strURI = "https://$NSIP/nitro/v1/config/nslicenseserver"

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
                                #$response.nslicenseserver | Format-List
                                $response.nslicenseserver | Select-Object @{N='Server Name'; E={$_.servername}}, 
                                                                    @{N='IP address'; E={$_.licenseserverip}}, 
                                                                    @{N='Port'; E={$_.port}}, 
                                                                    @{N='Status'; E={$_.status}}, 
                                                                    @{N='Grace status'; E={$_.grace}}, 
                                                                    @{N='Grace time left'; E={$_.gptimeleft}} | Format-List

                            }
                        
                        #endregion

                    }
                    If ($licensingMode -eq "Local")
                    {
                        Write-Output "This NetScaler is using a local license file."
                        #region Get License File(s) information

                            # Specifying the correct URL 
                            $strURI = ("https://$NSIP/nitro/v1/config/systemfile?args=filelocation:" + [System.Web.HttpUtility]::UrlEncode("/nsconfig/license"))

                            # Method #1: Making the REST API call to the NetScaler
                            try
                            {
                                # start with clean response variable
                                $responseFiles = $null
                                $responseFiles = Invoke-RestMethod -Method Get -Uri $strURI -ContentType $ContentType -WebSession $NetScalerSession -Verbose:$VerbosePreference -ErrorAction SilentlyContinue
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
                                If ($responseFiles.errorcode -eq 0)
                                {
                                    Write-Verbose "REST API call to retrieve all NetScaler system files in ""\nsconfig\license\"": successful"
                                    If ($responseFiles.systemfile.Count -gt 0)
                                    {
                                        [array]$licFiles = $responseFiles.systemfile | Where-Object {$_.filename -like "*.lic"}

                                        # Show how many license files were found
                                        Write-Output ""
                                        Write-Output ($licFiles.Count.ToString() + " license file(s) found in \nsconfig\license.")

                                        foreach ($licFile in $licFiles)
                                        {
                                            Write-Output ""
                                            Write-Output ("* License File: " + $licFile.filename)
                    
                                            # Specifying the correct URL (to retrieve the license file details)
                                            $strURI = ("https://$NSIP/nitro/v1/config/systemfile?args=filename:" + [System.Web.HttpUtility]::UrlEncode($licFile.filename) + ",filelocation:" + [System.Web.HttpUtility]::UrlEncode("/nsconfig/license")) 
                    
                                            # Method #1: Making the REST API call to the NetScaler
                                            try
                                            {
                                                # start with clean response variable
                                                $responseFile = $null
                                                $responseFile = Invoke-RestMethod -Method Get -Uri $strURI -ContentType $ContentType -WebSession $NetScalerSession -Verbose:$VerbosePreference #-ErrorAction SilentlyContinue
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
                                                $results = @()
                                                $featresults = @()
                                                #Get actual file content
                                                If ($responseFile.systemfile.fileencoding -eq "BASE64")
                                                {
                                                    # Get license file content
                                                    $FileContent = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($responseFile.systemfile.filecontent))
                                                }
                                                Else
                                                {
                                                    Throw ("Fileencoding " + $responseFile.systemfile.fileencoding + " unknown")
                                                }
                        
                                                #Process File Content
                                                #Grabs needed line that has licensing information
                                                $licLines = $FileContent.Split("`n")|where-object{$_ -like "*INCREMENT*"}
                                                $licdates = @()

                                                # Process the ecpire date
                                                foreach ($line in $licLines)
                                                {
                                                    $licdate = $line.Split()
    
                                                    if ($licdate[4] -like "permanent")
                                                    {
                                                        $expire = "PERMANENT"
                                                    }
                                                    else
                                                    {
                                                        $expire = [datetime]$licdate[4]
                                                    }
    
                                                    #adds date to object
                                                    $temp = New-Object PSObject -Property @{
                                                                expdate = $expire
                                                                feature = $licdate[1]
                                                                }
                                                    $licdates += $temp
                                                }

                                                # Get Feature descriptions from License File
                                                $featLines = $FileContent.Split("`n")|where-object{$_ -like "#CITRIXTERM*`tEN`t*"}
                                                foreach ($featLine in $featLines)
                                                {
                                                    $feat = $featLine.Split("`t",4)
                                                        # test[1] = feature name
                                                        # test[3] = feature description
                                                    $feattemp = New-Object PSObject -Property @{
                                                        LicenseFile = $licFile.filename
                                                        Feature = $feat[1]
                                                        Description = $feat[3]
                                                    }
                                                    $featresults += $feattemp
                                                }

                                                # Get License file information
                                                foreach ($date in $licdates)
                                                {
                                                    if ($date.expdate -like "PERMANENT")
                                                    {
                                                        $expires = "PERMANENT"
                                                        $span = "9999"
                                                    }
                                                    else
                                                    {
                                                        $expires = ($date.expdate).ToShortDateString()
                                                        $span = (New-TimeSpan -Start $nsDate -end ($date.expdate)).days
                                                    }
    
                                                    $temp = New-Object PSObject -Property @{
                                                        Expires = $expires
                                                        Feature = $date.feature
                                                        Description = ($featresults | Where-Object {($_.LicenseFile -eq $licFile.filename) -and ($_.Feature -eq $date.feature)}).Description # 
                                                        DaysLeft = $span
                                                        LicenseFile = $licFile.filename
                                                        ModifiedTime = $licFile.filemodifiedtime
                                                    }
                                                    # Link License Expire Date for each type to License description for each type
                                                    $results += $temp    
                                                }    
                                            $results | Select-Object Feature, @{N='Expire Date'; E={$_.Expires}}, DaysLeft, Description| Format-Table  # ModifiedTime, Expires not shown in results
                                            }
                                        }
                                    }
                                    Else
                                    {
                                        If ($response.errorcode -eq 0)
                                        {
                                            Write-Output ""
                                            Write-Warning "No license file(s) were found."
                                        }
                                        Else
                                        {
                                            Write-Error ("errorcode: " + $response.errorcode + ", message: " + $response.message)
                                        }
                                    }
                                }
                                Else
                                {
                                    Write-Warning "An error occured. (""errorcode: " + $response.errorcode + ", message: " + $response.message + """)"
                                }
                            }
                        #endregion

                    }
                }
                Else
                {
                    If ($response.errorcode -eq 0)
                    {
                        Write-Output ""
                        Write-Output "No license information was found."
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
    Get-NSLicenseDetails -NSIP $args[0] -NSUserName $args[1] -NSUserPW $args[2]
}
catch [System.Management.Automation.ParameterBindingException] {
    Write-Error "Couldn't bind parameter exception, Please make sure to provide all necessary parameters"
}
