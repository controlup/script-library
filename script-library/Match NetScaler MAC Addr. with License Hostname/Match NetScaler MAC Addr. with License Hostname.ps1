function Get-NSLicFromMAC {
    <#
    .SYNOPSIS
      Match the MAC Address of the first NetScaler interface with the Hostname in the License file on the NetScaler appliance.
    .DESCRIPTION
      Match the MAC Address of the first NetScaler interface with the Hostname in the License file on the NetScaler appliance (when using local license files for NetScaler licensing), using the Invoke-RestMethod cmdlet for the REST API calls.
    .NOTES
      Version:        0.3
      Author:         Esther Barthel, MSc
      Creation Date:  2018-03-18
      Updated:        2018-04-03
                      Changed Interface ID from 1/1 (NUC) to 0/1 (for the VirtualBox NetScaler)
      Updated:        2018-04-08
                      Using the unit number to filter the interfaces on the NetScaler to get the right interface MAC Address for the licensing check
      Purpose:        SBA - Created for ControlUp NetScaler Monitoring
      Credits:        A big shoutout to Ryan Butler (Citrix CTA) for showing me how to retrieve data from the license files of the NetScaler

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

    Write-Output "----------------------------------------------------------------- " #-ForegroundColor Yellow
    Write-Output "| Match MAC Address and License File Hostname on the NetScaler: | " #-ForegroundColor Yellow
    Write-Output "----------------------------------------------------------------- " #-ForegroundColor Yellow
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

    #region Check MAC Address on NetScaler first NIC
        ## Specifying the correct URL 
        # Filter to get the first NIC 
        # NOTE: eth0 interface id can vary between 0/1 and 1/1, so interface id is not a good filter!! Switched to unit.
        $strURI = ("https://$NSIP/nitro/v1/config/Interface?filter=unit:0")

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
                Write-Verbose "REST API call to retrieve NS license information: Successful"
                If ($response.interface)
                {
                    # debug info:
                    #$response.interface | Select-Object id, devicename, unit, decription, vlan, mac, ifnum, intftype
                    $MacAddress = $response.Interface.mac.Replace(":","")
                    $intID = [string]$response.Interface.id
                    Write-Host "NetScaler MAC Address: " -NoNewline
                    Write-Host ($MacAddress + " (int " + $intID + ")") -ForegroundColor Yellow
                }
                Else
                {
                    If ($response.errorcode -eq 0)
                    {
                        Write-Output ""
                        Write-Output "No interface information was found."
                    }
                    Else
                    {
                        Write-Error ("errorcode: " + $response.errorcode + ", message: " + $response.message)
                    }
                }
            }
        }
    #endregion

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
                Write-Verbose "REST API call to retrieve all NetScaler system files in ""\nsconfig\license\"": Successful"
                ## debug info
                #$responseFiles.systemfile
                $licCount = ($responseFiles.systemfile | Where-Object {($_.filesize -gt 0) -and ($_.filename -like "*.lic")}).filename.Count
                If ($licCount -gt 0)
                {
                    $licFiles = $responseFiles.systemfile | Where-Object {$_.filename -like "*.lic"}
                
                    # Show how many license files were found
                    Write-Output ""
                    Write-Host ($licCount.ToString() + " license file(s) found in \nsconfig\license.") -ForegroundColor DarkYellow

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
                            $licLines = $FileContent.Split("`n")|where-object{$_ -like "*SERVER this_host*"}
                            $licdates = @()

                            # Process the ecpire date
                            foreach ($line in $licLines)
                            {
                                $HostName = $line.Replace("SERVER this_host ", "")
                                Write-Host "`t" -NoNewline
                                Write-Host ("- Hostname: " + $HostName) -ForegroundColor Yellow
                                If ($MacAddress -eq $HostName)
                                {
                                    Write-Host "`t" -NoNewline
                                    Write-Host "`t" -NoNewline
                                    Write-Host "MATCH: MAC Address and Hostname are identical! This license file can be used on this appliance." -ForegroundColor Green
                                    Write-Output ""
                                }
                                Else
                                {
                                    Write-Host "`t" -NoNewline
                                    Write-Host "`t" -NoNewline
                                    Write-Warning "MAC Address and Hostname are NOT identical! This License file can NOT be used on this appliance."
                                    Write-Output ""
                                }
                            }
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
        Write-Output ""
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
    Get-NSLicFromMAC -NSIP $args[0] -NSUserName $args[1] -NSUserPW $args[2]
}
catch [System.Management.Automation.ParameterBindingException] {
    Write-Error "Couldn't bind parameter exception, Please make sure to provide all necessary parameters"
}
