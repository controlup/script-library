<#
.SYNOPSIS
    Set the CPU Yield parameter on the the Citrix ADC.
.DESCRIPTION
    Set the CPU Yield parameter on the the Citrix ADC, using NITRO.
.EXAMPLE
    Set-ADCVpxparamCpuyield -NSIP <ipaddress> 
.CONTEXT
    NetScalers
.MODIFICATION_HISTORY
    Esther Barthel, MSc - 05/01/20 - Original code
    Esther Barthel, MSc - 05/01/20 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
.LINK
    https://support.citrix.com/article/CTX2295559
.NOTES
    Version:        0.1
    Author:         Esther Barthel, MSc
    Creation Date:  2020-01-05
    Updated:        2020-01-05
                    Standardized the function, based on the ControlUp Standards (v0.2)
    Purpose:        Script Action, created for ControlUp Citrix ADC Management
        
    Copyright (c) cognition IT. All rights reserved.
#>

function Get-CUStoredCredential {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The system the credentials will be used for.")]
        [string]$System
    )

    # Get the stored credential object
    $strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
    
    try {
        Import-Clixml $strCUCredFolder\$($env:USERNAME)_$($System)_Cred.xml
    }
    catch {
        Write-Error ("The required PSCredential object could not be loaded. " + $_)
    }
}

function Get-ADCCredentials()
{
    <#
    .SYNOPSIS
        Retrieve the Citrix ADC Credentials.
    .DESCRIPTION
        Retrieve the Citrix ADC Credentials from either a stored credentials file or Get-Credential popup.
    .EXAMPLE
        Get-ADCCredentials
    .CONTEXT
        NetScalers
    .MODIFICATION_HISTORY
        Esther Barthel, MSc - 31/12/19 - Original code
        Esther Barthel, MSc - 31/12/19 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
    .COMPONENT
        Get-CUStoredCredential - to retreive the XML Credentials file for non-interactive/automated use of this script
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2019-12-31
        Updated:        2019-12-31
                        Standardized the function, based on the ControlUp Standards (v0.2)
        Purpose:        Script Action, created for ControlUp NetScaler Monitoring
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param()

    #region script settings
        # Stored ADC Credentials XML file
        $systemName = "ADC"
        $credTargetFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
        $credTarget = "$credTargetFolder\$($Env:Username)_$($systemName)_Cred.xml"
        # Declare ADC Credentials object
        [System.Management.Automation.PSCredential]$adcCredentials = $null
    #endregion

    Write-Verbose ""
    Write-Verbose "------------------------ "
    Write-Verbose "| Get ADC Credentials: | "
    Write-Verbose "------------------------ "
    Write-Verbose ""

    #region Load ADC Credentials either trough XML file import or Get-Credentials
        # Check for Stored Credentials
        If (Test-Path -Path $credTarget)
        # Stored credentials found, import credentials
        {
            try
            {
                $adcCredentials = Get-CUStoredCredential -System $systemName
            }
            catch
            {
                Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
                Exit
            }
            Write-Verbose "* ADC Credentials: Stored $systemName credentials XML file found. ADC credentials imported for Automated Action support."
        }
        Else
        # No Stored Credentials Found, ask for credentials
        {
            Write-Verbose "* ADC Credentials: Stored $systemName credentials XML file NOT found, using Get-Credential to retrieve ADC credentials."
            $adcCredentials = Get-Credential -Message "Enter your credentials for Citrix ADC $NSIP"
        }
    #endregion

    # Return the ADC Credentials (PSCredential object) for future use in the NITRO functions
    If (!($adcCredentials -eq $null))
    {
        Write-Verbose "* ADC Credentials: credentials returned."
        return $adcCredentials
    }
    Else
    {
        Write-Verbose "* ADC Credentials: NO credentials returned."
        Write-Error "No ADC Credentials retrieved, cannot perform NITRO actions."
        Exit
    }
}

function Get-ADCNsvpxparam ()
{
    <#
    .SYNOPSIS
        Retrieve nsvpxparam settings of the Citrix ADC.
    .DESCRIPTION
        Retrieve nsvpxparam settings of the Citrix ADC, using NITRO.
    .EXAMPLE
        Get-ADCNsvpxparam -NSIP 192.168.0.101
    .EXAMPLE
        Get-ADCNsvpxparam -NSIP 192.168.0.101 -NSCredentials $PSCredentialsObject
    .CONTEXT
        NetScalers
    .MODIFICATION_HISTORY
        Esther Barthel, MSc - 05/01/20 - Original code
        Esther Barthel, MSc - 05/01/20 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
    .LINK
        https://developer-docs.citrix.com/projects/netscaler-nitro-api/en/12.0/configuration/responder/responderpolicy_binding/responderpolicy_binding/
    .COMPONENT
        Get-ADCCredential - to retreive the XML Credentials file with the ADC credentials for non-interactive/automated use of this script
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-01-05
        Updated:        2020-01-05
                        Standardized the function, based on the ControlUp Standards (v0.2)
        Purpose:        Script Action, created for ControlUp NetScaler Monitoring
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Citrix ADC IP address to run the script on'
        )]
        [ValidateScript({$_ -match [IPAddress]$_ })]
        [string] $NSIP,
            
        [Parameter(
            Position=1, 
            Mandatory=$false, 
            HelpMessage='Enter a PSCredential object, containing the username and password'
        )]
        [System.Management.Automation.CredentialAttribute()] $ADCCredentials
    )    

    #region ControlUp Script Standards - version 0.2
        #Requires -Version 3.0

        # Configure a larger output width for the ControlUp PowerShell console
        [int]$outputWidth = 400
        # Altering the size of the PS Buffer
        $PSWindow = (Get-Host).UI.RawUI
        $WideDimensions = $PSWindow.BufferSize
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions

        # Ensure Debug information is shown, without the confirmation question after each Write-Debug
        If ($PSBoundParameters['Debug']) {$DebugPreference = "Continue"}
        If ($PSBoundParameters['Verbose']) {$VerbosePreference = "Continue"}
        $ErrorActionPreference = "Stop"
    #endregion

    #region script settings
        # Stored ADC Credentials XML file
        $systemName = "ADC"
        $credTargetFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
        $credTarget = "$credTargetFolder\$($Env:Username)_$($systemName)_Cred.xml"
        # NITRO Constants
        $ContentType = "application/json"
        # turn Verbose mode on
        #$VerbosePreference = "Continue"
        # turn Verbose mode off
        #$VerbosePreference="SilentlyContinue"
    #endregion

    Write-Verbose ""
    Write-Verbose "-------------------------------------------------- "
    Write-Verbose "| Get Citrix ADC nsvpxparam settings with NITRO: | "
    Write-Verbose "-------------------------------------------------- "
    Write-Verbose ""

    #region Load NSCredentials either trough XML file import or Get-Credentials
    If ($null -eq $ADCCredentials)
    {
        $ADCCredentials = Get-ADCCredentials
    }
    #endregion

    # Retieving username and password from PSCredentials for use with NITRO
    $NSUserName = $ADCCredentials.UserName
    $NSUserPW = $ADCCredentials.GetNetworkCredential().Password

    # ----------------------------------------
    # | Method #1: Using the SessionVariable |
    # ----------------------------------------
    #region Start NITRO Session
        #Force PowerShell to bypass validation for (self-signed) certificates and SSL connections
        # source: https://blogs.technet.microsoft.com/bshukla/2010/04/12/ignoring-ssl-trust-in-powershell-system-net-webclient/ 
        Write-Verbose "* Certificate Validation: Forcing PowerShell to trust all certificates (including the self-signed netScaler certificate)"
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

        #JSON payload
        $LoginJSON = ConvertTo-Json @{
            "login" = @{
                "username"=$NSUserName;
                "password"=$NSUserPW
            }
        }
        try
        {
            # Login to the NetScaler and create a session (stored in $NSSession)
            $invokeRestMethodParams = @{
                Uri             = "https://$NSIP/nitro/v1/config/login"
                Body            = $LoginJSON
                Method          = "Post"
                SessionVariable = "NSSession"
                ContentType     = $ContentType
            }
            $loginresponse = Invoke-RestMethod @invokeRestMethodParams
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct NetScaler IP-address. " + $_.Exception.Message)
            Exit
        }
        catch
        {
            If (($_.Exception.Message -like "*Unauthorized*") -and (Test-Path -Path $XMLFile))
            {
                Write-Warning "XML Stored NSCredentials were used, check if the stored NSCredentials are correct for this Citrix ADC (" + $NSIP + ")."
            }
            # Debug: 
            Write-Debug $_.Exception | Format-List -Force
            # Error:
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        # Check for REST API errors
        If ($loginresponse.errorcode -eq 0)
        {
            Write-Verbose "* NITRO: login successful"
        }
    #endregion Start NITRO Session

    # ------------------
    # | Get nsvpxparam |
    # ------------------
    #region nsvpxparam
        # start with clean response variable
        $adcNsvpxparam = $null
            
        # Create the Invoke-RestMethod params
        $uri = "https://$NSIP/nitro/v1/config/nsvpxparam"

        try
        {
            $invokeRestMethodParams = @{
                Uri         = $uri
                Method      = "GET"
                WebSession  = $NSSession
                ContentType = $ContentType
            }
            $adcNsvpxparam = Invoke-RestMethod @invokeRestMethodParams
        }
        catch
        {
            Write-Error $error[0] 
            Exit          
        }

        If ( -not ($adcNsvpxparam.errorcode -eq 0))
        # NITRO errorcode found (or results is empty)
        {
            Write-Error "NITRO: errorcode (" + $adcNsvpxparam.errorcode + ") " + $adcNsvpxparam.message + "."
            Exit
        }

        Write-Verbose "* NITRO: nsvpxparam retrieved successful"
        if ($adcNsvpxparam.nsvpxparam)
        {
            $results = $adcNsvpxparam.nsvpxparam
            return $results
        }
        else
        {
            Write-Warning "No nsvpxparam found."
        }
    #endregion

    #region End NetScaler NITRO Session
        #Disconnect from the NetScaler (cleanup session)
        $LogOut = @{
            "logout" = @{}
        } | ConvertTo-Json

        try
        {
            # Loout of the NetScaler and remove the session (stored in $NSSession)
            $invokeRestMethodParams = @{
                Uri             = "https://$NSIP/nitro/v1/config/logout"
                Body            = $LogOut
                Method          = "Post"
                WebSession      = $NSSession
                ContentType     = $ContentType
            }
            $logoutresponse = Invoke-RestMethod @invokeRestMethodParams
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct NetScaler IP-address. " + $_.Exception.Message)
            Exit
        }
        catch
        {
            # Debug: 
            Write-Debug $_.Exception | Format-List -Force
            # Error:
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        # Check for REST API errors
        If ($logoutresponse.errorcode -eq 0)
        {
            Write-Verbose "* NITRO: logout successful"
        }
    #endregion End NetScaler NITRO Session
}

function Get-ADCNshardware ()
{
    <#
    .SYNOPSIS
        Retrieve nshardware settings of the Citrix ADC.
    .DESCRIPTION
        Retrieve nshardware settings of the Citrix ADC, using NITRO.
    .EXAMPLE
        Get-ADCNshardware -NSIP 192.168.0.101
    .EXAMPLE
        Get-ADCNshardware -NSIP 192.168.0.101 -NSCredentials $PSCredentialsObject
    .CONTEXT
        NetScalers
    .MODIFICATION_HISTORY
        Esther Barthel, MSc - 05/01/20 - Original code
        Esther Barthel, MSc - 05/01/20 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
    .LINK
        https://developer-docs.citrix.com/projects/netscaler-nitro-api/en/12.0/configuration/responder/responderpolicy_binding/responderpolicy_binding/
    .COMPONENT
        Get-ADCCredential - to retreive the XML Credentials file with the ADC credentials for non-interactive/automated use of this script
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-01-05
        Updated:        2020-01-05
                        Standardized the function, based on the ControlUp Standards (v0.2)
        Purpose:        Script Action, created for ControlUp NetScaler Monitoring
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Citrix ADC IP address to run the script on'
        )]
        [ValidateScript({$_ -match [IPAddress]$_ })]
        [string] $NSIP,
            
        [Parameter(
            Position=1, 
            Mandatory=$false, 
            HelpMessage='Enter a PSCredential object, containing the username and password'
        )]
        [System.Management.Automation.CredentialAttribute()] $ADCCredentials
    )    

    #region ControlUp Script Standards - version 0.2
        #Requires -Version 3.0

        # Configure a larger output width for the ControlUp PowerShell console
        [int]$outputWidth = 400
        # Altering the size of the PS Buffer
        $PSWindow = (Get-Host).UI.RawUI
        $WideDimensions = $PSWindow.BufferSize
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions

        # Ensure Debug information is shown, without the confirmation question after each Write-Debug
        If ($PSBoundParameters['Debug']) {$DebugPreference = "Continue"}
        If ($PSBoundParameters['Verbose']) {$VerbosePreference = "Continue"}
        $ErrorActionPreference = "Stop"
    #endregion

    #region script settings
        # Stored ADC Credentials XML file
        $systemName = "ADC"
        $credTargetFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
        $credTarget = "$credTargetFolder\$($Env:Username)_$($systemName)_Cred.xml"
        # NITRO Constants
        $ContentType = "application/json"
        # turn Verbose mode on
        #$VerbosePreference = "Continue"
        # turn Verbose mode off
        #$VerbosePreference="SilentlyContinue"
    #endregion

    Write-Verbose ""
    Write-Verbose "-------------------------------------------------- "
    Write-Verbose "| Get Citrix ADC nshardware settings with NITRO: | "
    Write-Verbose "-------------------------------------------------- "
    Write-Verbose ""

    #region Load NSCredentials either trough XML file import or Get-Credentials
    If ($null -eq $ADCCredentials)
    {
        $ADCCredentials = Get-ADCCredentials
    }
    #endregion

    # Retieving username and password from PSCredentials for use with NITRO
    $NSUserName = $ADCCredentials.UserName
    $NSUserPW = $ADCCredentials.GetNetworkCredential().Password

    # ----------------------------------------
    # | Method #1: Using the SessionVariable |
    # ----------------------------------------
    #region Start NITRO Session
        #Force PowerShell to bypass validation for (self-signed) certificates and SSL connections
        # source: https://blogs.technet.microsoft.com/bshukla/2010/04/12/ignoring-ssl-trust-in-powershell-system-net-webclient/ 
        Write-Verbose "* Certificate Validation: Forcing PowerShell to trust all certificates (including the self-signed netScaler certificate)"
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

        #JSON payload
        $LoginJSON = ConvertTo-Json @{
            "login" = @{
                "username"=$NSUserName;
                "password"=$NSUserPW
            }
        }
        try
        {
            # Login to the NetScaler and create a session (stored in $NSSession)
            $invokeRestMethodParams = @{
                Uri             = "https://$NSIP/nitro/v1/config/login"
                Body            = $LoginJSON
                Method          = "Post"
                SessionVariable = "NSSession"
                ContentType     = $ContentType
            }
            $loginresponse = Invoke-RestMethod @invokeRestMethodParams
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct NetScaler IP-address. " + $_.Exception.Message)
            Exit
        }
        catch
        {
            If (($_.Exception.Message -like "*Unauthorized*") -and (Test-Path -Path $XMLFile))
            {
                Write-Warning "XML Stored NSCredentials were used, check if the stored NSCredentials are correct for this Citrix ADC (" + $NSIP + ")."
            }
            # Debug: 
            Write-Debug $_.Exception | Format-List -Force
            # Error:
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        # Check for REST API errors
        If ($loginresponse.errorcode -eq 0)
        {
            Write-Verbose "* NITRO: login successful"
        }
    #endregion Start NITRO Session

    # ------------------
    # | Get nshardware |
    # ------------------
    #region nshardware
        # start with clean response variable
        $adcNshardware = $null
            
        # Create the Invoke-RestMethod params
        $uri = "https://$NSIP/nitro/v1/config/nshardware"

        try
        {
            $invokeRestMethodParams = @{
                Uri         = $uri
                Method      = "GET"
                WebSession  = $NSSession
                ContentType = $ContentType
            }
            $adcNshardware = Invoke-RestMethod @invokeRestMethodParams
        }
        catch
        {
            Write-Error $error[0] 
            Exit          
        }

        If ( -not ($adcNshardware.errorcode -eq 0))
        # NITRO errorcode found (or results is empty)
        {
            Write-Error "NITRO: errorcode (" + $adcNshardware.errorcode + ") " + $adcNshardware.message + "."
            Exit
        }

        Write-Verbose "* NITRO: nsvpxparam retrieved successful"
        if ($adcNshardware.nshardware)
        {
            $results = $adcNshardware.nshardware
            return $results
        }
        else
        {
            Write-Warning "No nshardware found."
        }
    #endregion

    #region End NetScaler NITRO Session
        #Disconnect from the NetScaler (cleanup session)
        $LogOut = @{
            "logout" = @{}
        } | ConvertTo-Json

        try
        {
            # Loout of the NetScaler and remove the session (stored in $NSSession)
            $invokeRestMethodParams = @{
                Uri             = "https://$NSIP/nitro/v1/config/logout"
                Body            = $LogOut
                Method          = "Post"
                WebSession      = $NSSession
                ContentType     = $ContentType
            }
            $logoutresponse = Invoke-RestMethod @invokeRestMethodParams
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct NetScaler IP-address. " + $_.Exception.Message)
            Exit
        }
        catch
        {
            # Debug: 
            Write-Debug $_.Exception | Format-List -Force
            # Error:
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        # Check for REST API errors
        If ($logoutresponse.errorcode -eq 0)
        {
            Write-Verbose "* NITRO: logout successful"
        }
    #endregion End NetScaler NITRO Session
}

function Get-ADCNsversion ()
{
    <#
    .SYNOPSIS
        Retrieve nsversion settings of the Citrix ADC.
    .DESCRIPTION
        Retrieve nsversion settings of the Citrix ADC, using NITRO.
    .EXAMPLE
        Get-ADCNsversion -NSIP 192.168.0.101
    .EXAMPLE
        Get-ADCNsversion -NSIP 192.168.0.101 -NSCredentials $PSCredentialsObject
    .CONTEXT
        NetScalers
    .MODIFICATION_HISTORY
        Esther Barthel, MSc - 05/01/20 - Original code
        Esther Barthel, MSc - 05/01/20 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
    .LINK
        https://developer-docs.citrix.com/projects/netscaler-nitro-api/en/12.0/configuration/responder/responderpolicy_binding/responderpolicy_binding/
    .COMPONENT
        Get-ADCCredential - to retreive the XML Credentials file with the ADC credentials for non-interactive/automated use of this script
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-01-05
        Updated:        2020-01-05
                        Standardized the function, based on the ControlUp Standards (v0.2)
        Purpose:        Script Action, created for ControlUp NetScaler Monitoring
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Citrix ADC IP address to run the script on'
        )]
        [ValidateScript({$_ -match [IPAddress]$_ })]
        [string] $NSIP,
            
        [Parameter(
            Position=1, 
            Mandatory=$false, 
            HelpMessage='Enter a PSCredential object, containing the username and password'
        )]
        [System.Management.Automation.CredentialAttribute()] $ADCCredentials
    )    

    #region ControlUp Script Standards - version 0.2
        #Requires -Version 3.0

        # Configure a larger output width for the ControlUp PowerShell console
        [int]$outputWidth = 400
        # Altering the size of the PS Buffer
        $PSWindow = (Get-Host).UI.RawUI
        $WideDimensions = $PSWindow.BufferSize
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions

        # Ensure Debug information is shown, without the confirmation question after each Write-Debug
        If ($PSBoundParameters['Debug']) {$DebugPreference = "Continue"}
        If ($PSBoundParameters['Verbose']) {$VerbosePreference = "Continue"}
        $ErrorActionPreference = "Stop"
    #endregion

    #region script settings
        # Stored ADC Credentials XML file
        $systemName = "ADC"
        $credTargetFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
        $credTarget = "$credTargetFolder\$($Env:Username)_$($systemName)_Cred.xml"
        # NITRO Constants
        $ContentType = "application/json"
        # turn Verbose mode on
        #$VerbosePreference = "Continue"
        # turn Verbose mode off
        #$VerbosePreference="SilentlyContinue"
    #endregion

    Write-Verbose ""
    Write-Verbose "------------------------------------------------- "
    Write-Verbose "| Get Citrix ADC nsversion settings with NITRO: | "
    Write-Verbose "------------------------------------------------- "
    Write-Verbose ""

    #region Load NSCredentials either trough XML file import or Get-Credentials
    If ($null -eq $ADCCredentials)
    {
        $ADCCredentials = Get-ADCCredentials
    }
    #endregion

    # Retieving username and password from PSCredentials for use with NITRO
    $NSUserName = $ADCCredentials.UserName
    $NSUserPW = $ADCCredentials.GetNetworkCredential().Password

    # ----------------------------------------
    # | Method #1: Using the SessionVariable |
    # ----------------------------------------
    #region Start NITRO Session
        #Force PowerShell to bypass validation for (self-signed) certificates and SSL connections
        # source: https://blogs.technet.microsoft.com/bshukla/2010/04/12/ignoring-ssl-trust-in-powershell-system-net-webclient/ 
        Write-Verbose "* Certificate Validation: Forcing PowerShell to trust all certificates (including the self-signed netScaler certificate)"
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

        #JSON payload
        $LoginJSON = ConvertTo-Json @{
            "login" = @{
                "username"=$NSUserName;
                "password"=$NSUserPW
            }
        }
        try
        {
            # Login to the NetScaler and create a session (stored in $NSSession)
            $invokeRestMethodParams = @{
                Uri             = "https://$NSIP/nitro/v1/config/login"
                Body            = $LoginJSON
                Method          = "Post"
                SessionVariable = "NSSession"
                ContentType     = $ContentType
            }
            $loginresponse = Invoke-RestMethod @invokeRestMethodParams
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct NetScaler IP-address. " + $_.Exception.Message)
            Exit
        }
        catch
        {
            If (($_.Exception.Message -like "*Unauthorized*") -and (Test-Path -Path $XMLFile))
            {
                Write-Warning "XML Stored NSCredentials were used, check if the stored NSCredentials are correct for this Citrix ADC (" + $NSIP + ")."
            }
            # Debug: 
            Write-Debug $_.Exception | Format-List -Force
            # Error:
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        # Check for REST API errors
        If ($loginresponse.errorcode -eq 0)
        {
            Write-Verbose "* NITRO: login successful"
        }
    #endregion Start NITRO Session

    # -----------------
    # | Get nsversion |
    # -----------------
    #region nsversion
        # start with clean response variable
        $adcNsversion = $null
            
        # Create the Invoke-RestMethod params
        $uri = "https://$NSIP/nitro/v1/config/nsversion"

        try
        {
            $invokeRestMethodParams = @{
                Uri         = $uri
                Method      = "GET"
                WebSession  = $NSSession
                ContentType = $ContentType
            }
            $adcNsversion = Invoke-RestMethod @invokeRestMethodParams
        }
        catch
        {
            Write-Error $error[0] 
            Exit          
        }

        If ( -not ($adcNsversion.errorcode -eq 0))
        # NITRO errorcode found (or results is empty)
        {
            Write-Error "NITRO: errorcode (" + $adcNsversion.errorcode + ") " + $adcNsversion.message + "."
            Exit
        }

        Write-Verbose "* NITRO: nsvpxparam retrieved successful"
        if ($adcNsversion.nsversion)
        {
            $results = $adcNsversion.nsversion
            return $results
        }
        else
        {
            Write-Warning "No nsversion found."
        }
    #endregion

    #region End NetScaler NITRO Session
        #Disconnect from the NetScaler (cleanup session)
        $LogOut = @{
            "logout" = @{}
        } | ConvertTo-Json

        try
        {
            # Loout of the NetScaler and remove the session (stored in $NSSession)
            $invokeRestMethodParams = @{
                Uri             = "https://$NSIP/nitro/v1/config/logout"
                Body            = $LogOut
                Method          = "Post"
                WebSession      = $NSSession
                ContentType     = $ContentType
            }
            $logoutresponse = Invoke-RestMethod @invokeRestMethodParams
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct NetScaler IP-address. " + $_.Exception.Message)
            Exit
        }
        catch
        {
            # Debug: 
            Write-Debug $_.Exception | Format-List -Force
            # Error:
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        # Check for REST API errors
        If ($logoutresponse.errorcode -eq 0)
        {
            Write-Verbose "* NITRO: logout successful"
        }
    #endregion End NetScaler NITRO Session
}



function Set-ADCNsvpxparam ()
{
    <#
    .SYNOPSIS
        Configure nsvpxparam.
    .DESCRIPTION
        Configure nsvpxparam, using REST API.
    .EXAMPLE
        Set-Set-ADCNsvpxparam -NSIP 192.168.0.101 -CPUYield [DEFAULT|YES|NO]
    .EXAMPLE
        Set-Set-ADCNsvpxparam -NSIP 192.168.0.101 -CPUYield [DEFAULT|YES|NO] -NSCredentials $PSCredentialsObject
    .CONTEXT
        NetScalers
    .MODIFICATION_HISTORY
        Esther Barthel, MSc - 05/01/20 - Original code
        Esther Barthel, MSc - 05/01/20 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
    .LINK
        https://developer-docs.citrix.com/projects/netscaler-nitro-api/en/12.0/configuration/responder/responderpolicy_binding/responderpolicy_binding/
    .COMPONENT
        Get-ADCCredential - to retreive the XML Credentials file with the ADC credentials for non-interactive/automated use of this script
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-01-05
        Updated:        2020-01-05
                        Standardized the function, based on the ControlUp Standards (v0.2)
        Purpose:        Script Action, created for ControlUp NetScaler Monitoring
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Citrix ADC IP address to run the script on'
        )]
        [ValidateScript({$_ -match [IPAddress]$_ })]
        [string] $NSIP,
            
        [Parameter(
            Position=1, 
            Mandatory=$true, 
            HelpMessage='cpuyield setting appliance in virtual appliances. YES: Allow allocated but unused CPU resources to be used by another VM. NO: Reserve all CPU resources for the VM to which they have been allocated. This option shows higher percentage in hypervisor for VPX CPU usage.'
        )]
        [ValidateSet("YES","NO","DEFAULT")]
        [string] $CPUyield,

        [Parameter(
            Position=2, 
            Mandatory=$false, 
            HelpMessage='ID of the cluster node for which you are setting the cpuyield.'
        )]
        [int] $OwnerNode,

        [Parameter(
            Position=3, 
            Mandatory=$false, 
            HelpMessage='Enter a PSCredential object, containing the username and password'
        )]
        [System.Management.Automation.CredentialAttribute()] $ADCCredentials
    )    

    #region ControlUp Script Standards - version 0.2
        #Requires -Version 3.0

        # Configure a larger output width for the ControlUp PowerShell console
        [int]$outputWidth = 400
        # Altering the size of the PS Buffer
        $PSWindow = (Get-Host).UI.RawUI
        $WideDimensions = $PSWindow.BufferSize
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions

        # Ensure Debug information is shown, without the confirmation question after each Write-Debug
        If ($PSBoundParameters['Debug']) {$DebugPreference = "Continue"}
        If ($PSBoundParameters['Verbose']) {$VerbosePreference = "Continue"}
        $ErrorActionPreference = "Stop"
    #endregion

    #region script settings
        # Stored ADC Credentials XML file
        $systemName = "ADC"
        $credTargetFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
        $credTarget = "$credTargetFolder\$($Env:Username)_$($systemName)_Cred.xml"
        # NITRO Constants
        $ContentType = "application/json"
        # turn Verbose mode on
        #$VerbosePreference = "Continue"
        # turn Verbose mode off
        #$VerbosePreference="SilentlyContinue"
    #endregion

    Write-Verbose ""
    Write-Verbose "----------------------------------------- "
    Write-Verbose "| Set Citrix ADC nsvpxparam with NITRO: | "
    Write-Verbose "----------------------------------------- "
    Write-Verbose ""

    #region Load NSCredentials either trough XML file import or Get-Credentials
    If ($null -eq $ADCCredentials)
    {
        $ADCCredentials = Get-ADCCredentials
    }
    #endregion

    # Retieving username and password from PSCredentials for use with NITRO
    $NSUserName = $ADCCredentials.UserName
    $NSUserPW = $ADCCredentials.GetNetworkCredential().Password

    # ----------------------------------------
    # | Method #1: Using the SessionVariable |
    # ----------------------------------------
    #region Start NITRO Session
        #Force PowerShell to bypass validation for (self-signed) certificates and SSL connections
        # source: https://blogs.technet.microsoft.com/bshukla/2010/04/12/ignoring-ssl-trust-in-powershell-system-net-webclient/ 
        Write-Verbose "* Certificate Validation: Forcing PowerShell to trust all certificates (including the self-signed netScaler certificate)"
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

        #JSON payload
        $LoginJSON = ConvertTo-Json @{
            "login" = @{
                "username"=$NSUserName;
                "password"=$NSUserPW
            }
        }
        try
        {
            # Login to the NetScaler and create a session (stored in $NSSession)
            $invokeRestMethodParams = @{
                Uri             = "https://$NSIP/nitro/v1/config/login"
                Body            = $LoginJSON
                Method          = "Post"
                SessionVariable = "NSSession"
                ContentType     = $ContentType
            }
            $loginresponse = Invoke-RestMethod @invokeRestMethodParams
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct NetScaler IP-address. " + $_.Exception.Message)
            Exit
        }
        catch
        {
            If (($_.Exception.Message -like "*Unauthorized*") -and (Test-Path -Path $XMLFile))
            {
                Write-Warning "XML Stored NSCredentials were used, check if the stored NSCredentials are correct for this Citrix ADC (" + $NSIP + ")."
            }
            # Debug: 
            Write-Debug $_.Exception | Format-List -Force
            # Error:
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        # Check for REST API errors
        If ($loginresponse.errorcode -eq 0)
        {
            Write-Verbose "* NITRO: login successful"
        }
    #endregion Start NITRO Session

    # ------------------------
    # | Configure nsvpxparam |
    # ------------------------
    #region nsvpxparam
        # start with clean response variable
        $adcNsvpxparam = $null
            
        # Create the Invoke-RestMethod params
        $uri = "https://$NSIP/nitro/v1/config/nsvpxparam"

        # json payload
        $nsvpxparamHashtable = @{"cpuyield"=$CPUyield}
        If ($OwnerNode)
        {
            $nsvpxparamHashtable.Add("ownernode",$OwnerNode)
        }
        $jsonPayload = ConvertTo-Json @{
            "nsvpxparam"=$nsvpxparamHashtable
        } -Depth 5

        Write-Verbose ("json payload: " + $jsonPayload)
        try
        {
            $invokeRestMethodParams = @{
                Uri         = $uri
                Method      = "PUT"
                WebSession  = $NSSession
                ContentType = $ContentType
                Body        = $jsonPayload
            }
            $adcNsvpxparam = Invoke-RestMethod @invokeRestMethodParams
        }
        catch
        {
            Write-Error $error[0] 
            Exit          
        }
        if ($adcNsvpxparam.errorcode -ne 0)
        {
            Write-Error "nsvpxparam was not updated. Errorcode: $($adcNsvpxparam.errorcode)"
            Exit
        }
        Write-Verbose "* NITRO: nsvpxparam configured successful"
    #endregion

    #region End NetScaler NITRO Session
        #Disconnect from the NetScaler (cleanup session)
        $LogOut = @{
            "logout" = @{}
        } | ConvertTo-Json

        try
        {
            # Loout of the NetScaler and remove the session (stored in $NSSession)
            $invokeRestMethodParams = @{
                Uri             = "https://$NSIP/nitro/v1/config/logout"
                Body            = $LogOut
                Method          = "Post"
                WebSession      = $NSSession
                ContentType     = $ContentType
            }
            $logoutresponse = Invoke-RestMethod @invokeRestMethodParams
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Error ("A parameter binding ERROR occurred. Please provide the correct NetScaler IP-address. " + $_.Exception.Message)
            Exit
        }
        catch
        {
            # Debug: 
            Write-Debug $_.Exception | Format-List -Force
            # Error:
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        # Check for REST API errors
        If ($logoutresponse.errorcode -eq 0)
        {
            Write-Verbose "* NITRO: logout successful"
        }
    #endregion End NetScaler NITRO Session
}



#------------------------#
# Script Action workflow #
#------------------------#
Write-Host ""

## Retrieve input parameters
$NSIP = $args[0]

## Testing Script
#$NSIP = "192.168.0.99"

# Initiate variables
[System.Management.Automation.PSCredential]$adcCreds = $null
$isVPX = $false

# Step 0: Retrieve ADC Credentials
$adcCreds = Get-ADCCredentials #-Verbose

# Step 1: Check nshardware
Write-Host "* Step 1: Check if the Citrix ADC is a VPX appliance: " -ForegroundColor Yellow -NoNewline
$nshardware = Get-ADCNshardware -NSIP $NSIP -ADCCredentials $adcCreds #-Verbose -Debug
Write-Host "SUCCESS" -ForegroundColor Green
If ($nshardware.hwdescription -match "Virtual")
{
    $isVPX = $true
    Write-Host "`t=> The appliance hardware description `(""$($nshardware.hwdescription)""`) does indicate this is a VPX appliance." -ForegroundColor White
    Write-Host "`t=> The cpuyield setting will be changed to YES." -ForegroundColor White

    # Step 2: Check nsversion
    Write-Host "* Step 2: Check the Citrix ADC firmware version: " -ForegroundColor Yellow -NoNewline
    $nsversion = Get-ADCNsversion -NSIP $NSIP -ADCCredentials $adcCreds #-Verbose -Debug
    Write-Host "SUCCESS" -ForegroundColor Green

    #region Get NS Version information
	    # Extract NS version
	    $versionString = $nsversion.version
        $adcVersionRegex = "^([0-9]{2}[\.][0-9])"
	    if (($versionString -replace "NetScaler NS","") -match $adcVersionRegex)
		    {
			    Write-Verbose -Message "The Citrix ADC version string matched the supplied regex."
			    $adcVersion = ($versionString -replace "NetScaler NS","").Substring(0,4)
		    }
	    else
		    {
			    Write-Error "The NetScaler version string did not match the supplied regex."
                Exit
		    }
	
	    # Extract build number
	    $adcBuild = ($versionString.Split(",")).Split(":")[1].Trim().Replace("Build ","")
    #endregion
    Write-Host "`t=> firmware version: $($adcVersion) build $($adcBuild)." -ForegroundColor White

    # Step 3: Get current config
    Write-Host "* Step 3: Get current nsvpxparam settings: " -ForegroundColor Yellow -NoNewline
    $nsvpxparam = Get-ADCNsvpxparam -NSIP $NSIP -ADCCredentials $adcCreds #-Verbose -Debug
    Write-Host "SUCCESS" -ForegroundColor Green
    If ($nsvpxparam.cpuyield -eq "DEFAULT")
    {
        Switch ($adcVersion.Substring(0,2))
        {
            "12" {
                    Write-Host "`t=> cpuyield is currently set to DEFAULT (NO: Reserve all CPU resources for the VM to which they have been allocated. This option shows higher percentage in hypervisor for VPX CPU usage)." -ForegroundColor White
                    $changeCpuyield
                 }
            default {Write-Host "`t=> cpuyield is currently set to DEFAULT (YES: Allow allocated but unused CPU resources to be used by another VM)." -ForegroundColor White}
        }
    }
    Else
    {
        Write-Host "`t=> cpuyield is currently set to $($nsvpxparam.cpuyield)." -ForegroundColor White
    }

    # Step 4: Change the nsvpxparam cpuyield setting to YES
    Write-Host "* Step 4: Change cpuyield to YES: " -ForegroundColor Yellow -NoNewline
    Set-ADCNsvpxparam -NSIP $NSIP -CPUyield YES -ADCCredentials $adcCreds #-Verbose -Debug
    Write-Host "SUCCESS" -ForegroundColor Green

    # Step 5: Get new config
    Write-Host "* Step 5: Get changed nsvpxparam settings: " -ForegroundColor Yellow -NoNewline
    $nsvpxparam = Get-ADCNsvpxparam -NSIP $NSIP -ADCCredentials $adcCreds #-Verbose -Debug
    Write-Host "SUCCESS" -ForegroundColor Green
    Write-Host "`t=> cpuyield is now set to $($nsvpxparam.cpuyield)." -ForegroundColor White

    Write-Output ""
    Write-Host "The cpuyield setting was changed by this Script Action."
}
Else
{
    Write-Host "`t=> The appliance hardware description `(""$($nshardware.hwdescription)""`) does NOT indicate this is a VPX appliance." -ForegroundColor White
    Write-Output ""
    Write-Warning "The cpuyield setting will NOT be changed by this Script Action."
}

