<#
.SYNOPSIS
    Set WVD 2020 Spring Release Session Host Drain Mode.
.DESCRIPTION
    Set WVD 2020 Spring Release Session Host Drain Mode, using the new Az.Desktopvirtualization PowerShell Module and WVD ARM Architecture (2020 Spring Release).
.EXAMPLE
    Set-WVDSessionHostDrainMode -SessionHostName <SessionHostName>
.CONTEXT
    Windows Virtual Desktops
.MODIFICATION_HISTORY
    Esther Barthel, MSc - 19/05/20 - Original code
    Esther Barthel, MSc - 19/05/20 - Changed the script to support WVD 2020 Spring Release (ARM Architecture update)
.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/import-clixml?view=powershell-7
.COMPONENT
    Set-AzSPCredentials - The required Azure Service Principal (Subcription level) and tenantID information need to be securely stored in a Credentials File. The Set-AzSPCredentials Script Action will ensure the file is created according to ControlUp standards
    Az.Desktopvirtualization PowerShell Module - The Az.Desktopvirtualization PowerShell Module must be installed on the machine running this Script Action
.NOTES
    Version:        2.0
    Author:         Esther Barthel, MSc
    Creation Date:  2020-05-19
    Updated:        2020-05-19
                    Changed the script to support WVD 2020 Spring Release (ARM Architecture update)
    Updated:        2020-07-10
                    SBA rework to translate Session Host Name to HostPool ID (ControlUp WVD Script Action Integration workaround for current CU Console limitations)
    Purpose:        Script Action, created for ControlUp WVD Monitoring
        
    Copyright (c) cognition IT. All rights reserved.
#>
[CmdletBinding()]
Param
(
    [Parameter(
        Position=0, 
        Mandatory=$true, 
        HelpMessage='SBA parameter auto entry: Session Host NetBIOS Name'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $SessionHostName,

    [Parameter(
        Position=1, 
        Mandatory=$true, 
        HelpMessage='specify the Drain Mode'
    )]
    [ValidateSet("ON","OFF")]
    [string] $DrainMode
)    

# dot sourcing WVD Functions
# Windows Presentation Foundation (WPF) form to store WVD Service Principal information
[string]$mainformXAML = @'
<Window x:Class="wvdSP_Input_Form.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Esther_s_Input_Form"
        mc:Ignorable="d"
        Title="Enter the WVD Service Principal (SP) details" Height="389.336" Width="617.103">
    <Grid>
        <TextBox x:Name="textboxTenantId" HorizontalAlignment="Left" Height="31" Margin="176,50,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="398"/>
        <Label Content="SP Tenant ID" HorizontalAlignment="Left" Height="30" Margin="29,51,0,0" VerticalAlignment="Top" Width="117"/>
        <TextBox x:Name="textboxAppId" HorizontalAlignment="Left" Height="30" Margin="176,118,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="398"/>
        <Label Content="SP App ID" HorizontalAlignment="Left" Height="30" Margin="29,118,0,0" VerticalAlignment="Top" Width="117"/>
        <PasswordBox x:Name="textboxAppSecret" HorizontalAlignment="Left" Height="31" Margin="176,192,0,0" VerticalAlignment="Top" Width="398"/>
        <Label Content="SP App Secret" HorizontalAlignment="Left" Height="30" Margin="29,193,0,0" VerticalAlignment="Top" Width="117"/>
        <Button x:Name="buttonOK" Content="OK" HorizontalAlignment="Left" Height="46" Margin="29,274,0,0" VerticalAlignment="Top" Width="175" IsDefault="True"/>
        <Button x:Name="buttonCancel" Content="Cancel" HorizontalAlignment="Left" Height="46" Margin="244,274,0,0" VerticalAlignment="Top" Width="175" IsDefault="True"/>

    </Grid>
</Window>
'@

Function Invoke-WVDSPCredentialsForm {
# Created by Guy Leech - @guyrleech 17/05/2020
    Param
    (
        [Parameter(Mandatory=$true)]
        $inputXaml
    )

    $form = $null
    $inputXML = $inputXaml -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
    [xml]$xaml = $inputXML

    if( $xaml )
    {
        $reader = New-Object -TypeName Xml.XmlNodeReader -ArgumentList $xaml

        try
        {
            $form = [Windows.Markup.XamlReader]::Load( $reader )
        }
        catch
        {
            Throw "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$_"
        }

        $xaml.SelectNodes( '//*[@Name]' ) | ForEach-Object `
        {
            Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
        }
    }
    else
    {
        Throw "Failed to convert input XAML to WPF XML"
    }

    $form
}

function Get-AzSPStoredCredentials {
    <#
    .SYNOPSIS
        Retrieve the Azure Service Principal Stored Credentials.
    .DESCRIPTION
        Retrieve the Azure Service Principal Stored Credentials from a stored credentials file.
    .EXAMPLE
        Get-AzSPStoredCredentials
    .CONTEXT
        Azure
    .MODIFICATION_HISTORY
        Esther Barthel, MSc - 03/03/20 - Original code
        Esther Barthel, MSc - 03/03/20 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
    .COMPONENT
        Import-Clixml - https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/import-clixml?view=powershell-5.1
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-03-03
        Updated:        2020-03-03
                        Standardized the function, based on the ControlUp Standards (v0.2)
        Updated:        2020-05-08
                        Created a separate Azure Credentials function to support ARM architecture and Az PowerShell Module scripted actions
        Purpose:        Script Action, created for ControlUp WVD Monitoring
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param()

    #region function settings
        # Stored Credentials XML file
        $System = "AZ"
        $strAzSPCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
        $AzSPCredentials = $null
    #endregion

    Write-Verbose ""
    Write-Verbose "----------------------------- "
    Write-Verbose "| Get Azure SP Credentials: | "
    Write-Verbose "----------------------------- "
    Write-Verbose ""

    If (Test-Path -Path "$($strAzSPCredFolder)\$($env:USERNAME)_$($System)_Cred.xml")
    {
        try 
        {
            $AzSPCredentials = Import-Clixml -Path "$strAzSPCredFolder\$($env:USERNAME)_$($System)_Cred.xml"
        }
        catch 
        {
            Write-Error ("The required PSCredential object could not be loaded. " + $_)
        }
    }
    Else
    {
        Write-Error "The Azure Service Principal Credentials file stored for this user ($($env:USERNAME)) cannot be found. `nCreate the file with the Set-AzSPCredentials script action (prerequisite)."
        Exit
    }
    return $AzSPCredentials
}

function Set-AzSPStoredCredentials {
    <#
    .SYNOPSIS
        Store the Azure Service Principal Credentials.
    .DESCRIPTION
        Store the Azure Service Principal Credentials to an encrypted stored credentials file.
    .EXAMPLE
        Set-AzSPStoredCredentials
    .CONTEXT
        Azure
    .MODIFICATION_HISTORY
        Esther Barthel, MSc - 22/03/20 - Original code
        Esther Barthel, MSc - 22/03/20 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
    .COMPONENT
        Export-Clixml - https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/export-clixml?view=powershell-5.1
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-03-22
        Updated:        2020-03-22
                        Standardized the function, based on the ControlUp Standards (v0.2)
        Updated:        2020-05-08
                        Created a separate Azure Credentials function to support ARM architecture and Az PowerShell Module scripted actions
        Purpose:        Script Action, created for ControlUp WVD Monitoring
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param(
    )

    #region function settings
        # Stored Credentials XML file
        $System = "AZ"
        $strAzSPCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
        $AzSPCredentials = $null
    #endregion

    Write-Verbose ""
    Write-Verbose "------------------------------- "
    Write-Verbose "| Store Azure SP Credentials: | "
    Write-Verbose "------------------------------- "
    Write-Verbose ""

    If (!(Test-Path -Path "$($strAzSPCredFolder)"))
    {
        New-Item -ItemType Directory -Path "$($strAzSPCredFolder)"
        Write-Verbose "* AzSPCredentials: Path $($strAzSPCredFolder) created"
    }
    try 
    {
        Add-Type -AssemblyName PresentationFramework
        # Show the Form that will ask for the WVD Service Principal information (tenant ID, App ID, & App Secret)
        if( $mainForm = Invoke-WVDSPCredentialsForm -inputXaml $mainformXAML )
        {
            $WPFbuttonOK.Add_Click( {
                $_.Handled = $true
                $mainForm.DialogResult = $true
                $mainForm.Close()
            })
        
            $WPFbuttonCancel.Add_Click( {
                $_.Handled = $true
                $mainForm.DialogResult = $false
                $mainForm.Close()
            })
        
            $null = $WPFtextboxTenantId.Focus()
        
            if( $mainForm.ShowDialog() )
            {
                # Retrieve the form input (and check for errors)
                # tenant ID
                If ([string]::IsNullOrEmpty($($WPFtextboxTenantId.Text)))
                {
                    Write-Error "The provided tenant ID is empty!"
                    Exit
                }
                else 
                {
                    $tenantID = $($WPFtextboxTenantId.Text)
                }
                # app ID
                If ([string]::IsNullOrEmpty($($WPFtextboxAppId.Text)))
                {
                    Write-Error "The provided app ID is empty!"
                    Exit
                }
                else 
                {
                    $appID = $($WPFtextboxAppId.Text)
                }
                # app Secret
                If ([string]::IsNullOrEmpty($($WPFtextboxAppSecret.Password)))
                {
                    Write-Error "The provided app Secret is empty!"
                    Exit
                }
                else 
                {
                    $appSecret = $($WPFtextboxAppSecret.Password)
                }

                
                $appSecret = $($WPFtextboxAppSecret.Password)
            }
            else 
            {
                Write-Error "The required tenant ID, app ID and app Secret could not be retrieved from the form."
                Break
            }
        }
    }
    catch
    {
        Write-Error ("The required information could not be retrieved from the input form. " + $_)
        Exit        
    }
        # Create the SP Credentials, so they are encrypted before being stored in the XML file
        $secureAppSecret = ConvertTo-SecureString -String $appSecret -AsPlainText -Force
        $spCreds = New-Object System.Management.Automation.PSCredential($appID, $secureAppSecret)

    try
    {
        $hashAzSPCredentials = @{
            'tenantID' = $tenantID
            'spCreds' = $spCreds
        }
        $AzSPCredentials = Export-Clixml -Path "$strAzSPCredFolder\$($env:USERNAME)_$($System)_Cred.xml" -InputObject $hashAzSPCredentials -Force
    }
    catch 
    {
        Write-Error ("The required PSCredential object could not be exported. " + $_)
        Exit
    }
    Write-Verbose "* AzSPCredentials: Exported succesfully."
    return $hashAzSPCredentials
}

function Invoke-CheckInstallAndImportPSModulePrereq() {
    <#
    .SYNOPSIS
        Check, Install (if allowed) and Import the given PSModule prerequisite.
    .DESCRIPTION
        Check, Install (if allowed) and Import the given PSModule prerequisite.
    .EXAMPLE
        Invoke-CheckInstallAndImportPSModulePrereq -ModuleName Az.DesktopVirtualization
    .CONTEXT
        Windows Virtual Desktops
    .MODIFICATION_HISTORY
        Esther Barthel, MSc - 23/03/20 - Original code
        Esther Barthel, MSc - 23/03/20 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
    .COMPONENT
    .NOTES
        Version:        1.0
        Author:         Esther Barthel, MSc
        Creation Date:  2020-03-23
        Updated:        2020-03-23
                        Standardized the function, based on the ControlUp Standards (v0.2)
        Updated:        2020-07-22
                        Added an extra check to the Invoke-CheckInstallAndImportPSModuleRepreq function for the required NuGet packageprovider. If minimumversion 2.8.5.201 is not found it will install the NuGet PackageProvider
        Purpose:        Script Action, created for ControlUp WVD Monitoring
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the PowerShell Module name that needs to be installed and imported'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $ModuleName
    )

    #region Check if the given PowerShell Module is installed (and if not, install it with elevated right)
        # Check if the Module is loaded
        If (-not((Get-Module -Name $($ModuleName)).Name))
        # Module is not loaded
        {
            Write-Verbose "* CheckAndInstallPSModulePrereq: PowerShell Module $($ModuleName) is not loaded in current session"
            # Check if the Module is installed
            If (-not((Get-Module -Name $($ModuleName) -ListAvailable).Name))
            # Module is not installed on the system
            {
                # Check if session is evelated
                [bool]$isElevated = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                if ($isElevated)
                {
                    Write-Host ("The PowerShell Module $($ModuleName) is installed from an elevated session, Scope is set to AllUsers.") -ForegroundColor Yellow
                    $psScope = "AllUsers"
                }
                else
                {
                    Write-Warning "The PowerShell Module $($ModuleName) is NOT installed from an elevated session, Scope is set to CurrentUser."
                    $psScope = "CurrentUser"
                }

                # Check the version of the installed NuGet Provider and install if the version is lower than 2.8.5.201 or the provider is missing
                try
                {
                    If (!(([string](Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue).Version) -ge "2.8.5.201"))
                    {
                        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope $psScope -Force -Confirm:$false -WarningAction SilentlyContinue 
                    }
                }
                catch
                {
                    Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
                    Exit
                }
                Write-Verbose "* CheckAndInstallPSModulePrereq: The prerequired NuGet PackageProvider is installed."

                # Install the Module from the PSGallery
                try
                {
                    # Install the Module from the PSGallery
                    Write-Verbose "* CheckAndInstallPSModulePrereq: Installing the $($ModuleName) PowerShell Module from the PSGallery."
                    PowerShellGet\Install-Module -Name $($ModuleName) -Confirm:$false -Force -AllowClobber -Scope $psScope -WarningAction SilentlyContinue
                }
                catch
                {
                    Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
                    Exit
                }
                Write-Verbose "* CheckAndInstallPSModulePrereq: The $($ModuleName) PowerShell Module is installed from the PSGallery."
            }
        }
        # Import the Module
        try 
        {
            Import-Module -Name $($ModuleName) -Force
        }
        catch 
        {
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        Write-Verbose "* CheckAndInstallPSModulePrereq: The $($ModuleName) PowerShell Module is imported in the current session."
    #endregion Check PS Module status
}

function Invoke-NETFrameworkCheck() {
    <#
    .SYNOPSIS
        Check if the .NET Framework version is 4.7.2 or up.
    .DESCRIPTION
        Check if the .NET Framework version is 4.7.2 or up.
    .EXAMPLE
        Invoke-NETFrameworkCheck
    .CONTEXT
        Windows Virtual Desktops
    .MODIFICATION_HISTORY
        Esther Barthel, MSc - 17/05/20 - Original code
        Esther Barthel, MSc - 17/05/20 - Standardizing script, based on the ControlUp Scripting Standards (version 0.2)
    .COMPONENT
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-05-17
        Updated:        2020-05-17
                        Standardized the function, based on the ControlUp Standards (v0.2)
        Purpose:        Script Action, created for ControlUp WVD Monitoring
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param()

    #region Check if the current .NET Framework is 4.7.2 or higher (Release Value = 461808 or higher) prerequisite for the Az PowerShell Module
        If ((Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release).Release -ge 461808)
        {
            # Required .NET Framework found
            Write-Verbose ".NET Framework 4.7.2 or up is installed on this machine"
            return $true
        }
        Else
        {
            return $false
        }
    #endregion Check .NET Framework
}

function Get-WVDHostpoolSessionHost () {
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Subscription ID'
        )]
        [string] $SubscriptionID,
    
        [Parameter(
            Position=1, 
            Mandatory=$false, 
            HelpMessage='Enter the Host Pool name'
        )]
        [string] $HostPoolName,

    [Parameter(
        Position=2, 
        Mandatory=$false, 
        HelpMessage='Enter the Session Host name'
    )]
    [string] $SessionHostName
        )

    #Retrieve the WVD hostPool information for this Subscription
    [array]$psobjectCollection = Get-AzWvdHostPool -SubscriptionId $($azSubscription.Id.ToString()) | 
    ForEach-Object {
        $wvdHostPool = $_
        If ($wvdHostPool.Name -like "$HostPoolName*")
        {
            # Retrieve the SessionHost information for each HostPool
            Get-AzWvdSessionHost -HostPoolName $wvdHostpool.Name -ResourceGroupName $($wvdHostPool.Id.Split("/")[4]) | 
            ForEach-Object {
                $wvdSessionHost = $_
                If ($wvdSessionHost.Name.Split("/")[1] -like "$SessionHostName*")
                {
                    # retrieve both SessionHost and corresponding HostPool information for each SessionHost
                    $wvdSessionHost | Select @{Name='HostPoolName'; Expression={$wvdHostPool.Name}}, 
                        @{Name='SessionHostName'; Expression={$_.Name.Split("/")[1]}}, 
                        AgentVersion, 
                        AllowNewSession, 
                        @{Name='DrainMode'; Expression={if($_.AllowNewSession -eq "True"){return "Off"}else{return "On"}}}, 
                        AsignedUser, 
                        LastHeartBeat,
                        LastUpdateTime,
                        OSVersion,
                        Session, 
                        @{Name='ActiveSessions'; Expression={$_.Session}},  
                        Status,
                        StatusTimestamp,
                        SxSStackVersion, 
                        Type, 
                        UpdateErrorMessage, 
                        UpdateState, 
                        @{Name='HostPoolApplicationGroupReference'; Expression={$wvdHostPool.ApplicationGroupReference}}, 
                        @{Name='HostPoolCustomRdpProperty'; Expression={$wvdHostPool.CustomRdpProperty}}, 
                        @{Name='HostPoolDescription'; Expression={$wvdHostPool.Description}}, 
                        @{Name='HostPoolFriendlyName'; Expression={$wvdHostPool.FriendlyName}}, 
                        @{Name='HostPoolType'; Expression={$wvdHostPool.HostPoolType}}, 
                        @{Name='HostPoolLoadBalancerType'; Expression={$wvdHostPool.LoadBalancerType}}, 
                        @{Name='HostpoolLocation'; Expression={$wvdHostPool.Location}},
                        @{Name='HostpoolMaxSessionLimit'; Expression={$wvdHostPool.MaxSessionLimit}},
                        @{Name='HostpoolResourceGroup'; Expression={$_.Id.Split("/")[4]}}, 
                        @{Name='HostPoolSsoContext'; Expression={$wvdHostPool.SsoContext}}, 
                        @{Name='HostPoolVMTemplate'; Expression={$wvdHostPool.VMTemplate}}, 
                        @{Name='HostPoolValidationEnvironment'; Expression={$wvdHostPool.ValidationEnvironment}}
                } 
            }
        }
    }
    return $psobjectCollection
}

function Set-WVDSessionHostDrainMode () {
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Session Host Name'
        )]
        [string] $SessionHostName,
    
        [Parameter(
            Position=1, 
            Mandatory=$false, 
            HelpMessage='Enter the Host Pool name'
        )]
        [string] $HostPoolName,

        [Parameter(
            Position=2, 
            Mandatory=$false, 
            HelpMessage='Enter the Resource Group name'
        )]
        [string] $ResourceGroupName,

        [Parameter(
            Position=2, 
            Mandatory=$false, 
            HelpMessage='Enter the Drain Modee'
        )]
        [ValidateSet("ON","OFF")]
        [string] $DrainMode
    )

    # Translate Drain Mode to boolean value
    $boolDrainMode = $false
    If ($DrainMode -eq "OFF")
    {
        $boolDrainMode = $true
    }

    #Update the WVD Session Host Drain Mode (AllowNewSession parameter)
    try
    {
        Update-AzWvdSessionHost -HostPoolName $HostPoolName -ResourceGroupName $ResourceGroupName -Name $SessionHostName -AllowNewSession:$boolDrainMode | Out-Null

    }
    catch
    {
        Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
        Exit
    }
}

function Set-WVDHostpoolMaxSessionLimit () {
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Hostpool Name'
        )]
        [string] $HostpoolName,
    
        [Parameter(
            Position=1, 
            Mandatory=$true, 
            HelpMessage='Enter the Resource Group name'
        )]
        [string] $ResourceGroupName,

        [Parameter(
            Position=2, 
            Mandatory=$false, 
            HelpMessage='Enter the max session limit'
        )]
        [int] $MaxSessionLimit
    )

    #Update the WVD Hostpool Max Session Limit (AllowNewSession parameter)
    try
    {
        Update-AzWvdHostPool -Name $HostPoolName -ResourceGroupName $ResourceGroupName -MaxSessionLimit $MaxSessionLimit | Out-Null
    }
    catch
    {
        Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
        Exit
    }
}

function Set-WVDHostpoolLoadBalancerType () {
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Hostpool Name'
        )]
        [string] $HostpoolName,
    
        [Parameter(
            Position=1, 
            Mandatory=$true, 
            HelpMessage='Enter the Resource Group name'
        )]
        [string] $ResourceGroupName,

        [Parameter(
            Position=2, 
            Mandatory=$false, 
            HelpMessage='Enter the Load Balancer Type'
        )]
        [ValidateSet("BreadthFirst","DepthFirst")]
        [string] $LoadBalancerType
    )

    #Update the WVD Hostpool LoadBalancerType
    try
    {
        Update-AzWvdHostPool -Name $HostPoolName -ResourceGroupName $ResourceGroupName -LoadBalancerType $LoadBalancerType | Out-Null
    }
    catch
    {
        Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
        Exit
    }
}

function Get-WVDHostpoolRegistrationToken () {
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Hostpool Name'
        )]
        [string] $hostpoolName,
    
        [Parameter(
            Position=1, 
            Mandatory=$true, 
            HelpMessage='Enter the Resource Group name'
        )]
        [string] $resourceGroupName,

        [Parameter(
            Position=1, 
            Mandatory=$true, 
            HelpMessage='Enter the subscriptionID'
        )]
        [string] $subscriptionID
    )

    #Retrieve an existing Host Pool registration token
    $registrationInfo = $null
    try
    {
        $registrationInfo = Get-AzWvdRegistrationInfo -ResourceGroupName $resourceGroupName -HostPoolName $hostpoolName -SubscriptionId $subscriptionID # | Out-Null
    }
    catch
    {
        Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
        Exit
    }
    If ($registrationInfo.Count -gt 0)
    {
        return $($registrationInfo.Token)
    }
}

function Connect-WVDServicePrincipalAccount () {
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Service Principal credentials'
        )]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCredential] $SPCredentials,

        [Parameter(
            Position=1, 
            Mandatory=$true, 
            HelpMessage='Enter the Tenant ID'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $TenantID
    )

    # Connect to Azure with the WVD Service Principal credentials
    try
    {
        $wvdAzSession = Connect-AzAccount -Credential $SPCredentials -Tenant $TenantID -ServicePrincipal -WarningAction SilentlyContinue # | Out-Null
    }
    catch
    {
        Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
        Exit
    }
    return $($wvdAzSession)
}

function Get-WVDHostPool () {
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Subscription ID'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $SubscriptionID,

        [Parameter(
            Position=1, 
            Mandatory=$false, 
            HelpMessage='Enter the Hostpool Name'
        )]
        [string] $HostPoolName
    )

    # Check if the HostPoolName parameter is empty or not
    If ([string]::IsNullOrEmpty($HostPoolName))
    {
        $HostPoolName = ""
    }

    # Retrieve the HostPool information for the given Subscription ID (and optional HostPool Name)
    try
    {
        #Retrieve the WVD Host Pool information for this Subscription
        [array]$wvdHostPoolCollection = Get-AzWvdHostPool -SubscriptionId $SubscriptionID -WarningAction SilentlyContinue | 
        ForEach-Object {
        $wvdHostPool = $_

            If ($wvdHostPool.Name -like "$HostPoolName*")
            {
            # Retrieve the Host Pool information
            $wvdHostPool | Select Name, 
                ApplicationGroupReference, 
                @{Name='ApplicationGroupsCount'; Expression={$_.ApplicationGroupReference.Count}}, 
                CustomRdpProperty, 
                Description, 
                FriendlyName, 
                HostPoolType, 
                Id, 
                LoadBalancerType, 
                Location,
                MaxSessionLimit,
                PersonalDesktopAssignmentType, 
                RegistrationInfoExpirationTime, 
                RegistrationInfoRegistrationTokenOperation, 
                RegistrationInfoToken, 
                Ring, 
                SsoContext, 
                VMTemplate, 
                ValidationEnvironment, 
                @{Name='ResourceGroup'; Expression={$_.Id.Split("/")[4]}} 
            }
        }
    }
    catch
    {
        Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
        Exit
    }
    return $wvdHostPoolCollection
}

function Update-WVDHostpool () {
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Hostpool Name'
        )]
        [string] $HostpoolName,
    
        [Parameter(
            Position=1, 
            Mandatory=$true, 
            HelpMessage='Enter the Resource Group name'
        )]
        [string] $ResourceGroupName,

        [Parameter(
            Position=2, 
            Mandatory=$false, 
            HelpMessage='Select whether the Hostpool needs validation or not'
        )]
        [ValidateSet("Yes", "No")]
        [string] $ValidationEnvironment,

        [Parameter(
            Position=3, 
            Mandatory=$false, 
            HelpMessage='Enter a friendly name for the Hostpool'
        )]
        [string] $FriendlyName,

        [Parameter(
            Position=4, 
            Mandatory=$false, 
            HelpMessage='Enter a description for the Hostpool'
        )]
        [string] $Description,

        [Parameter(
            Position=5, 
            Mandatory=$false, 
            HelpMessage='Enter the max session limit'
        )]
        [int] $MaxSessionLimit,

        [Parameter(
            Position=6, 
            Mandatory=$false, 
            HelpMessage='Enter the load balancing algorith'
        )]
        [ValidateSet("BreadthFirst","DepthFirst")]
        [string] $LoadBalancingAlgorithm,

        [Parameter(
            Position=7, 
            Mandatory=$false, 
            HelpMessage='Enter the RDP properties as a semi-colon separated list'
        )]
        [string] $RDPProperty

    )

    # Configure the Hostpool parameters to be updated
    $objParams = @{
        "Name"=$HostpoolName;
        "ResourceGroupName"=$ResourceGroupName;
    }

    #Update the WVD Hostpool Max Session Limit (AllowNewSession parameter)
    If ($ValidationEnvironment)
    {
        $boolValidation = $false
        If ($ValidationEnvironment -eq "Yes") 
        {
            $boolValidation = $true
        }
        $objParams.Add("ValidationEnvironment",$boolValidation)
    }

    If (-not([string]::IsNullOrWhiteSpace($FriendlyName)))
    {
        $objParams.Add("FriendlyName",$FriendlyName)
    }

    If (-not([string]::IsNullOrWhiteSpace($Description)))
    {
        $objParams.Add("Description",$Description)
    }

    If ($MaxSessionLimit)
    {
        $objParams.Add("MaxSessionLimit",$MaxSessionLimit)
    }

    If (-not([string]::IsNullOrWhiteSpace($LoadBalancingAlgorithm)))
    {
        $objParams.Add("LoadBalancerType",$LoadBalancingAlgorithm)
    }

    If ($RDPProperty)
    {
        $objParams.Add("CustomRdpProperty",$RDPProperty)
    }

    #debug: $objParams

    try
    {
        Update-AzWvdHostPool @objParams | Out-Null
    }
    catch
    {
        Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
        Exit
    }
}

function Get-WVDHostPoolSessionHost () {
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Subscription ID'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $SubscriptionID,

        [Parameter(
            Position=1, 
            Mandatory=$false, 
            HelpMessage='Enter the Hostpool Name'
        )]
        [string] $HostPoolName,

        [Parameter(
            Position=2, 
            Mandatory=$false, 
            HelpMessage='Enter the Session Host Name'
        )]
        [string] $SessionHostName
    )

    # Check if the HostPoolName parameter is empty or not
    If ([string]::IsNullOrEmpty($HostPoolName))
    {
        $HostPoolName = ""
    }

    # Check if the SessionHostName parameter is empty or not
    If ([string]::IsNullOrEmpty($SessionHostName))
    {
        $SessionHostName = ""
    }
    
    # Retrieve the SessionHost information for the given Subscription ID (and optional HostPoolName and SessionHostName)
    try
    {
        #Retrieve the WVD SessionHost information for this Subscription
        [array]$wvdSessionHostCollection = Get-AzWvdHostPool -SubscriptionId $SubscriptionID -WarningAction SilentlyContinue | 
        foreach {
        $wvdHostPool = $_

            If ($wvdHostPool.Name -like "$HostPoolName*")
            {
                # Retrieve the SessionHost information
                Get-AzWvdSessionHost -HostPoolName $($wvdHostPool.Name) -ResourceGroupName $($wvdHostPool.Id.Split("/")[4]) |
                foreach {
                    $wvdSessionHost = $_
                    If ($wvdSessionHost.Name.Split("/")[1] -like "$SessionHostName*")
                    {
                        # Retrieve SessionHost and HostPool information for each SessionHost
                        $wvdSessionHost | Select @{Name='HostPoolName'; Expression={$wvdHostPool.Name}}, 
                        @{Name='SessionHostName'; Expression={$_.Name.Split("/")[1]}}, 
                        AgentVersion, 
                        AllowNewSession, 
                        @{Name='DrainMode'; Expression={if($_.AllowNewSession -eq "True"){return "Off"}else{return "On"}}}, 
                        AsignedUser, 
                        LastHeartBeat,
                        LastUpdateTime,
                        OSVersion,
                        Session, 
                        @{Name='ActiveSessions'; Expression={$_.Session}},  
                        Status,
                        StatusTimestamp,
                        SxSStackVersion, 
                        Type, 
                        UpdateErrorMessage, 
                        UpdateState, 
                        @{Name='HostPoolApplicationGroupReference'; Expression={$wvdHostPool.ApplicationGroupReference}}, 
                        @{Name='HostPoolCustomRdpProperty'; Expression={$wvdHostPool.CustomRdpProperty}}, 
                        @{Name='HostPoolDescription'; Expression={$wvdHostPool.Description}}, 
                        @{Name='HostPoolFriendlyName'; Expression={$wvdHostPool.FriendlyName}}, 
                        @{Name='HostPoolType'; Expression={$wvdHostPool.HostPoolType}}, 
                        @{Name='HostPoolLoadBalancerType'; Expression={$wvdHostPool.LoadBalancerType}}, 
                        @{Name='HostpoolLocation'; Expression={$wvdHostPool.Location}},
                        @{Name='HostpoolMaxSessionLimit'; Expression={$wvdHostPool.MaxSessionLimit}},
                        @{Name='HostpoolResourceGroup'; Expression={$_.Id.Split("/")[4]}}, 
                        @{Name='HostPoolSsoContext'; Expression={$wvdHostPool.SsoContext}}, 
                        @{Name='HostPoolVMTemplate'; Expression={$wvdHostPool.VMTemplate}}, 
                        @{Name='HostPoolValidationEnvironment'; Expression={$wvdHostPool.ValidationEnvironment}}
                    }
                }
            }
        }
    }
    catch
    {
        Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
        Exit
    }
    return $wvdSessionHostCollection
}

function Get-WVDSessionHostHostPool () {
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter the Subscription ID'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $SubscriptionID,

        [Parameter(
            Position=1, 
            Mandatory=$true, 
            HelpMessage='Enter the Session Host Name'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $SessionHostName
    )

    # Retrieve the HostPool Id from the given Subscription ID and SessionHostName
    try
    {
        #Retrieve the WVD SessionHost information for this Subscription
        [array]$wvdSessionHostHostpoolCollection = Get-AzWvdHostPool -SubscriptionId $SubscriptionID -WarningAction SilentlyContinue | 
        foreach {
        $wvdHostPool = $_
            # Retrieve the SessionHost information
            Get-AzWvdSessionHost -HostPoolName $($wvdHostPool.Name) -ResourceGroupName $($wvdHostPool.Id.Split("/")[4]) |
            foreach {
                $wvdSessionHost = $_
                If ($wvdSessionHost.Name.Split("/")[1] -like "$SessionHostName*")
                {
                    # Retrieve HostPool Id information for each SessionHost
                    $wvdSessionHost | Select @{Name='SessionHostName'; Expression={$wvdSessionHost.Name.Split("/")[1]}}, 
                    @{Name='SessionHostId'; Expression={$wvdSessionHost.Id}}, 
                    @{Name='HostPoolName'; Expression={$wvdHostPool.Name}}, 
                    @{Name='HostPoolId'; Expression={$wvdHostPool.Id}}, 
                    @{Name='HostpoolResourceGroup'; Expression={$_.Id.Split("/")[4]}} 
                }
            }
        }
    }
    catch
    {
        Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
        Exit
    }
    return $wvdSessionHostHostpoolCollection
}

# dot sourcing ControlUp Script Action settings
#region ControlUp Script Standards - version 0.2
    #Requires -Version 5.1

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

#------------------------#
# Script Action Workflow #
#------------------------#
Write-Host ""

#region Retrieve input parameters
#endregion

## Check if the required PowerShell Modules are installed and can be imported
Invoke-CheckInstallAndImportPSModulePrereq -ModuleName "Az.Accounts" #-Verbose
Invoke-CheckInstallAndImportPSModulePrereq -ModuleName "Az.DesktopVirtualization" #-Verbose

## Testing Script output
If (Invoke-NETFrameworkCheck)
{
    If ($azSPCredentials = Get-AzSPStoredCredentials)
    {
        # Sign in to Azure with a Service Principal with Contributor Role at Subscription level
        try
        {
            $azSPSession = Connect-WVDServicePrincipalAccount -SPCredentials $azSPCredentials.spCreds -TenantID $($azSPCredentials.tenantID).ToString()
        }
        catch
        {
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }

        # Retrieve the Subscription information for the Service Principal (that is logged on)
        $azSubscription = Get-AzSubscription

        try 
        {
            # Translate Session Host Name to HostPool ID (ControlUp WVD Script Action Integration specific function):
            $Id = (Get-WVDSessionHostHostPool -SubscriptionID $($azSubscription.Id) -SessionHostName $SessionHostName).SessionHostId
            # Translate ID to SessionHost Name
            $wvdSessionHostName = $Id.Replace("ID=","").Split("/")[10]
            If ([string]::IsNullOrEmpty($wvdSessionHostName))
            {
                $wvdSessionHostName = ""
            }
        }
        catch 
        {
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        try 
        {
            #Retrieve the WVD Session Host information, based on Name
            [array]$sessionHosts = Get-WVDHostpoolSessionHost -SubscriptionID $($azSubscription.Id) -SessionHostName $wvdSessionHostName
        }
        catch
        {
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }
        If ($sessionHosts.Count -gt 0)
        {
            ForEach ($sessionHost in $sessionHosts)
            {
                Set-WVDSessionHostDrainMode -SessionHostName $($sessionHost.SessionHostName) -HostPoolName $($sessionHost.HostPoolName) -ResourceGroupName $($sessionHost.HostpoolResourceGroup) -DrainMode $DrainMode
                Write-Host ("SessionHost '$($sessionHost.SessionHostName.Split(".")[0])' drain mode set: ") -ForegroundColor Yellow -NoNewline
                Write-Host $($DrainMode) -ForegroundColor Cyan
            }
        }
    }
    else 
    {
        Write-Warning "No Azure Credentials could be retrieved from the stored credentials file for this user."
    }
}
Else
{
    Write-Warning "The Az PowerShell Module, used by this Script Action, requires a minimum version of .NET Framework version 4.7.2. Please upgrade the .NET Framework version on this machine"
}
    
# Disconnect the Azure Session
Disconnect-AzAccount | Out-Null
