<#
.SYNOPSIS
    Get Azure Cost Analysis information, sorted by Name.
.DESCRIPTION
    Get Azure Cost Analysis information, using REST API calls.
.EXAMPLE
    Get-AzCostAnalysis
.EXAMPLE
    Get-AzCostAnalysis -SubscriptionID <string>
.CONTEXT
    Windows Virtual Desktops
.NOTES
    Version:        0.1
    Author:         Esther Barthel, MSc
    Creation Date:  2020-10-25
    Updated:        2020-10-25

    Purpose:        WVD Administration, through REST API calls
        
    Copyright (c) cognition IT. All rights reserved.
#>
[CmdletBinding()]
Param()    

# ------------------------------------
# | WVD Functions for REST API calls |
# ------------------------------------

#region Global Variables
$wvdApiVersion = "2019-12-10-preview"
#endregion Global Variables

#region Windows Presentation Foundation (WPF) form to store WVD Service Principal information
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

function Invoke-WVDSPCredentialsForm {
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
#endregion WPF form

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
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-08-03
        Purpose:        WVD Administration, through REST API calls
        
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

function Get-AzBearerToken {
    <#
    .SYNOPSIS
        Retrieve the Azure Bearer Token for an authentication session.
    .DESCRIPTION
        Retrieve the Azure Bearer Token for an authentication session, using a REST API call.
    .EXAMPLE
        Get-AzBearerToken -SPCredentials <PSCredentialObject> -TenantID <string>
    .CONTEXT
        Azure
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-03-22
        Updated:        2020-05-08
                        Created a separate Azure Credentials function to support ARM architecture and REST API scripted actions
        Purpose:        WVD Administration, through REST API calls
        
        Copyright (c) cognition IT. All rights reserved.
    #>
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

    #region Prep variables
        # URL for REST API call to authenticate with Azure (using the TenantID parameter)
        $uri = "https://login.microsoftonline.com/$TenantID/oauth2/token"
        
        # Create the Invoke-RestMethod Body (using the SPCredentials parameter)
        $body = @{
            grant_type="client_credentials";
            client_Id=$($SPCredentials.UserName);
            client_Secret=$($SPCredentials.GetNetworkCredential().Password);
            resource="https://management.azure.com"
        }
        #debug: $body

        # Create the Invoke-RestMethod parameters
        $invokeRestMethodParams = @{
            Uri             = $uri
            Body            = $body
            Method          = "POST"
            ContentType     = "application/x-www-form-urlencoded"
        }
    #endregion

    try 
    {
        $response = $null
        # Make the REST API call with the created parameters
        $response = Invoke-RestMethod @invokeRestMethodParams
    }
    catch 
    {
        Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
    }
    # return the JSON response
    return $response
}

function Get-AzSubscription {
    <#
    .SYNOPSIS
        Retrieve the Azure Subscription information.
    .DESCRIPTION
        Retrieve the Azure Subscription information, using a REST API call.
    .EXAMPLE
        Get-AzSubscription -BearerToken <string> -SubscriptionID <string>
    .CONTEXT
        Azure
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-09-20
        Updated:        2020-09-20
                        Created a separate Azure Credentials function to support ARM architecture and REST API scripted actions

        Purpose:        WVD Administration, through REST API calls
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter a valid bearer token'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $BearerToken,

        [Parameter(
            Position=1, 
            Mandatory=$false, 
            HelpMessage='Enter a subscriptionID'
        )]
        [string] $SubscriptionID
    )

    #region Prep variables
        # URL for REST API call to authenticate with Azure (using the TenantID parameter)
        $uri = "https://management.azure.com/subscriptions?api-version=2020-01-01"

        # Create the Invoke-RestMethod Header (using the bearertoken parameter)
        $header = @{
            "Authorization"="Bearer $BearerToken"; 
            "Content-Type" = "application/json"
        }
        #debug: $header

        # Create the Invoke-RestMethod parameters
        $invokeRestMethodParams = @{
            Uri             = $uri
            Method          = "GET"
            Headers          = $header
        }
        #debug: $invokeRestMethodParams
    #endregion

    try 
    {
        $response = $null
        # Make the REST API call with the created parameters
        $response = Invoke-RestMethod @invokeRestMethodParams
    }
    catch 
    {
        Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
    }
    # filter the response if a SubscriptionID was provided
    If (!([string]::IsNullOrEmpty($subScriptionID)))
    {
        $results = ($response.value).Where({$_.subscriptionId -like "$subScriptionID"})
    }
    else 
    {
        $results = $response.value
    }
    return $results
}

function Get-AzCostAnalysisActualDaily () {
    <#
    .SYNOPSIS
        Get Azure Cost Analysis information on the actual daily costs for this month.
    .DESCRIPTION
        Get Azure Cost Analysis information on the actual daily costs for this month, using a REST API call.
    .EXAMPLE
        Get-AzCostAnalysisActualDaily -BearerToken <string> -SubscriptionID <string>
    .CONTEXT
        Azure
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-10-25
        Updated:        2020-10-25
                        Created a separate Azure Cost Analysis function for the costs of this month, sorted by Resource Group

        Purpose:        WVD Administration, through REST API calls
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter a valid bearer token'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $BearerToken,

        [Parameter(
            Position=1, 
            Mandatory=$true, 
            HelpMessage='Enter the Subscription ID'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $SubscriptionID
    )
    #region Prep variables
        # URL for REST API call to list hostpools, based on given subscription ID
        $uri = "https://management.azure.com/subscriptions/$SubscriptionID/providers/Microsoft.CostManagement/query`?api-version=2019-11-01"#&`$top=5000"

        # Create the Invoke-RestMethod Header (using the bearertoken parameter)
        $header = @{
            "Authorization"="Bearer $BearerToken"; 
            "Content-Type" = "application/json"
        }
        #debug: $header

        # Create Custom time period for this month
        #$time = (Get-Date).AddMonths(-1).ToString("yyyy-MM")
        $time = $((Get-Date).AddMonths(0).ToString("yyyy-MM"))
        $lastdayofmonth = "$([System.DateTime]::DaysInMonth((Get-Date).AddMonths(0).Year,(Get-Date).AddMonths(0).Month))"
        # Create the JSON formatted body
        $body=@{
            "type"= "ActualCost";
            "dataSet"= @{
                "granularity"= "Daily";
                "aggregation"= @{
                    "totalCost"= @{"name"= "Cost";"function"= "Sum"}
                };
                "sorting"=@(
                    @{"direction"="ascending";"name"="UsageDate"}
                )
            };
            "timeframe"="Custom";
            "timePeriod"=@{"from"="$time`-01T00:00:00+00:00";"to"="$time`-$lastdayofmonth`T23:59:59+00:00"}
        }
        $bodyJSON = ConvertTo-Json -InputObject $body -Depth 10

        # Create the Invoke-RestMethod parameters
        $invokeRestMethodParams = @{
            Uri             = $uri
            Method          = "POST"
            Headers         = $header
            Body            = $bodyJSON
        }
        #debug: $invokeRestMethodParams
    #endregion

    try 
    {
        $response = $null
        # Make the REST API call with the created parameters
        $response = Invoke-RestMethod @invokeRestMethodParams
    }
    catch 
    {
        Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
    }
    #debug: $response
    $results = $response.properties
    return $results
}

function Get-AzCostAnalysisForecastDaily () {
    <#
    .SYNOPSIS
        Get Azure Cost Analysis information on the actual daily forcast for this month.
    .DESCRIPTION
        Get Azure Cost Analysis information on the actual daily forecast for this month, using a REST API call.
    .EXAMPLE
        Get-AzCostAnalysisForecastDaily -BearerToken <string> -SubscriptionID <string>
    .CONTEXT
        Azure
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-10-25
        Updated:        2020-10-25
                        Created a separate Azure Cost Analysis function for the costs of this month, sorted by Resource Group

        Purpose:        WVD Administration, through REST API calls
        
        Copyright (c) cognition IT. All rights reserved.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter a valid bearer token'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $BearerToken,

        [Parameter(
            Position=1, 
            Mandatory=$true, 
            HelpMessage='Enter the Subscription ID'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $SubscriptionID
    )
    #region Prep variables
        # URL for REST API call to list hostpools, based on given subscription ID
        $uri = "https://management.azure.com/subscriptions/$SubscriptionID/providers/Microsoft.CostManagement/forecast`?api-version=2019-11-01"#&`$top=5000"

        # Create the Invoke-RestMethod Header (using the bearertoken parameter)
        $header = @{
            "Authorization"="Bearer $BearerToken"; 
            "Content-Type" = "application/json"
        }
        #debug: $header

        # Create Custom time period for this month
        $months = 0
        $time = $((Get-Date).AddMonths($months).ToString("yyyy-MM"))
        $lastdayofmonth = "$([System.DateTime]::DaysInMonth((Get-Date).AddMonths($months).Year,(Get-Date).AddMonths($months).Month))"

        # Create the JSON formatted body
        $body=@{
            "type"= "ActualCost";
            "dataSet"= @{
                "granularity"= "Daily";
                "aggregation"= @{
                    "totalCost"= @{"name"= "Cost";"function"= "Sum"}
                };
                "sorting"=@(
                    @{"direction"="ascending";"name"="UsageDate"}
                )
            };
            "timeframe"="Custom";
            "timePeriod"=@{"from"="$time`-01T00:00:00+00:00";"to"="$time`-$lastdayofmonth`T23:59:59+00:00"};
            "includeActualCost"="false";
            "includeFreshPartialCost"="false"
        }
        $bodyJSON = ConvertTo-Json -InputObject $body -Depth 10

        # Create the Invoke-RestMethod parameters
        $invokeRestMethodParams = @{
            Uri             = $uri
            Method          = "POST"
            Headers         = $header
            Body            = $bodyJSON
        }
        #debug: $invokeRestMethodParams
    #endregion

    If ((Get-Date).Day -lt $lastdayofmonth)
    # forecast data should be available
    {
        try 
        {
            $response = $null
            # Make the REST API call with the created parameters
            $response = Invoke-RestMethod @invokeRestMethodParams
        }
        catch 
        {
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
        }
        #debug: $response
        $results = $response.properties
        return $results
    }
    else
    {
        Write-Warning "No forecast data available"
        return $null
    }
}


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



#-----------------#
# Script Workflow #
#-----------------#
Write-Output ""

# Script output
If ($azSPCredentials = Get-AzSPStoredCredentials)
{
    #debug: $azSPCredentials
    # Sign in to Azure with a Service Principal with Contributor Role at Subscription level and retrieve the brearer token
    try
    {
        $azBearerToken = $null
        $azBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $($azSPCredentials.tenantID).ToString()
        #debug: $azBearerToken
    }
    catch
    {
        Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
        Exit
    }

    # Retrieve the Subscription information for the Service Principal (that is logged on)
    $azSubscription = $null
    $azSubscription = Get-AzSubscription -bearerToken $($azBearerToken.access_token)
    #debug: Write-Output "DEBUG INFO - subscriptionID: $($azSubscription.subscriptionId)"

    # Retrieve the Cost Analysis - Actual Costs details
    $costAnalysisResults = Get-AzCostAnalysisActualDaily -bearerToken $($azBearerToken.access_token) -SubscriptionID $($azSubscription.subscriptionId.Split("/")[-1]) #-ResourceGroupName $ResourceGroupName

    # Build output dataset, based on custom object
    $dataResults = @()
    for ($row=0;$row -le (($costAnalysisResults.rows.Count)-1); $row++)
    {
        $htrow = New-Object psobject
        for ($col=0;$col -le (($costAnalysisResults.columns.Count)-1); $col++)
        {
            $htrow | Add-Member -Type NoteProperty -Name "$($costAnalysisResults.columns[$col].name)" -Value "$($costAnalysisResults.rows[$row][$col])"
        }
        $dataResults += $htrow
    }
    # Present the Cost information
    Write-Host "* Actual costs for this month, sorted by Day: " -ForegroundColor Yellow
    $totalCosts = $([math]::Round((($dataResults | Measure-Object -Property Cost -Sum).Sum),2))
    If ($(($dataResults | Measure-Object).Count) -ge 1)
    {
        $dataResults | Select Currency, 
            Cost, 
            @{Name='Date'; Expression={$($_.UsageDate)}}, 
            @{Name='Costs'; Expression={$([math]::Round($_.Cost,2))}}, 
            @{Name='Percentage'; Expression={$([math]::Round((($_.Cost/$totalCosts)*100),0))}} | `
                Sort Date | `
                Format-Table @{Name='Date    '; Expression={"$($_.Date)"};Align="left"}, 
                    @{Name='Costs        '; Expression={"$($_.Currency){0,10:N2}" -f($($_.Costs))}; Align="right"}, 
                    @{Name='Percentage'; Expression={"{0:N0} %" -f($($_.Percentage))}; Align="right"} -AutoSize

            #@{Name='Costs'; Expression={$_.Currency + " {0,8:N2}" -f($([math]::Round($_.Cost,2)))}} | Sort UsageDate | Format-Table UsageDate, Costs, @{Name='CostStatus'; Expression={"ActualCosts"}} -AutoSize
    }
    Else
    {
        Write-Output ""
        Write-Warning "No usage or purchases reported during this period"
        Write-Output ""
    }

    # Present forecast information too
    # Retrieve the Cost Analysis - Forecast details
    $forecastAnalysisResults = Get-AzCostAnalysisForecastDaily -bearerToken $($azBearerToken.access_token) -SubscriptionID $($azSubscription.subscriptionId.Split("/")[-1]) #-ResourceGroupName $ResourceGroupName

    # Build output dataset, based on custom object
    $forecastResults = @()
    for ($row=0;$row -le (($forecastAnalysisResults.rows.Count)-1); $row++)
    {
        $fhtrow = New-Object psobject
        for ($col=0;$col -le (($forecastAnalysisResults.columns.Count)-1); $col++)
        {
            $fhtrow | Add-Member -Type NoteProperty -Name "$($forecastAnalysisResults.columns[$col].name)" -Value "$($forecastAnalysisResults.rows[$row][$col])"
        }
        $forecastResults += $fhtrow
    }
    # Present the Forecast information
    Write-Host "* Forecast of the costs for the remainder of this month, sorted by Day: " -ForegroundColor Yellow
    $totalForecasts = $([math]::Round((($forecastResults | Measure-Object -Property Cost -Sum).Sum),2))
    $forecastResults | Select Currency, 
        Cost, 
        UsageDate, 
        CostStatus, 
        @{Name='Costs'; Expression={$_.Currency + " {0,8:N2}" -f($([math]::Round($_.Cost,2)))}}, 
        @{Name='Percentage'; Expression={$([math]::Round((($_.Cost/$totalCosts)*100),0))}} | 
            Sort UsageDate | 
            Format-Table @{Name='Date    '; Expression={"$($_.UsageDate)"};Align="left"}, 
                @{Name='Costs       '; Expression={"$($_.Currency) {0,8:N2}" -f($([math]::Round($_.Cost,2)))};Align="left"}, 
                @{Name='Cost status'; Expression={"$($_.CostStatus)"};Align="right"} -AutoSize

    # Present the Summary information
    Write-Host "* Costs summary for this month: " -ForegroundColor Yellow
    Write-Host "  - Actual Costs (incl. today):       " -ForegroundColor Cyan -NoNewline
    Write-Host ( "$($dataResults[0].Currency) {0:N2} " -f($([math]::Round((($dataResults | Measure-Object -Property Cost -Sum).Sum),2))) ) -ForegroundColor Yellow
    Write-Host "  - Predicted Costs (incl. Forecast): " -ForegroundColor Cyan -NoNewline
    Write-Host ( "$($forecastResults[0].Currency) {0:N2} " -f($([math]::Round( (($forecastResults | Measure-Object -Property Cost -Sum).Sum) + (($dataResults | Measure-Object -Property Cost -Sum).Sum),2))) ) -ForegroundColor Yellow -NoNewline
    Write-Host ( "(`+ " + "{0:N2}" -f($([math]::Round((($forecastResults | Measure-Object -Property Cost -Sum).Sum),2))) + ")" ) -ForegroundColor Yellow
    Write-Output ""
}
else 
{
    Write-Warning "No Azure Credentials could be retrieved from the stored credentials file for this user."
}

# SIG # Begin signature block
# MIINHAYJKoZIhvcNAQcCoIINDTCCDQkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU53CGNDVQEyRhUVvNUvri8s32
# x8SgggpeMIIFJjCCBA6gAwIBAgIQCyXBE0rAWScxh3bGfykLTjANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTIwMDgxMDAwMDAwMFoXDTIzMDgx
# NTEyMDAwMFowYzELMAkGA1UEBhMCTkwxDzANBgNVBAcTBkxlbW1lcjEVMBMGA1UE
# ChMMY29nbml0aW9uIElUMRUwEwYDVQQLEwxDb2RlIFNpZ25pbmcxFTATBgNVBAMT
# DGNvZ25pdGlvbiBJVDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAL2y
# YRbztz9wTtatpSJ5NMD0JZtDPhuFqAVTjoGc1jvn68M41zlGgi8fVvEccaH3nDTT
# 6T8edgFuEbsZVHZGmY109zHOPwXX+Zvp3T+Hk2Ys8Liwwirr6xw9dlneBu85j8gd
# Mamz+mNjzpyBg1eVlD7cV1JAL3oAXgONRiebdpD6DPvd3melPmeg84Un3VV6+W8M
# 8Y0Pec+TbxIda18Lr4DqnIl0a/Suk8kQ2DzZXDXoK+MCfA6zsqyEOSY5yI5OwdU0
# 93LC2PHFEKEkIogBlCiD0UQDbamPdu7wZnTAHPTDfifdMhCPLBA0y4pj4jm6ggFE
# 3ZuQMR/yU8JXSwy72ZECAwEAAaOCAcUwggHBMB8GA1UdIwQYMBaAFFrEuXsqCqOl
# 6nEDwGD5LfZldQ5YMB0GA1UdDgQWBBR2SeoVDh3RxGqV5iamn/FFU6J65zAOBgNV
# HQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1oDOg
# MYYvaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5j
# cmwwNaAzoDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQt
# Y3MtZzEuY3JsMEwGA1UdIARFMEMwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUHAgEW
# HGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCAYGZ4EMAQQBMIGEBggrBgEF
# BQcBAQR4MHYwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBO
# BggrBgEFBQcwAoZCaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# U0hBMkFzc3VyZWRJRENvZGVTaWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQCMAAwDQYJ
# KoZIhvcNAQELBQADggEBAL4Zda674x5WLL8B059a9cxnUIC05LcjD/3hkCLZgbMa
# krDrfsjNpA+KpMiTv2TW5pDRCXGJirJO27XRTojr2F8+gJAyIB+8ZLiyKmy3IcCV
# DXjjb6i/4TiGbDmGL3Ctl5pmWRpksnr3TKSMyxz2OogLS6w9pgRdA1hgJSfZMV+a
# KRrd4iW5YWKIwFZlYDeQqpBBtQ6ujzgQ/04FcmjyOlNch4hofJVLauzkSb1Tnzt1
# 6TyT2pJ9BzasoOlxYEFhn0ikXndlKVBb7gpFInqSf5DJtaVRIXojj0eqN6LZroUz
# 62m2YeR29uC06xcdF7fjo+YKxe+kdApdPfX0Nx9Moc8wggUwMIIEGKADAgECAhAE
# CRgbX9W7ZnVTQ7VvlVAIMA0GCSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVTMRUw
# EwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20x
# JDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMzEwMjIx
# MjAwMDBaFw0yODEwMjIxMjAwMDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMT
# KERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZz9D7RZmx
# OttE9X/lqJ3bMtdx6nadBS63j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnkoOn7p0WfT
# xvspJ8fTeyOU5JEjlpB3gvmhhCNmElQzUHSxKCa7JGnCwlLyFGeKiUXULaGj6Ygs
# IJWuHEqHCN8M9eJNYBi+qsSyrnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q8grkV7tK
# tel05iv+bMt+dDk2DZDv5LVOpKnqagqrhPOsZ061xPeM0SAlI+sIZD5SlsHyDxL0
# xY4PwaLoLFH3c7y9hbFig3NBggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPRAgMBAAGj
# ggHNMIIByTASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNV
# HSUEDDAKBggrBgEFBQcDAzB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0
# dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2Vy
# dHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYD
# VR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBPBgNVHSAESDBGMDgGCmCG
# SAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29t
# L0NQUzAKBghghkgBhv1sAzAdBgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt9mV1Dlgw
# HwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQELBQAD
# ggEBAD7sDVoks/Mi0RXILHwlKXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L/e8q3yBV
# N7Dh9tGSdQ9RtG6ljlriXiSBThCk7j9xjmMOE0ut119EefM2FAaK95xGTlz/kLEb
# Bw6RFfu6r7VRwo0kriTGxycqoSkoGjpxKAI8LpGjwCUR4pwUR6F6aGivm6dcIFzZ
# cbEMj7uo+MUSaJ/PQMtARKUT8OZkDCUIQjKyNookAv4vcn4c10lFluhZHen6dGRr
# sutmQ9qzsIzV6Q3d9gEgzpkxYz0IGhizgZtPxpMQBvwHgfqL2vmCSfdibqFT+hKU
# GIUukpHqaGxEMrJmoecYpJpkUe8xggIoMIICJAIBATCBhjByMQswCQYDVQQGEwJV
# UzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQu
# Y29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWdu
# aW5nIENBAhALJcETSsBZJzGHdsZ/KQtOMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBRPO1tUccm7
# JVysT4Do2GutKWJqJjANBgkqhkiG9w0BAQEFAASCAQBy1kE7n/8z65ac8EZCZ4S8
# qcob7TNY3I6R+usJsAqMbY5LWMWd4Ng0wLarGIRqOzh9S57lLgnYdEMS1QhZ5v3N
# gSAwsQ87Xy7v7CELeeZKEsbS5f5bipGeNd8NeQnj5hCEPSE5l4XrsamAZoTWGKvE
# tlXWTEatv10v4uJ2ZfqIjMQls5hALvjkwtjAM7Uuvt15pgpaeRFQdWRNsIRMbvg/
# ToDp/umLVgLG1/enRyA1O/OCvfcorZtUPD06ffaQBpchqeBytpyzpJQVor1mU9k1
# qzCHRRKboZFn1EhlifSpxh3v1XF5Psb/Rxi3afupiBk2Qfy+OgUwCI0tU4qTR/NI
# SIG # End signature block

