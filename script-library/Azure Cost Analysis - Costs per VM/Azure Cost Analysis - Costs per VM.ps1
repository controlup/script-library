<#
.SYNOPSIS
    Get Azure Cost Analysis information, sorted by Service, grouped by Resource.
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
Param(
    [Parameter(
        Position=0, 
        Mandatory=$false, 
        HelpMessage='ControlUp SBA parameter auto entry: Session Host Name'
    )]
    [string] $vmName
)    

If ([string]::IsNullOrEmpty($vmName))
{
    $vmName = ""
}

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

function Get-AzVM {
    <#
    .SYNOPSIS
        Retrieve the Azure VM information.
    .DESCRIPTION
        Retrieve the Azure VM information, using a REST API call.
    .EXAMPLE
        Get-AzVM -BearerToken <string> -VMname <string>
    .CONTEXT
        Azure
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-11-17
        Updated:        2020-11-17
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
            Mandatory=$true, 
            HelpMessage='Enter the Subscription ID'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $SubscriptionID,
    
        [Parameter(
            Position=2, 
            Mandatory=$false, 
            HelpMessage='Enter a VM name'
        )]
        [string] $vmName
    )

    #region Prep variables
        # URL for REST API call, based on given subscription ID
        $uri = "https://management.azure.com/subscriptions/$SubscriptionID/providers/Microsoft.Compute/virtualMachines`?api-version=2020-06-01"#&`$top=5000"

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
    # filter the response if a vmName was provided
    If (!([string]::IsNullOrEmpty($vmName)))
    {
        $results = ($response.value).Where({$_.name -like "$($vmName)"})
    }
    else 
    {
        $results = $response.value
    }
    return $results
}

function Get-AzCostAnalysisActualServiceByResource () {
    <#
    .SYNOPSIS
        Get Azure Cost Analysis information on the actual costs for this month sorted by Service, grouped by Resource.
    .DESCRIPTION
        Get Azure Cost Analysis information on the actual costs for this month sorted by Service, using a REST API call.
    .EXAMPLE
        Get-AzCostAnalysisActualServiceByResource -BearerToken <string> -SubscriptionID <string>
    .CONTEXT
        Azure
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-10-25
        Updated:        2020-10-25
                        Created a separate Azure Cost Analysis function for the costs of this month, sorted by Service

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
        $addMonth = 0
        $time = $((Get-Date).AddMonths($addMonth).ToString("yyyy-MM"))
        $lastdayofmonth = "$([System.DateTime]::DaysInMonth((Get-Date).AddMonths($addMonth).Year,(Get-Date).AddMonths($addMonth).Month))"

        # Create the JSON formatted body
        $body=@{
            "type"= "ActualCost";
            "dataSet"= @{
                "granularity"= "None";
                "aggregation"= @{
                    "totalCost"= @{"name"= "Cost";"function"= "Sum"}
                };
                "grouping"= @(
                    @{"type"="Dimension";"name"="ResourceId"}; 
                    @{"type"="Dimension";"name"="ResourceType"}; 
                    @{"type"="Dimension";"name"="ResourceLocation"}; 
                    @{"type"="Dimension";"name"="ChargeType"}; 
                    @{"type"="Dimension";"name"="ResourceGroupName"}; 
                    @{"type"="Dimension";"name"="PublisherType"}; 
                    @{"type"="Dimension";"name"="ServiceName"}; 
                    @{"type"="Dimension";"name"="ServiceTier"}; 
                    @{"type"="Dimension";"name"="Meter"}
                );
                "include"=@("Tags")
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

function Get-AzCostAnalysisForecastServiceByResource () {
    <#
    .SYNOPSIS
        Get Azure Cost Analysis information on the forecasted costs for this month sorted by Service, grouped by Resource.
    .DESCRIPTION
        Get Azure Cost Analysis information on the forecasted costs for this month sorted by Service, using a REST API call.
    .EXAMPLE
        Get-AzCostAnalysisForecastServiceByResource -BearerToken <string> -SubscriptionID <string>
    .CONTEXT
        Azure
    .NOTES
        Version:        0.1
        Author:         Esther Barthel, MSc
        Creation Date:  2020-10-25
        Updated:        2020-10-25
                        Created a separate Azure Cost Analysis function to support ARM architecture and REST API scripted actions

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
        #$time = (Get-Date).AddMonths(-1).ToString("yyyy-MM")
        $addMonth=0
        $time = $((Get-Date).AddMonths($addMonth).ToString("yyyy-MM"))
        $lastdayofmonth = "$([System.DateTime]::DaysInMonth((Get-Date).AddMonths($addMonth).Year,(Get-Date).AddMonths($addMonth).Month))"

        # Create the JSON formatted body
        $body=@{
            "type"= "ActualCost";
            "dataSet"= @{
                "granularity"= "None";
                "aggregation"= @{
                    "totalCost"= @{"name"= "Cost";"function"= "Sum"}
                };
                "grouping"= @(
                    @{"type"="Dimension";"name"="ResourceId"}; 
                    @{"type"="Dimension";"name"="ResourceType"}; 
                    @{"type"="Dimension";"name"="ResourceLocation"}; 
                    @{"type"="Dimension";"name"="ChargeType"}; 
                    @{"type"="Dimension";"name"="ResourceGroupName"}; 
                    @{"type"="Dimension";"name"="PublisherType"}; 
                    @{"type"="Dimension";"name"="ServiceName"}; 
                    @{"type"="Dimension";"name"="ServiceTier"}; 
                    @{"type"="Dimension";"name"="Meter"}
                );
                "include"=@("Tags")
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


    # Retrieve the VM information
    $azVM = $null
    $azVM = (Get-AzVM -bearerToken $($azBearerToken.access_token) -SubscriptionID $($azSubscription.subscriptionId.Split("/")[-1]) -vmName $vmName)
    #debug: Write-Output "DEBUG INFO - VM: $($azVM)"

    If ($(($azVM | Measure-Object).Count) -gt 1)
    {
        Write-Warning "More than one VM found, selecting first VM in list as resource."
        Write-Output ""
    }
    If ($(($azVM | Measure-Object).Count) -ge 1)
    {
        # Only process the first entry as we need do not want to blow up the Where ScriptBlock
        $selectedVMName = $($azVM[0].name)
        # Creating a dynamic Where filter
        $whereArray = @()
        # Add vmName to Where ScriptBlock
        $whereArray += "(`$_.Resource -like ""$($azVM[0].name)"")"
        # Add osDisk to Where ScriptBlock
        $whereArray += "(`$_.Resource -like ""$($azVM[0].properties.storageProfile.osDisk.name)"")"
        # Add dataDisks to Where ScriptBlock
        foreach ($dataDisk in $($azVM[0].properties.storageProfile.dataDisks))
        {
            $whereArray += "(`$_.Resource -like ""$($dataDisk.name)"")"
        }
        #Build the where array into a string by joining each statement with -and            
        $whereString = $whereArray -Join " -or "
        #debug: Write-Output "DEBUG INFO - whereString: $($whereString)"
            
        #Create the scriptblock with your final string            
        $whereBlock = [scriptblock]::Create($whereString)
    }

    # Retrieve the Cost Analysis - Actual Costs details
    $costAnalysisResults = Get-AzCostAnalysisActualServiceByResource -bearerToken $($azBearerToken.access_token) -SubscriptionID $($azSubscription.subscriptionId.Split("/")[-1])

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
    # Select the Cost information
    $dataset = $dataResults | Select Cost, 
        ResourceId, 
        @{Name='ResourceType'; Expression={$($_.ResourceType.Split("/")[-1])}}, 
        ResourceLocation, 
        ChargeType, 
        ResourceGroupName, 
        PublisherType, 
        ServiceName, 
        ServiceTier, 
        Meter, 
        Tags, 
        @{Name='Resource'; Expression={$($_.ResourceId.Split("/")[-1])}}, 
        @{Name='Costs'; Expression={$([math]::Round($_.Cost,2))}}, 
        @{Name='Location'; Expression={$_.ResourceLocation}}, 
        @{Name='Resource Group'; Expression={$($_.ResourceGroupName)}}, 
        @{Name='Percentage'; Expression={$([math]::Round((($_.Cost/$totalCosts)*100),0))}}, 
        Currency | 
            Sort Costs -Descending | 
            Sort Resource 

    # Filter the Cost Analysis - Actual Costs by Resource for VM and corresponding disks
    $vmCostsDataset = $dataset | Where -FilterScript $whereBlock

    # Retrieve the Cost Analysis - Forecast details
    $forecastAnalysisResults = Get-AzCostAnalysisForecastServiceByResource -bearerToken $($azBearerToken.access_token) -SubscriptionID $($azSubscription.subscriptionId.Split("/")[-1]) #-ResourceGroupName $ResourceGroupName

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

    # Calculate the Summary information
    $totalVMCosts = $([math]::Round((($vmCostsDataset | Measure-Object -Property Cost -Sum).Sum),2))
    $totalActualCosts = $([math]::Round((($dataResults | Measure-Object -Property Cost -Sum).Sum),2))
    $totalForecastCosts = $([math]::Round( (($forecastResults | Measure-Object -Property Cost -Sum).Sum),2))

    # Present the results
    Write-Host "Actual costs breakdown by service associated with VM: " -ForegroundColor Yellow -NoNewline
    Write-Host "$($selectedVMName)" -ForegroundColor Cyan
    $vmCostsDataset | 
        Where {$_.Costs -gt 0} | 
            Select ResourceId, ServiceName, ServiceTier, Meter, Costs, Currency | 
                Sort Resource, ServiceName, ServiceTier, Meter | 
                    Format-Table @{Name='Resource';Expression={$_.ResourceId.Split("/")[-1]}},
                        @{Name='Service name';Expression={$_.ServiceName}}, 
                        @{Name='Service tier';Expression={$_.ServiceTier}}, 
                        Meter, 
                        @{Name='Costs         '; Expression={"$($_.Currency) {0,10:N2}" -f($($_.Costs))}; Align="right"}
    Write-Host "Total costs for this VM: " -ForegroundColor Yellow -NoNewline
    Write-Host "$($vmCostsDataset[0].Currency) $($totalVMCosts)" -ForegroundColor Cyan -NoNewline
    Write-Host ", which is " -ForegroundColor Yellow -NoNewline
    Write-Host "~$([math]::Round((($totalVMCosts/$totalActualCosts)*100),2)) % " -ForegroundColor Cyan -NoNewline
    Write-Host "of the actual Azure subscription costs for this month " -ForegroundColor Yellow -NoNewline
    Write-Host "($($dataResults[0].Currency) $($totalActualCosts))" -ForegroundColor Cyan
}
else 
{
    Write-Warning "No Azure Credentials could be retrieved from the stored credentials file for this user."
}


