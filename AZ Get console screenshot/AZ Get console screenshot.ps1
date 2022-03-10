#require -version 3.0

<#
.SYNOPSIS
    Get Azure VM console screenshots

.DESCRIPTION
    Using REST API calls

.PARAMETER azid
    The relative URI of the Azure VM

.PARAMETER AZtenantId
    Optional Azure tenant id. Specify when there is a need to access multiple tenants with different credentials.

.NOTES
    Version:        0.1
    Author:         Guy Leech, BSc based on code from Esther Barthel, MSc
    Creation Date:  2021-09-15
    Updated:        2021-10-13 Guy Leech  Added support for credential files with tenant id in the file name rather than in the file
                    2021-10-13 Guy Leech  Workaround for VM name in Azid being passed as lowercase which breaks the blob URL
                    2022-02-22 Guy Leech  Look for _AZ_ credentials file if _Azure_ not found. Checks on AZ id and tenant id validity
#>

[CmdletBinding()]

Param
(
    [string]$AZid , ## passed by CU as the URL to the VM minus the FQDN
    [string]$AZtenantId ,
    [string]$screenshotFolder ,
    [string]$screenShotFile
)    

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 250
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

[string]$apiversion = '2021-04-01'
[string]$computeApiVersion = '2021-07-01'
[string]$baseURL = 'https://management.azure.com'
[string]$credentialType = 'Azure'

Write-Verbose -Message "AZid is $AZid"

Function Save-ScreenShot
{
    Param
    (
        [Parameter(Mandatory=$true,HelpMessage='Existing screenshot file')]
        [string]$imageFile ,
        [Parameter(Mandatory=$true,HelpMessage='Name of th VM for the screenshot file')]
        [string]$vmname  ,
        [Parameter(Mandatory=$true,HelpMessage='Folder to copy screenshot file to')]
        [string]$folder ,
        [string]$file ,
        [Parameter(Mandatory=$true,HelpMessage='Extension of the screenshot file')]
        [string]$extension ,
        [switch]$report
    )

    $folder = [System.Environment]::ExpandEnvironmentVariables( $folder ) -replace '%vmname%' , $vmname -replace '%month%' , (Get-Date -Format 'MM')  -replace '%year%' , (Get-Date -Format 'yyyy') `
                -replace '%day%' , (Get-Date -Format 'dd') -replace '%hour%' , (Get-Date -Format 'HH')  -replace '%minute%' , (Get-Date -Format 'mm') -replace '%second%' , (Get-Date -Format 'ss')

    if( -Not [string]::IsNullOrEmpty( $file ) )
    {
        $file = [System.Environment]::ExpandEnvironmentVariables( $file ) -replace '%vmname%' , $vmname -replace '%month%' , (Get-Date -Format 'MM')  -replace '%year%' , (Get-Date -Format 'yyyy') `
                -replace '%day%' , (Get-Date -Format 'dd') -replace '%hour%' , (Get-Date -Format 'HH')  -replace '%minute%' , (Get-Date -Format 'mm') -replace '%second%' , (Get-Date -Format 'ss')
    }
    else
    {
        $file = "$vmname.$((Get-Date -Format s) -replace ':' , '-').$extension"
    }
    if( -Not ( Test-Path -Path $folder -ErrorAction SilentlyContinue ) )
    {
        $null = New-Item -Path $folder -ItemType Directory -Force
    }
    if( Copy-Item -Path $imageFile -Destination (Join-Path -Path $folder -ChildPath $file) -PassThru )
    {
        if( $report )
        {
            Write-Output -InputObject "Screenshot file $file saved ok to $folder"
        }
    }
}


#region GUI
Function New-Form
{
    Param
    (
        [Parameter(Mandatory=$true)]
        $inputXaml
    )

    $form = $null
    $inputXML = $inputXaml -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
 
    if( [xml]$xaml = $inputXML )
    {
        $reader = New-Object -TypeName Xml.XmlNodeReader -ArgumentList $xaml

        try
        {
            if( $form = [Windows.Markup.XamlReader]::Load( $reader ) )
            {
                $xaml.SelectNodes( '//*[@Name]' ) | ForEach-Object `
                {
                    Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
                }
            }
        }
        catch
        {
            Write-Error -Message "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$_"
        }
 
    }
    else
    {
        Write-Error -Message 'Failed to convert input XAML to WPF XML'
    }

    $form ## return
}

Function Show-ImageInWindow
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$filename ,
        [string]$windowTitle
    )
    
    [string]$screenshotXAML = @'
<Window x:Class="VMWare_GUI.Screenshot"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:VMWare_GUI"
        mc:Ignorable="d"
        Title="Screenshot" Height="700" Width="950">
    <Grid>
        <Image x:Name="imgScreenshot" Margin="20,20,20,83"/>
    </Grid>
</Window>
'@
    Add-Type -AssemblyName PresentationCore,PresentationFramework #,System.Windows.Forms

    if( $screenshotWindow = New-Form -inputXaml $screenshotXAML )
    {
        if( $filestream = New-Object System.IO.FileStream -ArgumentList $filename , Open , Read )
        {
            $bitmap = New-Object -Typename System.Windows.Media.Imaging.BitmapImage
            $bitmap.BeginInit()
            $bitmap.StreamSource = $filestream
            $bitmap.EndInit()
            $wpfimgScreenshot.Source = $bitmap
            $screenshotWindow.Title = $windowTitle

            $null = $screenshotWindow.ShowDialog()
            $bitmap.StreamSource = $null
            $filestream.Close()
            $filestream = $null
            $bitmap = $null
        }
        else
        {
            Write-Error -Message "Failed to open image file $filename"
        }
    }
}
#endregion GUI

#region AzureFunctions

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
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$system ,
        [string]$tenantId
    )

    $strAzSPCredFolder = [System.IO.Path]::Combine( [environment]::GetFolderPath('CommonApplicationData') , 'ControlUp' , 'ScriptSupport' )
    $AzSPCredentials = $null

    Write-Verbose -Message "Get-AzSPStoredCredentials $system"

    [string]$credentialsFile = $(if( -Not [string]::IsNullOrEmpty( $tenantId ) )
        {
            [System.IO.Path]::Combine( $strAzSPCredFolder , "$($env:USERNAME)_$($tenantId)_$($System)_Cred.xml" )
        }
        else
        {
            [System.IO.Path]::Combine( $strAzSPCredFolder , "$($env:USERNAME)_$($System)_Cred.xml" )
        })

    Write-Verbose -Message "`tCredentials file is $credentialsFile"

    If (Test-Path -Path $credentialsFile)
    {
        try 
        {
            if( ( $AzSPCredentials = Import-Clixml -Path $credentialsFile ) -and -Not [string]::IsNullOrEmpty( $tenantId ) -and -Not $AzSPCredentials.ContainsKey( 'tenantid' ) )
            {
                $AzSPCredentials.Add(  'tenantID' , $tenantId )
            }
        }
        catch 
        {
            Write-Error -Message "The required PSCredential object could not be loaded from $credentialsFile : $_"
        }
    }
    Elseif( $system -eq 'Azure' )
    {
        ## try old Azure file name 
        $azSPCredentials = Get-AzSPStoredCredentials -system 'AZ' -tenantId $tenantId 
    }

    if( -not $AzSPCredentials )
    {
        Write-Error -Message "The Azure Service Principal Credentials file stored for this user ($($env:USERNAME)) cannot be found at $credentialsFile.`nCreate the file with the Set-AzSPCredentials script action (prerequisite)."
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
        [Parameter(Mandatory=$true, HelpMessage='Azure Service Principal credentials' )]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCredential] $SPCredentials,

        [Parameter(Mandatory=$true, HelpMessage='Azure Tenant ID' )]
        [ValidateNotNullOrEmpty()]
        [string] $TenantID
    )

    [string]$uri = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
        
    [hashtable]$body = @{
        grant_type    = 'client_credentials'
        client_Id     = $SPCredentials.UserName
        client_Secret = $SPCredentials.GetNetworkCredential().Password
        scope         = "$baseURL/.default"
    }

    [hashtable]$invokeRestMethodParams = @{
        Uri             = $uri
        Body            = $body
        Method          = 'POST'
        ContentType     = 'application/x-www-form-urlencoded'
    }

    Invoke-RestMethod @invokeRestMethodParams | Select-Object -ExpandProperty access_token -ErrorAction SilentlyContinue
}

function Invoke-AzureRestMethod {
   
    [CmdletBinding()]
    Param(
        [Parameter( Mandatory=$true, HelpMessage='A valid Azure bearer token' )]
        [ValidateNotNullOrEmpty()]
        [string]$BearerToken ,
        [string]$uri ,
        [ValidateSet('GET','POST','PUT')] ## add others as necessary
        [string]$method = 'GET' ,
        [hashtable]$body ,
        [string]$property = 'value' ,
        [string]$contentType = 'application/json'
    )

    [hashtable]$header = @{
        'Authorization' = "Bearer $BearerToken"
    }

    if( ! [string]::IsNullOrEmpty( $contentType ) )
    {
        $header.Add( 'Content-Type'  , $contentType )
    }

    [hashtable]$invokeRestMethodParams = @{
        Uri             = $uri
        Method          = $method
        Headers         = $header
    }

    if( $PSBoundParameters[ 'body' ] )
    {
        $invokeRestMethodParams.Add( 'Body' , ( $body | ConvertTo-Json ))
    }
    
    if( -not [String]::IsNullOrEmpty( $property ) )
    {
        Invoke-RestMethod @invokeRestMethodParams | Select-Object -ErrorAction SilentlyContinue -ExpandProperty $property
    }
    else
    {
        Invoke-RestMethod @invokeRestMethodParams ## don't pipe through select as will slow script down for large result sets if processed again after rreturn
    }
}

#endregion AzureFunctions

[string]$vmName = ($AZid -split '/')[-1]
    
if( [string]::IsNullOrEmpty( $vmName ) )
{
    Throw "Azure id `"$AZid`" does not appear valid - failed to find VM name"
}

if( -Not [string]::IsNullOrEmpty( $AZtenantId ) -and -Not ( $AZtenantId -as [guid] ) )
{
    Throw "Azure tenant id `"$AZtenantId`" is invalid"
}

if( ! ( $webClient = New-Object -TypeName System.Net.WebClient ) )
{
    Throw "Failed to create a System.Net.WebClient object"
}

If ($azSPCredentials = Get-AzSPStoredCredentials -system $credentialType -tenantId $AZtenantId )
{
    # Sign in to Azure with a Service Principal with Contributor Role at Subscription level and retrieve the bearer token
    Write-Verbose -Message "Authenticating to tenant $($azSPCredentials.tenantID) as $($azSPCredentials.spCreds.Username)"
    if( -Not ( $azBearerToken = Get-AzBearerToken -SPCredentials $azSPCredentials.spCreds -TenantID $azSPCredentials.tenantID ) )
    {
        Throw "Failed to get Azure bearer token"
    }

    $timeZone = Get-TimeZone
    [string]$currentTimeZone = $timeZone.StandardName
    if( $timeZone.IsDaylightSavingTime([datetime]::Now) )
    {
        $currentTimeZone = $timeZone.DaylightName
    }

    if( ( $vm = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid/?api-version=$computeApiVersion" -property $null ) -and ! [string]::IsNullOrEmpty( $vm.id ) )
    {
        ## GRL 2021-10-20 appears to be a CU bug/feature that transmogrifies the VM name to lowercase which produces a blob URL that doesn't work so we replace the Az ID with what is returned here
        $AZid = $vm.id
        $vmName = $vm.Name
    }

    ## https://docs.microsoft.com/en-us/rest/api/compute/virtual-machines/retrieve-boot-diagnostics-data
    if( $bootDiagnostics = Invoke-AzureRestMethod -BearerToken $azBearerToken -uri "$baseURL/$azid/retrieveBootDiagnosticsData?api-version=$computeApiVersion" -property $null -method POST )
    {
        if( [string]$screenshotURL = $bootDiagnostics.PSObject.Properties | Where-Object Name -match screenshot | Select-Object -ExpandProperty Value )
        {
            [string]$extension = $screenshotURL -replace '.*\.(\w+)\?.*$' , '$1'
            [string]$tempFile = Join-Path -Path $env:TEMP -ChildPath "$((New-Guid).Guid).$extension"
            [datetime]$retrievedDate = [datetime]::Now
            try
            {
                Write-Verbose -Message "Downloading $screenshotURL to $tempFile"
                $webClient.DownloadFile( $screenshotURL , $tempFile )
            }
            catch
            {
                Write-Warning -Message "Failed to download screenshot from $screenshotURL to $tempfile for VM $vmName : $_"
            }
            if( Test-Path -Path $tempFile )
            {
                if( $PSBoundParameters[ 'screenshotFolder' ] -or $PSBoundParameters[ 'screenShotFile' ] )
                {
                    Save-ScreenShot -imageFile $tempFile -folder $screenshotFolder -file $screenShotFile -extension $extension -vmname $vmName -report
                }
                else
                {
                    Show-ImageInWindow -filename $tempFile -windowTitle "Screenshot of $vmName retrieved at $(Get-Date -Date $retrievedDate -Format G) $currentTimeZone"
                }
                Remove-Item -Path $tempFile
            }
        }
        else
        {
            Write-Warning -Message "Failed to find screenshot URL for VM $vmName"
        }
    }
    else
    {
        Write-Warning -Message "Failed to get boot diagnostics for $vmName"
    }
}

