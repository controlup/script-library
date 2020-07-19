<#
.SYNOPSIS

Show details of installed configurations for Ivanti UWM agents

.DETAILS

Works for both MSI installed configurations (either via Management Centre or other mechanisms) and native configurations

.PARAMETER

None

.CONTEXT

Computer

.NOTES

For native configurations, so those not installed via MSI, there are no details such as names of configurations held on the end-point

In the 2019.1 release, product configurations are no longer prefixd with "Ivanti xxx Manager Configuration" and thus they cannot be retrieved via name for non-native configurations

.MODIFICATION_HISTORY:

@guyrleech 28/07/19

#>

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'

[string[]]$products = @( 
    'Performance'
)

[int]$outputWidth = 400

Function Get-ProductDetails
{
    [CmdletBinding()]

    Param
    (
        [string]$name ,
        [string]$configGUID ,
        [switch]$nativeConfig ,
        $configFile ,
        $agentProperties
    )

    [string]$configKey = Join-Path -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' -ChildPath $configGUID
    $configDetails = Get-ItemProperty -Path $configKey -ErrorAction SilentlyContinue
    $installed = $null 
    if( $configDetails )
    {
        $installed = [datetime]::ParseExact( $configDetails.InstallDate , 'yyyyMMdd' , $null )
    }
    elseif( ! $nativeConfig )
    {
        Write-Warning -Message "Failed to read config installation registry key $configKey"
    }
    $cca = Get-ItemProperty -Path 'HKLM:\SOFTWARE\AppSense Technologies\Communications Agent' -ErrorAction SilentlyContinue
    $result = New-Object -TypeName PSCustomObject 
    Add-Member -InputObject $result -MemberType NoteProperty -Name  'Item' -Value $name
    Add-Member -InputObject $result -MemberType NoteProperty -Name  'Name' -Value $(if( $configDetails ) { $configDetails.DisplayName } else { '-' })
    Add-Member -InputObject $result -MemberType NoteProperty -Name  'Native Config' -Value $(if( $nativeConfig ) { 'Yes' } else { 'No' } )
    if( $nativeConfig )
    {
        Add-Member -InputObject $result -MemberType NoteProperty -Name 'Configuration File' -Value $configFile.FullName
    }
    Add-Member -InputObject $result -MemberType NoteProperty -Name  'Install Date' -Value $(if( $installed ) { Get-Date -Date $installed -Format d } else { '-' })
    Add-Member -InputObject $result -MemberType NoteProperty -Name  'Config File Changed' -Value $configFile.LastWriteTime
    Add-Member -InputObject $result -MemberType NoteProperty -Name  'Config Version' -Value $(if( $configDetails ) { $configDetails.DisplayVersion } else { '-' })
    Add-Member -InputObject $result -MemberType NoteProperty -Name  'Agent Version' -Value ( "{0} ({1}.{2}.{3}.{4})" -f ( ($agentProperties.DisplayName -replace '^.*Agent\s' , '') ,
            (($agentProperties.Version -band 0xFF000000) -shr 24) ,
            (($agentProperties.Version -band 0x00FF0000) -shr 16) ,
             ($agentProperties.Version -band 0x000000FF)  ,
            (($agentProperties.Version -band 0x0000FF00) -shr 8) ))
    Add-Member -InputObject $result -MemberType NoteProperty -Name  'Management Center'  -Value $(if( $cca ) { if( $cca.PSObject.Properties[ 'WebSite' ] ) { $cca.WebSite } else { 'Missing Entry' } } else { 'Not Installed' } ) 
    $result
}

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

try
{
    $provider = Get-PSDrive -Name 'HKCR' -ErrorAction SilentlyContinue
}
catch
{
    $provider = $null
}

if( ! $provider )
{
    $null = New-PSDrive -Name HKCR -PSProvider Registry -Root "Registry::HKEY_CLASSES_ROOT"
}

[array]$results = @( ForEach( $product in $products )
{
    [string]$agentName = $( if( $product -eq 'Application' )
    {
        "^(Ivanti|AppSense) $product (Control|Manager) Agent"
    }
    else
    {
        "^(Ivanti|AppSense) $product Manager Agent"
    })

    $agentProperties = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | Where-Object { $_.PSObject.Properties[ 'DisplayName' ] -and $_.DisplayName -match $agentName }
    $result = $null

    if( $agentProperties )
    {
        ## Find where the config is in case non-default location, copy, unzip and read productcode from manifest.xml
        [string]$configFile = ( Get-ItemProperty -Path "HKLM:\SOFTWARE\AppSense\$Product Manager\Config" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Path -ErrorAction SilentlyContinue)
        [bool]$nativeConfig = $true

        if( [string]::IsNullOrEmpty( $configFile ) ) ## MSI config
        {
            $nativeConfig = $false
            [string]$configFilePath = Join-Path -Path ( Join-Path -Path $env:ALLUSERSPROFILE -ChildPath 'AppSense' ) -ChildPath "$product Manager"
            if( $product -eq 'Application' )
            {
                $configFilePath = Join-Path -Path $configFilePath -ChildPath 'Configuration'
            }
            $configFile = Join-Path -Path $configFilePath -ChildPath "Configuration.a$(($product.SubString(0,1).ToLower()))mp"
        }

        if( Test-Path -Path $configFile -PathType Leaf -ErrorAction SilentlyContinue )
        {
            [string]$tempFile = "$(([System.IO.Path]::GetTempFileName())).zip"
            Copy-Item -Path $configFile -Destination $tempFile
            
            [string]$tempFolder = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.Guid]::NewGuid())
            $null = New-Item -Path $tempFolder -ItemType Directory
            Add-Type -AssemblyName System.IO.Compression.FileSystem -Debug:$false

            [System.IO.Compression.ZipFile]::ExtractToDirectory( $tempFile , $tempFolder )
            if( ! $? )
            {
                Throw "Failed to extract `"$suiteZip`" to `"$installPath`""
            }
            else
            {
                [string]$manifestFile = Join-Path -Path $tempFolder -ChildPath 'Manifest.xml'
                if( Test-Path -Path $manifestFile -PathType Leaf -ErrorAction SilentlyContinue )
                {
                    [xml]$manifest = Get-Content -Path $manifestFile -Raw
                    Get-ProductDetails -configGUID "{$($manifest.Manifest.ProductCode)}" -name $($agentProperties.DisplayName -replace '\sAgent.*$' , '') -agentProperties $agentProperties -nativeConfig:$nativeConfig -configFile (Get-ItemProperty -Path $configFile)
                }
                else
                {
                    Write-Warning -Message "Unable to find manifest.xml in $configFile"
                }
                Remove-Item -Path $tempFolder -Force -Recurse -ErrorAction SilentlyContinue
            }
            Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
        }
        else
        {
            Write-Warning "Unable to find configuration file `"$configFile`""
        }
    }
    else
    {
        Write-Warning -Message "Unable to find agent for $product Manager installation details in registry"
    }
})

if( $results)
{
    $results | Sort-Object -Property Item | Format-Table -AutoSize
}

[array]$pending = @( Get-ItemProperty -Path 'HKLM:\SOFTWARE\AppSense Technologies\Communications Agent\installdefinitions\*' -ErrorAction SilentlyContinue | . { Process `
{
    $result = New-Object -TypeName PSCustomObject 
    Add-Member -InputObject $result -MemberType NoteProperty -Name 'Product' -Value $_.ProductName
    Add-Member -InputObject $result -MemberType NoteProperty -Name  'Package' -Value $_.PackageName
    Add-Member -InputObject $result -MemberType NoteProperty -Name  'Operation' -Value  $( if( $_.Action -eq 1 ) { 'Install' } else { 'Uninstall' })
    Add-Member -InputObject $result -MemberType NoteProperty -Name  'Type' -Value (Get-Culture).TextInfo.ToTitleCase( ($_.Type -split '/')[-1] )
    Add-Member -InputObject $result -MemberType NoteProperty -Name  'Version' -Value $(if( $_.PSObject.Properties[ 'PatchVersion_0' ] ) { $_.PatchVersion_0 } else { $_.Version } )
    $result
}})

if( $pending )
{
    Write-Output -InputObject "Pending work:"
    $pending | Sort-Object -Property Product,Type | Format-Table -AutoSize
}

