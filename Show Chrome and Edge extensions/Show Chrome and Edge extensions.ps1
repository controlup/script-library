<#
    Show Chrome extensions installed for all users by looking at their local profiles

    @guyrleech 19/05/20
#>

[CmdletBinding()]

Param
(
    [string]$username ## if not specified then all profiles
)

$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

[int]$outputWidth = 400

Function Get-ExtensionDetail
{
    [CmdletBinding()]

    Param
    (
        [Parameter(Mandatory)]
        [string]$folder ,
        [string]$browser = 'Chrome'
    )
    
    Get-ChildItem -Path $folder| Where-Object { $_.PSIsContainer -and $_.Name -cmatch '^[a-z]{32}$' } | . { Process `
    { 
        [string]$extensionid = $_.name
        [string]$extensionName = $extensions[ $extensionid ]
            
        if( [string]::IsNullOrEmpty( $extensionName ) )
        {
            ## The * is for a folder which will be one or more version numbers so get the latest created
            [string]$manifestFile = [System.IO.Path]::Combine( $folder , $extensionid , '*' , 'manifest.json' )
            if( ( Test-Path -Path $manifestFile -PathType Leaf -ErrorAction SilentlyContinue ) `
                -and ( $manifestFile = Get-ChildItem -Path $manifestFile | Sort-Object -Property CreationTime -Descending | Select-Object -First 1 ) `
                    -and ( $manifest = Get-Content -Path $manifestFile | ConvertFrom-Json ) `
                        -and ( $manifest.PSObject.Properties[ 'name' ] ))
            {
                if( ( $extensionName = $manifest.name ) -match '^__MSG_(.*)__' )
                {
                    [string]$appName = $Matches[1]
                    ## need to look up in locale messages file
                    if( ( [string]$messagesFile = [System.IO.Path]::Combine( (Split-Path -Path $manifestFile -Parent ) , '_locales' , ( (Get-Culture).Name -replace '\-' , '_' ) , 'messages.json' ) ) `
                        -and ( Test-Path -Path $messagesFile -PathType Leaf -ErrorAction SilentlyContinue ) `
                            -and ( $messages = Get-Content -Path $messagesFile | ConvertFrom-Json ) `
                                -and ( $messages.PSObject.Properties[ $appName ] ))
                    {
                        $extensionName = $messages.$appName | Select-Object -ExpandProperty 'Message'
                    }
                    ## if was for example en-GB then look up en
                    elseif( ( [string]$messagesFile = [System.IO.Path]::Combine( (Split-Path -Path $manifestFile -Parent ) , '_locales' , ( (Get-Culture).Name -replace '\-.*$' ) , 'messages.json' ) ) `
                        -and ( Test-Path -Path $messagesFile -PathType Leaf -ErrorAction SilentlyContinue ) `
                            -and ( $messages = Get-Content -Path $messagesFile | ConvertFrom-Json ) `
                                -and ( $messages.PSObject.Properties[ $appName ] ))
                    {
                        $extensionName = $messages.$appName | Select-Object -ExpandProperty 'Message'
                    }
                    elseif( ( [string]$messagesFile = [System.IO.Path]::Combine( (Split-Path -Path $manifestFile -Parent ) , '_locales' , 'en' , 'messages.json' ) ) `
                        -and ( Test-Path -Path $messagesFile -PathType Leaf -ErrorAction SilentlyContinue ) `
                            -and ( $messages = Get-Content -Path $messagesFile | ConvertFrom-Json ) `
                                -and ( $messages.PSObject.Properties[ $appName ] ))
                    {
                        $extensionName = $messages.$appName | Select-Object -ExpandProperty 'Message'
                    }
                }
                $extensions.Add( $extensionId , $extensionName )
            }
        }

        [pscustomobject]@{ 'id' = $extensionid ; 'Extension' = $extensionName ; 'Browser' = $browser ; 'Path' = $folder }
    }}
}

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

## we'll cache extensions we've successfuly looked up
[hashtable]$extensions = @{}

## We will look up folder redirections in the registry if the profile is loaded
try
{
    $provider = Get-PSDrive -Name HKU -ErrorAction SilentlyContinue
}
catch
{
    $provider = $null
}

if( ! $provider )
{
    if( ! ( $hku = New-PSDrive -Name 'HKU' -PSProvider Registry -Root "Registry::HKU" ) )
    {
        Write-Warning -Message "Failed to create HKU PS drive"
    }
}

## See if we have a Chrome machine policy for user data directory

[string]$chromePoliciesMachineKey = Join-Path -Path (Join-Path -Path 'HKU:' -ChildPath $profile.SID ) -ChildPath 'SOFTWARE\Policies\Google\Chrome'
[string]$computeruserdatadir = (Get-ItemProperty -Path $chromePoliciesMachineKey -Name 'UserDataDir' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'UserDataDir' -ErrorAction SilentlyContinue) `
    -replace '\${machine_name\}' , $env:COMPUTERNAME

[int]$counter = 0

[array]$profiles = @()

if( ! [string]::IsNullOrEmpty( $username ) )
{
    if( ! ( $sid = (New-Object System.Security.Principal.NTAccount($username)).Translate([System.Security.Principal.SecurityIdentifier]).value ) )
    {
        Throw "Unable to get SID for user $username"
    }
    $profiles = @( Get-CimInstance -ClassName win32_userprofile -ErrorAction SilentlyContinue -Filter "sid = '$sid'")
    if( ! $profiles -or ! $profiles.Count )
    {
        Throw "Unable to find profile for user $username in session $sessionid (sid $sid)"
    }
}
else
{
    $profiles = @( Get-CimInstance -ClassName win32_userprofile -ErrorAction SilentlyContinue -Filter "Special = 'FALSE'" )
}

[array]$results = @( $profiles | . { Process `
{
    $profile = $PSItem
    $counter++
    [string]$user = $null
    try
    {
        $user = ([System.Security.Principal.SecurityIdentifier]( $profile.SID )).Translate([System.Security.Principal.NTAccount]).Value
    }
    catch
    {
        $user = $null
    }

    Write-Verbose "$counter : user $user"

    ## see if local appdata has been redirected
    
    [string]$localAppdata = $null
    [string]$userdatadir = $computeruserdatadir

    if( $profile.Loaded )
    {
        [string]$folderRedirectionsKey = Join-Path -Path (Join-Path -Path 'HKU:' -ChildPath $profile.SID ) -ChildPath 'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        $localAppdata = Get-ItemProperty -Path $folderRedirectionsKey -Name 'Local AppData' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'Local AppData' -ErrorAction SilentlyContinue

        if( [string]::IsNullOrEmpty( $userdatadir ) )
        {
            ## Computer takes precedence over machine
            [string]$chromePoliciesKey = Join-Path -Path (Join-Path -Path 'HKU:' -ChildPath $profile.SID ) -ChildPath 'SOFTWARE\Policies\Google\Chrome'
            ## Replace Chrome building blocks https://www.chromium.org/administrators/policy-list-3/user-data-directory-variables
            $userdatadir = Get-ItemProperty -Path $chromePoliciesKey -Name 'UserDataDir' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'UserDataDir' -ErrorAction SilentlyContinue
        }
    }

    if( [string]::IsNullOrEmpty( $userdatadir ) -and [string]::IsNullOrEmpty( $localAppdata ) )
    {
        $localAppdata = Join-Path -Path $profile.LocalPath -ChildPath 'AppData\Local'
    }
    else
    {
        $userdatadir = $userdatadir -replace '\$\{profile\}' , $profile.LocalPath -replace '\$\{user_name\}' , $user ## More variables to do
    }
    
    ## If profile is loaded, look in HKU to see if localappdata is redirected
    [string]$chromeBaseFolder = $(if( ! [string]::IsNullOrEmpty( $userdatadir )  ) 
    {
        Join-Path -Path $userdatadir -ChildPath 'Default\Extensions'
    }
    else
    {
        Join-Path -Path $localAppdata -ChildPath 'Google\Chrome\User Data\Default\Extensions'
    })

    if( Test-Path -Path $chromeBaseFolder -PathType Container -ErrorAction SilentlyContinue )
    {
        Write-Verbose -Message "$counter : $chromeBaseFolder"
        Get-ExtensionDetail -folder $chromeBaseFolder -browser 'Chrome'
    }
 
    [string]$edgeBaseFolder = Join-Path -Path $profile.LocalPath -ChildPath 'AppData\Local\Microsoft\Edge\User Data\Default\Extensions'

    if( Test-Path -Path $edgeBaseFolder -PathType Container -ErrorAction SilentlyContinue )
    {
        Write-Verbose -Message "$counter : $edgeBaseFolder"
        Get-ExtensionDetail -folder $edgeBaseFolder -browser 'Edge'
    }
}})

## can be in "Google Chrome" key or a GUID
if( ! ( Get-ItemProperty -Path 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | Where-Object { $_.PSObject.Properties[ 'DisplayName' ] -and $_.DisplayName -eq 'Google Chrome' }  ) `
    -and ! ( Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | Where-Object { $_.PSObject.Properties[ 'DisplayName' ] -and $_.DisplayName -eq 'Google Chrome' } ) )
{
    Write-Warning -Message "Google Chrome is not installed"
}

if( $results -and $results.Count )
{
    if( [string]::IsNullOrEmpty( $username ) )
    {
        Write-Output -InputObject "Found $($extensions.Count) unique extensions for $counter local user profiles"
    }
    else
    {
        Write-Output -InputObject "Found $($extensions.Count) unique extensions for user $username"
    }
    ## Because we group on id and browser, Name will be "id, browser"
    $results|Group-Object -Property id,Browser | Select-Object @{n='Id';e={($_.Name -split ',')[0]}},@{n='Browser';e={($_.Name -split ',' , 2)[-1].Trim()}},@{n='Extension';e={$_.Group[0].Extension}},Count | Sort-Object -Property @{ e='Count' ; Descending = $true },@{ e='Extension'; Descending = $false} | Format-Table -AutoSize
}
else
{
    Write-Warning -Message "No Chrome extensions found"
}

