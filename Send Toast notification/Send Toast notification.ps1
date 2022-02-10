<#
.SYNOPSIS
Create a toast notification

.DESCRIPTION
To use this script in an automated action, take a copy of it, add the _clientMetricX parameters as required and set the message text in the param() block with {0}, {1}, etc as required and set any other parameters where you don't want the default

.PARAMETER _clientMetric1
Parameter passed from ControlUp record properties and replaced in message string. Where more than one is specified, the trailing digits are sorted numerically to determine order of replacement in the message string, eg {1} would be replaced by _clientMetric2, etc

.PARAMETER message
The message to display in the dialogue. If specifying client metrics with _ prefix, use {0} in the string to have it replaced with the first _ parameter numerically first, {1} for second, etc where trailing digits are sorted numerically to determine order
Add the message text in the script param block itself if using as an automated action.

.PARAMETER title
The title for the dialogue. If specified as an empty string or $null, no title bar is shown

.PARAMETER logo
Path of a graphic file to use for the popup or a string to match against .png files in \windows\systemresources.
Specify DefaultSystemNotification to get the default system notification icon otherwise a logo embedded into the script is used.

.PARAMETER audiosource
The notification sound name to be played (see https://docs.microsoft.com/en-us/uwp/schemas/tiles/toastschema/element-audio).
Will be silent if $null or empty string specified.

.PARAMETER application
Application to set as source for notification. Use Get-StartApps to see what is available, use the Name of the application.

.EXAMPLE

& '.\Show Toast Popup for Wifi signal.ps1' -_clientMetric1 42 -message "Poor WiFi Signal ({0}%)"

.NOTES

Code adapted from https://steviecoaster.dev/Toast-everything/

..MODIFICATION HISTORY

@guyrleech 2021-05-13  First public release
@guyrleech 2021-06-23  Added CU logo
@guyrleech 2021-06-28  Added ability to have no logo or override embedded one. Added option for silent notification sound if $null or empty string passed via -audiosource. Cache logo file. No default for -message
@guyrleech 2021-10-21  Fixes to allow working pre Win10
Ton de Vreede 2022-01-25 Added Windows version check and success confirmation

#>

[CmdletBinding()]

Param
(
    ## client metrics (or any parameter passed automagically via CU) must start with an underscore and have the number of the positional parameter in the $message string at the end (which they are sorted on before constructing the message string), eg _clientMetric2
    ## do not have digits anywhere else in the parameter name other than at the end
    ## if not passing any record properties via the SBA definition, delete the _ parameter(s) completely
    ## to show a number without decimal places, make the parameter an [int] type
    ## [int]$_clientMetric1 ,
	$_clientMetric1 ,
    $_clientMetric2 ,
    $_clientMetric3 ,
    $_clientMetric4 ,
    $_clientMetric5 ,

    [Parameter(Mandatory=$true,HelpMessage='Text to display to user in toast notification')]
    [string]$message , ## define the message here if using as an automated action

    [string]$title = "ControlUp Alert",
    [AllowEmptyString()][AllowNull()]
    [string]$logo  ,
    [string]$application = 'Windows Powershell*' ,
    [AllowEmptyString()][AllowNull()]
    [string]$audiosource = 'ms-winsoundevent:Notification.Default'
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

# Test if Windows version is supported
[CimInstance]$objWindows = Get-CimInstance -ClassName 'win32_operatingsystem' -Property Name, Version
[version]$verWindows = $objWindows.Version
[string]$strWindows = $objWindows.Name.Split('|')[0]
If ($verWindows.Major -lt '10.0') {
	Throw "This script only runs on Windows 10/Server 2016 or higher. This is $strWindows`."
}

# Test for correct session
[int]$sessionId = Get-Process -Id $pid | Select-Object -ExpandProperty SessionId

if( $sessionId -eq 0 )
{
    Throw "Toast notifications cannot be shown in session zero - set the script to run in the context of the users session"
}

$Priority = $null

try
{
    $Priority = [Windows.UI.Notifications.ToastNotificationPriority]::High ## doesn't seem any different to "Default" in appearance
}
catch
{
    Write-Verbose -Message "Error when setting Notification Priority. This is expected on Server 2016, notification will still show."
}

## get the underscore parameters from the parameters so we can expand the message string - put in hashtable keyed on number at the end of the parameter name so we can sort on that and check for duplicates
[hashtable]$messageStrings = @{}

ForEach( $parameter in $PSBoundParameters.GetEnumerator() )
{
    if( $parameter.Key -match '^_[^\d]*(\d*)$' )  ## _clientMetric1
    {
        try
        {
            $messageStrings.Add( [int]$Matches[1] , $parameter.Value )
        }
        catch
        {
            Throw "Already have an _ parameter ending in number $($Matches[1]) so can't use $($parameter.Key)"
        }
    }
}

Write-Verbose -Message "Got $($messageStrings.Count) parameters for message string"

if( $message -match '\{0\}' -and $messageStrings.Count -eq 0 )
{
    Write-Warning -Message "Message string contains {0} but no record properties were passed as parameters"
}

try
{
    $message = $message -f ($messageStrings.GetEnumerator() | Sort-Object -Property Key | Select-Object -ExpandProperty Value)
}
catch
{
    Write-Error "Failed to construct message string - are there sufficient {n} place holders for all the -_clientMetric parameters and vice versa?"
    Throw $_
}

Write-Verbose -Message "Expanded message text is `"$message`""

Write-Verbose -Message "Session id is $sessionId user name $env:USERNAME"

if( ! [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] `
    -or ! [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] )
{
    Throw "Failed to load required .NET classes"
}

try
{
    if( ! $PSBoundParameters.ContainsKey( 'logo' ) )
    {
        [string]$controlupLogo = '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'
        [string]$scriptSupportFolder = [System.IO.Path]::Combine( [Environment]::GetFolderPath( [Environment+SpecialFolder]::CommonApplicationData ) , 'ControlUp' , 'ScriptSupport' )
        if( ! (Test-Path -Path $scriptSupportFolder -PathType Container -ErrorAction SilentlyContinue ) -and ! ( New-Item -Path $scriptSupportFolder -ItemType Directory ) )
        {
            $scriptSupportFolder = $env:temp
        }
        [string]$ImageFile = Join-Path -Path $scriptSupportFolder -ChildPath "cu.toast.logo.png"
        [byte[]]$Bytes = [convert]::FromBase64String( $controlupLogo  )

        if( $Bytes.Count )
        {
            try
            {
                [System.IO.File]::WriteAllBytes( $ImageFile , $Bytes )
                $logo = $ImageFile
            }
            catch
            {
                Write-Warning -Message "Error writing to file `"$ImageFile`""
                $logo = $null
            }
        }
    }
    elseif( ! [string]::IsNullOrEmpty( $logo ) -and ! (Test-Path -Path $logo -PathType Leaf -ErrorAction SilentlyContinue ) ) ## null or empty -logo argument means no logo
    {
        $logoFolder = Join-Path -Path $env:SystemRoot -ChildPath 'SystemResources'
        if( ! ( $logoFile = Get-ChildItem -Path $logoFolder -Force -Recurse -File -Filter "*$logo*.png" | Select-Object -First 1 -ExpandProperty FullName ) )
        {
            Throw "Failed to find logo file $logo in $logoFolder"
        }
        $logo = $logoFile
    }

    [string]$audio = $null
    if( [string]::IsNullOrEmpty( $audiosource ) )
    {
        $audio = 'silent="true"'
    }
    else
    {
        $audio = "src=`"$audiosource`""
    }

    ## https://docs.microsoft.com/en-us/uwp/schemas/tiles/toastschema/schema-root
    $XmlString = @"
      <toast>
        <visual>
          <binding template="ToastGeneric">
            <text>$Title</text>
            <text>$Message</text>
            <image src="$Logo" placement="appLogoOverride" hint-crop="circle" />
          </binding>
        </visual>
        <audio $audio/>
      </toast>
"@

    try
    {
        Import-Module -Name StartLayout -Verbose:$false
    }
    catch
    {
        Import-Module -Name StartScreen -Verbose:$false
    }

    ## we need an AppID so grab one
    [string]$AppId = Get-StartApps -name $application | Where-Object AppID -NotMatch 'AutoGenerated' | Select-Object -ExpandProperty AppID -First 1
    Write-Verbose -Message "AppId is $appid"
    if( $ToastXml = New-Object -TypeName Windows.Data.Xml.Dom.XmlDocument )
    {
        $ToastXml.LoadXml($XmlString)
        ##if( $Toast = [Windows.UI.Notifications.ToastNotification]::new($ToastXml) )
        if( $Toast = New-Object -TypeName Windows.UI.Notifications.ToastNotification -ArgumentList $ToastXml)
        {
            if( $Priority )
            {
                $toast.Priority = $priority
            }
            [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppId).Show($Toast)
			Write-Output -InputObject "Toast notification sent to user."
        }
        else
        {
            Throw "Failed to create toast notification from XML"
        }
    }
}
catch
{
    Throw $_
}
finally
{
    if( $ImageFile -and ( Test-Path -Path $ImageFile -ErrorAction SilentlyContinue ) )
    {
        ## Notification is made by the Windows Push Notifications User Service so we can't delete the icon
        ##Remove-Item -Path $ImageFile -Force
    }
}

