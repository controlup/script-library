#requires -version 3

<#
.SYNOPSIS

Create a read-only WPF window containing a text block showing a message, optionally including parameters passed, with or without a title.

.DETAILS

To use this script in an automated action, take a copy of it, add the _clientMetricX parameters as required and set the message text in the param() block with {0}, {1}, etc as required and set any other parameters where you don't want the default

.PARAMETER _clientMetric1

Parameter passed from ControlUp record properties and replaced in message string. Where more than one is specified, the trailing digits are sorted numerically to determine order of replacement in the message string, eg {1} would be replaced by _clientMetric2

.PARAMETER message

The message to display in the dialogue. If specifying variables with _ prefix, use {0} in the string to have it replaced with the first _ parameter numerically first, {1} for second, etc. where trailing digits are sorted numerically to determine order

.PARAMETER title

The title for the dialogue. If specified as an empty string or $null, no title bar is shown

.PARAMETER fullScreen

The dialogue with cover the entire screen (primary monitor)

.PARAMETER screenPercentage

The percentage of the screen resolution to make the dialogue dimensions

.PARAMETER backgroundColour

The colour for the solid background specified as R,G,B values

.PARAMETER textColour

The colour for the text specified as R,G,B values

.PARAMETER fontSize

The font size to use

.PARAMETER fontFamily

The font family to use

.PARAMETER showForSeconds

Show the dialogue for this number of seconds. The default is to show until dismissed

.PARAMETER noClickToClose

A left mouse click on the dialogue will not close the dialogue

.PARAMETER position

The position on the screen to place the dialogue

.PARAMETER notTopmost

Do not make the dialogue system modal

.PARAMETER noClose

Do not allow the window to be closed. It will only close if -showForSeconds is specified and that number of seconds has been reached

.EXAMPLE

& '.\Message user session.ps1' -_clientMetric0 42 -message "Your WiFi signal is weak ({0}%)"

Display a popup in the bottom right hand corner of the session which reads "Your WiFi signal is weak (42%)"

.EXAMPLE

& '.\Message user session.ps1' -message "Please logoff and back on" -title "Message from $env:Username at $(Get-Date -Format G)" -noclose true

Display a popup in the bottom right hand corner of the session with the given text and title. This popup cannot be closed.

.EXAMPLE

& '.\Message user session.ps1' -message "Please logoff and back on" -position Center -screenPercentage 50 -showForSeconds 60 -fontSize 48

Display a popup in the center of the screen in the user's session which is 50% of the screen resolution with the given text, in 48 point, but with no title.
This popup can closed by clicking on the X, pressing Alt F4 or left clicking the popup.
The popup will close automatically after 60 seconds.

.CONTEXT

Session (user session)

.NOTES

Must run in the user's session (or ControlUp console) in order for the WPF code to display a dialogue

Most attributes of the window such as text colour, size, font can be controlled via parameters.

.MODIFICATION_HISTORY:

@guyrleech  2021-05-10  Initial release
@guyrleech  2021-05-12  Fix for colour RGB array values being flattened
@guyrleech  2021-05-13  Formatting codes added to $message
@guyrleech  2022-01-11  Changed [switch] parameters to work with CU console script calling mechanism
#>

[CmdletBinding()]

Param
(
	## client metrics (or any parameter passed automagically via CU) must start with an underscore and have the number of the positional parameter in the $message string at the end (which they are sorted on before constructing the message string), eg _clientMetric2
	## if not passing any record properties via the SBA definition, delete the parameter completely
	$_clientMetric1 ,

	[string]$message , ## don't make mandatory in case not passed and CU prompts silently for it

	[AllowEmptyString()][AllowNull()]
	[string]$title , ## if null or empty string then there will be no title
	[ValidateSet('Yes', 'No', 'True', 'False')]
	[string]$fullScreen = 'No',
	[ValidateSet('Yes', 'No', 'True', 'False')]
	[string]$noClickToClose = 'No', ## left mouse click does not close the dialogue
	[ValidateSet('TopLeft', 'TopRight', 'BottomLeft', 'BottomRight', 'Centre', 'Center')]
	[string]$position = 'BottomRight' ,
	[ValidateScript({ $_ -gt 0 -and $_ -le 100 })]
	[int]$screenPercentage = 25 ,
	[string[]]$backgroundColour,
	[string[]]$textColour ,
	[int]$fontSize ,
	[string]$fontFamily ,
	[int]$showForSeconds ,
	[ValidateSet('Yes', 'No', 'True', 'False')]
	[string]$notTopmost = 'No',
	[ValidateSet('Yes', 'No', 'True', 'False')]
	[string]$noClose = 'No'
)

#region WPF
[string]$notificationWindowXAML = @'
<Window x:Name="wndMain" x:Class="UserNotifier.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:UserNotifier"
        mc:Ignorable="d"
        Title="ControlUp Notification" SizeToContent="WidthAndHeight" ResizeMode="NoResize" Background="#CDCDCD">
        <TextBlock x:Name="txtblckMain" Grid.Column="2" HorizontalAlignment="Stretch" Margin="50,50,50,50" TextWrapping="Wrap" VerticalAlignment="Stretch" FontFamily="Raleway ExtraBold" FontWeight="Bold" FontSize="36" TextAlignment="Center"/>
</Window>
'@

$VerbosePreference = $(if ( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if ( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if ( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

Function New-Form {
	Param
	(
		[Parameter(Mandatory = $true)]
		$inputXaml
	)

	$form = $null
	$inputXML = $inputXaml -replace 'mc:Ignorable="d"' , '' -replace 'x:N' , 'N' -replace '^<Win.*' , '<Window'

	[xml]$xaml = $inputXML

	if ( $xaml ) {
		$reader = New-Object -TypeName Xml.XmlNodeReader -ArgumentList $xaml

		try {
			$form = [Windows.Markup.XamlReader]::Load( $reader )
		}
		catch {
			Throw "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$_"
		}

		$xaml.SelectNodes( '//*[@Name]' ) | ForEach-Object `
		{
			Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
		}
	}
	else {
		Throw "Failed to convert input XAML to WPF XML"
	}

	$form
}
#endregion WPF

## don't make mandatory so that can be set in the param() block
if ( [string]::IsNullOrEmpty( $message ) ) {
	Throw 'Must specify message text to display'
}

if ( $fullScreen -imatch '^(Yes|True)$' ) {
	if ( $PSBoundParameters[ 'percentage' ] -and $percentage -ne 100 ) {
		Throw 'Must not specify -percentage when not 100% if -fullscreen also specified'
	}
	$screenPercentage = 100
}

Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Windows.Forms

if ( ! ( $mainForm = New-Form -inputXaml $notificationWindowXAML ) ) {
	Exit 1
}
## arrays may have been flattened if not called natively from within PowerShell so resurrect
if ( $null -ne $backgroundColour -and $backgroundColour.Count -eq 1 -and $backgroundColour[0].IndexOf( ',' ) -ge 0 ) {
	$backgroundColour = $backgroundColour -split ','
}

if ( $null -ne $textColour -and $textColour.Count -eq 1 -and $textColour[0].IndexOf( ',' ) -ge 0 ) {
	$textColour = $textColour -split ','
}

## get the underscore parameters from the parameters so we can expand the message string - put in hashtable keyed on number at the end of the parameter name so we can sort on that and check for duplicates
[hashtable]$messageStrings = @{}

ForEach ( $parameter in $PSBoundParameters.GetEnumerator() ) {
	if ( $parameter.Key -match '^_[^\d]*(\d*)$' ) { ## _clientMetric0
		try {
			$messageStrings.Add( [int]$Matches[1] , $parameter.Value )
		}
		catch {
			Throw "Already have an _ parameter ending in number $($Matches[1]) so can't use $($parameter.Key)"
		}
	}
}

Write-Verbose -Message "Got $($messageStrings.Count) parameters for message string"

if ( $message -match '\{0\}' -and $messageStrings.Count -eq 0 ) {
	Write-Warning -Message "Message string contains {0} but no record properties were passed as parameters"
}

Add-Type -TypeDefinition  @'
    using System;
    using System.Runtime.InteropServices;

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public class user32
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(String sClassName, String sAppName);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SystemParametersInfo ( uint uiAction, uint uiParam, IntPtr lpvParam, uint fWinIni);
    }
'@

## Get size of taskbar as [System.Windows.Forms.Screen]::PrimaryScreen is unreliable

[string]$taskbarWindowClass = 'Shell_TrayWnd'
[IntPtr]$hWnd = [user32]::FindWindow( $taskbarWindowClass, $null )
$taskbar = New-Object -TypeName RECT

if ( $hwnd -ne [IntPtr]::Zero ) {
	if ( ! [user32]::GetWindowRect( $hWnd , [ref]$taskbar ) ) {
		Write-Warning -Message "Failed to get window rect for handle $hWnd"
	}
	else {
		$taskbar | Select-Object -Property * | Write-Verbose
	}
}
else {
	Write-Warning -Message "Found no window with class "
}

$WPFtxtblckMain.Text = $message -f ($messageStrings.GetEnumerator() | Sort-Object -Property Key | Select-Object -ExpandProperty Value)

Write-Verbose -Message "Expanded message text is `"$($WPFtxtblckMain.Text)`""

$mainForm.WindowStyle = 'ToolWindow'

if ( [string]::IsNullOrEmpty( $title ) ) {
	$mainForm.WindowStyle = 'None' ## no title bar/control - have to close with Alt F4 if -noClickToClose specified
}
else {
	$mainForm.Title = $title
}

$mainForm.Topmost = $notTopmost -notmatch '^(yes|true)$'

$primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen

Write-Verbose -Message "Primary screen $($primaryScreen.WorkingArea.Width) x $($primaryScreen.WorkingArea.Height), working area $($primaryScreen.Bounds.Width) x $($primaryScreen.Bounds.Height)"

[int]$primaryWidth = [math]::Min( $primaryScreen.WorkingArea.Width - $primaryScreen.WorkingArea.X , $primaryScreen.Bounds.Width - $primaryScreen.Bounds.X )
[int]$primaryHeight = [math]::Min( $primaryScreen.WorkingArea.Height - $primaryScreen.WorkingArea.Y , $primaryScreen.Bounds.Height - $primaryScreen.Bounds.Y )

$WPFtxtblckMain.Width = ( $primaryWidth * $screenPercentage / 100 )
$WPFtxtblckMain.Height = ( $primaryHeight * $screenPercentage / 100 )

## can't get/set size of containing window so figure out it's size from textblock and margins
[int]$width = $WPFtxtblckMain.Width + $WPFtxtblckMain.Margin.Left + $WPFtxtblckMain.Margin.Right
[int]$height = $WPFtxtblckMain.Height + $WPFtxtblckMain.Margin.Top + $WPFtxtblckMain.Margin.Bottom

if ( ! [string]::IsNullOrEmpty( $title ) ) {
	## adjust for title bar and borders
	$height += [System.Windows.SystemParameters]::CaptionHeight + [System.Windows.SystemParameters]::BorderWidth * 2
	$width += [System.Windows.SystemParameters]::BorderWidth * 2 ## left and right borders
}

if ( $fullScreen -imatch '^(Yes|True)$' ) { ## don't adjust for margins if fullscreen
	$mainForm.Left = $mainForm.Top = 0
}
else {
	[int]$taskbarWidth = [math]::Abs( $taskbar.Right - $taskbar.Left )
	[int]$taskbarHeight = [math]::Abs( $taskbar.Bottom - $taskbar.Top )

	switch -Regex ( $position ) {
		'TopLeft' {
			$mainForm.Left = 0
			$mainForm.Top = 0
		}
		'TopRight' {
			$mainForm.Left = $primaryWidth - $Width
			$mainForm.Top = 0
		}
		'BottomLeft' {
			$mainForm.Left = 0
			$mainForm.Top = $primaryHeight - $Height - $taskbarHeight
		}
		'BottomRight' {
			$mainForm.Left = $primaryWidth - $width
			$mainForm.Top = $primaryHeight - $height - $taskbarHeight
		}
		'Centre|Center' {
			$mainForm.Left = ( $primaryWidth - $Width  ) / 2
			$mainForm.Top = ( $primaryHeight - $Height ) / 2
		}
	}
}

if ( $PSBoundParameters[ 'backgroundColour' ] ) {
	if ( $backgroundColour.Count -ne 3 ) {
		Throw "Must specify R,G,B values for background colour"
	}
	$colour = [System.Windows.Media.Color]::FromRgb( $backgroundColour[0] , $backgroundColour[1] , $backgroundColour[2] )
	$mainForm.Background = [System.Windows.Media.SolidColorBrush]::new( $colour )
}

if ( $PSBoundParameters[ 'textColour' ] ) {
	if ( $textColour.Count -ne 3 ) {
		Throw "Must specify R,G,B values for text colour"
	}
	$colour = [System.Windows.Media.Color]::FromRgb( $textColour[0] , $textColour[1] , $textColour[2] )
	$WPFtxtblckMain.Foreground = [System.Windows.Media.SolidColorBrush]::new( $colour )
}

if ( $PSBoundParameters[ 'fontSize' ] ) {
	$WPFtxtblckMain.FontSize = $fontSize
}

if ( $PSBoundParameters[ 'fontFamily' ] ) {
	$WPFtxtblckMain.FontFamily = $fontFamily
}

$script:formTimer = $null
[bool]$script:timerExpired = $false

[scriptblock]$timerBlock = `
{
	$this.Stop()
	$script:timerExpired = $true
	$mainForm.Close()
}

if ( $PSBoundParameters.ContainsKey( 'showForSeconds' ) -and $showForSeconds -gt 0 ) {
	## https://richardspowershellblog.wordpress.com/2011/07/07/a-powershell-clock/
	$mainForm.Add_SourceInitialized({
			if ( $script:formTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer ) {
				$script:formTimer.Interval = New-TimeSpan -Seconds $showForSeconds
				$script:formTimer.Add_Tick( $timerBlock )
				$script:formTimer.Start()
			}
		})
}

## allow left button to close the form unless specified not to or the form itselt can't be closed
if ( $noClickToClose -notmatch '^(Yes|True)$' -and $noClose -notmatch '^(Yes|True)$' ) {
	$mainForm.add_MouseLeftButtonUp({
			$mainForm.Close()
		})
}

if ( $noClose -match '^(Yes|True)$' ) {
	$mainForm.add_Closing({
			$_.Cancel = ! $script:timerExpired ## do not allow close unless the timer has expired otherwise the timer can't close the window
		})
}

Write-Verbose -Message "Dialog $($Width) x $($Height) at $($mainForm.Left) , $($mainForm.Top)"

$mainForm.UpdateLayout()
$result = $mainForm.ShowDialog()

Write-Verbose -Message "Dialog after is at $($mainForm.Left) , $($mainForm.Top)"

if ( $script:formTimer ) {
	$script:formTimer.Stop()
	$script:formTimer.remove_Tick( $timerBlock )
	$script:formTimer = $null
	$timerBlock = $null
	Remove-Variable -Name timerBlock -Force -Confirm:$false
}

If (!$result) {
	Write-Output -InputObject 'User has seen and closed the message.'
}

