#requires -version 3

<#
.SYNOPSIS

Create a read-only WPF window containing a text block showing a message, with or without a title.

.DETAILS

Most attributes of the window such as text colour, size, font can be controlled via parameters

.PARAMETER message

The message to display in the dialogue

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

& '.\Message user session.ps1' -message "Please logoff and back on" -title "Message from $env:Username at $(Get-Date -Format G)" -noclose

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

.MODIFICATION_HISTORY:

@guyrleech  2021-05-10  Initial release

#>

[CmdletBinding()]

Param
(
    ## client metrics (or any parameter passed automagically via CU) must start with an underscore and have the number of the positional parameter in the $message string at the end (which they are sorted on before constructing the message string), eg _clientMetric2
    ## do not have digits anywhere else in the parameter name other than at the end
    ## if not passing any record properties via the SBA definition, delete the _ parameter(s) completely
    ## to show a number without decimal places, make the parameter an [int] type
    $_clientMetric1 ,
    $_clientMetric2 ,
    $_clientMetric3 ,
    $_clientMetric4 ,
    $_clientMetric5 ,
    [string]$message = 'Your WiFi signal is weak' ,
    [AllowEmptyString()][AllowNull()]
    [string]$title , ## if null or empty string then there will be no title
    [switch]$fullScreen ,
    [ValidateScript({$_ -gt 0 -and $_ -le 100})] 
    [int]$screenPercentage = 25 ,
    [string[]]$backgroundColour ,
    [string[]]$textColour ,
    [int]$fontSize ,
    [string]$fontFamily ,
    [int]$showForSeconds ,
    [switch]$noClickToClose , ## left mouse click does not close the dialogue
    [ValidateSet('TopLeft','TopRight','BottomLeft','BottomRight','Centre','Center')]
    [string]$position = 'BottomRight' ,
    [switch]$notTopmost ,
    [switch]$noClose 
)

[string]$notificationWindowXAML = @'
<Window x:Name="wndMain" x:Class="UserNotifier.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:UserNotifier"
        mc:Ignorable="d"
        Title="ControlUp Notification" SizeToContent="WidthAndHeight" ResizeMode="NoResize" Background="#FFE61414">
        <TextBlock x:Name="txtblckMain" Grid.Column="2" HorizontalAlignment="Stretch" Margin="50,50,50,50" TextWrapping="Wrap" VerticalAlignment="Stretch" FontFamily="Raleway ExtraBold" FontWeight="Bold" FontSize="36" TextAlignment="Center"/>
</Window>
'@

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$sessionId = Get-Process -Id $pid | Select-Object -ExpandProperty SessionId

if( $sessionId -eq 0 )
{
    Throw "Notifications cannot be shown in session zero - set the script to run in the context of the users session"
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

$message = $message -f ($messageStrings.GetEnumerator() | Sort-Object -Property Key | Select-Object -ExpandProperty Value)

#Write-Verbose -Message "Expanded message text is `"$($WPFtxtblckMain.Text)`"" ##Is this used?

Write-Verbose -Message "Session id is $sessionId user name $env:USERNAME"


Function New-Form
{
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

## don't make mandatory so that can be set in the param() block
if( [string]::IsNullOrEmpty( $message ) )
{
    Throw 'Must specify message text to display'
}

if( $PSBoundParameters[ 'fullScreen' ] )
{
    if( $PSBoundParameters[ 'percentage' ] -and $percentage -ne 100 )
    {
        Throw 'Must not specify -percentage when not 100% if -fullscreen also specified'
    }
    $screenPercentage = 100
}

Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Windows.Forms

if( ! ( $mainForm = New-Form -inputXaml $notificationWindowXAML ) )
{
    Exit 1
}

if( $null -ne $backgroundColour -and $backgroundColour.Count -eq 1 -and $backgroundColour[0].IndexOf( ',' ) -ge 0 )
{
    $backgroundColour = $backgroundColour -split ','
}

if( $null -ne $textColour -and $textColour.Count -eq 1 -and $textColour[0].IndexOf( ',' ) -ge 0 )
{
    $textColour = $textColour -split ','
}

[int]$appliedDPI = Get-ItemProperty -Path 'HKCU:\Control Panel\Desktop\WindowMetrics' -Name 'AppliedDPI' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'AppliedDPI' -ErrorAction SilentlyContinue
[double]$appliedDPIScaling = $appliedDPI / 96 ## 100% DPI scaling is 96dpi

Write-Verbose -Message "Applied DPI scaling is $appliedDPIScaling"

[double]$scaleFactor = 1
switch( $appliedDPI )
{
    96  { $scaleFactor = 1 }
    144 { $scaleFactor = 2 }
}

Write-Verbose -Message "Applied DPI scaling is $appliedDPIScaling, scale factor $scaleFactor"

$WPFtxtblckMain.Text = "$message"

$mainForm.WindowStyle = 'ToolWindow'

if( [string]::IsNullOrEmpty( $title ) )
{
    $mainForm.WindowStyle = 'None' ## no title bar/control - have to close with Alt F4 if -noClickToClose specified
}
else
{
    $mainForm.Title = $title
}

$mainForm.Topmost = ! $notTopmost

$primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen

Write-Verbose -Message "Primary screen $($primaryScreen.WorkingArea.Width) x $($primaryScreen.WorkingArea.Height), working area $($primaryScreen.Bounds.Width) x $($primaryScreen.Bounds.Height)"

$WPFtxtblckMain.Width  = ( $primaryScreen.WorkingArea.Width * $screenPercentage / 100 ) 
$WPFtxtblckMain.Height = ( $primaryScreen.WorkingArea.Height * $screenPercentage / 100 ) 

## can't get/set size of containing window so figure out it's size from textblock and margins
[int]$width  = $WPFtxtblckMain.Width + $WPFtxtblckMain.Margin.Left + $WPFtxtblckMain.Margin.Right
[int]$height = $WPFtxtblckMain.Height + $WPFtxtblckMain.Margin.Top + $WPFtxtblckMain.Margin.Bottom

if( ! [string]::IsNullOrEmpty( $title ) )
{
    ## adjust for title bar and borders
    $height += [System.Windows.SystemParameters]::CaptionHeight  + [System.Windows.SystemParameters]::BorderWidth * 2
    $width +=  [System.Windows.SystemParameters]::BorderWidth * 2 ## left and right borders
}

if( $fullScreen ) ## don't adjust for margins if fullscreen
{
    $mainForm.Left = $mainForm.Top = 0
}
else
{
    switch -Regex ( $position )
    {
        'TopLeft'
        {
            $mainForm.Left = 0
            $mainForm.Top = 0
        }
        'TopRight'
        {
            $mainForm.Left = $primaryScreen.WorkingArea.Width - $Width
            $mainForm.Top = 0
        }
        'BottomLeft'
        {
            $mainForm.Left = 0
            $mainForm.Top = $primaryScreen.WorkingArea.Height - $Height
        }
        'BottomRight'
        {
            $mainForm.Left = $primaryScreen.WorkingArea.Width  - $width
            $mainForm.Top  = $primaryScreen.WorkingArea.Height - $height
        }
        'Centre|Center'
        {
            $mainForm.Left = ( $primaryScreen.WorkingArea.Width  - $Width  ) / 2
            $mainForm.Top  = ( $primaryScreen.WorkingArea.Height - $Height ) / 2
        }
    }
}

if( $PSBoundParameters[ 'backgroundColour' ] )
{
    if( $backgroundColour.Count -ne 3 )
    {
        Throw "Must specify R,G,B values for background colour"
    }
    $colour = [System.Windows.Media.Color]::FromRgb( $backgroundColour[0] , $backgroundColour[1] , $backgroundColour[2] )
    $mainForm.Background = [System.Windows.Media.SolidColorBrush]::new( $colour )
}

if( $PSBoundParameters[ 'textColour' ] )
{
    if( $textColour.Count -ne 3 )
    {
        Throw "Must specify R,G,B values for text colour"
    }
    $colour = [System.Windows.Media.Color]::FromRgb( $textColour[0] , $textColour[1] , $textColour[2] )
    $WPFtxtblckMain.Foreground = [System.Windows.Media.SolidColorBrush]::new( $colour )
}

if( $PSBoundParameters[ 'fontSize' ] )
{
    $WPFtxtblckMain.FontSize = $fontSize
}

if( $PSBoundParameters[ 'fontFamily' ] )
{
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

if( $PSBoundParameters.ContainsKey( 'showForSeconds' ) -and $showForSeconds -gt 0 )
{
    ## https://richardspowershellblog.wordpress.com/2011/07/07/a-powershell-clock/
    $mainForm.Add_SourceInitialized({
        if( $script:formTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer )
        {
            $script:formTimer.Interval = New-TimeSpan -Seconds $showForSeconds
            $script:formTimer.Add_Tick( $timerBlock )
            $script:formTimer.Start()
        }
    })
}

## allow left button to close the form unless specified not to or the form itselt can't be closed
if( ! $noClickToClose -and ! $noClose )
{
    $mainForm.add_MouseLeftButtonUp({
        $mainForm.Close()
    })
}

if( $noClose )
{
    $mainForm.add_Closing({
        $_.Cancel = ! $script:timerExpired ## do not allow close unless the timer has expired otherwise the timer can't close the window
    })
}

Write-Verbose -Message "Dialog $($Width) x $($Height) at $($mainForm.Left) , $($mainForm.Top)"

$mainForm.Left /= $scaleFactor 
$mainForm.Top  /= $scaleFactor

Write-Verbose -Message "After scaling: dialog $($Width) x $($Height) at $($mainForm.Left) , $($mainForm.Top)"

$mainForm.UpdateLayout()
$result = $mainForm.ShowDialog()

Write-Verbose -Message "Dialog after is at $($mainForm.Left) , $($mainForm.Top)"

if( $script:formTimer )
{
    $script:formTimer.Stop()
    $script:formTimer.remove_Tick( $timerBlock )
    $script:formTimer = $null
    $timerBlock = $null
    Remove-Variable -Name timerBlock -Force -Confirm:$false
}

