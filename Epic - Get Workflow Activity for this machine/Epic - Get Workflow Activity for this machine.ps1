<#
    .SYNOPSIS
        Queries Epic SystemPulse for available Attributes to query

    .DESCRIPTION
        Queries Epic SystemPulse for available Attributes to query

    .EXAMPLE
        . .\EpicSystemPulseAttributes.ps1 -$EpicCitrixServer GLHS-PA-1034 -SystemPulseServer apporchard.epic.com -StartDate 2020-03-20 -StartTime 9:00AM -EndDate 2020-03-27 -EndTime 12:00PM
        Gets all events for the user "amttye" from domain "bottheory" on this machine

    .NOTES
        This script will be run on where the local ControlUp console.

    .CONTEXT
        Console

    .MODIFICATION_HISTORY
        Created TTYE : 2020-04-02

    AUTHOR: Trentent Tye
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='EpicCitrixServer')][ValidateNotNullOrEmpty()]                 [string]$EpicCitrixServer,
    [Parameter(Mandatory=$true,HelpMessage='SystemPulseServer')][ValidateNotNullOrEmpty()]                [uri]$SystemPulseServer,
    [Parameter(Mandatory=$false,HelpMessage='Runs with saved data silently')][ValidateNotNullOrEmpty()]   [switch]$Silent
)


$verbosePreference = "continue"

Start-Transcript "$env:temp\EpicWorkflow.log"

if ($silent) {
    try {
        Test-Path "$env:temp\EpicSBA.xml"
    } catch {
        Write-Verbose "Silent switch found but saved information not. Running in non-silent mode."
        $silent = $false
        break
    }
    #Load saved data
    if (Test-Path "$env:temp\EpicSBA.xml") {
        $syncHash = Import-Clixml "$env:temp\EpicSBA.xml"
    } else {
        $silent = $false
    }
}



if (-not($silent)) {
    #region Creates the Date Selector UI
    $syncHash = [hashtable]::Synchronized(@{})
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"         
    $newRunspace.Open()     
    $newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash) 
    $psCmd = [PowerShell]::Create().AddScript({   

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        Add-Type -AssemblyName PresentationFramework
        Add-Type –assemblyName PresentationCore
        Add-Type –assemblyName WindowsBase

        $syncHash.Window = New-Object System.Windows.Window -Property @{
            WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
            MinHeight  = 600
            MinWidth   = 600
            Height     = 600
            Width      = 600
            Title      = 'Epic - Select a date range (up to 14 days)'
            Topmost    = $true
        }

        # Create a grid container with 2 rows, one for the buttons, one for the datagrid
        $Grid =  New-Object Windows.Controls.Grid
        $Row1 = New-Object Windows.Controls.RowDefinition -Property @{
            MinHeight = 500
            Height = "4*"
        }
        $Row2 = New-Object Windows.Controls.RowDefinition -Property @{
            MinHeight = 50
        }
        $column1 = New-Object Windows.Controls.ColumnDefinition -Property @{
            Width = "2*"
        }
        $column2 = New-Object Windows.Controls.ColumnDefinition -Property @{
            Width = "2*"
        }
        $grid.ColumnDefinitions.Add($column1)
        $grid.ColumnDefinitions.Add($column2)
        $grid.RowDefinitions.Add($Row1)
        $grid.RowDefinitions.Add($Row2)
        $grid.ShowGridLines = $false    # set to true to debug UI

        $syncHash.WPFCalendar = New-Object System.Windows.Controls.Calendar -Property @{
            SelectionMode  = [System.Windows.Forms.SelectionMode]::MultiSimple
            DisplayDateEnd = $(Get-Date)
        }

        $syncHash.WPFViewBox = New-Object System.Windows.Controls.Viewbox -Property @{
            Stretch          = [System.Windows.Media.Stretch]::Fill
            StretchDirection = [System.Windows.Controls.StretchDirection]::Both
            Child            = $syncHash.WPFCalendar
        }
        $syncHash.WPFViewBox.SetValue([Windows.Controls.Grid]::RowProperty,0)
        $syncHash.WPFViewBox.SetValue([Windows.Controls.Grid]::ColumnSpanProperty,2)
        $syncHash.SelectDates = New-Object Windows.Controls.Button -Property @{
            Content  = "Select Dates"
            MinWidth    = 180
            MinHeight   = 50
            Margin   = "5,5,5,10"
            HorizontalAlignment = "Right"
            FontSize = 24
        }
        $syncHash.Button1ViewBox = New-Object System.Windows.Controls.Viewbox -Property @{
            Stretch          = [System.Windows.Media.Stretch]::Uniform
            StretchDirection = [System.Windows.Controls.StretchDirection]::Both
            Child            = $syncHash.SelectDates
            HorizontalAlignment = "Right"
        }
        $syncHash.Button1ViewBox.SetValue([Windows.Controls.Grid]::RowProperty,1)
        $syncHash.Button1ViewBox.SetValue([Windows.Controls.Grid]::ColumnProperty,1)

        $syncHash.Cancel = New-Object Windows.Controls.Button -Property @{
            Content  = "Cancel"
            MinWidth    = 180
            MinHeight   = 50
            Margin   = "5,5,5,10"
            HorizontalAlignment = "Right"
            FontSize = 24
        }

        $syncHash.CancelButtonViewBox = New-Object System.Windows.Controls.Viewbox -Property @{
            Stretch          = [System.Windows.Media.Stretch]::Uniform
            StretchDirection = [System.Windows.Controls.StretchDirection]::Both
            Child            = $syncHash.Cancel
            HorizontalAlignment = "Left"

        }
        $syncHash.CancelButtonViewBox.SetValue([Windows.Controls.Grid]::RowProperty,1)
        $syncHash.CancelButtonViewBox.SetValue([Windows.Controls.Grid]::ColumnProperty,0)
    
        #$syncHash.SelectDates.SetValue([Windows.Controls.Grid]::RowProperty,1)
        $grid.AddChild($syncHash.WPFViewBox)
        $grid.AddChild($syncHash.Button1ViewBox)
        $grid.AddChild($syncHash.CancelButtonViewBox)

        $syncHash.Window.Content = $Grid

        $syncHash.SelectedDatesPressed = $false
        $syncHash.CancelPressed = $false
        $syncHash.Closed = $false
    
        $syncHash.SelectDates.Add_Click({ 
            $syncHash.SelectedDatesPressed = $true
            $syncHash.Window.Close()
        })

        $syncHash.Cancel.Add_Click({
            $syncHash.CancelPressed = $true
            $syncHash.Window.Close()
            exit
        })

        $syncHash.Window.Add_Closing({
            $syncHash.Closed = $true
        })
        #$VerbosePreference = "continue"

        #When you are making a selection the layout gets updated. We set a variable to be $true ($syncHash.layoutUpdated) that we can check with "add_SelectedDatesChanged" method
        #add_SelectedDatesChanged fire when a date modification is made. When trying to limit it to 14 days we'd get infiniteloops because the 
        #SelectedDates.Clear() method would trigger the SelectedDatesChanged method.  By doing a variable check (and reset when we finish our modifications)
        #we can prevent the loop.  Hence the check if the variable is $true (UI has been updated) and setting it to $false at the end of our method (to break the loop)
        $syncHash.WPFCalendar.add_LayoutUpdated({
            $syncHash.LayoutUpdated = $true
            Write-Verbose "add_LayoutUpdated"
        })

        $syncHash.WPFCalendar.add_SelectedDatesChanged({
            if ($syncHash.layoutUpdated -eq $true) {
                if ($syncHash.WPFCalendar.SelectedDates.count -ne 0) {
                    $syncHash.StartDate = $syncHash.WPFCalendar.SelectedDates[0]
                    $syncHash.EndDate = $syncHash.WPFCalendar.SelectedDates[-1]
                    Write-Verbose "StartDate object reset to $($syncHash.StartDate)"    #debug by looking at $psCmd.Streams
                    Write-Verbose "EndDate object reset to $($syncHash.EndDate)"
                }
                Write-Verbose "SyncHash Date0: $($syncHash.StartDate)"
                Write-Verbose "SyncHash EndDate: $($syncHash.EndDate)"
            
                if ($syncHash.WPFCalendar.SelectedDates.count -ge 15) {
                    (New-Object Media.SoundPlayer "$($env:Windir)\Media\notify.wav").Play()
                    Write-Verbose "More than 14 days worth of data requested."
                    if ($syncHash.StartDate.AddDays(14) -gt $(Get-Date)) {
                        $syncHash.EndDate = $(Get-Date)
                    } else {
                        $syncHash.EndDate = $syncHash.StartDate.AddDays(13)
                    }
                    Write-Verbose "SyncHash Date1 Add: $($syncHash.StartDate.AddDays(13))"
                    $script:otherObj = $syncHash.WPFCalendar.Dispatcher.InvokeAsync([action]{
                        Write-Verbose "Clearing selected dates"
                        $syncHash.WPFCalendar.SelectedDates.Clear()
                        $syncHash.WPFCalendar.SelectedDates.AddRange($syncHash.StartDate,$syncHash.EndDate)
                    }, "Normal")
                }
                $syncHash.layoutUpdated = $false
            }
        })
        #$syncHash.Window.AddChild($syncHash.WPFViewBox)
        $syncHash.Window.ShowDialog()
    
        })
    $psCmd.Runspace = $newRunspace
    $data = $psCmd.BeginInvoke()
    #endregion Creates the Date Selector UI

    while ($data.IsCompleted -notcontains $true) {} #wait for selection
}

#################### Script Starts Here ##########################
function DateTime-ToEpicTimeCode ([dateTime]$dateTime) {
    return $dateTime.ToString("yyyy-MM-ddTHH:mm:sszzz")
}

function ConvertEpicJSON-toObjects ($JsonString) {
    #region create custom object for JSON data to get injested.  Much faster than relying on ConvertFrom-Json...
    Add-Type -TypeDefinition @"
namespace Data
{
    using System;
    using System.Collections.Generic;

    using System.Globalization;

public class Attribute
{
    public int ID { get; set; }
    public int Severity { get; set; }
    public string Value { get; set; }
}

public class Log
{
    public List<Attribute> Attributes { get; set; }
    public string Context { get; set; }
    public int ResourceID { get; set; }
    public DateTimeOffset Timestamp { get; set; }
}

public class RootObject
{
    public List<Log> Logs { get; set; }
}
}
"@

    $json = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer
    $json.MaxJsonLength = 104857600 #100mb as bytes, default is 2mb

    $QueryStartTime = Get-Date
    $data = $json.Deserialize($JsonString, [Data.RootObject])
    $QueryEndTime = Get-Date
    Write-Verbose "Creating object from JSON data took: $($(New-TimeSpan -Start $QueryStartTime -End $QueryEndTime).TotalSeconds) Seconds"
    return $data
    #endregion  create custom object for JSON data to get injested.  Much faster than relying on ConvertFrom-Json...
}

function Generate-PleaseWaitWindow  {
    if ($silent) {
        Write-Verbose "Generate PleaseWaitWindow in silent mode"
    } else {
        #check to see if the Please Wait Window is already running
        if ((Get-Runspace -Name PleaseWait).RunspaceAvailability -notcontains "Busy") {
            Write-Verbose "Creating `"Please Wait`" window."

            $Script:PleaseWaitHashTable = [hashtable]::Synchronized(@{})
            $newRunspace =[runspacefactory]::CreateRunspace()
            $newRunspace.ApartmentState = "STA"
            $newRunspace.ThreadOptions = "ReuseThread"
            $newRunspace.Name = "PleaseWait"    
            $newRunspace.Open()     
            $newRunspace.SessionStateProxy.SetVariable("PleaseWaitHashTable",$Script:PleaseWaitHashTable) 
            $psCmd = [PowerShell]::Create().AddScript({   

                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing
                Add-Type -AssemblyName PresentationFramework
                Add-Type –assemblyName PresentationCore
                Add-Type –assemblyName WindowsBase

                $PleaseWaitHashTable.Window = New-Object System.Windows.Window -Property @{
                    WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
                    Height     = 300
                    Width      = 300
                    Title      = 'Please Wait'
                    Topmost    = $true
                    ResizeMode = "NoResize"
                }

                # Create a grid container with 2 rows, one for the buttons, one for the datagrid
                $PleaseWaitHashTable.grid = New-Object Windows.Controls.Grid
                $Row1 = New-Object Windows.Controls.RowDefinition -Property @{ MinHeight = 50; Height = "4*" }
                $Row2 = New-Object Windows.Controls.RowDefinition -Property @{ MinHeight = 50 }
                $Row3 = New-Object Windows.Controls.RowDefinition -Property @{ MinHeight = 50 }

                $PleaseWaitHashTable.grid.RowDefinitions.Add($Row1)
                $PleaseWaitHashTable.grid.RowDefinitions.Add($Row2)
                $PleaseWaitHashTable.grid.RowDefinitions.Add($Row3)
                $PleaseWaitHashTable.grid.ShowGridLines = $true    # set to true to debug UI

                $PleaseWaitHashTable.StackPanel = [System.Windows.Controls.StackPanel]::new()
                $PleaseWaitHashTable.StackPanel.Background = "white"
                $PleaseWaitHashTable.StackPanel.Orientation = "Horizontal"
                $PleaseWaitHashTable.StackPanel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center

                [System.Windows.NameScope]::SetNameScope($PleaseWaitHashTable.StackPanel, $(New-Object System.Windows.NameScope) )

                $PleaseWaitHashTable.CancelButton = New-Object Windows.Controls.Button -Property @{
                    Content  = "Cancel"
                    MinWidth    = 180
                    MinHeight   = 10
                    Margin   = "3,5,3,5"
                    HorizontalAlignment = "Center"
                    FontSize = 14
                }
                $PleaseWaitHashTable.CancelButton.SetValue([Windows.Controls.Grid]::RowProperty,2)

                $PleaseWaitHashTable.Message = [System.Windows.Controls.TextBlock]::new()
                $PleaseWaitHashTable.Message.Text = ""
                $PleaseWaitHashTable.Message.FontSize = 12
                $PleaseWaitHashTable.Message.TextWrapping = [System.Windows.TextWrapping]::Wrap
                $PleaseWaitHashTable.Message.SetValue([Windows.Controls.Grid]::RowProperty,1)

                #Orange Line
                $PleaseWaitHashTable.lineOrange = [System.Windows.Shapes.Line]::new()
                $PleaseWaitHashTable.lineOrange.X1 = 0
                $PleaseWaitHashTable.lineOrange.Y1 = 140
                $PleaseWaitHashTable.lineOrange.X2 = 0
                $PleaseWaitHashTable.lineOrange.Y2 = 0
                $PleaseWaitHashTable.lineOrange.Stroke = "Orange"
                $PleaseWaitHashTable.lineOrange.StrokeThickness = 30
                $PleaseWaitHashTable.lineOrange.Name = "OrangeLine"

                $PleaseWaitHashTable.StackPanel.RegisterName($PleaseWaitHashTable.lineOrange.Name, $PleaseWaitHashTable.lineOrange)
                $PleaseWaitHashTable.StackPanel.Children.Add($PleaseWaitHashTable.lineOrange)

                #Green Line
                $PleaseWaitHashTable.lineGreen = [System.Windows.Shapes.Line]::new()
                $PleaseWaitHashTable.lineGreen.X1 = 30
                $PleaseWaitHashTable.lineGreen.Y1 = 140
                $PleaseWaitHashTable.lineGreen.X2 = 30
                $PleaseWaitHashTable.lineGreen.Y2 = 0
                $PleaseWaitHashTable.lineGreen.Stroke = "Green"
                $PleaseWaitHashTable.lineGreen.StrokeThickness = 30
                $PleaseWaitHashTable.lineGreen.Name = "GreenLine"

                $PleaseWaitHashTable.StackPanel.RegisterName($PleaseWaitHashTable.lineGreen.Name, $PleaseWaitHashTable.lineGreen)
                $PleaseWaitHashTable.StackPanel.Children.Add($PleaseWaitHashTable.lineGreen)

                #Blue Line
                $PleaseWaitHashTable.lineBlue = [System.Windows.Shapes.Line]::new()
                $PleaseWaitHashTable.lineBlue.X1 = 30
                $PleaseWaitHashTable.lineBlue.Y1 = 140
                $PleaseWaitHashTable.lineBlue.X2 = 30
                $PleaseWaitHashTable.lineBlue.Y2 = 0
                $PleaseWaitHashTable.lineBlue.Stroke = "Blue"
                $PleaseWaitHashTable.lineBlue.StrokeThickness = 30
                $PleaseWaitHashTable.lineBlue.Name = "BlueLine"

                $PleaseWaitHashTable.StackPanel.RegisterName($PleaseWaitHashTable.lineBlue.Name, $PleaseWaitHashTable.lineBlue)
                $PleaseWaitHashTable.StackPanel.Children.Add($PleaseWaitHashTable.lineBlue)

                #Red Line
                $PleaseWaitHashTable.lineRed = [System.Windows.Shapes.Line]::new()
                $PleaseWaitHashTable.lineRed.X1 = 30
                $PleaseWaitHashTable.lineRed.Y1 = 140
                $PleaseWaitHashTable.lineRed.X2 = 30
                $PleaseWaitHashTable.lineRed.Y2 = 0
                $PleaseWaitHashTable.lineRed.Stroke = "Red"
                $PleaseWaitHashTable.lineRed.StrokeThickness = 30
                $PleaseWaitHashTable.lineRed.Name = "RedLine"

                $PleaseWaitHashTable.StackPanel.RegisterName($PleaseWaitHashTable.lineRed.Name, $PleaseWaitHashTable.lineRed)
                $PleaseWaitHashTable.StackPanel.Children.Add($PleaseWaitHashTable.lineRed)

                # Create some animations and a storyboard.
                $OrangeLineAnimation = New-Object System.Windows.Media.Animation.DoubleAnimation(100,120, $(New-Object System.Windows.Duration($([System.TimeSpan]::FromSeconds(1.5)) )))
                [System.Windows.Media.Animation.Storyboard]::SetTargetName($OrangeLineAnimation,$PleaseWaitHashTable.lineOrange.Name)
                [System.Windows.Media.Animation.Storyboard]::SetTargetProperty($OrangeLineAnimation,$(New-Object System.Windows.PropertyPath([System.Windows.Shapes.Line]::Y2Property)))
                $OrangeLineAnimation.AutoReverse = $true
                $OrangeLineAnimation.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever

                $GreenLineAnimation = New-Object System.Windows.Media.Animation.DoubleAnimation(80,110, $(New-Object System.Windows.Duration($([System.TimeSpan]::FromSeconds(1.5)) )))
                [System.Windows.Media.Animation.Storyboard]::SetTargetName($GreenLineAnimation,$PleaseWaitHashTable.lineGreen.Name)
                [System.Windows.Media.Animation.Storyboard]::SetTargetProperty($GreenLineAnimation,$(New-Object System.Windows.PropertyPath([System.Windows.Shapes.Line]::Y2Property)))
                $GreenLineAnimation.AutoReverse = $true
                $GreenLineAnimation.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever

                $BlueLineAnimation = New-Object System.Windows.Media.Animation.DoubleAnimation(60,100, $(New-Object System.Windows.Duration($([System.TimeSpan]::FromSeconds(1.5)) )))
                [System.Windows.Media.Animation.Storyboard]::SetTargetName($BlueLineAnimation,$PleaseWaitHashTable.lineBlue.Name)
                [System.Windows.Media.Animation.Storyboard]::SetTargetProperty($BlueLineAnimation,$(New-Object System.Windows.PropertyPath([System.Windows.Shapes.Line]::Y2Property)))
                $BlueLineAnimation.AutoReverse = $true
                $BlueLineAnimation.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever

                $RedLineAnimation = New-Object System.Windows.Media.Animation.DoubleAnimation(40,90, $(New-Object System.Windows.Duration($([System.TimeSpan]::FromSeconds(1.5)) )))
                [System.Windows.Media.Animation.Storyboard]::SetTargetName($RedLineAnimation,$PleaseWaitHashTable.lineRed.Name)
                [System.Windows.Media.Animation.Storyboard]::SetTargetProperty($RedLineAnimation,$(New-Object System.Windows.PropertyPath([System.Windows.Shapes.Line]::Y2Property)))
                $RedLineAnimation.AutoReverse = $true
                $RedLineAnimation.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever


                $PleaseWaitHashTable.StoryBoard = New-Object System.Windows.Media.Animation.Storyboard

                [void]$PleaseWaitHashTable.StoryBoard.Children.Add($OrangeLineAnimation)
                [void]$PleaseWaitHashTable.StoryBoard.Children.Add($GreenLineAnimation)
                [void]$PleaseWaitHashTable.StoryBoard.Children.Add($BlueLineAnimation)
                [void]$PleaseWaitHashTable.StoryBoard.Children.Add($RedLineAnimation)

                $PleaseWaitHashTable.grid.AddChild($PleaseWaitHashTable.StackPanel)
                $PleaseWaitHashTable.grid.AddChild($PleaseWaitHashTable.CancelButton)
                $PleaseWaitHashTable.grid.AddChild($PleaseWaitHashTable.Message)
                $PleaseWaitHashTable.Window.Content = $PleaseWaitHashTable.Grid
                $PleaseWaitHashTable.StoryBoard.Begin($PleaseWaitHashTable.StackPanel)

                $PleaseWaitHashTable.Window.ShowDialog()

                })
            $psCmd.Runspace = $newRunspace
            $data = $psCmd.BeginInvoke()
        } else {
            Write-Verbose "Detected Running `"Please Wait`" window.  Skipping creation"
        }
    }
}

function Update-PleaseWaitWindow([string]$message)  {
    if ($silent) {
        Write-Verbose "Update-PleaseWaitWindow in silent mode"
    } else {

        #Needs to be set to a global variable to pass it into the action
        $script:message = $message
        if ((Get-Runspace -Name PleaseWait).RunspaceAvailability -contains "Busy") {
            if (-not($Script:PleaseWaitHashTable.Window.IsVisible)) {
                Sleep -Milliseconds 500
            }
            Write-Verbose "Updating Message: $Script:message"
                $Script:PleaseWaitHashTable.Message.Dispatcher.InvokeAsync([action]{
                $Script:PleaseWaitHashTable.Run = New-Object System.Windows.Documents.Run
                $Script:PleaseWaitHashTable.Run.Text = "$Script:message"
                $Script:PleaseWaitHashTable.Message.Inlines.Clear()
                $Script:PleaseWaitHashTable.Message.Inlines.Add($Script:PleaseWaitHashTable.Run)
            },
            "Normal"
            )
        } else {
            Write-Verbose "`"Please Wait`" window not detected. Aborting message send"
        }
    }
}

function Close-PleaseWaitWindow  {
    if ($silent) {
        Write-Verbose "Close-PleaseWaitWindow in silent mode"
    } else {
        if ((Get-Runspace -Name PleaseWait).RunspaceAvailability -contains "Busy") {
            Write-Verbose "Closing `"Please Wait`" Window..."
                $Script:PleaseWaitHashTable.Window.Dispatcher.InvokeAsync([action]{
                $Script:PleaseWaitHashTable.Window.Close()
            },
            "Normal"
            )
        } else {
            Write-Verbose "`"Please Wait`" window not detected. Aborting close command"
        }
    }
}



#Set TLS to version 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;

# Most internal servers have valid SSL Cert, thus ignore them
# https://stackoverflow.com/questions/11696944/powershell-v3-invoke-webrequest-https-error
add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

#optimize performance by ignoring GUI updates
$ProgressPreference = "silentlycontinue"

$startDate = $syncHash.StartDate
$endDate = ($syncHash.EndDate).AddDays(1) #add one days to encompass the selected day in the UI, otherwise it does it for 00:00:00 that day
$EpicStartTimeCodeWeb = $startDate.ToString("yyyy-MM-ddTHH:mm:sszzz")
$EpicEndTimeCodeWeb = $endDate.ToString("yyyy-MM-ddTHH:mm:sszzz")
Write-Verbose "Selected Date range: $startDate       - $endDate"
Write-Verbose "Epic Timecodes     : $EpicStartTimeCodeWeb - $EpicEndTimeCodeWeb"



if (($syncHash.CancelPressed -eq $false) -and ($syncHash.SelectedDatesPressed -eq $false)) {
    exit
    #exit button pressed
}

if (($syncHash.CancelPressed -eq $true)) {
    exit
    #cancel button pressed
}



if ($syncHash.SelectedDatesPressed -eq $true -or $silent -eq $true) {
    #optimize performance by ignoring GUI updates
    $ProgressPreference = "silentlycontinue"

    #generate Please Wait UI
    Write-Verbose "Generate Please Wait UI"
    Generate-PleaseWaitWindow
    

    If ($startDate -gt $endDate) {
        Write-Error "Start or End dates not aligned correctly"
    }
    Update-PleaseWaitWindow -message "Connecting to System Pulse server: $($SystemPulseServer)"

    $headers =  @{"Accept"="application/json"}
    #do we need to authenticate?
    try {
        $LogonTest = Invoke-RestMethod -Headers $headers -Method Get -Uri "$($SystemPulseServer.Scheme)://$($SystemPulseServer.Host)/SystemPulse/Services/DataProviderService.svc/api/v1/ResourceTypes" -UseBasicParsing
    } catch {
        $exception = $error[0]
        if ($exception.Exception.Response.StatusCode.value__ -eq 401) {
            Update-PleaseWaitWindow -message "Authentication required."
            if (-not(Test-Path "$env:temp\EpicCreds.xml")) {
                Write-Verbose "Saving new credentials..."
                $EpicCreds = Get-Credential -Message "Enter the credentials of an account that can authenticate to the SystemPulse server"
                $EpicCreds | Export-Clixml "$env:temp\EpicCreds.xml"
            } else {
                Update-PleaseWaitWindow -message "Found saved credentials."
                Write-Verbose "Found existing credentials. Importing..."
                $EpicCreds = Import-Clixml "$env:temp\EpicCreds.xml"
            }
            if ($EpicCreds.password.length -eq 0) {
                Remove-Item "$env:temp\EpicCreds.xml"  #If no password detected then remove the invalid cred file.
            }
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($EpicCreds.Password)
            $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($EpicCreds.UserName):$([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR))"))
            $headers += @{Authorization="Basic $encodedCredentials"}
            try {
                #try again with creds this time
                $LogonTest = Invoke-RestMethod -Headers $headers -Method Get -Uri "$($SystemPulseServer.Scheme)://$($SystemPulseServer.Host)/SystemPulse/Services/DataProviderService.svc/api/v1/ResourceTypes" -UseBasicParsing
            } catch {
                Update-PleaseWaitWindow -message "Authentication Failed."
                Remove-Item "$env:temp\EpicCreds.xml"
                Write-Error "Authentication failed! Exception code: $($exception.Exception.Response.StatusCode.value__)"
                break
            }

        } else {
            Write-Error "Unknown Status Code found: $($exception.Exception.Response.StatusCode.value__)"
            break
        }
    }

    Update-PleaseWaitWindow -message "Connected successfully."

    #Get Resource Type "Windows Host"
    Update-PleaseWaitWindow -message "Querying resource types"
    $resourceTypes = Invoke-RestMethod -Headers $headers -Method Get -Uri "$($SystemPulseServer.Scheme)://$($SystemPulseServer.Host)/SystemPulse/Services/DataProviderService.svc/api/v1/ResourceTypes" -UseBasicParsing
    Write-Verbose "Resource Types Discovered:"
    Write-Verbose "$($resourceTypes | Out-String)"

    #Find Windows Host ID number
    Update-PleaseWaitWindow -message "Querying Host ID number"
    $resourceTypeId = $resourceTypes | Where {$_.Name -eq "Epic Environment"}
    Write-Verbose "Epic Environment ID:"
    Write-Verbose "$($resourceTypeId.Id | Out-String)"


    ## Now we need to get the logTypes for this resource.  This resource is a part of a resourceTypeID that we can query. Apparently, all resources of a certain type should have the same logs available to query from them:
    Update-PleaseWaitWindow -message "Querying Log Types"
    $LogTypes = Invoke-RestMethod -Headers $headers -Method Get -Uri "$($SystemPulseServer.Scheme)://$($SystemPulseServer.Host)/SystemPulse/Services/DataProviderService.svc/api/v1/LogTypes?ResourceTypeID=$($resourceTypeId.Id)" -UseBasicParsing
    Write-Verbose "Found $($LogTypes.count) Log Types"

    ## Select log type for the Hyperspace workflow response times by client
    Update-PleaseWaitWindow -message "Querying Specific Log Type"
    $LogType = $LogTypes | where {$_.Description -eq "Hyperspace workflow step response times by client"}
    Write-Verbose "Hyperspace workflow step response times by client information:"
    Write-Verbose "$($LogType | Out-String)"

    ## Get Attributes
    Update-PleaseWaitWindow -message "Querying Attributes"
    $Attributes = Invoke-RestMethod -Headers $headers -Method Get -Uri "$($SystemPulseServer.Scheme)://$($SystemPulseServer.Host)/SystemPulse/Services/DataProviderService.svc/api/v1/Attributes/LogTypeID/$($LogType.Id)" -UseBasicParsing
    Write-Verbose "Found $($Attributes.count) Attributes"
    $WorkflowStepAttribute = $Attributes | Where {$_.Name -eq "WorkflowStep"}
    Write-Verbose "$($WorkflowStepAttribute | Out-String)"

    #region Async Download
    
    #region Synchronous download of SystemPulse data  -- using C# methods for performance
    Write-Verbose "Setting up Synchronous Requests"
    Update-PleaseWaitWindow -message "Starting synchronous download of data"
    $url = "$($SystemPulseServer.Scheme)://$($SystemPulseServer.Host)/SystemPulse/Services/DataProviderService.svc/api/v2/Data/LogTypeID/$($LogType.Id)/StartTime/$($EpicStartTimeCodeWeb)/EndTime/$($EpicEndTimeCodeWeb)?Contexts=$EpicCitrixServer"
    $start_time = Get-Date
    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add("Accept","application/json")
    $wc.Headers.Add("Authorization","Basic $encodedCredentials")
    
    <#
    $req = [System.Net.WebRequest]::Create($url)
    $req.Method = "GET"
    $req.Accept = "application/json"
    $req.Headers.Add("Authorization","Basic $encodedCredentials")
    [System.Net.WebResponse]$resp = $req.GetResponse()
    #>
    $responseLength = $resp.ContentLength
    $start_time = Get-Date
    $downloadedData = $wc.DownloadString($url)
    $end_time = Get-Date
    Write-Verbose "Synchronous download of Data took: $($(New-TimeSpan -Start $start_time -End $end_time).TotalSeconds) Seconds"
    #endregion Synchronous download of SystemPulse data
    $data = ConvertEpicJSON-toObjects -JsonString $downloadedData

    

    <#   ## for debugging if things go south
    $QueryStartTime = Get-Date
    $Data = Invoke-RestMethod -Headers $headers -Method Get -Uri "$($SystemPulseServer.Scheme)://$($SystemPulseServer.Host)/SystemPulse/Services/DataProviderService.svc/api/v2/Data/LogTypeID/$($LogType.Id)/StartTime/$($EpicStartTimeCodeWeb)/EndTime/$($EpicEndTimeCodeWeb)?Contexts=$EpicCitrixServer" -UseBasicParsing
    $QueryEndTime = Get-Date
    #>
    
    Update-PleaseWaitWindow -message "Generating attributes table"
    $AttributesTable = New-Object -TypeName "Collections.Generic.List[Object]"
    foreach ($attrib in $Attributes) {
        Add-Member -InputObject $AttributesTable -NotePropertyMembers @{
            $attrib.Id = $attrib.DisplayName
        }
    }


    $referenceObject = New-Object -TypeName psobject
    Add-Member -InputObject $referenceObject -NotePropertyMembers @{
        Time = 0
    }
    foreach ($attribute in $Data.logs[0].Attributes) {
        foreach ($id in $attribute) {
            $idName = ($Attributes.Where({$_.Id -eq $Attribute.Id})).DisplayName
            Add-Member -InputObject $referenceObject -NotePropertyMembers @{
                $idName = 0
            }
        }
    }
    Write-Verbose "Reference Object:"
    Write-Verbose "$($referenceObject | Out-String)"




    ## IF LIST WORKFLOW LIST:
    Update-PleaseWaitWindow -message "Getting list of workflow steps"
    $allWorkflowSteps = $data.logs.Attributes.Where{$_.Id -eq 331}.Value | Sort-Object -Unique
    
    $title = "Workflow Step"
    if (-not($silent)) {
        Close-PleaseWaitWindow
        <#
        $objectCollection=@()
        $allWorkflowSteps | ForEach-Object {
            $object = New-Object PSObject
            Add-Member -InputObject $object -MemberType NoteProperty -Name "$title" -Value "$_"
            $objectCollection += $object
        }
        #>
        $WorkFlowStepsList = $allWorkflowSteps | Out-GridView -Title "List of Workflows from $($StartDate.ToShortDateString()) to $($EndDate.ToShortDateString())" -PassThru
    } else {
        $workflowStepsList = $syncHash.WorkFlowStepsList
    }

    #Save Selections for next run
    if (-not($Silent)) {
        $SaveChoices = @{}
        $SaveChoices.StartDate = $syncHash.StartDate
        $SaveChoices.EndDate = $syncHash.EndDate
        $SaveChoices.WorkFlowStepsList = $WorkFlowStepsList
        Export-Clixml -InputObject $SaveChoices -Path "$env:temp\EpicSBA.xml" -Force
    }

    Generate-PleaseWaitWindow
    Update-PleaseWaitWindow -message "Filtering data to selected workflow steps"


    #Filter down Data to only the workflow steps wanted
    Write-Verbose "Workflow Steps:"
    foreach ($step in $WorkFlowStepsList) {
        Write-Verbose "$step"
    }

    $measureStartTime = Get-Date
    $PerClient = New-Object -TypeName "Collections.Generic.List[Object]"
    foreach ($log in $Data.logs) {
        foreach ($step in $WorkFlowStepsList) {
            if ($log.Attributes.Where{$_.Id -eq 331}.Value -eq "$step") {   ### check to see if contains is faster
                #$log.Attributes[4].Value
                #$log.Attributes
                $result = $referenceObject.psobject.Copy()
                $result.Time = $([System.DateTimeOffset]$log.Timestamp)
                foreach ($id in $log.attributes.Where{$_.Value.length -ne 0}) {
                    $idName = $AttributesTable.($id.Id)
                    $result.$($idName) = $($Id.Value)
                }
                $null = $PerClient.Add($result)
            }
        }
    }
    $measureEndTime = Get-Date
    Write-Verbose "Created PerClient object in $($(New-TimeSpan -Start $measureStartTime -End $measureEndTime).TotalSeconds)"
    Write-Verbose "Number of PerClient Objects : $($PerClient.Count)"

    if ($PerClient.count -eq 0) {  #put up a toast notification that no results were returned
        $app = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] ##need these two nulls to preload the assemblies
        $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
        $Template = [Windows.UI.Notifications.ToastTemplateType]::ToastImageAndText01
        [xml]$ToastTemplate = @"
        <toast launch="app-defined-string">
          <visual>
            <binding template="ToastGeneric">
              <text>ControlUp SBA</text>
              <text>No results returned.</text>
            </binding>
          </visual>
          <actions>
            <action activationType="background" content="OK" arguments="later"/>
          </actions>
        </toast>
"@
        $ToastXml = New-Object -TypeName Windows.Data.Xml.Dom.XmlDocument
        $ToastXml.LoadXml($ToastTemplate.OuterXml)
        $notify = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($app)
        $notify.Show($ToastXml)
        break
    }


    #HTML Color Codes
    $HTMLColor = @{
    "Silver" ="#C0C0C0"
    "Grey"   ="#808080"
    "Red"    ="#FF0000"
    "Maroon" ="#800000"
    "Yellow" ="#FFFF00"
    "Gold"  ="#FFD700"
    "Lime"   ="#00FF00"
    "Green"  ="#008000"
    "Aqua"   ="#00FFFF"
    "Teal"   ="#008080"
    "Blue"   ="#0000FF"
    "Navy"   ="#000080"
    "Fuchsia"="#FF00FF"
    "Purple" ="#800080"
    "firebrick"="#B22222"
    "springgreen"="#00FF7F"
    "lightseagreen"="#20B2AA"
    "midnightblue"="#191970"
    "darkviolet"="#9400D3"
    "darkslategray"="#2F4F4F"
    "goldenrod"="#DAA520"
    }



    Update-PleaseWaitWindow -message "Creating chart data"
    Write-Verbose "Creating Chart Template"
    $chartTemplate = '{"series":[{"name":"templateName","data":[[1583253600000,26],[1583253900000,26],[1583254200000,26],[1583254500000,27]]}],"chart":{"height":"400px","type":"line","zoom":{"enabled":true}},"dataLabels":{"enabled":false},"stroke":{"curve":"straight","width":2},"title":{"text":"templatetext","align":"left"},"grid":{"row":{"colors":["#f3f3f3","transparent"],"opacity":0.5}},"xaxis":{"type":"datetime"},"tooltip":{"x":{"format":"hh:mm TT"}},"colors":["#2E93fA", "#66DA26", "#546E7A", "#E91E63", "#FF9800","#00A100","#128FD9","#FFB200","#FF0000","#545454","#800080","#ffff00","#000000"]}'
    $chartTemplateObject = $chartTemplate | ConvertFrom-Json
    $chartTemplateObject.title.text = "$EpicCitrixServer"

    $1970 = New-Object System.DateTime (1970, 1, 1, 0, 0, 0, [System.DateTimeKind]::Utc)
    $charts = New-Object -TypeName "Collections.Generic.List[Object]"
    $chartData = New-Object -TypeName "Collections.Generic.List[Object]"
    $noteProperties = (Get-Member -InputObject $PerClient[0]).where{$_.MemberType -eq "NoteProperty" -and $_.Name -ne "Time" -and $_.Name -ne "Workflow Step"}.Name
    Write-Verbose "Note Properties:"
    Write-Verbose "$($noteProperties | Out-String)"
    foreach ($step in $WorkFlowStepsList) {
        $chartTemplateObject = $chartTemplate | ConvertFrom-Json
        $chartTemplateObject.title.text = $step
        $stepFilter = $PerClient.Where({$_.'Workflow Step' -eq $step})
        $stepFilter = $stepFilter | Sort-Object -Property Time
        $stepFilter.count
        foreach ($noteProperty in $noteProperties) {
            $chartData = New-Object -TypeName "Collections.Generic.List[Object]"
            if ($stepFilter.$noteProperty.where{$_ -ne 0}) { #if a property is all zero's skip adding it to the chart.
                if ($chartTemplateObject.series[0].name -ne "templateName") {
                    $chartTemplateObject.series += ($chartTemplateObject.series[0].psobject.Copy())
                    $chartTemplateObject.series[-1].name = $noteProperty
                    foreach($item in $stepFilter)
                    {
                        $epoch = ($item.time.Ticks - ($1970).Ticks) / 10000
                        [void]$chartData.Add(($epoch,$($item.$noteProperty)))
                    }
                    [array]$chartTemplateObject.series[-1].data = $chartData
                } else {
                    $chartTemplateObject.series[0].name = $noteProperty
                    $chartTemplateObject.series[0].data.Clear()
                    foreach($item in $stepFilter)
                    {
                        $epoch = ($item.time.Ticks - ($1970).Ticks) / 10000
                        [void]$chartData.Add(($epoch,$($item.$noteProperty)))
                    }
                    [array]$chartTemplateObject.series[0].data = $chartData
                }
            }
        }
        #set colors
        $colors = New-Object -TypeName "Collections.Generic.List[Object]"
        
        Foreach ($seriesName in $chartTemplateObject.series.name) {
            switch -Wildcard ($seriesName) {
                "Average Block Reads"              {Write-Host "$seriesName"; $colors.Add($HTMLColor['Lime'])}
                "Average Client CPU Time (sec)*"    {Write-Host "$seriesName"; $colors.Add($HTMLColor['Blue'])}
                "Average DB Network Time (VB)*"     {Write-Host "$seriesName"; $colors.Add($HTMLColor['Grey'])}
                "Average DB Network Time (Web)*"    {Write-Host "$seriesName"; $colors.Add($HTMLColor['springgreen'])}
                "Average DB Requests (VB)*"         {Write-Host "$seriesName"; $colors.Add($HTMLColor['midnightblue'])}
                "Average DB Requests (Web)*"        {Write-Host "$seriesName"; $colors.Add($HTMLColor['Fuchsia'])}
                "Average DB Time*"                  {Write-Host "$seriesName"; $colors.Add($HTMLColor['Navy'])}
                "Average GRefs*"                    {Write-Host "$seriesName"; $colors.Add($HTMLColor['Lime'])}
                "Average Web Network Time*"         {Write-Host "$seriesName"; $colors.Add($HTMLColor['Silver'])}
                "Average Web Requests*"             {Write-Host "$seriesName"; $colors.Add($HTMLColor['Purple'])}
                "Average Web Time*"                 {Write-Host "$seriesName"; $colors.Add($HTMLColor['Maroon'])}
                "Average Workflow Time"            {Write-Host "$seriesName"; $colors.Add($HTMLColor['Green'])}
                "Average Workflow Time (Red)"      {Write-Host "$seriesName"; $colors.Add($HTMLColor['firebrick'])}
                "Average Workflow Time (Yellow)"   {Write-Host "$seriesName"; $colors.Add($HTMLColor['Gold'])}
                "Workflow Count"                   {Write-Host "$seriesName"; $colors.Add($HTMLColor['Navy'])}
                "Workflow Count (Red)"             {Write-Host "$seriesName"; $colors.Add($HTMLColor['Red'])}
                "Workflow Count (Yellow)"          {Write-Host "$seriesName"; $colors.Add($HTMLColor['Yellow'])}
            }
        }
        $chartTemplateObject.colors = $colors
        Write-Verbose "Number of Colors added: $($colors.count)"
        Write-Verbose "Number of series added: $($chartTemplateObject.series.name.count)"
        $charts.Add($chartTemplateObject)
    }


    Update-PleaseWaitWindow -message "Creating HTML graphs"
    Write-Verbose "Creating HTML Template"
$HTMLTemplate = @"
<html>
   <head>
        <title>$($EpicCitrixServer)</title>
        <style>
            html {
                overflow-y: scroll;
            }            
            body {
                background-color: #f3f3f3;
            }
            .content {
                margin-top: 70px;
                margin-right: 60px;
                margin-left: 60px;
            }
            .sticky {
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                z-index: 999;
                background-color: white;
                box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.2), 0 2px 10px 0 rgba(0, 0, 0, 0.19);
            }
            .sticky img {
                width: 10%;                
                margin: 10px;
            }
            .widget {
                color: #33485b;
                font-size: 14px;
                font-weight: 900;
                line-height: 24px;
                font-family: Helvetica, Arial, sans-serif; 
                opacity: 1;
            }
            h1 {
                color: #2e3134;
                padding-top: 2px;
                padding-bottom: 14px;
                font-size: 14px;
                line-height: 24px;
            }           

            .floatingWidget {
                background-color: white;
                box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.2), 0 6px 20px 0 rgba(0, 0, 0, 0.19);
                margin-bottom: 15px;
            }
        </style>
    </head>
    <script src="https://cdn.jsdelivr.net/npm/apexcharts"></script>    

<body>
<div class="sticky">
  <img src="data:img/png;base64,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UAnXrV9vzwWKNjxzacEKSDUtiKC13I1/sRoAB769Ul0dKKpKfsJZU7W1tbbswtqHE0FGSgN1KriuxOuok/HUKYScjvq999/3rij5ghVWxGrdh+usICFa0QQ2H6DCcR5REBwJ/elnn9mpCoZUzorIFYRNntIwly1bJosXL/a+GUp+53owMeOc53UpHjiJdc2aNXYKB5fr9kH4xIO6vWHjRjv9EEqUuD7/JZO/PLWZeaL42ebN9ohoRIB5o1yLUrJgPdCLr1y1yvb8DJWymR7CICzCJC95qIDYV86Z490RTLbnq+OBvAzLN67HWqtbYDpB7osE//A7VZiGCYobV6LHLDEmmzJkXnz+/Pn2hNZsGyGEwyiMI9JxNTULZNPHHwcele6nMMobTunMH2UwQ3IdMWLebN68eXZogyhRieghEQoqE45CzoVz1l88p65Or6iwx1B/uH69VFVV2WFWUHqCwknE4Qd+4SeVnjAYztEIEdQ3Vq60VnYspk2bZv3B+f0nvtNzuIoAi4i0PfblFZ95iBlLHMePH2/viUwT+UR6U2W68QO//H67PKT8082CV1+VvxuBZW5yxowZttOgPhIH8iRd7cNfpyh/HHnJ6bDUqeqaai9G0Zn50kwbt8g44V+q889KOPbhlff/UHgA0drWZudznChRUGGWTKYgJIb4c42olpeXD1xMAtLz4OFDewS1s2apbMng5mqJ1ygjNOONCJVPnOh9mzjEp/HXX6XLNAZOIAVGEfOr42tImaKhocHmF+VOmSOqS5YujToH6OA3DadO2bSRz/ye8nt14ULvjtRgvp2n527unA60xgzfR44a5d2RWUhfS0uLtJo61e61EepUUm3DpKHYpIGOeYSxRsdQp8rKUlqjSpzONDZagXWjBzrqucZ4Sif68OoZcQmroihKLFRYHSL/D3GTakYlltNOAAAAAElFTkSuQmCC"></img>
</div>
<div class=content>
"@

    $chartCount = 0
    foreach ($chart in $charts) {
    $chartCount = $chartCount+1
    $HTMLTemplate += "
    <div id=$("Chart$chartCount") class=floatingWidget></div>
    <script>
    var $("Chart$chartCount") = JSON.parse(`'$($chart | ConvertTo-Json -Depth 99 -Compress)`')
    var chart = new ApexCharts(document.querySelector(`"#$("Chart$chartCount")`"), $("Chart$chartCount"));
    chart.render();
    </script>
    "
    }
    $HTMLTemplate += "
    </body>
    </html>
    "

    $Random = Get-Random
    $HTMLTemplate | Set-Content -Path "$env:temp\$($EpicCitrixServer)_$Random.html"
    Close-PleaseWaitWindow

    Start-Process -FilePath "$env:temp\$($EpicCitrixServer)_$Random.html"
}
