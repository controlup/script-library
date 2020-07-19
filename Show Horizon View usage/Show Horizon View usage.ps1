#requires -version 3

<#
.SYNOPSIS

Draw a graph with the event data written to the VMware Horizon View event database once a day which reports the maximum usage for that day

.DETAILS

Uses VMware PowerCLI to connect to a Horizon View Connection Server and retrieve the relevant events which then drawn using Windows Forms in a separate thread

.PARAMETER server

The hostname of the Horizon View Server to connect to

.PARAMETER domain

The domain name of the Horizon View Server to connect to

.PARAMETER username

The username of the user with sufficient rights in Horizon View to query the events database

.PARAMETER password

The password for the user  with sufficient rights in Horizon View to query the events database

.PARAMETER daysBack

The number of days in the past to fetch events from

.PARAMETER timeout

The timeout in seconds to wait for the thread with the GUI to be ready

.CONTEXT

Computer

.NOTES

Needs to run on the ControlUp Console machine as it shows a user interface. This machine must also have VMware PowerCLI installed

With code from https://github.com/andyjmorgan/HorizonPowershellScripts/blob/master/unique%20logons%20per%20day%20in%20the%20last%20week.ps1

.MODIFICATION_HISTORY:

@guyrleech 28/07/19 Initial release
@guyrleech 11/11/19 Tweaks durng blogpost writing

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage="Horizon View Connection Server NetBIOS Name")]
    [string]$server ,
    [Parameter(Mandatory=$true,HelpMessage="Horizon View Connection Server domain Name")]
    [string]$domain ,
    [Parameter(Mandatory=$true,HelpMessage="Horizon View Connection Server admin username")]
    [string]$username ,
    [Parameter(Mandatory=$true,HelpMessage="Horizon View Connection Server admin password")]
    [string]$password ,
    [Parameter(Mandatory=$true,HelpMessage="Number of previous days to report on")]
    [int]$daysBack ,
    [Parameter(Mandatory=$false,HelpMessage="Number of seconds to wait for form to appear")]
    [int]$timeout = 30
)


$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { 'Continue' } else { 'SilentlyContinue' })
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { 'Continue' } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'ErrorAction' ] ) { $ErrorActionPreference } else { 'Stop' })

[int]$exitCode = 0

Function Get-EventSummaryView{
    Param(
        [parameter(mandatory=$true)]
        $hvServer,
        [parameter(mandatory=$false)]
        $days = 7
    )
    $query_service_helper = New-Object VMware.Hv.QueryServiceService
    $query = New-Object VMware.Hv.QueryDefinition
    $query.queryEntityType = 'EventSummaryView'
    $services = Get-ViewAPIService -hvServer $hvServer
    
    $OrFilter = New-Object VMware.Hv.QueryFilterOr
    ForEach( $eventType in @( 'BROKER_DAILY_MAX_APP_USERS' , 'BROKER_DAILY_MAX_CCU_USERS' ,  'BROKER_DAILY_MAX_DESKTOP_SESSIONS' ) ) ## 'BROKER_DAILY_MAX_NU_USERS' seems to be a static count
    {
        $filter = New-Object VMware.Hv.QueryFilterEquals
        $filter.memberName = 'data.eventType'
        $filter.value = $eventType
        $Orfilter.Filters += $filter
    }

    $Andfilter = New-Object VMware.Hv.QueryFilterAnd
    
    $date= Get-Date -Hour 0 -Minute 00 -Second 00  ## midnight

    $DateFilter = New-Object VMware.Hv.QueryFilterBetween
    $DateFilter.FromValue = $date.AddDays( -$days)
    $DateFilter.ToValue = $date
    $datefilter.MemberName= 'data.time'
    $Andfilter.Filters += $DateFilter
    $andfilter.Filters += $OrFilter
    $query.Filter = $AndFilter
    
    $queryResponse = $query_service_helper.QueryService_Create( $services , $query )
    
    $results = [System.Collections.Generic.List[object]]$queryResponse.Results

    if($queryResponse.RemainingCount -gt 0){
        Write-Verbose  -Message "Found further results to retrieve"
        $remaining = $queryResponse.RemainingCount
        do{
            
            $latestResponse = $query_service_helper.QueryService_GetNext($services,$queryResponse.Id)
            $results += $latestResponse.Results
            Write-Verbose  -Message "Pulled an additional $($latestResponse.Results.Count) item(s)"
            $remaining = $latestResponse.RemainingCount
        } while( $remaining -gt 0 )
    }

    $query_service_helper.QueryService_Delete($services,$queryResponse.Id)

    $results
}

Function Get-ViewAPIService {
  Param(
    [Parameter(Mandatory = $false)]
    $HvServer
  )
  if ($null -ne $hvServer) {
    if ($hvServer.GetType().name -ne 'ViewServerImpl') {
      $type = $hvServer.GetType().name
      Write-Error -Message "Expected hvServer type is ViewServerImpl, but received: [$type]"
      return $null
    }
    elseif ($hvServer.IsConnected) {
      return $hvServer.ExtensionData
    }
  } elseif ($global:DefaultHVServers.Length -gt 0) {
     $hvServer = $global:DefaultHVServers[0]
     return $hvServer.ExtensionData
  }
  return $null
}

Write-Verbose -Message "$(Get-Date -Format G) : Importing PowerCLI module ..."
$oldVerbosePreference = $VerbosePreference
$VerbosePreference = 'SilentlyContinue'
##Import-Module -Name VMware.VimAutomation.Core -Verbose:$false -Debug:$false
$module = Import-Module -Name VMware.VimAutomation.HorizonView -Verbose:$false -Debug:$false -PassThru
if( ! $module )
{
    Throw "Unable to load Horizon View PowerShell module"
}
$VerbosePreference = $oldVerbosePreference
[hashtable]$powerCLISettings = @{
    'ParticipateInCeip' = $false 
    'InvalidCertificateAction' = 'Ignore'
    'DisplayDeprecationWarnings' = $false
    'Confirm' = $false
    'Scope' = 'Session'
}

[void](Set-PowerCLIConfiguration @powerCLISettings )

$oldWarningPreference = $WarningPreference
$WarningPreference = 'SilentlyContinue' ## may warn about certs or CEIP
Write-Verbose -Message "$(Get-Date -Format G) : Connecting to $server as $username ..."
$hvServer = Connect-HVServer -Server $server -User $username -Password $password -Domain $domain -Force -WarningAction SilentlyContinue
$password = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
$WarningPreference = $oldWarningPreference

if( ! $hvServer )
{
    Throw "Unable to connect to Horizon View Connection Server $server as $domain\$username"
}

Write-Verbose -Message "$(Get-Date -Format G) : Connected to server $server, fetching events from last $daysBack days ..."

$logevents = Get-EventSummaryView $hvserver -days $daysBack

Disconnect-HVServer -Server $hvServer -Confirm:$false
$hvServer = $null

if( ! $logevents -or ! $logevents.Count )
{
    Write-Warning -Message "Failed to find any relevant events in the last $daysBack days"
}
else
{ 
    [hashtable]$eventsByType = $logEvents | Select-Object -ExpandProperty Data | Group-Object -Property EventType -AsHashTable

    Write-Verbose -Message "$(Get-Date -Format G) : Got $($logEvents.Count) events grouped into $($eventsByType.Count) groups by type"
    ## code from TTYE
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"         
    $newRunspace.Open()
    $syncHash = [hashtable]::Synchronized(@{})
    $newRunspace.SessionStateProxy.SetVariable( 'syncHash' , $syncHash ) 
    $syncHash.eventsByType = $eventsByType
    $syncHash.daysBack = $daysBack

    $psCmd = [PowerShell]::Create().AddScript({
    try
    {
        $VerbosePreference = 'SilentlyContinue'
        ##Start-Transcript -Path (Join-Path -Path $env:TEMP -ChildPath 'view-sba.log' )
        Add-Type -AssemblyName System.Windows.Forms.DataVisualization

        $chartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    
        $Chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
        $chart.width = 1000
        $chart.Height = 700
        [void]$chart.Titles.Add( "VMware Horizon View Connections in last $($syncHash.daysBack) days" )
        $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
        $Chart.ChartAreas.Add($ChartArea)
        $ChartArea.AxisY.Title = 'Number'
        ## Turn off grid lines
        $ChartArea.AxisY.MajorGrid.Enabled = $ChartArea.AxisY.MinorGrid.Enabled = $ChartArea.AxisX.MajorGrid.Enabled = $ChartArea.AxisX.MinorGrid.Enabled = $false

        ForEach( $eventType in $syncHash.eventsByType.GetEnumerator() )
        {
            [string]$legendName = Switch( $eventType.Key )
            {
                'BROKER_DAILY_MAX_APP_USERS' { 'Max Application Users' }
                'BROKER_DAILY_MAX_CCU_USERS' { 'Max Concurrent Users' }
                'BROKER_DAILY_MAX_DESKTOP_SESSIONS' { 'Max Desktop Users' } 
            }
            $chartSeries = $Chart.Series.Add($legendName)
            $ChartSeries.ChartType = $chartType
            $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
            $legend.Name = $legendName
            $chartSeries.ToolTip = $legendName
            $chartSeries.IsValueShownAsLabel = $false
            $Chart.Legends.Add($legend)

            $eventType.Value | Sort-Object -Property Time | . { Process `
            {
                $event = $_
                if( $event.Message -match '\d+$' )
                {
                    $point = $chartSeries.Points.AddXY( (Get-Date -Date $event.Time -Format d) , $Matches[0] )
                }
            }}
        }

        $AnchorAll = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor
            [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
        $syncHash.Form = New-Object Windows.Forms.Form
        $syncHash.Form.Width = $chart.Width
        $syncHash.Form.Height = $chart.Height + 50
        $syncHash.Form.AutoSize = $true
        $syncHash.Form.controls.add($Chart)
        $Chart.Anchor = $AnchorAll

        ## add a variable that we check for to indicate that dialogue has been activated
        [void]$syncHash.Form.Add_Shown({ $syncHash.Form.Activate() ; $syncHash.ReadyToGo = $true })
        $syncHash.Form.Visible = $false
        $syncHash.Form.TopMost = $true

        Write-Verbose -Message 'Showing chart'
        
        [void]$syncHash.Form.ShowDialog()
       }
       catch
       {
            ## $_ | Out-File -FilePath (Join-Path -Path $env:temp -ChildPath 'cu-view-sba.log')
            $syncHash.Form = $null
            Throw $_
        }
        finally
        {
            ##Stop-Transcript
        }
    })
  
    $psCmd.Runspace = $newRunspace
    $data = $psCmd.BeginInvoke()
  
    $signature = @'
        public enum WindowShowStyle : uint
        {
            Hide = 0,
            ShowNormal = 1,
            ShowMinimized = 2,
            ShowMaximized = 3,
            Maximize = 3,
            ShowNormalNoActivate = 4,
            Show = 5,
            Minimize = 6,
            ShowMinNoActivate = 7,
            ShowNoActivate = 8,
            Restore = 9,
            ShowDefault = 10,
            ForceMinimized = 11
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, WindowShowStyle nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
'@

    ## Wait for form to be visible
    $timer = [Diagnostics.Stopwatch]::StartNew()
    [bool]$timedOut = $false
  
    ## Start-Sleep doesn't return anything so we are just sleeping part way through the while statement :-)
    do
    {
        try
        {
            $notDone = ( ! $data.IsCompleted -and ! $timedOut  -and ! $syncHash.Contains( 'ReadyToGo' ) -and ! $syncHash.Contains( 'Form' ) -and ! (Start-Sleep -Milliseconds 333) -and ! $syncHash.Form.PSObject.Properties[ 'Handle' ]  )
        }
        catch
        {
            $notDone = $True
            Write-Warning -Message $_
        }
        if( ! $notDone -and $timer.Elapsed.TotalSeconds -gt $timeout )
        {
            Write-Error -Message "Timeout waiting for form to appear"
            $timedOut = $true
            $notDone = $true
            $exitCode = 2
        }
    } while( $notDone )
    
    $timer.Stop()

    if( $syncHash.Contains( 'Form' ) -and $syncHash.Form )
    {
        [void](Add-Type -MemberDefinition $signature -Name 'Windows' -Namespace Win32Functions -PassThru -Debug:$false)

        $timer.Start()
        do
        {
            [bool]$failed = $false
            try
            {
                $syncHash.Form.Visible = $true
                [void][Win32Functions.Windows]::ShowWindow( $syncHash.Form.Handle , [Win32Functions.Windows+WindowShowStyle]::Show )
                [void][Win32Functions.Windows]::SetForegroundWindow( $syncHash.Form.Handle )
                Start-Sleep -Milliseconds 500
                [void][Win32Functions.Windows]::ShowWindow( $syncHash.Form.Handle , [Win32Functions.Windows+WindowShowStyle]::Show )
                [void][Win32Functions.Windows]::SetForegroundWindow( $syncHash.Form.Handle )
            }
            catch
            {
                if( $timer.Elapsed.TotalSeconds -gt $timeout )
                {
                    Write-Error -Message "Timeout waiting for form handle to appear"
                    $failed = $false
                    $exitCode = 3
                }
                else
                {
                    $failed = $true
                    Start-Sleep -Milliseconds 333
                }
            }
        } while( $failed )
        $timer.Stop()

        Write-Output -InputObject "A graph of the activity over the last $daysBack days should now have been shown in a separate window"

        ## Wait for window to close and then exit
        While( $syncHash.Form.Visible )
        {
            Start-Sleep -Milliseconds 500
        }
    }
    elseif( ! $data.IsCompleted )
    {
        ## Terminate thread
        $pscmd.Stop()
        $pscmd.Dispose()
    }
    else
    {
        $result = $psCmd.EndInvoke( $data )
        if( $result )
        {
            Write-Error -Message "Failed to create Windows form: $result"
            $exitCode = 1
        }
    }
}

Exit $exitCode

