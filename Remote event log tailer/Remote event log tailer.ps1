#requires -version 3.0

<#
    Monitor specified event logs on specified machines via event notifications


    Modification History:

    @guyrleech 16/07/2021  Initial version
    @guyrleech 19/07/2021  Zero or negative duration means run forever
    @guyrleech 20/07/2021  Added markers for start and end of tailing
    @guyrleech 20/07/2021  Added highlight options
    @guyrleech 21/07/2021  Added ability to pattern match with * on event log names. Added enabling of disabled logs
    @guyrleech 23/07/2021  Added parameter to ignore analytic/debug logs as cannot have watchers. Fixed script not exiting when grid view closed when no end time
    @guyrleech 26/11/2021  Added thread to monitor grid view and exit main loop if closed (via window title)
#>

[CmdletBinding()]

Param
(
	[string[]]$computers ,
	[decimal]$runForMinutes = 1 ,
	[ValidateSet('yes', 'no')]
	[string]$problemsOnly = 'no' ,
	[ValidateSet('yes', 'no')]
	[string]$enableDisabledLogs = 'no' ,
	[ValidateSet('yes', 'no')]
	[string]$ignoreAnalyticAndDebugLogs = 'yes' ,
	[string[]]$eventLogs = @( 
		'*-TerminalServices-*/*' ,
		'*-User Profile Service*/Operational' ,
		'*-GroupPolicy*/Operational' ,
		'*-RemoteDesktopServices*/Operational' ,
		'Application' ,
		'System' ) ,
	[string]$highlightProvider ,
	[string]$highlightMessage ,
	[string]$highlightId ,
	[string]$marker = '------' ,
	[string]$highlightBefore = 'V ' ,
	[string]$highlightAfter = '^ '
)

[datetime]$startTime = [datetime]::Now
[string]$eventquery = '*'
[string]$sourceIdentifierBase = 'GLEvent'
$sourceIdentifiers = New-Object -TypeName System.Collections.Generic.List[string] -ArgumentList @()
$watchers = New-Object -TypeName System.Collections.Generic.List[object] -ArgumentList @()
## key on computer name to store sessions so we can both enumerate at end to dispose but also so we can get session later to disable any logs we enabled
[hashtable]$sessions = @{}
$enabledLogs = New-Object -TypeName System.Collections.Generic.List[object] -ArgumentList @()
[string]$sourceIdentifier = $null

if ( $problemsOnly -and $problemsOnly[0] -ieq 'y') {
	$eventquery = '*[System[(Level=1 or Level=2 or Level=3)]]'
}

if ( $computers -and $computers.Count -eq 1 -and $computers[0].IndexOf( ',' ) -ge 0 ) {
	$computers = @( $computers -split ',' )
}

if ( $eventLogs -and $eventLogs.Count -eq 1 -and $eventLogs[0].IndexOf( ',' ) -ge 0 ) {
	$eventLogs = @( $eventLogs -split ',' )
}

try {
	ForEach ( $computer in $computers ) {
		[int]$counter = 0
		if ( ! ( $session = New-Object -TypeName System.Diagnostics.Eventing.Reader.EventLogSession -ArgumentList $computer ) ) {
			Write-Warning -Message "Failed to start event log session on $computer"
			continue
		}
		try {
			$sessions.Add( $computer , $session )
		}
		catch {
			Write-Warning -Message "Duplicate computer $computer"
			$computer = $null
		}

		if ( $null -ne $computer ) {
			[string[]]$allEventLogs = @()
			try {
				$allEventLogs = @( $session.GetLogNames() )
			}
			catch {
				$allEventLogs = $null
				Write-Warning -Message "Failed to get event log names from $computer : $_"
			}

			ForEach ( $eventLogPattern in $eventLogs ) {
				## verify each event log exists even if no wildcard so if doesn't exist, we don't set a watcher
				[int]$matchingLogs = 0
				$allEventLogs.Where( { $_ -like $eventLogPattern -and ( [string]::IsNullOrEmpty( $ignoreAnalyticAndDebugLogs ) -or $ignoreAnalyticAndDebugLogs[0] -ieq 'n' -or $_ -notmatch '/(analytic|debug)$' ) } ) | ForEach-Object `
				{
					$eventLog = $_
					$matchingLogs++
					if ( $query = New-Object Diagnostics.Eventing.Reader.EventLogQuery ($eventLog, [Diagnostics.Eventing.Reader.PathType]::LogName, $eventquery ) ) {
						$query.Session = $session
						$query.TolerateQueryErrors = $true

						if ( $watcher = New-Object Diagnostics.Eventing.Reader.EventLogWatcher -ArgumentList $query ) {
							$sourceIdentifier = $sourceIdentifierBase + ".$computer.$counter"
							$counter++
							Write-Verbose -Message "source identifier $sourceIdentifier for $eventlog"
							Unregister-Event -SourceIdentifier $sourceIdentifier -Force -ErrorAction SilentlyContinue
							Register-ObjectEvent -InputObject $watcher -EventName EventRecordWritten -SourceIdentifier $sourceIdentifier

							if ( $eventLogConfiguration = New-Object Diagnostics.Eventing.Reader.EventLogConfiguration -ArgumentList $eventLog , $session ) {
								[bool]$continue = $true

								if ( ! [string]::IsNullOrEmpty( $ignoreAnalyticAndDebugLogs ) -and $ignoreAnalyticAndDebugLogs[0] -ieq 'y' `
										-and ( $eventLogConfiguration.LogType -eq [System.Diagnostics.Eventing.Reader.EventLogType]::Analytical -or $eventLogConfiguration.LogType -eq [System.Diagnostics.Eventing.Reader.EventLogType]::Debug ) ) {
									$continue = $false
									Write-Verbose -Message "Ignoring log `"$eventlog`" on $computer because type is $($eventLogConfiguration.LogType)"
									$watcher.Dispose()
									$watcher = $null
								}
                                
								if ( $continue -and ! [string]::IsNullOrEmpty( $enableDisabledLogs ) -and $enableDisabledLogs[0] -ieq 'y' -and ! $eventLogConfiguration.IsEnabled ) {
									try {
										$eventLogConfiguration.IsEnabled = $true
										$eventLogConfiguration.SaveChanges()
										Write-Warning -Message "Log `"$eventLog`" on $computer was disabled so enabled it"
										$enabledLogs.Add( "$($computer):$eventLog" )
										$watcher.Enabled = $true
										$continue = $true
									}
									catch {
										Write-Warning -Message "Problem enabling log `"$eventLog`" on $computer : $_"
										$watcher.Dispose()
										$watcher = $null
									}
								}

								$eventLogConfiguration.Dispose()
								$eventLogConfiguration = $null
							}
							else {
								Write-Warning -Message "Failed to get configuration for log `"$eventLog`" on $computer so cannot tell if disabled or analytic/debug"
							}

							if ( $watcher ) {
								## Enabling on some logs throws "The request is not supported" which is probably because they are analytic/debug which don't appear to alow watchers
								try {
									$watcher.Enabled = $true
								}
								catch {
									$exception = $_
									Write-Warning -Message "Problem setting watcher on log `"$eventLog`" on $computer : $exception"
									Unregister-Event -SourceIdentifier $sourceIdentifier -Force -ErrorAction SilentlyContinue
									$watcher.Dispose()
									$watcher = $null
								}
								if ( $watcher ) {
									$watchers.Add( $watcher )
									$sourceIdentifiers.Add( $sourceIdentifier )
								}
							}
							else {
								Unregister-Event -SourceIdentifier $sourceIdentifier -Force -ErrorAction SilentlyContinue
							}
						}
						else {
							Write-Warning -Message "Failed to create watcher for event log $eventlog on $computer with query $eventquery"
						}
					}
					else {
						Write-Warning -Message "Failed to create query for event log $eventlog on $computer"
					}
				}
				if ( ! $matchingLogs ) {
					Write-Warning -Message "No event logs found on $computer matching `"$eventLogPattern`""
				}
			}
		}
		if ( ! $counter -and $computer ) {
			Write-Warning -Message "No event logs found to monitor on $computer"
		}
	}
    
	# Test if any logs are being monitored at all
	if( $watchers.Count -eq 0 )
    {
        Throw "Nothing to monitor. This could be because the machines specified cannot be reached or the logs requested do not exist."
    }

	[string]$message = "Monitoring $($watchers.Count) event logs across $($sessions.Count) machines from $(Get-Date -Format G)"

	$endTime = $null
	if ( $runForMinutes -gt 0 ) {
		$endTime = (Get-Date).AddSeconds( $runForMinutes * 60 )
		$message += " until $(Get-Date -Date $endTime -Format G)"
	}

	## start a thread to monitor gridview so we can quit if closed
	$monitorThread = $null
	$sharedVariables = $null
	[string]$eventToSignal = 'GuysWindowTitleChangedEventThingy'

	if ( ! ( $SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault() ) ) {
		Write-Warning "Failed to create runspace initial session state to monitor grid view"
	}
	else {
		## when the thread is ready to exit, it modified this collection which generates an event - courtesy of Bo Prox via https://social.technet.microsoft.com/Forums/SECURITY/en-US/4ce2efd2-1852-44a5-85a7-29659e5bf704/raise-a-powershell-events-from-within-a-runspace-which-has-no-gui-elements-to-the-powershell-host?forum=winserverpowershell
		$observableCollection = New-Object -TypeName System.Collections.ObjectModel.ObservableCollection[string]
		Register-ObjectEvent -InputObject $observableCollection -EventName CollectionChanged -SourceIdentifier $eventToSignal
 
		## sharing variables with the runspaces so they can signal the main thread
		$sharedVariables = [hashtable]::Synchronized(@{ Exit = $false ; SignalCollection = $observableCollection })
		## must add shared variables before creating runspace pool
		$sessionState.Variables.Add( (New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList 'sharedVariables' , $sharedVariables , 'Shared Variables hashtable' ) )
        
		if ( $RunspacePool = [runspacefactory]::CreateRunspacePool(
				1, ## Min Runspaces
				2 , ## Max parallel runspaces ,
				$sessionstate ,
				$host.psobject.Copy() )) {
			$RunspacePool.Open()
            ($powerShell = [PowerShell]::Create()).RunspacePool = $RunspacePool

			[void]$powerShell.AddScript({
					Param( [string]$message , [int]$pollPeriodMilliseconds = 10000 , $eventToSignal = 'DoneWatching' )

                
					## https://www.linkedin.com/pulse/fun-powershell-finding-suspicious-cmd-processes-britton-manahan/

					Add-Type -TypeDefinition  @"
            using System;
            using System.Text;
            using System.Collections.Generic;
            using System.Runtime.InteropServices;

            namespace Api
            {

             public class WinStruct
             {
               public string WinTitle {get; set; }
               public int WinHwnd { get; set; }
               public int PID { get; set; }
             }

             public class ApiDef
             {
               private delegate bool CallBackPtr(int hwnd, int lParam);
               private static CallBackPtr callBackPtr = Callback;
               private static List<WinStruct> _WinStructList = new List<WinStruct>();

               [DllImport("User32.dll")]
               [return: MarshalAs(UnmanagedType.Bool)]
               private static extern bool EnumWindows(CallBackPtr lpEnumFunc, IntPtr lParam);

               [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
               static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
   
               [DllImport("user32.dll")]
               static extern bool IsWindowVisible(IntPtr hWnd);
   
                [DllImport("user32.dll")]
                public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
 
               private static bool Callback(int hWnd, int pid)
               {
                    if( IsWindowVisible( (IntPtr)hWnd ) )
                    {
                        int ipid = 0 ;
                        GetWindowThreadProcessId( (IntPtr)hWnd , out ipid );
                        if( ipid == pid )
                        {
                            StringBuilder sb = new StringBuilder(256);
                            int res = GetWindowText((IntPtr)hWnd, sb, 256);
                            _WinStructList.Add( new WinStruct { WinHwnd = hWnd, WinTitle = sb.ToString() , PID = ipid  });
                        }
                    }
                    return true;
               }   

               public static List<WinStruct> GetWindows( int pid )
               {
                  _WinStructList = new List<WinStruct>();
                  EnumWindows(callBackPtr, (IntPtr)pid );
                  return _WinStructList;
               }

             }
            }
"@
					[bool]$haveSeenTitle = $false
					[int]$waitMilliseconds = 500 ## initially poll frequently to get the window title then we can back off
                
					[System.Diagnostics.Trace]::WriteLine( "Pid is $pid poll $pollPeriodMilliseconds message $message" )
					Do {
						if ( $thisProcess = Get-Process -Id $pid ) {
							if ( [string]::IsNullOrEmpty( $thisProcess.MainWindowTitle) -or $thisProcess.MainWindowTitle -ne $message ) {
								if ( $haveSeenTitle ) {
									## enumerate all visible windows for this process to see if the grid view is still present but not the foreground window currently
									if ( -Not ( $windows = [Api.Apidef]::GetWindows( $pid ) ) -or -Not $windows.Where( { $_.WinTitle -eq $message } )) {
										[System.Diagnostics.Trace]::WriteLine( "Setting exit as title now `"$($thisProcess.MainWindowTitle)`" and window title not found" )
										[void]$sharedVariables.SignalCollection.Add( "Window title: $($thisProcess.MainWindowTitle)" ) ## main loop is waiting on an event which this will signal
										$sharedVariables.Exit = $true
										break
									}
									else {
										[System.Diagnostics.Trace]::WriteLine( "Process still has a window with title `"$message`" althouhg is now `"$($thisProcess.MainWindowTitle)`"" )
									}
								}
								## else not seen title yet so don't exit
							}
							elseif ( $thisProcess.MainWindowTitle -eq $message ) {
								$haveSeenTitle = $true ## must see title at least once since we could be checking before grid view is viewable
								[System.Diagnostics.Trace]::WriteLine( "Seen title $message" )
								$waitMilliseconds = $pollPeriodMilliseconds
							}
						}
						else {
							Write-Warning -Message "Failed to get pid $pid"
						}
						[System.Diagnostics.Trace]::WriteLine( "Sleeping for $waitMilliseconds ms" )

						Start-Sleep -Milliseconds $waitMilliseconds
					}
					While ( $true )
					## implicit exit from thread
				})
        
			[void]$powerShell.AddParameters( @{ message = $message ; eventToSignal = $eventToSignal } )
			$monitorThread = [pscustomobject]@{
				'PowerShell' = $powerShell
				'Handle'     = $powerShell.BeginInvoke( ) 
			}
		}
	}

	Write-Verbose -Message $message

	## put loop in a scriptblock so it can stream its output before it has finished https://powershell.one/tricks/loops/make-loops-stream
	$selected = & {
		[pscustomobject][ordered]@{  Time = Get-Date -Format 'HH:mm:ss.fff' ; MachineName = $marker ; Id = 0 ; LevelDisplayName = $marker ; OpcodeDisplayName = $marker ; TaskDisplayName = $marker; LogName = $marker; ProviderName = $marker; Message = 'STARTED TAILING EVENT LOGS  ' }
        
		do {
			if ( $eventRaised = Wait-Event -Timeout $( if ( $endTime ) { ($endTime - [datetime]::Now).TotalSeconds } else { -1 } ) ) { ## when no endtime, we must periodically come out of Wait-Event in case grid view has been closed
				if ( $eventRaised.TimeGenerated -ge $startTime ) {
					if ( $sourceIdentifiers -contains $eventRaised.SourceIdentifier -and $eventRaised.SourceEventArgs -and ( $event = $eventRaised.SourceEventArgs | Select-Object -ExpandProperty EventRecord )) { ## could make hash table for speedier lookup
						[string]$messageText = $event.FormatDescription()
						if ( [string]::IsNullOrEmpty( $messageText ) ) {
							$messageText = $event.Properties | Select-Object -ExpandProperty Value
						}
						## cannot use $PSBoundParameters on script parameters as we are in a scriptblock which has its own
						[bool]$highlight = ! [string]::IsNullOrEmpty( $highlightMessage ) -and $messageText -match $highlightMessage
						if ( ! $highlight ) {
							if ( ! ( $highlight = ! [string]::IsNullOrEmpty( $highlightId ) -and $event.Id -match $highlightId ) ) {
								$highlight = ! [string]::IsNullOrEmpty( $highlightProvider ) -and $event.ProviderName -match $highlightProvider 
							}
						}
						if ( $highlight ) {
							[string]$markertext = $highlightBefore * 3
							[pscustomobject][ordered]@{  Time = Get-Date -Date $event.TimeCreated -Format 'HH:mm:ss.fff' ; MachineName = $markertext ; Id = 0 ; LevelDisplayName = $markertext ; OpcodeDisplayName = $markertext ; TaskDisplayName = $markertext; LogName = $markertext; ProviderName = $markertext; Message = $markertext }
						}
						$event | Select-Object -Property @{n = 'Time'; e = { Get-Date -Date $_.TimeCreated -Format 'HH:mm:ss.fff' } }, MachineName, Id, LevelDisplayName, OpcodeDisplayName, TaskDisplayName, LogName, ProviderName, @{n = 'Message'; e = { $messageText } }
						if ( $highlight ) {
							[string]$markertext = $highlightAfter * 3
							[pscustomobject][ordered]@{  Time = Get-Date -Date $event.TimeCreated -Format 'HH:mm:ss.fff' ; MachineName = $markertext ; Id = 0 ; LevelDisplayName = $markertext ; OpcodeDisplayName = $markertext ; TaskDisplayName = $markertext; LogName = $markertext; ProviderName = $markertext; Message = $markertext }
						}
					}
					elseif ( $eventRaised.SourceIdentifier -eq $eventToSignal ) { ## event from our thread watching the main window
						Write-Verbose -Message "$(Get-Date -Format G) : got event from thread : $($eventRaised.SourceEventArgs | Select-Object -ExpandProperty NewItems)" 
						break
					}
				}
				## else event raised before we started so ignore
        
				$eventRaised | Remove-Event
				$eventRaised = $null
			}
		} while ( ( ! $endTime -or [datetime]::Now -le $endTime ) -and -not $sharedVariables.Exit )

		[pscustomobject][ordered]@{  Time = Get-Date -Format 'HH:mm:ss.fff' ; MachineName = $marker ; Id = 0 ; LevelDisplayName = $marker ; OpcodeDisplayName = $marker ; TaskDisplayName = $marker; LogName = $marker; ProviderName = $marker; Message = 'FINISHED TAILING EVENT LOGS' }

	} | Out-GridView -Title $message -PassThru

	if ( $null -ne $selected ) {
		$selected | Set-Clipboard
	}
}
catch {
	throw $_
}
finally {
	Write-Verbose -Message "$(Get-Date -Format G) cleaning up"

	if ( $monitorThread ) {
		[void]$monitorThread.PowerShell.Stop()
		[void]$monitorThread.PowerShell.Dispose()
	}

	Unregister-Event -SourceIdentifier $eventToSignal

	ForEach ( $sourceIdentifier in $sourceIdentifiers ) {
		Unregister-Event -SourceIdentifier $sourceIdentifier -Force -ErrorAction SilentlyContinue
	}

	ForEach ( $watcher in $watchers ) {
		$watcher.Enabled = $false
		$watcher.Dispose()
		$watcher = $null
	}

	ForEach ( $enabledLog in $enabledLogs ) {
		## TODO disable the log
		$computer, $eventlog = $enabledLog -split ':' , 2
		if ( $session = $sessions[ $computer ] ) {
			if ( $eventLogConfiguration = New-Object Diagnostics.Eventing.Reader.EventLogConfiguration -ArgumentList $eventLog , $session ) {
				if ( $eventLogConfiguration.IsEnabled ) {
					try {
						$eventLogConfiguration.IsEnabled = $false
						$eventLogConfiguration.SaveChanges()
						Write-Verbose -Message "Log `"$eventLog`" on $computer disabled ok"
					}
					catch {
						Write-Warning -Message "Problem disabling log `"$eventLog`" on $computer : $_"
					}
				}
				else {
					Write-Warning -Message "Log `"$eventLog`" on $computer is not enabled but we enabled it"
				}
				$eventLogConfiguration.Dispose()
				$eventLogConfiguration = $null
			}
			else {
				Write-Warning -Message "Failed to get configuration for log `"$eventLog`" on $computer"
			}
		}
		else {
			Write-Warning -Message "Failed to get session for $computer to disable log `"$eventLog`""
		}
	}

	ForEach ( $session in $sessions.GetEnumerator() ) {
		$session.Value.Dispose()
		$session.Value = $null
	}

	$watchers.Clear()
	$sessions.Clear()
}

