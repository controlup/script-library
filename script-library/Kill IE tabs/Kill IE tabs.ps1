#requires -version 3

<#
    Find process for IE tabs

    Uses window enumeration code found at https://powertoe.wordpress.com/2010/11/10/finding-the-thread-pid-that-belongs-to-a-tab-in-ie-8-with-powershell/

    @guyrleech 2019
#>

[int]$sessionId = if( $args.Count -and $args[0] )
{
    $args[0] -as [int]
}
else
{
    Throw 'Must pass the session id of the session on the command line'
}

[string]$URLtoKill = if( $args.Count -ge 2 -and $args[1] )
{
    $args[1]
}
else
{
    Throw 'Must pass the URL to kill on the command line'
}

[bool]$forceKill = ( $args.Count -ge 3 -and ( $args[ 2 ] -eq 'Yes' -or $args[2] -match 'True' ))
[bool]$killRegardless = ( $args.Count -ge 4 -and ( $args[ 3 ] -eq 'Yes' -or $args[3] -match 'True' ))

[int]$outputWidth = 400

$VerbosePreference = 'SilentlyContinue'

Add-Type -Namespace User32 -Name Util -UsingNamespace System.Text -MemberDefinition @'
[DllImport("user32.dll", SetLastError=true)]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

[DllImport("user32.dll")]
public static extern IntPtr GetTopWindow(IntPtr hWnd);

[DllImport("user32.dll", SetLastError = true)]
public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

public enum GetWindow_Cmd : uint {
    GW_HWNDFIRST = 0,
    GW_HWNDLAST = 1,
    GW_HWNDNEXT = 2,
    GW_HWNDPREV = 3,
    GW_OWNER = 4,
    GW_CHILD = 5,
    GW_ENABLEDPOPUP = 6
}

[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

[DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
public static extern int GetWindowTextLength(IntPtr hWnd);
'@

Function Get-HKCU
{
    Param
    ( 
        [Parameter(Mandatory=$true)]
        $process 
    )

    $owner = Invoke-CimMethod -InputObject $process -MethodName GetOwner

    if( $owner )
    {
        [string]$sid = (New-Object -TypeName System.Security.Principal.NTAccount( "$($owner.Domain)\$($owner.User)" )).Translate([System.Security.Principal.SecurityIdentifier]).Value

        if( [string]::IsNullOrEmpty( $sid ) )
        {
            Write-Warning "Failed to get sid for user $($owner.Domain)\$($owner.User))"
        }
        else
        {
            if( ! (  Get-PSdrive -Name HKU -ErrorAction SilentlyContinue ) )
            {
                [void](New-PSDrive -Name HKU -PSProvider Registry -Root 'Registry::HKEY_USERS' -Scope Script )
            }

            Join-Path -Path 'HKU:\' -ChildPath $sid ## return
        }
    }
    else
    {
        Write-Warning "Failed to get owner for process id $($process.ProcessId) ($($process.Name))"
    }
}

[hashtable]$ieProcesses = @{}

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

## command line not returned on Server 2016 when run as non-admin - we will check later that it's a child of another IE process otherwise likely to kill multiple tabs
Get-CimInstance -Class Win32_Process -Filter "Name = 'iexplore.exe' and SessionId = '$sessionId'" | ForEach-Object `
{
    $ieProcesses.Add( $_.ProcessId -as [int] , $_ )
}

[hashtable]$toKill = @{}
[int]$tabsFound = 0
[hashtable]$notToKill = @{}
[hashtable]$topLevelTabs = @{}

if( ! $ieProcesses -or ! $ieProcesses.Count )
{
    Write-Warning "No child iexplore.exe processes found in session $sessionid"
}
else
{
    $window = [User32.Util]::GetTopWindow( [IntPtr]::Zero )

    if( ! $window )
    {
        Write-Error "Failed to get top most window in session $sessionId so cannot get tabs"
    }
    while ($window -ne [IntPtr]::Zero ) 
    {
        [int]$windowPid = 0
        [void]([User32.util]::GetWindowThreadProcessId( $window , [ref]$windowPid ))
        $ieProcess = $ieProcesses[ $windowPid ]
        if( $ieProcess ) 
        {
            $length = [User32.Util]::GetWindowTextLength( $window )
            if ($length -gt 0)
            {
                $string = New-Object System.Text.Stringbuilder 1024
                if( [User32.Util]::GetWindowText( $window , $string , ( $length + 1 )) -gt 0 )
                {
                    ## can be hidden blank pages but ignore anyway since won't contain content
                    if( $string.tostring() -notmatch '^(MSCTFIME UI|Default IME|SysFader|MCI command handling window|Tooltip|DDE Server Window|PseudoServerHiddenWindow|Blank Page - Internet Explorer)$' )
                    {
                        ## only look to kill if it is a child process not the parent as that will kill everything but still count the tab for reporting/checking purposes
                        if( (Get-Process -Id $ieProcess.ParentProcessId -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name) -eq 'iexplore')
                        {
                            $tabsFound++
                            if( $string.ToString() -match $URLtoKill )
                            {
                                Write-Verbose "Will kill pid $windowPid with text `"$($string.ToString())`""
                                ## Can't kill processes yet as it can mess up fetching of windows
                                try
                                {
                                    $toKill.Add( $windowPid , [pscustomobject]@{ Title = $string.tostring() ; 'Process' = $ieProcess } )
                                }
                                catch
                                {
                                    ## already have it so don't try and kill twice
                                }
                            }
                            else
                            {
                                Write-Verbose "Not killing pid $windowPid with text `"$($string.ToString())`" as no match"
                                ## keep a list of innocent pids so we can warn if we are going to kill them
                                $existing = $notToKill[ $windowPid ]
                                if( $existing )
                                {
                                    [void]$existing.Add( $string.ToString() )
                                }
                                else
                                {
                                    $notToKill.Add( $windowPid , ( [System.Collections.ArrayList]@( $string.ToString() ) ) )
                                }
                            }
                        }
                        else
                        {
                            Write-Verbose "May not kill pid $windowPid with text `"$($string.ToString())`" as top level process"
                            try
                            {
                                $topLevelTabs.Add( $string.tostring() , [pscustomobject]@{ Title = $string.tostring() ; 'Process' = $ieProcess } )
                            }
                            catch
                            {
                                ## already got it - may be same tab or another tab of same name but we can't tell
                            }
                        }
                    }
                }
            }
        }
        $window = [User32.Util]::GetWindow( $window , 2 ) ## GW_HWNDNEXT
    }
}

[bool]$checkedUserSettings = $false
[string]$regPath = $null
[int]$oldValue = -1
[string]$hkcu = $null
[int]$killedCount = 0

## if number of tabs found exceeds number of IE child processes then we haven't got a 1:1 mapping so killing one process could kill multiple tabs
[int]$ieChildProcesses = ( $ieProcesses.GetEnumerator() | Where-Object { ((Get-Process -Id $_.Value.ParentProcessId -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name) -eq 'iexplore') } | Measure-Object | Select -ExpandProperty Count )

if( $tabsFound -gt $ieChildProcesses -or ! $toKill.Count )
{
    if( ! $hkcu )
    {
        $hkcu = Get-HKCU -process ($ieProcesses.GetEnumerator()|Select-Object -First 1 -ExpandProperty Value)
    }
    ## check user settings to see if tab will respawn if killed
    if( $hkcu )
    {
        $regPath = Join-Path -Path $hkcu -ChildPath 'SOFTWARE\Microsoft\Internet Explorer\Main'
        $value = Get-ItemProperty -Path $regPath -Name 'TabProcGrowth' -ErrorAction SilentlyContinue | Select -ExpandProperty 'TabProcGrowth'
        ## if missing then is a problem on Windows 10 but not Server 2016
        [bool]$isWindows10 = ((Get-WmiObject -Class win32_operatingsystem|select -ExpandProperty Name) -match '^Microsoft Windows 10 ')
        if( $value -eq [int]0 -or $value -lt $tabsFound -or ( $value -eq $null -and $isWindows10 ) )
        {
            Write-Warning "Found $($tabsFound + $topLevelTabs.Count) tabs in total but only $ieChildProcesses IE child processes so consider setting 'TabProcGrowth' in 'HKCU\SOFTWARE\Microsoft\Internet Explorer\Main' to at least $([math]::Max($tabsFound,1))"
        }
    }
    ## Now see if have just top level tabs with just our URL in which case we will kill that otherwise we won't
    [int]$topLevelInnocentTabs = 0
    if( ! $toKill.Count )
    {
        $topLevelTabs.GetEnumerator() | ForEach-Object `
        {
            if( $_.Name -notmatch $URLtoKill )
            {
                $topLevelInnocentTabs++
            }
            else
            {
                $toKill.Add( $_.Value.Process.ProcessId , $_.Value ) ## we will delete later if innocents found
            }
        }
        if( $topLevelInnocentTabs -and $toKill.Count )
        {
            if( $killRegardless )
            {
                Write-Warning "Tab with specified text is a shared top level tab so killing it will kill those tabs too" 
            }
            else
            {            
                Write-Warning "Found $topLevelInnocentTabs other tab$(if( $topLevelInnocentTabs -ne 1 ) { 's' }) in top level IE process so not killing"
                $toKill = @{}
            }
        }
    }
}

if( $toKill.Count )
{
    ForEach( $kill in $toKill.GetEnumerator() )
    {
        if( ! $checkedUserSettings -and $kill.Value.Process )
        {
            if( ! $hkcu )
            {
                $hkcu = Get-HKCU -process $kill.Value.Process
            }
            ## check user settings to see if tab will respawn if killed
            if( $hkcu )
            {
                $regPath = Join-Path -Path $hkcu -ChildPath 'SOFTWARE\Microsoft\Internet Explorer\Recovery'
                $value = Get-ItemProperty -Path $regPath -Name 'AutoRecover' -ErrorAction SilentlyContinue | Select -ExpandProperty 'AutoRecover'
                if( $value -ne 2 ) ## if missing then also respawns
                {
                    [string]$warning = "User has the `"Enable automatic crash recovery`" IE setting enabled so "
                    if( $forceKill )
                    {
                        Set-ItemProperty -Path $regPath -Name 'AutoRecover' -Value ([int]2) -Force
                        $warning += 'disabled this'
                        $oldValue = $value
                    }
                    else
                    {
                        $warning += 'killed tabs will re-open'
                    }
                    Write-Warning $warning
                }
            }

            $checkedUserSettings = $true
        }
        if( $notToKill.Count )
        {
            ## check if any tab not being killed shares the same process as one that is
            $innocent = $notToKill[ $kill.Name ]
            if( $innocent )
            {
                if( $killRegardless )
                {
                    Write-Warning "PID $($kill.Name) also hosts these tabs which will be killed:`n`t$($innocent -join "`n`t")"
                }
                else
                {
                    Write-Warning "Found other tabs in the same IE process so not killing"
                    continue
                }
            }
        }
        $killed = Invoke-CimMethod -InputObject $kill.Value.Process -MethodName Terminate
        if( ! $killed -or $killed.ReturnValue )
        {
            Write-Warning "Failed to kill PID $($kill.Name), returned $($killed|Select-Object -ExpandProperty ReturnValue)"
        }
        $killedCount++
    }
}

<#
if( $forceKill -and $oldValue -ge 0 -and $regPath )
{
    Set-ItemProperty -Path $regPath -Name 'AutoRecover' -Value $oldValue
}
#>

"Killed $killedCount process$(if( $killedCount -ne 1) { 'es' })"

if( $VerbosePreference -eq 'Continue' -and $killedCount )
{
    $toKill.GetEnumerator() | Select -ExpandProperty Value | Where-Object { ! $_.Process.PSobject.properties[ 'CUNotKilled' ] } | Format-Table @{n='  Process Id';e={("{0,10}" -f $_.Process.ProcessId)}},Title | Out-String ## left padding pid with spaces
}

