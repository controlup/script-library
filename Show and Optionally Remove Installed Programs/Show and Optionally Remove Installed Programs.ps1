#requires -version 3
<#
.SYNOPSIS
Retrieve information on installed programs

.DESCRIPTION
Does not use WMI/CIM since Win32_Product only retrieves MSI installed software and can be slow. Instead it processes the "Uninstall" registry key(s) which is also usually much faster
From https://github.com/guyrleech/Microsoft/blob/master/Get%20installed%20software.ps1
Based on code from https://blogs.technet.microsoft.com/heyscriptingguy/2011/11/13/use-powershell-to-quickly-find-installed-software/

.PARAMETER computers
Comma separated list of computer names to query. Use . to represent the local computer

.PARAMETER exportcsv
The name and optional path to a non-existent csv file which will have the results written to it

.PARAMETER productname
Only show products where the display name matches this regular expression 

.PARAMETER vendor
Only show products where the publisher matches this regular expression 

.PARAMETER gridview
The output will be presented in an on screen filterable/sortable grid view. Lines selected when OK is clicked will be placed in the clipboard

.PARAMETER removePattern
Comma separated list of one or more package names, or patterns that match one or more package names, that will be uninstalled.

.PARAMETER silent
Try and run the uninstall silently. This only works where the uninstall program is msiexec.exe

.PARAMETER asjson
Output objects as json

.PARAMETER ascsv
Output objects as csv

.PARAMETER importcsv
A csv file containing a list of computers to process where the computer name is in the "ComputerName" column unless specified via the -computerColumnName parameter

.PARAMETER computerColumnName
The column name in the csv specified via -importcsv which contains the name of the computer to query

.PARAMETER includeEmptyDisplayNames
Includes registry keys which have no "DisplayName" value which may not be valid installed packages

.EXAMPLE
& '.\Get installed software.ps1' -gridview -computers .
Retrieve installed software details on the local computer and show in a grid view

.EXAMPLE
& '.\Get installed software.ps1' -gridview -computers . -ascsv -quiet -productname 'Teams' -vendor 'Microsoft'
Retrieve installed software details on the local computer where the display name contains "Teams" and the vendor contains "Microsoft".
Output the results to the pipeline in csv format , whilst suppressing warnings.

.EXAMPLE
& '.\Get installed software.ps1' -computers computer1,computer2,computer3 -exportcsv installed.software.csv
Retrieve installed software details on the computers computer1, computer2 and computer3 and write to the CSV file "installed.software.csv" in the current directory

.EXAMPLE
& '.\Get installed software.ps1' -gridview -importcsv computers.csv -computerColumnName Machine
Retrieve installed software details on computers in the CSV file computers.csv in the current directory where the column name "Machine" contains the name of the computer and write the results to standard output

.EXAMPLE
& '.\Get installed software.ps1' -gridview -computers . -uninstall
Retrieve installed software details on the local computer and show in a grid view. Packages selected after OK is clicked in the grid view will be uninstalled.

.EXAMPLE
& '.\Get installed software.ps1' -computers . -remove 'Notepad++*' -Confirm:$false
Retrieve installed software details on the local computer and and remove Notepad++ without asking for confirmation.

.EXAMPLE
& '.\Get installed software.ps1' -computers . -remove '*Acrobat Reader*' -Confirm:$false -silent
Retrieve installed software details on the local computer and and remove Adobe Acrobat Reader silently, so without any user prompts, and without asking for confirmation.

.NOTES
THE SCRIPT IS PROVIDED IN AN "AS IS" CONDITION, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL CONTROLUP, ANY AUTHORS OR ANY COPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SCRIPT OR THE USE OR OTHER DEALINGS IN THE SCRIPT.

Modification History:
16/11/18	GRL		Added functionality to uninstall items selected in grid view
18/11/18	GR		Added -remove to remove one or more packages without needing grid view display
					Added -silent to run silent uninstalls where installer is msiexec
					Added -quiet option
02/03/19	GRL		Added SystemComponent value
23/07/19	GRL		Added HKU searching
14/10/20	GRL		Added default parameter set name and hiding error if can't open reg key
					Added InstallSource
22/05/21	GRL		Added json and csv direct output. Added -delimiter for csv outputs for ze Dutch
24/05/21	GRL		Added parameters to filter on product name and/or vendor
07/02/22	GRL		Changed removeRegex to removePattern. Added maximumRemovals and timeoutMinutes. Added checks for child processes of uninstall process to complete
#>

[CmdletBinding()]
Param
(
    [ValidateSet('yes','no')]
    [string]$noRemoval = 'yes' ,
	[decimal]$timeoutMinutes = 5 ,
    [string]$removePattern ,
    [string]$productname ,
    [string]$vendor ,
    [switch]$silent = $true,
    [switch]$quiet ,
    [switch]$asjson ,
    [switch]$ascsv ,
    [string]$delimiter = ',' ,
    [switch]$includeEmptyDisplayNames
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

## technically the script can remove any number of packages but to reduce the risk of inadvertent removals, it is limited to 1 - CHANGE AT YOUR OWN RISK
[int]$maximumRemovals = 1

[int]$outputWidth = 400
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

Function Stop-ChildProcesses
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        $ParentProcess ,
        [switch]$kill ## could just be counting to see if there are child processes we need to wait for as some uninstallers spawn a new process but don't wait for it to exit
    )
    
    Get-CimInstance -ClassName win32_process -ErrorAction SilentlyContinue -Filter "ParentProcessId = '$($ParentProcess.Id)'" -Verbose:$false | ForEach-Object `
    {
        if( $thisProcess = Get-Process -Id $_.ProcessId -ErrorAction SilentlyContinue -IncludeUserName )
        {
            if( $thisProcess.StartTime -gt $parentProcess.StartTime -and $parentProcess.SessionId -eq $thisProcess.SessionId -and $thisProcess.UserName -eq $parentProcess.UserName )
            {
                Stop-ChildProcesses -ParentProcess $thisProcess -kill:$kill
                Write-Verbose -Message "$(if( -Not $kill ) { 'Not ' })Killing $($thisProcess.Name) ($($thisProcess.Id)), parent $($parentProcess.Name) ($($parentProcess.Id))"
                if( $kill )
                {
                    Stop-Process -InputObject $thisProcess -Force -ErrorAction SilentlyContinue
                }
                Add-Member -InputObject $thisProcess -MemberType NoteProperty -Name ParentProcess -Value $parentProcess -PassThru ## return
            }
            else
            {
                Write-Verbose -Message "Ignoring child process $($thisProcess.Name) ($($thisProcess.Id)), parent $($parentProcess.Name) ($($parentProcess.Id)), parent sid $($parentProcess.SessionId) this sid $($thisProcess.SessionId) username $($thisProcess.UserName) vs $($parentProcess.UserName)"
            }
        }
    }
}

Function Remove-Package()
{
    [CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]
    Param
    (
        $package ,
        [switch]$silent ,
        [int]$timeoutSeconds
    )

    [bool]$uninstallerRan = $false

    Write-Verbose "Removing `"$($package.DisplayName)`""

    if( ! [string]::IsNullOrEmpty( $package.Uninstall ) )
    {
        ## need to split uninstall line so we can pass to Start-Process since we need to wait for each to finish in turn
        [string]$executable = $null
        [string]$arguments = $null
        if( $package.Uninstall -match '^"([^"]*)"\s?(.*)$' `
            -or $package.Uninstall -match '^(.*\.exe)\s?(.*)$' ) ## cope with spaces in path but no quotes
        {
            $executable = $Matches[1]
            $arguments = $Matches[2].Trim()
        }
        else ## unquoted so see if there's a space delimiting exe and arguments
        {
            [int]$space = $package.Uninstall.IndexOf( ' ' )
            if( $space -lt 0 )
            {
                $executable = $package.Uninstall
            }
            else
            {
                $executable = $package.Uninstall.SubString( 0 , $space )
                if( $space -lt $package.Uninstall.Length )
                {
                    $arguments = $package.Uninstall.SubString( $space ).Trim()
                }
            }
        }
        [hashtable]$processArguments = @{
            'FilePath' = $executable
            'PassThru' = $true
            'Wait' = $false
        }
        if( $silent )
        {
            if( $executable -match '^msiexec\.exe$' -or $executable -match '^msiexec$' -or $executable -match '[^a-z0-9_]msiexec\.exe$' -or $executable -match '[^a-z0-9_]msiexec$' )
            {
                ## Some uninstallers pass /I as they are meant to be interactive so we'll change this to /X
                $arguments = ($arguments -replace '/I' , '/X') + ' /qn /norestart'
            }
            else
            {
                ## NSIS installers use /S (case sensitive)
                ## if unins000.exe then try /VerySilent /NoRestart
                Write-Warning "Don't know how to run silent uninstall for package `"$($package.DisplayName)`", uninstaller `"$executable`""
            }
        }
        if( ! [string]::IsNullOrEmpty( $arguments ) )
        {
            $processArguments.Add( 'ArgumentList' , $arguments )
        }
        Write-Verbose "Running $executable `"$arguments`" for $($package.DisplayName) ..."
        [System.Diagnostics.Trace]::WriteLine( "Running $executable `"$arguments`" for $($package.DisplayName) ..." )
 
        $uninstallProcess = $null
        $uninstallProcess = Start-Process @processArguments

        if( $uninstallProcess ) 
        {
            [datetime]$timeoutTime = $uninstallProcess.StartTime.AddSeconds( $timeoutSeconds )

            Write-Verbose -Message "$(Get-Date -Format G): Uninstaller pid $($uninstallProcess.Id) , timeout at $(Get-Date -Date $timeoutTime -Format G)"

            $waitError = $null
            [hashtable]$waitParameters = @{
                InputObject = $uninstallProcess
                ErrorAction = 'SilentlyContinue'
                ErrorVariable = 'waitError'
            }
            if( $timeoutSeconds -gt 0 )
            {
                $waitParameters.Add( 'Timeout' , $timeoutSeconds )
            }

            ## when we get child processes, we need the username of the installer process as will not be $username if running as system (will be computername$)
            Add-Member -InputObject $uninstallProcess -MemberType NoteProperty -Name Username -Value (Get-Process -InputObject $uninstallProcess -IncludeUserName | Select-Object -ExpandProperty Username) -Force

            Wait-Process @waitParameters

            [bool]$waitResult = $?
            [int]$childProcesses = 0
            [array]$childProcesses = @( )
            [bool]$timedOut = $false

            if( $waitResult )
            {
                ## need to see if any child processes still running as the main uninstaller may have started another process which hasn't exited (e.g. notepad++ uninstaller)    
                do
                {
                    $childProcesses = @( Stop-ChildProcesses -ParentProcess $uninstallProcess -kill:$false | Sort-Object -Property StartTime )
                    if( $childProcesses -and $childProcesses.Count -gt 0 )
                    {
                        try
                        {
                            Wait-Process -Id $childProcesses[0].Id -Timeout ($timeoutTime - [datetime]::Now).TotalSeconds -ErrorAction SilentlyContinue
                        }
                        catch
                        {
                            ## so we don't have to verify that the timeout is positive
                        }
                    }
                } while( [datetime]::now -le $timeoutTime -and $childProcesses -and $childProcesses.Count -gt 0 )
                
                if( $timedOut = $childProcesses -and $childProcesses.Count -gt 0 )
                {
                    $waitResult = $false
                }
            }

            if( $waitResult )
            {
                Write-Verbose "Uninstall exited with code $($uninstallProcess.ExitCode)"
                ## https://docs.microsoft.com/en-us/windows/desktop/Msi/error-codes
                if( $uninstallProcess.ExitCode -eq 3010 ) ## maybe should check it's msiexec that ran
                {
                    Write-Warning "Uninstall of `"$($package.DisplayName)`" requires a reboot"
                }
                $uninstallerRan = $true
            }
            elseif( $timedOut -or ( $waitError -and $waitError.Count -and $waitError[0].Exception -and $waitError[0].Exception.HResult -eq 0x80131505 ) ) ## language neutral
            {
                Write-Warning -Message "Timeout waiting for uninstall process $($package.Uninstall) (pid $($uninstallProcess.Id)) to finish"
                ## kill it and all child processes
                [array]$killed = @( Stop-ChildProcesses -ParentProcess $uninstallProcess -kill:$true )
                Stop-Process -InputObject $uninstallProcess -Force
                Write-Host -Object "Killed uninstall process and $($killed.Count) child processes"
            }
            else
            {
                Write-Warning -Message "Unexpected issue waiting for uninstall process $($package.Uninstall) (pid $($uninstallProcess.Id)) to finish - $waitError"
            }
        }
    }
    else
    {
        Write-Warning "Unable to uninstall `"$($package.DisplayName)`" as it has no uninstall string"
    }
    $uninstallerRan
}

Function Process-RegistryKey
{
    [CmdletBinding()]
    Param
    (
        [string]$hive ,
        $reg ,
        [string[]]$UninstallKeys ,
        [switch]$includeEmptyDisplayNames ,
        [AllowNull()]
        [string]$username ,
        [AllowNull()]
        [string]$productname ,
        [AllowNull()]
        [string]$vendor
    )

    ForEach( $UninstallKey in $UninstallKeys )
    {
        $regkey = $reg.OpenSubKey($UninstallKey) 
    
        if( $regkey )
        {
            [string]$architecture = if( $UninstallKey -match '\\wow6432node\\' ){ '32 bit' } else { 'Native' } 

            $subkeys = $regkey.GetSubKeyNames() 
    
            foreach($key in $subkeys)
            {
                $thisKey = Join-Path -Path $UninstallKey -ChildPath $key 

                $thisSubKey = $reg.OpenSubKey($thisKey) 

                if( $includeEmptyDisplayNames -or ! [string]::IsNullOrEmpty( $thisSubKey.GetValue('DisplayName') ) )
                {
                    [string]$installDate = $thisSubKey.GetValue('InstallDate')
                    $installedOn = New-Object -TypeName 'DateTime'
                    if( [string]::IsNullOrEmpty( $installDate ) -or ! [datetime]::TryParseExact( $installDate , 'yyyyMMdd' , [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$installedOn ) )
                    {
                        $installedOn = $null
                    }
                    $size = New-Object -TypeName 'Int'
                    if( ! [int]::TryParse( $thisSubKey.GetValue('EstimatedSize') , [ref]$size ) )
                    {
                        $size = $null
                    }
                    else
                    {
                        $size = [math]::Round( $size / 1KB , 1 ) ## already in KB
                    }

                    if( $thisSubKey.GetValue('DisplayName') -match $productname -and $thisSubKey.GetValue('Publisher') -match $vendor )
                    {
                        [pscustomobject][ordered]@{
                            ##'ComputerName' = $computername
                            'Hive' = $Hive
                            'User' = $username
                            'Key' = $key
                            'Architecture' = $architecture
                            'DisplayName' = $($thisSubKey.GetValue('DisplayName'))
                            'DisplayVersion' = $($thisSubKey.GetValue('DisplayVersion'))
                            'InstallLocation' = $($thisSubKey.GetValue('InstallLocation'))
                            'InstallSource' = $($thisSubKey.GetValue('InstallSource'))
                            'Publisher' = $($thisSubKey.GetValue('Publisher'))
                            'InstallDate' = $(if( $installedOn ) { Get-Date -Date $installedOn -Format d }) ## never includes time so no point outputting 00:00
                            'Size (MB)' = $size
                            'System Component' = $($thisSubKey.GetValue('SystemComponent') -eq 1)
                            'Comments' = $($thisSubKey.GetValue('Comments'))
                            'Contact' = $($thisSubKey.GetValue('Contact'))
                            'HelpLink' = $($thisSubKey.GetValue('HelpLink'))
                            'HelpTelephone' = $($thisSubKey.GetValue('HelpTelephone'))
                            'Uninstall' = $($thisSubKey.GetValue('UninstallString'))
                        }
                    }
                }
                else
                {
                    ## Write-Warning "Ignoring `"$hive\$thisKey`" on $computername as has no DisplayName entry"
                }

                $thisSubKey.Close()
            } 
            $regKey.Close()
        }
        elseif( $hive -eq 'HKLM' )
        {
            Write-Warning "Failed to open `"$hive\$UninstallKey`" on $computername"
        }
    }
}

if( $quiet )
{
    $VerbosePreference = $WarningPreference = 'SilentlyContinue'
}

[string[]]$UninstallKeys = @( 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall' , 'SOFTWARE\\wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall' )

if( $PSBoundParameters[ 'removePattern' ] -and -Not [string]::IsNullOrEmpty( $removePattern ) )
{
    ## check if string is all * characters
    if( $removePattern -match '^\**$' )
    {
        Throw "Search pattern $removePattern is too broad and not allowed by this script. If you would like to list all programs leave the Search Pattern empty."
    }
}

[array]$installed = @(
    [string]$computername = $env:COMPUTERNAME

    $reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey( 'LocalMachine' , $computername )
    
    if( $? -and $reg )
    {
        Process-RegistryKey -Hive 'HKLM' -reg $reg -UninstallKeys $UninstallKeys -includeEmptyDisplayNames:$includeEmptyDisplayNames -productname $productname -vendor $vendor
        $reg.Close()
    }
    else
    {
        Throw "Failed to open HKLM on $computername"
    }

    $reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey( 'Users' , $computername )
    
    if( $? -and $reg )
    {
        ## get each user SID key and process that for per-user installed apps
        ForEach( $subkey in $reg.GetSubKeyNames() )
        {
            try
            {
                if( $userReg = $reg.OpenSubKey( $subKey ) )
                {
                    [string]$username = $null
                    try
                    {
                        $username = ([System.Security.Principal.SecurityIdentifier]($subKey)).Translate([System.Security.Principal.NTAccount]).Value
                    }
                    catch
                    {
                        $username = $null
                    }
                    Process-RegistryKey -Hive (Join-Path -Path 'HKU' -ChildPath $subkey) -reg $userReg -UninstallKeys $UninstallKeys -includeEmptyDisplayNames:$includeEmptyDisplayNames -user $username  -productname $productname -vendor $vendor
                    $userReg.Close()
                }
            }
            catch
            {
            }
        }
        $reg.Close()
    }
    else
    {
        Write-Warning "Failed to open HKU on $computername"
    }
) | Sort -Property ComputerName, DisplayName

[int]$uninstalled = 0

if( $installed -and $installed.Count )
{
    Write-Verbose "Found $($installed.Count)"
    
    [System.Collections.Generic.List[string]]$excluded = @( 'Hive' , 'Key' , 'Help*', 'Uninstall' , 'Contact' , 'InstallSource' )
    
    if( ($installed | Where-Object { ! [string]::IsNullOrEmpty( $_.User ) } | Measure-Object ).Count -eq 0 )
    {
        $excluded.Add( 'User' )
    }

    if( $PSBoundParameters[ 'removePattern' ] -and ! [string]::IsNullOrEmpty( $removePattern ) )
    {
        [int]$matched = 0
        $removedPackages = New-Object -TypeName System.Collections.Generic.List[object] -ArgumentList @()
        [array]$toRemove = @( $installed | Where-Object { $_.DisplayName -like $removePattern } )
        if( $toRemove.Count -gt $maximumRemovals )
        {
            Throw "$($toRemove.Count) packages matched `"$removePattern`" which exceeds script maximum removal count of $maximumRemovals"
        }

        $removed = @( ForEach( $package in $toRemove )
        {
            $removedPackages.Add( $package )
            if( $noRemoval -and $noRemoval.Length -and $noRemoval[0] -ine 'y' )
            {
                if( Remove-Package -Package $package -silent:$silent -timeoutSeconds ( $timeoutMinutes * 60 ) )
                {
                    $uninstalled++
                    $package
                }
            }
        })
        
        [string]$before = $null
        [string]$after = $null
        if( $noRemoval -and $noRemoval.Length -and $noRemoval[0] -ine 'y' )
        {
            $before = "Ran"
            $after = ", $uninstalled uninstallers ran ok"
        }
        else
        {
            $before = "Would have run"
        }
        
        if( ! $removedPackages -or ! $removedPackages.Count )
        {
            Write-Warning "No uninstallers $($before.ToLower()) as none of the $($installed.Count) packages found matched $removePattern"
        }
        else
        {
            Write-Output "`n$before uninstaller for $($removedPackages.Count) matches$after :"
        
            $removedPackages | Select-Object -Property * -ExcludeProperty $excluded | Format-Table -AutoSize
        }
    }
    else
    {
        $installed | Select-Object -Property * -ExcludeProperty $excluded | Format-Table -AutoSize
    }
}
else
{
    Write-Warning "Found no installed products in the registry"
}

