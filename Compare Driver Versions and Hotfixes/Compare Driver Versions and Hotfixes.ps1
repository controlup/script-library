#requires -version 3

<#
.SYNOPSIS
    Show driver differences between 2 machines
    
.DESCRIPTION
    Account running the script must have remote WMI/CIM access to the toher machine.

.PARAMETER otherMachine
    The name of the other machine to compare with this one

.PARAMETER runningOnly
    Whether to compare all drivers or jsut those currently running on the machine where the script is run

.NOTES
    Modification History:

    2024/01/30  @guyrleech  Script born
    2024/02/23  @guyrleech  Added help
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true)]
    [string]$otherMachine ,
    [ValidateSet('yes','no','true','false')]
    [string]$runningOnly = 'yes'
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 400
try
{
    if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
    {
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions
    }
}
catch
{
    ## not a showstopper
}

if( $otherMachine -ieq $env:COMPUTERNAME )
{
    Throw "Both machines specified are $otherMachine"
}

## Win32_PnPSignedDriver gives us version info, win32_systemdriver does not but does give us full path to driver file

$remoteSessionOptions = New-CimSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
$remoteSession = $null

try
{
    $remoteSession = New-CimSession -ComputerName $otherMachine -SessionOption $remoteSessionOptions -ErrorAction Continue
    if( $null -eq $remoteSession )
    {
        Throw "Failed to create remote CIM session to $otherMachine"
    }

    $localOS = Get-CimInstance -ClassName Win32_OperatingSystem
    $remoteOS = $null
    $remoteOS = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $remoteSession

    [array]$localSystemDrivers  = @( Get-CimInstance -ClassName Win32_SystemDriver )
    [array]$remoteSystemDrivers = @( Get-CimInstance -ClassName Win32_SystemDriver -CimSession $remoteSession)
    $notFoundRemotely = New-Object -TypeName System.Collections.Generic.List[object]

    Write-Verbose -Message "$($localSystemDrivers.Count) local system drivers, $($remoteSystemDrivers.Count) remote"

    [array]$differences = @( ForEach( $driver in $localSystemDrivers ) ## this gives us driver file name so we can look at actual versions which we need to do reagardless of what driver info says as that comes from inf file and cannot be trusted
    {
        if( ( $runningOnly -match 'yes|true' -and $driver.State -ieq 'Running' ) -or $runningOnly -match 'no|false' )
        {
            $localFileDetails = Get-ItemProperty -Path ([Environment]::ExpandEnvironmentVariables( $driver.PathName ) -replace '^\\\?\?\\' ) | Select-Object -ExpandProperty VersionInfo
            $remoteMatches = $remoteSystemDrivers | Where-Object { $_.Name -eq $driver.Name } ## bowser is one that has different description on Server 2016!
            if( $null -eq $remoteMatches -or $remoteMatches.Count -eq 0 )
            {
                $notFoundRemotely.Add( (Add-Member -InputObject $driver -MemberType NoteProperty -Name FileDetails -Value $localFileDetails -PassThru ) )
            }
            else ## some matching drivers present
            {
                if( $remoteMatches -and $remoteMatches -is [array] -and $remoteMatches.Count -gt 1 )
                {
                    ## TODO check if same version but if more than 1, how do we differentiate and do we need to look for other local drivers matching and process them all here and not check the similar ones again later ?
                    ## will get multiple matches for things like processors
                    Write-Verbose -Message "$($driver.Description): $($driver.FriendlyName) has $($remoteMatches.Count) matches"
                }
                else ## only 1 matching driver
                {   
                    $remoteFileDetails = $null
                    ## some paths have \\??\ at the start so strip that and escape \ with extra \
                    $remoteFileDetails = Get-CimInstance -ClassName CIM_DataFile -Filter "Name = '$([Environment]::ExpandEnvironmentVariables( $driver.PathName ) -replace '^\\\?\?\\' -replace '\\' , '\\')'" -CimSession $remoteSession -Verbose:$false
                    if( $localFileDetails.FileVersionRaw.ToString() -ne $remoteFileDetails.Version )
                    {
                        Write-Verbose -Message "** $($driver.Caption) has different version numbers **"
                        $result = [pscustomobject]@{
                            Driver = $driver.Caption
                            File = $driver.PathName
                            'Version' = $localFileDetails.FileVersionRaw
                            'Remote Version' = $remoteFileDetails.Version
                            'Local Newer' = $localFileDetails.FileVersionRaw -gt [version]$remoteFileDetails.Version
                        }
                        if( $runningOnly -notmatch 'yes|true' )
                        {
                            Add-Member -InputObject $result -MemberType NoteProperty -Name State -Value $driver.State
                        }
                        Add-Member -InputObject $result -MemberType NoteProperty -Name 'Remote State' -Value $remoteMatches.State
                        $result ## output
                    }
                }
            }
        }
        else
        {
            Write-Verbose -Message "Ignoring `"$($driver.DisplayName)`" as not running"
        }
    })

    Write-Verbose -Message "Got $($notFoundRemotely.count) drivers out of $($localSystemDrivers.Count) not found on $otherMachine ($($remoteSystemDrivers.Count) drivers), $($differences.Count) with differences"

    [array]$notFoundLocally = @( ForEach( $remoteDriver in $remoteSystemDrivers )
    {
        ## check present locally
        $localMatches = @( $localSystemDrivers | Where-Object Name -ieq $remoteDriver.Name )
        if( $null -eq $localMatches -or $localMatches.Count -eq 0 )
        {
            ## fetch remote file version info so we have vendor in case not obvious/Microsoft
            $remoteFileDetails = $null
            ## environment variables expanded locally not on remote system but most likely are the same
            $remoteFileDetails = Get-CimInstance -ClassName CIM_DataFile -Filter "Name = '$( [Environment]::ExpandEnvironmentVariables( $remoteDriver.PathName ) -replace '^\\\?\?\\' -replace '\\' , '\\')'" -CimSession $remoteSession -Verbose:$false
            Add-Member -InputObject $remoteDriver -MemberType NoteProperty -Name FileDetails -Value $remoteFileDetails -PassThru ## output
        }
    })

    Write-Verbose -Message "Found $($notFoundLocally.Count) drivers which exist remotely but are not present locally"

    if( $localOS.Caption -ine $remoteOS.Caption )
    {
        Write-Warning -Message "Comparing different operating systems : $($localOS.Caption) and $($remoteOS.Caption)"
    }
    elseif( $localOS.BuildNumber -ine $remoteOS.BuildNumber )
    {
        Write-Warning -Message "Comparing different build numbers of $($localOS.Caption) : $($localOS.BuildNumber) and $($remoteOS.BuildNumber)"
    }

    Write-Output -InputObject "Found $($differences.Count) drivers with different driver versions:"
    $differences | Format-Table -AutoSize

    Write-Output -InputObject "Found $($notFoundLocally.Count) drivers on remote system that are not present locally"
    $notFoundLocally | Select-Object -Property Name,ServiceType,Description,@{name='Manufacturer';expression= {$_.FileDetails.Manufacturer}},@{name='Version';expression= {$_.FileDetails.Version}} | Format-Table -AutoSize

    Write-Output -InputObject "Found $($notFoundRemotely.Count) drivers locally that are not present on remote system"
    $notFoundRemotely | Select-Object -Property Name,State,ServiceType,@{name='Description';expression= {$_.FileDetails.FileDescription}},@{name='Manufacturer';expression= {$_.FileDetails.CompanyName}},@{name='Version';expression= {$_.FileDetails.FileVersion}} | Format-Table -AutoSize

    [hashtable]$localHotfixes = @{}
    [hashtable]$remoteHotfixes = @{}
    Get-CimInstance -ClassName Win32_QuickFixEngineering | ForEach-Object `
    {
        try
        {
            $localHotfixes.Add( $_.HotFixId , $_ )
        }
        catch
        {
        $null
        }
    }
    Get-CimInstance -ClassName Win32_QuickFixEngineering -CimSession $remoteSession | ForEach-Object `
    {
        try
        {
            $remoteHotfixes.Add( $_.HotFixID , $_ )
        }
        catch
        {
        $null
        }
    }

    [array]$hotfixOnlyLocally = @( $localHotfixes.GetEnumerator() | Where-Object { -Not $remotehotfixes[ $_.Name ] } )
    [array]$hotfixOnlyRemote  = @( $remoteHotfixes.GetEnumerator() | Where-Object { -Not $localHotfixes[ $_.Name ] } )

    if( $hotfixOnlyLocally.Count -eq 0 -and $hotfixOnlyRemote.Count -eq 0 )
    {
        Write-Output -InputObject "Both systems have the same $($localHotfixes.Count) hotfixes"
    }
    if( $hotfixOnlyLocally.Count -gt 0 )
    {
        Write-Output -InputObject "$($hotfixOnlyLocally.Count) hotfixes out of $($localHotfixes.Count) only on local system:"
        $hotfixOnlyLocally.GetEnumerator() | Select-Object -ExpandProperty value | Sort-Object -Property InstalledOn | Select-Object HotfixId,Description,@{name='Installed';expression={ $_.InstalledOn.ToString( 'd' ) }} | Format-Table -AutoSize
    }
    if( $hotfixOnlyRemote.Count -gt 0 )
    {
        Write-Output -InputObject "$($hotfixOnlyRemote.Count) hotfixes out of $($remoteHotfixes.Count) only on remote system:"
        $hotfixOnlyRemote.GetEnumerator() | Select-Object -ExpandProperty value | Sort-Object -Property InstalledOn | Select-Object HotfixId,Description,@{name='Installed';expression={ $_.InstalledOn.ToString( 'd' ) }} | Format-Table -AutoSize
    }
}
catch
{
    throw
}
finally
{
    if( $remoteSession )
    {
        Remove-CimSession -CimSession $remoteSession
        $remoteSession = $null
    }
}
