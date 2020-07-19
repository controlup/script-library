#Requires -version 3.0

<#
    Change or show StoreFront tracing settings on one or more SF servers.
    Note that this will restart various services so a brief loss of service is possible when seting the trace level

    Use of this script is entirely at your own risk - the author cannot be held responsible for any undesired effects deemed to have been caused by this script.

    @guyrleech, 2018
#>

## arguments
##   1 required trace level
##   2 operate on all servers in a cluster

[bool]$cluster = $false
[string[]]$servers = @( $env:COMPUTERNAME ) 
[string]$traceLevel = $null
[string[]]$validTraceLevels = @('Off', 'Error','Warning','Info','Verbose')
[int]$outputWidth = 400
$VerbosePreference = 'SilentlyContinue'

if( $args.Count -ge 2 )
{
    $traceLevel = $args[1]
}

if( $traceLevel -and $traceLevel -notin $validTraceLevels  )
{
    Throw "Illegal trace level `"$tracelevel`" specified - valid values are $($validTraceLevels -join ' , ')"
}

if( $args.Count -ge 1 )
{
    $cluster = ( $args[0] -eq 'true' )    
}

## Variables below here generally should not need to be changed
[string]$webConfig = 'web.config' 
[string]$installDirKey = 'SOFTWARE\Citrix\DeliveryServices'
[string]$installDirValue = 'InstallDir'
[string]$moduleInstaller = 'Scripts\ImportModules.ps1'
[string]$diagnosticsNode = 'configuration/system.diagnostics/switches/add'
[string]$logfileNode = 'configuration/system.diagnostics/sharedListeners/add'

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

if( [string]::IsNullOrEmpty( $traceLevel ) )
{
    [string]$snapin = 'Citrix.DeliveryServices.Web.Commands'
}
else
{
    [string]$snapin = 'Citrix.DeliveryServices.Framework.Commands' 
}

## retrieve versions so store for second pass
[hashtable]$StoreFrontVersions = @{}

## We will retrieve all the cluster members - although multiple SF servers may have been specified, they may be members of different clusters so we get all unique names
if( $cluster )
{
    $newServers = New-Object -TypeName System.Collections.ArrayList
    ForEach( $server in $servers )
    {
        ## Read install dir from registry so we don't have to load all SF cmdlets
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$server)
        if( $reg )
        {
            $RegSubKey = $Reg.OpenSubKey($installDirKey)

            if( $RegSubKey )
            {
                $installDir = $RegSubKey.GetValue($installDirValue) 
                if( ! [string]::IsNullOrEmpty( $installDir ) )
                {
                    $script = Join-Path $installDir $moduleInstaller
                    [string]$version,[string[]]$clusterMembers = Invoke-Command -ComputerName $server -ScriptBlock `
                    {
                        & $using:script
                        (Get-DSVersion).StoreFrontVersion
                        @( (Get-DSClusterMembersName).Hostnames )
                    } 
                    if( ! [string]::IsNullOrEmpty( $version ) )
                    {
                        $StoreFrontVersions.Add( $server , $version )
                    }
                    if( $clusterMembers -and $clusterMembers.Count )
                    {
                        ## now iterate through and add to servers array if not present already although indirectly since we are already iterating over this array so cannot change it
                        $clusterMembers | ForEach-Object `
                        {
                            if( $servers -notcontains $_ -and $newServers -notcontains $_ )
                            {
                                $null = $newServers.Add( $_ )
                            }
                        }
                    }
                    else
                    {
                        Write-Warning "No cluster members found via $server"
                    }
                }
                else
                {
                    Write-Error "Failed to read value `"$installDirValue`" from key HKLM\$installDirKey on $server"
                }
                $RegSubKey.Close()
            }
            else
            {
                Write-Error "Failed to open key HKLM\$installDirKey on $server"
            }
            $reg.Close()
        }
        else
        {
            Write-Error "Failed to open key HKLM on $server"
        }
    }
    if( $newServers -and $newServers.Count )
    {
        Write-Verbose "Adding $($newServers.Count) servers to action list: $($newServers -join ',')"
        $servers += $newServers
    }
    else
    {
        Write-Warning "Only a single server found in this cluster"
    }
}

[int]$badServers = 0 
           
[array]$results = @( ForEach( $server in $servers )
{
    if( [string]::IsNullOrEmpty( $traceLevel ) )
    {
        ## keyed on file name with value as the XML from that file
        [hashtable]$configFiles = Invoke-Command -ComputerName $server -ScriptBlock `
        {
            Add-PSSnapin $using:snapin
            [hashtable]$files = @{}
            $dsWebSite = Get-DSWebSite
            $dsWebSite.Applications | ForEach-Object `
            {
                $app = $_
                [string]$configFile = Join-Path $app.Folder $using:webConfig
                [xml]$content = $null

                if( Test-Path $configFile -ErrorAction SilentlyContinue )
                {
                    $content = (Get-Content $configFile)
                    if( $content )
                    {
                        $files.Add( $configFile , $content )
                    }
                }
            }
            $files
        }
        Write-Verbose "Got $($configFiles.Count) $webConfig files from $server"
        [hashtable]$states = @{}
        $configFiles.GetEnumerator() | ForEach-Object `
        {
            [xml]$node = $_.Value
            [string]$fileName = $_.Key
            $diags = $null
            try
            {
                $diags = @( $node.SelectNodes( "//$diagnosticsNode" ) )
                $logFile = $node.SelectSingleNode( "//$logfileNode" )  ## should only be one
            }
            catch { }
            if( $diags )
            {
                $diags | ForEach-Object `
                {
                    $thisSwitch = $_
                    [string]$module = ($thisSwitch.Name -split '\.')[-1]
                    $info = $null 
                    try
                    {
                        $info = [pscustomobject]@{ 'Server' = $server ; 'Trace Level' = $thisSwitch.Value ; 'Config File' = $fileName ; 'Module' = $module }
                        [string]$version = $StoreFrontVersions[ $server ]
                        if( ! [string]::IsNullOrEmpty( $version ) )
                        {
                            Add-Member -InputObject $info -MemberType NoteProperty -Name 'StoreFront Version' -Value $version
                        }
                        if( $logFile )
                        {
                           Add-Member -InputObject $info -NotePropertyMembers @{ 'Log File' = $logFile.initializeData ; 'Max Size (KB)' = $logFile.maxFileSizeKB  }
                           ## Seems this isn't present for all SF versions
                           if( Get-Member -InputObject $logFile -Name fileCount -ErrorAction SilentlyContinue )
                           {
                               Add-Member -InputObject $info -MemberType NoteProperty -Name 'Log File Count' -Value $logfile.fileCount
                           }
                        }
                        $states.Add( $thisSwitch.Value , [System.Collections.ArrayList]( @( $info ) ) )
                    }
                    catch
                    {
                        if( ! [string]::IsNullOrEmpty( $thisSwitch.Name ) -and $info )
                        {
                            $null = $states[ $thisSwitch.Value ].Add( $info )
                        }
                    }
                }
            }
        }
        $states.GetEnumerator() | select -ExpandProperty Value ## push into results array
        if( $states.Count -gt 1 )
        {
            Write-Warning "Trace levels are inconsistent on $server - $(($states.GetEnumerator()|Select -ExpandProperty Name) -join ',')"
            $states.GetEnumerator() | ForEach-Object { Write-Verbose "$($_.Name) :`n`t$(($_.Value|select -ExpandProperty 'Config File') -join ""`n`t"" )" }
            $badServers++
        }
        elseif( ! $states.Count )
        {
            Write-Warning "No trace levels found on $server"
        }
        Write-Host "$server : logging level is $(($states.GetEnumerator()|select -ExpandProperty Name ) -join ',')" ## Can't be Write-Output otherwise will be captured into the results array
    }
    else
    {
        Write-Host "Setting trace level to $traceLevel on $server"
        Invoke-Command -ComputerName $server -ScriptBlock { Add-PSSnapin $using:snapin ; Set-DSTraceLevel -All -TraceLevel $using:traceLevel }
    }
} )

if( $results.Count )
{
    [string]$title = $( if( $badServers )
    {
        "Inconsistent settings found on $badServers out of"
    }
    else
    {
        ## Now check it's the same consistent setting across all servers
        [string]$lastLevel = $null
        [bool]$matching = $true
        [int]$different = 0
        ForEach( $server in $servers )
        {
            [string]$thisLevel = $results | Where-Object { $_.Server -eq $server } | Select -First 1 -ExpandProperty 'Trace Level'
            if( $lastLevel -and $thisLevel -ne $lastLevel)
            {
                $matching = $false
                $different++
            }
            $lastLevel = $thisLevel
        }
        if( $matching )
        {
            "Consistent settings found on all"
        }
        else
        {
            "Different settings found on $different out of"
        }
    } ) + " $($servers.Count) StoreFront servers $($servers -join ' ')"
    
    $title

    [hashtable]$params = @{ 'AutoSize' = $true}
    if( $badServers )
    {
        $fields = [System.Collections.ArrayList]( @( 'Module','Trace Level','Config File','Log File','Max Size (KB)','Log File Count' ) )
        $params.Add( 'GroupBy' , 'Trace Level' )
    }
    else
    {
        $fields = [System.Collections.ArrayList]( @( 'Log File','Max Size (KB)','Log File Count' ) )
    }
    if( $servers.Count -gt 1 )
    {
        $fields.Insert( 0 , 'Server' )
    }
    if( $badServers )
    {
        $results | Select $fields | Format-Table @params
    }
    else
    {
        $results | Select $fields | Sort -Unique 'Log File' | Format-Table @params
    }
}

