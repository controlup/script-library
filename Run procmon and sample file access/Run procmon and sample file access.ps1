#requires -version 3.0
<#
    Perform an automated procmon capture, convert results to csv and process for ControlUp SBA

    @guyrleech 2018

    Modification History:

    17/10/18  GRL  Added watcher process to terminate and remove backing file if SBA times out

    26/10/18  GRL  Workaround for procmon conversion bug when pml file is empty. Remove any *-1.pml, etc files

    30/10/18  GRL  Fixed issue where wouldn't download procmon. Added checks for procmon.exe signing

    06/08/19 GRL  Fixed certificate validation issue

#>

<#
    arguments:
        0    Pid or Session Id
        1    Run time
        2    Backing file
        3    Path to procmon
#>

$VerbosePreference = 'SilentlyContinue'

[bool]$perSession = $false
[bool]$perComputer = $false
[string]$procmon = $null
[int]$runTime = 15
[string]$backingFile = ( Join-Path $env:temp "controlup.sba.$pid.pml" )
[int]$processId = -1
[int]$sessionId = -1
[int]$outputWidth = 400
[int]$otherArgumentsStartIndex = 1
[bool]$downloadedProcmon = $false
[string[]]$arguments = @(
    'runTime' ,
    'backingFile' ,
    'procmon' )

Function Find-UnicodeString
{
    Param
    (
        [byte[]]$data ,
        [string]$searchFor ,
        [long]$offset = 0
    )
    $encoding = [System.Text.Encoding]::Unicode
    [byte[]]$searchBytes = $encoding.GetBytes( $searchFor )
    [int]$stringOffset = 0
    [long]$foundAt = -1
    For( [int]$index = 0 ; $index -lt $data.Count -and $foundAt -lt 0 ; $index++ )
    {
        if( $index -ge $offset -and $data[ $index ] -eq $searchBytes[ $stringOffset ] )
        {
            $stringOffset++
            if( $stringOffset -ge $searchBytes.Count )
            {
                $foundAt = $index - $searchBytes.Count + 1
            }
        }
        else
        {
            $stringOffset = 0
        }
    }
    $foundAt
}

if( $perSession )
{
    $sessionId = $args[0]
    $theSession = quser.exe $sessionId
    if( ! $theSession -or ! $theSession.Count )
    {
        Throw "Session id $sessionId not found"
    }
}
elseif( $perComputer )
{
    $otherArgumentsStartIndex = 0
}
else
{
    $processId = $args[0]
    if( ! ( Get-Process -Id $processId -ErrorAction SilentlyContinue ) )
    {
        Throw "Process id $processId no longer exists"
    }
}

[hashtable]$passedArguments = @{}

For( [int]$index = $otherArgumentsStartIndex ; $index -lt $args.Count ; $index++ )
{
    if( $args[ $index ] )
    {
        Set-Variable -Name $arguments[ $index - $otherArgumentsStartIndex ] -Value $args[ $index ] -ErrorAction Stop
        $passedArguments.Add( $arguments[ $index - $otherArgumentsStartIndex ] , $args[ $index ] )
    }
}

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

[string]$csvTrace = $backingFile -replace '\.[a-z0-9]+$' , '.csv'

# Get the procmon filter and inject our pid into it
# Filter: out: registry, network, profiling, process
# Filter: in: pid 123456
# Columns: add: duration, category
# Options: drop filtered packets

[string]$procmonFilter = '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'
[byte[]]$filter = [System.Convert]::FromBase64String( $procmonFilter )
[string]$pmcFile = $null
$monitorProcess = $null
$convertProcess = $null 

if( $processId -ge 0 )
{
    $filterRulesOffset = Find-UnicodeString -data $filter -searchFor 'FilterRules'
    if( $filterRulesOffset -lt 0 )
    {
        Write-Warning "Unable to find filter rule section in embedded procmon config file"
    }
    else
    {
        $pidOffset = Find-UnicodeString -data $filter -searchFor '123456' -offset $filterRulesOffset
        if( $pidOffset -lt 0 )
        {
            Write-Warning "Unable to find pid '123456' in embedded procmon config file"
        }
        else
        {
            [string]$paddedPid = $processId.ToString('000000')
            For( [int]$index = 0 ; $index -lt $paddedPid.Length ; $index++ )
            {
                $filter[ ($pidOffset + ($index * 2)) ] = [byte]$paddedPid[ $index ]
            }
            # also needs the binary pid in the filter after the null terminator in little endian format
            [int]$binaryPidOffset = $pidOffset + ($index + 1) * 2
            $filter[ $binaryPidOffset ] = [byte]($processId -band 0xff)
            $filter[ $binaryPidOffset + 1 ] = [byte](($processId -band 0xff00) -shr 8)
            $filter[ $binaryPidOffset + 2 ] = [byte](($processId -band 0xff0000) -shr 16)
            $filter[ $binaryPidOffset + 3 ] = [byte](($processId -band 0xff000000) -shr 24)

            $pmcFile = $backingFile -replace '\.[a-z0-9]+$' , '.pmc'

            $fileStream = New-Object System.IO.FileStream($pmcFile, [System.IO.FileMode]'Create', [System.IO.FileAccess]'Write')
            $fileStream.Write($filter,0,$filter.Count)
            $fileStream.Close()

            if( ! ( Test-Path -Path $pmcFile -ErrorAction SilentlyContinue ) )
            {
                Write-Warning "Failed to create procmon configuration file `"$pmcFile`""
            }
        }
    }
}

try
{
    $existingProcmons = Get-Process -Name procmon* -ErrorAction SilentlyContinue | Select -ExpandProperty Id
    if( $existingProcmons -and $existingProcmons.Count )
    {
        Throw "Procmon is already running (Process ids $($existingProcmons -join ','))"
    }

    Remove-Item -Path $backingFile -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $csvTrace -Force -ErrorAction SilentlyContinue

    if( [string]::IsNullOrEmpty( $procmon ) )
    {
        $procmon = Join-Path $env:TEMP 'procmon.exe'
        if( Test-Path $procmon -ErrorAction SilentlyContinue -PathType Leaf )
        {
            Throw "Procmon already exists at $procmon so not overwriting"
        }
        else
        {
            Write-Verbose "Downloading procmon ..."
            [Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            (New-Object System.Net.WebClient).DownloadFile( 'https://live.sysinternals.com/procmon.exe' , $procmon )
            if( ! ( Test-Path $procmon -ErrorAction SilentlyContinue -PathType Leaf ) )
            {
                Throw "Failed to download procmon"
            }
            Unblock-File -Path $procmon
            $downloadedProcmon = $true
            $signing = Get-AuthenticodeSignature -FilePath $procmon -ErrorAction SilentlyContinue
            if( ! $signing )
            {
                Throw "Could not get signing information from `"$procmon`""
            }
            if( $signing.Status -ne 'Valid' )
            {
                Throw "Certificate status for `"$procmon`" is $($signing.Status), not `"Valid`""
            }
            if( $signing.SignerCertificate.Subject -notmatch '^CN=Microsoft Corporation,' )
            {
                Throw "`"$procmon`" is not signed by Microsoft Corporation, found $($signing.SignerCertificate.Subject)"
            }
        }
    }
    elseif( ! ( Test-Path $procmon -ErrorAction SilentlyContinue -PathType Leaf ) )
    {
        Throw "Procmon not found at $procmon"
    }

    [string]$procmonArguments = "/Quiet /AcceptEula /BackingFile `"$backingFile`" /Minimized /Runtime $runTime /Nofilter"
<#
    ## A bug in procmon means that if the pml file produced is empty then the procmon run to convert to csv will get a  "the file is corrupt and cannot be opened" error which hangs it so we only filter then, not when capturing
    if( $pmcFile )
    {
        $procmonArguments += " /LoadConfig `"$pmcFile`""
    }
#>
    Write-Output "$(Get-Date -Format G): running procmon to capture data for $runtime seconds ..."
        
    $monitorProcess = Start-Process -FilePath $procmon -ArgumentList $procmonArguments -PassThru

    if( ! $monitorProcess )
    {
        Throw "Failed to run procmon `"$procmon`""
    }
    
    ## Start a separate process to monitor the procmon process and terminate it if it does not exit when this script exits which is probably because the SBA timed out due to procmon hanging
    [string]$extraBackingFiles = $null
    if( ! $passedArguments[ 'backingFile' ] )
    {
        ## tidy up any *-1.pml, etc
        $extraBackingFiles = ",""{0}""" -f ( Join-Path $env:temp "controlup.sba.$($pid)-*.pml" )
    }
    [string]$watcherAppArguments = "-ExecutionPolicy Bypass -NonInteractive -WindowStyle Hidden -Command ""&{ Wait-Process -Id $pid ; Stop-Process -name 'procmon','procmon64' -Force -PassThru ; Start-Sleep -seconds 10 ; Remove-Item -Path ""$backingFile"",""$pmcFile""$extraBackingFiles -Force "
    if( $downloadedProcmon )
    {
        $watcherAppArguments += "; Remove-Item -Force ""$procmon"""
    }
    $watcherAppArguments += "}"""

    $watcherProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList $watcherAppArguments -PassThru
        
    if( ! $watcherProcess )
    {
        Write-Warning "Failed to launch watcher script to ensure procmon is stopped if this script times out"
    }

    Wait-Process -Id $monitorProcess.Id

    Write-Output "$(Get-Date -Format G): running procmon to convert capture file ..."
        
    $procmonArguments = "/Quiet /AcceptEula /OpenLog `"$backingFile`" /Minimized /SaveAs `"$csvTrace`" /SaveApplyFilter"
    if( $pmcFile )
    {
        $procmonArguments += " /LoadConfig `"$pmcFile`""
    }
    $convertProcess = Start-Process -FilePath $procmon -ArgumentList $procmonArguments -PassThru -Wait

    if( ! $convertProcess )
    {
        Throw "Failed to run procmon conversion `"$procmon`""
    }
    if( ! ( Test-Path -Path $csvTrace -ErrorAction SilentlyContinue ) )
    {
        Throw "Procmon failed to save csv trace file to `"$csvTrace`""
    }

    Write-Verbose "$(Get-Date -Format G): reading events ..."

    $traceEvents = @( Import-Csv -Path $csvTrace )
    Write-Verbose "Got $($traceEvents.Count) trace events from `"$csvTrace`""
    if( $traceEvents -and $traceEvents.Count )
    {
        # Calculate sum of durations by "operation" field
        $durationsbyoperation = @( ($traceEvents | Group-Object -Property Operation -AsHashTable).GetEnumerator()|ForEach-Object `
        {
            [pscustomobject][ordered]@{ 'Operation' = $_.Key ; 'TotalSec' = [math]::Round( ($_.Value | Measure -Property Duration -Sum|Select -ExpandProperty Sum ) , 3 ) } 
        })

        # Display the operation categories that took the longest
        if ($durationsbyoperation -and $durationsbyoperation.Count )
        {
            Write-Output "`r`nTop 10 operation categories by duration"
            Write-Output "======================================="
            $durationsbyoperation | sort TotalSec -Descending | select -first 10 | Format-Table Operation, TotalSec -Wrap
        }

        # Display individual operations that took the longest
        $top10ops = @( $traceEvents | sort Duration -Descending | select -first 10 )
        if ($top10ops -and $top10ops.COunt )
        {
            Write-Output "`r`nTop 10 individual operations by duration"
            Write-Output "========================================"
            $top10ops | Format-Table Operation,Path,Category,@{n='Duration';e={[math]::Round( $_.Duration , 3 )}} -Wrap
        }
        
        $top10paths = @( $traceEvents | Where-Object { $_.Path -match '^[a-z]:\\' } | Group-Object -Property Path | Sort Count -Descending | Select Name, Count -First 10 )
        if( $top10paths -and $top10paths.Count )
        {
            Write-Output "`r`nTop 10 paths by access count"
            Write-Output "============================="
            $top10paths | Format-Table -Wrap
        }
    }
    else
    {
        Write-Warning "No events captured"
    }
}
catch
{
    Write-Error "Line: $($_.InvocationInfo.ScriptLineNumber) : $_"
}
finally
{
    if( $watcherProcess )
    {
        $watcherProcess | Stop-Process -Force
    }
    if( $monitorProcess -or $convertProcess )
    {
        Stop-Process -name 'procmon*' -Force
    }
    Remove-Item -Path $backingFile -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $csvTrace -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $pmcFile -Force -ErrorAction SilentlyContinue
    if( ! $passedArguments[ 'backingFile' ] )
    {
        ## tidy up any *-1.pml, etc
        Remove-Item -Path ( Join-Path $env:temp "controlup.sba.$($pid)-*.pml" ) -Force -ErrorAction SilentlyContinue
    }
    if( $downloadedProcmon )
    {
        Remove-Item -Force -Path $procmon -ErrorAction SilentlyContinue
    }
}
