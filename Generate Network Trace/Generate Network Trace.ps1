#requires -RunAsAdministrator

<#
.SYNOPSIS
    Use PowerShell to capture a network trace and look for TLS handshakes in the captured data

.NOTES
    https://devblogs.microsoft.com/scripting/packet-sniffing-with-powershell-getting-started/
    https://pcapng.com/

    Modification History:

    2024/09/30  Guy Leech  Script born
    2024/10/02  Guy Leech  Added filter parameters
    2024/10/03  Guy Leech  Added array splitting. Added code to get latest download. Changed code to keep etl if converter not downloaded
    2024/10/04  Guy Leech  TLS1.2 code added. Improved mechanism for getting latest download URL
    2024/10/09  Guy Leech  Workaround for trace file passed as empty string. Fixed bug where new session creation fails. Added -outputfile parameter
#>

[CmdletBinding()]

Param
(
    [int]$durationSeconds = 60 ,
    [ValidateSet('Yes','No','True','False')]
    [string]$overWrite = 'No' ,
    [string[]]$addresses ,
    [string[]]$protocols ,
    [string[]]$etherTypes ,
    [string]$traceName = 'ControlUp_trace' ,
    [string]$traceFile ,
    [string]$outputFile ,
    [int]$maxFileSizeMB = 512 ,
    [string]$fallbackURL = 'https://github.com/microsoft/etl2pcapng/releases/download/v1.11.0/etl2pcapng.exe' ,
    [string]$releasesPage = 'https://api.github.com/repos/microsoft/etl2pcapng/releases' ,
    [int]$truncationLength = 1500 ,
    [string]$expectedSignerCertificateSubjectRegex = '^CN=Microsoft Corporation,'
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 250
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
    ## not a showstopper, will just may not be wide enough to stop output wrapping
}

Function Resolve-RecurisveDNSName
{
    [CmdletBinding()]
    Param
    (
        [string]$Name
    )
    $resolved = $null
    $resolved = @( Resolve-DnsName -Name $Name -PipelineVariable resolvedItem | ForEach-Object -Process `
    {
        if( $resolvedItem.GetType().Name -ieq 'DnsRecord_PTR' )
        {
            Resolve-RecurisveDNSName -Name $resolvedItem.NameHost ## TODO could this go infintely recursive such that we need to record hosts already resolved?
        }
        else
        {
            $resolvedItem.Address
        }
    })
    $resolved | Sort-Object -Unique
}

$( "Script Parameters:" ; $PSBoundParameters.GetEnumerator())| Write-Verbose

## argument passing can flatten arrays so we need to reconstitute them
if( $null -ne $addresses )
{
    $addresses = @( $addresses -split ',' )
}
if( $null -ne $protocols )
{
    $protocols = @( $protocols -split ',' )
}
if( $null -ne $etherTypes )
{
    $etherTypes = @( $etherTypes -split ',' )
}
[string]$etl2pcapConverterPath = Join-Path -Path $env:Temp -ChildPath (Split-Path -Path $fallbackURL -Leaf)

Import-Module -Name NetEventPacketCapture -Verbose:$false -Debug:$false

## process early so if any errors we don't have stuff to undo/delete

$captureProvider = $null
[hashtable]$captureParameters = @{
    SessionName = $traceName
    TruncationLength = $truncationLength
}
if( $null -ne $addresses -and $addresses.Count -gt 0 )
{
    Import-Module -Name DnsClient -Verbose:$false -Debug:$false
    [string[]]$ipAddressList = @( ForEach( $address in $addresses) ## parameter passing can cause empty elements
    {
        if( -Not[string]::IsNullOrEmpty( $address ))
        {
            ## filter only takes IP addresses so look up DNS names if not an IP address
            if( -Not ( $address -as [ipaddress]))
            {
                Resolve-RecurisveDNSName -Name $address
            }
            else
            {
                $address
            }
        }
    })
    Write-Verbose -Message "IPAddresses: $($ipAddressList -join ' , ')"
    $captureParameters.Add( 'IPAddresses' , $ipAddressList )
}
if( $null -ne $etherTypes -and $ethertypes.Count -gt 0)
{
    [uint16[]]$etherTypesArray = @( ForEach( $etherType in $etherTypes)
    {
        switch ($etherType)
        {
            'IPv4' { 0x0800 }
            'IPv6' { 0x86DD }
            'ARP' { 0x0806 }
            'AppleTalk' { 0x0801 }
            'RARP' { 0x0805 }
            'MPLS unicast' { 0x8847 }
            'MPLS multicast' { 0x8848 }
            'VLAN' { 0x8100 }
            'LLDP' { 0x88CC }
            'MACsec' { 0x88E1 }
            'EAPOL' { 0x8915 }
            'PPPoE Discovery Stage' { 0x8863 }
            'PPPoE Session Stage' { 0x8864 }
            '' {}
            default { Throw "Unknown etherType $etherType" }
        }
    } )
    $captureParameters.Add( 'EtherType' , $etherTypesArray )
}
if( -Not [string]::IsNullOrEmpty( $protocols ))
{
    [byte[]]$protocolByteArray = @( ForEach( $protocol in $protocols )
    {
        switch( $protocol )
        {
            'TCP' { 6 }
            'UDP' { 17 }
            'ICMP' { 1 }
            'IGMP' { 2 }
            'IPv6 encapsulation' { 41 }
            'GRE' { 47 }
            'ESP' { 50 }
            'AH' { 51 }
            'ICMPv6' { 58 }
            'OSPF' { 89 }
            '' {}
            <#
            1: ICMP (Internet Control Message Protocol)
            2: IGMP (Internet Group Management Protocol)
            6: TCP (Transmission Control Protocol)
            17: UDP (User Datagram Protocol)
            41: IPv6 encapsulation (used in IPv6-in-IPv4 tunneling)
            47: GRE (Generic Routing Encapsulation)
            50: ESP (Encapsulating Security Payload)
            51: AH (Authentication Header)
            58: ICMPv6 (Internet Control Message Protocol for IPv6)
            89: OSPF (Open Shortest Path First)
            #>
            default { Throw "Unknown protocol $protocol" }
        }
    } )
    $captureParameters.Add( 'IPProtocols' , $protocolByteArray )
}

if( [string]::IsNullOrEmpty( $traceFile ) )
{
    $traceFile = Join-Path -Path $env:temp -ChildPath "ControlUpTrace-$([datetime]::now.Ticks)$pid.etl"
}

if( $fileProperties = Get-ItemProperty -Path $traceFile -ErrorAction SilentlyContinue)
{
    if( $fileProperties.Length -eq 0 )
    {
        Remove-Item -Path $newSession.LocalFilePath
    }
    elseif( $overWrite -notmatch 'Yes|True' )
    {
        Throw "Capture file `"$($fileProperties.FullName) already exists, is $([math]::Round( $fileProperties.Length / 1MB , 2))MB, created $($fileProperties.CreationTime.ToString('G'))"
    }
}

## try and get the latest download URL
[string]$downloadURL = $fallbackURL
try
{
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    [string]$tryURL = ((Invoke-RestMethod -URI $releasesPage )|Select-Object -Property @{n='version';e={($_.tag_name -replace '^v') -as [version]}},assets|Sort-Object version -Descending|Select-Object -first 1 -ExpandProperty assets|Where-Object browser_download_url -match '\.exe$'|Select-Object -ExpandProperty browser_download_url)
    if( -Not[string]::IsNullOrEmpty( $tryURL ))
    {
        if( $tryURL -ine $fallbackURL )
        {
            Write-Verbose -Message "Trying $tryURL"
            (New-Object -TypeName System.Net.WebClient).Downloadfile( $tryURL , $etl2pcapConverterPath )
            if( $? )
            {
                Write-Verbose -Message "Downloaded newer binary from $tryURL"
                $downloadURL = $null ## so that later it won't download again
            }
        }
        else
        {
            Write-Verbose -Message "Not trying $tryURL as it is the same as the fallback URL"
        }
        ## else same URL as we already have so will use that
    }
}
catch
{
    Write-Verbose -Message "Dynamic download problem : $_"
    ## fallback to the hard coded URL
}

if( -Not [string]::IsNullOrEmpty($downloadURL))
{
    Write-Verbose -Message "$([datetime]::Now.ToString('G')): downloading from $downloadURL to $etl2pcapConverterPath"
    
    (New-Object -TypeName System.Net.WebClient).Downloadfile( $downloadURL , $etl2pcapConverterPath )
}

if( -Not $? -or $null -eq ($exeProperties = Get-ItemProperty -Path $etl2pcapConverterPath) -or $exeProperties.Length -lt 10KB)
{
    Write-Warning "Problem downloading from $downloadURL to $etl2pcapConverterPath"
}
else
{
    
    $signing = $null
    $signing = Get-AuthenticodeSignature -FilePath $etl2pcapConverterPath
    if( $null -eq $signing -or $signing.Status -ne 'Valid' )
    {
        Throw "No valid Authenticode signature found on $etl2pcapConverterPath"
    }
    if( $signing.SignerCertificate.Subject -notmatch $expectedSignerCertificateSubjectRegex )
    {
        Throw "Unexpected signer certificate subject $($signing.SignerCertificate.Subject) on $etl2pcapConverterPath"
    }
}

$newSessionError = $null
$newSession = $null
$newSession = New-NetEventSession -Name $traceName -LocalFilePath $traceFile -MaxFileSize $maxFileSizeMB -CaptureMode SaveToFile -ErrorAction SilentlyContinue -ErrorVariable newSessionError
if( -Not $? -or $null -eq $newSession )
{
    [array]$existingTraces = @( Get-NetEventSession )
    if( $null -ne $existingTraces -and $existingTraces.Count -gt 0 )
    {
        if( $existingTraces.Name -contains $traceName )
        {
            if( $overWrite -match 'Yes|True' )
            {
                Remove-NetEventSession -Name $traceName -Confirm:$false
                $newSession = New-NetEventSession -Name $traceName -LocalFilePath $traceFile -MaxFileSize $maxFileSizeMB -CaptureMode SaveToFile -ErrorAction SilentlyContinue -ErrorVariable newSessionError
                if( -Not $? -or $null -eq $newSession )
                {
                    Throw "Tracing session $traceName already existed and was removed but still failed to create new one: $newsSessionError"
                }
            }
            else
            {
                Throw "There are already $($existingTraces.Count) tracing sessions ($($existingTraces.Name -join ' , ')) - please remove them and rerun the script or use -overwrite"
            }
        }
        else
        {
            Throw "There are already $($existingTraces.Count) tracing sessions ($($existingTraces.Name -join ' , ')) - please remove them and rerun the script" ## -overwrite won't work as not traces we knowingly created
        }
    }
    else
    {
        Throw "Failed to create tracing session - $newsessionError"
    }
}

Write-Verbose "New session: $newSession"

try
{
    $( "Capture Parameters:" ; $captureParameters.GetEnumerator()|Format-Table | Out-String) | Write-Verbose
    $captureProvider = Add-NetEventPacketCaptureProvider @captureParameters
    if( -Not $? -or $null -eq $captureProvider )
    {
        Throw "Failed to add packet capture provider to session $traceName"
    }
    
    if( [string]::IsNullOrEmpty( $outputFile ) )
    {
        $outputfile = $newSession.LocalFilePath -replace '\.\w+$' , '.pcapng'
    }

    if( ( $fileProperties = Get-ItemProperty -Path $outputFile -ErrorAction SilentlyContinue) -and $fileProperties.Length -gt 0 )
    {
        ## see if open, eg Wireshark has previous trace open
        try
        {
            $fileStream = [System.IO.File]::Open($outputFile, 'Open', 'ReadWrite', 'None')
            $fileStream.Close()
            $fileStream.Dispose()
        }
        catch
        {
            Throw "Failed to open $outputFile - is it open in another application?"
        }
        if( $overWrite -notmatch 'Yes|True' )
        {
            Throw "Capture file `"$($outputFile)`" already exists, is $([math]::Round( $fileProperties.Length / 1MB , 2))MB, created $($fileProperties.CreationTime.ToString('G'))"
        }
        else
        {
            Remove-Item -Path $outputFile
        }
    }

    if( -Not ( Start-NetEventSession -Name $traceName -PassThru ))
    {
        Throw "Failed to start trace"
    }
    Write-Verbose -Message "$([datetime]::Now.ToString('G')): sleeping for $durationSeconds seconds"

    Start-Sleep -Seconds $durationSeconds

    Write-Verbose -Message "$([datetime]::Now.ToString('G')): back from sleep"

    $stoppedSession = $null
    $stoppedSession = Stop-NetEventSession -Name $traceName -PassThru
    if( $null -eq $stoppedSession )
    {
        Write-Warning "Possible problem stopping trace"
    }

    ## might have been unable to download the converter in which case we don't delete the trace file
    if( -Not (Test-Path -Path $etl2pcapConverterPath -PathType Leaf))
    {
        Write-Output -InputObject "Raw trace file `"$($newsession.LocalFilePath)`" created, size is $([math]::Round( $fileProperties.Length / 1MB , 2))MB"
        Write-Output -InputObject "Convert to pcapng file with etl2pcapng.exe utility from https://github.com/microsoft/etl2pcapng"
    }
    else
    {      
        ## not using start-process so we can get error output if required
        $processInfo = New-Object -TypeName System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $etl2pcapConverterPath
        $processInfo.Arguments = "`"$($newSession.LocalFilePath)`" `"$outputFile`""
        $processInfo.RedirectStandardError = $true
        $processInfo.RedirectStandardOutput = $true
        $processInfo.UseShellExecute = $false
        $processInfo.WindowStyle = 'Hidden'
        $processInfo.CreateNoWindow = $true
        $process = New-Object -TypeName System.Diagnostics.Process
        $process.StartInfo = $processInfo
        if( $process.Start() )
        {
            $process.WaitForExit()
            
                if( $process.ExitCode -ne 0 )
                {
                    Write-Warning -Message "$($processInfo.FileName) exited with status $($process.ExitCode)"
                    $process.StandardError.ReadToEnd()|Write-Warning
                    $process.StandardOutput.ReadToEnd()|Write-Warning
                }
        }
        else
        {
            Throw "Failed to run `"$($processInfo.FileName)`""
        }
        if( -Not ( $fileProperties = Get-ItemProperty -Path $outputFile ))
        {
            Throw "$($process.Path) failed to create capture file"
        }
        elseif( $fileProperties.Length -eq 0 )
        {
            Throw "$($process.Path) created an empty capture file"
        }

        Write-Output -InputObject "Capture file `"$($outputFile)`" created, size is $([math]::Round( $fileProperties.Length / 1MB , 2))MB"
        
        Remove-Item -Path $traceFile
        $traceFile = $null
    }
}
catch
{
    throw
}
finally
{
    if( ( $currentSession = Get-NetEventSession  -name $traceName -ErrorAction SilentlyContinue ) -and $currentSession.SessionStatus -eq 'Running' )
    {
        $stopError = $null
        $stoppedSession = Stop-NetEventSession -Name $traceName -PassThru -ErrorAction SilentlyContinue -ErrorVariable stopError
        if( $null -eq $stoppedSession)
        {
            Write-Warning "Problem stopping trace $traceName : $stopError"
        }
    }
    
    Remove-NetEventSession -Name $traceName

    if( $null -ne $etl2pcapConverterPath -and (Test-Path -Path $etl2pcapConverterPath -PathType Leaf))
    {
        ## get errror if we try to delete too soon
        [datetime]$endTime = [datetime]::Now.AddSeconds( 15 )  ## yuck, hard coding!
        do
        {
            Remove-Item -Path $etl2pcapConverterPath -ErrorAction SilentlyContinue
            if( -Not $?)
            {
                Write-Verbose -Message "$([datetime]::Now.ToString('G')): waiting for $etl2pcapConverterPath to delete"
                Start-Sleep -Seconds 1 ## yuck, hard coding!
            }
            ## else delete worked ok
        } while( (Test-Path -Path $etl2pcapConverterPath -PathType Leaf) -and [datetime]::Now -lt $endTime )
    }
}
