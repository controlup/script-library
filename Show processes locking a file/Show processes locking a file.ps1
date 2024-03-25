<#
    Use SysInternals handle.exe tool to find who has a certain file open

    @guyrleech 2018
#>

[string]$fileName = $args[0]
[string]$pathToHandle = $null
[bool]$downloadedHandle = $false
[int]$outputWidth = 400

if( $args.Count -ge 2 -and $args[1] )
{
    $pathToHandle = $args[1]
    if( ! ( Test-Path -Path $pathToHandle -ErrorAction SilentlyContinue ) )
    {
        Throw "Unable to find handle.exe at `"$pathToHandle`""
    }
}
else
{
    $pathToHandle = Join-Path $env:temp 'handle.controlup.exe'
    Write-Verbose "Downloading handle.exe to `"$pathToHandle`" ..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    (New-Object System.Net.WebClient).DownloadFile( 'https://live.sysinternals.com/handle.exe' , $pathToHandle )
    if( ! ( Test-Path $pathToHandle -ErrorAction SilentlyContinue -PathType Leaf ) )
    {
        Throw "Failed to download handle to `"$pathToHandle`""
    }
    Unblock-File -Path $pathToHandle
    $downloadedHandle = $true
    $signing = Get-AuthenticodeSignature -FilePath $pathToHandle -ErrorAction SilentlyContinue
    if( ! $signing )
    {
        Throw "Could not get signing information from `"$pathToHandle`""
    }
    if( ! $signing.Status -ne 'Valid' )
    {
        Throw "Certificate status for `"$pathToHandle`" is $($signing.Status), not `"Valid`""
    }
    if( $signing.SignerCertificate.Subject -notmatch '^CN=Microsoft Corporation,' )
    {
        Throw "`"$pathToHandle`" is not signed by Microsoft Corporation, found $($signing.SignerCertificate.Subject)"
    }
}

[string]$escapedFileName = [regex]::Escape( $filename )
## an apparent bug in handle.exe where it fails to match on a full path means we get the whole output and grab our lines from them
[hashtable]$processLine = $null
[array]$results = @( & $pathToHandle -accepteula -nobanner -u | ForEach-Object `
{
    ## username could be "\<unable to open process>"
    ## SCService64.exe pid: 3468 NT AUTHORITY\NETWORK SERVICE
    if( $_ -match "^(?<Process>.*)\s+pid:\s+(?<Pid>\d+)\s(?<Username>[a-z0-9_\s<>\-\.]*\\[a-z0-9_\s<>\-\.]+)" )
    {
        $processLine = $Matches.Clone()
    }
    ##  420: File  (R--)   C:\Windows\assembly\pubpol3.dat
    elseif( $_ -match "(?<Handle>[0-9A-F]+):\s*File\s+\((?<Flags>[^\)]*)\)\s+(?<File>.*$escapedFileName.*)$" )
    {
        $fileProperties = Get-ItemProperty -Path $Matches[ 'File' ] -ErrorAction SilentlyContinue
        [pscustomobject][ordered]@{
            'File' = $Matches[ 'File' ]
            'File Owner' = ( Get-Acl -Path $Matches[ 'File' ] -ErrorAction SilentlyContinue | Select -ExpandProperty Owner )
            'Last Modified' = $( if( $fileProperties ) { (Get-Date $fileProperties.LastWriteTime -Format G) } )
            'Process' = $processLine[ 'Process' ]
            'Flags' = $Matches[ 'Flags' ]
            'Pid' = $processLine[ 'Pid' ]
            'User' = $(if( $processLine[ 'username' ] -notlike '*unable to open process*' ) { $processLine[ 'username' ] } )
            'Handle' = $Matches[ 'Handle' ] }
    }
})

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

"Found $($results.Count) open instances of $fileName"

$results | Sort -Property 'File','Process' | Select-Object -Property * -ExcludeProperty Handle | Format-Table -AutoSize

if( $downloadedHandle )
{
    Remove-Item -Path $pathToHandle -Force
}

