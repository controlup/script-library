#require -version 3

<#
.SYNOPSIS 
    Get date/time and location of where FSlogix disks mounted via the .metadata file

.PARAMETER shares
    Comma separated list of shares to examine. If not specified, the registry will be checked for FSlogix settings
    
.PARAMETER boottime
    Include the last boot time of the machines where the disk is mounted. Could be slow if machines are not booted and/or if WMI/CIM access not available

.PARAMETER filePattern
    The file pattern to search to find the meta data files. It is not recommended to change this

.PARAMETER operationTimeoutSeconds
    How long in seconds to allow the WMI/CIM operation to retrieve the last boot time to run

.NOTES
    Modification History:

    2022/10/06  @guyrleech  Initial Release
    2022/10/31  @guyrleech  Exclude CORRUPT_* files
#>

[CmdletBinding()]

Param
(
    [ValidateSet('yes','no')]
    [string]$bootTime = 'yes' ,
    [string[]]$shares = @( '*' ) ,  # placeholder so CU can pass * to mean get from registry
    [string]$filePattern = '*.metadata' ,
    [int]$operationTimeoutSeconds = 30
)

#region ControlUp_Standards
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
    ## not the ennd of the world but we don't want this to stop the script from completing
}

#endregion ControlUp_Standards

if( $null -eq $shares -or $shares.Count -eq 0 -or $shares[0].Length -eq 1 )
{
    [string]$vhdLocationsValueName = 'VHDLocations'
    $shares = @( ForEach( $key in @( 'HKLM:\Software\FSlogix\Profiles' , 'HKLM:\SOFTWARE\Policies\FSLogix\ODFC' ) )
    {
        if( $fslogixprofileKey = Get-ItemProperty -Path $key -Name $vhdLocationsValueName -ErrorAction SilentlyContinue )
        {
            $fslogixprofileKey.VHDLocations
        }
        else
        {
            Write-Warning "No FSlogix profile root passed and no '$vhdLocationsValueName' value in FSlogix registry key $key"
        }
    })
}
elseif( $shares.Count -eq 1 -and $shares[0].IndexOf( ',' ) -ge 0 ) ## array can be flattened if called outside of PS, eg scheduled task
{
    $shares = @( $shares -split ',' )
}

if( $null -eq $shares -or $shares.Count -eq 0 )
{
    Throw "No shares passed and no FSlogix registry configuration present"
}

[double]$totalVHDSize = 0
[int]$totalVHDs = 0
[string]$baseFileRegex = [regex]::Escape( ( $filePattern -replace '^\*' ) )
[hashtable]$machineOsInfo = @{}

[array]$results = @( ForEach( $share in $shares )
{
    Write-Verbose -Message "Enumerating share $share"
    Get-ChildItem -Path $share -Include $filePattern -Recurse -Force | Where-Object Name -NotMatch '^CORRUPT_' | ForEach-Object `
    {
        $file = $_
        [string]$mountedOn = 'N/A'
        [byte[]]$bytes = @()
        ## can't use Get-Content as -Encoding Byte not available in PS 7 so this way keeps code easier and file should only be small
        $bytes = [System.IO.File]::ReadAllBytes( $file.FullName )

        [int]$lastCharacter = -1

        if( $null -ne $bytes -and $bytes.Count -gt 0 )
        {
            ## get first non-control character (it is actually unicode) and build string from there until 00 00 terminator
            <#
            00000000   01 00 00 00 01 00 00 00 01 00 00 00 47 00 4C 00  ............G.L.
            00000010   57 00 31 00 30 00 43 00 54 00 58 00 4D 00 43 00  W.1.0.C.T.X.M.C.
            00000020   53 00 50 00 30 00 31 00 2E 00 67 00 75 00 79 00  S.P.0.1...g.u.y.
            00000030   72 00 6C 00 65 00 65 00 63 00 68 00 2E 00 6C 00  r.l.e.e.c.h...l.
            00000040   6F 00 63 00 61 00 6C 00 00 00 00 00 00 00 00 00  o.c.a.l.........
            00000050   00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00  ................
            #>
            [int]$startIndex = -1
            [int]$endIndex = -1
            For( [int]$index = 0 ; $index -lt $bytes.Count ; $index++ )
            {
                if( $startIndex -lt 0 -and $bytes[ $index ] -ge 32 -and $index -gt 0 -and $lastCharacter -eq 0x0 ) ## ascii space with preceding zero byte as Unicode
                {
                    $startIndex = $index - 1
                }
                elseif( $startIndex -ge 0 -and $bytes[ $index ] -eq 0x0 -and $lastCharacter -eq 0x0 )
                {
                    $endIndex = $index
                    break
                }
                $lastCharacter = $bytes[ $index ]
            }

            if( $startIndex -ge 0 )
            {
                if( $endIndex -gt $startIndex )
                {
                    $mountedOn = [System.Text.Encoding]::BigEndianUnicode.GetString( $bytes , $startindex , $endIndex - $startIndex - 1 )
                }
                else
                {
                    Write-Warning -Message "Unable to find end of string in `"$($file.FullName)`""
                }
            }
            else
            {
                Write-Warning -Message "Unable to find start of string in `"$($file.FullName)`""
            }
        }
        elseif( $fileError )
        {
            Write-Warning -Message "Error reading `"$($file.FullName)`" - $fileError"
        }

        [string]$username = 'N/A'

        $parentPath = $null
        if( $parentPath = (Split-Path -Path (Split-Path -path $file -parent) -Leaf) )
        {
            if( $parentPath -match '(S-\d-\d-\d+-\d+-\d+-\d+-\d+)' )
            {
                [string]$sid = $Matches[ 1 ]
                if( $resolvedSid = ([System.Security.Principal.SecurityIdentifier]( $sid )).Translate([System.Security.Principal.NTAccount]).Value )
                {
                    $username = $resolvedSid
                }
                else
                {
                    $username = $sid
                }
            }
            else ## TODO fallback looking for username in file or path
            {
            }
        }

        $baseFile = $null    
        $baseFile = Get-ItemProperty -Path ($file.FullName -replace "$baseFileRegex`$") -ErrorAction SilentlyContinue

        $result = [pscustomobject]@{
            'Disk' = $baseFile | Select-Object -ExpandProperty Name
            'Location' = $baseFile | Select-Object -ExpandProperty DirectoryName
            'Username' = $username
            'Mounted'  = $file.CreationTime
            'Machine'  = $mountedOn
            'Disk Created' = $baseFile | Select-Object -ExpandProperty CreationTime
            'Disk Last Modified' = $baseFile | Select-Object -ExpandProperty LastWriteTime
            'Disk Last Accessed' = $baseFile | Select-Object -ExpandProperty LastAccessTime
            'Disk Size (MB)' = [math]::Round( ( $baseFile | Select-Object -ExpandProperty Length ) / 1MB , 1 )
        }

        if( $bootTime -ieq 'yes' -and $mountedOn -ne 'N/A' )
        {
            if( $null -eq ( $osinfo =  $machineOsInfo[ $mountedOn ] ) )
            {
                Write-Verbose -Message "$(Get-Date -Format G): getting boot time of $mountedOn"
                $osinfo = Get-CimInstance -ClassName Win32_OperatingSystem -OperationTimeoutSec $operationTimeoutSeconds -ComputerName $mountedOn -ErrorAction SilentlyContinue
                if( $osinfo )
                {
                    $machineOsInfo.Add( $mountedOn , $osinfo )
                }
                else
                {
                    ## add even if $null so we don't try a non-booted/bad machine again which will delay script further
                    $machineOsInfo.Add( $mountedOn , $false )
                }              
            }

            Add-Member -InputObject $result -MemberType NoteProperty -Name 'Boot Time' -Value ($osinfo | Select-Object -ExpandProperty LastBootUpTime -ErrorAction SilentlyContinue)
        }

        $result
    }
})

## Remove Format-Table to use outside of ControlUp
$results | Format-Table -AutoSize

