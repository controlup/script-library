<#
.SYNOPSIS

Backup specified DNS zones

.DETAILS

Uses dnscmd.exe /ZoneExport to backup files to a subfolder created in system32\dns

.PARAMETER zoneName

The full name or a regular expression matching the zone or zones that are to be operated upon

.PARAMETER overWrite

If set to "yes" will overwrite existing backup files otherwise the backup for that zone will not happen

.PARAMETER subFolder

The name of the subfolder to create the backup files in

.CONTEXT

Computer (must be a DNS server)

.MODIFICATION_HISTORY:

@guyrleech 27/09/19

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='Zone name to backup/restore')]
    [string]$zoneName ,
    [Parameter(Mandatory=$false)]
    [ValidateSet('Yes','No')]
    [string]$overWrite = 'No' ,
    [Parameter(Mandatory=$false)]
    [string]$subFolder = 'Controlup' ,
    [Parameter(Mandatory=$false)]
    [ValidateSet('Yes','No')]
    [string]$restore = 'No'
)

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'

if( ! ( Get-Command -Name dnscmd.exe ) )
{
    Throw 'Unable to locate dnscmd.exe'
}

## Need RSAT
##Import-Module -Name DnsServer -Verbose:$false

[string]$folder = Join-Path -Path (Join-Path -Path ([Environment]::GetFolderPath('System')) -ChildPath 'dns') -ChildPath $subfolder

if( ! ( Test-Path -Path $folder -ErrorAction SilentlyContinue ) )
{
    if( $restore -eq 'Yes' )
    {
        Throw "Folder `"$folder`" does not exist so cannot restore"
    }
    else
    {
        $newFolder = New-Item -Path $folder -ItemType Directory
        if( ! $newFolder )
        {
            Throw "Failed to create backup folder `"$folder`""
        }
    }
}

[int]$counter = 0

if( $restore -eq 'No' )
{
    ##[array]$zones = @( Get-DnsServerZone | Where-Object { $_.ZoneType -eq 'Primary' -and ! $_.IsAutoCreated -and $_.ZoneName -match $zoneName } )
    [bool]$seenHeader = $false
    [array]$zones = @( ForEach( $line in (dnscmd.exe /EnumZones /Primary) )
    {
        [array]$fields = @( $line -split '\s+' )
        if( $fields.Count -ge 6 )
        {
            if( $seenHeader -and $fields[ 1 ] -match $zoneName)
            {
                $result = New-Object -TypeName PSCustomObject
                Add-Member -InputObject $result -MemberType NoteProperty -Name ZoneName -Value $fields[1]
                $result
            }
            else
            {
                $seenHeader = $true
            }
        }
    })

    Write-Verbose -Message "Got $($zones.Count) primary zones"

    if( ! $zones.Count )
    {
        Throw "No primary zones found matching `"$zoneName`""
    }

    ForEach( $zone in $zones )
    {
        [string]$backupFile = Join-Path -Path $folder -ChildPath ( $zone.ZoneName + '.backup' )
        [bool]$carryOn = $true

        $backupFileDetails = Get-ItemProperty -Path $backupFile -ErrorAction SilentlyContinue
        
        if( $backupFileDetails )
        {
            if( $overWrite -ne 'Yes' )
            {
                Write-Warning -Message "Unable to backup zone $($zone.zoneName) as backup file already exists from $(Get-Date -Date $backupFileDetails.LastWriteTime -Format G)"
                $carryOn = $false
            }
            else
            {
                Write-Warning -Message "Overwriting backup file for zone $($zone.zoneName) from $(Get-Date -Date $backupFileDetails.LastWriteTime -Format G)"
                Remove-Item -Path $backupFile -Force
                if( ! $? )
                {
                    Write-Warning -Message "Failed to delete backup file `"$backupFile`""
                    $carryOn = $false
                }
            }
        }

        if( $carryOn )
        {
            $result = Start-Process -FilePath 'dnscmd.exe' -ArgumentList "$env:COMPUTERNAME /ZoneExport $($zone.ZoneName) `"$(Join-Path -Path $subfolder -ChildPath ($zone.ZoneName + '.backup' ))`"" -PassThru -Wait -WindowStyle Hidden
            if( $result )
            {
                if( $result.ExitCode )
                {
                    Write-Warning -Message "Error code $($result.ExitCode) returned from dnscmd.exe for zone $($zone.ZoneName)"
                }
				Write-Output -InputObject "Zone $($zone.zonename) backed up to $backupFile"
            }
            else
            {
                Write-Warning -Message "Failed to run dnscmd.exe for zone $($zone.ZoneName)"
            }
        }
    }
}
else ## restore
{
    Throw 'Restoration is not yet implemented - use DNS console or dnscmd.exe /ZoneAdd'
}
