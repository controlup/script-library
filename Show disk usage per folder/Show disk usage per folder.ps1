#Requires -version 3.0
<#
    Calculate folder sizes of given folder and display largest first which have been modified within a given number of days

    @guyrleech, 2018
#>

[string]$startingFolder = ( Join-Path $env:SystemDrive '\' )
[int]$depth = 1
[int]$showFirst = 10
[int]$lastWrittenDays = 1000
[int]$outputWidth = 400

$VerbosePreference = 'SilentlyContinue'

if( $args.Count )
{
    if( ! [string]::IsNullOrEmpty( $args[0] ) )
    {
        $startingFolder = $args[0]
    }
    if( $args.Count -ge 2 -and ! [string]::IsNullOrEmpty( $args[1] ) )
    {
        $depth = $args[1]
    }
    if( $args.Count -ge 3 -and ! [string]::IsNullOrEmpty( $args[2] ) )
    {
        $showFirst = $args[2]
    }
    if( $args.Count -ge 4 -and ! [string]::IsNullOrEmpty( $args[3] ) )
    {
        $lastWrittenDays = $args[3]
    }
}

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

## Do it this way so we can exclude junction point folders and sym linked files and stop at a specific level
$items = Get-Item $startingFolder -Force
[int]$level = 1

[datetime]$lastWritten = (Get-Date).AddDays( -$lastWrittenDays )

[array]$allItems = while( $items -and $level -le $depth )
{
    $newitems = $items | Get-ChildItem -Attributes !ReparsePoint+!SparseFile+Directory -ErrorAction SilentlyContinue -Force
    $items = $newitems | Where-Object -Property Attributes -Like *Directory*
    $newItems | select *,@{n='Level';e={$level}}
    $level++
}

[hashtable]$childFolderSizes = @{}

For( [int]$thisLevel = $depth ; $thisLevel -gt 0 ; $thisLevel-- )
{
    [long]$levelSize = 0

    $allItems | Where-Object { $_.Level -eq $thisLevel } | ForEach-Object `
    {
        [long]$totalSize = 0
        [string]$thisFolder = $_.FullName
        Get-ChildItem -Path $thisFolder -Attributes !ReparsePoint+!SparseFile+Directory -ErrorAction SilentlyContinue -Force | ForEach-Object `
        {
            Write-Verbose "$thisLevel : $($_.FullName)"
            [string]$folderName = $_.FullName
            [long]$thisSize = $childFolderSizes[ $folderName ]
            if( ! $thisSize )
            {
                $thisSize = Get-ChildItem -Path $folderName -Force -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $lastWritten } | Measure-Object -Sum -Property Length | Select -ExpandProperty Sum
            }
            if( $thisSize -gt 0 )
            {
                $totalSize += $thisSize
            }
        }
        [long]$fileSizes = Get-ChildItem -Path $thisFolder -Attributes !ReparsePoint+!SparseFile+!Directory -Force -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $lastWritten } | Measure-Object -Sum -Property Length | Select -ExpandProperty Sum
        $totalSize += $fileSizes
        $childFolderSizes.Add( $thisFolder , $( if ( $totalSize -gt 0 ) { $totalSize } else { -1 } ) )

        Add-Member -InputObject $_ -MemberType NoteProperty -Name Size -Value $totalSize
        $levelSize += $totalSize
    }
    Write-Verbose "Done level $thisLevel"
}

Write-Output ( "Files modified since {2} are consuming {1:N2}GB in `"{0}`"" -f $startingFolder , ( $levelSize / 1GB ) , ( Get-Date $lastWritten -Format G) )
$allItems |  Where-Object { $_.Size } | Sort -Property Size -Descending | Select -First $showFirst | Format-Table -AutoSize -Property @{n='Folder';e={$_.FullName}},@{n='Size (GB)';e={ '{0,7:f2}' -f ( $_.Size / 1GB) }},@{n='Folder Last Modified';e={$_.LastWriteTime}}

