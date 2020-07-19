<#
    Find files over a givensize and show the largest

    @guyrleech 2018
#>

[string]$startingFolder = ( Join-Path $env:SystemDrive '\' )
[long]$overSize = 100MB
[int]$showFirst = 15
[int]$modifiedWithinDays = 365
[int]$outputWidth = 200 ## could expose this as a parameter
[string]$sortBy = 'Length'
$VerbosePreference = 'SilentlyContinue'

if( $args.Count )
{
    if( ! [string]::IsNullOrEmpty( $args[0] ) )
    {
        $startingFolder = $args[0]
    }
    if( $args.Count -ge 2 -and ! [string]::IsNullOrEmpty( $args[1] ) )
    {
        $overSize = Invoke-Expression $args[1]
    }
    if( $args.Count -ge 3 -and ! [string]::IsNullOrEmpty( $args[2] ) )
    {
        $showFirst = $args[2]
    }
    if( $args.Count -ge 4 -and ! [string]::IsNullOrEmpty( $args[3] ) )
    {
        $ModifiedWithinDays = $args[3]
    }
    if( $args.Count -ge 5 -and ! [string]::IsNullOrEmpty( $args[4] ) )
    {
        if( $args[4] -eq 'Age' )
        {
            $sortBy = 'LastWriteTime'
        }
        elseif( $args[4] -eq 'Size' )
        {
            $sortBy = 'Length'
        }
        else
        {
            Write-Error "Unexpected sort field `"$($args[4])`" specified"
            Exit
        }
    }
}

# Altering the size of the PS Buffer
if( $PSWindow = (Get-Host).UI.RawUI )
{
    if( $WideDimensions = $PSWindow.BufferSize ) 
    {
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions
    }
}
[dateTime]$newerThan = (Get-Date).AddDays( -$modifiedWithinDays )
[datetime]$now = Get-Date

Get-ChildItem -Path $startingFolder -Attributes !ReparsePoint+!SparseFile -File -ErrorAction SilentlyContinue -Force -Recurse | ?{ $_.Length -gt$overSize -and $_.LastWriteTime -ge $modifiedWithinDays }| sort $sortBy -Descending | select -first $showFirst | Format-Table -Property FullName,@{n='Modified Hours Ago';e={'{0:N1}' -f ($now - $_.LastWriteTime).TotalHours}},@{n='Size (MB)';e={ ( $_.Length / 1MB ) -as [int] } }

