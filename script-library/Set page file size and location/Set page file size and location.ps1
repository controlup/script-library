#requires -version 3.0
<#
    Set page file location and size

    @guyrleech 2018

    Modification History:

    15/10/18  GRL  Extra error handling and re-enable automatic management if we disabled it

    23/11/18  GRL  Bugfixes and changed so no new name specified means set to automatically managed
#>

[string]$newLocation = $args[0]
[int]$initialSize = 0
[int]$maximumSize = 0

if( $args.Count -ge 2 -and $args[1] )
{
    $initialSize = $args[1]
}

if( $args.Count -ge 3 -and $args[2] )
{
    $maximumSize = $args[2]
}

if( $maximumSize -lt $initialSize )
{
    Throw 'Maximum size cannot be less than intial size'
}

if( $newLocation -eq '0' -or $newLocation -match '^auto' ) ## 0 if left as blank as SBA shifts parameters down
{
    $newLocation = $null
}

$computer = Get-CimInstance -ClassName Win32_ComputerSystem
$changed = $null
[bool]$wasAutomaticallyManaged = $false

if( ! [string]::IsNullOrEmpty( $newLocation ) )
{
    $logicalDisk = $null
    $drive = Split-Path -Path $newLocation -Qualifier -ErrorAction SilentlyContinue
    if( $drive )
    {
        $logicalDisk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "Name = '$drive' and DriveType = '3'" -ErrorAction SilentlyContinue
    }
    if( ! $logicalDisk )
    {
        Throw "Cannot find local fixed drive for $newLocation"
    }
    elseif( $logicalDisk.FreeSpace / 1MB -lt $maximumSize )
    {
        Write-Warning "Drive $drive only has $([math]::Round($logicalDisk.FreeSpace / 1MB))MB free space"
    }

    $newPageFile = Join-Path $drive 'pagefile.sys' ## default if only a drive has been specified rather than drive+folder
    [string]$folder = '\'
    if( $newLocation -match '^[a-z]:\\(.+)$' )
    {
        $file = Split-Path -Path $newLocation -Leaf
        if( ! $file )
        {
            $newLocation = Join-Path -Path $newLocation -ChildPath 'pagefile.sys'
        }
        elseif( $file -ne 'pagefile.sys' )
        {
            Throw "Page file name `"$file`" is not allowed as can only be called ""pagefiles.sys"""
        }
        $folder = Split-Path -Path $newLocation -Parent
        if( ! ( Test-Path -Path $folder -ErrorAction SilentlyContinue ) )
        {
            [void](New-Item -Path $folder -ItemType Directory -ErrorAction Stop)
        }
        $newPageFile = $newLocation
    }
    elseif( $newLocation -match '^[a-z]:[^\\]' )
    {
        Throw "Path `"$newLocation`" is invalid as an absolute path must be specified"
    }

    # Works around a PowerShell bug when trying to find if the new page file already exists
    $fileProperties = Get-ChildItem -Path $folder -ErrorAction SilentlyContinue -Force | Where-Object { $_.Name -eq (Split-Path -Path $newPageFile -Leaf) }
    if( $fileProperties )
    {
        Write-Warning "New page file $newPageFile already exists, size $([math]::Round($fileProperties.Length / 1MB))MB, created $(Get-Date $fileProperties.CreationTime -Format G), last modified $(Get-Date $fileProperties.LastWriteTime -Format G)"
    }
    $wasAutomaticallyManaged = $computer.AutomaticManagedPagefile
    if( $wasAutomaticallyManaged)
    {
        $computer.AutomaticManagedPagefile = $false
        $changed = Set-CimInstance -InputObject $computer -PassThru
        if( ! $? -or ! $changed )
        {
            Throw "Failed to enable automatic page file management"
        }
    }
}
else ## no name so setting to automatically managed
{
    if ($computer.AutomaticManagedPagefile)
    {
        Write-Warning 'Already using automatically managed pagefile so making no changes'
    }
    else
    {
        $computer.AutomaticManagedPagefile = $true
        $changed = Set-CimInstance -InputObject $computer -PassThru
        if( ! $? -or ! $changed )
        {
            Throw 'Failed to enable automatic page file management'
        }
        else
        {
            Write-Output "Page file successfully changed to being automatically managed`n`nYou must reboot before this takes effect"
        }
    }
    Exit 0
}

try
{
    [string]$oldPagefile = ''

    $currentPageFile = @( Get-CimInstance -ClassName Win32_PageFileSetting )
    
    [string]$currentPageFileName = $null

    if( ! $currentPageFile -or ! $currentPageFile.Count )
    {
        Write-Warning "There is currently no page file set"
    }
    elseif( $currentPageFile.Count -gt 1 )
    {
        Throw "There are $($currentPageFile.Count) page files already ($(($currentPageFile|Select -ExpandProperty Name) -join ' , ')) so don't know which one to change"
    }
    else
    {
        $currentPageFileName = $currentPageFile[0].Name
        if( ! $wasAutomaticallyManaged )
        {
            if( $currentPageFileName -eq $newPageFile -and $currentPageFile[0].InitialSize -eq $initialSize -and $currentPageFile[0].MaximumSize -eq $maximumSize )
            {
                Write-Warning "Current page file is already $newPageFile with same sizes so no changes required"
                Exit 0
            }

            $oldPagefile = " from $currentPageFileName"
        }
        ## delete current page file settings and create new
        $currentPageFile | Remove-CimInstance
    }
    
    $newSetting = New-CimInstance -ClassName Win32_PageFileSetting -Property  @{ 'Name' = $newPageFile } -ErrorAction Stop  

    if( ! $newSetting )
    {
        Throw "Failed to change page file$oldPagefile to $newPageFile"
    }
    else
    {
        $changedSize = Set-CimInstance -PassThru -InputObject $newSetting -Property @{
            'InitialSize' = [uint32]$initialSize
            'MaximumSize' = [uint32]$maximumSize
        }
        if( ! $? -or ! $changedSize )
        {
            Throw "Failed to change $newPageFile initial size to $initialSize MB and maximum $maximumSize MB"
        }
    }

    Write-Output "Page file successfully changed$oldPagefile to $newPageFile`n`nYou must reboot before this takes effect"
}
catch
{
    if ($changed)
    {
        $computer.AutomaticManagedPagefile = $true
        $changedBack = Set-CimInstance -InputObject $computer -ErrorAction Stop -PassThru
        if( ! $? -or ! $changedBack )
        {
            Write-Warning 'Failed to re-enable automatic page file management'
        }
    }
    Throw $_
}

