#Requires -version 3

<#
  .SYNOPSIS
  Check User Profile sizes on a machine.

  .DESCRIPTION
  This function Check Profile Sizes examines user profiles for all or selected users
  on the target machine, grouping the results by file type, using the extension. 
  For each group of files, if the size of the group exceeds a threshold 
  (default 15% of the total profile size) the individual files are listed
  Options provide for sorting and summarization preferences.

  .PARAMETER ThresholdPercentToExpand
  Specifies the threshold percentage for any group of files.

  .PARAMETER SamAccountNameList
  Specifies the profiles to be checked, using the sAMAccountName. 
  Specifying 'All' will check all profiles found on the machine, including those
  for local accounts.

  .PARAMETER SortBy
  When set to 'Size', items in a file-extension group will be ordered in descending order of size.
  When set to 'Path', items in a file-extension group will be ordered in ascending order by path (FullName).

  .PARAMETER PreSummarySize
  An integer, specifying the maximum number of items in a file-extension group that will be displayed in full.
  Any remaining items will be displayed in summary form.

  .NOTES
  The script uses the WMI class Win32_UserProfile to report the profiles on
  the machine, so will detect profiles regardless of where they are stored.
  The script will enumerate every file in the profile, which may cause files
  to be fetched from network locations for some profile types (e.g. Citrix
  Profiles - "UPM") and generate peaks in network traffic.
  The script outputs exact file sizes, in bytes.
  The script will differentiate between active AD accounts (using ADSI)
  and local accounts by querying WinNT accounts. Any accounts not detected by
  these techniques may correspond to deleted accounts in AD and will be identified
  as such in the output.

   
  Modification History:
  2023-03-10   Bill Powell       Initial public release

  .LINK
  For more information refer to:
    https://www.controlup.com

  .LINK
  Stay in touch:
  https://twitter.com/asoftman

  .EXAMPLE
  PS> .\Check-ProfileSize.ps1

  .EXAMPLE
  PS> .\Check-ProfileSize.ps1 -ThresholdPercentToExpand 20 -SamAccountNameList "tom,dick,harry"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false,
               HelpMessage = "Threshold of total size above which files are analyzed")]
               [int]$ThresholdPercentToExpand = 15,
    [Parameter(Mandatory=$false,
               HelpMessage = "List of account names to analyze")]
               [string]$SamAccountNameList,
    [Parameter(Mandatory=$false,
               HelpMessage = "sort order to present results")]
               [ValidateSet('Path','Size')]
               [string]$SortBy = 'Size',
    [Parameter(Mandatory=$false,
               HelpMessage = "where there are many files, the first ones are itemised, the remainder shown by folder")]
               [int]$PreSummarySize = 6
)

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'

function Get-FilesizeString {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][long]$Filesize
    )
    $Units = $null
    if ($Filesize -ge 1pb) {
        $Units = "PB"
        $DecimalSize = [decimal]($Filesize / 1PB)
    }
    elseif ($Filesize -ge 1tb) {
        $Units = "TB"
        $DecimalSize = [decimal]($Filesize / 1TB)
    }
    elseif ($Filesize -ge 1gb) {
        $Units = "GB"
        $DecimalSize = [decimal]($Filesize / 1GB)
    }
    elseif ($Filesize -ge 1mb) {
        $Units = "MB"
        $DecimalSize = [decimal]($Filesize / 1MB)
    }
    elseif ($Filesize -ge 1kb) {
        $Units = "KB"
        $DecimalSize = [decimal]($Filesize / 1KB)
    }
    if ($Units -ne $null) {
        if ($DecimalSize -ge 100) {
            [string]$result = "{0} $Units" -f [math]::Round($DecimalSize)
        }
        elseif ($DecimalSize -ge 10) {
            [string]$result = "{0:n1} $Units" -f [math]::Round($DecimalSize,1)
        }
        else {
            [string]$result = "{0:n2} $Units" -f [math]::Round($DecimalSize,2)
        }
    }
    else {
        [string]$result = $Filesize.ToString('d')
    }
    $result.PadLeft(7)
}

<#
#
# test data
$TestNumbers = @(0,1,10,100,1000,1024,1025,123456,12345678,123456789,1234567890,12345678901)

$TestNumbers | foreach {
    $Test = $_
    $message = "Input {0} -> '{1}'" -f $Test, (Get-FilesizeString $Test)
    Write-Host $message
}
#>

[int]$outputWidth = 400
try
{
    if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
    {
        Write-Verbose -Message "Setting output width to $outputWidth"
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions
        Write-Verbose -Message "Set output width to $($WideDimensions.width)"
    }
}
catch
{
    ## not much we can do but will hide the error since it is not fundamental to script functionality, just output
    Write-Warning -Message "Failed to set output width to $($WideDimensions.width) : $_"
}

Write-Output "**********************************************************************************************************"
Write-Output "System:             $($env:COMPUTERNAME)"
Write-Output "Date/Time:          $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))"
Write-Output "ThresholdPercent:   $ThresholdPercentToExpand"
Write-Output "SamAccountNameList: $SamAccountNameList"
Write-Output "SortBy:             $SortBy"
Write-Output "PreSummarySize:     $PreSummarySize"
Write-Output ""

if ($SamAccountNameList -eq "All") {
    $SamAccountNameList = $null
}
$SamAccountNameArray = @()

if (-not [string]::IsNullOrWhiteSpace($SamAccountNameList)) {
    $SamAccountNameList -split ',' -replace "^\s+",'' -replace "\s+$",'' | foreach {
        $SamAccountNameArray += $_
    }
}

$ProfileFieldList = "LocalPath,SID,SamAccountName,userPrincipalName,accountExpires" -split ','
$AllUserProfiles = @()
Get-CimInstance win32_userprofile | 
  where {$_.SID -like 'S-1-5-21-*'} | 
  Select-Object $ProfileFieldList | 
  foreach {
    $Profile = $_
    $strSID=$Profile.SID
    # see https://serverfault.com/questions/120411/retrieve-user-details-from-active-directory-using-sid
    try {
        $uSid = [ADSI]"LDAP://<SID=$strSID>"
        $Profile.SamAccountName = [string]($uSid.sAMAccountName)
        $Profile.userPrincipalName = [string]($uSid.userPrincipalName)
        $Profile.accountExpires = [string]($uSid.accountExpires)
    }
    catch {
        $e = $_
    }
    if ([string]::IsNullOrWhiteSpace($Profile.SamAccountName)) {
        $Name = Split-Path $Profile.LocalPath -Leaf
        try {
            $LocalUser = [adsi]"WinNT://./$Name,user"
            if (-not [string]::IsNullOrWhiteSpace($LocalUser.Name)) {
                $Profile.SamAccountName = $Name + " (Local account - SID $($Profile.SID) not found in AD)"
            }
            else {
                $Profile.SamAccountName = $Name + " (Deleted/anomalous account - SID $($Profile.SID) not found in AD or local accounts)"
            }
        }
        catch {
            $e = $_
            $Profile.SamAccountName = $Name + " (Deleted/anomalous account - SID $($Profile.SID) not found in AD or local accounts)"
        }
    }
    if ($SamAccountNameArray.Count -eq 0) {
        # empty list -> all profiles checked
        $AllUserProfiles += $Profile
    }
    elseif ($Profile.SamAccountName -in $SamAccountNameArray) {
        # non-empty list -> only check profiles passed as parameter
        $AllUserProfiles += $Profile
    }
}

if ($AllUserProfiles.Count -eq 0) {
    throw "No user profiles found matching $SamAccountNameList"
    exit 0
}

$script:ProfileFolderLookup = @{}
$script:AppDataPrefix = $null
#
# we want to classify files by Hidden/AppData/Normal
function Get-FileClass {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]$FileObject
    )
    if ($FileObject.Mode -match "[hs]" ) {
        return 'hidden'
    }
    elseif (([string]$FileObject.FullName).StartsWith($script:AppDataPrefix)) {
        return 'appdata'
    }
    return 'standard'
}

filter Partition-LargeFilesets ([int]$PreSummarySize) {
    begin {
        $IndividualItems = @()
        $SummaryItems = @()
        $SummaryItemSize = 0
        $ItemCount = 0
        $FolderList = @{}
    }
    process {
        if ($null -eq $PreSummarySize) {
            $IndividualItems += $_
        }
        elseif ($IndividualItems.Count -lt $PreSummarySize) {
            $IndividualItems += $_
        }
        else {
            $SummaryItems += $_
            $SummaryItemSize += $_.Length
            $StatsForDirectory = $FolderList[$_.DirectoryName]
            if ($null -eq $StatsForDirectory) {
                $StatsForDirectory = New-Object PSObject | Select-Object -Property FileCount,FileSize,Folder
                $StatsForDirectory.FileCount = 0
                $StatsForDirectory.FileSize = 0
                $StatsForDirectory.Folder = $_.DirectoryName
            }
            $StatsForDirectory.FileCount += 1
            $StatsForDirectory.FileSize += $_.Length
            $FolderList[$_.DirectoryName] = $StatsForDirectory
        }
    }
    end {
        New-Object PSObject |
            Add-Member NoteProperty IndividualItems     $IndividualItems       -PassThru |
            Add-Member NoteProperty SummaryItems        $SummaryItems          -PassThru |
            Add-Member NoteProperty SummaryItemSize     $SummaryItemSize       -PassThru |
            Add-Member NoteProperty SummaryItemCount    $SummaryItems.Count    -PassThru |
            Add-Member NoteProperty SummaryItemFolders  $FolderList            -PassThru
    }
}

switch ($SortBy) {
    "Path" {
            #
            # sort by path, ascending
            $SortExpression = @{Expression={($_.FullName)};Descending=$false}
        }
    "Size" {
            #
            # sort by file size, descending
            $SortExpression = @{Expression={($_.Length)};Descending=$true}
        }
    default {
            #
            # sort by file size, descending
            $SortExpression = @{Expression={($_.FullName)};Descending=$false}
        }
}

$AllUserProfiles | foreach {
    $FilesByExtensionHash = @{}
    $TotalFileSize = 0
    $TotalFileCount = 0
    $SystemHiddenFileSize = 0
    $SystemHiddenFileCount = 0
    $AppDataFileSize = 0
    $AppDataFileCount = 0
    $StandardFileSize = 0
    $StandardFileCount = 0

    $Profile = $_
    #
    # we need to classify the folders - at least to identify AppData 
    $script:ProfileFolderLookup = @{}
    (Get-Item $profile.LocalPath).GetDirectories() | foreach {
        $TopLevelFolder = $_
        $ProfileFolderLookup[$TopLevelFolder.Name] = $TopLevelFolder
        if ($TopLevelFolder.Mode -match "l") {
            #
            # link - can ignore
        }
        elseif ($TopLevelFolder.Mode -match "h") {
            #
            # hidden, but not link - this will be AppData
            $script:AppDataPrefix = $TopLevelFolder.FullName
        }
    }
    #
    # now we classify each of the files
    Get-ChildItem $Profile.LocalPath -Recurse -File -Force -ErrorAction SilentlyContinue | foreach {
        $FileObj = $_
        $FileClass = Get-FileClass $FileObj
        $TotalFileSize += $FileObj.Length
        $TotalFileCount++
        switch ($FileClass) {
            'hidden' {
                    $SystemHiddenFileSize += $FileObj.Length
                    $SystemHiddenFileCount++
                }
            'appdata' {
                    $AppDataFileSize += $FileObj.Length
                    $AppDataFileCount++
                }
            'standard' {
                    $StandardFileSize += $FileObj.Length
                    $StandardFileCount++
                    #
                    # now we record specifics by extension
                    $previous = $FilesByExtensionHash[$FileObj.Extension]
                    if ($previous -eq $null) {
                        $previous = New-Object psobject | select-object Extension,TotalSize,Items
                        $previous.Extension = $FileObj.Extension
                        $previous.TotalSize = $FileObj.Length
                        $previous.Items = @($FileObj)
                    }
                    else {
                        $previous.TotalSize += $FileObj.Length
                        $previous.Items += $FileObj
                    }
                    $FilesByExtensionHash[$FileObj.Extension] = $previous
                }
        }
    }

    Write-Output "**********************************************************************************************************"
    Write-Output "User: $($Profile.SamAccountName)"
                ("    Total files:         {0,12}     Total File size:         {1,15} (bytes) / {2} , comprising:" -f $TotalFileCount,$TotalFileSize,(Get-FilesizeString $TotalFileSize)) | Out-Default
                ("    Standard files:      {0,12}     Standard File size:      {1,15} (bytes) / {2}"              -f $StandardFileCount,$StandardFileSize,(Get-FilesizeString $StandardFileSize)) | Out-Default
                ("    AppData files:       {0,12}     AppData File size:       {1,15} (bytes) / {2}"              -f $AppDataFileCount,$AppDataFileSize,(Get-FilesizeString $AppDataFileSize)) | Out-Default
                ("    Hidden/System files: {0,12}     Hidden/System File size: {1,15} (bytes) / {2}"              -f $SystemHiddenFileCount,$SystemHiddenFileSize,(Get-FilesizeString $SystemHiddenFileSize)) | Out-Default
    Write-Output "**********************************************************************************************************"

    $ThresholdLength = $StandardFileSize * $ThresholdPercentToExpand / 100

    $ExtensionStats = $FilesByExtensionHash.GetEnumerator() | % {$_.Value} | Sort-Object -Property TotalSize -Descending

    $ExtensionStats | Format-Table -Property Extension,@{Label = "Total (Bytes)";Expression={$_.TotalSize}},@{Label = "Total";Expression={(Get-FilesizeString $_.TotalSize)}}

    $ExtensionStats | where {$_.TotalSize -gt $ThresholdLength} | foreach {
        Write-Output "==============================================="
        Write-Output "Expanding extension $($_.Extension)"
        $_.Items | Sort-Object -Property $SortExpression |
          #  Select-Object -Property Mode,FullName,Length | 
            Partition-LargeFilesets -PreSummarySize $PreSummarySize | 
            ForEach-Object {
                $_.IndividualItems |
                    Format-Table -Property @{Label = "Attributes";Expression={$_.Mode}},FullName,@{Label = "FileSize (bytes)";Expression={$_.Length}},@{Label = "FileSize";Expression={(Get-FilesizeString $_.Length)}}
                if ($_.SummaryItems.Count -gt 0) {
                    Write-Output "plus $($_.SummaryItemCount) items totalling $($_.SummaryItemSize) bytes ($(Get-FilesizeString $_.SummaryItemSize)) in the folders below:"
                    $_.SummaryItemFolders.GetEnumerator() | ForEach-Object {$_.Value} | Sort-Object -Property FileSize -Descending | Format-Table -Property FileCount,@{Label = "FileSize (bytes)";Expression={$_.FileSize}},@{Label = "FileSize";Expression={(Get-FilesizeString $_.FileSize)}},Folder
                }
            }
    }
}

Write-Output "Complete"

