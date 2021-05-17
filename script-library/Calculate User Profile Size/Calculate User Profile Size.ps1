<#
.SYNOPSIS
        Caclulate the user profile folder and sub-folders size
.DESCRIPTION
        This script runs against the user profile folder and collects information about
        the number of files and file size.
.PARAMETER <paramName>
        Non at this point
.EXAMPLE
        <Script Path>\script.ps1
.INPUTS
        Positional argument of the Down-Level Logon Name (Domain\User)
.OUTPUTS
        User Profile Folder Size (and number of files), AppData size and lists the user profile subfolders size
        and number of files. 
.LINK
        See http://www.controlup.com
#>

$ProfileRoot = $env:USERPROFILE
$ItemSizeList = @()
$ItemList = (Get-ChildItem $ProfileRoot -force -recurse -erroraction SilentlyContinue `
| Measure-Object -property length -sum -erroraction SilentlyContinue)
$Aggregate  = "{0:N2}" -f ($ItemList.sum / 1MB) + " MB `($($ItemList.Count) files`)"

if (get-item "$ProfileRoot\Appdata\Local" -ErrorAction SilentlyContinue) {
$ItemList = (Get-ChildItem $ProfileRoot\Appdata\Local -force -recurse -erroraction SilentlyContinue `
| Measure-Object -property length -sum -erroraction SilentlyContinue)
$LocalSize = "{0:N2}" -f ($ItemList.sum / 1MB) + " MB"
}

$ItemList = (Get-ChildItem $ProfileRoot -force -erroraction SilentlyContinue | Where-Object {$_.PSIsContainer} | Sort-Object)
foreach ($i in $ItemList) {
        $Folder = New-Object System.Object
        $Folder | Add-Member -MemberType NoteProperty -Name "SubFolder Name" -Value $i.Name
        $Size = $null
        $SubFolderItemList = (Get-ChildItem $i.FullName -force -recurse -erroraction SilentlyContinue `
        | Measure-Object -property length -sum -erroraction SilentlyContinue)
        $Size = [decimal]::round($SubFolderItemList.sum / 1MB)
        $FileSC = $SubFolderItemList.count
        $Folder | Add-Member -MemberType NoteProperty -Name "Size (MB)" -Value $Size
        $Folder | Add-Member -MemberType NoteProperty -Name "File Count" -Value $FileSC
        $ItemSizeList += $Folder
}

Write-Output "Total profile Size: $Aggregate
AppData\Local Size: $LocalSize"
$ItemSizeList | Sort-Object "Size (MB)" -Descending | Format-Table -AutoSize
