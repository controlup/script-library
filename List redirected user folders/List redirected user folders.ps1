Function Get-RegistryKeyPropertiesAndValues
{
  <#
    This function is used here to retrieve registry values while omitting the PS properties
    Example: Get-RegistryKeyPropertiesAndValues -path 'HKCU:\Volatile Environment'
    Origin: Http://www.ScriptingGuys.com/blog
    Via: http://stackoverflow.com/questions/13350577/can-powershell-get-childproperty-get-a-list-of-real-registry-keys-like-reg-query
  #>

 Param(
  [Parameter(Mandatory=$true)]
  [string]$path
  )

  Push-Location
  Set-Location -Path $path
  Get-Item . |
  Select-Object -ExpandProperty property |
  ForEach-Object {
      New-Object psobject -Property @{"Folder"=$_;
        "RedirectedLocation" = (Get-ItemProperty -Path . -Name $_).$_}}
  Pop-Location
}

# Get the user profile path, while escaping special characters because we are going to use the -match operator on it
$Profilepath = [regex]::Escape($env:USERPROFILE)

# List all folders
$RedirectedFolders = Get-RegistryKeyPropertiesAndValues -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" | Where-Object {$_.RedirectedLocation -notmatch "$Profilepath"}
if ($RedirectedFolders -eq $null) {
    Write-Output "No folders are redirected for this user"
} else {
    $RedirectedFolders | format-list *
}
