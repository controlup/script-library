<# 
	.SYNOPSIS 
 		This script will set the correct permissions on HKLM\SOFTWARE\Microsoft\SystemCertificates\AuthRoot to fix CAPI event id 4110.
	.DESCRIPTION 
 		This script will solve CAPI event id 4110 by allowing NT SERVICE\CryptSvc full control on the HKLM\SOFTWARE\Microsoft\SystemCertificates\AuthRoot registry key and it's children	
	.EXAMPLE 
 		&'.\Set AuthRoot Registry Permission.ps1'
    .CONTEXT
        Computer
    .MODIFICATION_HISTORY 
        Full name - When (date format DD/MM/YY) - What changed 
        Drew Robbins - 26/02/19 - Initial version
        Matthew Fritz - 26/02/19 - Initial version
        Dennis Geerlings - 18/10/19 - Added error handling, comments and Get-Help comment block 
    .LINK 
        https://social.technet.microsoft.com/Forums/windowsserver/en-US/2b7e774d-2bd7-4833-818c-1429c7398ef1/correct-procedure-to-add-registry-key-permissions-for-certsvc?forum=winservergen
    .LINK 
        https://social.technet.microsoft.com/Forums/windowsserver/en-US/1b620576-98e1-4fe9-aa0e-3e73eda92058/capi2-error-access-denied?forum=winserversecurity
    .LINK 
        http://dieterboonen.blogspot.com/2017/10/root-certificate-update-issue-on-server.html
#>

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'

## Get's a reference to the Access Control List (ACL) set on the AuthRoot registry key.
$ACL  = Get-ACL HKLM:\SOFTWARE\Microsoft\SystemCertificates\AuthRoot
## Defines the account to set permissions for. 
$LocalAccount = "NT SERVICE\CryptSvc"
## Creates a new access control rule to allow the account mentioned above full control on the registry key.
$Rule = New-Object System.Security.AccessControl.RegistryAccessRule ($LocalAccount,"FullControl","Allow")
## Apply the rule to the ACL reference.
$ACL.SetAccessRule($Rule)
## Commit the changes to the ACL to the registry key. 
$ACL |Set-ACL -Path HKLM:\SOFTWARE\Microsoft\SystemCertificates\AuthRoot

# Set subkeys
$Dir = Get-Childitem "HKLM:\SOFTWARE\Microsoft\SystemCertificates\AuthRoot" -Recurse
foreach ($Folder in $Dir)
{
    $ACL.SetAccessRule($Rule)
    $ACL | Set-Acl $Folder.PSPath 
}

write-host Done!
