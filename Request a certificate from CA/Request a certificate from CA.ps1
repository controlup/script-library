<#
.SYNOPSIS
    Requests a certificate using the FQDN of the selected machine from an Enterprise CA, installs it and sets the IIS default site to use this certificate

.DESCRIPTION
    Requests a certificate using the FQDN of the selected machine from an Enterprise CA, installs it and sets the IIS default site to use this certificate

.PARAMETER CertificateTemplate
    The CertificateTemplate parameter needs to match the "template name" -- not the "template display name". For the default Web Server template,
    the "Template Display Name" is "Web Server" but the "Template Name" is "WebServer".

.PARAMETER CertificateAuthority
    Certificate Authority (CA) server name. This is looked up using LDAP so this parameter needs to be 'Common Name' of the CA you want. If you leave this parameter
    blank then the default discovered CA will be used

.NOTES
    This script runs as the SYSTEM account via ControlUp. In order for the certificate to be successfully requested the Computer Active Directory object
    of the machine needs the ability to READ-ENROLL on the template. This can be done via a computers group or adding the machine to the template permissions direcly.
    
    Modification History:

    2023/10/05   TTYE   Initial public release
#>
[CmdletBinding()]

Param
(
    [Parameter(Position=0,Mandatory=$true,HelpMessage='Enter the certificate template name (eg, WebServer)')]
    [string]$CertificateTemplate ,

    [Parameter(Position=1,Mandatory=$false,HelpMessage='Enter a preferred Certificate Authority common name')][AllowEmptyString()][AllowNull()]
    [string]$CertificateAuthority 
)

$FQDN = "$((Resolve-DnsName -Name $(((Get-NetIPConfiguration)[0]).IPv4Address.ipaddress)).NameHost)"
Write-Output "FQDN of this machine was detected as: $FQDN `nThis name will be used on the certificate.`n`n"

#requests and installs the certificate
try {
    $cert = Get-Certificate -URL "ldap:///CN=$CertificateAuthority" -Template $CertificateTemplate -CertStoreLocation Cert:\LocalMachine\My -DnsName "$FQDN"
} catch {
    Write-Error $_
    exit
}

Write-Output "Certificate was    : $($cert.Status)"
Write-Output "Certificate DNS    : $($cert.Certificate.DnsNameList.punycode)"
Write-Output "Cert Expiry Date   : $($cert.Certificate.NotAfter)"
Write-Output "Cert Thumbprint    : $($cert.Certificate.Thumbprint)"

Write-Output "`nBinding Certificate to port 443 on the Default Web Site"
$Binding = Get-WebBinding -Name "Default Web Site" -Port 443 
if ($Binding -eq $null) {
    Write-Output "No WebBinding was found for port 443. Creating one..."
    New-WebBinding -Name "Default Web Site" -IP "*" -Port 443 -Protocol https
    $Binding = Get-WebBinding -Name "Default Web Site" -Port 443 
} else {
    Write-Output "`nAn existing WebBind was found. Assigning certificate to it."
}
$binding.AddSslCertificate($cert.Certificate.Thumbprint, "my")
Write-Output "Results:"
Write-Output "$($Binding | Where-Object {$_.Protocol -like "https"} | Select-Object "protocol","bindingInformation","certificateHash" | Format-List * | Out-String)"


