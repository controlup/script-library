<#
    Get module details for a running process

    @guyrleech 2018

    Modification History:

    14/11/18   GRL  Group by module directory
#>

$processId = $args[0]
$outputWidth = 400

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

if( ! ( Get-Process -Id $processId -ErrorAction SilentlyContinue ) )
{
    Throw "Unable to find process with id $processId"
}

[hashtable]$modulesDone = @{}

[array]$Modules = @( Get-Process -id $processId -ErrorAction Stop | Select -ExpandProperty Modules | ForEach-Object `
{
    $module = Get-ItemProperty -Path $_.FileName
    if( ! $modulesDone[ $module.FullName ] ) ## Get-Process sometimes repeats a module
    {
        $modulesDone.Add( $module.FullName , $true )
        $signing = $null
        $versionInfo = $module.VersionInfo
        try
        {
            ## for signed files, as in not externally signed via a catalogue, Get-AuthenticodeSignature does not return the correct certificate
            $cert = [System.Security.Cryptography.X509Certificates.X509Certificate]::CreateFromSignedFile( $module.FullName )
        }
        catch
        {
            $cert = $null
            $signing = Get-AuthenticodeSignature -FilePath $module.FullName -ErrorAction SilentlyContinue
        }
        [string]$expired = '-'
        [string]$expiryDate = '-'
        [string]$certificateSigner = '-'

        if( $cert )
        {
            $theExpiryDate = New-Object -TypeName DateTime
            if( [datetime]::TryParse( $cert.GetExpirationDateString() , [ref]$theExpiryDate ) )
            {
                $expired = if( $theExpiryDate -gt [datetime]::Now ) { 'No' } else { 'Yes' }
                $expiryDate = Get-Date -Date $theExpiryDate -Format G
            }
            else
            {
                $expired = '-'
            }
            $certificateSigner = ($cert.GetName() -split 'CN=')[-1]
        }
        elseif( $signing )
        {
            if( $signing.Status -eq 'Valid' )
            {
                $expired = if( [datetime]::Now -gt $signing.SignerCertificate.NotBefore -and [datetime]::Now -lt $signing.SignerCertificate.NotAfter ) { 'No' } else { 'Yes' }
                $expiryDate = Get-Date -Date $signing.SignerCertificate.NotAfter -Format G
                $certificateSigner = ($signing.SignerCertificate.GetName() -split 'CN=')[-1]
            }
        }
        else
        {
            $expired = '-'
        }

        $result = [pscustomobject][ordered]@{
            'Path' = $module.DirectoryName
            'Module' = $module.Name
            'Version' = if( $versionInfo ) { $versionInfo.ProductVersion } else { '-' }
            ##'File Owner' = ( Get-Acl -Path $_.FileName | Select -ExpandProperty Owner )
            'Created' = $module.CreationTime
            'Last Modified' = $module.LastWriteTime
            'Size (KB)' = [math]::Round( $module.Length / 1KB )
            'Cert Expiry' = $expiryDate
            'Cert Expired' = $expired
            'Cert Signer' = $certificateSigner
        }
        $result
    }
})

"Analysed $($Modules.Count) modules:"

$Modules | Sort 'Path','Module' | Format-Table -AutoSize -GroupBy 'Path' -Property 'Module','Version','Created','Last Modified','Size (KB)','Cert Expiry','Cert Expired','Cert Signer'

