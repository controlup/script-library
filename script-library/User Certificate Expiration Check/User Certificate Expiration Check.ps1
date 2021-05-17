$threshold = $args[0]

#Set deadline date
$deadline = (Get-Date).AddDays($threshold)

$Certs = Get-ChildItem Cert:\CurrentUser\My | where {$_.notafter -lt $deadline} |
    select issuer, subject,notafter, @{Label="Expires In (Days)";Expression={($_.NotAfter - (Get-Date)).Days}}

If ($Certs) {$Certs} Else { Write-Host "There are no certificates expiring in $threshold days." }

