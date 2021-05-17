$ComputerName = $args[0]
$ComputerDomain = $args[1]
$UserName = $args[2]
$FQDN = $ComputerName + "." + $ComputerDomain
$Output = "$env:temp\GPOResults.html"

If (!(Test-Path (Split-Path $Output))) {
    New-Item (Split-Path $Output) -ItemType directory | Out-Null
    If (!($?)) {
        Write-Host "Unable to create output path, please check and try again."
        Exit 1
    }
}

& gpresult.exe /S $FQDN /User $UserName /H $Output /f
$IE = (Get-ChildItem -path "$env:ProgramFiles" -include "iexplore.exe" -recurse -ea SilentlyContinue).FullName
If (!$IE) {
    If (Test-Path $Output) {
        Write-Host "Could not find Internet Explorer in the Program Files directory. Please launch $Output from another browser."
        Exit 1
    } Else {
        Write-Host "Could not write the output file. Please check the path and try again."
        Exit 1
    }
}
Start-Process $IE -ArgumentList "file://$Output"
Write-Host "SBA complete"
