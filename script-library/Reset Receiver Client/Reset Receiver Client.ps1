<#####
Script looks for cleanup.exe and selfservice.exe in the default Receiver install directory.  If they are both found it will
perform a receiver reset, wait for it to finish, and then poll storefront.  If storefront is not configured by policy or from 
Studio it will not be able to poll.
#####>

If (Test-Path 'C:\Program Files (x86)') {
    $cleanup = Get-ChildItem 'C:\Program Files (x86)\Citrix' -Recurse | Where {$_.name -eq "cleanup.exe"} | select -ExpandProperty fullname
    $selfservice = Get-ChildItem 'C:\Program Files (x86)\Citrix' -Recurse | Where {$_.name -eq "selfservice.exe"} | select -ExpandProperty fullname
} Else {
    $cleanup = Get-ChildItem 'C:\Program Files\Citrix' -Recurse | Where {$_.name -eq "cleanup.exe"} | select -ExpandProperty fullname
    $selfservice = Get-ChildItem 'C:\Program Files\Citrix' -Recurse | Where {$_.name -eq "selfservice.exe"} | select -ExpandProperty fullname
}

if ($cleanup -eq $null -or $selfservice -eq $null) {
    Write-Host "The required executables are not found in the default path."
    Write-Host "Please check that cleanup.exe and selfservice.exe are installed properly."
    Exit 1
}

$app = Start-Process $cleanup -ArgumentList "-cleanUser -silent" -PassThru
Wait-Process $app.id

if ($app.exitcode -ne "1") {
    write-host "Receiver reset successfully."
    $app1 = Start-Process $selfservice -ArgumentList "-poll -logon" -PassThru
    Wait-Process $app1.Id
} else {
    write-host "Receiver reset failed!"
    Exit 1
}
if ($app1.exitcode -ne "1") {
    write-host "StoreFront Polled successfully"
} else {
    write-host "Polling StoreFront Failed"
    Exit 1
}

