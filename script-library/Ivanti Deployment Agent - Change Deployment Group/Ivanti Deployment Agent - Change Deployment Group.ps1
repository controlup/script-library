$DeploymentGroup = $args[1]
$URL = $args[0]

If ((Test-Path "C:\Program Files\AppSense\Management Center\Communications Agent\CcaCmd.exe") -eq $false){
    "CCA not found"
    exit(1)
}

"Unregistering the computer from the Management Server"
$output = & 'C:\Program Files\AppSense\Management Center\Communications Agent\ccacmd.exe' /unregister

If (-not ($output -eq "The operation succeeded.")){
    "Error unregistering the computer from the Management Server"
    exit(1)
}

"Deleting the Communications Agent key"
Try{
    remove-item -path 'HKLM:\SOFTWARE\AppSense Technologies\Communications Agent' -Recurse -Force}
Catch{
    "Error removing the Coummications Agent key"
    exit(1)
}

"Starting the CCA"
Try{
    start-service -name 'AppSense Client Communications Agent'}
Catch{
    "Error starting the CCA"
    exit(1)
}

"Joining the $DeploymentGroup deployment group"
$output = & 'C:\Program Files\AppSense\Management Center\Communications Agent\ccacmd.exe' /URL $URL $DeploymentGroup

If (-not ($output -eq "The operation succeeded.")){
    "Error unregistering the computer from the Management Server"
    exit(1)
}

"Restarting the CCA"
Try{
    Stop-Service -name 'AppSense Client Communications Agent' -Force
    Start-Service -name 'AppSense Client Communications Agent'}
Catch{
    "Error Restarting the service."
}

"Done"
