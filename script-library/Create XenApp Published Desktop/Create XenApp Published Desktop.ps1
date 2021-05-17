Add-PsSnapin Citrix.XenApp.Commands

$appname = $args[0]+" Desktop"
$account = $args[1]

New-XAApplication -ApplicationType ServerDesktop -DisplayName $appname -FolderPath "Applications" -Description "Admin Desktop for Remote Administration" -WindowType "99%" -ColorDepth Colors32Bit -Accounts $account  -Servernames $args[0]
