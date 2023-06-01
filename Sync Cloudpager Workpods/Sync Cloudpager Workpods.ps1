#requires -Version 5.0
#requires -Modules Cloudpaging

<#
.SYNOPSIS
 Initiates a sync of Cloudpager Workpods via the Cloudpager client.
.DESCRIPTION
  Initiates a sync of Cloudpager Workpods via the Cloudpager client.
  If the Cloudpager client is not installed, the script will return an error notifying the client is missing.
.EXAMPLE
 .\SyncCloudpagerWorkpods.ps1
.NOTES
 To successfully sync, the Cloudpager client must be installed on the machine.
#>

Import-Module Cloudpaging

$ProgramFilesX86 =[Environment]::getfolderpath([environment+specialfolder]::ProgramFilesX86)

$CloudpagerExePath = Join-Path $ProgramFilesX86 "Numecent\Cloudpager\Cloudpager.exe"
$arg = " /autodeploy"

if (-not(Test-Path -Path $CloudpagerExePath)) {
    try {
        Write-Error "The Cloudpager client is not installed on this device."
    }
    catch {
        throw $_.Exception.Message
    }
}
else {
    $runsync = Start-Process $CloudpagerExePath -ArgumentList $arg -Wait -NoNewWindow -PassThru
    $runsync.HasExited
    $excode = $runsync.ExitCode
    if ($excode -eq '0') {
        Write-Output "Completed"
    }
    else {
        Write-Error "Failed"
    }
}

