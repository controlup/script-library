#requires -version 3
$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    This script will delete files to clean disks

    .DESCRIPTION
    This script will delete files and folders contents that are know to be safe to be removed.
    The -Force parameter is used to ensure as much as possible is removed without asking for confirmation.
    
    .PARAMETER bolCleanmgr
    True/False value indicating if CLEANMGR (with all options selected except delete System Restore points) should be run

    .PARAMETER bolVolumeShadowCopy
    True/False value indicating if System Restore points should be deleted

    .CHANGES
                7-12-2020, Ton de Vreede, Added DISM.EXE fallback (for OS version > 2008R2), increased CLEANMGR timeout
#>

# Parse the arguments
If ($args[0] -eq 'True') { [bool]$bolCleanmgr = $true } Else { [bool]$bolCleanmgr = $false }
If ($args[1] -eq 'True') { [bool]$bolVolumeShadowCopy = $true } Else { [bool]$bolVolumeShadowCopy = $false }
If ($args[2] -eq 'True') { [bool]$bolDISMFallback = $true } Else { [bool]$bolDISMFallback = $false }

# Array of files to be cleaned
[array]$arrFilesToBeDeleted = @(
    '%SystemRoot%\memory.dmp',
    '%SystemRoot%\Minidump.dmp'
)

# Array of folders to be cleaned
[array]$arrFoldersToBeCleaned = @(
    '%systemroot%\Downloaded Program Files',
    '%systemroot%\Temp',
    '%systemdrive%\Windows.old',
    '%systemdrive%\Temp',
    '%systemdrive%\MSOCache\All Users',
    '%allusersprofile%\Adobe\Setup',
    '%allusersprofile%\Microsoft\Windows Defender\Definition Updates',
    '%allusersprofile%\Microsoft\Windows Defender\Scans',
    '%allusersprofile%\Microsoft\Windows\WER'
)

# Keys to set CLEANMGR to clean, remove unwanted entries
[array]$arrSageSetKeys = @(
    'Active Setup Temp Folders',
    'BranchCache',
    'Compress System Disk',
    'Content Indexer Cleaner',
    'D3D Shader Cache',
    'Delivery Optimization Files',
    'Device Driver Packages',
    'Diagnostic Data Viewer database files',
    'Downloaded Program Files',
    'Internet Cache Files',
    'Offline Pages Files',
    'Old ChkDsk Files',
    'Previous Installations',
    'Recycle Bin',
    'RetailDemo Offline Content',
    'Service Pack Cleanup',
    'Setup Log Files',
    'System error memory dump files',
    'System error minidump files',
    'Temporary Files',
    'Temporary Setup Files',
    'Temporary Sync Files',
    'Thumbnail Cache',
    'Update Cleanup',
    'Upgrade Discarded Files',
    'User file versions',
    'Users Download Folder',
    'Windows Defender',
    'Windows Error Reporting Files',
    'Windows ESD installation files',
    'Windows Upgrade Log Files'
)

# Get current disk free space
[string]$strSystemDriveLetter = [System.Environment]::ExpandEnvironmentVariables("%systemdrive%") -replace (':', '')
$SystemDrive = Get-PSDrive -Name $strSystemDriveLetter
[double]$dblFreeSpace = $SystemDrive.Free

# Create counter to optionally record how many files could not be deleted. Not output by default
[long]$lngSkippedFileCount = 0

Function Feedback ($strFeedbackString) {
    # This function provides feedback in the console on errors or progress, and aborts if error has occured.
    If ($error.count -eq 0) {
        # Write content of feedback string
        Write-Host -Object $strFeedbackString -ForegroundColor 'Green'
    }
  
    # If an error occured report it, and exit the script with ErrorLevel 1
    Else {
        # Write content of feedback string but in red
        Write-Host -Object $strFeedbackString -ForegroundColor 'Red'
    
        # Display error details
        Write-Host 'Details: ' $error[0].Exception.Message -ForegroundColor 'Red'

        Exit 1
    }
}

Function Remove-AllFilesInFolder ($strFolder) {
    $ExpFolder = [System.Environment]::ExpandEnvironmentVariables("$strFolder")
    # Make sure folder exists, Get-Childitem -recurse can hang on folders that don't exist
    If ((Test-Path -Path "$ExpFolder") -eq $true) {
        $Files = Get-ChildItem -Path $ExpFolder -Recurse -File -Force

        # Call the function to remove the files
        Remove-FilesInArray $Files
    }
}
 
Function Remove-FilesInArray ($arrFiles) {
    Foreach ($File in $Files) {
        try {
            # Remove the file
            Remove-Item -Path $file.Fullname -Force
        }
        catch {
            $script:lngSkippedFileCount += 1
        }
    }
}

# Run CLEANMGR if required
If ($bolCleanmgr) {
    # Is CLEANMGR available on the system?
    If (Test-Path ([System.Environment]::ExpandEnvironmentVariables("%systemroot%\System32\cleanmgr.exe"))) {
        # Create the SAGESET for CLEANMGR in registry
        [string]$strRegPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches'
        Foreach ($key in $arrSageSetKeys) {
            try {
                # Check if the path actually exists before trying to set the key (keys exisitng can depend on OS< patch level etc)
                If (Test-Path "$strRegPath\$key") {
                    New-ItemProperty -Path "$strRegPath\$key" -Name 'StateFlags0066' -Value 2 -PropertyType DWORD -Force | Out-Null
                }
            }
            catch {
                Feedback "SAGESET registry keys for CLEANMGR could not be set."
            }
        }
        # Run CLEANMGR as a Job, so a time out can be used because cleanmgr sometimes hangs if run silently
        try {
            # Set timeout 
            [int]$intTimeOut = 1800
            $CodeBlock = {
                Start-Process cleanmgr.exe -Wait -ArgumentList '/SAGERUN:66'
            }
            # Start the job
            $Job = Start-Job -ScriptBlock $CodeBlock

            # Wait for the job to complete
            Wait-Job $Job -Timeout $intTimeOut | Out-Null

            # Has the job completed?
            If ($Job.State -ne 'Completed') {
                Feedback "CLEANMGR did not complete in the specified time of $intTimeOut seconds. It may have run but failed tot exit. Script will continue."
                Stop-Process -Name cleanmgr -Force
            }

            # Cleanup
            Stop-Job $Job
            Remove-Job $Job
        }
        catch {
            Feedback "CLEANMGR failed to run."
        }
    }
    elseif ($bolDISMFallback) { 
        # Test for different Wndows, as DISM has different command line options depending on OS build.
        [version]$verWindows = [Environment]::OSVersion.Version
        if ($verWindows.Major -eq 6) {
            if ($verWindows.Minor -le 1) {
                Write-Host "CLEANMGR.EXE is not available on this system and the OS version is 2008R2 or lower, so the required DISM options are available."
            }
            elseif ($verWindows.Minor -eq 2) {
                Write-Host "CLEANMGR.EXE is not available on this system, running DISM /online /Cleanup-Image /StartComponentCleanup to remove updates and old windows version(s) instead."
                try {
                    dism.exe /online /Cleanup-Image /StartComponentCleanup | Out-Null
                    If ($LastExitCode -eq 0) {
                        Write-Host "DISM.EXE ran successfully."
                    }
                    ElseIf ($LastExitCode -eq 2) {
                        Write-Host 'DISM.exe returned exit code 2. Usually, the DISM commands have done (most) their work regardless of this error. We suggest that if you do not experience any other problems with the machine you ignore this error.'
                    }
                    Else {
                        Write-Host "DISM.exe did not return the expected exit code 0, the exit code was: $lastExitCode"
                    }
                }
                catch {
                    Write-Host "There was an error running DISM.exe: $($_.Exception)"
                }
            }
        }
        else {
            Write-Host "CLEANMGR.EXE is not available on this system, running DISM /online /Cleanup-Image /StartComponentCleanup /ResetBase to remove updates and old windows version(s) instead."
            try {
                dism.exe /online /Cleanup-Image /StartComponentCleanup /ResetBase | Out-Null
                If ($LastExitCode -eq 0) {
                    Write-Host "DISM.EXE ran successfully."
                }
                ElseIf ($LastExitCode -eq 2) {
                    Write-Host 'DISM.exe returned exit code 2. Usually, the DISM commands have done (most) their work regardless of this error. We suggest that if you do not experience any other problems with the machine you ignore this error.'
                }
                Else {
                    Write-Host "DISM.exe did not return the expected exit code 0, the exit code was: $lastExitCode"
                }
            }
            catch {
                Write-Host "There was an error running DISM.exe: $($_.Exception)"
            }
        }
    }
}

# Clean Volume Shadow Copies if required
If ($bolVolumeShadowCopy) {
    try {
        vssadmin.exe delete shadows /All /Quiet | Out-Null
    }
    catch {
        Feedback "Volume Shadow Copies could not be deleted."
    }
}

# Pass the folder array to function that gathers files
Foreach ($folder in $arrFoldersToBeCleaned) {
    Remove-AllFilesInFolder $folder
}

# Delete the specified files
Foreach ($File in $arrFilesToBeDeleted) {
    Try {
        # Check if the file exists at all before trying to remove it, because default files may not exist on system
        $File = [System.Environment]::ExpandEnvironmentVariables("$File")
        If (Test-Path -Path $File) {
            Remove-Item -Path $File -Force
        }
    }
    catch {
        # File WAS FOUND but the file object could not be retreived, increase the Skipped File counter
        $lngSkippedFileCount += 1
    }
}

# Write space gained to console
$SystemDrive = Get-PSDrive -Name $strSystemDriveLetter
[double]$dblNewFreeSpace = $SystemDrive.Free
[double]$dblSpaceFreed = [math]::Round((($dblNewFreeSpace - $dblFreeSpace) / 1MB), 2)

Write-Host "$dblSpaceFreed MB of disk space freed"
