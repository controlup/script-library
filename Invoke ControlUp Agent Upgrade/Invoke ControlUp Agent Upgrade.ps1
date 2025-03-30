 <# 
.SYNOPSIS 
    Ensures ControlUp Automation module meets the minimum required version, then schedules an Invoke-CUAgentUpdate task.

.DESCRIPTION 
    This script checks if the ControlUp.Automation module is installed and updated to the required version.
    If needed, it installs or upgrades the module. Then, it creates a scheduled task to execute Invoke-CUAgentUpdate.
    The scheduled task is deleted after execution.

    WARNING: This script will install NuGet and the ControlUp.Automation module if they are not present.

.PARAMETER desiredVersion
    Specifies the desired version of CUAgent to install. Use 'latest' to install the newest version.

.PARAMETER zipFilePath
    Specifies the path to a local CUAgent ZIP file or the string 'useOnline' to download from the official source.

.PARAMETER Tags
    If set to `"true"`, the script will check and remove a registry value related to ControlUp agent errors and log new errors.

.EXAMPLE 
    .\CUAgentUpdate.ps1 -desiredVersion "9.1.0.597" -zipFilePath "C:\temp\CUAgent.zip" -Tags "true"

.NOTES 
    Version:        2.3
    Context:        ControlUp - Automated Agent Update
    Author:         Chris Twiest
    Requires:       Task Scheduler
    Creation Date:  2024-02-11
    Updated:        2024-02-14 - No NuGet Needed
#>

param (
    [Parameter(Mandatory=$false, HelpMessage='Version string')]
    [ValidateNotNullOrEmpty()]
    [string]$desiredVersion,

    [Parameter(Mandatory=$false, HelpMessage='Path to private CUAgent ZIP file or useOnline')]
    [ValidateNotNullOrEmpty()]
    [string]$zipFilePath,

    [Parameter(Mandatory=$false, HelpMessage='Enable registry cleanup and error logging for ControlUp Agent tags')]
    [string]$Tags = "false"
)

# Set script behavior preferences
$ErrorActionPreference = 'Stop'
$VerbosePreference     = 'SilentlyContinue'
$DebugPreference       = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'

# Internal parameter: Minimum required version for ControlUp.Automation
$CUAVersion = "1.0.3"

# Apply default values after validation
if (-not $desiredVersion) { $desiredVersion = "latest" }
if (-not $zipFilePath) { $zipFilePath = "useOnline" }

Write-Output "Starting ControlUp Agent Upgrade Process..."
Write-Output "Desired CUAgent version: $desiredVersion"
Write-Output "Zip file path or source: $zipFilePath"
Write-Output "Tags enabled: $Tags"


# Function to check if a module is installed and meets minimum version
function Test-ModuleInstalled {
    param (
        [string]$ModuleName,
        [string]$MinVersion
    )

    $module = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1

    if ($module) {
        $installedVersion = [System.Version]$module.Version
        $requiredVersion = [System.Version]$MinVersion

        if ($installedVersion -ge $requiredVersion) {
            Write-Verbose "$ModuleName is installed and meets the required version ($installedVersion)."
            return $true
        } else {
            Write-Verbose "Installed version ($installedVersion) is lower than required version ($requiredVersion)."
            return $false
        }
    }

    Write-Verbose "$ModuleName is not installed."
    return $false
}

# Ensure ControlUp.Automation module meets the required version
Write-Output "Checking ControlUp.Automation module installation..."
$moduleInstalled = Test-ModuleInstalled -ModuleName "ControlUp.Automation" -MinVersion $CUAVersion

if (-not $moduleInstalled) {
    Write-Output "The module is missing or outdated. Proceeding with installation..."
       
     # Define parameters
    $ModuleUrl = "https://www.powershellgallery.com/api/v2/package/ControlUp.Automation/$CUAVersion"
    $ModuleName = "ControlUp.Automation"
    $ModuleVersion = $CUAVersion
    $DownloadPath = "$env:TEMP\$ModuleName.$ModuleVersion.nupkg"
    $ZipPath = "$env:TEMP\$ModuleName.$ModuleVersion.zip"
    $ExtractPath = "$env:TEMP\$ModuleName.$ModuleVersion"
    $ModuleInstallPath = "C:\Program Files\WindowsPowerShell\Modules\$ModuleName"

    # Step 1: Download the .nupkg file
    Write-Output "Downloading $ModuleName module..."
    Invoke-WebRequest -Uri $ModuleUrl -OutFile $DownloadPath

    # Step 2: Rename .nupkg to .zip to allow extraction
    Write-Output "Renaming .nupkg to .zip..."
    If (Test-Path $ZipPath) { Remove-Item -Force $ZipPath }
    Rename-Item -Path $DownloadPath -NewName $ZipPath -Force

    # Step 3: Extract the .zip file
    Write-Output "Extracting module..."
    if (Test-Path $ExtractPath) { Remove-Item -Recurse -Force $ExtractPath }
    Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath

    # Step 4: Identify the correct module folder
    $Psd1File = Get-ChildItem -Path $ExtractPath -Filter "*.psd1" -Recurse | Select-Object -First 1
    $DllFiles = Get-ChildItem -Path $ExtractPath -Filter "*.dll" -Recurse

    if (-not $Psd1File) {
        Write-Output "Error: Could not find the module manifest (.psd1)."
        exit 1
    }

    # Step 5: Prepare the installation directory
    Write-Output "Setting up module directory..."
    if (Test-Path $ModuleInstallPath) { Remove-Item -Recurse -Force $ModuleInstallPath }
    New-Item -ItemType Directory -Path $ModuleInstallPath | Out-Null

    # Step 6: Move necessary files into the module folder
    Write-Output "Copying module files..."
    Copy-Item -Path $Psd1File.FullName -Destination $ModuleInstallPath -Force
    Copy-Item -Path $DllFiles.FullName -Destination $ModuleInstallPath -Force

    Write-Output "Module installed successfully in $ModuleInstallPath"

    # Step 7: Import the module
    Write-Output "Importing module..."
    Import-Module $ModuleName -Force

    # Step 8: Verify installation
    Write-Output "Verifying module installation..."
    if (Get-Module -ListAvailable -Name $ModuleName) {
        Write-Output "$ModuleName successfully installed and available."
    } else {
        Write-Output "Error: Module not found after installation."
    }
}

# Generate timestamped script filename
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$scriptPath = "$env:TEMP\Invoke-CUAgentUpdate-$timestamp.ps1"
$taskName = "Invoke-CUAgentUpdate-Task-$timestamp"

# Construct the Invoke-CUAgentUpdate command dynamically
$upgradeCmd = "Invoke-CUAgentUpdate"

if ($desiredVersion -ne "latest") {
    $upgradeCmd += " -Version `"$desiredVersion`""
}

if ($zipFilePath -ne "useOnline") {
    $upgradeCmd += " -ZipFilePath `"$zipFilePath`""
}

# Create the PowerShell script that will be executed by the scheduled task
$scriptBlock = @"
#This script is generated with the Invoke-CUAgentUpdate Action in the ControlUp console.
Import-Module ControlUp.Automation
"@

# If Tags is "true", add the registry cleanup logic before executing the upgrade command
if ($Tags -eq "true") {
    $scriptBlock += @"

# Define the registry path
`$RegPath = "HKLM:\SOFTWARE\Smart-X\ControlUp\Agent\ComputerTags"

# Define the name of the registry value to check
`$RegValueName = "InstallCUA_Error"

# Check if the registry value exists
if (Get-ItemProperty -Path `$RegPath -Name `$RegValueName -ErrorAction SilentlyContinue) {
    try {
        # Remove the registry value
        Remove-ItemProperty -Path `$RegPath -Name `$RegValueName -Force
        Write-Output "Registry value '`$RegValueName' found and removed successfully."
    } catch {
        Write-Error "Failed to remove registry value '`$RegValueName'. Error: `$_"
    }
} else {
    Write-Output "Registry value '`$RegValueName' does not exist. No action needed."
}
"@
}

# Assign the upgrade command to $info before execution
$scriptBlock += @"

Write-Output "Executing upgrade command..."
`$info = $upgradeCmd
"@

# If Tags is "true", add error logging logic after the upgrade command
if ($Tags -eq "true") {
    $scriptBlock += @"

# Capture errors from the upgrade process
`$Installerror = `$info.error

if (`$Installerror) {
    # Get Error info
    if (`$info.Error -match "Exception:\s*(.*?)\s*at ControlUp.Automation") {
        `$cleanError = `$matches[1]
        Write-Output `$cleanError
        } else {
        `$cleanError = `$info.Error
    }

    Write-Output "Adding Error to Tag `$RegValueName"
    `$RegValueData = `$cleanError

    # Ensure the registry path exists, create it if it does not
    if (-not (Test-Path `$RegPath)) {
        New-Item -Path `$RegPath -Force | Out-Null
    }

    # Create or set the registry value
    New-ItemProperty -Path `$RegPath -Name `$RegValueName -Value `$RegValueData -PropertyType String -Force

    Write-Output "Registry value '`$RegValueName' created successfully in '`$RegPath'."
}
"@
}

# Cleanup scheduled task after execution
$scriptBlock += @"

Write-Output "Cleaning up scheduled task: $taskName"
`$null = Unregister-ScheduledTask -TaskName "$taskName" -Confirm:`$false
"@

# Save script to a timestamped file
$scriptBlock | Set-Content -Path $scriptPath -Encoding UTF8

# Define the scheduled task action
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`""

# Define the scheduled task principal to run as SYSTEM (to prevent permission issues)
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

#Definte the task trigger
$TriggerTime = (Get-Date).AddSeconds(20)
$Trigger = New-ScheduledTaskTrigger -Once -At $TriggerTime

# Define task settings
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Seconds 0)

# Register and start the scheduled task
try {
    Register-ScheduledTask -TaskName $taskName -Description "This script is generated with the Invoke-CUAgentUpdate Action in the ControlUp console." -Action $action -Principal $principal -Settings $settings -Trigger $trigger -Force
    Write-Output "Scheduled task successfully registered."

} catch {
    Write-Error "Failed to create or start scheduled task: $_"
    exit 1
}

Write-Output "The upgrade to $desiredVersion will now start. The upgrade may take a few minutes."
