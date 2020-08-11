#requires -Version 3.0
$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    This script will take a screenshot of the CONSOLE of an ESX virtual machine.

    .DESCRIPTION
    This script uses CreateScreenshot_Task of an ESXi virtual machine through vCenter. The screenshot is the moved from the datastore folder of the VM to a location of choice.

    .PARAMETER strHypervisorPlatform
    The hypervisor platform, will be passed from ControlUp Console

    .PARAMETER strVCenter
    The name of the vcenter server that will be connected to to run the PowerCLI commands, will be passed from ControlUp Console

    .PARAMETER strVMName
    The name of the virtual machine for the screenshot, will be passed from ControlUp Console

    .PARAMETER strScreenShotPath
    The location the screenshot should be saved, without a trailing backslash (ie. \\server\share\screenshots). Environment
    variables such as %USERPROFILE% may be used.

    .NOTES
    Screenshots are placed in the virtual machine configuration folder by default. The script moves the screenshot to the desired target folder. For these steps to succeed the account running the script needs the following priviliges:
    1. Virtual Machine - Interaction - Create screenshot
    2. Datastore - Browse Datastore
    3. Datastore - Low level file operations

    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli
    
    .COMPONENT
    VMWare PowerCLI 6.5.0R1 or higher
#>

# Hypervisor platform
[string]$strHypervisorPlatform = $args[0]
# Name of Vcenter, parsed from Console input
[string]$strVCenter = $args[1].Replace('https://', '').Split('/')[0]
# Name of virtual machine
[string]$strVMName = $args[2]
# Path to store the screenshot
[string]$strScreenShotPath = $args[3]
# Ignore certificate errors for connecting the environment
[string]$strIgnoreCertificateError = $args[4]

Function Out-CUConsole {
    <# This function provides feedback in the console on errors or progress, and aborts if error has occured.
      If only Message is passed this message is displayed
      If Warning is specified the message is displayed in the warning stream (Message must be included)
      If Stop is specified the stop message is displayed in the warning stream and an exception with the Stop message is thrown (Message must be included)
      If an Exception is passed a warning is displayed and the exception is thrown
      If an Exception AND Message is passed the Message message is displayed in the warning stream and the exception is thrown
    #>

    Param (
        [Parameter(Mandatory = $false)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [switch]$Warning,
        [Parameter(Mandatory = $false)]
        [switch]$Stop,
        [Parameter(Mandatory = $false)]
        $Exception
    )
    # Throw error, include $Exception details if they exist
    if ($Exception) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        If ($Message) {
            Write-Warning -Message "$Message`n$($Exception.CategoryInfo.Category)`nPlease see the Error tab for the exception details."
            Write-Error "$Message`n$($Exception.Exception.Message)`n$($Exception.CategoryInfo)`n$($Exception.Exception.ErrorRecord)" -ErrorAction Stop
        }
        Else {
            Write-Warning "There was an unexpected error: $($Exception.CategoryInfo.Category)`nPlease see the Error tab for details."
            Throw $Exception
        }
    }
    elseif ($Stop) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        Write-Warning -Message "There was an error.`n$Message"
        Throw $Message
    }
    elseif ($Warning) {
        # Write the warning to Warning stream, thats it. It's a warning.
        Write-Warning -Message $Message
    }
    else {
        # Not an exception or a warning, output the message
        Write-Output -InputObject $Message
    }
}

Function Test-ArgsCount {
    <# This function checks that the correct amount of arguments have been passed to the script. As the arguments are passed from the Console or Monitor, the reason this could be that not all the infrastructure was connected to or there is a problem retreiving the information.
      This will cause a script to fail, and in worst case scenarios the script running but using the wrong arguments.
      The possible reason for the issue is passed as the $Reason.
      Example: Test-ArgsCount -ArgsCount 3 -Reason 'The Console may not be connected to the Horizon View environment, please check this.'
      Success: no ouput
      Failure: "The script did not get enough arguments from the Console. The Console may not be connected to the Horizon View environment, please check this.", and the script will exit with error code 1
      Test-ArgsCount -ArgsCount $args -Reason 'Please check you are connectect to the XXXXX environment in the Console'
    #>    
    Param (
        [Parameter(Mandatory = $true)]
        [int]$ArgsCount,
        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    # Check all the arguments have been passed
    if ($args.Count -ne $ArgsCount) {
        Out-CUConsole -Message "The script did not get enough arguments from the Console. $Reason" -Stop
    }
}

function Load-VMWareModules {
    <# Imports VMware modules
      NOTES:
      - The required modules to be loaded are passed as an array.
      - In versions of PowerCLI below 6.5 some of the modules can't be imported (below version 6 it is Snapins only) using so Add-PSSnapin is used (which automatically loads all VMWare modules)
  #>

    param (    
        [parameter(Mandatory = $true,
            HelpMessage = "The VMware module to be loaded. Can be single or multiple values (as array).")]
        [array]$Components
    )

    # Try Import-Module for each passed component, try Add-PSSnapin if this fails (only if -Prefix was not specified)
    # Import each module, if Import-Module fails try Add-PSSnapin
    foreach ($component in $Components) {
        try {
            $null = Import-Module -Name VMware.$component
        }
        catch {
            try {
                $null = Add-PSSnapin -Name VMware
            }
            catch {
                Out-CUConsole -Message 'The required VMWare modules were not found as modules or snapins. Please check the .NOTES and .COMPONENTS sections in the Comments of this script for details.' -Stop
            }
        }
    }
}

Function Connect-VCenterServer {
    param (
        [parameter(Mandatory = $true)]
        [string]$VcenterName
    )
    Try {
        # Connect to VCenter server
        Connect-VIServer $VCenterName
    }
    Catch {
        Out-CUConsole -Message "There was a problem connecting to VCenter server $VCenterName. Please correct the error and try again." -Exception $_
    }
}

Function Disconnect-VCenterServer {
    # This function closes the connection with the VCenter server 'VCenter'
    param (
        [parameter(Mandatory = $true)]
        $Vcenter
    )
    try {
        # Disconnect from the VCenter server
        Disconnect-VIServer -Server $VCenter -Confirm:$false
    }
    catch {
        Out-CUConsole -Message "There was a problem disconnecting from VCenter server $VCenter. This it not a serious problem as the connection will close when this script ends but keep an eye on this reccuring." -Warning
        Exit 0
    }
}

Function Get-VMWareVirtualMachine {
    param (
        [parameter(Mandatory = $true)]
        [string]$VMName,
        [parameter(Mandatory = $true)]
        $VCenter
    )
    # This function retrieves the VMware VM
    try {
        Get-VM -Name $VMName -Server $VCenter
    }
    catch {
        Out-CUConsole -Message "There was a problem retrieving virtual machine $VMName." -Exception $_
    }
}

Function Get-VMWareVIMException {
    param (
        [parameter(Mandatory = $true)]
        [Management.Automation.ErrorRecord]$Exception
    )
    If ($_.Exception.MethodFault.GetType().Name -eq 'NoPermission') {
        [string]$Privilege = $_.Exception.MethodFault.PrivilegeID
        [string]$Explanation = (Get-VIPrivilege -Id $Privilege).Description
        [string]$ObjectType = $_.Exception.MethodFault.Object.Type
        [string]$ObjectName = (Get-View -Id $_.Exception.MethodFault.Object).Name | Get-Unique
        Write-Output "The account used to run this is script is missing a required vCenter privilege for $ObjectType $ObjectName`:`n$Privilege ($Explanation)."
    }
    else {
        Write-Output $Exception.Exception.Message
    }
}

Function Move-DatastoreFile {
    param (
        [parameter(Mandatory = $true)]
        [string]$SourceFile,
        [parameter(Mandatory = $true)]
        [string]$DestinationFile
    )
    # Copy the file to the desired location
    try {
        Copy-DatastoreItem -Item $SourceFile -Destination $DestinationFile
    }
    catch {
        If ($_.Exception.Message.Trim().EndsWith('Response status code does not indicate success: 401 (Unauthorized).')) {
            Out-CUConsole -Message "Screenshot was created but could not be copied to the desired location. This indicates a permission problem, please ensure that the account used to run this script has the 'Low level file operation' privilege on datastore $($strScreenshotFullPath.Split('\')[2]) (or at least the folder this VM resides in)." -Stop
        }
        else {
            Out-CUConsole -Message "There was an unexpected error while copying the screenshot file to the target storage path. This could be a permission issue on the target." -Exception $_
        }
    }
    # Remove the original file
    try { 
        Remove-Item -Path $SourceFile -Force
    }
    catch {
        Out-CUConsole -Message "There was an unexpected error while deleting the screenshot file from the virtual machine folder on the datastore." -Exception $_
    }
}

Function Test-HypervisorPlatform ([string]$HypervisorPlatform) {
    # This function checks if the hypervisor is supported by this script.
    If ($strHypervisorPlatform -ne 'VMWare') {
        Out-CUConsole -Message "Currently this script based action only supports VMWare, selected guest $strVMName is not running on VMWare" -Stop
    }
}

# Test correct ammount of arguments was passed
Test-ArgsCount -ArgsCount 5 -Reason 'The Console or Monitor may not be connected to the ESX environment, please check this.'

# Import the VMware PowerCLI modules
Load-VMwareModules -Components @('VimAutomation.Core')

# Set PowerCLI configuration
if ($strIgnoreCertificateError -eq 'True') {
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -ParticipateInCeip:$false -Scope Session -Confirm:$false | Out-Null
}
else {
    Set-PowerCLIConfiguration -InvalidCertificateAction Fail -ParticipateInCeip:$false -Scope Session -Confirm:$false | Out-Null
}

# Connect to VCenter server for VMWare
$objVCenter = Connect-VCenterServer -VcenterName $strVCenter

# Get the virtual machine
$objVM = Get-VMWareVirtualMachine -VMName $strVMName -VCenter $objVCenter

# Create View
try {
    $objView = $objVM | Get-View
}
catch {
    Out-CUConsole -Message "There was an issue creating the View object used to take the screenshot." -Exception $_
}

# Take the screenshot
try {
    $strScreenshotResult = $objView.CreateScreenshot().Replace('[', '').Replace('] ', '\').Replace('/', '\')
}
catch [VMware.Vim.VimException] {
    # Call handler function
    Out-CUConsole -Message $(Get-VMWareVIMException -Exception $_) -Stop
}
catch {
    Out-CUConsole -Message "There was an unexpected error while taking the screenshot." -Exception $_
}

# Create the full path to the screenshot source
[string]$strDatacenterName = (Get-Datacenter -VM $objVM).Name
[string]$strScreenshotFullPath = "vmstore:\$strDatacenterName\$strScreenshotResult"

# Set destination filepath and name
[string]$strNow = (Get-Date).ToString("yyyyMMdd-HHmmss")
$strScreenShotDestinationFileName = "$strNow-$strVMName`.png"
$strFileDestinationPathAndName = ([System.Environment]::ExpandEnvironmentVariables("$strScreenShotPath\$strScreenShotDestinationFileName"))

# Test if account has access to the file
if (!(Test-Path $strScreenShotFullPath)) {
    Out-CUConsole -Message "Screenshot was created but could not be found. This indicates a permission problem, please ensure that the account used to run this script has the 'Browse datastore' privilege on datastore $($strScreenshotFullPath.Split('\')[2]) (or at least the folder this VM resides in)." -Stop
}

# Move the screenshot to the target folder
Move-DatastoreFile -SourceFile $strScreenshotFullPath -DestinationFile $strFileDestinationPathAndName

# Report success
Out-CUConsole "Screenshot successfully created and moved to the target folder. The screenshot was saved with the following name format: yyyyMMdd-HHmmss-VMNAME.png"

# Disconnect Vcenter server
Disconnect-VCenterServer $objVCenter
