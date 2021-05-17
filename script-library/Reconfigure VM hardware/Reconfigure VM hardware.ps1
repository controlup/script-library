<#
    .SYNOPSIS
    This script can be used to reconfigure a VMWare Virtual Machine hardware configuration

    .DESCRIPTION
    The script allows changing the (basic) hardware configuration of a VMware VM. The number of CPUs, memory size (in GB) and hard disk size (in GB) can be changed'
    - If (CPU,Mem or HDD) input is empty or similar to current configuration it will be left unchanged
    - Hard disks can only remain similar or be expanded
    - If more than one hard disk is found, the script will exit as this is considererd a risky operation
    - If the machine is powered on, VMWare tools will need to be present on the machine as it needs to be powered down gracefully. If these tools are not rpesent script will exit
    - After configuration the machine is returend to it's previous power state (if it was on it will be turned on again, if it was off it will be left off)

    .PARAMETER HypervisorPlatform
    The hypervisor platform, will be passed from ControlUp Console

    .PARAMETER strVCenter
    The name of the vcenter server that will be connected to to run the PowerCLI commands

    .PARAMETER VMName
    The name of the virtual machine the action is to be performed on/for.

    .PARAMETER CPUCount
    The target number of CPUs. If the input is '0' then the current configuration will not be changed.

    .PARAMETER MemGB
    The target amount of memory in GB. If the input is '0' then the current configuration will not be changed.

    .PARAMETER HarddiskGB
    The target hard disk size in GB. If the input is '0' then the current configuration will not be changed.

    .EXAMPLE
    Example is not relevant as this script will be called through ControlUp Console

    .NOTES
    VMware PowerCLI needs to be installed on the machine running the script.
 
#>

$strHypervisorPlatform = $args[0]
$strVCenter = $args[1]
$strVMName = $args[2]
$user = $args[3]
$password = $args[4]
$intTargetCPU = $args[5]
$intTargetMemGB = $args[6]
$intTargetHDDGB = $args[7]

If ($args.count -ne 8) {
    Write-Host "ControlUp must be connected to the VM's hypervisor for this command to work. Please connect to the appropriate hypervisor and try again."
    Exit 1
}

# Clear $error in case script is run in persistent environment. Set ErrorAction preference.
$error.Clear()
$ErrorActionPreference = 'Stop'

Function Feedback ($strFeedbackString)
{
  # This function provides feedback in the console on errors or progress, and aborts if error has occured.
  If ($error.count -eq 0)
  {
    # Write content of feedback string
    Write-Host -Object $strFeedbackString -ForegroundColor 'Green'
  }
  
  # If an error occured report it, and exit the script with ErrorLevel 1
  Else
  {
    # Write content of feedback string but in red
    Write-Host -Object $strFeedbackString -ForegroundColor 'Red'
    
    # Display error details
    Write-Host 'Details: ' $error[0].Exception.Message -ForegroundColor 'Red'

    Exit 1
  }
}
Function Load-PowerCLI ()
{
    if (Get-Command "*find-module")
    {
	$PCLIver = 	(Find-Module "*vmware.powercli").Version.major
    }
    If ($PCLIver -eq $null) 
	{
        $PCLIver = (Get-ItemProperty HKLM:\software\Microsoft\Windows\CurrentVersion\Uninstall\* | where displayname -match 'PowerCLI').DisplayVersion
		If ($PCLIver -eq $null) 
		{
			$PCLIver = (Get-ItemProperty HKLM:\software\WoW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | where displayname -match 'PowerCLI').DisplayVersion
		}
        If ($PCLIver -eq $null) 
		{
            Write-Host "PowerCLI is not installed on this computer. Please download and install the PowerCLI package from VMware at https://www.vmware.com/support/developer/PowerCLI/"
            Exit
        }
    }
	($PCLIver| Out-String).Split(".")[0]

	If ($PCLIver -ge 10) 
	{
		$PCLI = "vmware.powercli"
            Try 
			{
				Import-Module -Name $PCLI
            } 
			Catch 
			{
                  Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
                  Exit 1
            }
    }
    elseIf ($PCLIver -ge "6") 
	{
		$PCLI = "VMware.VimAutomation.Core"
        If ((Get-Module -Name $PCLI -ErrorAction SilentlyContinue) -eq $null) 
		{
            Try 
			{
                  Import-Module $PCLI
            } 
			Catch 
			{
                  Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
                  Exit 1
            }
        }
    }
	ElseIf ($PCLIver -ge "5") 
	{
		$PCLI = "VMware.VimAutomation.Core"
        If ((Get-PSSnapin $PCLI -ErrorAction "SilentlyContinue") -eq $null) 
		{
            Try 
			{
                Add-PSSnapin $PCLI
            } Catch 
			{
                Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
                Exit 1
            }
        }
    } 
	Else 
	{
        Write-Host "This version of PowerCLI seems to be unsupported. Please upgrade to the latest version of PowerCLI and try again."
    }
}
Load-PowerCLI


Function Connect-VCenterServer ($strVCenter)
{ 
    Try {
        # Connect to VCenter server
        Connect-VIServer $strVCenter -user $user -password $password -WarningAction SilentlyContinue -Force | Out-Null
    }
    Catch {
        Feedback  "There was a problem connecting to VCenter server $strVCenter. Please correct the error and try again."
    }
}

Function Disconnect-VCenterServer ($strVCenter)
{
  # This function closes the connection with the VCenter server 'VCenter'
  try
  {
    # Disconnect from the VCenter server
    Disconnect-VIServer -Server $strVCenter -Confirm:$false
  }
  catch
  {
    Feedback "There was a problem disconnecting from VCenter server $strVCenter"
  }
}

Function Check-HypervisorPlatform ($strHypervisorPlatform)
{
  # This function checks if the hypervisor is supported by this script.
  If ($strHypervisorPlatform -ne 'VMWare')
  {
    Feedback "Currently this script based action only supports VMWare, selected guest $strVMName is not running VMWare"
    Exit 0
  }
}

Function Get-VMWareVirtualMachine ($strVMName)
{
  # This function retrieves the VMware VM
  try
  {
    $script:objVM = Get-VM -Name $strVMName
  }
  catch
  {
    Feedback ("There was a problem retrieving virtual machine $strVMName.")
  }
}

Function Get-CurrentHDD ($objVM,$intTargetHDDGB)
{
  # It's too dangerous to reconfigure VMs with more than 1 HDD, if this is the case exit script. Otherwise get the HDD size.
  $script:CurrentHDD = Get-HardDisk -VM $objVM
  If ($CurrentHDD.Count -eq 1 ) 
  {
    [int]$script:intCurrentHDDGB = $CurrentHDD.CapacityGB
  }
  Else 
  {
    Feedback 'More than one hard disk found. This script only works on VMs with 1 hard disk.'
    Exit 0
  }
    
  # HDD size can only be increased, so if it is smaller than current size, exit.
  If ($intTargetHDDGB -lt $intCurrentHDDGB)
  {
    Feedback 'HDD size can only remain equal or be increased, target HDD size is smaller than current HDD size.'
    Exit 0
  }
}

Function Get-CurrentMemCPUAndPowerState ($objVM)
{
  # Retrieve the current configuration 
  [int]$script:intCurrentCPU = $objVM.NumCpu
  [int]$script:intCurrentMemGB = $objVM.MemoryGB
  [string]$script:strPowerState = $objVM.PowerState
}
  
Function Check-VMwareToolsStatus ($objVM)
{
  # This function checks if VMware Tools are isntalled. If no version of tools is returned this means they are not installed.
  If ((Get-VMGuest -VM $objVM).ToolsVersion -eq '')
  {
    Feedback 'There are no VMware Guest Tools installed in the guest OS and the VM is powered on. Powered machines can only be reconfigured if these tools are installed.'
    Exit 0
  }
}
 
Function Stop-VirtualMachine ($strVMName)
{
  # Issues a Shutdown on the target machine and waits for the PowerState to report it's really Off
  try
  {
    Stop-VMGuest -VM $strVMName -Confirm:$false
    
    # Wait until machine is down, or fail after three minutes
    [int]$intWaitCount = 0
    do 
    {
      Start-Sleep -Seconds 3
          
      # Increase Wait Count
      $intWaitCount++
          
      # If Wait Count is 60 then three minutes have passed, too long, fail the script
      If ($intWaitCount -eq 60)
      {
        Feedback 'Shutting down the target machine took too long (over three minutes), script will exit.'
        Exit 1
      }
    } until((Get-VM -Name $strVMName).Powerstate -eq 'PoweredOff')
  }
  catch
  {
    Feedback ('There was a problem shutting down the target machine.')
  }
}
 
Function Set-NewVMWareMachineConfig ($strVMName,$intTargetCPU,$intTargetMemGB,$intTargetHDDGB,$bolHDDUnchanged)
{
  # Reconfigure the machine
  # HDD configuration is skipped if HDDSizeUnchanged is $true
  # CPU and memory configuration is awlays set, allowing for a 'hard reset'  
  Try
  {
    # (Re)configure CPU and memory
    Set-VM -VM $strVMName -NumCpu $intTargetCPU -MemoryGB $intTargetMemGB -Confirm:$false >$null
    
    # Only call HDD reconfiguration if target size is different than current size
    If ($bolHDDUnchanged -ne $true)
    {
      Get-HardDisk -VM $objVM |  Set-HardDisk -CapacityGB $intTargetHDDGB -Confirm:$false >$null
    }
  }
  Catch
  {
    Feedback 'There was a problem reconfiguring the target machine.'
  }
}

Function Start-VirtualMachine ($objVM)
{
  try
  {
    Start-VM -VM $objVM -Confirm:$false
  }
  catch
  {
    Feedback 'Machine has been reconfigured but there was a problem powering it on again.'
  }
}

Function Report-VMConfig ($strVMName)
{
  # Reload VM properties
  $objReportVM = Get-VM -Name $strVMName
  
  # Create custom object for output
  $script:objOutput = New-Object -TypeName PSObject -Property @{
    VM        = $strVMName
    'CPU Count' = $objReportVM.NumCpu
    'Memory GB' = $objReportVM.MemoryGB
    'HDD GB'  = (Get-HardDisk $objReportVM).CapacityGB
  }
}
    
# Main    
# Check if hypervisor is supported (VMWare)
Check-HypervisorPlatform $strHypervisorPlatform

# Check if any relevant values were entered, if all values were '0' no change is required.
If (($intTargetCPU -eq 0) -and ($intTargetMemGB -eq 0) -and ($intTargetHDDGB -eq 0))
{
  Feedback 'Only values of 0 were given for CPU, memory and HDD size. This means no reconfiguration is required.'
  Exit 0
}
 
# Connect to the VCenter server
Connect-VCenterServer $strVCenter
 
# Get the target virtual machine
Get-VMWareVirtualMachine $strVMName > $null
 
# If HDD reconfiguration is required, check if there is only 1 HDD and that the target size is equal to or greater than current size.
If ($intTargetHDDGB -gt 0) 
{
  Get-CurrentHDD $objVM $intTargetHDDGB > $null
}
 
# Get the configuration and powerstate
Get-CurrentMemCPUAndPowerState $objVM > $null

# Fill the variables if they were '0'
If ($intTargetCPU -eq 0) 
{
  $intTargetCPU = $intCurrentCPU
}
If ($intTargetMemGB -eq 0) 
{
  $intTargetMemGB = $intCurrentMemGB
}

# Because HDD is configured in another task than CPU and memory, check if this task needs to be run at all
If (($intTargetHDDGB -eq $CurrentHDD.CapacityGB) -or ($intTargetHDDGB -eq 0)) 
{
  [bool]$bolHDDUnchanged = $true
}

# Check if there is actually any change in configuration
If (($intTargetCPU -eq $intCurrentCPU) -and ($intTargetMemGB -eq $intCurrentMemGB) -and ($bolHDDUnchanged))
{
  Feedback 'Machine already has desired hardware configuration.'
  Exit 0
}
 
# If the VM is On the VMWare Guest Tools need to be installed for the script to power manage it. VM is then turned off.
If ($strPowerState -ne 'PoweredOff')
{
  Check-VMwareToolsStatus $objVM > $null
  Stop-VirtualMachine $objVM > $null
}
 
# Reconfigure the virtual machine
Set-NewVMWareMachineConfig $strVMName $intTargetCPU $intTargetMemGB $intTargetHDDGB $bolHDDUnchanged > $null

# If the machine was On when the script started, switch it back on again
If ($strPowerState -eq 'PoweredOn') 
{
  Start-VirtualMachine $objVM > $null
}

# Output the current (new) configuration as a table
Report-VMConfig $strVMName > $null
$objOutput | Format-Table -Property VM, 'CPU Count', 'Memory GB', 'HDD GB' -AutoSize

# Disconnect from VCenter
Disconnect-VCenterServer $strVCenter

