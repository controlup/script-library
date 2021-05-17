<#
    .SYNOPSIS: open VMare Remote Console
                Using two subroutines which can be reused:
                    LoadPowerCLI     &
                    Open-VMRC from Allen Derusha
#>
$debug="yes"      # set to yes for debug messages to appear

Function Load-PowerCLI ()
{
    if (get-installedmodule | where{$_.name -like "*vmware.powercli"})   # Check for Version 10 or 11 deployed via nuGet
        {
		$PCLIverString = 	(Find-Module "*vmware.powercli").Version.major
        }
 
    If ($PCLIverString -eq $null)   # could not find version via nuGet, check for installed package
	{
        $PCLIverString = (Get-ItemProperty HKLM:\software\Microsoft\Windows\CurrentVersion\Uninstall\* | where displayname -match 'PowerCLI').DisplayVersion
		If ($PCLIverString -eq $null) 
		{
			$PCLIverString = (Get-ItemProperty HKLM:\software\WoW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | where displayname -match 'PowerCLI').DisplayVersion
		}
        If ($PCLIverString -eq $null) 
		{
            Write-Host "PowerCLI is not installed on this computer. Please download and install the PowerCLI package from VMware at https://www.vmware.com/support/developer/PowerCLI/"
            Exit
        }
        if ($debug -eq "yes") {Write-host "Found PowerCLI version $PCLIverString"}
    }
	
    $PCLIver=[int]($PCLIverString| Out-String).Split(".")[0]    # Convert to Integer the Mayor version of PowerCLI

    if ($debug -eq "yes") {Write-host "will use $PCLIver for PowerCLI load version command"}

    $PCLI = "VMware.VimAutomation.Core" # PowrCLI module name for versions 5 and 6

    Switch -Regex ($PCLIver)
        {
        5    {  If ((Get-PSSnapin $PCLI -ErrorAction "SilentlyContinue") -eq $null) 
		           { 
                    Try    { Add-PSSnapin $PCLI  } 
                    Catch  { Write-Host "Failed to load  PowerCLI $PCLIver. Exiting"; Exit 1 }
                   }
              }

        6     { If ((Get-Module -Name $PCLI -ErrorAction SilentlyContinue) -eq $null) 
		           {
                    Try    { Import-Module $PCLI } 
			        Catch  { Write-Host "Failed to load  PowerCLI $PCLIver. Exiting"; Exit 1 }
                   }
              }

        "^1[0-1]$" { $PCLI = "vmware.powercli"
                    Try    { Import-Module -Name $PCLI } 
			        Catch  { Write-Host "Failed to load  PowerCLI $PCLIver. Exiting"; Exit 1 }
              }

        default{
                Write-Host "This version of PowerCLI seems to be unsupported. `
                            Please upgrade to the latest version of PowerCLI and try again."
               }
        }  # --> end Switch
    if ($debug -eq "yes") {Write-host "Sucessfully loaded PowerCLI, continuing"}
}

Function Open-VMRC {
    <#
    .SYNOPSIS 
      Launch a VMware Remote Console for a named VM
    .DESCRIPTION
      The script will launch the VMware Remote Console for a named VM
      This script requires the following tools from VMware to be installed locally:
      PowerCLI: https://www.vmware.com/support/developer/PowerCLI/
      VMRC: https://www.vmware.com/support/developer/vmrc/
    .COMPONENT
      This script requires the following tools from VMware to be installed on the local system
      PowerCLI: https://www.vmware.com/support/developer/PowerCLI/
      VMRC: https://www.vmware.com/support/pubs/vmrc_pubs.html or http://www.vmware.com/go/download-vmrc
    .PARAMETER Name
      Specify the name of the virtual machine
    .PARAMETER Server
      Specify the vSphere server to connect to
    .PARAMETER Username
      Specify the user name you want to use for authenticating with the server.
    .PARAMETER Password
      Specify the password you want to use for authenticating with the server.
    .EXAMPLE
      PS> RTS_Launch-VMRC.ps1 -Name guest1 -Server vcenter2 -username administrator -password password
    .NOTES
      Author: Allen Derusha
    #>

    Param(
      [Parameter(Mandatory=$True,HelpMessage="Virtual Machine name")][String]$Name,
      [Parameter(Mandatory=$True,HelpMessage="vSphere server name")][String]$Server,
      [Parameter(HelpMessage="vSphere user name")][String][String]$Username,
      [Parameter(HelpMessage="vSphere user password")][String][String]$Password
    )

    $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $vSphereCreds = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)

    # Check to see if we can find VMRC.exe and tell the user where to download it if we can't find it
    If (Test-Path "${env:ProgramFiles(x86)}\VMware\VMware Remote Console\vmrc.exe") {
      $VMRCpath = "${env:ProgramFiles(x86)}\VMware\VMware Remote Console\vmrc.exe"
    } ElseIf (Test-Path "$env:ProgramFiles\VMware\VMware Remote Console\vmrc.exe") {
      $VMRCpath = "$env:ProgramFiles\VMware\VMware Remote Console\vmrc.exe"
    } Else {
      Write-Host "Could not find VMRC.exe.  Download and install the VMRC package from VMware at https://www.vmware.com/support/pubs/vmrc_pubs.html or http://www.vmware.com/go/download-vmrc"
    }


    
    # Connect to the provided vCenter server with user provided creds if provided or with session creds if not
    Try {
        Connect-VIServer -Server $Server -Credential $vSphereCreds -WarningAction Ignore | Out-Null
    } Catch {
        Write-Host "ERROR: failed to authenticate to vSphere server $Server as user $Username"
        Exit 1
    }

    # Get the MoRef for the provided VM, fail if we can't find it
    $VMobject = Get-VM $Name
    $VMmoRef = $VMobject.extensiondata.moref.value
    If (! $VMmoRef) {
      Write-Host "ERROR: Failed to find the VM managed object reference for $Name"
      Exit 1
    }

    # Get a Clone Ticket for opening a remote console
    $Session = Get-View -Id Sessionmanager
    $Ticket = $Session.AcquireCloneTicket()
    If (! $Ticket) {
      Write-Host "ERROR: Failed to acquire session ticket"
      Exit 1
    }

    # Launch VMRC
    Start-Process -FilePath $VMRCpath -ArgumentList "vmrc://clone:$Ticket@$Server`:443/?moid=$VMmoRef"
}    # --> End Open-VMRC

###############  Main thread  #############

    $ErrorActionPreference = "Stop"
    
    if ($debug -eq "yes") {Write-host "Starting..."}
    #if ($debug -eq "yes") {Write-host "inputs: $args[0] $args[1] $args[2] $args[4]"}
    # Check the necesary variables were received
    if ($args.count -ne 5) {
        Write-Host "ControlUp must be connected to the VM's hypervisor for this command to work. Please connect to the appropriate hypervisor and try again."
        Exit 1
    }
    
    $Hyper = $args[4]
    
    If ($Hyper -ne "VMware") {
        Write-Host "This script only supports VMs on vSphere/ESXi, not other hypervisors. Please pick another computer and try again."
        Exit 1
    }
    
    Load-PowerCLI
    
    Open-VMRC -Server $args[1] -Name $args[0] -Username $args[2] -Password $args[3]
    
    Write-Host "Done"

