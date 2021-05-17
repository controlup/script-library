<#  .SYNOPSIS    VNC_to_a_Machine
                 Use the IP address or hostname of a machine to VNC to it.
                 This requires the other device to have VNC server installed as well as the VNC viewer in the console
                 Validated with RealVNC and TightVNC_2.8. TightVNC_1.3 fails and goes to listen mode

    .EXAMPLE     .\VNC_to_a_Machine.ps1 -vncPath 'C:\Program Files\RealVNC' -vncPort 5900  '192.168.214.117 
                 .\VNC_to_a_Machine.ps1 -vncPath 'C:\Program Files\TightVNC\tvnviewer.exe'-vncPort 5900  '192.168.214.117 
    .CONTEXT     Machine
    .COMPONENT   VNCviewer. Validated with RealVNC and TightVNC_2.8. TightVNC_1.3 fails and goes to listen mode
    .TAGS        $Machine 
    .HISTORY
                 Marcel Calef     - 2020-05-10 - Initial release
 #>

 [CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='VNC path. Will look for *viewer.exe there')]
    [ValidateNotNullOrEmpty()]                                               [string]$vncPath,
    [Parameter(Mandatory=$true,HelpMessage='TCP port to connect')]           [string]$vncPort,
    [Parameter(Mandatory=$true,HelpMessage='IP of the machine')]             [string]$inputIP,
    [Parameter(Mandatory=$false,HelpMessage='alt-IP-not used')]              [string]$altinputIP
      )

Set-StrictMode -Version Latest
$ErrorActionPreference = "continue"

$VerbosePreference = "continue"  ## comment this line to disable verbose debug output

Write-Verbose "Variables:"
Write-Verbose "          vncPath : $vncPath"
Write-Verbose "          vncPort : $vncPort"
Write-Verbose "          inputIP : $inputIP"

$vncIP  = $inputIP.split([Environment]::NewLine)[0]   # Grab the first IP if many presented
Write-Verbose "          vncIP   : $vncIP"

#Test the Path provided exists
If ((Test-Path -Path $vncPath) -ne 'True') {Write-output "$vncPath does not exist" | Msg *; exit}

# Search for the VNC viewer exact filename (*viewer.exe) and build the path for the VNC viewer
Write-Verbose "  searching for vnc exe"
$vncViewerExe = (Get-ChildItem $vncPath -Include *viewer.exe -Recurse)

if([string]::IsNullOrEmpty($vncViewerExe)){Write-Output  "Could not find a VNCviewer.exe in $vncPath or its subdirectories" | Msg *; exit}

Write-Verbose "  $vncViewerExe"


# Start a wait loop while checking if the TCP port provided responds
if(1) {  
        # using a faster test # $testVNC = Test-NetConnection -ComputerName $vncIP -port $vncPort
        # Open a pop-up message (as a job that can be killed) for up to 20 seconds while the test runs (when successful fast, you will not even see it)
        $popup = Start-Job -ScriptBlock {$wshell = New-Object -ComObject Wscript.Shell; $wshell.Popup("Testing the TCP $vncPort port",20,"Testing",64)}
        # Simple test of the TCP port
        $testVNC = New-Object System.Net.Sockets.TCPClient; $testVNC.ReceiveTimeout = 300; $testVNC.SendTimeout = 300;
        $testVNC.Connect($vncIP, $vncPort)
        
        # Check the result. If Failed, Report and exit
        if ($testVNC.Connected -ne "True") {
                        Stop-Job -id $popup.id # close the pop-up if still showing
                        Write-Output  "The proveded Machine and TCP port $vncPort does not allow connection" | Msg * 
                        exit }
      }  
      Stop-Job -id $popup.id # When sucessful, immediatly close the pop-up (stop the pop-up job)
                   
Start-Process -FilePath $vncViewerExe -ArgumentList $vncIP':'$vncPort 

