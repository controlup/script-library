<#  
.SYNOPSIS
      This script shows the network bandwidth used by a given VMware Blast session.
.DESCRIPTION
      This script measures the bandwidth of a given active Blast session, and breaks down the bandwidth consumption into the most useable Blast virtual channels.
      The output shows the bandwidth usage in kbps (kilobit per second) of each session
.PARAMETER 
      This script has 3 parameters:
      SessionId - the Id of the session from the ControlUp Console(e.g. 1)
      UserName - the user name from the console. (e.g. CONTROLUP\rotema).
      ViewClientProtocol - the Horizon session protocol name (e.g. "BLAST")
      The 2 parameters create the session name like the Get-Counter command requires.
.EXAMPLE
        ./AnalyzeBlastSession.ps1 "1" "CONTROLUP\rotema" "BLAST"

.NOTES
    Changelog
        27-10-2020 - Wouter Kursten - 

.OUTPUTS
        A list of the measured VMware Blast metrics.
.LINK
        See http://www.ControlUp.com
#>

#Defining all the parameters from the console

$SessionId=$args[0]
$username=$args[1]
$Protocol=$args[2]

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

Test-ArgsCount -ArgsCount 3 -Reason 'Did not receive all the required arguments, please check that EUC environment and the agent are connected.'


if ($Protocol.StartsWith("PCOIP")) {
    Write-Host "This is a PCOIP session, Please re-run the script against a VMware Blast session"
    exit 1
}
if ($Protocol.StartsWith("RDP")) {
    Write-Host "This is a RDP session, Please re-run the script against a VMware Blast session"
    exit 1
}



#Determine the Horizon Agent Version
$HorizonAgentVersion = Get-ItemProperty -Path 'HKLM:\SOFTWARE\VMware, Inc.\VMware VDM\' | Select ProductVersion

#######################
#defining the correct session name to the Get-Counter command session naming convention
$sessionname = "Session $SessionId $username"
$Samples = 10
$sampleinterval = 1
$receivedbytes = 0
$transmittedbytes = 0
$fps = 0
$rtt = 0
$bandwidth = 0
$jitter = 0
$packetloss = o

$counters = Get-Counter  -Counter "\VMware Blast Session Counters(session id: $SessionID; (main))\Received Bytes", "\VMware Blast Session Counters(session id: $SessionID; (main))\Transmitted Bytes", "\VMware Blast Imaging Counters(Session ID: $SessionID; Channel: Imaging; (main))\Frames per second", "\VMware Blast Session Counters(session id: $SessionID; (main))\RTT", "\VMware Blast Session Counters(session id: $SessionID; (main))\Estimated Bandwidth (Uplink)","\VMware Blast Audio Counters(session id: $SessionID; Channel: Audio; (main))\Received Bytes","\VMware Blast Audio Counters(session id: $SessionID; Channel: Audio; (main))\Transmitted Bytes","\VMware Blast CDR Counters(session id: $SessionID; Channel: CDR; (main))\Received Bytes","\VMware Blast CDR Counters(session id: $SessionID; Channel: CDR; (main))\Transmitted Bytes","\VMware Blast Clipboard Counters(session id: $SessionID; Channel: Clipboard; (main))\Received Bytes","\VMware Blast Clipboard Counters(session id: $SessionID; Channel: Clipboard; (main))\Transmitted Bytes","\VMware Blast HTML5 MMR Counters(session id: $SessionID; Channel: HTML5MMR; (main))\Received Bytes","\VMware Blast HTML5 MMR Counters(session id: $SessionID; Channel: HTML5MMR; (main))\Transmitted Bytes","\VMware Blast Imaging Counters(session id: $SessionID; Channel: Imaging; (main))\Received Bytes","\VMware Blast Imaging Counters(session id: $SessionID; Channel: Imaging; (main))\Transmitted Bytes","\VMware Blast RTAV Counters(session id: $SessionID; Channel: RTAV; (main))\Received Bytes","\VMware Blast RTAV Counters(session id: $SessionID; Channel: RTAV; (main))\Transmitted Bytes","\VMware Blast Session Counters(session id: $SessionID; (main))\Jitter (Uplink)","\VMware Blast Session Counters(session id: $SessionID; (main))\Packet Loss (Uplink)" -SampleInterval $sampleinterval -MaxSamples $Samples

$receivedbytes = $counters[-1].CounterSamples[0].CookedValue - $counters[0].CounterSamples[3].CookedValue
$transmittedbytes = $counters[-1].CounterSamples[1].CookedValue - $counters[0].CounterSamples[4].CookedValue
$audioreceivedbytes = $counters[-1].CounterSamples[5].CookedValue - $counters[0].CounterSamples[5].CookedValue
$audiotransmittedbytes = $counters[-1].CounterSamples[6].CookedValue - $counters[0].CounterSamples[6].CookedValue
$CDRreceivedbytes = $counters[-1].CounterSamples[7].CookedValue - $counters[0].CounterSamples[7].CookedValue
$CDRtransmittedbytes = $counters[-1].CounterSamples[8].CookedValue - $counters[0].CounterSamples[8].CookedValue
$Clipboardreceivedbytes = $counters[-1].CounterSamples[9].CookedValue - $counters[0].CounterSamples[9].CookedValue
$Clipboardtransmittedbytes = $counters[-1].CounterSamples[10].CookedValue - $counters[0].CounterSamples[10].CookedValue
$html5mmrreceivedbytes = $counters[-1].CounterSamples[11].CookedValue - $counters[0].CounterSamples[11].CookedValue
$html5mmrtransmittedbytes = $counters[-1].CounterSamples[12].CookedValue - $counters[0].CounterSamples[12].CookedValue
$imagingreceivedbytes = $counters[-1].CounterSamples[13].CookedValue - $counters[0].CounterSamples[13].CookedValue
$imagingtransmittedbytes = $counters[-1].CounterSamples[14].CookedValue - $counters[0].CounterSamples[14].CookedValue
$rtavreceivedbytes = $counters[-1].CounterSamples[15].CookedValue - $counters[0].CounterSamples[15].CookedValue
$rtavtransmittedbytes = $counters[-1].CounterSamples[16].CookedValue - $counters[0].CounterSamples[16].CookedValue

foreach ($counter in $counters) {

    $fps += $counter.CounterSamples[2].CookedValue
    $rtt += $counter.CounterSamples[3].CookedValue
    $bandwidth += $counter.CounterSamples[4].CookedValue
    $jitter += $counter.CounterSamples[17].CookedValue
    $packetloss += $counter.CounterSamples[18].CookedValue
}
Write-Host "__________________________________________________________________________"
Write-Host "Average VMware Blast metrics for session: .::$sessionname::. "
Write-Host "Averages are per sample for $samples samples with a $sampleinterval second interval."
Write-Host "--------------------------------------------------------------------------"


$receivedbytes = $receivedbytes / $Samples / 1024
$rounded = [math]::Round($receivedbytes)
Write-Host "Session Received Bytes`t`t`t`t`t:" $rounded "KBytes"

$transmittedbytes = $transmittedbytes / $Samples / 1024
$rounded = [math]::Round($transmittedbytes)
Write-Host "Session Transmitted Bytes`t`t`t`t:" $rounded "KBytes"

$audioreceivedbytes = $audioreceivedbytes / $Samples / 1024
$rounded = [math]::Round($audioreceivedbytes)
Write-Host "Audio Transmitted Bytes`t`t`t`t`t:" $rounded "KBytes"

$audiotransmittedbytes = $audiotransmittedbytes / $Samples / 1024
$rounded = [math]::Round($audiotransmittedbytes)
Write-Host "Audio Total Transmitted Bytes`t`t`t`t:" $rounded "KBytes"

$CDRreceivedbytes = $CDRreceivedbytes / $Samples / 1024
$rounded = [math]::Round($CDRreceivedbytes)
Write-Host "CDR Transmitted Bytes`t`t`t`t`t:" $rounded "KBytes"

$CDRtransmittedbytes = $CDRtransmittedbytes / $Samples / 1024
$rounded = [math]::Round($CDRtransmittedbytes)
Write-Host "CDR Transmitted Bytes`t`t`t`t`t:" $rounded "KBytes"

$Clipboardreceivedbytes = $Clipboardreceivedbytes / $Samples / 1024
$rounded = [math]::Round($Clipboardreceivedbytes)
Write-Host "Clipboard Received Bytes`t`t`t`t:" $rounded "KBytes"

$Clipboardtransmittedbytes = $Clipboardtransmittedbytes / $Samples / 1024
$rounded = [math]::Round($Clipboardtransmittedbytes)
Write-Host "Clipboard Transmitted Bytes`t`t`t`t:" $rounded "KBytes"

$html5mmrreceivedbytes = $html5mmrreceivedbytes / $Samples / 1024
$rounded = [math]::Round($html5mmrreceivedbytes)
Write-Host "HTML5 Multimedia Redirection Received Bytes`t`t:" $rounded "KBytes"

$html5mmrtransmittedbytes = $html5mmrtransmittedbytes / $Samples / 1024
$rounded = [math]::Round($html5mmrtransmittedbytes)
Write-Host "HTML5 Multimedia Redirection Transmitted Bytes`t`t:" $rounded "KBytes"

$imagingreceivedbytes = $imagingreceivedbytes / $Samples / 1024
$rounded = [math]::Round($imagingreceivedbytes)
Write-Host "Imaging Received Bytes`t`t`t`t`t:" $rounded "KBytes"

$imagingtransmittedbytes = $imagingtransmittedbytes / $Samples / 1024
$rounded = [math]::Round($imagingtransmittedbytes)
Write-Host "Imaging Transmitted Bytes`t`t`t`t:" $rounded "KBytes"

$rtavreceivedbytes = $rtavreceivedbytes / $Samples / 1024
$rounded = [math]::Round($rtavreceivedbytes)
Write-Host "RTAV Received Bytes`t`t`t`t`t:" $rounded "KBytes"

$rtavtransmittedbytes = $rtavtransmittedbytes / $Samples / 1024
$rounded = [math]::Round($rtavtransmittedbytes)
Write-Host "RTAV Transmitted Bytes`t`t`t`t`t:" $rounded "KBytes"

$bandwidth = $bandwidth / $Samples
$rounded = [math]::Round($bandwidth)
Write-Host "Average Estimated Bandwidth`t`t`t`t:"$rounded "Kbps"

$fps = $fps / $Samples
$rounded = [math]::Round($fps)
Write-Host "Average Frames per second`t`t`t`t:"$rounded "fps"

$rtt = $rtt / $Samples
$rounded = [math]::Round($rtt)
Write-Host "Average Round-trip time`t`t`t`t`t:"$rounded "milliseconds"

$jitter = $jitter / $Samples
Write-Host "Average jitter`t`t`t`t`t`t:"$jitter

$packetloss = $packetloss / $Samples
Write-Host "Average packet loss`t`t`t`t`t:"$packetloss "packets"


Write-Host "__________________________________________________________________________"
Write-Host "Session Details"
Write-Host "--------------------------------------------------------------------------"

# this part gets the encoder
$bol = $false
$maxretry = 10
$waitbetweentries = 1
$counter = 1 

# As VMware is using the file very frequently, try to copy the log file in the temp folder
while (($bol -eq $false) -and ($counter -le $maxretry)){
    try {
        Copy-Item "C:\ProgramData\VMware\VMware Blast\Blast-Worker-SessionId$SessionId.log" $ENV:TEMP -ErrorAction Stop
        $bol = $true
    }
    catch{
        Start-Sleep -Seconds $waitbetweentries
        $bol = $false
        $counter+=1
    }
}


# Check if the file as been successfully copied and then analyze it
if ($bol){
    $VNCRegionEncoders = Get-Content "$ENV:TEMP\Blast-Worker-SessionId$SessionId.log" | Where-Object { $_.Contains("VNCRegionEncoder") }
    $LastVNCRegionEncoder = $VNCRegionEncoders[$VNCRegionEncoders.Count - 1]
    $encoder=($LastVNCRegionEncoder.Substring($LastVNCRegionEncoder.IndexOf("VNCRegionEncoder_Create")).Split(".")[0] ).replace("VNCRegionEncoder_Create: region encoder ","")
    
    Write-host "Used Blast Codec`t`t`t`t`t:" $encoder
}
else {
    Write-Error "Impossible to retrieve the Blast log file after $maxretry tries"
}
