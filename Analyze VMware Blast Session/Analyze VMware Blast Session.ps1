<#  
.SYNOPSIS
      This script shows the network bandwidth used by a given VMware Blast session.
.DESCRIPTION
      This script measures the bandwidth of a given active Blast session, and breaks down the bandwidth consumption into the most useable ICA virtual channels.
      The output shows the bandwidth usage in kbps (kilobit per second) of each session
.PARAMETER 
      This script has 3 parameters:
      SessionName - the name of the session from the ControlUp Console(e.g. "Console")
      UserName - the user name from the console. (e.g. CONTROLUP\rotema).
      ViewClientProtocol - the Horizon session protocol name (e.g. "BLAST")
      The 2 parameters create the session name like the Get-Counter command requires.
.EXAMPLE
        ./AnalyzeBlastSession.ps1 "Console" "CONTROLUP\rotema" "BLAST"
.OUTPUTS
        A list of the measured VMware Blast metrics.
.LINK
        See http://www.ControlUp.com
#>

#Defining all the parameters from the console

$originalSession = $args[0].ToString().Replace("#" , " ")
if ($args[2].StartsWith("PCOIP")) {
    Write-Host "This is a PCOIP session, Please re-run the script against a VMware Blast session"
    exit 1
}
if ($args[2].StartsWith("RDP")) {
    Write-Host "This is a RDP session, Please re-run the script against a VMware Blast session"
    exit 1
}
$username = $args[1].ToString().Split("\")

#Determine the Horizon Agent Version
$HorizonAgentVersion = Get-ItemProperty -Path 'HKLM:\SOFTWARE\VMware, Inc.\VMware VDM\' | Select ProductVersion

#######################
#defining the correct session name to the Get-Counter command session naming convention
$correctUserName = $username[1]
$sessionname = "$originalSession ($correctUserName)"
$Samples = 10
$receivedbytes = 0
$transmittedbytes = 0
$fps = 0 
$rtt = 0
$bandwidth = 0

If ([version]$HorizonAgentVersion.ProductVersion -ge [version]7.3) {

    $querys = Get-Counter  -Counter "\VMware Blast Session Counters(*)\Received Bytes", "\VMware Blast Session Counters(*)\Transmitted Bytes", "\VMware Blast Imaging Counters(*)\Frames per second", "\VMware Blast Session Counters(*)\RTT", "\VMware Blast Session Counters(*)\Estimated Bandwidth (Uplink)" -SampleInterval 1 -MaxSamples $Samples

    foreach ($query in $querys) {
        $receivedbytes += $query.CounterSamples[0].CookedValue
        $transmittedbytes += $query.CounterSamples[1].CookedValue
        $fps += $query.CounterSamples[2].CookedValue
        $rtt += $query.CounterSamples[3].CookedValue
        $bandwidth += $query.CounterSamples[4].CookedValue
    }
    Write-Host "__________________________________________________________________________"
    Write-Host "Average VMware Blast metrics for session: .::$sessionname::."
    Write-Host "--------------------------------------------------------------------------"

    $receivedbytes = $receivedbytes / 1024
    $rounded = [math]::Round($receivedbytes)
    Write-Host "Total Received Bytes`t`t:" $rounded "KBytes"

    $transmittedbytes = $transmittedbytes / 1024
    $rounded = [math]::Round($transmittedbytes)
    Write-Host "Total Transmitted Bytes`t:" $rounded "KBytes"

    $bandwidth = $bandwidth / $Samples
    $rounded = [math]::Round($bandwidth)
    Write-Host "Estimated Bandwidth`t:"$rounded "Kbps"

    $fps = $fps / $Samples
    $rounded = [math]::Round($fps)
    Write-Host "Frames per second`t:"$rounded "fps"

    $rtt = $rtt / $Samples
    $rounded = [math]::Round($rtt)
    Write-Host "Round-trip time`t`t:"$rounded "milliseconds"

    Write-Host "Samples`t`t`t:" $samples`
}

else {
    $querys = Get-Counter  -Counter "\VMware Blast(*)\Estimated throughput", "\VMware Blast(*)\Estimated fps", "\VMware Blast(*)\Estimated rtt", "\VMware Blast(*)\Estimated bandwidth" -SampleInterval 1 -MaxSamples $Samples
    foreach ($query in $querys) {
        $throughput += $query.CounterSamples[0].CookedValue
        $fps += $query.CounterSamples[1].CookedValue
        $rtt += $query.CounterSamples[2].CookedValue
        $bandwidth += $query.CounterSamples[3].CookedValue
    }
    Write-Host "__________________________________________________________________________"
    Write-Host "Average VMware Blast metrics for session: .::$sessionname::."
    Write-Host "--------------------------------------------------------------------------"

    $throughput = $throughput / $Samples / 1024 * 8
    $rounded = [math]::Round($throughput)
    Write-Host "Throughput`t`t:" $rounded "kbps"

    $bandwidth = $bandwidth / $Samples / 1024 * 8
    $rounded = [math]::Round($bandwidth)
    Write-Host "Estimated Bandwidth`t:"$rounded "kbps"

    $fps = $fps / $Samples
    $rounded = [math]::Round($fps)
    Write-Host "Frames per second`t:"$rounded "fps"

    $rtt = $rtt / $Samples
    $rounded = [math]::Round($rtt)
    Write-Host "Round-trip time`t`t:"$rounded "microseconds"

    Write-Host "Samples`t`t`t:" $samples`
}
