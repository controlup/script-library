$TargetName = "$env:COMPUTERNAME"

## To Do:
##  setup instance names to filter results when the sessions are on an RDS server. Filtering is not meaningful for VDI (workstations). This version is for VDI only.

function sGet-Wmi {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [Parameter(Mandatory = $false)]
        [string]$Namespace = "root\Cimv2",
        [Parameter(Mandatory = $true)]
        [string]$Class,
        [Parameter(Mandatory = $false)]
        $Property,
        [Parameter(Mandatory = $false)]
        $Filter
    )
    
    # Base string
    $wmiCommand = "gwmi -ComputerName $ComputerName -Namespace $Namespace -Class $Class -ErrorAction Stop"

    # If available, add Filter parameter
    if ($Filter)
    {
        # $Filter = ($Filter -join ',').ToString()
        $Filter = [char]34 + $Filter + [char]34
        $wmiCommand += " -Filter $Filter"
    }

    # If available, add Property parameter
    if ($Property)
    {
        $Property = ($Property -join ',').ToString()
        $wmiCommand += " -Property $Property"
    }

    # Try to connect
    $ResultCode = "1"
    Try
    {
        # $wmiCommand
        $wmiResult = iex $wmiCommand
    }
    Catch
    {
        $wmiResult = $_.Exception.Message
        $ResultCode = "0"
    }
    
    # If wmiResult is null
    if ($wmiResult -eq $null)
    {
        $wmiResult = "Result is null"
        $ResultCode = "2"
    }

    Return $wmiResult, $ResultCode
}

$PCOIPGeneralClass = "Win32_PerfRawData_TeradiciPerf_PCoIPSessionGeneralStatistics"
$PCOIPNetworkClass = "Win32_PerfRawData_TeradiciPerf_PCoIPSessionNetworkStatistics"
$PCOIPImagingClass = "Win32_PerfRawData_TeradiciPerf_PCoIPSessionImagingStatistics"
$PCOIPAudioClass = "Win32_PerfRawData_TeradiciPerf_PCoIPSessionAudioStatistics"
$PCOIPUSBClass = "Win32_PerfRawData_TeradiciPerf_PCoIPSessionUSBStatistics"

$GeneralPropertyList = @(
	"BytesReceived",
	"BytesSent",
	"PacketsReceived",
	"PacketsSent",
	"RXPacketsLost",
	"TXPacketsLost",
	"SessionDurationSeconds"
	)

$NetworkPropertyList = @(
	"RoundTripLatencyms",
    "TXBWLimitkbitPersec",
	"Timestamp_Sys100NS",
	"Frequency_Sys100NS"
	)

$ImagingPropertyList = @(
	"ImagingBytesReceived",
	"ImagingBytesSent"
	)

$AudioPropertyList = @(
	"AudioBytesReceived",
	"AudioBytesSent"
	)

$USBPropertyList = @(
	"USBBytesReceived",
	"USBBytesSent"
	)

#region Get initial values

$CompPCOIPGeneral1 = sGet-WMI -ComputerName $TargetName -Class $PCOIPGeneralClass -Property $GeneralPropertyList

If ($CompPCOIPGeneral1[1] -eq 1)
{
    $BytesReceived1 = $CompPCOIPGeneral1[0].BytesReceived
    $BytesSent1 = $CompPCOIPGeneral1[0].BytesSent
    $PacketsReceived1 = $CompPCOIPGeneral1[0].PacketsReceived
    $PacketsSent1 = $CompPCOIPGeneral1[0].PacketsSent
    $RXPacketsLost1 = $CompPCOIPGeneral1[0].RXPacketsLost
    $TXPacketsLost1 = $CompPCOIPGeneral1[0].TXPacketsLost
    $SessionDurationSeconds1 = $CompPCOIPGeneral1[0].SessionDurationSeconds
}
Else
{
    Write-Host "Error in retrieving initial general values. Game over."
    Write-Host "Result code: $CompPCOIPGeneral1[1]"
    Break
}

$CompPCOIPNetwork1 = sGet-WMI -ComputerName $TargetName -Class $PCOIPNetworkClass -Property $NetworkPropertyList

If ($CompPCOIPNetwork1[1] -eq 1)
{
    $Timestamp_Sys100NS1 = $CompPCOIPNetwork1[0].Timestamp_Sys100NS
    $Frequency_Sys100NS = $CompPCOIPNetwork1[0].Frequency_Sys100NS
}
Else
{
    Write-Host "Error in retrieving initial network values. Game over."
    Write-Host "Result code: $CompPCOIPNetwork1[1]"
    Break
}

$CompPCOIPImaging1 = sGet-WMI -ComputerName $TargetName -Class $PCOIPImagingClass -Property $ImagingPropertyList

If ($CompPCOIPImaging1[1] -eq 1)
{
    $ImagingBytesReceived1 = $CompPCOIPImaging1[0].ImagingBytesReceived
    $ImagingBytesSent1 = $CompPCOIPImaging1[0].ImagingBytesSent
}
Else
{
    Write-Host "Error in retrieving initial imaging values. Game over."
    Write-Host "Result code: $CompPCOIPImaging1[1]"
    Break
}

$CompPCOIPAudio1 = sGet-WMI -ComputerName $TargetName -Class $PCOIPAudioClass -Property $AudioPropertyList

If ($CompPCOIPAudio1[1] -eq 1)
{
    $AudioBytesReceived1 = $CompPCOIPAudio1[0].AudioBytesReceived
    $AudioBytesSent1 = $CompPCOIPAudio1[0].AudioBytesSent
}
Else
{
    Write-Host "Error in retrieving initial audio values. Game over."
    Write-Host "Result code: $CompPCOIPAudio1[1]"
    Break
}

$CompPCOIPUSB1 = sGet-WMI -ComputerName $TargetName -Class $PCOIPUSBClass -Property $USBPropertyList

If ($CompPCOIPUSB1[1] -eq 1)
{
    $USBBytesReceived1 = $CompPCOIPUSB1[0].USBBytesReceived
    $USBBytesSent1 = $CompPCOIPUSB1[0].USBBytesSent
}
Else
{
    Write-Host "Error in retrieving initial USB values. Game over."
    Write-Host "Result code: $CompPCOIPUSB1[1]"
    Break
}

#endregion Get initial values

###                        ###
##                          ##
#    getting delta values    #
##                          ##
###                        ###

#region Getting deltas

Start-Sleep -Seconds 1

$CompPCOIPGeneral2 = sGet-WMI -ComputerName $TargetName -Class $PCOIPGeneralClass -Property $GeneralPropertyList

If ($CompPCOIPGeneral2[1] -eq 1)
{
    $BytesReceived2 = $CompPCOIPGeneral2[0].BytesReceived
    $BytesSent2 = $CompPCOIPGeneral2[0].BytesSent
    $PacketsReceived2 = $CompPCOIPGeneral2[0].PacketsReceived
    $PacketsSent2 = $CompPCOIPGeneral2[0].PacketsSent
    $RXPacketsLost2 = $CompPCOIPGeneral2[0].RXPacketsLost
    $TXPacketsLost2 = $CompPCOIPGeneral2[0].TXPacketsLost
    $SessionDurationSeconds2 = $CompPCOIPGeneral2[0].SessionDurationSeconds
}
Else
{
    Write-Host "Error in retrieving General values. Game over."
    Write-Host "Result code: $CompPCOIPGeneral2[1]"
    Break
}

$CompPCOIPNetwork2 = sGet-WMI -ComputerName $TargetName -Class $PCOIPNetworkClass -Property $NetworkPropertyList

If ($CompPCOIPNetwork2[1] -eq 1)
{
    $RoundTripLatencyms = $CompPCOIPNetwork2[0].RoundTripLatencyms
    $TXBWLimitkbitPersec = $CompPCOIPNetwork2[0].TXBWLimitkbitPersec
    $Timestamp_Sys100NS2 = $CompPCOIPNetwork2[0].Timestamp_Sys100NS
}
Else
{
    Write-Host "Error in retrieving Network values. Game over."
    Write-Host "Result code: $CompPCOIPNetwork2[1]"
    Break
}

$CompPCOIPImaging2 = sGet-WMI -ComputerName $TargetName -Class $PCOIPImagingClass -Property $ImagingPropertyList

If ($CompPCOIPImaging2[1] -eq 1)
{
    $ImagingBytesReceived2 = $CompPCOIPImaging2[0].ImagingBytesReceived
    $ImagingBytesSent2 = $CompPCOIPImaging2[0].ImagingBytesSent
}
Else
{
    Write-Host "Error in retrieving Imaging values. Game over."
    Write-Host "Result code: $CompPCOIPImaging2[1]"
    Break
}

$CompPCOIPAudio2 = sGet-WMI -ComputerName $TargetName -Class $PCOIPAudioClass -Property $AudioPropertyList

If ($CompPCOIPAudio2[1] -eq 1)
{
    $AudioBytesReceived2 = $CompPCOIPAudio2[0].AudioBytesReceived
    $AudioBytesSent2 = $CompPCOIPAudio2[0].AudioBytesSent
}
Else
{
    Write-Host "Error in retrieving audio values. Game over."
    Write-Host "Result code: $CompPCOIPAudio2[1]"
    Break
}

$CompPCOIPUSB2 = sGet-WMI -ComputerName $TargetName -Class $PCOIPUSBClass -Property $USBPropertyList

If ($CompPCOIPUSB2[1] -eq 1)
{
    $USBBytesReceived2 = $CompPCOIPUSB2[0].USBBytesReceived
    $USBBytesSent2 = $CompPCOIPUSB2[0].USBBytesSent
}
Else
{
    Write-Host "Error in retrieving USB values. Game over."
    Write-Host "Result code: $CompPCOIPUSB2[1]"
    Break
}

$Time = (Get-Date).toLongTimeString()

#endregion Getting deltas

###                        ###
##                          ##
#      calculate values      #
##                          ##
###                        ###

#region Calculate values

If ($SessionDurationSeconds2 -eq $SessionDurationSeconds1 -and $SessionDurationSeconds2 -ne "0") {$SessionStatus = "ForceCloseCrash"}
ElseIf ($SessionDurationSeconds2 -eq "0") {$SessionStatus = "Closed"}
Else {
    $SessionStatus = "Normal"
}

$PerfTimeSec = ($Timestamp_Sys100NS2 - $Timestamp_Sys100NS1) / $Frequency_Sys100NS
$RXBWkbitPersec = ($BytesReceived2-$BytesReceived1) * 8 / (1024 * $PerfTimeSec)
$TXBWkbitPersec = ($BytesSent2-$BytesSent1) * 8 / (1024 * $PerfTimeSec)
$RXBWkbitPersecRnd = [Math]::round($RXBWkbitPersec,2)
$TXBWkbitPersecRnd = [Math]::round($TXBWkbitPersec,2)

$TXPacketsLost = $TXPacketsLost2 - $TXPacketsLost1
$RXPacketsLost = $RXPacketsLost2 - $RXPacketsLost1
$TXPacketLossPercent = [Math]::Round(($TXPacketsLost / ($TXPacketsLost + ($PacketsSent2-$PacketsSent1))) * 100,2)
$RXPacketLossPercent = [Math]::Round(($RXPacketsLost / ($RXPacketsLost + ($PacketsReceived2-$PacketsReceived1))) * 100,2)

$TXBWLimitkbitPersec = [Math]::Round($TXBWLimitkbitPersec,2)

$ImagingBytesSentKbitPerSec = [Math]::Round((($ImagingBytesSent2 - $ImagingBytesSent1) * 8 / (1024 * $PerfTimeSec)),2)
$ImagingBytesReceivedKbitPerSec = [Math]::Round((($ImagingBytesReceived2 - $ImagingBytesReceived1)* 8 / (1024 * $PerfTimeSec)),2)

$AudioBytesSentKbitPerSec = [Math]::Round((($AudioBytesSent2 - $AudioBytesSent1)* 8 / (1024 * $PerfTimeSec)),2)
$AudioBytesReceivedKbitPerSec = [Math]::Round((($AudioBytesReceived2 - $AudioBytesReceived1)* 8 / (1024 * $PerfTimeSec)),2)

$USBBytesSentKbitPerSec = [Math]::Round((($USBBytesSent2 - $USBBytesSent1)* 8 / (1024 * $PerfTimeSec)),2)
$USBBytesReceivedKbitPerSec = [Math]::Round((($USBBytesReceived2 - $USBBytesReceived1)* 8 / (1024 * $PerfTimeSec)),2)

#endregion Calculate values

###                        ###
##                          ##
#           output           #
##                          ##
###                        ###

#region Output

$MainOutput = "
PCoIP Bandwidth Usage Details

(Sample Period: $PerfTimeSec seconds)

Overall Bandwidth Usage:

    Bandwidth Utilization Tx: $TXBWkbitPersecRnd Kbit/sec
    Bandwidth Utilization Rx: $RXBWkbitPersecRnd Kbit/sec

Bandwidth Usage Breakdown:

    Imaging Bytes Tx: $ImagingBytesSentKbitPerSec Kbit/sec
    Imaging Bytes Rx: $ImagingBytesReceivedKbitPerSec Kbit/sec
    
    Audio Bytes Tx: $AudioBytesSentKbitPerSec Kbit/sec
    Audio Bytes Rx: $AudioBytesReceivedKbitPerSec Kbit/sec
    (does not include audio within USB data)

    USB Bytes Tx: $USBBytesSentKbitPerSec Kbit/sec
    USB Bytes Rx: $USBBytesReceivedKbitPerSec Kbit/sec
    (Zero Client only)

Advanced Statistics:

    Round Trip Latency: $RoundTripLatencyms ms
    Tx Bandwidth Limit: $TXBWLimitkbitPersec Kbit/sec
    
    Percentage Tx Packet Loss (during interval): $TXPacketLossPercent %
    Percentage RX Packet Loss (during interval): $RXPacketLossPercent %

    Tx Packet Loss (during interval): $TXPacketsLost packets
    Rx Packet Loss (during interval): $RXPacketsLost packets
"
If ($SessionStatus -eq "Normal") {
    $MainOutput
}
ElseIf ($SessionStatus -eq "Closed") {
    Write-Host "The PCoIP session just ended."
}
ElseIf ($SessionStatus -eq "ForceCloseCrash") {
    Write-Host "The PCoIP session just crashed or was forced closed."
}
Else {
    Write-Host "There was an error retrieving session statistics."
}

#endregion Output
