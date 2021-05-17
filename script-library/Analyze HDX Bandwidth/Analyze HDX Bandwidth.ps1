
<#  
.SYNOPSIS
      This script shows th network bandwidth used by a given HDX session.
.DESCRIPTION
      This script measures the bandwidth of a given active HDX session, and breaks down the bandwidth consumption into the most useable ICA virtual channels.
      The output shows the bandwidth usage in kbps (kilobit per second) of each virtual channel and the total session.
      This version shows only the session output (download), and not the upload.
.PARAMETER 
      This script has 3 parameters:
      ServerName - The target server that the script should run on.
      SessionName - the name of the session from the ControlUp Console(e.g. ICA-TCP#1, ICA-CGP#2)
      UserName - the user name from the console. (e.g. controlup\matan). 
      The 2 parameters create the session name like the Get-Counter command requires.
.EXAMPLE
        ./AnalyzeHDXbandwidth.ps1 "ICA-TCP#2" "controlup\matan" "cuxen65ts03"
.OUTPUTS
        A list of the measured virtual channels with the bandwidth consumption in kbps.
.LINK
        See http://www.ControlUp.com
#>

#Defininig all the paramerters from the console

$originalSession = $args[0].ToString().Replace("#" , " ")
if($originalSession.StartsWith("RDP")){
    Write-Host "This is an RDP session, Please re-run the script against an HDX session"
    exit 1
}

$username = $args[1].ToString().Split("\")
$servername = $args[2]
#######################
#defining the correct session name to the Get-Counter command session naming convenstion
$correctUserName = $username[1]
$sessionname = "$originalSession ($correctUserName)"
$Samples = 10
$queries = Get-Counter -ComputerName $servername -Counter "\ICA Session($sessionname)\Output ThinWire Bandwidth","\ICA Session($sessionname)\Output Audio Bandwidth", "\ICA Session($sessionname)\Output TWAIN Bandwidth","\ICA Session($sessionname)\Output COM Bandwidth","ICA Session($sessionname)\Output Drive Bandwidth", "\ICA Session($sessionname)\Output Printer Bandwidth" , "\ICA Session($sessionname)\Output Clipboard Bandwidth", "\ICA Session($sessionname)\Output Session Bandwidth" -SampleInterval 1 -MaxSamples $Samples
$ThinWire = 0
$Audio = 0 
$TWAIN = 0
$COMPORT = 0
$Drive = 0
$printer = 0
$ClipBoard = 0
$total = 0
foreach($querie in $queries){
    $ThinWire += $querie.CounterSamples[0].CookedValue
    $Audio += $querie.CounterSamples[1].CookedValue
    $TWAIN +=  $querie.CounterSamples[2].CookedValue
    $COMPORT +=  $querie.CounterSamples[3].CookedValue
    $Drive +=  $querie.CounterSamples[4].CookedValue
    $printer +=  $querie.CounterSamples[5].CookedValue
    $ClipBoard +=  $querie.CounterSamples[6].CookedValue
    $total +=  $querie.CounterSamples[7].CookedValue
}
Write-Host "__________________________________________________________________________"
Write-Host "Average Download Bandwidth for session: .::$sessionname::."
Write-Host "--------------------------------------------------------------------------"
$ThinWire = $ThinWire / $Samples / 1024
$rounded = [math]::Round($ThinWire)
Write-Host "Thinwire Bandwidth`t`t:" $rounded "kbps"
$Audio = $Audio / $Samples / 1024
$rounded = [math]::Round($Audio)
Write-Host "Audio Bandwidth`t`t`t:"$rounded "kbps"
$TWAIN = $TWAIN / $Samples / 1024
$rounded = [math]::Round($TWAIN)
Write-Host "TWAIN Devices Bandwidth`t`t:"$rounded "kbps"
$COMPORT = $COMPORT / $Samples / 1024
$rounded = [math]::Round($COMPORT)
Write-Host "COM Port Bandwidth`t`t:"$rounded "kbps"
$Drive = $Drive / $Samples / 1024
$rounded = [math]::Round($Drive)
Write-Host "Drive Bandwidth`t`t`t:"$rounded "kbps"
$printer = $printer / $Samples / 1024
$rounded = [math]::Round($printer)
Write-Host "Printers Bandwidth`t`t:"$rounded "kbps"
$ClipBoard = $ClipBoard / $Samples / 1024
$rounded = [math]::Round($ClipBoard)
Write-Host "Clipboard Bandwidth`t`t:"$rounded "kbps"
$total = $total / $Samples / 1024
$rounded = [math]::Round($total)
Write-Host "--------------------------------------------------------------------------"
Write-Host "Total Session Bandwidth`t`t:"$rounded "kbps"
Write-Host "Samples`t`t`t`t:" $samples
