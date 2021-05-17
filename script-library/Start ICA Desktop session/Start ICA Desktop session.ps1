<# 
.SYNOPSIS
    Start an ICA session with the target computer
.PARAMETER TargetServer
   The computer that will be the remote session target. - automatically supplied by CU
.PARAMETER Fullscreen
   Go full screen or use the default 1024x768
#>

<#
Credits to: Joel Bennett for the Set-IniValue function
http://poshcode.org/160
#>

$TargetServer = $args[0]  # address of the target server - name or IP
$FullScreen = $args[1]    # ask "Fullscreen (Yes/No)?"     ^(?i)(yes|y|no|n)$

Function Set-IniValue($inifile,$section,$name,$value)
{
   $lines = gc $inifile
   $sections = select-string "^\[.*\]" $inifile
   $start,$end = 0,0
   for($l=0; $l -lt $sections.Count; ++$l){
      if($sections[$l].Line.Trim() -eq "[$section]") {
         $start = $sections[$l].LineNumber
         if($l+1 -ge $sections.Count) {
            $end = $lines.length-1;
         } else {
            $end = $sections[$l+1].LineNumber -2
         }
      }
   }
   
   if($start -and $end) {
      $done = $false
      for($l=$start;$l -le $end;++$l){
         if( $lines[$l] -match "^\s*$name\s*=" ) {
            $lines[$l] = "{0} = {1}" -f $name, $value
            $done = $true
            break;
         }
      }
      if(!$done) {
         $output = $lines[0..$start]
         $output += "{0} = {1}" -f $name, $value
         $output += $lines[($start+1)..($lines.Length-1)]
         $lines = $output
      }
   }
   Set-Content $inifile $lines
}

# let's replace the %CitrixServer% variable everywhere with the actual name and make a copy of the file elsewhere
Get-Content $Env:AppData\ControlUp\TemplateICAFile.ica | ForEach { $_ -Replace ("%CitrixServer%",$TargetServer) } | Set-Content $env:temp\sba-$TargetServer.ica -Encoding ascii

# set color depth and remove the shadow.exe program instruction (since we want a regular desktop)
Set-IniValue $env:temp\sba-$TargetServer.ica $TargetServer DesiredColor 24
Set-IniValue $env:temp\sba-$TargetServer.ica $TargetServer InitialProgram

If ($FullScreen -match "y")
{
    Set-IniValue $env:temp\sba-$TargetServer.ica $TargetServer DesiredHRes ([System.Int32]::MaxValue)
    Set-IniValue $env:temp\sba-$TargetServer.ica $TargetServer DesiredVRes ([System.Int32]::MaxValue)
} Else {
    Set-IniValue $env:temp\sba-$TargetServer.ica $TargetServer DesiredHRes 1024
    Set-IniValue $env:temp\sba-$TargetServer.ica $TargetServer DesiredVRes 768
}

# we're assuming that the .ica extension is associated with the Citrix Receiver here, and Windows
# will automatically start the proper program
Start-Process $env:temp\sba-$TargetServer.ica

