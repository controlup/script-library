Set WshShell = Wscript.CreateObject("Wscript.Shell")
Set FSO = Wscript.CreateObject("Scripting.FileSystemObject")

If NOT FSO.FileExists(WshShell.ExpandEnvironmentStrings("%allusersprofile%") & "\AppSense\Application Manager\Configuration\Configuration.aamp") Then
	Wscript.Echo "Application Control not installed."
	Wscript.Quit
End If

If NOT FSO.FileExists(WshShell.ExpandEnvironmentStrings("%temp%") & "\Configuration.aamp") Then
	Wscript.Echo "Machine Locked."
	Wscript.Quit
End If

'Create the configuration
Dim Configuration
Set Configuration = CreateObject("AM.Configuration.5")

'Create the configuration helper
Dim ConfigurationHelper
Set ConfigurationHelper = CreateObject("AM.ConfigurationHelper.1")

Dim ConfigurationXML
ConfigurationXML = ConfigurationHelper.LoadLocalConfiguration(WshShell.ExpandEnvironmentStrings("%temp%") & "\Configuration.aamp")

Configuration.ParseXML ConfigurationXml
ConfigurationHelper.SaveLiveConfiguration Configuration.Xml

FSO.DeleteFile WshShell.ExpandEnvironmentStrings("%temp%") & "\Configuration.aamp"

Wscript.Echo "Machine Locked."
