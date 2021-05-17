Set WshShell = Wscript.CreateObject("Wscript.Shell")
Set FSO = Wscript.CreateObject("Scripting.FileSystemObject")

If NOT FSO.FileExists(WshShell.ExpandEnvironmentStrings("%allusersprofile%") & "\AppSense\Application Manager\Configuration\Configuration.aamp") Then
	Wscript.Echo "Application Control not installed."
	Wscript.Quit
End If

If FSO.FileExists(WshShell.ExpandEnvironmentStrings("%temp%") & "\Configuration.aamp") Then
	Wscript.Echo "Machine Unlocked."
	Wscript.Quit
End If

'Create the configuration
Dim Configuration
Set Configuration = CreateObject("AM.Configuration.5")

'Create the configuration helper
Dim ConfigurationHelper
Set ConfigurationHelper = CreateObject("AM.ConfigurationHelper.1")

Dim ConfigurationXml
ConfigurationXml = ConfigurationHelper.LoadLiveConfiguration

Configuration.ParseXML ConfigurationXml
ConfigurationHelper.SaveLocalConfiguration WshShell.ExpandEnvironmentStrings("%temp%") & "\Configuration.aamp", Configuration.Xml

'Load the default configuration
Configuration.ParseXML ConfigurationHelper.DefaultConfiguration

Configuration.DefaultRules.TrustedOwnershipChecking = False
Configuration.DefaultRules.ApplicationAccessEnabled = False
Configuration.DefaultRules.ANACEnabled = False

'Save the blank configuration to file.
ConfigurationHelper.SaveLiveConfiguration Configuration.Xml

Wscript.Echo "Machine Unlocked."
