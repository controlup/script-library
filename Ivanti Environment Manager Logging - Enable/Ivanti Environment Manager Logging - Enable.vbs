Const HKEY_LOCAL_MACHINE =  &H80000002
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

If oReg.EnumKey(HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "", "") <> 0 Then
  Wscript.Echo "Environment Manager not installed."
  Wscript.Quit
End If

oReg.SetDWordValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "LockdownLogging", 0
oReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "LockdownLoggingProcesses", ""
oReg.SetDWordValue  HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "LockdownDebugLevel", 0
oReg.SetDWordValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "FbrCorruptionChecking", 0
oReg.SetDWordValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "EMLoaderDebugging", 0
oReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "DebugPath", ""
oReg.SetDWordValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "EmLoaderDebugLevel", 0

If oReg.EnumKey(HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "", "") <> 0 Then
  oReg.CreateKey HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing"
End If

oReg.SetDWordValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "Enabled", 1
oReg.SetDWordValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "LogLevel", 4
oReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "LogSessionName", "EM_ETW_Session"
oReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "LogFileName", "C:\Logs\EmTraceLog.etl"
oReg.SetDWordValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "LogFileSizeLimitMB", 512
oReg.SetDWordValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "BufferSizeKB", 256
oReg.SetDWordValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "MaxBuffers", 64
oReg.SetDWordValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "MinBuffers", 16
oReg.SetDWordValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "LoggingMode", 2
arrStringValues = Array("emcoreservice", "emuser", "emsystem", "emuserlogoff", "emloggedonuser", "emexit", "emauthpackage", "winlogon_notify_package", "winlogon_detour", "emcredentialmanager", "emlogoffuiapp", "emwow64", "emvirtualizationhost")
oReg.SetMultiStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "Components", arrStringValues
Wscript.Echo "Environment Manager logging enabled."
