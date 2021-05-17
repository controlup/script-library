Const HKEY_LOCAL_MACHINE =  &H80000002
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

If oReg.EnumKey(HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "", "") <> 0 Then
  Wscript.Echo "Environment Manager not installed."
  Wscript.Quit
End If

oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "LockdownLogging"
oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "LockdownLoggingProcesses"
oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "LockdownDebugLevel"
oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "FbrCorruptionChecking"
oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "EMLoaderDebugging"
oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "DebugPath"
oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager", "EmLoaderDebugLevel"

If oReg.EnumKey(HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "", "") = 0 Then
  oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "Enabled"
  oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "LogLevel"
  oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "LogSessionName"
  oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "LogFileName"
  oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "LogFileSizeLimitMB"
  oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "BufferSizeKB"
  oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "MaxBuffers"
  oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "MinBuffers"
  oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "LoggingMode"
  oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\AppSense\Environment Manager\EventTracing", "Components"
End If

Wscript.Echo "Environment Manager logging disabled."
