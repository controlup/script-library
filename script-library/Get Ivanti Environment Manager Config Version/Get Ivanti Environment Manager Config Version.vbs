strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSoftware = objWMIService.ExecQuery ("Select * from Win32_Product WHERE Name LIKE 'Ivanti Environment Manager Configuration%'")
Config = ""
For Each objSoftware in colSoftware
  Config = Replace(objSoftware.Caption, "Ivanti Environment Manager Configuration ", "")
  Wscript.Echo Config
Next
If Config = "" Then
  Wscript.Echo "EM config not installed."
End If

