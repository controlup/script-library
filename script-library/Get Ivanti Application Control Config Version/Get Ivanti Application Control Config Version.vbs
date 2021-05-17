strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSoftware = objWMIService.ExecQuery ("Select * from Win32_Product WHERE Name LIKE 'Ivanti Application Control Configuration%'")
For Each objSoftware in colSoftware
  Config = Replace(objSoftware.Caption, "Ivanti Application Control Configuration ", "")
  Wscript.Echo Config
Next
