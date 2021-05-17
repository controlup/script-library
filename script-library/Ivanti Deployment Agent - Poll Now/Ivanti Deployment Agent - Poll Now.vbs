'' Find the installation folder for the AppSense/Ivanti Deployment Agent/CCA and invoke ccacmd.exe to force it to poll the Management Center

Option Explicit

Dim WshShell, FSO , strUploadDir , strCCA , iStatus

Set WshShell = Wscript.CreateObject("Wscript.Shell")
Set FSO = Wscript.CreateObject("Scripting.FileSystemObject")

on error resume next
strUploadDir = WshShell.RegRead( "HKLM\SOFTWARE\AppSense Technologies\Communications Agent\upload dir" )
on error goto 0

If StrUploadDir = "" Then
  Wscript.echo "Error: failed to find installation folder, is the Deployment Agent installed?"
  Wscript.Quit 1
End If

' Remove any trailing \
If Right( strUploadDir , 1 ) = "\" Then
	strUploadDir = Left( strUploadDir , Len( strUploadDir ) - 1 )
End If

strCCA = Left( strUploadDir , InStrRev( strUploadDir , "\" )) & "ccacmd.exe"

If NOT Fso.FileExists( strCCA ) Then
  Wscript.Echo "Error: " & strCCA & " not found."
  Wscript.Quit 2
Else
  iStatus = WshShell.Run( """" & strCCA & """ /UpdateConfigs", 0, True )

  Wscript.Echo "Poll Now complete - status is " & iStatus
End If

