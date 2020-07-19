dim sProcName

dim sPID 

dim iYesNo 

dim oShell

sProcName = WScript.Arguments.Item(0)

sPID = WScript.Arguments.Item(1)

iYesNo = msgbox("Your "& sProcName &" consumes excessive resources and may be interrupting the work of other users on this computer. Is it safe to abort it?",4,"ControlUp")
if iYesNo = 6 then
   Set oShell = CreateObject("WScript.Shell")
   oShell.Run "taskkill /PID "&sPID,0 , False
end if
