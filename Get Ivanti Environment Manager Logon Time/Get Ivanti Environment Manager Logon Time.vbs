Dim EMUserStart
Dim Message
Dim EventTime
Dim TotalTime
Dim LogonTrigger
Dim WinLogon
Dim Userinit

Dim Offset
Offset = -5

Set WshShell = WScript.CreateObject("WScript.Shell")

SessionID = Wscript.Arguments.Item(0)
UserSID = GetSID(SessionID)
Username = WshShell.RegRead("HKEY_USERS\" & UserSID & "\Volatile Environment\USERNAME")
Userdomain = WshShell.RegRead("HKEY_USERS\" & UserSID & "\Volatile Environment\USERDOMAIN")

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'emuser.exe' AND SessionID=" & SessionID)

For Each objProcess in colProcessList
    dtmStartTime = objProcess.CreationDate
    strReturn = WMIDateStringToDate(dtmStartTime)
    EMUserStart = strReturn 
Next

Set colLoggedEvents = objWMIService.ExecQuery("Select * from Win32_NTLogEvent where LogFile = 'AppSense' AND User='" & UserDomain & "\\" & UserName & "' AND Message like '%PRE_DESKTOP%'")
For Each objEvent in colLoggedEvents
 Message = objEvent.Message
 EventTime = WMIDateStringToDate(objEvent.TimeGenerated)
 EventTime = DateAdd("h", Offset, EventTime)
 Exit For
Next

TotalTime = DateDiff("s", EMUserStart, EventTime)
LogonTrigger = Right(Message, Len(Message) - Instr(Message, "Duration") + 1)
LogonTrigger = Replace(LogonTrigger, "Duration: ", "")
LogonTrigger = Replace(LogonTrigger, ".", "")

If LogonTrigger = "" Then
wscript.echo "EM Logs not enabled. Please enable the 9662 event."
wscript.quit
End If

'wscript.echo "EMUser started at: " & EMUserStart
'wscript.echo "EMUser finished at: " & EventTime
wscript.echo "EM took " & TotalTime & " seconds at logon."
wscript.echo "The Pre-Desktop trigger took " & LogonTrigger & "."

Wscript.Quit 0


Function GetEMSessionVariable(name)
    GetEMSessionVariable = ""

    If (IsEmpty(sessionVariableReader)) Then
        Set sessionVariableReader = CreateObject("EMValue.EMGetValue")
    end if
        
    sessionVariableReader.Name = name    
    
    Dim errorCode
    errorCode = sessionVariableReader.Apply("")
    If errorCode = 0 Then
        GetEMSessionVariable = sessionVariableReader.Value
    end if        
End Function


Function WMIDateStringToDate(dtmStart)
    WMIDateStringToDate = CDate(Mid(dtmStart, 5, 2) & "/" & _
        Mid(dtmStart, 7, 2) & "/" & Left(dtmStart, 4) _
            & " " & Mid (dtmStart, 9, 2) & ":" & _
                Mid(dtmStart, 11, 2) & ":" & Mid(dtmStart, _
                    13, 2))
End Function

Function GetSessionID
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colProcess = objWMIService.ExecQuery ("Select * From Win32_Process")

For Each objProcess In colProcess
    If InStr (objProcess.CommandLine, WScript.ScriptName) <> 0 Then
      GetSessionID = objProcess.SessionID
    End If
Next
End Function


Function GetSID(SessionID)


Set WshShell = WScript.CreateObject("WScript.Shell")
Set Output = WshShell.Exec("query user " & SessionID)
Output.Stdout.ReadLine
Line = Output.Stdout.Readline
Line = Mid(Line, 2)
Username = Left(Line, Instr(Line, " ") - 1)


strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colProfiles = objWMIService.ExecQuery("Select * from win32_userprofile Where localpath like '%\\users\\" & Username & "'")

For Each objProfile in colProfiles
    GetSID = objProfile.SID
Next

End Function

