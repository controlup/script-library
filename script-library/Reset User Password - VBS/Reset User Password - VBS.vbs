On Error Resume Next

Set objStdErr = WScript.StdErr

If wscript.arguments(1) <> wscript.arguments(2) Then
    objStdErr.Write("The password and its confirmation do not match. Please try again.")
    Wscript.Quit(1)
Else
    ' Constants for the NameTranslate object.
    Const ADS_NAME_INITTYPE_GC = 3
    Const ADS_NAME_TYPE_NT4 = 3
    Const ADS_NAME_TYPE_1779 = 1

    strUserAccount = wscript.arguments(0)

    ' Use the NameTranslate object to convert the NT user name to the
    ' Distinguished Name required for the LDAP provider.
    Set objTrans = CreateObject("NameTranslate")
    objTrans.Init ADS_NAME_INITTYPE_GC, ""
    objTrans.Set ADS_NAME_TYPE_NT4, strUserAccount
    strUserDN = objTrans.Get(ADS_NAME_TYPE_1779)

    ' Escape any "/" characters with backslash escape character.
    ' All other characters that need to be escaped will be escaped.
    strUserDN = Replace(strUserDN, "/", "\/")

    ' Reset a User's password

    Set objUser = GetObject("LDAP://" & strUserDN)

    objUser.SetPassword wscript.arguments(1)
	
    If Err.Number <> 0 Then
        objStdErr.Write("Error in setting password: " & vbCrLf)
        objStdErr.Write("Source: " & Err.Source & vbCrLf)
        objStdErr.Write("Description: " & Err.Description & vbCrLf)
        Err.Clear
        wscript.Quit(1)
    End If

    Wscript.Echo "The password was changed."
End If

