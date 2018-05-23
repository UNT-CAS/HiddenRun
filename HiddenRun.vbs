' Usage:
'   wscript.exe //E:vbscript HiddenRun.vbs C:\Temp\CustomFoo.bat
'
' Will run your batch in a completely hidden command prompt, same as if you called it like this:
'   C:\Temp\CustomFoo.bat
'
' More Info: https://github.com/UNT-CAS/HiddenRun

Set oShell = CreateObject("Wscript.Shell")

Const LOG_EVENT_SUCCESS = 0
Const LOG_EVENT_ERROR = 1
Const LOG_EVENT_INFORMATION = 4

Dim iExitStatus : iExitStatus = LOG_EVENT_SUCCESS
Dim sArgs : sArgs = ""
Dim sMessage : sMessage = ""

For Each sArg in Wscript.Arguments
    If Len(sArg) > 0 Then
        sArgs = sArgs & " "
    End If
    
    If InStr(sArg, " ") > 0 Then
        ' If there's a space in the argument, wrap it in quotes.
        sArgs = sArgs & """" & sArg & """"
    Else
        sArgs = sArgs & sArg
    End If
Next

sMessage = "HiddenRun Running: " _
            & vbCrLf & vbTab & Wscript.ScriptFullName _
            & vbCrLf & vbTab & sArgs
oShell.LogEvent LOG_EVENT_INFORMATION, sMessage

iReturn = oShell.Run(sArgs, 0, True)

If iReturn <> 0 Then    
    iExitStatus = LOG_EVENT_ERROR
End If

sMessage = "HiddenRun Exited: " _
            & vbCrLf & vbTab & Wscript.ScriptFullName _
            & vbCrLf & vbTab & sArgs _
            & vbCrLf & vbTab & "Exit Code: " & iReturn
oShell.LogEvent iExitStatus, sMessage

Wscript.Quit iReturn
