Option Explicit

Dim files, shell, baseDir, interpreter, entryPoint, command
Set files = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

baseDir = files.GetParentFolderName(WScript.ScriptFullName)
interpreter = files.BuildPath(baseDir, ".venv\Scripts\pythonw.exe")
entryPoint = files.BuildPath(baseDir, "app.pyw")

If Not files.FileExists(interpreter) Then
    MsgBox "The local Python environment was not found:" & vbCrLf & interpreter, vbCritical, "KeyFlip"
    WScript.Quit 1
End If
If Not files.FileExists(entryPoint) Then
    MsgBox "The application file was not found:" & vbCrLf & entryPoint, vbCritical, "KeyFlip"
    WScript.Quit 1
End If

shell.CurrentDirectory = baseDir
command = Chr(34) & interpreter & Chr(34) & " " & Chr(34) & entryPoint & Chr(34)
On Error Resume Next
shell.Run command, 0, False
If Err.Number <> 0 Then
    MsgBox "Cannot start KeyFlip:" & vbCrLf & Err.Description, vbCritical, "KeyFlip"
    WScript.Quit 1
End If
