Option Explicit

Dim files, shell, baseDir, interpreter, entryPoint, command
Dim shortcut, shortcutPath, iconPath
Set files = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

baseDir = files.GetParentFolderName(WScript.ScriptFullName)
interpreter = files.BuildPath(baseDir, ".venv\Scripts\pythonw.exe")
entryPoint = files.BuildPath(baseDir, "app.pyw")

If Not files.FileExists(interpreter) Then
    MsgBox "The local Python environment was not found:" & vbCrLf & interpreter & vbCrLf & vbCrLf & _
           "Create it in this project folder:" & vbCrLf & _
           "py -3 -m venv .venv" & vbCrLf & _
           ".venv\Scripts\python.exe -m pip install -r requirements.txt", vbCritical, "KeyFlip"
    WScript.Quit 1
End If
If Not files.FileExists(entryPoint) Then
    MsgBox "The application file was not found:" & vbCrLf & entryPoint, vbCritical, "KeyFlip"
    WScript.Quit 1
End If

' A .lnk stores machine-specific paths. Rebuild the local shortcut from this script's folder.
shortcutPath = files.BuildPath(baseDir, "KeyFlip.lnk")
iconPath = files.BuildPath(baseDir, "resources\keyflip.ico")
On Error Resume Next
Set shortcut = shell.CreateShortcut(shortcutPath)
shortcut.TargetPath = files.BuildPath(files.GetParentFolderName(WScript.FullName), "wscript.exe")
shortcut.Arguments = Chr(34) & WScript.ScriptFullName & Chr(34)
shortcut.WorkingDirectory = baseDir
If files.FileExists(iconPath) Then shortcut.IconLocation = iconPath & ",0"
shortcut.Save
Err.Clear
On Error GoTo 0

shell.CurrentDirectory = baseDir
command = Chr(34) & interpreter & Chr(34) & " " & Chr(34) & entryPoint & Chr(34)
On Error Resume Next
shell.Run command, 0, False
If Err.Number <> 0 Then
    MsgBox "Cannot start KeyFlip:" & vbCrLf & Err.Description, vbCritical, "KeyFlip"
    WScript.Quit 1
End If
