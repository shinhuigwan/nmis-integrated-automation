Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strDir = fso.GetParentFolderName(WScript.ScriptFullName)
quote = Chr(34)
runtimePythonwPath = strDir & "\.venv_runtime\Scripts\pythonw.exe"
pythonwPath = strDir & "\.venv\Scripts\pythonw.exe"
scriptPath = strDir & "\nmis_slip_ui.py"

If fso.FileExists(runtimePythonwPath) Then
    WshShell.Run quote & runtimePythonwPath & quote & " " & quote & scriptPath & quote, 0, False
ElseIf fso.FileExists(pythonwPath) Then
    WshShell.Run quote & pythonwPath & quote & " " & quote & scriptPath & quote, 0, False
Else
    WshShell.Run quote & strDir & "\run_ui.bat" & quote & " hide", 0, False
End If
