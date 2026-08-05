Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonwPath = strDir & "\.venv\Scripts\pythonw.exe"
scriptPath = strDir & "\nmis_slip_ui.py"

If fso.FileExists(pythonwPath) Then
    WshShell.Run """" & pythonwPath & """ """ & scriptPath & """", 0, False
Else
    WshShell.Run """" & strDir & "\통합 자동화 시스템 실행.bat""", 0, False
End If
