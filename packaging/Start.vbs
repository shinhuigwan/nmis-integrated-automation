Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
exe = folder & "\NMIS-Automation.exe"
If Not fso.FileExists(exe) Then
    MsgBox "Extract the entire ZIP before running Start.vbs.", 16, "NMIS"
    WScript.Quit 1
End If
shell.CurrentDirectory = folder
shell.Run Chr(34) & exe & Chr(34), 1, False
