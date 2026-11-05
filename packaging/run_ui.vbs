Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = folder
shell.Run Chr(34) & folder & "\NMIS-Automation.exe" & Chr(34), 1, False
