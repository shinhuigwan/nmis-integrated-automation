Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strDir = fso.GetParentFolderName(WScript.ScriptFullName)
quote = Chr(34)
WshShell.Run quote & strDir & "\run_ui.bat" & quote & " hide", 0, False
