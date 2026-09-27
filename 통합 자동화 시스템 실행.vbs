Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strDir = fso.GetParentFolderName(WScript.ScriptFullName)
quote = Chr(34)
WshShell.Run quote & strDir & "\통합 자동화 시스템 실행.bat" & quote & " hide", 0, False
