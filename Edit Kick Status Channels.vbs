Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = fso.BuildPath(scriptDir, "EditKickStatusConfig.ps1")
shell.Run "powershell.exe -ExecutionPolicy Bypass -NoProfile -File """ & scriptPath & """", 0, False
