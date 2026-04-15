Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
exePath = fso.BuildPath(scriptDir, "DestinyStatusDesktop.exe")
If Not fso.FileExists(exePath) Then
    exePath = fso.BuildPath(fso.BuildPath(scriptDir, "dist"), "DestinyStatusDesktop.exe")
End If
shell.Run """" & exePath & """", 0, False
