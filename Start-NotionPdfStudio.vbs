Option Explicit

Dim shell
Dim fso
Dim appRoot
Dim psScript
Dim command

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

appRoot = fso.GetParentFolderName(WScript.ScriptFullName)
psScript = fso.BuildPath(appRoot, "Start-NotionPdfStudio.ps1")
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & Chr(34) & psScript & Chr(34)

shell.Run command, 0, False
