Set oWSH = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
MyFolder = oFSO.GetParentFolderName(WScript.ScriptFullName)
oWSH.CurrentDirectory = MyFolder

Architecture = oWSH.environment("PROCESS").item("PROCESSOR_ARCHITECTURE")
SystemRoot = oWSH.ExpandEnvironmentStrings("%SystemRoot%")
WOWPath = SystemRoot & "\SysWOW64\"
ExePath = SystemRoot & "\System32\"
If Architecture="x86" And oFSO.FolderExists(WOWPath) Then ExePath = WOWPath

oWSH.Run ExePath & "MSHTA.exe """ & MyFolder & "\TimeRemaining.hta""",1,False