Set oWSH = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
MyFolder = oFSO.GetParentFolderName(WScript.ScriptFullName)
oWSH.CurrentDirectory = MyFolder
oWSH.Run "C:\Windows\SysWOW64\MSHTA.exe """ & MyFolder & "\TimeRemaining.hta""",1,False