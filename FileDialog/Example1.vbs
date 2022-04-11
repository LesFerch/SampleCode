Set oWSH = CreateObject("Wscript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
oWSH.Run ".\FileDialog.exe Open ""*.m3u|*.m3u"" ThisPC ""~Select a playlist:""",0,True
File = oWSH.RegRead("HKCU\Software\FileDialog\")
MsgBox File