Set oWSH = CreateObject("Wscript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
oWSH.Run ".\FileDialog.exe Open ""*.m3u|*.m3u"" ThisPC ""Select one or more playlists:""",0,True
ItemList = oWSH.RegRead("HKCU\Software\FileDialog\ItemList")
ArrItemList = Split(ItemList,"|")
For i = 0 To UBound(ArrItemList)
  MsgBox ArrItemList(i)
Next