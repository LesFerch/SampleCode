Set oWSH = CreateObject("Wscript.Shell")
CurDir = oWSH.CurrentDirectory
oWSH.Run "sc create ""ServiceName"" binPath=""" & CurDir & "\pathToExE"""
oWSH.Run "sc start ""ServiceName"""
