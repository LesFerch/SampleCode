SetLocale("en-us")

Const ForReading = 1
Const ForWriting = 2
Const Ansi = 0

Set oWSH = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count<2 Then
  WScript.Echo "Usage: hexdump inputfile outputfile"
  WScript.Quit
End If

InputFile = WScript.Arguments.Item(0)
OutputFile = WScript.Arguments.Item(1)

If Not oFSO.FileExists(InputFile) Then
  WScript.Echo "Input file not found"
  WScript.Quit
End If

Set oInput = oFSO.OpenTextFile(InputFile,ForReading)
Set oOutput = oFSO.OpenTextFile(OutputFile,ForWriting,True,Ansi)

Do Until oInput.AtEndOfStream
  Char = oInput.Read(1)
  oOutput.Write Hex(Asc(Char))
Loop

oInput.Close
oOutput.Close