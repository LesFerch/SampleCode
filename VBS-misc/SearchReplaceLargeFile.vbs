Set oFSO = CreateObject("Scripting.FileSystemObject")
InFile = "Eingabe.txt"
OutFile = "Ausgabe.txt"

aSearch = Array("  ","//","**")
aReplace = Array(" ","<i>","<b>")

Set oInFile = oFSO.OpenTextFile(InFile)
Set oOutFile = oFSO.CreateTextFile(OutFile,True)

Do While Not oInFile.AtEndOfStream
  Line = oInFile.ReadLine
  For i = 0 To UBound(aSearch)
    Do While Instr(Line,aSearch(i))>0
      Line = Replace(Line,aSearch(i),aReplace(i))
    Loop
  Next
  oOutFile.WriteLine Line
Loop

oInFile.Close
oOutFile.Close