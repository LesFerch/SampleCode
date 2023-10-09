Set oFSO = CreateObject("Scripting.FileSystemObject")
InFile = "Eingabe.txt"
OutFile = "Ausgabe.txt"

aSearch = Array("  ","//","**")
aReplace = Array(" ","<i>","<b>")

Contents = oFSO.OpenTextFile(InFile).ReadAll

For i = 0 To UBound(aSearch)
  Do While Instr(Contents,aSearch(i))>0
    Contents = Replace(Contents,aSearch(i),aReplace(i))
  Loop
Next

oFSO.CreateTextFile(OutFile,True).Write(Contents)