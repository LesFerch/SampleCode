set oFSO = CreateObject("Scripting.FileSystemObject")
InFile = "Eingabe.txt"
OutFile = "Ausgabe.txt"

aSearch = Array("  ","//","**")
aReplace = Array(" ","<i>","<b>")

Contents = oFSO.OpenTextFile(InFile).ReadAll

For i = 0 To UBound(aSearch)
  Contents = Replace(Contents,aSearch(i),aReplace(i))
Next

oFSO.CreateTextFile(OutFile,True).Write(Contents)