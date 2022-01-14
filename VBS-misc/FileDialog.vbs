Set wShell=CreateObject("WScript.Shell")
Set oExec=wShell.Exec("mshta.exe ""about:<meta http-equiv=""X-UA-Compatible"" content=""IE=10""><input type=file id=FILE multiple accept=""video/*""><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
sFileSelected = oExec.StdOut.ReadLine
WScript.Echo sFileSelected