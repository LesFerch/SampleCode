Const DefaultWait = 2.5
L1 = "Enter amount of time to wait before shutting down"
L2 = "Enter a value in minutes or hours:minutes"
L3 = "0 = Shutdown immediately"
L4 = "Click X or Cancel to abort the shutdown"
Z = VBCRLF & VBCRLF
Set oWSH = CreateObject("Wscript.Shell")
Do
  Wait = InputBox(L1 & Z & L2 & Z & L3 & Z & L4,WScript.ScriptName,DefaultWait)
  If Wait="" Then WScript.Quit
  If IsNumeric(Replace(Wait,":",".")) Then Exit Do
Loop
If InStr(Wait,":")>0 Then
  aWait = Split(Wait,":")
  Wait = aWait(0)*60 + aWait(1)
End If
WScript.Sleep Wait * 60 * 1000
oWSH.Run "Shutdown.exe -r -f",1,False