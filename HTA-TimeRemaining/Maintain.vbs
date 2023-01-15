Set oWSH = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWMI = GetObject("winmgmts:\root\cimv2")

MyFolder = oFSO.GetParentFolderName(WScript.ScriptFullName)
oWSH.CurrentDirectory = MyFolder
Dim Value
UserSID = "": UserName = ""
Do
  WScript.Sleep 1000
  Set colProcessList = oWMI.ExecQuery("Select Name from Win32_Process where Name = 'Explorer.exe'")
  For Each oProcess in colProcessList
    oProcess.GetOwnerSID UserSID
    oProcess.GetOwner UserName
  Next
  oWSH.RegWrite "HKLM\Software\TimeRemainingHTA\UserSID",UserSID
Loop Until UserSID<>""

RegPath = "HKEY_USERS\" & UserSID & "\Software\TimeRemainingHTA\TimeLeftInSeconds"

Sub ReStartHTA
  On Error Resume Next
  oWSH.RegWrite "HKLM\Software\TimeRemainingHTA\TimeStamp",Now()
  If Err.Number=0 Then oWSH.Run "Schtasks.exe /run /tn TimeRemaining",1,True
  WScript.Sleep 5000
  On Error Goto 0
End Sub

Sub GetTimeLeft
  On Error Resume Next
  Value = oWSH.RegRead(RegPath)
  If Err.Number<>0 Then WScript.Quit 'User logged out
  On Error Goto 0
End Sub

Do
  GetTimeLeft
  If Value<>"" Then
    RegTimeLeft = Value
    oWSH.RegWrite RegPath,""
  End If
  WScript.Sleep 2000
  GetTimeLeft
  If Value="" Then
    oWSH.RegWrite RegPath,RegTimeLeft
    ReStartHTA
  End If
Loop
