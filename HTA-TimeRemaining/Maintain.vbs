Set oWSH = CreateObject("WScript.Shell")
UserSID = oWSH.RegRead("HKLM\Software\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\LastLoggedOnUserSID")
RegPath = "HKEY_USERS\" & UserSID & "\Software\TimeRemainingHTA\TimeLeftInSeconds"
oWSH.RegWrite "HKLM\Software\TimeRemainingHTA\TimeStamp\" & UserSID,"Start"

Sub ReStartHTA
  On Error Resume Next
  oWSH.RegWrite "HKLM\Software\TimeRemainingHTA\TimeStamp\" & UserSID,Now()
  If Err.Number=0 Then
    oWSH.Run "Schtasks.exe /run /tn TimeRemaining",1,True
    WScript.Sleep 5000
  End If
  On Error Goto 0
End Sub

Function GetTimeLeft()
  On Error Resume Next
  GetTimeLeft = oWSH.RegRead(RegPath)
  If Err.Number<>0 Then WScript.Quit 'User logged out
  On Error Goto 0
End Function

Do
  RegTimeLeft = GetTimeLeft()
  WScript.Sleep 2000
  x = GetTimeLeft()
  If RegTimeLeft=x And x>0 Then ReStartHTA
Loop
