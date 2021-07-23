Set oShell = CreateObject("Wscript.Shell")
Set oShApp = CreateObject("Shell.Application")
Set oWMI = GetObject("winmgmts:\\.\root\cimv2")

Function ProcessExist(Exe,File)
  ProcessExist = False
  On Error Resume Next
  Set oProcesses = oWMI.ExecQuery("SELECT * FROM Win32_Process")
  For Each oProcess In oProcesses
    If InStr(oProcess.CommandLine,Exe) Then
      If InStr(oProcess.CommandLine,File) Then
        ProcessExist = True
        Exit Function
      End If
    End If
  Next
  On Error Goto 0
End Function

oShell.Run ".\MsgBox.vbs ""Message Here"" 0+48 ""Notification""",1,False

Do While ProcessExist("\WScript.exe","\MsgBox.vbs")
  oShApp.MinimizeAll
  WScript.Sleep 500
Loop
