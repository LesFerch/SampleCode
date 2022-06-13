'Sets the date added for all Edge IE Mode pages to 2099
'This causes the expiry dates to also be 2099

Silent = False
Const ForReading = 1
Const ForWriting = 2
Const Ansi = 0
DateAdded = "15741381600000000" '10/28/2099

Set oWSH = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")  
PrefsFile = oWSH.ExpandEnvironmentStrings("%LocalAppData%") & "\Microsoft\Edge\User Data\Default\" & "Preferences"

If Not Silent Then
  Response = MsgBox("Change expiry of all Edge IE Mode pages to 2099?",VBOKCancel)
  If Response=VBCancel Then WScript.Quit
End If

'Edge must be closed to modify the Preferences file
oWSH.Run "TaskKill /im MSEdge.exe /f",0,True

'Read contents of Edge Preferences file into a variable
Set oInput = oFSO.OpenTextFile(PrefsFile,ForReading)
Data = oInput.ReadAll
oInput.Close

'Find and change ever IE Mode page entry
'Possible enhancement: replace this loop with a regexp
StartPos = 1
Do
  FoundPos = InStr(StartPos,Data,"date_added")
  If FoundPos=0 Then Exit Do
  Data = Mid(Data,1,FoundPos + 12) & DateAdded & Mid(Data,FoundPos + 30)
  StartPos = FoundPos + 1
Loop

'Overwrite the Preferences file with the new data
Set oOutput = oFSO.OpenTextFile(PrefsFile,ForWriting,True,Ansi)
oOutput.Write Data
oOutput.Close

If Not Silent Then MsgBox "Done"
