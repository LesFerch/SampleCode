'Sets the date added for all Edge IE Mode pages to 2099
'This causes the expiry dates to also be 2099
'How to use:
'1. Add your IE Mode pages in Microsoft Edge
'2. Close Microsoft Edge
'3. Run this script
'Repeat the above steps whenever you add more IE Mode pages

Silent = False
Const ForReading = 1
Const ForWriting = 2
Const Ansi = 0
Dim PrefsFile
DateAdded = "15741381600000000" '10/28/2099

Set oWSH = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
EdgeData = oWSH.ExpandEnvironmentStrings("%LocalAppData%") & "\Microsoft\Edge\User Data\"

If Not Silent Then
  Response = MsgBox("Change expiry of all Edge IE Mode pages to 2099?",VBOKCancel)
  If Response=VBCancel Then WScript.Quit
End If

'Edge must be closed to modify the Preferences file
oWSH.Run "TaskKill /im MSEdge.exe /f",0,True

Sub EditProfile
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
End Sub

PrefsFile = EdgeData & "Default\Preferences"
If oFSO.FileExists(PrefsFile) Then EditProfile

Do
  i = i + 1
  PrefsFile = EdgeData & "Profile " & i & "\Preferences"
  If Not oFSO.FileExists(PrefsFile) Then Exit Do
  EditProfile
Loop

If Not Silent Then MsgBox "Done"
