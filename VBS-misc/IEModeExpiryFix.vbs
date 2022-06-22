'Sets the date added for all Edge IE Mode pages to any date you specify below
'This causes the expiry dates to be the specified date plus 30 days
'The default date added is a date in 2099, making the expiry a long way in the future

'How to use:
'1. Add your IE Mode pages in Microsoft Edge (or add them to the AddSites variable below)
'2. Close Microsoft Edge (if open)
'3. Run this script
'Repeat the above steps to add more IE Mode pages

Silent = False 'Change to True for no prompts
Setlocale("en-us") 'Locale setting must be consistent with the date format
DateAdded = "10/28/2099 10:00:00 PM" 'Specify the date here

'To add sites, uncomment and edit the AddSites line below. Separate each page entry with a |.
'Entries must end with a slash unless the URL ends with a file such as .html, .aspx, etc.
'AddSites = "http://www.fiat.it/|http://www.ferrari.it/"

'To find and replace a URL, uncomment and edit the FindReplace line below.
'Separate find and replace strings with a comma and separate each find/replace pair with a |.
'FindReplace = "http://localServer/register/login.aspx/,http://localserver/register/login.aspx"

Const ForReading = 1
Const ForWriting = 2
Const Ansi = 0
Dim PrefsFile

'Convert AddSites list to an array"
aAddSites = Split(AddSites,"|")
aFindReplace = Split(FindReplace,"|")

'Convert the date 
Set oDateTime = CreateObject("WbemScripting.SWbemDateTime")
Call oDateTime.SetVarDate(DateAdded,True)
EdgeDateAdded = Left(oDateTime.GetFileTime,17)

Set oWSH = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
EdgeData = oWSH.ExpandEnvironmentStrings("%LocalAppData%") & "\Microsoft\Edge\User Data\"

If Not Silent Then
  Response = MsgBox("Change expiry of all Edge IE Mode pages to:" & VBCRLF & VBCRLF & DateAdded & " + 30 days?",VBOKCancel)
  If Response=VBCancel Then WScript.Quit
End If

'Edge must be closed to modify the Preferences file
oWSH.Run "TaskKill /im MSEdge.exe /f",0,True

Sub EditProfile
  'Read contents of Edge Preferences file into a variable
  Set oInput = oFSO.OpenTextFile(PrefsFile,ForReading)
  Data = oInput.ReadAll
  oInput.Close

  'Exit if user is signed in
  If InStr(Data,"""account_info"":[]")=0 Then
    MsgBox ("Microsoft Edge sign-in detected." & VBCRLF & VBCRLF & "Sorry, this script only works with a local (signed-out) Edge profile.")
    WScript.Quit
  End If

  'Find and change ever IE Mode page entry
  'Possible enhancement: replace this loop with a regexp
  StartPos = 1
  Do
    FoundPos = InStr(StartPos,Data,"date_added")
    If FoundPos=0 Then Exit Do
    Data = Mid(Data,1,FoundPos + 12) & EdgeDateAdded & Mid(Data,FoundPos + 30)
    StartPos = FoundPos + 1
  Loop
  
  'Add any sites specified with the AddSites variable
  For i = 0 To UBound(aAddSites)
    AddSite = LCase(aAddSites(i))
    If Instr(AddSite,"://")=0 Then AddSite = "http://" & AddSite
    If Instr(Data,"user_list_data_1")=0 Then Data = Replace(Data,"},""edge"":{",",""user_list_data_1"":{}},""edge"":{")
    If Instr(Data,AddSite)=0 Then Data = Replace(Data,"""user_list_data_1"":{","""user_list_data_1"":{""" & AddSite & """:{""date_added"":""" & EdgeDateAdded & """,""engine"":2,""visits_after_expiration"":0},")
    Data = Replace(Data,"},}},","}}},")
  Next
  
  'Find and replace strings specified with the FindReplace variable
  For i = 0 To UBound(aFindReplace)
    aFindReplacePair = Split(aFindReplace(i),",")
    Data = Replace(Data,aFindReplacePair(0),aFindReplacePair(1))
  Next
  
  'Set "Allow sites to be reloaded in Internet Explorer mode" to "Allow"
  Data = Replace(Data,"{""ie_user""","{""enabled_state"":1,""ie_user""")
  Data = Replace(Data,"{""enabled_state"":0,""ie_user""","{""enabled_state"":1,""ie_user""")
  Data = Replace(Data,"{""enabled_state"":2,""ie_user""","{""enabled_state"":1,""ie_user""")

  'Overwrite the Preferences file with the new data
  Set oOutput = oFSO.OpenTextFile(PrefsFile,ForWriting,True,Ansi)
  oOutput.Write Data
  oOutput.Close
End Sub

PrefsFile = EdgeData & "Default\Preferences"
If oFSO.FileExists(PrefsFile) Then EditProfile

For Each oFolder In oFSO.GetFolder(EdgeData).SubFolders
  PrefsFile = oFolder.Path & "\Preferences"
  If oFSO.FileExists(PrefsFile) Then EditProfile
Next

If Not Silent Then MsgBox "Done"
