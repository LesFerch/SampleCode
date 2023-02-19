'Sets the date added for all Edge IE Mode pages to any date you specify below
'This causes the expiry dates to be the specified date plus 30 days
'The default date added is a date in 2099, making the expiry a long way in the future

'This script only works with completely local Edge profiles. It will not work if Edge is signed-in.

'Optionally run via CScript (i.e. CScript IEModeExpiryFix.vbs) to get console output instead of message boxes.

'Note for system administrators: If your computers are in Active Directory, please consider
'using the Enterprise Mode Site List instead of this script. See this link:
'https://docs.microsoft.com/en-us/internet-explorer/ie11-deploy-guide/what-is-enterprise-mode

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
'The domain part of the entry must be all lowercase.
'AddSites = "http://www.fiat.it/|http://www.ferrari.it/"

'To find and replace a URL, uncomment and edit the FindReplace line below.
'Separate find and replace strings with a comma and separate each find/replace pair with a |.
'FindReplace = "http://localServer/register/login.aspx/,http://localserver/register/login.aspx"

Const ForReading = 1
Const ForWriting = 2
Const Ansi = 0
Dim PrefsFile,MyLog

'Convert AddSites and FindReplace lists to arrays"
aAddSites = Split(AddSites,"|")
aFindReplace = Split(FindReplace,"|")

'Convert the date 
Set oDateTime = CreateObject("WbemScripting.SWbemDateTime")
Call oDateTime.SetVarDate(DateAdded,True)
EdgeDateAdded = Left(oDateTime.GetFileTime,17)

Set oWSH = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

CScript = InStr(LCase(WScript.FullName),"cscript")>0

If Not Silent And Not CScript Then
  Response = MsgBox("Change expiry of all Edge IE Mode pages to:" & VBCRLF & VBCRLF & DateAdded & " + 30 days?",VBOKCancel)
  If Response=VBCancel Then WScript.Quit
End If

'Edge must be closed to modify the Preferences file
oWSH.Run "TaskKill /im MSEdge.exe /f",0,True

LocalAppData = oWSH.ExpandEnvironmentStrings("%LocalAppData%")

LogMsg "Profiles processed:"

ProcessProfiles("Edge") 'For released Edge profile
ProcessProfiles("Edge Beta") 'For Beta Edge profile
ProcessProfiles("Edge Dev") 'For Dev Edge profile
ProcessProfiles("Edge SxS") 'For Canary Edge profile

Sub LogMsg(Msg)
  MyLog = MyLog & Msg & VBCRLF & VBCRLF
End Sub

Function BadURL(ByVal URL)
  URL = LCase(URL)
  NoDoubleSlash = Instr(URL,"://")=0
  TooManySlashes = Instr(URL,"http:///")>0 Or Instr(URL,"https:///")>0 Or Instr(URL,"file:////")>0
  BadURL = NoDoubleSlash Or TooManySlashes
End Function

Function FixURL(byVal URL)
  URL = URL & " "
  For i = 2 To Len(URL)
    If Mid(URL,i,1)="/" And Mid(URL,i-1,1)<>"/" And Mid(URL,i+1,1)<>"/" Then Exit For
  Next
  URL = Trim(URL)
  If i>Len(URL) Then URL = URL & "/"
  FixURL = LCase(Left(URL,i)) & Mid(URL,i+1)
End Function

Sub EditProfile
  'Read contents of Edge Preferences file into a variable
  Set oInput = oFSO.OpenTextFile(PrefsFile,ForReading)
  Data = oInput.ReadAll
  oInput.Close

  LogMsg PrefsFile

  'Exit if user is signed in
  If InStr(Data,"""account_info"":[]")=0 And InStr(Data,"account_info")>0 Then
    LogMsg "Edge profile sign-in detected. Profile cannot be updated."
    Exit Sub
  End If

  OriginalData = Data

  'Find and change every IE Mode page entry
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
    If Not BadURL(aAddSites(i)) Then
      AddSite = FixURL(aAddSites(i))
      If Instr(Data,"user_list_data_1")=0 Then Data = Replace(Data,"},""edge"":{",",""user_list_data_1"":{}},""edge"":{")
      If Instr(Data,AddSite)=0 Then Data = Replace(Data,"""user_list_data_1"":{","""user_list_data_1"":{""" & AddSite & """:{""date_added"":""" & EdgeDateAdded & """,""engine"":2,""visits_after_expiration"":0},")
      Data = Replace(Data,"},}},","}}},")
    End If
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
  If Data<>OriginalData Then
    LogMsg "Profile updated"
    Set oOutput = oFSO.OpenTextFile(PrefsFile,ForWriting,True,Ansi)
    oOutput.Write Data
    oOutput.Close
  Else
    LogMsg "Profile already updated"
  End If
End Sub

Sub ProcessProfiles(ProfileFolder)
  EdgeData = LocalAppData & "\Microsoft\" & ProfileFolder & "\User Data\"
  If oFSO.FolderExists(EdgeData) Then
    For Each oFolder In oFSO.GetFolder(EdgeData).SubFolders
      PrefsFile = oFolder.Path & "\Preferences"
      If oFSO.FileExists(PrefsFile) Then EditProfile
    Next
  End If
End Sub

If Not Silent Then
  WScript.Echo MyLog
End If