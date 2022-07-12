Const URL = "http://localhost:8090/api/ToggleTray?Traynr=1"
Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
Set oItems = oWMI.ExecQuery("Select * From Win32_PnPEntity")

Do
  For Each oItem in oItems
    If InStr(LCase(oItem.Description),"unknown")>0 Then
      Set oH = CreateObject("Msxml2.XMLHttp.6.0")
      oH.Open "GET", URL, False
      oH.Send
      If oH.Status = 200 Then
          'success
      Else
          'fail
      End If
      oH.Close
    End If
  Next
  WScript.Sleep 1000
Loop
