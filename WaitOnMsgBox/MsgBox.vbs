On Error Resume Next
Message = WScript.Arguments.Item(0)
Icon    = Eval(WScript.Arguments.Item(1))
Title   = WScript.Arguments.Item(2)
On Error Goto 0
MsgBox Message,Icon,Title