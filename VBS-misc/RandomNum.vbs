Dim max,min,rand
max=5
min=1
Do
  Randomize
  rand = Int((max-min+1)*Rnd+min)
  RetVal = MsgBox(rand,vbOKCancel )
  If RetVal=vbCancel Then Exit Do
Loop