Const wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
'&hFFFFFFFF will result in error: The Vector's Type is not compatible with this operation
'c = &h00FFFFFF 'white
'c = &hFF000000 'black
'c = &hFF0000FF 'blue
c = &hFF2D7D9A 'cool blue

ImageWidth = 1600
ImageHeight = 900
Pad = Int(0.1 * ImageWidth)

Set v = CreateObject("WIA.Vector")
For i = 1 To 400
  v.Add c
Next
Set Img = v.ImageFile(20,20)
Set IP = CreateObject("WIA.ImageProcess")
IP.Filters.Add IP.FilterInfos("Scale").FilterID
IP.Filters(1).Properties("MaximumWidth") = ImageWidth + Pad
IP.Filters(1).Properties("MaximumHeight") = ImageHeight + Pad
IP.Filters(1).Properties("PreserveAspectRatio") = KeepAspect
IP.Filters.Add IP.FilterInfos("Crop").FilterID
IP.Filters(2).Properties("Right") = Pad
IP.Filters(2).Properties("Bottom") = Pad
IP.Filters.Add IP.FilterInfos("Convert").FilterID
IP.Filters(3).Properties("FormatID").Value = wiaFormatPNG
Set Img = IP.Apply(Img)

Img.SaveFile ".\Test.png"
