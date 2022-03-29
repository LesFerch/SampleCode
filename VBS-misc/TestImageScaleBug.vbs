Set oWSH = CreateObject("Wscript.Shell")
Const wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
c = &hFF2D7D9A 'cool blue

ImageWidth = 800
ImageHeight = 600

Set v = CreateObject("WIA.Vector")
v.Add c
Set Img = v.ImageFile(1,1)
Set IP = CreateObject("WIA.ImageProcess")
IP.Filters.Add IP.FilterInfos("Scale").FilterID
IP.Filters(1).Properties("MaximumWidth") = ImageWidth
IP.Filters(1).Properties("MaximumHeight") = ImageHeight
IP.Filters(1).Properties("PreserveAspectRatio") = KeepAspect
IP.Filters.Add IP.FilterInfos("Convert").FilterID
IP.Filters(2).Properties("FormatID").Value = wiaFormatPNG
Set Img = IP.Apply(Img)

Img.SaveFile "test.png"
oWSH.Run "mspaint test.png"
