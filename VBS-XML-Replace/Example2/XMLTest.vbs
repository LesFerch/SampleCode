Set xmlDoc =  CreateObject("Microsoft.XMLDOM")
xmlFile = ".\XMLTest.xml"
xmlDoc.Load xmlFile

Set xmlNodes = xmlDoc.getElementsByTagName("Text")

For Each xmlNode in xmlNodes
  s = xmlNode.text
  x = InStr(s,"/")
  s = Mid(s,x+1)
  xmlNode.text = s
Next

xmlDoc.Save xmlFile