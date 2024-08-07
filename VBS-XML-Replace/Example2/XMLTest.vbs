Set xmlDoc =  CreateObject("Microsoft.XMLDOM")
xmlFile = ".\XMLTest.xml"
xmlDoc.Load xmlFile

Set xmlNodes = xmlDoc.getElementsByTagName("Text")

For Each xmlNode in xmlNodes
  s = xmlNode.text
  x = InStr(s,"/")
  x = Mid(s,x+1)
  xmlNode.text = x
Next

xmlDoc.Save xmlFile