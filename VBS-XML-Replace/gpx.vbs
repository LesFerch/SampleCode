Set xmlDoc =  CreateObject("Microsoft.XMLDOM")
xmlFile = ".\gpx.xml"
xmlDoc.Load XmlFile

Set xmlNodes = xmlDoc.getElementsByTagName("desc")

For Each xmlNode in xmlNodes
  xmlNode.text = xmlNode.text * 15
Next

xmlDoc.Save xmlFile