'Sub ParseXmlDocument()
Set xmlDoc = CreateObject("Msxml2.DOMDocument")
   'Dim success As Boolean

'xmlDoc.Load("http://gatling.nelonenmedia.fi/media-xml-cache?id=3193752")
'Dim xmlDoc As New MSXML2.DOMDocument30
Dim nodeBook 'As IXMLDOMElement 
Dim nodeId 'As IXMLDOMAttribute
'Dim sIdValue "" 'As String
xmlDoc.async = False
xmlDoc.Load ("http://gatling.nelonenmedia.fi/media-xml-cache?id=3193752")
If (xmlDoc.parseError.errorCode <> 0) Then
   Dim myErr
   Set myErr = xmlDoc.parseError
   MsgBox ("You have error " & myErr.reason)
Else
   Set nodeBook = xmlDoc.selectSingleNode("//Program")
   Set nodeId = nodeBook.getAttributeNode("description")
   'sIdValue = nodeId.xml
   wscript.Echo nodeId.value ' returns the program description! finally, took a while to figure this out!
End If
'End Sub
