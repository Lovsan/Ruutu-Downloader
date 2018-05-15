' Ruutu + Streamer. made by lovsan, feel free to made adjustments or what ever you want
' version 0.2 - only to stream with VLC player
' Currently stores the videos in the same folder where the script is located along with ffmpeg.exe and vlc cam also be used to stream the video
' TODO - choose folder via promt, after download is done, get mediainfo of downloaded file and write an NFO. Also get more data from the xml file and also store that data to the NFO
' TODO - make butttons to start stream or download.
dim gatlingurl, xmlid, fullurl, xmltag, xmltag2, colNodes, xmlCol2, description, info, vlc

'MsgBox "Ruutu Streamer. Kopioi numerosarja linkin lopusta, minka haluat toistaa.", "Ruutu Streamer" 

xmlid=inputbox("Ruutu+ ID.", "Ruutu Streamer")
gatlingurl = "http://gatling.nelonenmedia.fi/media-xml-cache?id=" & xmlid
' https://www.ruutu.fi/video/3223089 - example 3223089
fullurl = gatlingurl
xmltag = "CastMediaFile"
' vlcPath = "" TODO

If xmlid = "" then wscript.Quit 1' if no ID typed, quit program
set xmlDoc = WScript.createobject("MSXML2.DOMDocument")
xmlDoc.async = "false"
xmlDoc.load (fullurl)

If xmlDoc.parseError <> 0 Then ' Quits the program if invalid id has been typed.
WScript.Echo xmlDoc.parseError.reason
WScript.Quit 1
End If

Set xmlCol = xmlDoc.getElementsByTagName(xmltag)
For Each Elem In xmlCol
'wscript.Echo(Elem.firstChild.nodeValue)
'store playlist url in sFilePath
sFilePath = Elem.firstChild.nodeValue
'wscript.Echo sFilePath ' check if xml file contains the playlist.
if sFilePath = "" then wscript.Echo "This video cant streamed - DRM active"
'wscript.Echo "try new ID" - if faulty ID, promt for new
'loop
Next

set vlc = CreateObject("Wscript.Shell")
vlcCommand = "vlc.exe """ + sFilePath +""""
'if sFilePath = "" then wscript.quit 1
'wscript.Quit 1 ' no playlist found
'wscript.Echo vlcCommand
wscript.Echo "Stream Starting!" & sFilePath
vlc.Run vlcCommand, 1, True

Set oShell = Nothing
Set vlc = Nothing
Set oFSO = Nothing
Set xmlCol = Nothing
Set xmlDoc = Nothing
