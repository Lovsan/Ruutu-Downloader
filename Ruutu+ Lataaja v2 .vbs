' Ruutu + lataaja. made by lovsan
' version 0.2
' updated @ 15.5 19.54 GMT +2
' Choose folder to download the files into, or stream the video with VLC player. working perfectly
' TODO - choose folder via promt[DONE], 
' TODO get mediainfo of downloaded file and write an NFO.
' TODO Also get more data from the xml file and also store that data to the NFO
' TODO - add VLC support, so just open the video in vlc instead downloading it. - done!!!
' TODO - something more to do maybe... will see...
dim gatlingurl, xmlid, fullurl, xmltag, xmltag2, colNodes, xmlCol2, description, info

MsgBox "Tervetuloa Ruutu+ Lataajaan. Kopioi numerosarja linkin lopusta, minka haluat ladata.", ,"Ruutu+ Lataaja" 

xmlid=inputbox("Ruutu+ ID.", "Ruutu + lataaja", 3223089 )
gatlingurl = "http://gatling.nelonenmedia.fi/media-xml-cache?id=" & xmlid
' https://www.ruutu.fi/video/3223089 - example 3223089
fullurl = gatlingurl
xmltag = "CastMediaFile"
' folder = "downloads" TODO
'vlcPath = "C:\Program Files (x86)\VideoLAN\VLC\vlc.exe" ' TODO, says file not found? path should be correct???
' ffmpegPath "" TODO

If xmlid = "" then wscript.Quit 1' if no ID typed, quit program
set xmlDoc = WScript.createobject("MSXML2.DOMDocument")
xmlDoc.async = "false"
xmlDoc.load (fullurl)

If xmlDoc.parseError <> 0 Then
'WScript.Echo xmlDoc.parseError.reason
WScript.Echo "Invalid xml ID -  Exiting Program."
WScript.Quit 1
End If

Set xmlCol = xmlDoc.getElementsByTagName(xmltag)
For Each Elem In xmlCol
'wscript.Echo(Elem.firstChild.nodeValue)
sFilePath = Elem.firstChild.nodeValue
Next

dim vlcd
vlcd=inputbox("type > stream = Open with vlc player,                          type > download = Download the file", "Ruutu+ Lataaja", "download")
If vlcd = "" then wscript.echo "Input Empty. Closing program" 'wscript.Quit 1'  if no ID typed, quit program
If vlcd = "stream" then wscript.Echo ("Starting stream with VLC Player")

Function BrowseForFolder()
  Dim oFolder
  Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0,"Select a Folder",0,0)' "/"
   If (oFolder Is Nothing) Then
     BrowseForFolder = Empty
   Else 
     BrowseForFolder = oFolder.Self.Path
   End If
End Function
Select Case vlcd

Case "download" 'Download the file using ffmpeg
Dim oFSO, oShell, sCommand, filext, folderPath
Dim xmlCol,sFilePath, sTempFilePath, vlc, filename
filext = ".mkv" ' you can use .ts .mp4 or .mkv
sTempFilePath=inputbox("Choose filename.", "Ruutu + lataaja")
 oFolder = BrowseForFolder() & "\"

FilePath = oFolder & sTempFilePath & filext 
'if (sFilePath Is Nothing) then wscript.Echo ("This video cant be downloaded or streamed - DRM active"), 1
	if sFilePath = Empty then
	wscript.Echo "This video cant be downloaded or streamed - DRM active"
	wscript.Quit 1
	end if
WScript.Echo "Saving To file ->" & FilePath

sCommand = "ffmpeg -y -i """ + sFilePath + """ -c:v copy -c:a copy """ + FilePath + """"
if FilePath = "" then wscript.Quit 1
Set oShell = WScript.CreateObject("WScript.Shell")
oShell.Run sCommand, 1, True
wscript.Echo "file download complete - " & FilePath

'Stream file with vlc player
case "stream"
set vlc = WScript.CreateObject("Wscript.Shell")
vlcCommand = "vlc.exe """ + sFilePath +""""
'wscript.Echo vlcCommand
wscript.Echo "Opening playlist - " & sFilePath
vlc.Run vlcCommand, 1, True
End Select
wscript.Echo "Thank you for using Ruutu+ lataaja!"
' show some info of the media after ffmpeg has executed? 'mediainfo? TODO - version 0.3

Set oShell = Nothing
Set vlc = Nothing
Set oFSO = Nothing
Set xmlCol = Nothing
Set xmlDoc = Nothing
