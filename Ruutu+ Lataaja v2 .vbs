' Ruutu + lataaja. made by lovsan
' version 0.2.1
' Choose folder to download the files into, or stream the video with VLC player. working perfectly
' TODO - choose folder via promt[DONE], 
' TODO get mediainfo of downloaded file and write an NFO.
' TODO Also get more data from the xml file and also store that data to the NFO
' TODO - add VLC support, so just open the video in vlc instead downloading it. - done!!!
' TODO - something more to do maybe... will see...
dim gatlingurl, xmlid, fullurl, xmltag, xmltag2, colNodes, xmlCol2, description, info, mediaInfoPath

MsgBox "Tervetuloa Ruutu+ Lataajaan. Kopioi numerosarja linkin lopusta, minka haluat ladata.", ,"Ruutu+ Lataaja" 

xmlid=inputbox("Ruutu+ ID.", "Ruutu + lataaja", 3223089 )
gatlingurl = "http://gatling.nelonenmedia.fi/media-xml-cache?id=" & xmlid
' https://www.ruutu.fi/video/3223089 - example 3223089
fullurl = gatlingurl
xmltag = "CastMediaFile"
' folder = "downloads" TODO
' vlcPath = "C:\Program Files (x86)\VideoLAN\VLC\vlc.exe " ' TODO, says file not found? path should be correct???
' ffmpegPath "" TODO
mediaInfoPath = "C:\Program Files\Mediainfo\Mediainfo.exe"

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
	if sFilePath = Empty then
	wscript.Echo "This video cant be downloaded or streamed - DRM active"
	wscript.Quit 1
	end if
dim vlcd
vlcd=inputbox("type > stream = Open with vlc player,                       type > download = Download the file                      type > encode = Encode the file", "Ruutu+ Lataaja", "download")
If vlcd = "" then wscript.echo "Input Empty. Closing program" 'wscript.Quit 1'  if no ID typed, quit program
If vlcd = "stream" then wscript.Echo ("Starting stream with VLC Player")
if vlcd = "encode" then wscript.Echo ("Starting to encode the playlist.")
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
if sTempFilePath = Empty then sTempFilePath = xmlid
FilePath = oFolder & sTempFilePath & filext 
WScript.Echo "Saving To file ->" & FilePath

sCommand = "ffmpeg -y -i """ + sFilePath + """ -c:v copy -c:a copy """ + FilePath + """"
if FilePath = "" then wscript.Quit 1
Set oShell = WScript.CreateObject("WScript.Shell")
oShell.Run sCommand, 1, True
wscript.Echo "file download complete - " & FilePath' , "Ruutu + Lataaja"

'Stream file with vlc player
case "stream"
set vlc = WScript.CreateObject("Wscript.Shell")
'vlcCommand = "" + vlcPath + "" + sFilePath + """" ' not working yet
vlcCommand = "vlc.exe """ + sFilePath +""""
'wscript.Echo vlcCommand
'wscript.Echo "Opening playlist - " & sFilePath
vlc.Run vlcCommand, 1, True

Case "encode"
Dim codec, bitrate, crf, encode, filext2, encodeFilePath, encodeFilePath2
filext2 = ".mkv" ' you can use .ts .mp4 or .mkv
'encodeCommand = "ffmpeg -y -i """ + sFilePath + """ -c:v """ + codec + """ -c:a copy """ + FilePath + """"
encodeFilePath=inputbox("Choose filename.", "Ruutu + lataaja")
 oFolder = BrowseForFolder() & "\"
if encodeFilePath = Empty then encodeFilePath = xmlid
encodeFilePath2 = oFolder & encodeFilePath & filext2
 msgbox "To encode the stream, select the codec and bitrates", ,"Ruutu+ Encoder" 
 
codec=inputbox("choose libx264, libx265 or libxvid                           Caution! Do not use libx255 (HEVC on slow machines)")
 'wscript.Echo "Current coded is " & codec
 
bitrate=inputbox("choose bitrate between 0-3000, default is 1000 or press cancel to use crf,                           Suggested bitrate for Hevc(libx255) is 500.                            LEAVE THIS EMPTY TO USE CRF",,"1000")
 'wscript.Echo "Current bitrate is set to: " & bitrate
	if bitrate = Empty then crf=inputbox("choose crf value here, default is 18", ,"18")
'wscript.Echo crf
'command to use if crf has been selected
encodeCommand2 = "ffmpeg -y -i """ + sFilePath + """ -c:v """ + codec + """ -crf """ + crf + """ -c:a copy """ + encodeFilePath2 + """"
Set encode = Wscript.CreateObject("Wscript.Shell")
'command to use if bitrate has been selected
encodeCommand = "ffmpeg -y -i """ + sFilePath + """ -c:v """ + codec + """ -bitrate """ + bitrate + """ -c:a copy """ + encodeFilePath2 + """" ' still broken somehow? or encoding is not possible?
'wscript.Echo encodeCommand
if crf = Empty then
encode.Run encodeCommand, 1, True
else
encode.Run encodeCommand2, 1,True
' playFileonFinish() TODO
' WriteNFO() TODO
End Select
wscript.Echo "Thank you for using Ruutu+ lataaja!"
' show some info of the media after ffmpeg has finished? with ffmpeg or mediainfo?  TODO - version 0.3

set WriteNFO = Nothing
Set oShell = Nothing
Set vlc = Nothing
Set oFSO = Nothing
Set xmlCol = Nothing
Set xmlDoc = Nothing
