' Ruutu + lataaja. made by lovsan
' version 0.2.3
' TODO get mediainfo of downloaded file and write an NFO.
' TODO Also get more data from the xml file and also store that data to the NFO [Work in progress]- finally manageded to read some data from xml, took a while to figure!
' TODO - add VLC support, so just open the video in vlc instead downloading it. - done!!!
' TODO - proper NFO file

dim gatlingurl, xmlid, fullurl, xmltag, xmltag2, colNodes, xmlCol2, description, info, mediaInfoPath, appVersion, github_url
appVersion = "0.2.3"
github_url = ""
MsgBox "Tervetuloa Ruutu+ Lataajaan. Kopioi numerosarja linkin lopusta, minka haluat ladata.", ,"Ruutu+ Lataaja" 

xmlid=inputbox("Ruutu+ ID.", "Ruutu + lataaja", 3223089 )
gatlingurl = "http://gatling.nelonenmedia.fi/media-xml-cache?id=" & xmlid
' https://www.ruutu.fi/video/3223089 - example 3223089
fullurl = gatlingurl
xmltag = "CastMediaFile"
xmltag2 = "Program"
'xmltag3 = "Program"
'xmltag4 = "Program"
' folder = "downloads" TODO
' vlcPath = "C:\Program Files (x86)\VideoLAN\VLC\vlc.exe " ' TODO, says file not found? path should be correct???
' ffmpegPath "" TODO
'mediaInfoPath = "C:\Program Files\Mediainfo\Mediainfo.exe"

If xmlid = "" then wscript.Quit 1' if no ID typed, quit program
set xmlDoc = WScript.createobject("MSXML2.DOMDocument")
xmlDoc.async = "false"
xmlDoc.load (fullurl)


Set xmlCol = xmlDoc.getElementsByTagName(xmltag)
For Each Elem In xmlCol
'wscript.Echo(Elem.firstChild.nodeValue)
sFilePath = Elem.firstChild.nodeValue

If xmlDoc.parseError <> 0 Then
'WScript.Echo xmlDoc.parseError.reason
WScript.Echo "Invalid xml ID -  Exiting Program.", ,"Ruutu+ Lataaja"
WScript.Quit 1
End If

' get show description
Dim nodeBook 'As IXMLDOMElement 
Dim nodeId, nodeID2 'As IXMLDOMAttribute
'Dim sIdValue "" 'As String

If (xmlDoc.parseError.errorCode <> 0) Then
   Dim myErr
   Set myErr = xmlDoc.parseError
   MsgBox ("You have error " & myErr.reason)
Else
   Set nodeBook = xmlDoc.selectSingleNode("//Program")
   Set nodeId = nodeBook.getAttributeNode("description")
   Set nodeId2 = nodeBook.getAttributeNode("program_name")
   'sIdValue = nodeId.xml
   
  ' comment out next 2 lines if you dont want to see the popups
   wscript.Echo nodeId2.value ' Display Show name and Episode
   wscript.Echo nodeId.value ' Display show description Once XML is loaded.
   description = nodeId.value
   program_name = nodeId2.value
End If
'Continue Program
Next
	if sFilePath = Empty then
	wscript.Echo "This video cant be downloaded or streamed - DRM active"
	wscript.Quit 1
end if
	
dim vlcd
vlcd=inputbox(" type > stream = Open with vlc player_type" & vbcrlf & " type > download = Download the file" & vbcrlf & " type > encode = Encode the file" & vbcrlf & " and Cancel to Quit.", "Ruutu+ Lataaja", "download")
If vlcd = "" then wscript.echo "Input Empty. Closing program" 'wscript.Quit 1'  if no ID typed, quit program
If vlcd = "stream" then wscript.Echo ("Starting stream with VLC Player")
if vlcd = "encode" then wscript.Echo ("Starting to encode the playlist.")

'create log
Function createLog()
Dim fso, outFile, nfoFile, fullurl, xItem, program, dldate, ruutuURL
Dim resolution, acodec, abitrate, dltime
nfoExt = ".nfo"
nfoFile = sTempFilePath & nfoExt 
ruutuURL = "https://www.ruutu.fi/video/" &xmlid
dldate = date() 
dltime = time()

resolution = "1080x720"
vcodec = "h264"
acodec = "aac"
abitrate = "128 Kb/s"
bitrate = "avg 3Mb/s"

'if sTempFilePath Empty then nfoFile = xmlid & nfoExt
'nfoFile=inputbox("Press no to cancel Nfo Creation", "Ruutu+ Lataaja","nfo name")& nfoExt
'if nfoFile is Empty then nfoFile = day() & nfoExt
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile(nfoFile, True)

' TODO - Figure how to get the show description and program name out of the xmlfile grrrr!!!!
' TODO - add some mediainfo into the .nfo
outFile.WriteLine("Program name: ") & program_name							
outFile.WriteLine("Description: ") & description
outFile.WriteLine("File Name: ") & sTempFilePath & filext
outFile.WriteLine("")
outFile.WriteLine("-----Mediainfo -----")
outFile.WriteLine("")
outFile.WriteLine("Resolution : ") & resolution
outFile.WriteLine("VideoCodec : ") & vcodec 'how to do empty lines otherways?
outFile.WriteLine("Bitrate: ") & bitrate
outFile.WriteLine("Audio Codec: ") & acodec
outFile.WriteLine("Audio Bitrate: ") & abitrate
outFile.WriteLine("")
outFile.WriteLine("---- End MediaInfo -------")
outFile.WriteLine("")
outFile.WriteLine("")
outfile.WriteLine("watch the show @ ") & ruutuURL
outFile.WriteLine("file downloaded at ") & dldate
outFile.WriteLine("Downloaded with Ruutu+ Lataaja")
outFile.WriteLine("https://github.com/Lovsan/Ruutu-Downloader")
outFile.WriteLine("version 0.2.2")
outFile.Close
wscript.Echo "Nfo File Created " & nfoFile
End Function

'Choose download folder
Function BrowseForFolder()
  Dim oFolder
  Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0,"Select a Folder",0,0)' "/"
   If (oFolder Is Nothing) Then
     BrowseForFolder = Empty
   Else 
     BrowseForFolder = oFolder.Self.Path
   End If
End Function
'WScript.Echo BrowseForFolder' oFolder.Self.Path '& "/"

Select Case vlcd

Case "download" 'Download the file using ffmpeg
Dim oFSO, oShell, sCommand, filext, folderPath,day

Dim xmlCol,sFilePath, sTempFilePath, vlc, filename
filext = ".mkv" ' you can use .ts .mp4 or .mkv
sTempFilePath=inputbox("Choose filename.", "Ruutu + lataaja")
 oFolder = BrowseForFolder() & "\"
if sTempFilePath = Empty then sTempFilePath = program_name 
FilePath = oFolder & sTempFilePath & filext 
WScript.Echo "Saving To file ->" & FilePath

sCommand = "ffmpeg -y -i """ + sFilePath + """ -c:v copy -c:a copy """ + FilePath + """"
if FilePath = "" then wscript.Quit 1
Set oShell = WScript.CreateObject("WScript.Shell")
oShell.Run sCommand, 1, True
wscript.Echo "file download complete - " & FilePath' , "Ruutu + Lataaja"
'msgbox "file download complete - " & FilePath , "Ruutu+ Lataaja" ', 1
'playFileonFinish()
' WriteNFO() TODO

'Stream file with vlc player
case "stream"
set vlc = WScript.CreateObject("Wscript.Shell")
'vlcCommand = "" + vlcPath + "" + sFilePath + """" ' not working yet
vlcCommand = "bin\vlc.exe """ + sFilePath +""""
'wscript.Echo vlcCommand
'wscript.Echo "Opening playlist - " & sFilePath
vlc.Run vlcCommand, 1, True

Case "encode"
Dim codec, bitrate, crf, encode, filext2, encodeFilePath, encodeFilePath2
filext2 = ".mkv" ' you can use .ts .mp4 or .mkv
'encodeCommand = "ffmpeg -y -i """ + sFilePath + """ -c:v """ + codec + """ -c:a copy """ + FilePath + """"
encodeFilePath=inputbox("Choose filename.", "Ruutu + lataaja")
 oFolder = BrowseForFolder() & "\"
if encodeFilePath = Empty then encodeFilePath = program_name
encodeFilePath2 = oFolder & encodeFilePath & filext2
 msgbox "To encode the stream, select the codec and bitrates", ,"Ruutu+ Lataaja" 
 
codec=inputbox("choose libx264, libx265 or libxvid" & vbcrlf & " (libx264 is set as default) ALSO Caution!" & vbcrlf & " Do not use libx255 (HEVC on slow machines)", ,"libx264")
 'wscript.Echo "Current coded is " & codec
'audiocode=inputbox("select aac, mp3 or ac3",,"ac3") ' aac,mp4,ac3
'audiobitrate=inputbox("128kbit/s is selected as default", , "128") ' 128, 256
bitrate=inputbox(" choose bitrate between 500-3000, default is 1000." & vbcrlf & " 3000 is avg when doing 720p." & vbcrlf & " Suggested bitrate for Hevc(libx265) is 500-100." & vbcrlf & " !!LEAVE THIS EMPTY TO USE CRF or press cancel",,"1000")
 'msgbox "Current bitrate is set to: " & bitrate
	if bitrate = Empty then crf=inputbox("choose crf value here, default is 18", ,"18")
'command to use if crf has been selected
encodeCommand2 = "bin/ffmpeg -y -i """ + sFilePath + """ -c:v """ + codec + """ -crf """ + crf + """ -c:a copy """ + encodeFilePath2 + """"
Set encode = Wscript.CreateObject("Wscript.Shell")
'command to use if bitrate has been selected
encodeCommand = "bin\ffmpeg -y -i """ + sFilePath + """ -c:v """ + codec + """ -bitrate """ + bitrate + """ -c:a copy """ + encodeFilePath2 + """" ' still broken somehow? or encoding is not possible?
'wscript.Echo encodeCommand

if crf = Empty then
encode.Run encodeCommand, 1, True
else
encode.Run encodeCommand2, 1,True
' playFileonFinish() TODO
' WriteNFO() TODO
End if
End Select
createLog() 'comment this line this disable creation of NFO file.
'playFile() ' uncomment this to play the file once download or stream is finished.
wscript.Echo "Thank you for using Ruutu+ lataaja!"
' show some info of the media after ffmpeg has finished? with ffmpeg or mediainfo?  TODO - version 0.3

set WriteNFO = Nothing
Set oShell = Nothing
Set vlc = Nothing
Set oFSO = Nothing
Set xmlCol = Nothing
Set xmlDoc = Nothing
'checkVersion() todo 0.2.5-0.3
'MsgBox download new version from & github_url
