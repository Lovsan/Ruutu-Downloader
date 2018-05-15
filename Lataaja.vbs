' Ruutu + lataaja. made by lovsan, feel free to made adjustments or what ever you want
' version 0.2
' Currently stores the videos in the same folder where the script is located along with ffmpeg.exe and vlc cam also be used to stream the video
' TODO - choose folder via promt, after download is done, get mediainfo of downloaded file and write an NFO. Also get more data from the xml file and also store that data to the NFO
' TODO - make butttons to start stream or download.
dim gatlingurl, xmlid, fullurl, xmltag, xmltag2, colNodes, xmlCol2, description, info

MsgBox "Tervetuloa Ruutu+ Lataajaan. Kopioi numerosarja linkin lopusta, minka haluat ladata.", vbYes ,"Ruutu+ Lataaja" 

xmlid=inputbox("Ruutu+ ID.", "Ruutu + lataaja", 3223089 )
gatlingurl = "http://gatling.nelonenmedia.fi/media-xml-cache?id=" & xmlid
' https://www.ruutu.fi/video/3223089 - example 3223089
fullurl = gatlingurl
xmltag = "CastMediaFile"
' xmltag2 = "" TODO
' xmltag3 = "" TODO
' folder = "downloads" TODO
' vlcPath = "" TODO
' ffmpegPath "" TODO

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
' check if xml file contains the playlist.
'if sFilePath = "" then wscript.Echo = "This video cant be downloaded or streamed - DRM active"
'wscript.Quit 1
'End If
Next

Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
sTempFilePath=inputbox("Choose filename.", "Ruutu + lataaja")

'Choose download folder
Function BrowseForFolder()
  Dim oFolder
  'Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0,"Select a Folder",0,0)' "/"
  Set oFolder = CreateObject("Shell.Application").BrowseForFolder( 0, "Select Folder", 0, myStartFolder )
   If (oFolder Is Nothing) Then
     BrowseForFolder = Empty
   Else 
     BrowseForFolder = oFolder.Self.Path
   End If
   On Error Goto 0
End Function
'WScript.Echo BrowseForFolder' oFolder.Self.Path '& "/"


Dim oFSO, oShell, sCommand, filext, folderPath
Dim xmlCol,sFilePath, sTempFilePath, vlc, vlcCommand
filext = ".mkv" ' you can use .ts .mp4 or .mkv

 oFolder = BrowseForFolder() & "\"
 'wscript.Echo oFolder

' If IsEmpty(folderPath) Then
 '  MsgBox "No Folder found."
' Else
 '  MsgBox folderPath
 'End If

FilePath = oFolder & sTempFilePath & filext 
if sFilePath = "" then wscript.Echo ("This video cant be downloaded or streamed - DRM active")
'FilePath = & oFolder & sTempFilePath & filext 'if sFilePath = "" then wscript.Echo ("This video cant be downloaded or streamed - DRM active")
'wscript.Quit 1 ' no playlist found, Quit the program
'if sTempFilePath = "" then wscript.Quit 1

WScript.Echo "Saving To file ->" & FilePath
sCommand = "ffmpeg -y -i """ + sFilePath + """ -c:v copy -c:a copy """ + FilePath + """"
Set oShell = WScript.CreateObject("WScript.Shell")
oShell.Run sCommand, 1, True
wscript.Echo "file download complete - " & FilePath
'Program Ends.
Set nfoShell = Nothing
Set oShell = Nothing
Set oFSO = Nothing
Set xmlCol = Nothing
Set xmlDoc = Nothing
