# Ruutu+ Downloader
Ruutu + Downloader
version 0.2.3 - last updated 16.5.2018

The current version can download, stream or encode the Ruutu+ Content, the App will simple quide you on each step what to do.
Encode option is not really suggested unless you really want to download the content with other Codec and Settings.
Download option will automatically get the best version of the Ruutu+ content, which is 720p and bitrate around 3Mb/s

Usage:

Go to Ruutu.fi with browser of your choise, and open the Show you want to get in new tab, 
then copy the numbers from the end of the link 'https://www.ruutu.fi/video/3224241'
this is what you need : 3224241

open the app(ruutu+ Lataaja.vbs) and paste the numbers in it. it will then ask you to download, stream or encode
If you just want to get copy of the file, press ok as 'download' is preselected.It will then prompt your for filename then folder, if filenames is left empty, it will save the file as 3224241.mkv, then it will then launch ffmpeg and get the file for you.
If you select 'stream' it will simply launch the Show in vlc player, you can simply fastforward/reverse and so on.

If you want to encode. follow directions in app, i will write better quide here and add more options in the app also.


vlc needed to be installed and vlc.exe copied into same folder where script is located, same goes with ffmpeg.exe
without those files the script cannot fucnction.
