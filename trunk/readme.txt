Title: PDFCreator Version 0.7.0
Authors: Philip Chinery, Frank Heindörfer
Releasedate: 10.06.2003

Look at the readme.txt-Files in 'Setup', 'Printer', 'Printer\Redmon' and 'Setup\Upx'.

Necessary additional files:

IPDK:
	Install the IPDK files to use error messages in your native language.
	http://msdn.microsoft.com/vbasic/downloads/tools/ipdk.aspx

Systemfiles: 
	Run Win9x_CopySystemfiles.bat or WinNt_CopySystemfiles.bat from 'Additional files\Systemfiles' to copy the Systemfiles for the setup.

Ghostscript Files:
	Download Ghostscript 8.00
	Install Ghostscript in the standarddirectory c:\gs
	Run Win9x_CopyGhost.bat or WinNt_CopyGhost.bat from 'Additional files\Ghostscript' to copy the necessary Gohstscript-files
	You can delete this files with Win9x_DelGhost.bat or WinNt_DelGhost.bat

Redmon Files:
	Download the Reddmon-Files from http://www.cs.wisc.edu/~ghost/redmon/ (ftp://mirror.cs.wisc.edu/pub/mirrors/ghost/ghostgum/redmon17.zip)
	Extract the archive in Printer\Redmon
	You need only redmon95.dll and redmonnt.dll.

Upx (Version 1.24):
	Download the exe-packer 'upx' from http://upx.sourceforge.net/
	You need only upx.exe.