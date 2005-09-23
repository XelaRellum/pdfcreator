' FTP upload script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.1.0.0
' Date: September, 1. 2005
' Author: Frank Heindörfer

Option Explicit

Const AppTitle = "PDFCreator - FTPUpload"

Dim objArgs, fname, domain, user, pass, rdir

domain="127.0.0.1"
rdir=""
user="anonymous"
pass="anonymous@"

Set objArgs = WScript.Arguments

If objArgs.Count = 0 Then
 MsgBox "This script needs a parameter!", vbExclamation, AppTitle
 WScript.Quit
End If

fname = objArgs(0)


Call FTPUpload(domain, rdir, user, pass, fname)

Private Sub FTPUpload(domain, rdir, user, pass, fname)
 Dim fso, ftpo
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set ftpo = CreateObject("InetCtls.Inet.1")
 ftpo.URL = "ftp://" & domain
 ftpo.UserName = user
 ftpo.Password = pass
 ftpo.Execute , "CD " & rdir

 Do
  WScript.Sleep 100
 Loop while ftpo.StillExecuting

 ftpo.Execute , "Put """ & fname & """ """ & fso.GetFilename(fname) & """"
 
 Do
  WScript.Sleep 100
 Loop while ftpo.StillExecuting

 ftpo.Execute , "Close"
End Sub 