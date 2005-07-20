' AddWatermarkToPDF script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: July, 18. 2005
' Author: Frank Heindörfer
' Comments: This script needs pdftk. 
'           For more informations about the freeware pdftk use this link:
'           http://www.accesspdf.com

Option Explicit

Const AppTitle = "PDFCreator - AddWatermarkToPDF"
Const PathToPdftk = "c:\pdftk-1.12\pdftk.exe"
Const WatermarkPDF = "watermark.pdf"

Dim objArgs, fname, tfname, fso, WshShell, oExec

Set objArgs = WScript.Arguments

If objArgs.Count = 0 Then
 MsgBox "You can't call the script from commandline!", vbExclamation, AppTitle
 WScript.Quit
End If

fname = objArgs(0)

Set fso = CreateObject("Scripting.FileSystemObject")

If Ucase(fso.GetExtensionName(fname)) <> "PDF" Then
 MsgBox "This script works only with pdf files!", vbExclamation, AppTitle
 WScript.Quit
End If

If Not fso.FileExists(PathToPdftk) Then
 MsgBox "You need pdftk for this script!" & vbcrlf & vbcrlf & _
  "Please go to http://www.accesspdf.com and download it.", vbExclamation, AppTitle
 WScript.Quit
End If

If Not fso.FileExists(WatermarkPDF) Then
 MsgBox "Can't find the watermark pdf file!", vbExclamation, AppTitle
 WScript.Quit
End If

Set WshShell = CreateObject("WScript.Shell")

tfname = fso.GetTempName 
WshShell.Run PathToPdftk & " """ & fname & """ background " & WatermarkPDF & " output """ & tfname & """",0,true

If Not fso.FileExists(tfname) Then
 MsgBox "There was an error using ""pdftk""!", vbCritical, AppTitle
 WScript.Quit
End If

If fso.FileExists(fname) Then
 fso.DeleteFile(fname)
End If

fso.MoveFile tfname, fname
