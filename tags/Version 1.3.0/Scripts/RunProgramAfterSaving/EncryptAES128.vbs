' EncryptAES128 script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: September, 23. 2010
' Author: Frank Heindörfer
' Comments: Encrypt a pdf file with the aes methode.

Option Explicit

Const AppTitle = "EncryptAES128"

Dim objArgs, fname, tfname, fso, WshShell, oExec, pdf, enc

Set objArgs = WScript.Arguments

If objArgs.Count = 0 Then
 MsgBox "This script needs a parameter!", vbExclamation, AppTitle
 WScript.Quit
End If

fname = objArgs(0)

Set fso = CreateObject("Scripting.FileSystemObject")

If Ucase(fso.GetExtensionName(fname)) <> "PDF" Then
 MsgBox "This script works only with pdf files!", vbExclamation, AppTitle
 WScript.Quit
End If

tfname = fso.GetTempName

Set WshShell = CreateObject("WScript.Shell")

Set enc = WScript.CreateObject("pdfforge.PDF.PDFEncryptor")
enc.AllowAssembly = false
enc.AllowCopy = false
enc.AllowFillIn = true
enc.AllowModifyAnnotations = false
enc.AllowModifyContents = false
enc.AllowPrinting = true
enc.AllowPrintingHighResolution = false
enc.AllowScreenreaders = false
enc.EncryptionMethode = 2 ' AES 128 bit encryption
enc.OwnerPassword = "pdfforge"
Set pdf = WScript.CreateObject("pdfforge.pdf.pdf")

pdf.EncryptPDFFile fname, tfname, (enc)

If Not fso.FileExists(tfname) Then
 MsgBox "There was an error during enrypting!", vbCritical, AppTitle
 WScript.Quit
End If

If fso.FileExists(fname) Then
 fso.DeleteFile(fname)
End If

fso.MoveFile tfname, fname

Set enc = Nothing
Set pdf = Nothing
Set fso = Nothing
Set WshShell = Nothing
Set objArgs = Nothing
