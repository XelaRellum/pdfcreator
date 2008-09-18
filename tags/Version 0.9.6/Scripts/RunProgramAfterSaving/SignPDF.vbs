' SignPDF script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: December, 11. 2007
' Author: Frank Heindörfer
' Comments: This script signs a PDF using a exported P12/PFX Zertifikat.
Option Explicit

Const AppTitle = "PDFCreator - SignPDF"
Const Certificate = "C:\MyCertificate.p12" '  "C:\MyCertificate.pfx"
Const signatureReason = ""
Const signatureContact = ""
Const signatureLocation = ""
Const signatureVisible = false
Const signaturePositionLowerLeftX = 100
Const signaturePositionLowerLeftY = 100
Const signaturePositionUpperRightX = 200
Const signaturePositionUpperRightY = 200
Const multiSignatures = true

				
Dim objArgs, fname, tfname, fso, WshShell, oExec, pdfforge, certificatePassword

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

On Error Resume Next
Set pdfforge = CreateObject("pdfforge.pdf.PDFTools")
If err.number = 429 Then
 MsgBox "The pdfforge.dll coming with PDFCreator is not installed! A possible reason can be a missing Microsoft .Net 1.1!", vbExclamation, AppTitle
 WScript.Quit
End If
On Error Goto 0

If Not fso.FileExists(Certificate) Then
 MsgBox "Can't find the certficate file '" &  Certificate & "'!", vbExclamation, AppTitle
 WScript.Quit
End If

tfname = fso.GetTempName

certificatePassword = InputBox("Enter the certificate password", AppTitle) 

If IsEmpty(certificatePassword) Then
 MsgBox "Script has been canceled by user!", vbExclamation, AppTitle
 WScript.Quit
End If
				
pdfforge.signPDF fname, tfname, Certificate, certificatePassword, signatureReason, signatureContact, signatureLocation, signatureVisible, signaturePositionLowerLeftX, signaturePositionLowerLeftY, signaturePositionUpperRightX, signaturePositionUpperRightY, multiSignatures

If fso.FileExists(fname) Then
 fso.DeleteFile(fname)
End If

fso.MoveFile tfname, fname

Set objArgs = Nothing
Set pdfforge = Nothing
Set fso = Nothing
