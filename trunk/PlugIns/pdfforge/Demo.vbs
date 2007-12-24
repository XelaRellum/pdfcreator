' Createa test-document script showing some functions of the  pdfforge.dll.
' Part of PDFCreator\pdfforge.dll
' License: FairPlay
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: December, 24. 2007
' Author: Frank Heindörfer
' Comments: Shows some functions of the pdfforge.dll.

Option Explicit

Dim pdfforge, pdftext, tools, fso, WshShell, ScriptBaseName, AppTitle, i, s, p, fname, tfname, imfname

Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
         
ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "pdfforge.dll - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

WshShell.Popup "Please wait some seconds ...", 5, AppTitle, vbInformation


Set pdfText = Wscript.CreateObject("pdfforge.pdf.PDFText")
pdfText.FontColorBlue = 0
pdfText.FontColorRed = 255
pdfText.FontColorGreen = 1
'pdfText.FontName = "BarcodeFont.ttf"
pdfText.FontName = "Arial.ttf"
pdfText.FontPath = WshShell.SpecialFolders("Fonts")
pdfText.FontSize = 36
pdfText.Rotation = 90
pdfText.Text = "www.pdfforge.org"
pdfText.XPosition = 580
pdfText.YPosition = 500

For i = 1 To 50
 s = s + "0123456789 ÄÖÜäöüß "
Next

p = fso.GetParentFolderName (Wscript.ScriptFullname)
if Right(p,1) <> "\" then p = p & "\"

tfname = p + "~tmp.pdf"
fname = p + "TestDocument.pdf"

Set pdfforge = Wscript.CreateObject("pdfforge.pdf.PDF")
pdfforge.CreatePDFTestdocument fname, 20, s

pdfforge.AddTextToPDFFile fname, tfname, 1, 2, (pdfText)
DeleteFile fname

pdfforge.SetBackgroundColor tfname, fname, 1, 2, 0, 255, 255
DeleteFile tfname

imfname = p & "TestImage.png"
Set tools = Wscript.CreateObject("pdfforge.tools")
tools.CreateTestImage imfname, 255, 0, 0

pdfforge.StampPDFFileWithImage fname, tfname, imfname, 5, 6, true, 1, 9
DeleteFile fname
DeleteFile imfname

tools.CreateTestImage imfname, 0, 255, 0
pdfforge.StampPDFFileWithImage tfname, fname, imfname, 7, 8, true, 1, 9
DeleteFile tfname
DeleteFile imfname

tools.CreateTestImage imfname, 0, 0, 255
pdfforge.StampPDFFileWithImage fname, tfname, imfname, 9, 10, true, 1, 9
DeleteFile fname
DeleteFile imfname

pdfforge.NUp tfname, fname, 4
DeleteFile tfname

Set pdfforge = Nothing

MsgBox "Ready"

Private Sub DeleteFile(filename)
 Dim fso
 Set fso = CreateObject("Scripting.FileSystemObject")
 If fso.FileExists(filename) Then
  fso.DeleteFile(filename)
 End If
End Sub