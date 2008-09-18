' StampPDFFileWithImage script
' Part of PDFCreator\pdfforge.dll
' License: FairPlay
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: December, 24. 2007
' Author: Frank Heindörfer
' Comments: Create a pdf testdocument with 100 pages and stamp the pages 1 until 4 with an external image.

Option Explicit

Dim pdfforge, tools, fso, ScriptBaseName, AppTitle, i, s, p, im, f1, f2, f3
Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "pdfforge.dll - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set pdfforge = Wscript.CreateObject("pdfforge.pdf.pdf")
Set tools = Wscript.CreateObject("pdfforge.tools")

For i = 1 To 50
 s = s + "0123456789 ÄÖÜäöüß "
Next

p = fso.GetParentFolderName (Wscript.ScriptFullname)
if Right(p,1) <> "\" then p = p & "\"

f1 = p & "TestDocument.pdf"
f2 = p & "~Stamped.pdf"
f3 = p & "TestDocumentStamped.pdf"
im = p & "TestImage.png"
pdfforge.CreatePDFTestdocument f1, 4, s
tools.CreateTestImage im, 255, 0, 0
pdfforge.StampPDFFileWithImage f1, f2, im, 1, 2, true, 0.8, 0
tools.CreateTestImage im, 0, 255, 0
pdfforge.StampPDFFileWithImage f2, f3, im, 3, 4, true, 1, 9
fso.DeleteFile f1
fso.DeleteFile f2
fso.DeleteFile im
