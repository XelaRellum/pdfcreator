' StampPDFFileWithImage script
' Part of PDFCreator\pdfforge.dll
' License: FairPlay
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: December, 24. 2007
' Author: Frank Heindörfer
' Comments: Create a pdf testdocument with 100 pages and copy the pages from 2 until 7 in a new pdf document.

Option Explicit

Dim pdfforge, tools, fso, ScriptBaseName, AppTitle, i, s
Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "pdfforge.dll - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set pdfforge = Wscript.CreateObject("pdfforge.pdf.pdf")
Set tools = Wscript.CreateObject("pdfforge.tools.tools")

For i = 1 To 50
 s = s + "0123456789 ÄÖÜäöüß "
Next

pdfforge.CreatePDFTestdocument "TestDocument.pdf", 4, s

For i = 1 To 50
 s = s + "Stamp "
Next

tools.CreateTestImage "TestImage.png"
pdfforge.StampPDFFileWithImage "TestDocument.pdf", "TestDocumentStamped.pdf", "TestImage.png", true, 0.8, 0, 1, 2
fso.DeleteFile "TestImage.png"
