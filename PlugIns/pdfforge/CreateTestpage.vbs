' CreateTestpage script
' Part of PDFCreator\pdfforge.dll
' License: FairPlay
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: December, 24. 2007
' Author: Frank Heind�rfer
' Comments: Create a pdf testdocument with 100 pages using pdfforge.dll.

Option Explicit

Dim pdfforge, fso, ScriptBaseName, AppTitle, i, s

Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "pdfforge.dll - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

For i = 1 To 50
 s = s + "0123456789 ������� "
Next

Set pdfforge = Wscript.CreateObject("pdfforge.pdf.pdf")
pdfforge.CreatePDFTestdocument "TestDocument.pdf", 100, s
pdfforge = Nothing
