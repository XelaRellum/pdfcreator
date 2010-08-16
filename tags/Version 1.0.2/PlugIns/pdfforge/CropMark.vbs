' AddLineText script
' Part of PDFCreator\pdfforge.dll
' License: FairPlay
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: September, 18. 2009
' Author: Frank Heindörfer
' Comments: Create a pdf testdocument with 10 pages and add a text and a line.

Option Explicit

Dim pdfforge, fso, ScriptBaseName, AppTitle, i, s, pdfLine, pdfText

Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "pdfforge.dll - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set pdfforge = Wscript.CreateObject("pdfforge.pdf.pdf")

Set pdfLine  = Wscript.CreateObject("pdfforge.pdf.pdfline")
' Add crop marks to a pdf using the standard line object.
pdfforge.AddCropMarksToPDFFile "TestDocument.pdf", "TestDocument_1.pdf", 1, 2, 3, 3, 3, 3, (pdfline)

Set pdfLine = Nothing
Set pdfforge = Nothing
Set fso = Nothing
MsgBox "Ready"
