' SplitPDFFile script
' Part of PDFCreator\pdfforge.dll
' License: FairPlay
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.1
' Date: April, 26. 2010
' Author: Frank Heindörfer
' Comments: Splits a pdf file.

Option Explicit

Dim pdfforge, fso, ScriptBaseName, AppTitle, i, s, p, im, f1
Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "pdfforge.dll - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set pdfforge = Wscript.CreateObject("pdfforge.pdf.pdf")

For i = 1 To 50
 s = s + "0123456789 ÄÖÜäöüß "
Next

p = fso.GetParentFolderName (Wscript.ScriptFullname)
if Right(p,1) <> "\" then p = p & "\"

f1 = p & "TestDocument.pdf"
pdfforge.CreatePDFTestdocument f1, 4, s
pdfforge.SplitPDFFile f1, f1

Set pdfforge = Nothing
Set fso = Nothing
MsgBox "Ready"
