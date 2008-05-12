' nUp script
' Part of PDFCreator\pdfforge.dll
' License: FairPlay
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: May, 12. 2008
' Author: Frank Heindörfer
' Comments: Create a pdf testdocument with 4 pages per sheet.

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
f2 = p & "nUp.pdf"
pdfforge.CreatePDFTestdocument f1, 16, s
pdfforge.NUp f1, f2, 4
