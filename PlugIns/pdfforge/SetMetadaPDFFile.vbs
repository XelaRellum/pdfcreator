' SetMetadaFromPDFFile script
' Part of PDFCreator\pdfforge.dll
' License: FairPlay
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: December, 24. 2007
' Author: Frank Heindörfer
' Comments: Set a new metadata value of a pdf file.

Option Explicit

Dim pdfforge, fso, ScriptBaseName, AppTitle, i, s, p, f1, f2

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

f1 = p + "TestDocument.pdf"
f2 = p + "~temp.pdf"

pdfforge.CreatePDFTestdocument f1, 100, s
MsgBox "Old value for author: " & pdfforge.GetMetadata(f1, "Author")
pdfforge.SetMetadataKey f1, f2, "Author", "Frank & Philip"
fso.DeleteFile f1
fso.MoveFile f2, f1
MsgBox "New value for author: " & pdfforge.GetMetadata(f1, "Author")
