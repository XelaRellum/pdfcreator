' AddLineText script
' Part of PDFCreator\pdfforge.dll
' License: FairPlay
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.1
' Date: April, 26. 2010
' Author: Frank Heindörfer
' Comments: Create a pdf testdocument with 10 pages and add a text and add a line.

Option Explicit

Dim pdfforge, fso, ScriptBaseName, AppTitle, i, s, pdfLine, pdfText

Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "pdfforge.dll - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

For i = 1 To 50
 s = s + "0123456789 ÄÖÜäöüß "
Next

Set pdfforge = Wscript.CreateObject("pdfforge.pdf.pdf")
pdfforge.CreatePDFTestdocument "TestDocument.pdf", 10, s

Set pdfLine  = Wscript.CreateObject("pdfforge.pdf.pdfline")
' Add crop marks to a pdf using the standard line object.
pdfforge.AddCropMarksToPDFFile "TestDocument.pdf", "Result.pdf", 1, 3, 4, 4, 4, 4, (pdfline)

' Change some line parameters and add a line to a pdf.
pdfLine.FromX = 15
pdfLine.FromY = 20
pdfLine.ToX = pdfLine.FromX + 180
pdfLine.ToY = pdfLine.FromY
pdfLine.LineColorBlue = 255
pdfLine.LineThickness = 4
pdfLine.UnitsOn = 10
pdfLine.UnitsOff = 10
pdfLine.Phase = 5
pdfforge.AddLineToPDFFile "Result.pdf", "Result2.pdf", 1, 1, (pdfLine)

' Add a text to a pdf.
Set pdfText  = Wscript.CreateObject("pdfforge.pdf.pdfText")
pdfText.Text = "For eyes only"
pdfText.FontColorRed = 255
pdfText.FontName = "Arial.ttf"
pdfText.FontPath = "C:\Windows\Fonts\"
pdfText.FontSize = 72
pdfText.Rotation = 45
pdfText.XPosition = 50
pdfText.YPosition = 100
pdfText.FillOpacity = 0.5
pdfforge.AddTextToPDFFile "Result2.pdf", "Result3.pdf", 1, 1, (pdfText)

Set pdfforge = Nothing
Set fso = Nothing
MsgBox "Ready"