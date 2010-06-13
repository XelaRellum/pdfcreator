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

Dim objArgs, fso, pdfforge, ScriptBaseName, AppTitle, i
Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "pdfforge.dll - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set objArgs = WScript.Arguments

If objArgs.Count = 0 Then
 MsgBox "Syntax: " & vbtab & Scriptname & " <Filename>" & vbcrlf & vbtab & "or use ""Drag and Drop""!", vbExclamation + vbSystemModal, AppTitle
 WScript.Quit
End If

Set pdfforge = Wscript.CreateObject("pdfforge.pdf.pdf")

For i = 0 to objArgs.Count - 1
 pdfforge.SplitPDFFile objArgs(i), objArgs(i)
Next

Set objArgs = Nothing
Set pdfforge = Nothing
Set fso = Nothing
MsgBox "Ready"
