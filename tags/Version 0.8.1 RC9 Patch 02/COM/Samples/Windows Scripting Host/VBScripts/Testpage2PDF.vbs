' Testpage2PDF script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: September, 1. 2005
' Author: Frank Heindörfer
' Comments: Save the test page as pdf-file using
'           the com interface of PDFCreator.

Option Explicit

Const maxTime = 10    ' in seconds
Const sleepTime = 250 ' in milliseconds

Dim fso, WshShell, PDFCreator, DefaultPrinter, ReadyState, c, _
 AppTitle, Scriptname, Scriptbasename

Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "PDFCreator - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set WshShell = WScript.CreateObject("WScript.Shell")

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")
PDFCreator.cStart "/NoProcessingAtStartup"

ReadyState = 0
With PDFCreator
 .cOption("UseAutosave") = 1
 .cOption("UseAutosaveDirectory") = 1
 .cOption("AutosaveDirectory") = fso.GetParentFolderName(Wscript.ScriptFullname)
 .cOption("AutosaveFilename") = "Testpage - PDFCreator"
 .cOption("AutosaveFormat") = 0                            ' 0 = PDF
 DefaultPrinter = .cDefaultprinter
 .cDefaultprinter = "PDFCreator"
 .cClearcache
 .cPrintPDFCreatorTestpage
 .cPrinterStop = false
End With

c = 0

Do While (ReadyState = 0) and (c < (maxTime * 1000 / sleepTime))
 c = c + 1
 Wscript.Sleep sleepTime
Loop

With PDFCreator
 .cDefaultprinter = DefaultPrinter
 WScript.Sleep 200
 .cClose
End With

If ReadyState = 0 then
 MsgBox "Creating test page as pdf." & vbcrlf & vbcrlf & _
  "An error is occured: Time is up!", vbExclamation + vbSystemModal, AppTitle
End If

'--- PDFCreator events ---

Public Sub PDFCreator_eReady()
 ReadyState = 1
End Sub

Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub