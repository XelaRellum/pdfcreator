' URL2PDF script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: September, 21. 2006
' Author: Frank Heindörfer
' Comments: This script saves a web page as pdf file using 
'           the com interface of PDFCreator.

Option Explicit

Const maxTime = 30    ' in seconds
Const sleepTime = 250 ' in milliseconds
Const URLs = "PDFCreator URLs"


Dim fso, PDFCreator, DefaultPrinter, ReadyState, _
 c, AppTitle, Scriptname, ScriptBasename

Set fso = CreateObject("Scripting.FileSystemObject")

Scriptname = fso.GetFileName(Wscript.ScriptFullname)
ScriptBasename = fso.GetFileName(Wscript.ScriptFullname)

AppTitle = "PDFCreator - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 ifnox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")

PDFCreator.cStart "/NoProcessingAtStartup"
With PDFCreator
 .cOption("UseAutosave") = 1
 .cOption("UseAutosaveDirectory") = 1
 .cOption("AutosaveFormat") = 0                              ' 0 = PDF
 DefaultPrinter = .cDefaultprinter
 .cDefaultprinter = "PDFCreator"
 .cClearcache
End With

 With PDFCreator

  ReadyState = 0
  .cOption("AutosaveDirectory") = fso.GetParentFolderName(Wscript.ScriptFullname)
  .cOption("AutosaveFilename") = URLs
  .cPrintURL "http://www.pdfforge.org", 1500
  .cPrintURL "http://www.pdfforge.org/products/pdfcreator", 1500
  .cPrintURL "http://www.pdfforge.org/products/pdfcreator/screenshots", 1500
  WScript.Sleep 5000
  .cCombineAll
  .cPrinterStop = false

  c = 0
  Do While (ReadyState = 0) and (c < (maxTime * 1000 / sleepTime))
   c = c + 1
   Wscript.Sleep sleepTime
  Loop
  If ReadyState = 0 then
   MsgBox "Converting: " & URLs & vbcrlf & vbcrlf & _
   "An error is occured: Time is up!", vbExclamation + vbSystemModal, AppTitle
  End If
 End With

With PDFCreator
 .cDefaultprinter = DefaultPrinter
 .cClearcache
 WScript.Sleep 200
 .cClose
End With

'--- PDFCreator events ---

Public Sub PDFCreator_eReady()
 ReadyState = 1
End Sub

Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub
