' Testpage2PDFSendEmail script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.1.0.0
' Date: October, 9. 2008
' Author: Frank Heindörfer
' Comments: Save the test page as pdf-file and send as email using
'           the com interface of PDFCreator.

Option Explicit

Const maxTime = 10    ' in seconds
Const sleepTime = 250 ' in milliseconds

Dim fso, WshShell, PDFCreator, PDFCreatorOptions, DefaultPrinter, ReadyState, c, _
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
Set PDFCreatorOptions = Wscript.CreateObject("PDFCreator.clsPDFCreatorOptions")

PDFCreator.cStart "/NoProcessingAtStartup"

Set PDFCreatorOptions = PDFCreator.cOptions
With PDFCreatorOptions
 .UseAutosave = 1
 .UseAutosaveDirectory = 1
 .AutosaveDirectory = fso.GetParentFolderName(Wscript.ScriptFullname)
 .AutosaveFilename = "Testpage - PDFCreator"
 .AutosaveFormat = 0                                       ' 0 = PDF
 .StandardSubject = "Here is the pdf file"
End With
Set PDFCreator.cOptions = PDFCreatorOptions

With PDFCreator
 DefaultPrinter = .cDefaultprinter
 .cDefaultprinter = "PDFCreator"
 .cClearcache
 ReadyState = 0
 .cPrintPDFCreatorTestpage
 .cPrinterstop = false
End With

c = 0

Do While (ReadyState = 0) and (c < (maxTime * 1000 / sleepTime))
 c = c + 1
 Wscript.Sleep sleepTime
Loop

If ReadyState = 0 then
 MsgBox "Creating test page as pdf." & vbcrlf & vbcrlf & _
  "An error is occured: Time is up!", vbExclamation + vbSystemModal, AppTitle
 With PDFCreator
  .cDefaultprinter = DefaultPrinter
  WScript.Sleep 200
  .cClose
 End With
 Wscript.Quit
End If

With PDFCreator
 .cDefaultprinter = DefaultPrinter
 .cSendMail .cOutputfilename, "email@address.com"
 Wscript.Sleep 2000
 fso.Deletefile .cOutputfilename
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