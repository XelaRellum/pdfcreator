' TestCompression1 script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: September, 1. 2005
' Author: Frank Heindörfer
' Comments: This script use different compressions options
'           to convert the testpage to a pdf file using 
'           the com interface of PDFCreator.
'           This script use the auto-save mode.

Option Explicit

Const maxTime = 30    ' in seconds
Const sleepTime = 250 ' in milliseconds

Dim WshShell, fso, PDFCreator, DefaultPrinter, ReadyState, opath, _
 ScriptBaseName, AppTitle

Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "PDFCreator - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set WshShell = WScript.CreateObject("WScript.Shell")

opath = fso.GetParentFolderName(Wscript.ScriptFullname)

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")
PDFCreator.cStart "/NoProcessingAtStartup"

With PDFCreator
 .cOption("UseAutosave") = 1
 .cOption("UseAutosaveDirectory") = 1
 .cOption("AutosaveDirectory") = opath
 .cOption("AutosaveFormat") = 0                            ' PDF
 DefaultPrinter = .cDefaultprinter
 .cDefaultprinter = "PDFCreator"
 .cClearcache
 .cPrinterStop = false
End With

WshShell.Popup "Please wait a moment.", 2, AppTitle, 64

CheckCompression "P1-TestPage1", 0, 0, 0, 300, 0, 0, 0, 300, 0, 0, 0, 1200
CheckCompression "P1-TestPage2", 1, 1, 0, 30,  1, 1, 0, 30,  0, 1, 0, 30
CheckCompression "P1-TestPage3", 1, 1, 0, 10,  1, 1, 0, 10,  0, 1, 0, 10

With PDFCreator
 .cDefaultprinter = DefaultPrinter
 Wscript.Sleep 200
 .cClose
End With
 
Private Sub CheckCompression(Filename, PDFCompressionColorCompressionChoice, _
                                       PDFCompressionColorResample, _
                                       PDFCompressionColorResampleChoice, _
                                       PDFCompressionColorResolution, _
                                       PDFCompressionGreyCompressionChoice, _
                                       PDFCompressionGreyResample, _
                                       PDFCompressionGreyResampleChoice, _
                                       PDFCompressionGreyResolution, _
                                       PDFCompressionMonoCompressionChoice, _
                                       PDFCompressionMonoResample, _
                                       PDFCompressionMonoResampleChoice, _
                                       PDFCompressionMonoResolution)
 Dim c
 
 With PDFCreator
  .cOption("AutosaveFilename") = Filename
  .cOption("PDFCompressionColorCompressionChoice") = PDFCompressionColorCompressionChoice
  .cOption("PDFCompressionColorResample") = PDFCompressionColorResample
  .cOption("PDFCompressionColorResampleChoice") = PDFCompressionColorResampleChoice
  .cOption("PDFCompressionColorResolution") = PDFCompressionColorResolution
  .cOption("PDFCompressionGreyCompressionChoice") = PDFCompressionGreyCompressionChoice
  .cOption("PDFCompressionGreyResample") = PDFCompressionGreyResample
  .cOption("PDFCompressionGreyResampleChoice") = PDFCompressionGreyResampleChoice
  .cOption("PDFCompressionGreyResolution") = PDFCompressionGreyResolution
  .cOption("PDFCompressionMonoCompressionChoice") = PDFCompressionMonoCompressionChoice
  .cOption("PDFCompressionMonoResample") = PDFCompressionMonoResample
  .cOption("PDFCompressionMonoResampleChoice") = PDFCompressionMonoResampleChoice
  .cOption("PDFCompressionMonoResolution") = PDFCompressionMonoResolution
  .cPrintPDFCreatorTestpage
 End With
 
 c = 0
 ReadyState = 0
 Do While (ReadyState = 0) and (c < (maxTime * 1000 / sleepTime))
  c = c + 1
  Wscript.Sleep sleepTime
 Loop

 If ReadyState = 0 then
  MsgBox "Max. time is up!", vbExclamation + vbSystemModal, AppTitle
  Wscript.quit
 End If
end Sub

'--- PDFCreator events ---

Public Sub PDFCreator_eReady()
 ReadyState = 1
End Sub

Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub
