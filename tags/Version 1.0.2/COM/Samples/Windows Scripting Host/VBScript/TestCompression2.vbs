' TestCompression2 script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: September, 1. 2005
' Author: Frank Heindörfer
' Comments: This script use different compressions options
'           to convert the testpage to a pdf file using 
'           the com interface of PDFCreator.
'           This script use the function "cConvertPostscriptfile".

Option Explicit

Const maxTime = 10    ' in seconds
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
 .cClearcache
 .cPrintPDFcreatorTestpage
End With

WshShell.Popup "Please wait a moment.", 2, AppTitle, 64

CheckCompression "P2-TestPage1", 0, 0, 0, 300, 0, 0, 0, 300, 0, 0, 0, 1200
CheckCompression "P2-TestPage2", 1, 1, 0, 30,  1, 1, 0, 30,  0, 1, 0, 30
CheckCompression "P2-TestPage3", 1, 1, 0, 10,  1, 1, 0, 10,  0, 1, 0, 10

With PDFCreator
 .cClearcache
 WScript.Sleep 200
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
  ReadyState = 0
  .cConvertPostscriptfile .cPrintjobFilename(1), CompletePath(opath) & Filename & ".pdf"
 End With
 
 c = 0
 Do While (ReadyState = 0) and (c < (maxTime * 1000 / sleepTime))
  c = c + 1
  Wscript.Sleep sleepTime
 Loop

 If ReadyState = 0 then
  MsgBox "Max. time is up!", vbExclamation + vbSystemModal, AppTitle
  Wscript.quit
 End If
End Sub

Private Function CompletePath(Path)
 If Right(Path, 1) <> "\" Then
   CompletePath = Path & "\"
  Else
   CompletePath = Path
 End If
End Function

'--- PDFCreator events ---

Public Sub PDFCreator_eReady()
 ReadyState = 1
End Sub

Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub
