' TestCompression3 script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: September, 1. 2005
' Author: Frank Heindörfer
' Comments: This script use different values for the color and gray jpeg compressions factor
'           to convert the testpage to a pdf file using the com interface of PDFCreator.

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
 .cClearcache
 .cPrintPDFcreatorTestpage
 .cOption("PDFCompressionColorCompressionChoice") = 1      ' JPEG Maximum
 .cOption("PDFCompressionGreyCompressionChoice") = 1       ' JPEG Maximum
End With

WshShell.Popup "Please wait a moment.", 2, AppTitle, 64

CheckCompression "P3-TestPage1",  2,  2
CheckCompression "P3-TestPage2", 10, 10
CheckCompression "P3-TestPage3", 20, 20

With PDFCreator
 .cClearcache
 Wscript.Sleep 200
 .cClose
End With
 
Private Sub CheckCompression(Filename, PDFCompressionColorCompressionJPEGMaximumFactor, _
                                       PDFCompressionGreyCompressionJPEGMaximumFactor)
 Dim c
 
 With PDFCreator
  .cOption("PDFCompressionColorCompressionJPEGMaximumFactor") = PDFCompressionColorCompressionJPEGMaximumFactor
  .cOption("PDFCompressionGreyCompressionJPEGMaximumFactor") = PDFCompressionGreyCompressionJPEGmaximumFactor
  .cConvertPostscriptfile .cPrintjobFilename(1), CompletePath(opath) & Filename & ".pdf"
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
