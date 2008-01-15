' SaveOptionsToFile script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 3.2.0.0
' Date: December, 24. 2007
' Author: Frank Heindörfer
' Comments: Save the pdfcreator options as ini-file.

Option Explicit

Const sleepTime = 1000

Dim fso, PDFCreator, AppTitle, Scriptbasename, PDFCreatorOptions, ProgramIsRunning

Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "PDFCreator - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")

With PDFCreator
 ProgramIsRunning = .cProgramIsRunning
 .cStart "/NoProcessingAtStartup", true
 .cSaveOptionsToFile fso.GetParentFolderName(Wscript.ScriptFullname) & "\PDFCreator.ini"
 WScript.Sleep sleepTime
 If ProgramIsRunning = false then
  .cClose
 End If 
End With

'--- PDFCreator events ---

Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub