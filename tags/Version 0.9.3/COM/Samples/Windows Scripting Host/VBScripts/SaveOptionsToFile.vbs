' SaveOptionsToFile script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 2.0.0.0
' Date: Aprl, 14. 2005
' Author: Frank Heindörfer
' Comments: Save the pdfcreator as ini-file.

Option Explicit

Const sleepTime = 250 ' in milliseconds

Dim fso, PDFCreator, AppTitle, Scriptbasename, PDFCreatorOptions

Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "PDFCreator - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")

With PDFCreator
 .cStart "/NoProcessingAtStartup"
 .cSaveOptionsToFile fso.GetParentFolderName(Wscript.ScriptFullname) & "\PDFCreator.ini"
 WScript.Sleep sleepTime
 .cClose
End With

'--- PDFCreator events ---

Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub