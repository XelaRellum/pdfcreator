' Testpage2PDF script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: September, 1. 2005
' Author: Frank Heindörfer
' Comments: This script shows the gui functions 
'           of the com interface of PDFCreator.

Option Explicit

Dim fso, PDFCreator, aw, ScriptBaseName, AppTitle

Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "PDFCreator - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")

MsgBox "This script shows the gui functions of the com interface of PDFCreator.", vbInformation + vbSystemModal, AppTitle

With PDFCreator
 MsgBox "Set Visible = False", vbInformation + vbSystemModal
 .cVisible = False
 MsgBox "Start PDFCreator. PDFCreator is not on the desktop or in the systray.", vbInformation + vbSystemModal, AppTitle
 .cStart "/NoProcessingAtStartup"
 MsgBox "Set Visible = True", vbInformation + vbSystemModal, AppTitle
 .cVisible = true
 MsgBox "PDFCreator is on the desktop or in the systray now.", vbInformation + vbSystemModal, AppTitle
 MsgBox "Show the options dialog.", vbInformation + vbSystemModal, AppTitle
 .cShowOptionsDialog True
 Wscript.Sleep 250
 Do
  aw = Msgbox("Please close the options dialog.", vbOkCancel + vbInformation + vbSystemModal, AppTitle)
 Loop Until (aw=vbOk and .cIsOptionsDialogDisplayed=False) or aw=vbCancel
 If aw=vbCancel Then
  MsgBox "User cancel the script!", vbExclamation + vbSystemModal, AppTitle
  Wscript.Quit
 End If
 MsgBox "Show the options dialog.", vbInformation + vbSystemModal, AppTitle
 .cShowOptionsDialog True
 Wscript.Sleep 250
 MsgBox "PDFCreator closes the options dialog." & vbcrlf & _
  "This is the same like the user press cancel.", vbInformation + vbSystemModal, AppTitle
 .cShowOptionsDialog False
 MsgBox "Show the logfile dialog.", vbInformation + vbSystemModal, AppTitle
 .cShowLogfileDialog True
 Wscript.Sleep 250
 Do
  aw = Msgbox("Please close the logfile dialog.", vbOkCancel + vbInformation + vbSystemModal, AppTitle)
 Loop Until (aw=vbOk and .cIsLogfileDialogDisplayed=False) or aw=vbCancel
 If aw=vbCancel Then
  MsgBox "User cancel the script!", vbExclamation + vbSystemModal, AppTitle
  Wscript.Quit
 End If
 MsgBox "Show the logfile dialog.", vbInformation + vbSystemModal, AppTitle
 .cShowLogfileDialog True
 Wscript.Sleep 250
 MsgBox "PDFCreator closes the logfile dialog." & vbcrlf & _
  "This is the same like the user press cancel.", vbInformation + vbSystemModal, AppTitle
 .cShowlogfileDialog False
 WScript.Sleep 200
 .cClose
End With
 
MsgBox "Ready", vbInformation + vbSystemModal, AppTitle

'--- PDFCreator events ---

Public Sub PDFCreator_eReady()
 ReadyState = 1
End Sub

Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub
