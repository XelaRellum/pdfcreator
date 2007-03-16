' Testpage2PDF script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: September, 1. 2005
' Author: Frank Heindörfer
' Comments: Print a file and shows the infos about this printjob.

Option Explicit

Dim fso, WshShell, PDFCreator, DefaultPrinter, ReadyState, c, _
 ScriptBaseName, AppTitle

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
 DefaultPrinter = .cDefaultprinter
 .cDefaultprinter = "PDFCreator"
 .cClearcache
 .cPrintPrinterTestpage
 WshShell.Popup "Please wait a moment.", 2, AppTitle, 64
 WScript.Sleep 4000                                         ' Wait until the file is printed
 If .cCountOfPrintJobs>0 Then
   Msgbox GetPrintjobInfo(.cPrintjobFilename(1)), vbOKOnly + vbInformation + vbSystemModal, AppTitle
  Else
   Msgbox "There is no printjob!", vbCritical + vbSystemModal, AppTitle
 End If
End With

With PDFCreator
 .cClearcache
 .cDefaultprinter = DefaultPrinter
 Wscript.Sleep 200
 .cClose
End With

Private Function GetPrintjobInfo(Filename)
 Dim tStr
 With PDFCreator
  tStr = "COMPUTER:" & vbtab & vbtab & .cPrintjobInfo(CStr(Filename), "COMPUTER") & vbcrlf
  tStr = tStr & "CREATED:" & vbtab & vbtab & .cPrintjobInfo(CStr(Filename), "CREATED") & vbcrlf
  tStr = tStr & "SPOOLERACCOUNT:" & vbtab & .cPrintjobInfo(CStr(Filename), "SPOOLERACCOUNT") & vbcrlf
  tStr = tStr & "SPOOLFILENAME:" & vbtab & vbtab & .cPrintjobInfo(CStr(Filename), "SPOOLFILENAME") & vbcrlf & vbcrlf
  tStr = tStr & "REDMON_DOCNAME:" & vbtab & .cPrintjobInfo(CStr(Filename), "REDMON_DOCNAME") & vbcrlf
  tStr = tStr & "REDMON_FILENAME:" & vbtab & .cPrintjobInfo(CStr(Filename), "REDMON_FILENAME") & vbcrlf
  tStr = tStr & "REDMON_JOB:" & vbtab & vbtab & .cPrintjobInfo(CStr(Filename), "REDMON_JOB") & vbcrlf
  tStr = tStr & "REDMON_MACHINE:" & vbtab & .cPrintjobInfo(CStr(Filename), "REDMON_MACHINE") & vbcrlf
  tStr = tStr & "REDMON_PORT:" & vbtab & vbtab & .cPrintjobInfo(CStr(Filename), "REDMON_PORT") & vbcrlf
  tStr = tStr & "REDMON_PRINTER:" & vbtab & .cPrintjobInfo(CStr(Filename), "REDMON_PRINTER") & vbcrlf
  tStr = tStr & "REDMON_SESSIONID:" & vbtab & .cPrintjobInfo(CStr(Filename), "REDMON_SESSIONID") & vbcrlf
  tStr = tStr & "REDMON_USER:" & vbtab & vbtab & .cPrintjobInfo(CStr(Filename), "REDMON_USER")
 End With
 GetPrintjobInfo = tStr
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
