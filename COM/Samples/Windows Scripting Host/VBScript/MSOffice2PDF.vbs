' MSOffice2PDF.vbs script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: February, 16. 2009
' Author: Frank Heindörfer
' Comments: This script convert a MS Office document (Word, Excel, Powerpoint) in a pdf-file using 
'           the com interface of PDFCreator.

Option Explicit

Const maxTime = 30    ' in seconds
Const sleepTime = 250 ' in milliseconds
Const PrinterName = "PDFCreator" ' PDFCreator printer

Dim objArgs, ifname, fso, PDFCreator, i, AppTitle, Scriptname, ScriptBasename, ext
 
Set fso = CreateObject("Scripting.FileSystemObject")

Scriptname = fso.GetFileName(Wscript.ScriptFullname)
ScriptBasename = fso.GetFileName(Wscript.ScriptFullname)

AppTitle = "PDFCreator - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set objArgs = WScript.Arguments

If objArgs.Count = 0 Then
 MsgBox "Syntax: " & vbtab & Scriptname & " <Filename>" & vbcrlf & vbtab & "or use ""Drag and Drop""!", vbExclamation + vbSystemModal, AppTitle
 WScript.Quit
End If

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")
PDFCreator.cStart "/NoProcessingAtStartup"

For i = 0 to objArgs.Count - 1
 ifname = objArgs(i)
 ext = LCase(fso.GetExtensionName(ifname))
 If (ext <> "doc") and (ext <> "docx") and (ext <> "xls") and (ext <> "xlsx")  and (ext <> "ppt") and (ext <> "pptx")then
  MsgBox "File type '*." & ext & "' not supported!", vbExclamation + vbSystemModal + vbOKCancel, AppTitle
  Exit For
 end if

 If Not fso.FileExists(ifname) Then
  MsgBox "Can't find the file: " & ifname, vbExclamation + vbSystemModal, AppTitle
  Exit For
 End If
 
 If (ext = "doc") or (ext = "docx") then
  ConvertWord ifname
 end if
 If (ext = "xls") or (ext = "xlsx") then
  ConvertExcel ifname
 end if
 If (ext = "ppt") or (ext = "pptx") then
  ConvertPowerpoint ifname
 end if
Next

PDFCreator.cClose

MsgBox "Ready", vbInformation + vbSystemModal, AppTitle

'--- Internal sub routines ---
Private Sub ConvertWord(sourceFilename)
 Dim objWord, tempFile, outputFolder, outputFile, DefaultPrinter

 outputFolder = fso.GetParentFolderName(sourceFilename)
 If Mid(outputFolder, Len(outputFolder), 1) <> "\" then outputFolder = outputFolder & "\"
 outputFile = outputFolder + fso.GetBaseName(sourceFilename) & ".pdf"

 Set objWord = CreateObject("Word.Application")
 objWord.Visible = False
 objWord.Documents.Open sourceFilename
 tempFile = GetTempFileName
 DefaultPrinter = objWord.ActivePrinter
 objWord.ActivePrinter = PrinterName
 objWord.PrintOut False, , , tempFile, , , , , , , True

 PDFCreator.cConvertFile tempFile, outputFile

 objWord.ActiveDocument.Close false
 objWord.ActivePrinter = DefaultPrinter
 objWord.Quit
 Set objWord = Nothing
End Sub

Private Sub ConvertExcel(sourceFilename)
 Dim objExcel, tempFile, outputFolder, outputFile

 outputFolder = fso.GetParentFolderName(sourceFilename)
 If Mid(outputFolder, Len(outputFolder), 1) <> "\" then outputFolder = outputFolder & "\"
 outputFile = outputFolder + fso.GetBaseName(sourceFilename) & ".pdf"

 Set objExcel = CreateObject("Excel.Application")
 objExcel.Visible = False
 objExcel.Workbooks.Open sourceFilename
 tempFile = GetTempFileName 
 objExcel.ActiveWorkbook.PrintOut , , , , PrinterName, True, , tempFile

 PDFCreator.cConvertFile tempFile, outputFile

 objExcel.ActiveWorkbook.Close false
 objExcel.Quit
 Set objExcel = Nothing
End Sub

Private Sub ConvertPowerpoint(sourceFilename)
 Dim objPowerpoint, tempFile, outputFolder, outputFile

 outputFolder = fso.GetParentFolderName(sourceFilename)
 If Mid(outputFolder, Len(outputFolder), 1) <> "\" then outputFolder = outputFolder & "\"
 outputFile = outputFolder + fso.GetBaseName(sourceFilename) & ".pdf"

 Set objPowerpoint = CreateObject("Powerpoint.Application")
 objPowerpoint.Visible = True
 objPowerpoint.Presentations.Open sourceFilename
 tempFile = GetTempFileName
 
 objPowerpoint.ActivePresentation.PrintOptions.ActivePrinter  = PrinterName
 objPowerpoint.ActivePresentation.PrintOptions.PrintInBackground = False
 objPowerpoint.ActivePresentation.PrintOut , , tempFile

 PDFCreator.cConvertFile tempFile, outputFile

 objPowerpoint.ActivePresentation.Close
 objPowerpoint.Quit
 Set objPowerpoint = Nothing
End Sub

Private Function GetTempFileName()
 Dim fo
 fo = fso.GetSpecialFolder(2)
 if Len(fo) = 0 Then
   fo = ".\"
  Else
   if Mid(fo, Len(fo), 1) <> "\" Then
    fo = fo & "\"
   End if
 end if
 GetTempFileName = fo + fso.GetTempName
End Function

'--- PDFCreator events ---
Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub
