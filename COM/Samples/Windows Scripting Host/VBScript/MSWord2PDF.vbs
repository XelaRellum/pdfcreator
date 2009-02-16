' MSWord2PDF.vbs script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: February, 16. 2009
' Author: Frank Heindörfer
' Comments: This script convert a word document in a pdf-file using 
'           the com interface of PDFCreator.

Option Explicit

Const maxTime = 30    ' in seconds
Const sleepTime = 250 ' in milliseconds

Dim objArgs, ifname, fso, PDFCreator, DefaultPrinter, ReadyState, _
 i, c, AppTitle, Scriptname, ScriptBasename, tempFile, ts1, ts2, objWord
 
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
With PDFCreator
 DefaultPrinter = .cDefaultprinter
 .cDefaultprinter = "PDFCreator"
End With

Set objWord = CreateObject("Word.Application")
For i = 0 to objArgs.Count - 1
 ifname = objArgs(i)
 If Not fso.FileExists(ifname) Then
  MsgBox "Can't find the file: " & ifname, vbExclamation + vbSystemModal, AppTitle
  Exit For
 End If
 
 objWord.Visible = False
 objWord.Documents.Open ifname
 tempFile = GetTempFileName 
 objWord.PrintOut True, , , tempFile, , , , , , , True
 ts1 = fso.GetParentFolderName(ifname)
 If Mid(ts1, Len(ts1), 1) <> "\" then ts1 = ts1 & "\"
 ts2 = ts1 + fso.GetBaseName(ifname) & ".pdf"
 c = 0
 While (fso.FileExists(tempFile) = false) and (c < (maxTime * 1000 / sleepTime))
  c = c + 1
  Wscript.Sleep sleepTime
 Wend

 If Not fso.FileExists(tempFile) Then
  MsgBox "Can't find the print file: " & tempFile, vbExclamation + vbSystemModal, AppTitle
  Exit For
 End If
 c = 0
 While (FileInUse(tempFile) = true) and (c < (maxTime * 1000 / sleepTime))
  c = c + 1
  Wscript.Sleep sleepTime
 Wend
 If FileInUse(tempFile) Then
  MsgBox "Print file to big, increase maxTime!" & tempFile, vbExclamation + vbSystemModal, AppTitle
  Exit For
 End If

 PDFCreator.cConvertFile tempFile, ts2
 objWord.ActiveDocument.Close false
Next
objWord.Quit

With PDFCreator
 .cDefaultprinter = DefaultPrinter
 WScript.Sleep 200
 .cClose
End With

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

Function FileInUse(sFileName)
 On Error Resume Next
 Const ForAppending = 8
 Dim f
 If fso.FileExists(sFileName) Then
  Set f = fso.OpenTextFile(sFileName, ForAppending, False)
  If Err.Number = 70 Then
   FileInUse = True
  Else
   f.Close
   Set f = Nothing
   FileInUse =  False
  End If
 Else
  FileInUse = False
 End If
End Function

'--- PDFCreator events ---
Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub
