' CombineAndAddBookmarks script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: December, 10. 2005
' Author: Frank Heindörfer
' Comments: This script combines some printjobs in one pdf and add a bookmark for each file.

Option Explicit

Const ForReading = 1, ForAppending = 8
Const maxTime = 30    ' in seconds
Const sleepTime = 250 ' in milliseconds

Dim objArgs, ifname, fso, PDFCreator, DefaultPrinter, ReadyState, _
 i, c, AppTitle, Scriptname, ScriptBasename, FileInfo()

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

Redim FileInfo(1, objArgs.Count - 1)

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")
PDFCreator.cStart "/NoProcessingAtStartup", True
With PDFCreator
 .cPrinterstop =  true
 .cOption("UseAutosave") = 1
 .cOption("UseAutosaveDirectory") = 1
 .cOption("AutosaveFormat") = 0                              ' 0 = PDF
 DefaultPrinter = .cDefaultprinter
 .cDefaultprinter = "PDFCreator"
 .cClearcache
End With

For i = 0 to objArgs.Count - 1
 With PDFCreator
  ifname = objArgs(i)
  If Not fso.FileExists(ifname) Then
   MsgBox "Can't find the file: " & ifname, vbExclamation + vbSystemModal, AppTitle
   Exit For
  End If
  if Not .cIsPrintable(CStr(ifname)) Then
   MsgBox "Converting: " & ifname & vbcrlf & vbcrlf & _
    "An error is occured: File is not printable!", vbExclamation + vbSystemModal, AppTitle
   WScript.Quit
  End if

  .cOption("AutosaveDirectory") = fso.GetParentFolderName(ifname)
  .cOption("AutosaveFilename") = fso.GetBaseName(ifname)
  .cPrintfile cStr(ifname)
  c = 0
  Do While (.cCountOfPrintjobs < i + 1) and (c < (maxTime * 1000 / sleepTime))
   c = c + 1
   Wscript.Sleep sleepTime
  Loop
  FileInfo(0, i) = fso.GetBasename(ifname)
  FileInfo(1, i) = GetCountOfPagesFromPostscriptfile(.cPrintjobFilename(i + 1))
 End With
Next

With PDFCreator
 .cCombineAll
 c = 0
 Do While (.cCountOfPrintjobs <> 1) and (c < (maxTime * 1000 / sleepTime))
  c = c + 1
  Wscript.Sleep sleepTime
 Loop
 ReadyState = 0
 AppendBookmarks .cPrintjobFilename(1)
 .cPrinterStop = false

 c = 0
 Do While (ReadyState = 0) and (c < (maxTime * 1000 / sleepTime))
  c = c + 1
  Wscript.Sleep sleepTime
 Loop
 If ReadyState = 0 then
  MsgBox "Converting: " & ifname & vbcrlf & vbcrlf & _
   "An error is occured: Time is up!", vbExclamation + vbSystemModal, AppTitle
   WScript.Quit
 End If
 .cDefaultprinter = DefaultPrinter
 .cClearcache
 WScript.Sleep 200
 .cClose
End With

Private Sub AppendBookmarks(PostscriptFile)
 Dim fso, f, i, c
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set f = fso.OpenTextFile(PostscriptFile, ForAppending, True)
 f.writeline "[/Page " & 1 & "/View[/Fit]/Title(" & FileInfo(0, 0) & ")/OUT pdfmark"
 For i = 2 to objArgs.Count
  c = c + CLng(FileInfo(1, i - 2))
  f.writeline "[/Page " & c + 1 & "/View[/Fit]/Title(" & FileInfo(0, i - 1) & ")/OUT pdfmark"
 Next
 f.WriteLine "[/PageMode/UseOutlines/Page 1/View[/Fit]/DOCVIEW pdfmark"
 f.Close
End Sub

Private Function GetCountOfPagesFromPostscriptfile(PostscriptFile)
 Dim fso, f, fstr, pp
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set f = fso.OpenTextFile(PostscriptFile, ForReading, True)
 fstr = f.ReadAll
 f.Close
 pp = InstrRev(fstr, "%%Pages:", -1, 1)
 If pp <= 0 Then
  GetCountOfPagesFromPostscriptfile = 1
  Exit Function
 End If
 pp = Instr(pp, fstr," ", 1)
 If pp <= 0 Then
  GetCountOfPagesFromPostscriptfile = 1
  Exit Function
 End If
 fstr = Trim(Mid(fstr,pp))
 fstr = Replace(fstr, chr(10), " ", 1, -1, 1)
 fstr = Replace(fstr, chr(13), " ", 1, -1, 1)
 pp = Instr(1, fstr," ", 1)
 If pp <= 0 Then
  GetCountOfPagesFromPostscriptfile = 1
  Exit Function
 End If
 fstr=mid(fstr,1,pp-1)
 If Not IsNumeric(fstr) Then
  fstr = 1
 End If
 GetCountOfPagesFromPostscriptfile = fstr
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