' CombineJobs script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: September, 1. 2005
' Author: Frank Heindörfer
' Comments: This script combines some printjobs in one pdf.

Option Explicit

Const maxTime = 30    ' in seconds
Const sleepTime = 250 ' in milliseconds

Dim PDFCreator, DefaultPrinter, ReadyState, fso, c, opath, _
 AppTitle, ScriptBasename, WshShell

Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "PDFCreator - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

opath = CompletePath(fso.GetParentFolderName(Wscript.ScriptFullname))

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Popup "Please wait a moment.", 2, AppTitle, 64

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")
PDFCreator.cStart "/NoProcessingAtStartup"

With PDFCreator
 .cOption("UseAutosave") = 1
 .cOption("UseAutosaveDirectory") = 1
 .cOption("AutosaveDirectory") = opath
 .cOption("AutosaveFilename") = Scriptbasename
 DefaultPrinter = .cDefaultprinter
 .cDefaultprinter = "PDFCreator"
 .cClearcache

 ' 1. page
 CreateTextfileAndPrint opath & "1.txt", "1"

 ' 2. page
 CreateTextfileAndPrint opath & "2.txt", "2"

 ' 3. page
 CreateTextfileAndPrint opath & "3.txt", "3"

 ' 4. page
 CreateTextfileAndPrint opath & "4.txt", "4"

 Wscript.Sleep 2000                                        ' Wait until all files are printed

 ' Page order: 1 2 3 4

 .cMovePrintjobTop 3 
 ' Page order: 3 1 2 4

  
 .cMovePrintjobBottom 2
 ' Page order: 3 2 4 1

 .cMovePrintjobDown 2
 ' Page order: 3 4 2 1

 .cMovePrintjobUp 2
 ' Page order: 4 3 2 1

 .cDeletePrintjob 1
 ' Page order: 3 2 1

 .cCombineAll

 .cPrinterStop = False
End With

c = 0
ReadyState = 0
Do While (ReadyState = 0) and (c < (maxTime * 1000 / sleepTime))
 c = c + 1
 Wscript.Sleep sleepTime
Loop

If ReadyState = 0 then
 MsgBox "An error is occured: Time is up!", vbExclamation + vbSystemModal, AppTitle
 Wscript.quit
End If

With PDFCreator
 .cDefaultprinter = DefaultPrinter
 Wscript.Sleep 200
 .cClose
End With

Public Sub CreateTextfileAndPrint(Filename, Content)
 Dim fso, f
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set f = fso.CreateTextFile(Filename, True)
 f.WriteLine(Content)
 f.Close
 PDFCreator.cPrintfile cStr(Filename)
 WScript.Sleep 2000                                        ' Wait until the file is printed
 fso.DeleteFile(Filename)
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
