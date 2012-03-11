' ConvertJPEG2PDF.vbs script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: August, 10. 2008
' Author: Frank Heindörfer
' Comments: This script convert JPEG-filee in a pdf-file using 
'           the com interface of PDFCreator.

Option Explicit

Dim objArgs, ifname, fso, PDFCreator, i, AppTitle, Scriptname, ScriptBasename

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
 With PDFCreator
  ifname = objArgs(i)
  If Not fso.FileExists(ifname) Then
   MsgBox "Can't find the file: " & ifname, vbExclamation + vbSystemModal, AppTitle
   Exit For
  End If
  .cConvertFile ifname, CompletePath(fso.GetParentFolderName(ifname)) & fso.GetBaseName(ifname) & ".pdf"
 End With
Next

With PDFCreator
 WScript.Sleep 200
 .cClose
End With

Private Function CompletePath(Path)
 If Right(Path, 1) <> "\" Then
   CompletePath = Path & "\"
  Else
   CompletePath = Path
 End If
End Function

'--- PDFCreator events ---

Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub
