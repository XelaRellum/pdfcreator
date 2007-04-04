' ShowOptions script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: September, 1. 2005
' Author: Frank Heindörfer
' Comments: This script shows all options of PDFCreator.

Option Explicit

Const HTMLFile = "PDFCreator_options.htm"

Dim fso, WshShell, oExec, PDFCreator, opt, optnames, _
 AppTitle, ScriptBasename

Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBasename = fso.GetFileName(Wscript.ScriptFullname)

AppTitle = "PDFCreator - " & Scriptbasename

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")
PDFCreator.cVisible = False
PDFCreator.cStart "/NoProcessingAtStartup"

Set optNames = PDFCreator.cOptionsNames

CreateHTMLFile HTMLFile, Header & Table & Footer

WScript.Sleep 200
PDFCreator.cClose

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run HTMLFile


Private Function Header
 Dim tStr, title
 title = "PDFCreator options"
 tStr = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01//EN"""
 tStr = tStr & vbcrlf & "<html>"
 tStr = tStr & vbcrlf & "<head>"
 tStr = tStr & vbcrlf & "<title>" & title & "</title>"
 tStr = tStr & vbcrlf & "</head>"
 tStr = tStr & vbcrlf & "<body>"
 tStr = tStr & vbcrlf & "<h1>" & title & "</h1>"
 tStr = tStr & vbcrlf & "<p>Windows version: " & PDFCreator.cWindowsversion & "</p>"
 tStr = tStr & vbcrlf & "<p>Program release: " & PDFCreator.cProgramRelease & "</p>"
 If PDFCreator.cInstalledAsServer then
   tStr = tStr & vbcrlf & "<p>Installation mode: Server</p>"
  else
   tStr = tStr & vbcrlf & "<p>Installation mode: Standard</p>"
 End if
 tStr = tStr & vbcrlf & "<p>Count of options: " & optnames.Count & "</p>"
 Header = tStr & vbcrlf
End Function

Private Function Table
 Dim tStr, item, c
 tStr = "<table border=""1"" style=""border-collapse:collapse;empty-cells:show;"">"
 tStr = tStr & vbcrlf & "<tr>"
 tStr = tStr & vbcrlf & "<th></th>"
 tStr = tStr & vbcrlf & "<th>Option</th>"
 tStr = tStr & vbcrlf & "<th>Value</th>"
 tStr = tStr & vbcrlf & "<th>Standard value</th>"
 tStr = tStr & vbcrlf & "</tr>"

 c = 0
 For Each item in optnames
  c = c + 1
  tStr = tStr & vbcrlf & "<tr>"
  tStr = tStr & vbcrlf & "<td>" & c & "</td>"
  tStr = tStr & vbcrlf & "<td>" & ReplaceForbiddenChars(item) & "</td>"
  If ReplaceForbiddenChars(PDFCreator.cOption(CStr(item))) <> ReplaceForbiddenChars(PDFCreator.cStandardOption(CStr(item))) Then
    tStr = tStr & vbcrlf & "<td bgcolor=""#00C0C0"">" & ReplaceForbiddenChars(PDFCreator.cOption(CStr(item))) & "</td>"
   Else
    tStr = tStr & vbcrlf & "<td>" & ReplaceForbiddenChars(PDFCreator.cOption(CStr(item))) & "</td>"
  End IF 
  tStr = tStr & vbcrlf & "<td>" & ReplaceForbiddenChars(PDFCreator.cStandardOption(CStr(item))) & "</td>"
  tStr = tStr & vbcrlf & "</tr>"
 Next

 tStr = tStr & vbcrlf & "</table>"

 Table = tStr & vbcrlf
End Function

Private Function Footer
 Dim tStr
 tStr = "</body>"
 tStr = tStr & vbcrlf & "</html>"
 Footer = tStr
End Function

Private Sub CreateHTMLFile(Filename, Content)
 Dim fso, tf
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set tf = fso.CreateTextFile(Filename, True)
 tf.Write Content
 tf.Close
End Sub

Private Function ReplaceForbiddenChars(value)
 Dim tStr
 tStr=Replace(value, "&", "&amp;")
 tStr=Replace(tStr, "<", "&lt;")
 tStr=Replace(tStr, ">", "&gt;")
 tStr=Replace(tStr, """", "&quot;")
 ReplaceForbiddenChars = tStr
End Function

'--- PDFCreator events ---

Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub