' ShowLogfile script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: September, 1. 2005
' Author: Frank Heindörfer
' Comments: This script shows the logfile of PDFCreator.

Option Explicit

Const HTMLFile = "PDFCreator_logfile.htm"

Dim fso, WshShell, PDFCreator, opt, AppTitle, ScriptBasename

Set fso = CreateObject("Scripting.FileSystemObject")

Scriptbasename = fso.GetFileName(Wscript.ScriptFullname)

AppTitle = "PDFCreator - " & ScriptBasename

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")
PDFCreator.cVisible = False
PDFCreator.cStart "/NoProcessingAtStartup"

CreateHTMLFile HTMLFile, Header & LogFile & Footer

WScript.Sleep 200
PDFCreator.cClose

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run HTMLFile


Private Function Header
 Dim tStr, title
 title = "PDFCreator logfile"
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
 If PDFCreator.cOption("Logging") = 1 then
   tStr = tStr & vbcrlf & "<p>Logging: is activ</p>"
  else
   tStr = tStr & vbcrlf & "<p>Logging: is NOT activ</p>"
 End if
 Header = tStr & vbcrlf & "<p>--------------------------------</p>" & vbcrlf
End Function

Private Function LogFile
 LogFile = Replace(ReplaceForbiddenChars(CStr(PDFCreator.cGetLogfile)),vbcrlf,"<br>") & vbcrlf
End Function

Private Function Footer
 Footer = "</body>" & vbcrlf & "</html>"
End Function

Private Sub CreateHTMLFile(Filename, Content)
 Dim tf
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
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal
 Wscript.Quit
End Sub
