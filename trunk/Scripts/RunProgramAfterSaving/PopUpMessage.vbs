' PopUpMessage script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: July, 18. 2005
' Author: Frank Heindörfer

Option Explicit

Const AppTitle = "PDFCreator - PopUpMessage"
Const SecondsToWait = 5

Dim objArgs, WshShell

Set objArgs = WScript.Arguments

If objArgs.Count = 0 Then
 MsgBox "You can't call the sctipt from commandline!", vbExclamation, AppTitle
 WScript.Quit
End If

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.Popup "PDFCreator: File was created." & vbcrlf & vbcrlf & _
 "Time:" & vbtab & vbtab & Now & vbcrlf & _
 "Filename:" & vbtab & vbtab & objArgs(0) & vbcrlf & _
 "User:" & vbtab & vbtab & objArgs(1) & vbcrlf & _
 "Computer:" & vbtab & Replace(objArgs(2),"\\",""), SecondsToWait, AppTitle, 0
