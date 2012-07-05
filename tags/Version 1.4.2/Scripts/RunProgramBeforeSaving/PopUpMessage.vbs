' PopUpMessage script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.1.0.0
' Date: September, 1. 2005
' Author: Frank Heindörfer

Option Explicit

Const AppTitle = "PDFCreator - PopUpMessage"
Const SecondsToWait = 5

Dim objArgs, WshShell

Set objArgs = WScript.Arguments

If objArgs.Count = 0 Then
 MsgBox "This script needs a parameter!", vbExclamation, AppTitle
 WScript.Quit
End If

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.Popup "PDFCreator: File was created." & vbcrlf & vbcrlf & _
 "Filename:" & vbtab & vbtab & objArgs(0) & vbcrlf, SecondsToWait, AppTitle, 0
