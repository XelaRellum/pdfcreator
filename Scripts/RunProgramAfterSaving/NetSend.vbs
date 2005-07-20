' NetSend script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: July, 18. 2005
' Author: Frank Heindörfer
' Comments: Please be shure that the messenger service is started!

'Option Explicit

Const AppTitle = "PDFCreator - NetSend"

Dim objArgs, WshShell

Set objArgs = WScript.Arguments

If objArgs.Count = 0 Then
 MsgBox "You can't call the sctipt from commandline!", vbExclamation, AppTitle
 WScript.Quit
End If

Set WshShell = WScript.CreateObject("WScript.Shell")

SendToUser
'SendToComputer


Private Sub SendToUser
 WshShell.Run "net send " & objArgs(1) & " " & _
  "PDFCreator: File was created." & vbcrlf & _
  "Filename: " & objArgs(0) & vbcrlf & _
  "User: " & objArgs(1) & vbcrlf & _
  "Computer:" & Replace(objArgs(2),"\\","")
End Sub 

Private Sub SendToComputer
 WshShell.Run "net send " & Replace(objArgs(2),"\\","")  & " " & _
  "PDFCreator: File was created." & vbcrlf & _
  "Filename: " & objArgs(0) & vbcrlf & _
  "User: " & objArgs(1) & vbcrlf & _
  "Computer:" & Replace(objArgs(2),"\\","")
End Sub 