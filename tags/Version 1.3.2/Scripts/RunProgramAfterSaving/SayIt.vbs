' SayIt script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: July, 18. 2005
' Author: Frank Heindörfer
' Comments: This script needs MS SAPI runtime and a Text-to-speech engine.
'           MS SAPI runtime:       http://www.microsoft.com/MSAGENT/downloads/user.asp#sapi
'           Text-to-speech engine: http://www.microsoft.com/MSAGENT/downloads/user.asp#tts

Option Explicit

Const AppTitle = "PDFCreator - SayIt"
Const TextToSpeech1 = "P D F Creator"
Const TextToSpeech2 = "File was created!"

Dim objArgs, vt

Set objArgs = WScript.Arguments

On Error Resume Next

set vt = WScript.CreateObject("Speech.VoiceText")
If Err.Number <> 0 Then
 MsgBox "This script needs MS SAPI runtime." & vbcrlf & vbcrlf & _
  "For more informations use this link." & vbcrlf & _
  "http://www.microsoft.com/MSAGENT/downloads/user.asp#sapi" & vbcrlf & vbcrlf & _
  Err.Number & " " & Err.Description, vbCritical, AppTitle
 WScript.Quit
End If

vt.Register "", WScript.ScriptName
vt.Speak TextToSpeech1, 1
WScript.Sleep(100)
vt.Speak TextToSpeech2, 1

do while vt.IsSpeaking
 WScript.Sleep(100)
loop

WScript.Quit 