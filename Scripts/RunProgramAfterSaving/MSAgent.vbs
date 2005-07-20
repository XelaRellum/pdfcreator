' MSAgent script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: July, 18. 2005
' Author: Frank Heindörfer
' Comments: This script needs MS SAPI runtime.
'           MS SAPI runtime: http://www.microsoft.com/MSAGENT/downloads/user.asp#sapi
'           Characters:      http://www.microsoft.com/MSAGENT/downloads/user.asp#character
'           Localizations:   http://www.microsoft.com/MSAGENT/downloads/user.asp#core

Option Explicit

Const AppTitle = "PDFCreator - MSAgent"
Const TextToSpeech = "PDFCreator: File was created!"
Const AgentName = "Merlin"

Dim objAgent, objCharacter, c, HideID, LastID

LastID = 0

On Error Resume Next

Set objAgent = CreateObject("Agent.Control.2")
If Err.Number <> 0 Then
 MsgBox "This script needs MS SAPI runtime." & vbcrlf & vbcrlf & _
  "For more informations use this link." & vbcrlf & _
  "http://www.microsoft.com/MSAGENT/downloads/user.asp#sapi" & vbcrlf & vbcrlf & _
  Err.Number & " " & Err.Description, vbCritical, AppTitle
 WScript.Quit
End If

WScript.ConnectObject objAgent, "Agent_"

objAgent.Connected = TRUE
objAgent.Characters.Load AgentName, AgentName & ".acs"
If Err.Number <> 0 Then
 MsgBox "Try to load the agent character file: " & AgentName & ".acs" & vbcrlf & vbcrlf & _
  "For more informations use this link." & vbcrlf & _
  "http://www.microsoft.com/MSAGENT/downloads/user.asp#character" & vbcrlf & vbcrlf & _
  Err.Number & " " & Err.Description, vbCritical, AppTitle
 WScript.Quit
End If

Set objCharacter = objAgent.Characters.Character(AgentName)

objCharacter.Show
objCharacter.Play "GetAttention"
objCharacter.Speak TextToSpeech
Set HideID = objCharacter.Hide

c = 150 ' Don't wait more than 15 seconds
Do While (c > 0) and (LastID <> HideID)
 c = c -1
 Wscript.Sleep 100
Loop

Public Sub Agent_RequestComplete(ByVal Request)
 LastID = Request
End Sub
