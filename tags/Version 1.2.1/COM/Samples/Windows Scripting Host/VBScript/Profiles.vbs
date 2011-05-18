' Profiles.vbs script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: March, 11. 2010
' Author: Frank Heindörfer
' Comments: Show handling of the com profiles functions of PDFCreator.

Option Explicit

Dim fso, PDFCreator, _
 AppTitle, Scriptname, Scriptbasename, TestProfile, RenamedProfile, aw
 
Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "PDFCreator - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")

' Show all profiles
 ShowProfiles
 
 TestProfile =  "Test profile 1"
 RenamedProfile =  "Renamed test profile 1"
 aw = MsgBox("Install a profile now.", vbOkCancel + vbInformation, AppTitle)
 If aw = vbCancel Then
  WScript.Quit
 End If

 If Not PDFCreator.cProfileExists(TestProfile) Then
  PDFCreator.cAddProfile (TestProfile)
  ShowProfiles
 End If
 aw = MsgBox("Rename a profile now.", vbOkCancel + vbInformation, AppTitle)
 If aw = vbCancel Then
  WScript.Quit
 End If
 
 PDFCreator.cRenameProfile (TestProfile), (RenamedProfile)
 ShowProfiles
 
 aw = MsgBox("Delete a profile now.", vbOkCancel + vbInformation, AppTitle)
 If aw = vbCancel Then
  WScript.Quit
 End If

 PDFCreator.cDeleteProfile  (RenamedProfile)
 ShowProfiles
 
 Msgbox "Ready", vbOkCancel + vbInformation, AppTitle

Private Sub ShowProfiles
 Dim s, Profiles, i
 Set Profiles = PDFCreator.cGetProfileNames
 If Profiles.Count > 0 Then
   s = Profiles(1)
   For i = 2 To Profiles.Count
    s = s & vbCrLf & Profiles(i)
   Next
   MsgBox "Installed additional profiles:" & vbCrLf & s, vbOkCancel + vbInformation, AppTitle
  Else
   MsgBox "No additional profiles installed." & vbCrLf & s, vbOkCancel + vbInformation, AppTitle
 End If  
End Sub

'--- PDFCreator events ---

Public Sub PDFCreator_eReady()
 ReadyState = 1
End Sub

Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub