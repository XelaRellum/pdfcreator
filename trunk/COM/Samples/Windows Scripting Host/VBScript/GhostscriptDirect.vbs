' GhostscriptDirect script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.pdfforge.org
' Version: 1.0.0.0
' Date: September, 27. 2006
' Author: Frank Heindörfer
' Comments: This script can convert postscript and pdf files in any Ghostscript output format
'           using Ghostscript direct and the com interface of PDFCreator.

Option Explicit

Const OutputFormat = "PNG"
Const GsDevice = "png16m"
Const BitmapResolution = "300"

Dim objArgs, ifname, fso, PDFCreator, WshShell, tStr, _
 i, AppTitle, Scriptname, ScriptBasename, gsArguments(), _
 initArray

initArray = false 
Set fso = CreateObject("Scripting.FileSystemObject")

Scriptname = fso.GetFileName(WScript.ScriptFullname)
ScriptBasename = fso.GetFileName(WScript.ScriptFullname)

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

Set WshShell = WScript.CreateObject("WScript.Shell")

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")
PDFCreator.cStart "/NoProcessingAtStartup", true

For i = 0 to objArgs.Count - 1
 With PDFCreator
 ifname = objArgs(i)
  If Not fso.FileExists(ifname) Then
   MsgBox "Can't find the file: " & ifname, vbExclamation + vbSystemModal, AppTitle
   Exit For
  End If
  tStr = .cOption("DirectoryGhostscriptLibraries") & ";" & .cOption("DirectoryGhostscriptFonts")
  If LenB(LTrim(.cOption("DirectoryGhostscriptResource"))) > 0 Then
   tStr = tStr & ";" & LTrim(.cOption("DirectoryGhostscriptResource"))
  End If
  If LenB(LTrim(.cOption("AdditionalGhostscriptSearchpath"))) > 0 Then
   tStr = tStr & ";" & LTrim(.cOption("AdditionalGhostscriptSearchpath"))
  End If
  AddGsArgument "-I" & tStr
  AddGsArgument "-q"
  AddGsArgument "-dNOPAUSE"
  AddGsArgument "-dSAFER"
  AddGsArgument "-dBATCH"
  If LenB(WshShell.SpecialFolders("Fonts")) > 0 And .cOption("AddWindowsFontpath") = 1 Then
   AddGsArgument "-sFONTPATH=" & WshShell.SpecialFolders("Fonts")
  End If
  AddGsArgument "-sDEVICE=" & GsDevice
  AddGsArgument "-r" & BitmapResolution & "x" & BitmapResolution
  AddGsArgument "-sOutputFile=" & fso.GetParentFolderName(objArgs(i)) & _
   "\" & fso.GetBaseName(objArgs(i)) & "." & LCase(OutputFormat)
  AddGsArgument "-f"
  AddGsArgument objArgs(i)
  .cGhostscriptRun(gsArguments)
 End With
Next

With PDFCreator
 WScript.Sleep 200
 .cClose
End With

Private Sub AddGsArgument(GsArgument)
 If initArray = false Then
   ReDim gsArguments(0)
   initArray = true
  else
   ReDim Preserve gsArguments(Ubound(gsArguments) + 1)
 End If 
 gsArguments(Ubound(gsArguments)) = GsArgument
End Sub

'--- PDFCreator events ---
Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub
