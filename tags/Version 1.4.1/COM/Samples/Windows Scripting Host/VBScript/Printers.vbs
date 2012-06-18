' Printers.vbs script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: March, 11. 2010
' Author: Frank Heindörfer
' Comments: Show handling of the com printers functions of PDFCreator.

Option Explicit

Dim fso, PDFCreator, PrinterName, TestProfile, _
 AppTitle, Scriptname, Scriptbasename, aw
 
Set fso = CreateObject("Scripting.FileSystemObject")

ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)

AppTitle = "PDFCreator - " & ScriptBaseName

If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
 MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End if

Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")

If Not PDFCreator.cIsAdministrator Then
 MsgBox "Please run this script as administrator!", vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End If

ShowInstalledPDFCreatorPrinters

PrinterName = "NewPDFCreator"
TestProfile = "NewTestProfile"

aw = MsgBox("Install printer """ & Printername & """ now!", vbOkCancel + vbInformation, AppTitle)
If aw = vbCancel Then
 WScript.Quit
End If
 
If Not PDFCreator.cPrinterIsInstalled(PrinterName) Then
  If Not PDFCreator.cProfileExists(TestProfile) Then
   PDFCreator.cAddProfile (TestProfile)
  End If
  PDFCreator.cAddPDFCreatorPrinter PrinterName, TestProfile
 Else
  MsgBox "Printer """ & Printername & """ is already installed!", vbOk + vbInformation, AppTitle
  WScript.Quit
End If
ShowInstalledPDFCreatorPrinters

aw = MsgBox("Delete printer """ & Printername & """ now!", vbOkCancel + vbInformation, AppTitle)
If aw = vbCancel Then
 WScript.Quit
End If
PDFCreator.cDeletePDFCreatorPrinter(PrinterName)
ShowInstalledPDFCreatorPrinters

Msgbox "Ready"


Public Sub ShowInstalledPDFCreatorPrinters
 Dim s, PDFCreatorPrinters, i
 Set PDFCreatorPrinters = PDFCreator.cGetPDFCreatorPrinters
 If PDFCreatorPrinters.Count > 0 Then
   s = PDFCreatorPrinters(1)
   For i = 2 To PDFCreatorPrinters.Count
    s = s & vbCrLf & PDFCreatorPrinters(i)
   Next
   MsgBox "Installed PDFCreator printers:" & vbCrLf & s, vbOk + vbInformation, AppTitle
  Else
   MsgBox "No PDFCreator printers." & vbCrLf & s, vbOk + vbInformation, AppTitle
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