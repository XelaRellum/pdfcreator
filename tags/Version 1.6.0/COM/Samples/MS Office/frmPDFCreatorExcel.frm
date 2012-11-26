VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPDFCreator 
   Caption         =   "UserForm1"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7530
   OleObjectBlob   =   "frmPDFCreatorExcel.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmPDFCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

' Add a reference to PDFCreator
Private WithEvents PDFCreator1 As PDFCreator.clsPDFCreator
Attribute PDFCreator1.VB_VarHelpID = -1

Private ReadyState As Boolean, DefaultPrinter As String

Private Sub CommandButton1_Click()
 Dim outName As String, i As Long
 If InStr(1, ActiveWorkbook.Name, ".", vbTextCompare) > 1 Then
   outName = Mid(ActiveWorkbook.Name, 1, InStr(1, ActiveWorkbook.Name, ".", vbTextCompare) - 1)
  Else
   outName = ActiveWorkbook.Name
 End If
 CommandButton1.Enabled = False
 If OptionButton1.Value = True Then
  With PDFCreator1
   .cOption("UseAutosave") = 1
   .cOption("UseAutosaveDirectory") = 1
   .cOption("AutosaveDirectory") = ActiveWorkbook.Path
   .cOption("AutosaveFilename") = outName
   .cOption("AutosaveFormat") = 0                            ' 0 = PDF
   .cClearCache
  End With
  For i = 1 To Application.Sheets.Count
   Application.Sheets(i).PrintOut Copies:=1, ActivePrinter:="PDFCreator"
  Next i
  Do Until PDFCreator1.cCountOfPrintjobs = Application.Sheets.Count
   DoEvents
   Sleep 1000
  Loop
  Sleep 1000
  PDFCreator1.cCombineAll
  Sleep 1000
  PDFCreator1.cPrinterStop = False
 End If
 If OptionButton2.Value = True Then
  With PDFCreator1
   .cOption("UseAutosave") = 1
   .cOption("UseAutosaveDirectory") = 1
   .cOption("AutosaveDirectory") = ActiveWorkbook.Path
   Debug.Print outName & "-" & ActiveSheet.Name
   .cOption("AutosaveFilename") = outName & "-" & ActiveSheet.Name
   .cOption("AutosaveFormat") = 0                            ' 0 = PDF
   .cClearCache
  End With
  ActiveSheet.PrintOut Copies:=1, ActivePrinter:="PDFCreator"
  Do Until PDFCreator1.cCountOfPrintjobs = 1
   DoEvents
   Sleep 1000
  Loop
  Sleep 1000
  PDFCreator1.cPrinterStop = False
 End If
End Sub

Private Sub PrintPage(PageNumber As Integer)
 Dim cPages As Long
 cPages = Selection.Information(wdNumberOfPagesInDocument)
 If PageNumber > cPages Then
  MsgBox "This document has only " & cPages & " pages!", vbExclamation
 End If
 DoEvents
 ActiveDocument.PrintOut Background:=False, Range:=wdPrintFromTo, From:=CStr(PageNumber), To:=CStr(PageNumber)
 DoEvents
End Sub

Private Sub PDFCreator1_eError()
 AddStatus "ERROR [" & PDFCreator1.cErrorDetail("Number") & "]: " & PDFCreator1.cErrorDetail("Description")
End Sub

Private Sub PDFCreator1_eReady()
 AddStatus "File'" & PDFCreator1.cOutputFilename & "' was saved."
 PDFCreator1.cPrinterStop = True
 CommandButton1.Enabled = True
End Sub

Private Sub UserForm_Initialize()
 If Len(ActiveWorkbook.Path) = 0 Then
  MsgBox "Please save the document first!", vbExclamation
  End
 End If
 Set PDFCreator1 = New clsPDFCreator
 With PDFCreator1
  If .cStart("/NoProcessingAtStartup") = False Then
   CommandButton1.Enabled = False
   AddStatus "Can't initialize PDFCreator."
   Exit Sub
  End If
 End With
 AddStatus "PDFCreator initialized."
End Sub

Private Sub AddStatus(Str1 As String)
 With TextBox1
  If Len(.Text) = 0 Then
    .Text = Now & ": " & Str1
   Else
    .Text = .Text & vbCrLf & Now & ": " & Str1
  End If
  .SelStart = Len(.Text)
  .SetFocus
 End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 PDFCreator1.cClose
 Set PDFCreator1 = Nothing
 Sleep 250
 DoEvents
End Sub
