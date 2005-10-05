VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPDFCreator
   Caption         =   "UserForm1"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7485
   OleObjectBlob   =   "frmPDFCreatorWord.frx":0000
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
 Dim outName As String
 AddStatus "Start ..."
 If InStr(1, ActiveDocument.Name, ".", vbTextCompare) > 1 Then
   outName = Mid(ActiveDocument.Name, 1, InStr(1, ActiveDocument.Name, ".", vbTextCompare) - 1)
  Else
   outName = ActiveDocument.Name
 End If
 With PDFCreator1
  .cOption("UseAutosave") = 1
  .cOption("UseAutosaveDirectory") = 1
  .cOption("AutosaveDirectory") = ActiveDocument.Path
  .cOption("AutosaveFilename") = outName
  .cOption("AutosaveFormat") = 0                            ' 0 = PDF
  .cClearCache
  DoEvents
  ActiveDocument.PrintOut False
  DoEvents
  .cPrinterStop = False
 End With
End Sub

Private Sub PDFCreator1_eError()
 AddStatus "ERROR [" & PDFCreator1.cErrorDetail("Number") & "]: " & PDFCreator1.cErrorDetail("Description")
End Sub

Private Sub PDFCreator1_eReady()
 AddStatus "File'" & PDFCreator1.cOutputFilename & "' was saved."
End Sub

Private Sub UserForm_Initialize()
 If Len(ActiveDocument.Path) = 0 Then
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
  DefaultPrinter = ActivePrinter
  SetPrinter "PDFCreator"
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

Private Sub SetPrinter(Printername As String)
 With Dialogs(wdDialogFilePrintSetup)
  .Printer = Printername
  .DoNotSetAsSysDefault = True
  .Execute
 End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 SetPrinter DefaultPrinter
 PDFCreator1.cClose
 Set PDFCreator1 = Nothing
 Sleep 250
 DoEvents
End Sub
