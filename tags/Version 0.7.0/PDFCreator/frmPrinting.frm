VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrinting 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "PDFCreator"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmPrinting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComCtl2.Animation anmProcess 
      Height          =   825
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   1455
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   283
      FullHeight      =   55
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdNow 
      Caption         =   "Now"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtCreateFor 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   5295
   End
   Begin VB.TextBox txtCreationDate 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4335
   End
   Begin VB.CheckBox chkStartStandardProgram 
      Caption         =   "After saving open the document with the standardprogram."
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   5175
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   4440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdWaiting 
      Appearance      =   0  '2D
      BackColor       =   &H00D9E5E9&
      Caption         =   "Waiting"
      Height          =   495
      Left            =   120
      MaskColor       =   &H00D9E5E9&
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "eMail"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5295
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  '2D
      BackColor       =   &H00D9E5E9&
      Caption         =   "Save"
      Height          =   495
      Left            =   4200
      MaskColor       =   &H00D9E5E9&
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblCreateFor 
      Caption         =   "Author:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lblCreationDate 
      Caption         =   "Creation Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lblTitle 
      Caption         =   "Document Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Zentriert
      Caption         =   "Creating file..."
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   11080
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "frmPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SaveFilename As String, SaveFilterIndex As Long, SaveCancel As Boolean

Dim PSHeader As tPSHeader

Private Sub chkStartStandardReader_Click()
 If chkStartStandardProgram.Value = 1 Then
   Options.StartStandardProgram = 1
  Else
   Options.StartStandardProgram = 0
 End If
 SaveOptions Options
End Sub

Private Sub cmdEMail_Click()
 Dim mail As clsPDFCreatorMail, PDFFile As String

 If GsDllLoaded = 0 Then
  MsgBox LanguageStrings.MessagesMsg08
  SetPrinterStop True
  frmMain.Visible = True
  Unload Me
  Exit Sub
 End If

 Set mail = New clsPDFCreatorMail
 ShowAnimation True

 PDFFile = Trim$(Create_eDoc)
 If Dir(PDFFile) <> "" And Len(Trim$(PDFFile)) > 0 Then
  If mail.Send(Trim$(PDFFile)) <> 0 Then
   MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
  End If
 End If

 Me.Visible = False
 ShowAnimation False
 Set mail = Nothing
 Unload Me
End Sub

Private Sub cmdNow_Click()
 txtCreationDate.Text = CStr(Now)
End Sub

Private Sub cmdOptions_Click()
 frmOptions.Show , Me
End Sub

Private Sub cmdSave_Click()
 If GsDllLoaded = 0 Then
  MsgBox LanguageStrings.MessagesMsg08
  SetPrinterStop True
  frmMain.Visible = True
  Unload Me
  Exit Sub
 End If
 Create_eDoc
 Unload Me
End Sub

Private Function GetFilename(FilterIndex As Long) As String
 SaveFilename = ReplaceForbiddenChars(txtTitle.Text) & ".pdf"
 SaveFilterIndex = 0: SaveCancel = False
 frmSave.Show vbModal, Me
 GetFilename = SaveFilename
 FilterIndex = SaveFilterIndex
End Function

Private Function Create_eDoc() As String
 On Local Error GoTo ErrorHandler
 Dim OutputFile As String, Path As String, tStr As String, _
  tErrNumber As Long, FileName As String, FilterIndex As Long

 FileName = GetFilename(FilterIndex)
 If SaveCancel = True Then
  Exit Function
 End If

 Me.MousePointer = vbHourglass

 DoEvents
 Options.StartStandardProgram = chkStartStandardProgram.Value
 OutputFile = Trim$(SaveFilename)
 SplitPath OutputFile, , Path
 Options.LastSaveDirectory = Path
 SaveOptions Options
 ShowAnimation True
 DoEvents
'   SavePDFTitle PDFSpoolfile, txtTitle.Text
 PSHeader.CreateFor.Comment = txtCreateFor.Text
 PSHeader.CreationDate.Comment = txtCreationDate.Text
 PSHeader.Creator.Comment = App.EXEName & _
  " Version " & App.Major & "." & App.Minor & "." & App.Revision
 PSHeader.Title.Comment = txtTitle.Text
 PutPSHeader PDFSpoolfile, PSHeader
 Select Case SaveFilterIndex + 1
  Case 1:
   CallGScript PDFSpoolfile, OutputFile, Options, PDFWriter
  Case 2:
   CallGScript PDFSpoolfile, OutputFile, Options, PNGWriter
  Case 3:
   CallGScript PDFSpoolfile, OutputFile, Options, JPEGWriter
  Case 4:
   CallGScript PDFSpoolfile, OutputFile, Options, BMPWriter
  Case 5:
   CallGScript PDFSpoolfile, OutputFile, Options, PCXWriter
  Case 6:
   CallGScript PDFSpoolfile, OutputFile, Options, TIFFWriter
  Case 7:
   CallGScript PDFSpoolfile, OutputFile, Options, PSWriter
  Case 8:
   CallGScript PDFSpoolfile, OutputFile, Options, EPSWriter
 End Select
 If chkStartStandardProgram.Value = 1 Then
  OpenDocument OutputFile
 End If
 Create_eDoc = OutputFile
 If Dir(PDFSpoolfile) <> "" Then
  Kill PDFSpoolfile
 End If
 Me.MousePointer = vbNormal
 Me.Visible = False
 ShowAnimation False
 Exit Function
ErrorHandler:
 Me.MousePointer = vbNormal
 tErrNumber = Err.number
 tStr = Err.number & ", " & Err.Description
 On Error Resume Next
 Me.Hide
 If tErrNumber <> 32755 Then
  If Dir(PDFSpoolfile) <> "" Then
   Kill PDFSpoolfile
  End If
  IfLoggingWriteLogfile "Error: " & tStr
  IfLoggingShowLogfile frmLog, frmMain
 End If
 Unload Me
End Function

Private Sub ShowAnimation(Show As Boolean)
 DoEvents
 lblTitle.Visible = Not Show
 lblCreationDate.Visible = Not Show
 lblCreateFor.Visible = Not Show
 txtTitle.Visible = Not Show
 txtCreationDate.Visible = Not Show
 txtCreateFor.Visible = Not Show

 cmdWaiting.Visible = Not Show
 cmdEMail.Visible = Not Show
 cmdSave.Visible = Not Show
 chkStartStandardProgram.Visible = Not Show
 cmdNow.Visible = Not Show

 anmProcess.Visible = Show
 lblStatus.Visible = Show

 If Show = True Then
   Me.Height = 1760
   Me.Width = 4530
   ResAnimate anmProcess, ranOpen, 100
   ResAnimate anmProcess, ranPlay
  Else
   ResAnimate anmProcess, ranStop
   ResAnimate anmProcess, ranClose
   Me.Height = 2520
 End If
End Sub

Private Sub cmdWaiting_Click()
 SetPrinterStop True
 frmMain.Visible = True
 Unload Me
End Sub

Private Sub Form_Load()
 Printing = True
 RemoveX Me
 If frmMain.Visible = False Then
  FormInTaskbar Me, True, True
 End If
 With LanguageStrings
  lblTitle.Caption = .PrintingDocumentTitle
  lblStatus.Caption = .PrintingStatus
  lblCreationDate.Caption = .PrintingCreationDate
  lblCreateFor.Caption = .PrintingAuthor
  chkStartStandardProgram.Caption = .PrintingStartStandardProgram
  cmdWaiting.Caption = .PrintingWaiting
  cmdOptions.Caption = .DialogPrinterOptions
  cmdEMail.Caption = .PrintingEMail
  cmdSave.Caption = .PrintingSave
  cmdNow.Caption = .PrintingNow
 End With
 If Options.StartStandardProgram = 1 Then
   chkStartStandardProgram.Value = 1
  Else
   chkStartStandardProgram.Value = 0
 End If
' txtTitle.Text = GetPDFTitle(PDFSpoolfile)
 PSHeader = GetPSHeader(PDFSpoolfile)
 With PSHeader
'  txtTitle.Text = Trim$(.Title.Comment)
  txtTitle.Text = GetSubstFilename(PDFSpoolfile, Options.SaveFilename)
  If Options.UseStandardAuthor = 1 Then
    txtCreateFor.Text = Options.StandardAuthor
   Else
    txtCreateFor.Text = Trim$(.CreateFor.Comment)
  End If
  If Options.UseCreationDateNow = 1 Then
    txtCreationDate.Text = Now
   Else
    txtCreationDate.Text = Trim$(.CreationDate.Comment)
  End If
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
 SetTopMost Me, False, False
 Printing = False
End Sub
