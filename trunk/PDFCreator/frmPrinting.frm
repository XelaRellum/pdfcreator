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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkStartStandardProgram.Value = 1 Then
50020    Options.StartStandardProgram = 1
50030   Else
50040    Options.StartStandardProgram = 0
50050  End If
50060  SaveOptions Options
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "chkStartStandardReader_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdEMail_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim mail As clsPDFCreatorMail, PDFFile As String
50020
50030  If GsDllLoaded = 0 Then
50040   MsgBox LanguageStrings.MessagesMsg08
50050   SetPrinterStop True
50060   frmMain.Visible = True
50070   Unload Me
50080   Exit Sub
50090  End If
50100
50110  Set mail = New clsPDFCreatorMail
50120
50130  PDFFile = Trim$(Create_eDoc)
50140  If Dir(PDFFile) <> "" And Len(Trim$(PDFFile)) > 0 Then
50150   If mail.Send(Trim$(PDFFile)) <> 0 Then
50160    MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
50170   End If
50180  End If
50190
50200  Me.Visible = False
50210  Set mail = Nothing
50220  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "cmdEMail_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdNow_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtCreationDate.Text = CStr(Now)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "cmdNow_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdOptions_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  frmOptions.Show , Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "cmdOptions_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdSave_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If GsDllLoaded = 0 Then
50020   MsgBox LanguageStrings.MessagesMsg08
50030   SetPrinterStop True
50040   frmMain.Visible = True
50050   Unload Me
50060   Exit Sub
50070  End If
50080  Create_eDoc
50090  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "cmdSave_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetFilename(FilterIndex As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SaveFilename = ReplaceForbiddenChars(txtTitle.Text) & ".pdf"
50020  SaveFilterIndex = 0: SaveCancel = False
50030  frmSave.Show vbModal, Me
50040  GetFilename = SaveFilename
50050  FilterIndex = SaveFilterIndex
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "GetFilename")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function Create_eDoc() As String
 On Error GoTo ErrorHandler
 Dim OutputFile As String, Path As String, tStr As String, _
  tErrNumber As Long, FileName As String, FilterIndex As Long

 FileName = GetFilename(FilterIndex)
 If SaveCancel = True Then
  Exit Function
 End If

 'Stop Timer: Important! If Animation will be stopped and the timer runs,
 'this form will unload immediately (sendmessage in modResAvi)!
 frmMain.Timer1.Enabled = False
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
 frmMain.Timer1.Enabled = True
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
 frmMain.Timer1.Enabled = True
 Unload Me
End Function

Private Sub ShowAnimation(Show As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DoEvents
50020  lblTitle.Visible = Not Show
50030  lblCreationDate.Visible = Not Show
50040  lblCreateFor.Visible = Not Show
50050  txtTitle.Visible = Not Show
50060  txtCreationDate.Visible = Not Show
50070  txtCreateFor.Visible = Not Show
50080
50090  cmdWaiting.Visible = Not Show
50100  cmdEMail.Visible = Not Show
50110  cmdSave.Visible = Not Show
50120  chkStartStandardProgram.Visible = Not Show
50130  cmdNow.Visible = Not Show
50140
50150  anmProcess.Visible = Show
50160  lblStatus.Visible = Show
50170
50180  If Show = True Then
50190    Me.Height = 1760
50200    Me.Width = 4530
50210    ResAnimate anmProcess, ranOpen, 100
50220    ResAnimate anmProcess, ranPlay
50230   Else
50240    ResAnimate anmProcess, ranStop
50250    ResAnimate anmProcess, ranClose
50260    Me.Height = 2520
50270  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "ShowAnimation")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdWaiting_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetPrinterStop True
50020  frmMain.Visible = True
50030  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "cmdWaiting_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Load()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Printing = True
50020  RemoveX Me
50030  If frmMain.Visible = False Then
50040   FormInTaskbar Me, True, True
50050  End If
50060  With LanguageStrings
50070   lblTitle.Caption = .PrintingDocumentTitle
50080   lblStatus.Caption = .PrintingStatus
50090   lblCreationDate.Caption = .PrintingCreationDate
50100   lblCreateFor.Caption = .PrintingAuthor
50110   chkStartStandardProgram.Caption = .PrintingStartStandardProgram
50120   cmdWaiting.Caption = .PrintingWaiting
50130   cmdOptions.Caption = .DialogPrinterOptions
50140   cmdEMail.Caption = .PrintingEMail
50150   cmdSave.Caption = .PrintingSave
50160   cmdNow.Caption = .PrintingNow
50170  End With
50180  If Options.StartStandardProgram = 1 Then
50190    chkStartStandardProgram.Value = 1
50200   Else
50210    chkStartStandardProgram.Value = 0
50220  End If
50230 ' txtTitle.Text = GetPDFTitle(PDFSpoolfile)
50240  PSHeader = GetPSHeader(PDFSpoolfile)
50250  With PSHeader
50260 '  txtTitle.Text = Trim$(.Title.Comment)
50270   txtTitle.Text = GetSubstFilename(PDFSpoolfile, Options.SaveFilename)
50280   If Options.UseStandardAuthor = 1 Then
50290     txtCreateFor.Text = Options.StandardAuthor
50300    Else
50310     txtCreateFor.Text = Trim$(.CreateFor.Comment)
50320   End If
50330   If Options.UseCreationDateNow = 1 Then
50340     txtCreationDate.Text = Now
50350    Else
50360     txtCreationDate.Text = Trim$(.CreationDate.Comment)
50370   End If
50380  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Unload(Cancel As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetTopMost Me, False, False
50020  Printing = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "Form_Unload")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
