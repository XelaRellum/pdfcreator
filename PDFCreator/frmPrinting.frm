VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrinting 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "PDFCreator"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmPrinting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtKeywords 
      Height          =   285
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   5295
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   5295
   End
   Begin VB.TextBox txtModifyDate 
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   4335
   End
   Begin VB.CommandButton cmdNow 
      Caption         =   "Now"
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   14
      Top             =   1800
      Width           =   735
   End
   Begin MSComCtl2.Animation anmProcess 
      Height          =   825
      Left            =   360
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
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdNow 
      Caption         =   "Now"
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtCreateFor 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2520
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
      Top             =   4440
      Width           =   5295
   End
   Begin VB.CommandButton cmdWaiting 
      Appearance      =   0  '2D
      BackColor       =   &H00D9E5E9&
      Caption         =   "Waiting"
      Height          =   495
      Left            =   120
      MaskColor       =   &H00D9E5E9&
      TabIndex        =   6
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "eMail"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   5160
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
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lblKeywords 
      Caption         =   "Keywords:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   585
   End
   Begin VB.Label lblModifyDate 
      AutoSize        =   -1  'True
      Caption         =   "Modify Date:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label lblCreateFor 
      AutoSize        =   -1  'True
      Caption         =   "Author:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   510
   End
   Begin VB.Label lblCreationDate 
      AutoSize        =   -1  'True
      Caption         =   "Creation Date:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Document Title:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1125
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

Private SaveFilename As String, SaveFilterIndex As Long

Private PSHeader As tPSHeader

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
50030  ' UnloadDLLComplete GsDllLoaded
50040  ' GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50050
50060  If GsDllLoaded = 0 Then
50070   MsgBox LanguageStrings.MessagesMsg08
50080   SetPrinterStop True
50090   frmMain.Visible = True
50100   Unload Me
50110   Exit Sub
50120  End If
50130
50140  Set mail = New clsPDFCreatorMail
50150
50160  PDFFile = Trim$(Create_eDoc)
50170
50180  ' UnloadDLLComplete GsDllLoaded
50190
50200  If Dir(PDFFile) <> "" And Len(Trim$(PDFFile)) > 0 Then
50210   If mail.Send(PDFFile) <> 0 Then
50220    MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
50230   End If
50240  End If
50250
50260  Me.Visible = False
50270  Set mail = Nothing
50280  Unload Me
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


Private Sub cmdNow_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 0:
50030    txtCreationDate.Text = CStr(Now)
50040   Case 1:
50050    txtModifyDate.Text = CStr(Now)
50060  End Select
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
50010  Dim PDFFile As String
50020
50030  'UnloadDLLComplete GsDllLoaded
50040  'GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50050
50060  If GsDllLoaded = 0 Then
50070   MsgBox LanguageStrings.MessagesMsg08
50080   SetPrinterStop True
50090   frmMain.Visible = True
50100   Unload Me
50110   Exit Sub
50120  End If
50130
50140  PDFFile = Trim$(Create_eDoc)
50150  If PDFFile <> vbNullString Then
50160   Unload Me
50170  End If
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

Private Function Create_eDoc() As String
 On Error GoTo ErrorHandler
 Dim OutputFile As String, Path As String, tStr As String, _
  tErrNumber As Long, Filename As String, Filterindex As Long, _
  PDFDocInfo As tPDFDocInfo, Filter As String, Cancel As Boolean, _
  Files As Collection

 With LanguageStrings
  Filter = .ListPDFFiles & " (*.pdf)|*.pdf|" & _
   .PrintingPNGFiles & " (*.png)|*.png|" & _
   .PrintingJPEGFiles & " (*.jpg)|*.jpg|" & _
   .PrintingBMPFiles & " (*.bmp)|*.bmp|" & _
   .PrintingPCXFiles & " (*.pcx)|*.pcx|" & _
   .PrintingTIFFFiles & " (*.tif)|*.tif|" & _
   .PrintingPSFiles & " (*.ps)|*.ps|" & _
   .PrintingEPSFiles & " (*.eps)|*.eps"
 End With
 With LanguageStrings
  PSHeader = GetPSHeader(PDFSpoolfile)
'  If WinXP = False Then
'    SaveFilterIndex = SaveFileDialog(SaveFilename, ReplaceForbiddenChars(PSHeader.Title.Comment), _
'     Filter, _
'     "*.pdf", _
'     GetMyFiles, _
'     .SaveOpenSaveTitle, _
'      OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_LONGNAMES + OFN_NODEREFERENCELINKS + OFN_HIDEREADONLY, _
'      Me.hwnd)
'    If SaveFilterIndex < 0 Then
'     Exit Function
'    End If
'   Else
    SaveFilename = ReplaceForbiddenChars(txtTitle.Text) & ".pdf"
    Set Files = GetFilename(SaveFilename, Options.LastSaveDirectory, Filterindex, Filter, saveFile, Cancel, Me)
    If SaveOpenCancel = True Then
     Exit Function
    End If
    If Files.Count <> 1 Then
     Exit Function
    End If
    SaveFilterIndex = Filterindex
    SaveFilename = Files.item(1)
'  End If
 End With
 
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
 
 With PDFDocInfo
  .Author = txtCreateFor.Text
  If IsDate(txtCreationDate.Text) = True Then
    .CreationDate = Format(txtCreationDate.Text, "YYYYMMDDHHNNSS")
   Else
    .CreationDate = Format(Now, "YYYYMMDDHHNNSS")
  End If
  .Creator = App.EXEName & " Version " & App.Major & "." & App.Minor & "." & App.Revision
  .Keywords = txtKeywords.Text
  If IsDate(txtModifyDate.Text) = True Then
    .ModifyDate = Format(txtModifyDate.Text, "YYYYMMDDHHNNSS")
   Else
    .ModifyDate = Format(Now, "YYYYMMDDHHNNSS")
  End If
  .Subject = txtSubject.Text
  .Title = txtTitle.Text
 End With
 
 AppendPDFDocInfo PDFSpoolfile, PDFDocInfo
 
' PutPSHeader PDFSpoolfile, PSHeader
 Select Case SaveFilterIndex
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
 On Error GoTo 0
 On Error Resume Next
 Me.Hide
 If tErrNumber <> 32755 Then
  MsgBox Err.Description
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
50130  cmdNow(0).Visible = Not Show: cmdNow(1).Visible = Not Show
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyCode = vbKeyF1 Then
50020   KeyCode = 0
50030   Call HTMLHelp_ShowTopic("html\welcome.htm")
50040  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "Form_KeyDown")
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
50010  Me.KeyPreview = True
50020  Caption = App.Title & " " & GetProgramReleaseStr ' & " " & LanguageStrings.CommonTitle
50030  Printing = True
50040  RemoveX Me
50050  If frmMain.Visible = False Then
50060   FormInTaskbar Me, True, True
50070  End If
50080  With LanguageStrings
50090   lblTitle.Caption = .PrintingDocumentTitle
50100   lblStatus.Caption = .PrintingStatus
50110   lblCreationDate.Caption = .PrintingCreationDate
50120   lblCreateFor.Caption = .PrintingAuthor
50130   lblModifyDate.Caption = .PrintingModifyDate
50140   lblSubject.Caption = .PrintingSubject
50150   lblKeywords.Caption = .PrintingKeywords
50160   chkStartStandardProgram.Caption = .PrintingStartStandardProgram
50170   cmdWaiting.Caption = .PrintingWaiting
50180   cmdOptions.Caption = .DialogPrinterOptions
50190   cmdEMail.Caption = .PrintingEMail
50200   cmdSave.Caption = .PrintingSave
50210   cmdNow(0).Caption = .PrintingNow
50220   cmdNow(1).Caption = .PrintingNow
50230  End With
50240  If Options.StartStandardProgram = 1 Then
50250    chkStartStandardProgram.Value = 1
50260   Else
50270    chkStartStandardProgram.Value = 0
50280  End If
50290 ' txtTitle.Text = GetPDFTitle(PDFSpoolfile)
50300  PSHeader = GetPSHeader(PDFSpoolfile)
50310  With PSHeader
50320 '  txtTitle.Text = Trim$(.Title.Comment)
50330   txtTitle.Text = GetSubstFilename(PDFSpoolfile, Options.SaveFilename)
50340   If Options.UseStandardAuthor = 1 Then
50350 '    txtCreateFor.Text = Options.StandardAuthor
50360     txtCreateFor.Text = GetSubstFilename(PDFSpoolfile, Options.StandardAuthor, True)
50370    Else
50380     txtCreateFor.Text = Trim$(.CreateFor.Comment)
50390   End If
50400   If Options.UseCreationDateNow = 1 Then
50410     txtCreationDate.Text = Now
50420    Else
50430     If IsDate(.CreationDate.Comment) = True Then
50440       txtCreationDate.Text = CStr(CDate(Trim$(.CreationDate.Comment)))
50450      Else
50460       txtCreationDate.Text = Trim$(.CreationDate.Comment)
50470     End If
50480   End If
50490   txtModifyDate.Text = txtCreationDate.Text
50500  End With
50510  If Options.OptionsEnabled = 0 Then
50520   cmdOptions.Enabled = False
50530  End If
50540  If Options.OptionsVisible = 0 Then
50550   cmdOptions.Visible = False
50560  End If
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
