VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrinting 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "PDFCreator"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "frmPrinting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   105
      TabIndex        =   22
      Top             =   5160
      Width           =   1350
   End
   Begin VB.CommandButton cmdWaiting 
      Caption         =   "&Waiting"
      Height          =   495
      Left            =   1680
      TabIndex        =   18
      Top             =   5160
      Width           =   1350
   End
   Begin VB.CommandButton cmdNow 
      Caption         =   "&Now"
      Height          =   300
      Index           =   0
      Left            =   6465
      TabIndex        =   16
      Top             =   1073
      Width           =   1260
   End
   Begin VB.TextBox txtKeywords 
      Height          =   285
      Left            =   105
      TabIndex        =   5
      Top             =   3960
      Width           =   7620
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   105
      TabIndex        =   4
      Top             =   3240
      Width           =   7620
   End
   Begin VB.TextBox txtModifyDate 
      Height          =   285
      Left            =   105
      TabIndex        =   2
      Top             =   1800
      Width           =   6240
   End
   Begin MSComCtl2.Animation anmProcess 
      Height          =   960
      Left            =   1935
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      _Version        =   393216
      AutoPlay        =   -1  'True
      BackStyle       =   1
      FullWidth       =   64
      FullHeight      =   64
   End
   Begin VB.TextBox txtCreateFor 
      Height          =   285
      Left            =   105
      TabIndex        =   3
      Top             =   2520
      Width           =   7620
   End
   Begin VB.TextBox txtCreationDate 
      Height          =   285
      Left            =   105
      TabIndex        =   1
      Top             =   1080
      Width           =   6240
   End
   Begin VB.CheckBox chkStartStandardProgram 
      Caption         =   "After saving open the document with the standardprogram."
      Height          =   615
      Left            =   105
      TabIndex        =   6
      Top             =   4440
      Width           =   7620
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Top             =   360
      Width           =   7620
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   15
      Top             =   0
      Width           =   0
   End
   Begin VB.CommandButton cmdNow 
      Caption         =   "&Now"
      Height          =   300
      Index           =   1
      Left            =   6465
      TabIndex        =   17
      Top             =   1793
      Width           =   1260
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options"
      Height          =   495
      Left            =   3255
      TabIndex        =   19
      Top             =   5160
      Width           =   1350
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "&eMail"
      Height          =   495
      Left            =   4830
      TabIndex        =   20
      Top             =   5160
      Width           =   1350
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   6375
      TabIndex        =   21
      Top             =   5160
      Width           =   1350
   End
   Begin VB.Label lblKeywords 
      AutoSize        =   -1  'True
      Caption         =   "Keywords:"
      Height          =   195
      Left            =   105
      TabIndex        =   12
      Top             =   3720
      Width           =   7620
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   105
      TabIndex        =   11
      Top             =   3000
      Width           =   7620
   End
   Begin VB.Label lblModifyDate 
      AutoSize        =   -1  'True
      Caption         =   "Modify Date:"
      Height          =   195
      Left            =   105
      TabIndex        =   9
      Top             =   1560
      Width           =   6240
   End
   Begin VB.Label lblCreateFor 
      AutoSize        =   -1  'True
      Caption         =   "Author:"
      Height          =   195
      Left            =   105
      TabIndex        =   10
      Top             =   2280
      Width           =   7620
   End
   Begin VB.Label lblCreationDate 
      AutoSize        =   -1  'True
      Caption         =   "Creation Date:"
      Height          =   195
      Left            =   105
      TabIndex        =   8
      Top             =   840
      Width           =   6240
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Document Title:"
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   120
      Width           =   7620
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Zentriert
      Caption         =   "Creating file..."
      Height          =   255
      Left            =   480
      TabIndex        =   14
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

Private PSHeader As tPSHeader, FirstFormActivate As Boolean

Private Sub chkStartStandardProgram_Click()
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
Select Case ErrPtnr.OnError("frmPrinting", "chkStartStandardProgram_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdCancel_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KillFile PDFSpoolfile
50020  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "cmdCancel_Click")
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
50140  PDFFile = Trim$(Create_eDoc)
50150
50160  If Len(PDFFile) > 0 And FileExists(PDFFile) = True Then
50170   If Options.RunProgramAfterSaving = 1 Then
50180    RunProgramAfterSaving Me.hwnd, PDFFile, _
   Options.RunProgramAfterSavingProgramParameters, _
   Options.RunProgramAfterSavingWindowstyle, PDFSpoolfile
50210   End If
50220   Set mail = New clsPDFCreatorMail
50230   If mail.Send(PDFFile, txtSubject.Text, Options.SendMailMethod) <> 0 Then
50240    MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
50250   End If
50260   Set mail = Nothing
50270  End If
50280
50290  KillFile PDFSpoolfile
50300  KillInfoSpoolfile PDFSpoolfile
50310
50320  Me.Visible = False
50330  Unload Me
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
50030    txtCreationDate.Text = Format(CStr(Now), Options.StandardDateformat)
50040   Case 1:
50050    txtModifyDate.Text = Format(CStr(Now), Options.StandardDateformat)
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
50010  If FormISLoaded("frmOptions") = False Then
50020   frmOptions.Show vbModal, Me
50030  End If
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
50010  SaveEDoc
50020  Me.Visible = False
50030  Unload Me
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

Private Sub cmdWaiting_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetPrinterStop True
50020  With frmMain
50030   .WindowState = vbNormal
50040   .Visible = True
50050   SetTopMost frmMain, True, True
50060   SetTopMost frmMain, False, True
50070   SetActiveWindow .hwnd
50080  End With
50090  Unload Me
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

Private Sub Form_Activate()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If FirstFormActivate = False Then
50020   txtTitle.SetFocus
50030   FirstFormActivate = True
50040  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "Form_Activate")
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
50010  Dim tDate As Date, tStr As String
50020  Me.KeyPreview = True
50030  Caption = App.EXEName
50040
50050  Caption = App.Title & " " & GetProgramReleaseStr ' & " " & LanguageStrings.CommonTitle
50060  Printing = True
50070 ' RemoveX Me
50080
50090  With anmProcess
50100   .Top = 0
50110   .Left = 0
50120   .Width = 260 * Screen.TwipsPerPixelX
50130   .Height = 66 * Screen.TwipsPerPixelY
50140  End With
50150
50160  If frmMain.Visible = False Then
50170   FormInTaskbar Me, True, True
50180  End If
50190  With LanguageStrings
50200   lblTitle.Caption = .PrintingDocumentTitle
50210   lblStatus.Caption = .PrintingStatus
50220   lblCreationDate.Caption = .PrintingCreationDate
50230   lblCreateFor.Caption = .PrintingAuthor
50240   lblModifyDate.Caption = .PrintingModifyDate
50250   lblSubject.Caption = .PrintingSubject
50260   lblKeywords.Caption = .PrintingKeywords
50270   chkStartStandardProgram.Caption = .PrintingStartStandardProgram
50280   cmdWaiting.Caption = .PrintingWaiting
50290   cmdOptions.Caption = .DialogPrinterOptions
50300   cmdEMail.Caption = .PrintingEMail
50310   cmdSave.Caption = .PrintingSave
50320   cmdNow(0).Caption = .PrintingNow
50330   cmdNow(1).Caption = .PrintingNow
50340   If LenB(.PrintingCancel) = 0 Then
50350     cmdCancel.Caption = .OptionsCancel
50360    Else
50370     cmdCancel.Caption = .PrintingCancel
50380   End If
50390  End With
50400  If Options.StartStandardProgram = 1 Then
50410    chkStartStandardProgram.Value = 1
50420   Else
50430    chkStartStandardProgram.Value = 0
50440  End If
50450  PSHeader = GetPSHeader(PDFSpoolfile)
50460  With PSHeader
50470   If Len(Trim$(Options.StandardTitle)) > 0 Then
50480     txtTitle.Text = GetSubstFilename(PDFSpoolfile, _
     RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardTitle)))
50500    Else
50510     txtTitle.Text = GetSubstFilename(PDFSpoolfile, Options.SaveFilename)
50520   End If
50530   If Options.UseStandardAuthor = 1 Then
50540     txtCreateFor.Text = GetSubstFilename(PDFSpoolfile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardAuthor)), True)
50550    Else
50560     txtCreateFor.Text = GetDocUsername(PDFSpoolfile, False)
50570   End If
50580   If Len(Trim$(Options.StandardKeywords)) > 0 Then
50590    txtKeywords.Text = GetSubstFilename(PDFSpoolfile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardKeywords)))
50600   End If
50610   If Len(Trim$(Options.StandardSubject)) > 0 Then
50620    txtSubject.Text = GetSubstFilename(PDFSpoolfile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardSubject)))
50630   End If
50640
50650   tDate = Now
50660   If LenB(PSHeader.CreationDate.Comment) > 0 Then
50670     tStr = FormatPrintDocumentDate(PSHeader.CreationDate.Comment)
50680    Else
50690     tStr = CStr(tDate)
50700   End If
50710   txtCreationDate.Text = GetDocDate(Options.StandardCreationdate, Options.StandardDateformat, FormatPrintDocumentDate(tStr))
50720   'tStr = CStr(tDate)
50730   txtModifyDate.Text = GetDocDate(Options.StandardModifydate, Options.StandardDateformat, FormatPrintDocumentDate(tStr))
50740  End With
50750  If Options.OptionsEnabled = 0 Or FormISLoaded("frmOptions") = True Then
50760   cmdOptions.Enabled = False
50770  End If
50780  If Options.OptionsVisible = 0 Then
50790   cmdOptions.Visible = False
50800  End If
50810  If Options.DisableEmail = 1 Then
50820   cmdEMail.Enabled = False
50830  End If
50840  Height = cmdWaiting.Top + cmdWaiting.Height + (Height - ScaleHeight) + 100
50850  With txtTitle
50860   .SelStart = 0
50870   .SelLength = Len(.Text)
50880  End With
50890  SetTopMost Me, True, True
50900  SetTopMost Me, False, True
50910  SetActiveWindow hwnd
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If UnloadMode = vbFormControlMenu Then
50020   KillFile PDFSpoolfile
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "Form_QueryUnload")
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
50010  SetTopMost frmMain, False, False
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

Private Sub txtCreateFor_GotFocus()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With txtCreateFor
50020   If Len(.Text) > 0 Then
50030    .SelStart = 0
50040    .SelLength = Len(.Text)
50050   End If
50060  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "txtCreateFor_GotFocus")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtCreateFor_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyAscii = vbKeyReturn Then
50020   SaveEDoc
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "txtCreateFor_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtCreationDate_GotFocus()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With txtCreationDate
50020   If Len(.Text) > 0 Then
50030    .SelStart = 0
50040    .SelLength = Len(.Text)
50050   End If
50060  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "txtCreationDate_GotFocus")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtCreationDate_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyAscii = vbKeyReturn Then
50020   SaveEDoc
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "txtCreationDate_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtKeywords_GotFocus()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With txtKeywords
50020   If Len(.Text) > 0 Then
50030    .SelStart = 0
50040    .SelLength = Len(.Text)
50050   End If
50060  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "txtKeywords_GotFocus")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtKeywords_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyAscii = vbKeyReturn Then
50020   SaveEDoc
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "txtKeywords_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtModifyDate_GotFocus()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With txtModifyDate
50020   If Len(.Text) > 0 Then
50030    .SelStart = 0
50040    .SelLength = Len(.Text)
50050   End If
50060  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "txtModifyDate_GotFocus")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtModifyDate_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyAscii = vbKeyReturn Then
50020   SaveEDoc
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "txtModifyDate_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtSubject_GotFocus()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With txtSubject
50020   If Len(.Text) > 0 Then
50030    .SelStart = 0
50040    .SelLength = Len(.Text)
50050   End If
50060  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "txtSubject_GotFocus")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtSubject_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyAscii = vbKeyReturn Then
50020   SaveEDoc
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "txtSubject_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtTitle_GotFocus()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With txtTitle
50020   If Len(.Text) > 0 Then
50030    .SelStart = 0
50040    .SelLength = Len(.Text)
50050   End If
50060  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "txtTitle_GotFocus")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyAscii = vbKeyReturn Then
50020   SaveEDoc
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "txtTitle_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function SaveEDoc() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PDFFile As String
50020
50030  SaveEDoc = False
50040
50050  If GsDllLoaded = 0 Then
50060   MsgBox LanguageStrings.MessagesMsg08
50070   SetPrinterStop True
50080   frmMain.Visible = True
50090   Unload Me
50100   Exit Function
50110  End If
50120
50130  PDFFile = Trim$(Create_eDoc)
50140
50150  If PDFFile <> vbNullString Then
50160   SaveEDoc = True
50170   If Options.RunProgramAfterSaving = 1 Then
50180    If Options.OnePagePerFile = 1 Then
50190     PDFFile = Replace$(PDFFile, "%d", "1", , , vbTextCompare)
50200    End If
50210    If Options.RunProgramAfterSaving = 1 Then
50220     RunProgramAfterSaving Me.hwnd, PDFFile, _
     Options.RunProgramAfterSavingProgramParameters, _
     Options.RunProgramAfterSavingWindowstyle, PDFSpoolfile
50250    End If
50260   End If
50270   If chkStartStandardProgram.Value = 1 Then
50280    If Options.OnePagePerFile = 1 Then
50290      OpenDocument Replace$(PDFFile, "%d", "1", , , vbTextCompare)
50300     Else
50310      OpenDocument PDFFile
50320    End If
50330   End If
50340
50350   KillFile PDFSpoolfile
50360   KillInfoSpoolfile PDFSpoolfile
50370  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "SaveEDoc")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function Create_eDoc() As String
 On Error GoTo ErrorHandler
 Dim OutputFile As String, Path As String, tStr As String, Filter As String, _
  tErrNumber As Long, Filename As String, FilterIndex As Long, _
  Cancel As Boolean, PDFDocInfo As tPDFDocInfo, Files As Collection, _
  tStrf() As String, i As Long, Ext As String, Ext2 As String

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
 If IsInIDE Then
  Filter = Filter & "|(*.txt)|*.txt"
 End If
 FilterIndex = 1
 If InStr(1, Filter, "|", vbTextCompare) > 0 Then
  tStrf = Split(Filter, "|")
  If InStr(1, Options.StandardSaveformat, ".") Then
    SplitPath Options.StandardSaveformat, , , , , Ext2
   Else
    Ext2 = Options.StandardSaveformat
  End If
  Ext2 = UCase$(Ext2)
  For i = 1 To UBound(tStrf) Step 2
   SplitPath tStrf(i), , , , , Ext
   If Ext2 = UCase$(Ext) Then
    FilterIndex = (i + 1) \ 2
    Exit For
   End If
  Next i
 End If
 With LanguageStrings
  PSHeader = GetPSHeader(PDFSpoolfile)
  If Len(txtTitle.Text) > 0 And Options.RemoveAllKnownFileExtensions = 1 Then
    SaveFilename = ReplaceForbiddenChars(RemoveAllKnownFileExtensions(txtTitle.Text))
   Else
    SaveFilename = ReplaceForbiddenChars(txtTitle.Text)
  End If
  Set Files = GetFilename(SaveFilename, GetSubstFilename2(Options.LastSaveDirectory), FilterIndex, Filter, SaveFile, Cancel, Me)
  If SaveOpenCancel = True Then
   Exit Function
  End If
  If Files.Count <> 1 Then
   Exit Function
  End If
  SaveFilterIndex = FilterIndex
  SaveFilename = Files.Item(1)
  If FileExists(Files.Item(1)) = True Then
   If FileInUse(Files.Item(1)) = True Then
    MsgBox LanguageStrings.MessagesMsg34
    Exit Function
   End If
  End If
 End With

 'Stop Timer: Important! If Animation will be stopped and the timer runs,
 'this form will unload immediately (sendmessage in modResAvi)!
 frmMain.Timer1.Enabled = False
 Screen.MousePointer = vbHourglass
 DoEvents

 Options.StartStandardProgram = chkStartStandardProgram.Value
 OutputFile = Trim$(SaveFilename)
 SplitPath OutputFile, , Path
 Options.LastSaveDirectory = Path
 SaveOption Options, "LastSaveDirectory"
 frmMain.SetSystrayIcon 3
 If Options.ShowAnimation = 1 Then
   ShowAnimation True
  Else
   Me.Visible = False
 End If
 DoEvents

 PSHeader.CreateFor.Comment = txtCreateFor.Text
 PSHeader.CreationDate.Comment = txtCreationDate.Text
 PSHeader.Creator.Comment = App.EXEName & _
  " Version " & App.Major & "." & App.Minor & "." & App.Revision
 PSHeader.Title.Comment = txtTitle.Text

 With PDFDocInfo
  .Author = txtCreateFor.Text
  .CreationDate = txtCreationDate.Text
  .Creator = App.EXEName & " Version " & App.Major & "." & App.Minor & "." & App.Revision
  .Keywords = GetSubstFilename(PDFSpoolfile, txtKeywords.Text)
  .ModifyDate = txtModifyDate.Text
  .Subject = GetSubstFilename(PDFSpoolfile, txtSubject.Text)
  .Title = GetSubstFilename(PDFSpoolfile, txtTitle.Text)
 End With

 AppendPDFDocInfo PDFSpoolfile, PDFDocInfo
 
 CheckForStamping PDFSpoolfile
 
 If Options.RunProgramBeforeSaving = 1 Then
  RunProgramBeforeSaving Me.hwnd, PDFSpoolfile, _
  Options.RunProgramBeforeSavingProgramParameters, _
  Options.RunProgramBeforeSavingWindowstyle
 End If
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
  Case 9:
   CallGScript PDFSpoolfile, OutputFile, Options, TXTWriter
 End Select
 Create_eDoc = OutputFile
 CheckForPrintingAfterSaving PDFSpoolfile, Options
 Screen.MousePointer = vbNormal
 Me.Visible = False
 If Options.ShowAnimation = 1 Then
  ShowAnimation False
 End If
 ConvertedOutputFilename = OutputFile
 ReadyConverting = True
 frmMain.SetSystrayIcon 2
 frmMain.Timer1.Enabled = True
 Exit Function
ErrorHandler:
 Screen.MousePointer = vbNormal
 tErrNumber = Err.Number
 tStr = Err.Number & ", " & Err.Description
 On Error GoTo 0
 On Error Resume Next
 Me.Hide
 If tErrNumber <> 32755 Then
  MsgBox Err.Description
  KillFile PDFSpoolfile
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
50010  Dim tL As Long, BorderWidth As Long
50020  DoEvents
50030  lblTitle.Visible = Not Show
50040  lblCreationDate.Visible = Not Show
50050  lblCreateFor.Visible = Not Show
50060  txtTitle.Visible = Not Show
50070  txtCreationDate.Visible = Not Show
50080  txtCreateFor.Visible = Not Show
50090
50100  cmdWaiting.Visible = Not Show
50110  cmdEMail.Visible = Not Show
50120  cmdSave.Visible = Not Show
50130  chkStartStandardProgram.Visible = Not Show
50140  cmdNow(0).Visible = Not Show
50150  cmdNow(1).Visible = Not Show
50160  anmProcess.Visible = Show
50170
50180  lblStatus.Visible = Show
50190  'Dim n As FormBorderStyleConstants
50200
50210  If Show = True Then
50220    ResAnimate anmProcess, ranOpen, 100
50230    With Me
50240     BorderWidth = 3
50250     anmProcess.Left = BorderWidth * Screen.TwipsPerPixelX
50260     anmProcess.Top = BorderWidth * Screen.TwipsPerPixelY
50270     .Height = anmProcess.Height + (Height - ScaleHeight) + 2 * BorderWidth * Screen.TwipsPerPixelY
50280     .Width = anmProcess.Width + 4 * BorderWidth * Screen.TwipsPerPixelX
50290     .BorderStyle = vbBSNone
50300     .Caption = .Caption
50310     tL = .Width
50320     .Width = tL - Screen.TwipsPerPixelX
50330     .Width = tL
50340    End With
50350    DrawBorder3D Me, 4, BorderWidth
50360    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
50370    ResAnimate anmProcess, ranPlay
50380   Else
50390    ResAnimate anmProcess, ranStop
50400    ResAnimate anmProcess, ranClose
50410    Me.Height = 2520
50420  End If
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
