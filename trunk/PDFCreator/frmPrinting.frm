VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrinting 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "PDFCreator"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmPrinting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtKeywords 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   6030
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   6030
   End
   Begin VB.TextBox txtModifyDate 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4680
   End
   Begin VB.CommandButton cmdNow 
      Caption         =   "Now"
      Height          =   285
      Index           =   1
      Left            =   4890
      TabIndex        =   4
      Top             =   1800
      Width           =   1260
   End
   Begin MSComCtl2.Animation anmProcess 
      Height          =   960
      Left            =   360
      TabIndex        =   19
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
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options"
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   5160
      Width           =   1350
   End
   Begin VB.CommandButton cmdNow 
      Caption         =   "Now"
      Height          =   285
      Index           =   0
      Left            =   4890
      TabIndex        =   2
      Top             =   1080
      Width           =   1260
   End
   Begin VB.TextBox txtCreateFor 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   6030
   End
   Begin VB.TextBox txtCreationDate 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4650
   End
   Begin VB.CheckBox chkStartStandardProgram 
      Caption         =   "After saving open the document with the standardprogram."
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   6030
   End
   Begin VB.CommandButton cmdWaiting 
      Appearance      =   0  '2D
      BackColor       =   &H00D9E5E9&
      Caption         =   "Waiting"
      Height          =   495
      Left            =   105
      MaskColor       =   &H00D9E5E9&
      TabIndex        =   12
      Top             =   5160
      Width           =   1350
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "eMail"
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   5160
      Width           =   1350
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6030
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  '2D
      BackColor       =   &H00D9E5E9&
      Caption         =   "Save"
      Height          =   495
      Left            =   4800
      MaskColor       =   &H00D9E5E9&
      TabIndex        =   9
      Top             =   5160
      Width           =   1350
   End
   Begin VB.Label lblKeywords 
      AutoSize        =   -1  'True
      Caption         =   "Keywords:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   585
   End
   Begin VB.Label lblModifyDate 
      AutoSize        =   -1  'True
      Caption         =   "Modify Date:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label lblCreateFor 
      AutoSize        =   -1  'True
      Caption         =   "Author:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   510
   End
   Begin VB.Label lblCreationDate 
      AutoSize        =   -1  'True
      Caption         =   "Creation Date:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Document Title:"
      Height          =   195
      Left            =   90
      TabIndex        =   13
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Zentriert
      Caption         =   "Creating file..."
      Height          =   255
      Left            =   480
      TabIndex        =   20
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

Private Sub cmdEMail_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim mail As clsPDFCreatorMail, PDFFile As String
50050
50060  ' UnloadDLLComplete GsDllLoaded
50070  ' GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50080
50090  If GsDllLoaded = 0 Then
50100   MsgBox LanguageStrings.MessagesMsg08
50110   SetPrinterStop True
50120   frmMain.Visible = True
50130   Unload Me
50140   Exit Sub
50150  End If
50160
50170  PDFFile = Trim$(Create_eDoc)
50180
50190  If Len(PDFFile) > 0 And FileExists(PDFFile) = True Then
50200   If Options.RunProgramAfterSaving = 1 Then
50210    RunProgramAfterSaving Me.hwnd, GetShortName(PDFFile), GetRunProgramAfterSavingProgramParameters, Options.RunProgramAfterSavingWindowstyle
50220   End If
50230   Set mail = New clsPDFCreatorMail
50240   If mail.Send(PDFFile, txtSubject.Text, Options.SendMailMethod) <> 0 Then
50250    MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
50260   End If
50270   Set mail = Nothing
50280  End If
50290
50300  Me.Visible = False
50310  Unload Me
50320 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50330 Exit Sub
ErrPtnr_OnError:
50351 Select Case ErrPtnr.OnError("frmPrinting", "cmdEMail_Click")
      Case 0: Resume
50370 Case 1: Resume Next
50380 Case 2: Exit Sub
50390 Case 3: End
50400 End Select
50410 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdNow_Click(Index As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50041  Select Case Index
        Case 0:
50060    txtCreationDate.Text = Format(CStr(Now), Options.StandardDateformat)
50070   Case 1:
50080    txtModifyDate.Text = Format(CStr(Now), Options.StandardDateformat)
50090  End Select
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmPrinting", "cmdNow_Click")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SaveEDoc
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmPrinting", "cmdSave_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdWaiting_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SetPrinterStop True
50050  With frmMain
50060   .Visible = True
50070   .WindowState = 0
50080   SetTopMost frmMain, True, True
50090   SetTopMost frmMain, False, True
50100   SetActiveWindow .hwnd
50110  End With
50120  Unload Me
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Sub
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("frmPrinting", "cmdWaiting_Click")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Sub
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If KeyCode = vbKeyF1 Then
50050   KeyCode = 0
50060   Call HTMLHelp_ShowTopic("html\welcome.htm")
50070  End If
50080 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50090 Exit Sub
ErrPtnr_OnError:
50111 Select Case ErrPtnr.OnError("frmPrinting", "Form_KeyDown")
      Case 0: Resume
50130 Case 1: Resume Next
50140 Case 2: Exit Sub
50150 Case 3: End
50160 End Select
50170 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Load()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tDate As Date, tStr As String
50050  Me.KeyPreview = True
50060  Caption = App.EXEName
50070
50080  Caption = App.Title & " " & GetProgramReleaseStr ' & " " & LanguageStrings.CommonTitle
50090  Printing = True
50100  RemoveX Me
50110
50120  With anmProcess
50130   .Top = 0
50140   .Left = 0
50150   .Width = 260 * Screen.TwipsPerPixelX
50160   .Height = 66 * Screen.TwipsPerPixelY
50170  End With
50180
50190  If frmMain.Visible = False Then
50200   FormInTaskbar Me, True, True
50210  End If
50220  With LanguageStrings
50230   lblTitle.Caption = .PrintingDocumentTitle
50240   lblStatus.Caption = .PrintingStatus
50250   lblCreationDate.Caption = .PrintingCreationDate
50260   lblCreateFor.Caption = .PrintingAuthor
50270   lblModifyDate.Caption = .PrintingModifyDate
50280   lblSubject.Caption = .PrintingSubject
50290   lblKeywords.Caption = .PrintingKeywords
50300   chkStartStandardProgram.Caption = .PrintingStartStandardProgram
50310   cmdWaiting.Caption = .PrintingWaiting
50320   cmdOptions.Caption = .DialogPrinterOptions
50330   cmdEMail.Caption = .PrintingEMail
50340   cmdSave.Caption = .PrintingSave
50350   cmdNow(0).Caption = .PrintingNow
50360   cmdNow(1).Caption = .PrintingNow
50370  End With
50380  If Options.StartStandardProgram = 1 Then
50390    chkStartStandardProgram.Value = 1
50400   Else
50410    chkStartStandardProgram.Value = 0
50420  End If
50430  PSHeader = GetPSHeader(PDFSpoolfile)
50440  With PSHeader
50450   If Len(Trim$(Options.StandardTitle)) > 0 Then
50460     txtTitle.Text = GetSubstFilename(PDFSpoolfile, _
     RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardTitle)))
50480    Else
50490     txtTitle.Text = GetSubstFilename(PDFSpoolfile, Options.SaveFilename)
50500   End If
50510   If Len(txtTitle.Text) > 0 And Options.RemoveAllKnownFileExtensions = 1 Then
50520    txtTitle.Text = RemoveAllKnownFileExtensions(txtTitle.Text)
50530   End If
50540   If Options.UseStandardAuthor = 1 Then
50550     txtCreateFor.Text = GetSubstFilename(PDFSpoolfile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardAuthor)), True)
50560    Else
50570     txtCreateFor.Text = GetDocUsername(PDFSpoolfile, False)
50580   End If
50590   If Len(Trim$(Options.StandardKeywords)) > 0 Then
50600    txtKeywords.Text = GetSubstFilename(PDFSpoolfile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardKeywords)))
50610   End If
50620   If Len(Trim$(Options.StandardSubject)) > 0 Then
50630    txtSubject.Text = GetSubstFilename(PDFSpoolfile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardSubject)))
50640   End If
50650
50660   tDate = Now
50670   If LenB(PSHeader.CreationDate.Comment) > 0 Then
50680     tStr = FormatPrintDocumentDate(PSHeader.CreationDate.Comment)
50690    Else
50700     tStr = CStr(tDate)
50710   End If
50720   txtCreationDate.Text = GetDocDate(Options.StandardCreationdate, Options.StandardDateformat, FormatPrintDocumentDate(tStr))
50730   'tStr = CStr(tDate)
50740   txtModifyDate.Text = GetDocDate(Options.StandardModifydate, Options.StandardDateformat, FormatPrintDocumentDate(tStr))
50750  End With
50760  If Options.OptionsEnabled = 0 Or FormISLoaded("frmOptions") = True Then
50770   cmdOptions.Enabled = False
50780  End If
50790  If Options.OptionsVisible = 0 Then
50800   cmdOptions.Visible = False
50810  End If
50820  SetTopMost Me, True, True
50830  SetTopMost Me, False, True
50840  SetActiveWindow hwnd
50850 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50860 Exit Sub
ErrPtnr_OnError:
50881 Select Case ErrPtnr.OnError("frmPrinting", "Form_Load")
      Case 0: Resume
50900 Case 1: Resume Next
50910 Case 2: Exit Sub
50920 Case 3: End
50930 End Select
50940 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Unload(Cancel As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SetTopMost frmMain, False, False
50050  Printing = False
50060 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50070 Exit Sub
ErrPtnr_OnError:
50091 Select Case ErrPtnr.OnError("frmPrinting", "Form_Unload")
      Case 0: Resume
50110 Case 1: Resume Next
50120 Case 2: Exit Sub
50130 Case 3: End
50140 End Select
50150 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtCreateFor_GotFocus()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  With txtCreateFor
50050   If Len(.Text) > 0 Then
50060    .SelStart = 0
50070    .SelLength = Len(.Text)
50080   End If
50090  End With
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmPrinting", "txtCreateFor_GotFocus")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtCreateFor_KeyPress(KeyAscii As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If KeyAscii = vbKeyReturn Then
50050   SaveEDoc
50060  End If
50070 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50080 Exit Sub
ErrPtnr_OnError:
50101 Select Case ErrPtnr.OnError("frmPrinting", "txtCreateFor_KeyPress")
      Case 0: Resume
50120 Case 1: Resume Next
50130 Case 2: Exit Sub
50140 Case 3: End
50150 End Select
50160 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtCreationDate_GotFocus()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  With txtCreationDate
50050   If Len(.Text) > 0 Then
50060    .SelStart = 0
50070    .SelLength = Len(.Text)
50080   End If
50090  End With
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmPrinting", "txtCreationDate_GotFocus")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtCreationDate_KeyPress(KeyAscii As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If KeyAscii = vbKeyReturn Then
50050   SaveEDoc
50060  End If
50070 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50080 Exit Sub
ErrPtnr_OnError:
50101 Select Case ErrPtnr.OnError("frmPrinting", "txtCreationDate_KeyPress")
      Case 0: Resume
50120 Case 1: Resume Next
50130 Case 2: Exit Sub
50140 Case 3: End
50150 End Select
50160 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtKeywords_GotFocus()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  With txtKeywords
50050   If Len(.Text) > 0 Then
50060    .SelStart = 0
50070    .SelLength = Len(.Text)
50080   End If
50090  End With
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmPrinting", "txtKeywords_GotFocus")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtKeywords_KeyPress(KeyAscii As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If KeyAscii = vbKeyReturn Then
50050   SaveEDoc
50060  End If
50070 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50080 Exit Sub
ErrPtnr_OnError:
50101 Select Case ErrPtnr.OnError("frmPrinting", "txtKeywords_KeyPress")
      Case 0: Resume
50120 Case 1: Resume Next
50130 Case 2: Exit Sub
50140 Case 3: End
50150 End Select
50160 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtModifyDate_GotFocus()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  With txtModifyDate
50050   If Len(.Text) > 0 Then
50060    .SelStart = 0
50070    .SelLength = Len(.Text)
50080   End If
50090  End With
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmPrinting", "txtModifyDate_GotFocus")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtModifyDate_KeyPress(KeyAscii As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If KeyAscii = vbKeyReturn Then
50050   SaveEDoc
50060  End If
50070 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50080 Exit Sub
ErrPtnr_OnError:
50101 Select Case ErrPtnr.OnError("frmPrinting", "txtModifyDate_KeyPress")
      Case 0: Resume
50120 Case 1: Resume Next
50130 Case 2: Exit Sub
50140 Case 3: End
50150 End Select
50160 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtSubject_GotFocus()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  With txtSubject
50050   If Len(.Text) > 0 Then
50060    .SelStart = 0
50070    .SelLength = Len(.Text)
50080   End If
50090  End With
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmPrinting", "txtSubject_GotFocus")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtSubject_KeyPress(KeyAscii As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If KeyAscii = vbKeyReturn Then
50050   SaveEDoc
50060  End If
50070 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50080 Exit Sub
ErrPtnr_OnError:
50101 Select Case ErrPtnr.OnError("frmPrinting", "txtSubject_KeyPress")
      Case 0: Resume
50120 Case 1: Resume Next
50130 Case 2: Exit Sub
50140 Case 3: End
50150 End Select
50160 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtTitle_GotFocus()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  With txtTitle
50050   If Len(.Text) > 0 Then
50060    .SelStart = 0
50070    .SelLength = Len(.Text)
50080   End If
50090  End With
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmPrinting", "txtTitle_GotFocus")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If KeyAscii = vbKeyReturn Then
50050   SaveEDoc
50060  End If
50070 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50080 Exit Sub
ErrPtnr_OnError:
50101 Select Case ErrPtnr.OnError("frmPrinting", "txtTitle_KeyPress")
      Case 0: Resume
50120 Case 1: Resume Next
50130 Case 2: Exit Sub
50140 Case 3: End
50150 End Select
50160 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SaveEDoc()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim PDFFile As String
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
50160   If Options.RunProgramAfterSaving = 1 Then
50170    If Options.OnePagePerFile = 1 Then
50180     PDFFile = Replace$(PDFFile, "%d", "1", , , vbTextCompare)
50190    End If
50200    RunProgramAfterSaving Me.hwnd, GetShortName(PDFFile), GetRunProgramAfterSavingProgramParameters, Options.RunProgramAfterSavingWindowstyle
50210   End If
50220   If chkStartStandardProgram.Value = 1 Then
50230    If Options.OnePagePerFile = 1 Then
50240      OpenDocument Replace$(PDFFile, "%d", "1", , , vbTextCompare)
50250     Else
50260      OpenDocument PDFFile
50270    End If
50280   End If
50290   Unload Me
50300  End If
50310 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50320 Exit Sub
ErrPtnr_OnError:
50341 Select Case ErrPtnr.OnError("frmPrinting", "SaveEDoc")
      Case 0: Resume
50360 Case 1: Resume Next
50370 Case 2: Exit Sub
50380 Case 3: End
50390 End Select
50400 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

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
  SaveFilename = ReplaceForbiddenChars(txtTitle.Text)
  
  Set Files = GetFilename(SaveFilename, Options.LastSaveDirectory, FilterIndex, Filter, SaveFile, Cancel, Me)
  If SaveOpenCancel = True Then
   Exit Function
  End If
  If Files.Count <> 1 Then
   Exit Function
  End If
  SaveFilterIndex = FilterIndex
  SaveFilename = Files.item(1)
  If FileExists(Files.item(1)) = True Then
   If FileInUse(Files.item(1)) = True Then
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
 SaveOptions Options
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
 Create_eDoc = OutputFile
 KillFile PDFSpoolfile
 Screen.MousePointer = vbNormal
 Me.Visible = False
 If Options.ShowAnimation = 1 Then
  ShowAnimation False
 End If
 frmMain.Timer1.Enabled = True
 Exit Function
ErrorHandler:
 Screen.MousePointer = vbNormal
 tErrNumber = Err.number
 tStr = Err.number & ", " & Err.Description
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
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tL As Long, BorderWidth As Long
50050  DoEvents
50060  lblTitle.Visible = Not Show
50070  lblCreationDate.Visible = Not Show
50080  lblCreateFor.Visible = Not Show
50090  txtTitle.Visible = Not Show
50100  txtCreationDate.Visible = Not Show
50110  txtCreateFor.Visible = Not Show
50120
50130  cmdWaiting.Visible = Not Show
50140  cmdEMail.Visible = Not Show
50150  cmdSave.Visible = Not Show
50160  chkStartStandardProgram.Visible = Not Show
50170  cmdNow(0).Visible = Not Show: cmdNow(1).Visible = Not Show
50180
50190  anmProcess.Visible = Show
50200
50210  lblStatus.Visible = Show
50220  Dim n As FormBorderStyleConstants
50230
50240  If Show = True Then
50250    ResAnimate anmProcess, ranOpen, 100
50260    With Me
50270     BorderWidth = 3
50280     anmProcess.Left = BorderWidth * Screen.TwipsPerPixelX
50290     anmProcess.Top = BorderWidth * Screen.TwipsPerPixelY
50300     .Height = anmProcess.Height + 380 + 2 * BorderWidth * Screen.TwipsPerPixelY
50310     .Width = anmProcess.Width + 4 * BorderWidth * Screen.TwipsPerPixelX
50320     .BorderStyle = vbBSNone
50330     .Caption = .Caption
50340     tL = .Width
50350     .Width = tL - Screen.TwipsPerPixelX
50360     .Width = tL
50370    End With
50380    DrawBorder3D Me, 4, BorderWidth
50390    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
50400    ResAnimate anmProcess, ranPlay
50410   Else
50420    ResAnimate anmProcess, ranStop
50430    ResAnimate anmProcess, ranClose
50440    Me.Height = 2520
50450  End If
50460 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50470 Exit Sub
ErrPtnr_OnError:
50491 Select Case ErrPtnr.OnError("frmPrinting", "ShowAnimation")
      Case 0: Resume
50510 Case 1: Resume Next
50520 Case 2: Exit Sub
50530 Case 3: End
50540 End Select
50550 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

