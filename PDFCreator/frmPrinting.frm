VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrinting 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "PDFCreator"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "frmPrinting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fraProfile 
      Caption         =   "Profile"
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   7575
      Begin VB.ComboBox cmbProfile 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   24
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   105
      TabIndex        =   15
      Top             =   5880
      Width           =   1350
   End
   Begin VB.CommandButton cmdCollect 
      Caption         =   "&Wait - Collect"
      Height          =   495
      Left            =   1680
      TabIndex        =   16
      Top             =   5880
      Width           =   1350
   End
   Begin VB.CommandButton cmdNow 
      Caption         =   "&Now"
      Height          =   300
      Index           =   0
      Left            =   6465
      TabIndex        =   4
      Top             =   1073
      Width           =   1260
   End
   Begin VB.TextBox txtKeywords 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   105
      TabIndex        =   13
      Top             =   3960
      Width           =   7620
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   105
      TabIndex        =   11
      Top             =   3240
      Width           =   7620
   End
   Begin VB.TextBox txtModifyDate 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   105
      TabIndex        =   6
      Top             =   1800
      Width           =   6240
   End
   Begin MSComCtl2.Animation anmProcess 
      Height          =   960
      Left            =   1935
      TabIndex        =   20
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
      Appearance      =   0  '2D
      Height          =   285
      Left            =   105
      TabIndex        =   9
      Top             =   2520
      Width           =   7620
   End
   Begin VB.TextBox txtCreationDate 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   105
      TabIndex        =   3
      Top             =   1080
      Width           =   6240
   End
   Begin VB.CheckBox chkStartStandardProgram 
      Appearance      =   0  '2D
      Caption         =   "After saving open the document with the standardprogram."
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   105
      TabIndex        =   14
      Top             =   5160
      Width           =   7620
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   105
      TabIndex        =   1
      Top             =   360
      Width           =   7620
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   21
      Top             =   0
      Width           =   0
   End
   Begin VB.CommandButton cmdNow 
      Caption         =   "&Now"
      Height          =   300
      Index           =   1
      Left            =   6465
      TabIndex        =   7
      Top             =   1793
      Width           =   1260
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options"
      Height          =   495
      Left            =   3255
      TabIndex        =   17
      Top             =   5880
      Width           =   1350
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "&eMail"
      Height          =   495
      Left            =   4830
      TabIndex        =   18
      Top             =   5880
      Width           =   1350
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   495
      Left            =   6375
      TabIndex        =   19
      Top             =   5880
      Width           =   1350
   End
   Begin VB.Label lblKeywords 
      AutoSize        =   -1  'True
      Caption         =   "Keywords:"
      Height          =   195
      Left            =   105
      TabIndex        =   12
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   105
      TabIndex        =   10
      Top             =   3000
      Width           =   585
   End
   Begin VB.Label lblModifyDate 
      AutoSize        =   -1  'True
      Caption         =   "Modify Date:"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label lblCreateFor 
      AutoSize        =   -1  'True
      Caption         =   "Author:"
      Height          =   195
      Left            =   105
      TabIndex        =   8
      Top             =   2280
      Width           =   510
   End
   Begin VB.Label lblCreationDate 
      AutoSize        =   -1  'True
      Caption         =   "Creation Date:"
      Height          =   195
      Left            =   105
      TabIndex        =   2
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Document Title:"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Zentriert
      Caption         =   "Creating file..."
      Height          =   255
      Left            =   480
      TabIndex        =   22
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
Private OldProfile As Long
Public PrinterProfile As String

Private Sub chkStartStandardProgram_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkStartStandardProgram.value = 1 Then
50020    Options.StartStandardProgram = 1
50030   Else
50040    Options.StartStandardProgram = 0
50050  End If
50060  If cmbProfile.ListIndex = 0 Then
50070    SaveOption Options, "StartStandardProgram"
50080   Else
50090    SaveOption Options, "StartStandardProgram", cmbProfile.List(cmbProfile.ListIndex)
50100  End If
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

Private Sub cmbProfile_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tOpt As tOptions
50020  If cmbProfile.ListIndex <> OldProfile Then
50030   OldProfile = cmbProfile.ListIndex
50040   If cmbProfile.ListIndex = 0 Then
50050     Options = ReadOptions
50060    Else
50070     tOpt = Options
50080     Options = ReadOptions(, , cmbProfile.List(cmbProfile.ListIndex))
50090     Options.LastUpdateCheck = tOpt.LastUpdateCheck
50100     Options.UpdateInterval = tOpt.UpdateInterval
50110   End If
50120   InitForm
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "cmbProfile_Click")
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
50010  If FileExists(CurrentInfoSpoolFile) = False Then
50020   Exit Sub
50030  End If
50040  KillInfoSpoolFiles CurrentInfoSpoolFile
50050  Unload Me
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
50290  KillInfoSpoolFiles CurrentInfoSpoolFile
50300
50310  Options.Counter = Options.Counter + 1
50320
50330  Me.Visible = False
50340  Unload Me
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
50030    txtCreationDate.Text = Format(CStr(Now), "YYYYMMDDHHNNSS")
50040   Case 1:
50050    txtModifyDate.Text = Format(CStr(Now), "YYYYMMDDHHNNSS")
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
50020   If cmbProfile.ListIndex = 0 Then
50030     CurrentPrinterProfile = ""
50040    Else
50050     CurrentPrinterProfile = cmbProfile.List(cmbProfile.ListIndex)
50060   End If
50070   frmOptions.Show vbModal, Me
50080   UpdateProfiles
50090  End If
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

Private Sub cmdCollect_Click()
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
Select Case ErrPtnr.OnError("frmPrinting", "cmdCollect_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UpdateProfiles()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Profiles As Collection, profile As Variant, i As Long, ppi As Long
50020  cmbProfile.Clear
50030  cmbProfile.AddItem LanguageStrings.OptionsProfileDefaultName
50040  Set Profiles = GetProfiles
50050  For Each profile In Profiles
50060   cmbProfile.AddItem profile
50070  Next profile
50080  For i = 0 To cmbProfile.ListCount - 1
50090   If StrComp(cmbProfile.List(i), PrinterProfile, vbTextCompare) = 0 Then
50100    ppi = i
50110    Exit For
50120   End If
50130  Next i
50140
50150  cmbProfile.ListIndex = ppi
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "UpdateProfiles")
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
50030   Call HTMLHelp_ShowTopic("html\pdfcreator-user-manual.html")
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
50010  Me.Icon = LoadResPicture(2120, vbResIcon)
50020  Me.KeyPreview = True
50030  Caption = App.EXEName
50040  Caption = App.title & " " & GetProgramReleaseStr
50050  Printing = True
50060
50070  With anmProcess
50080   .Top = 0
50090   .Left = 0
50100   .Width = 260 * Screen.TwipsPerPixelX
50110   .Height = 66 * Screen.TwipsPerPixelY
50120  End With
50130
50140  If frmMain.Visible = False Then
50150   FormInTaskbar Me, True, True
50160  End If
50170
50180  ChangeLanguage
50190
50200  InitForm
50210
50220  UpdateProfiles
50230
50240  With Options
50250   SetFontControls Me.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50260  End With
50270
50280  ShowAcceleratorsInForm Me, True
50290  SetTopMost Me, True, True
50300  SetTopMost Me, False, True
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

Private Sub InitForm()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tDate As Date, tStr As String
50020  If Options.StartStandardProgram = 1 Then
50030    chkStartStandardProgram.value = 1
50040   Else
50050    chkStartStandardProgram.value = 0
50060  End If
50070  PSHeader = GetPSHeader(PDFSpoolfile)
50080  With PSHeader
50090   If Len(Trim$(Options.StandardTitle)) > 0 Then
50100     txtTitle.Text = GetSubstFilename(CurrentInfoSpoolFile, _
     RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardTitle)), , , True)
50120    Else
50130     txtTitle.Text = GetSubstFilename(CurrentInfoSpoolFile, Options.SaveFilename, , , True)
50140   End If
50150   If Options.UseStandardAuthor = 1 Then
50160     txtCreateFor.Text = GetSubstFilename(CurrentInfoSpoolFile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardAuthor)), True, , True)
50170    Else
50180     txtCreateFor.Text = GetDocUsername(CurrentInfoSpoolFile, False)
50190   End If
50200   If Len(Trim$(Options.StandardKeywords)) > 0 Then
50210    txtKeywords.Text = GetSubstFilename(CurrentInfoSpoolFile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardKeywords)), , , True)
50220   End If
50230   If Len(Trim$(Options.StandardSubject)) > 0 Then
50240    txtSubject.Text = GetSubstFilename(CurrentInfoSpoolFile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardSubject)), , , True)
50250   End If
50260
50270   tDate = Now
50280   If LenB(PSHeader.CreationDate.Comment) > 0 Then
50290     tStr = FormatPrintDocumentDate(PSHeader.CreationDate.Comment)
50300    Else
50310     tStr = CStr(tDate)
50320   End If
50330   txtCreationDate.Text = GetDocDate(Options.StandardCreationdate, Options.StandardDateformat, tStr)
50340
50350   txtModifyDate.Text = GetDocDate(Options.StandardModifydate, Options.StandardDateformat, tStr)
50360  End With
50370  If Options.OptionsEnabled = 0 Or FormISLoaded("frmOptions") = True Then
50380   cmdOptions.Enabled = False
50390  End If
50400  If Options.OptionsVisible = 0 Then
50410   cmdOptions.Visible = False
50420  End If
50430  If Options.DisableEmail = 1 Then
50440   cmdEMail.Enabled = False
50450  End If
50460  Height = cmdCollect.Top + cmdCollect.Height + (Height - ScaleHeight) + 100
50470  With txtTitle
50480   .SelStart = 0
50490   .SelLength = Len(.Text)
50500  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "InitForm")
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

Private Function SaveEDoc() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PDFFile As String, tOpt As tOptions
50020
50030  SaveEDoc = False
50040  IsConverted = False
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
50270
50280   If cmbProfile.ListIndex = 0 Then
50290     Options.Counter = Options.Counter + 1
50300     SaveOption Options, "Counter"
50310    Else
50320     tOpt = ReadOptions(True, , cmbProfile.List(cmbProfile.ListIndex))
50330     tOpt.Counter = tOpt.Counter + 1
50340     SaveOption tOpt, "Counter", cmbProfile.List(cmbProfile.ListIndex)
50350   End If
50360
50370   If chkStartStandardProgram.value = 1 Then
50380    If Options.OnePagePerFile = 1 Then
50390      OpenDocument Replace$(PDFFile, "%d", "1", , , vbTextCompare)
50400     Else
50410      OpenDocument PDFFile
50420    End If
50430   End If
50440   IsConverted = True
50450   KillInfoSpoolFiles CurrentInfoSpoolFile
50460  End If
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  On Error GoTo ErrorHandler
50020  Dim OutputFile As String, Path As String, tStr As String, Filter As String, _
  tErrNumber As Long, FilterIndex As Long, Cancel As Boolean, _
  PDFDocInfo As tPDFDocInfo, files As Collection, Extf(12) As String
50050  Dim PDFDocInfoFile As String, StampFile As String
50060
50070  Extf(0) = "*.pdf"
50080  Extf(1) = "*.png"
50090  Extf(2) = "*.jpg"
50100  Extf(3) = "*.bmp"
50110  Extf(4) = "*.pcx"
50120  Extf(5) = "*.tif"
50130  Extf(6) = "*.ps"
50140  Extf(7) = "*.eps"
50150  Extf(8) = "*.txt"
50160  Extf(9) = "*.psd"
50170  Extf(10) = "*.pcl"
50180  Extf(11) = "*.raw"
50190  Extf(12) = "*.svg"
50200  With LanguageStrings
50210   Filter = .ListPDFFiles & " (" & Extf(0) & ")|" & Extf(0) & "|" & _
   .PrintingPDFAFiles & " (" & Extf(0) & ")|" & Extf(0) & "|" & _
   .PrintingPDFXFiles & " (" & Extf(0) & ")|" & Extf(0) & "|" & _
   .PrintingPNGFiles & " (" & Extf(1) & ")|" & Extf(1) & "|" & _
   .PrintingJPEGFiles & " (" & Extf(2) & ")|" & Extf(2) & "|" & _
   .PrintingBMPFiles & " (" & Extf(3) & ")|" & Extf(3) & "|" & _
   .PrintingPCXFiles & " (" & Extf(4) & "|" & Extf(4) & "|" & _
   .PrintingTIFFFiles & " (" & Extf(5) & ")|" & Extf(5) & "|" & _
   .PrintingPSFiles & " (" & Extf(6) & ")|" & Extf(6) & "|" & _
   .PrintingEPSFiles & " (" & Extf(7) & ")|" & Extf(7) & "|" & _
   .PrintingTXTFiles & " (" & Extf(8) & ")|" & Extf(8) & "|" & _
   .PrintingPSDFiles & " (" & Extf(9) & ")|" & Extf(9) & "|" & _
   .PrintingPCLFiles & " (" & Extf(10) & ")|" & Extf(10) & "|" & _
   .PrintingRAWFiles & " (" & Extf(11) & ")|" & Extf(11) & "|" & _
   .PrintingSVGFiles & " (" & Extf(12) & ")|" & Extf(12)
50360  End With
50370
50380  FilterIndex = Options.StandardSaveformat + 1
50390  With LanguageStrings
50400   PSHeader = GetPSHeader(PDFSpoolfile)
50410   If Len(txtTitle.Text) > 0 And Options.RemoveAllKnownFileExtensions = 1 Then
50420     SaveFilename = ReplaceForbiddenChars(RemoveAllKnownFileExtensions(txtTitle.Text))
50430    Else
50440     SaveFilename = ReplaceForbiddenChars(txtTitle.Text)
50450   End If
50460   Set files = GetFilename(SaveFilename, GetSubstFilename2(Options.LastSaveDirectory), FilterIndex, Filter, SaveFile, Cancel, Me)
50470   If SaveOpenCancel = True Then
50480    Exit Function
50490   End If
50500   If files.Count <> 1 Then
50510    Exit Function
50520   End If
50530   SaveFilterIndex = FilterIndex
50540   SaveFilename = files.Item(1)
50550   If FileExists(files.Item(1)) = True Then
50560    If FileInUse(files.Item(1)) = True Then
50570     MsgBox LanguageStrings.MessagesMsg34
50580     Exit Function
50590    End If
50600   End If
50610  End With
50620
50630  Screen.MousePointer = vbHourglass
50640  DoEvents
50650
50660  Options.StartStandardProgram = chkStartStandardProgram.value
50670  OutputFile = Trim$(SaveFilename)
50680  SplitPath OutputFile, , Path
50690  Options.LastSaveDirectory = Path
50700  If cmbProfile.ListIndex = 0 Then
50710    SaveOption Options, "LastSaveDirectory"
50720   Else
50730    SaveOption Options, "LastSaveDirectory", cmbProfile.List(cmbProfile.ListIndex)
50740  End If
50750  frmMain.SetSystrayIcon 3
50760  If Options.ShowAnimation = 1 Then
50770    ShowAnimation True
50780   Else
50790    Me.Visible = False
50800  End If
50810  DoEvents
50820
50830  PSHeader.CreateFor.Comment = txtCreateFor.Text
50840  PSHeader.CreationDate.Comment = txtCreationDate.Text
50850  PSHeader.Creator.Comment = App.EXEName & _
  " Version " & App.Major & "." & App.Minor & "." & App.Revision
50870  PSHeader.title.Comment = txtTitle.Text
50880
50890  With PDFDocInfo
50900   .Author = txtCreateFor.Text
50910   .CreationDate = txtCreationDate.Text
50920   .Creator = App.EXEName & " Version " & App.Major & "." & App.Minor & "." & App.Revision
50930   .Keywords = GetSubstFilename(PDFSpoolfile, txtKeywords.Text, , , True)
50940   .ModifyDate = txtModifyDate.Text
50950   .Subject = GetSubstFilename(PDFSpoolfile, txtSubject.Text, , , True)
50960   .title = GetSubstFilename(PDFSpoolfile, txtTitle.Text, , , True)
50970  End With
50980
50990  PDFDocInfoFile = CreatePDFDocInfoFile(CurrentInfoSpoolFile, PDFDocInfo)
51000  'AppendPDFDocInfo PDFSpoolfile, PDFDocInfo
51010
51020  StampFile = CreateStampFile(CurrentInfoSpoolFile)
51030  'CheckForStamping PDFSpoolfile
51040
51050
51060  If Options.RunProgramBeforeSaving = 1 Then
51070   RunProgramBeforeSaving Me.hwnd, PDFSpoolfile, _
  Options.RunProgramBeforeSavingProgramParameters, _
  Options.RunProgramBeforeSavingWindowstyle
51100  End If
51111  Select Case SaveFilterIndex
        Case 1:
51130    'CallGScript PDFSpoolfile, OutputFile, Options, PDFWriter
51140    CallGScript CurrentInfoSpoolFile, OutputFile, Options, PDFWriter, PDFDocInfoFile, StampFile
51150   Case 2:
51160    CallGScript CurrentInfoSpoolFile, OutputFile, Options, PDFAWriter, PDFDocInfoFile, StampFile
51170   Case 3:
51180    CallGScript CurrentInfoSpoolFile, OutputFile, Options, PDFXWriter, PDFDocInfoFile, StampFile
51190   Case 4:
51200    CallGScript CurrentInfoSpoolFile, OutputFile, Options, PNGWriter, PDFDocInfoFile, StampFile
51210   Case 5:
51220    CallGScript CurrentInfoSpoolFile, OutputFile, Options, JPEGWriter, PDFDocInfoFile, StampFile
51230   Case 6:
51240    CallGScript CurrentInfoSpoolFile, OutputFile, Options, BMPWriter, PDFDocInfoFile, StampFile
51250   Case 7:
51260    CallGScript CurrentInfoSpoolFile, OutputFile, Options, PCXWriter, PDFDocInfoFile, StampFile
51270   Case 8:
51280    CallGScript CurrentInfoSpoolFile, OutputFile, Options, TIFFWriter, PDFDocInfoFile, StampFile
51290   Case 9:
51300    CallGScript CurrentInfoSpoolFile, OutputFile, Options, PSWriter, PDFDocInfoFile, StampFile
51310   Case 10:
51320    CallGScript CurrentInfoSpoolFile, OutputFile, Options, EPSWriter, PDFDocInfoFile, StampFile
51330   Case 11:
51340    CallGScript CurrentInfoSpoolFile, OutputFile, Options, TXTWriter, PDFDocInfoFile, StampFile
51350   Case 12:
51360    CallGScript CurrentInfoSpoolFile, OutputFile, Options, PSDWriter, PDFDocInfoFile, StampFile
51370   Case 13:
51380    CallGScript CurrentInfoSpoolFile, OutputFile, Options, PCLWriter, PDFDocInfoFile, StampFile
51390   Case 14:
51400    CallGScript CurrentInfoSpoolFile, OutputFile, Options, RAWWriter, PDFDocInfoFile, StampFile
51410   Case 15:
51420    CallGScript CurrentInfoSpoolFile, OutputFile, Options, SVGWriter, PDFDocInfoFile, StampFile
51430  End Select
51440  Create_eDoc = OutputFile
51450  CheckForPrintingAfterSaving CurrentInfoSpoolFile, Options
51460  Screen.MousePointer = vbNormal
51470  Me.Visible = False
51480  If Options.ShowAnimation = 1 Then
51490   ShowAnimation False
51500  End If
51510  ConvertedOutputFilename = OutputFile
51520
51530  ReadyConverting = True
51540  frmMain.SetSystrayIcon 2
51550  Exit Function
ErrorHandler:
51570  Screen.MousePointer = vbNormal
51580  tErrNumber = Err.Number
51590  tStr = "Methode ""Create_eDoc"": " & Err.Number & ", " & Err.Description
51600  On Error GoTo ErrPtnr_OnError
51610  On Error Resume Next
51620  Me.Hide
51630  If tErrNumber <> 32755 Then
51640   KillFile PDFSpoolfile
51650   IfLoggingWriteLogfile "Error: " & tStr
51660   IfLoggingShowLogfile frmLog, frmMain
51670  End If
51680  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "Create_eDoc")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50100  cmdCollect.Visible = Not Show
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

Public Sub ChangeLanguage()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With LanguageStrings
50020   lblTitle.Caption = .PrintingDocumentTitle
50030   lblStatus.Caption = .PrintingStatus
50040   lblCreationDate.Caption = .PrintingCreationDate
50050   lblCreateFor.Caption = .PrintingAuthor
50060   lblModifyDate.Caption = .PrintingModifyDate
50070   lblSubject.Caption = .PrintingSubject
50080   lblKeywords.Caption = .PrintingKeywords
50090   chkStartStandardProgram.Caption = .PrintingStartStandardProgram
50100   cmdCollect.Caption = .PrintingCollect
50110   cmdOptions.Caption = .DialogPrinterOptions
50120   cmdEMail.Caption = .PrintingEMail
50130   cmdSave.Caption = .PrintingSave
50140   cmdNow(0).Caption = .PrintingNow
50150   cmdNow(1).Caption = .PrintingNow
50160   If LenB(.PrintingCancel) = 0 Then
50170     cmdCancel.Caption = .OptionsCancel
50180    Else
50190     cmdCancel.Caption = .PrintingCancel
50200   End If
50210   fraProfile.Caption = .PrintingProfile
50220  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "ChangeLanguage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
