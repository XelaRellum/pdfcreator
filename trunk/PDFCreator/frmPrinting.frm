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
      Left            =   120
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
50140  PDFFile = Trim$(Create_eDoc)
50150
50160  If Len(PDFFile) > 0 And FileExists(PDFFile) = True Then
50170   If Options.RunProgramAfterSaving = 1 Then
50180    RunProgramAfterSaving GetShortName(PDFFile), GetRunProgramAfterSavingProgramParameters, Options.RunProgramAfterSavingWindowstyle
50190   End If
50200   Set mail = New clsPDFCreatorMail
50210   If mail.Send(PDFFile) <> 0 Then
50220    MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
50230   End If
50240   Set mail = Nothing
50250  End If
50260
50270  Me.Visible = False
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
50010  SaveEDoc
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
 Dim OutputFile As String, Path As String, tStr As String, Filter As String, _
  tErrNumber As Long, Filename As String, Filterindex As Long, _
  Cancel As Boolean, PDFDocInfo As tPDFDocInfo, Files As Collection, _
  Ext As String

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
  SaveFilename = ReplaceForbiddenChars(txtTitle.Text)
  If InStr(SaveFilename, ".") > 0 Then
   SaveFilename = Mid(SaveFilename, 1, InStr(SaveFilename, ".") - 1)
  End If
  Set Files = GetFilename(SaveFilename, Options.LastSaveDirectory, Filterindex, Filter, SaveFile, Cancel, Me)
  If SaveOpenCancel = True Then
   Exit Function
  End If
  If Files.Count <> 1 Then
   Exit Function
  End If
  SaveFilterIndex = Filterindex
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
50140  cmdNow(0).Visible = Not Show: cmdNow(1).Visible = Not Show
50150
50160  anmProcess.Visible = Show
50170
50180  lblStatus.Visible = Show
50190  Dim n As FormBorderStyleConstants
50200
50210  If Show = True Then
50220    ResAnimate anmProcess, ranOpen, 100
50230    With Me
50240     BorderWidth = 3
50250     anmProcess.Left = BorderWidth * Screen.TwipsPerPixelX
50260     anmProcess.Top = BorderWidth * Screen.TwipsPerPixelY
50270     .Height = anmProcess.Height + 380 + 2 * BorderWidth * Screen.TwipsPerPixelY
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
50010  Dim tDate As Date, tStr As String, Path As String, Filename As String
50020  Me.KeyPreview = True
50030  Caption = App.EXEName
50040  Me.Icon = frmMain.Icon
50050  Caption = App.Title & " " & GetProgramReleaseStr ' & " " & LanguageStrings.CommonTitle
50060  Printing = True
50070  RemoveX Me
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
50340  End With
50350  If Options.StartStandardProgram = 1 Then
50360    chkStartStandardProgram.Value = 1
50370   Else
50380    chkStartStandardProgram.Value = 0
50390  End If
50400  PSHeader = GetPSHeader(PDFSpoolfile)
50410  With PSHeader
50420   If Len(Trim$(Options.StandardTitle)) > 0 Then
50430     txtTitle.Text = GetSubstFilename(PDFSpoolfile, _
     RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardTitle)))
50450    Else
50460     txtTitle.Text = GetSubstFilename(PDFSpoolfile, Options.SaveFilename)
50470   End If
50480   SplitPath txtTitle.Text, , Path, , Filename
50490   If Len(Filename) > 0 Then
50500    txtTitle.Text = Filename
50510   End If
50520   If Options.UseStandardAuthor = 1 Then
50530     txtCreateFor.Text = GetSubstFilename(PDFSpoolfile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardAuthor)), True)
50540    Else
50550     txtCreateFor.Text = GetDocUsername(PDFSpoolfile, False)
50560   End If
50570   If Len(Trim$(Options.StandardKeywords)) > 0 Then
50580    txtKeywords.Text = GetSubstFilename(PDFSpoolfile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardKeywords)))
50590   End If
50600   If Len(Trim$(Options.StandardSubject)) > 0 Then
50610    txtSubject.Text = GetSubstFilename(PDFSpoolfile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardSubject)))
50620   End If
50630
50640   tDate = Now
50650   If LenB(PSHeader.CreationDate.Comment) > 0 Then
50660     tStr = FormatPrintDocumentDate(PSHeader.CreationDate.Comment)
50670    Else
50680     tStr = CStr(tDate)
50690   End If
50700   txtCreationDate.Text = GetDocDate(Options.StandardCreationdate, Options.StandardDateformat, FormatPrintDocumentDate(tStr))
50710   'tStr = CStr(tDate)
50720   txtModifyDate.Text = GetDocDate(Options.StandardModifydate, Options.StandardDateformat, FormatPrintDocumentDate(tStr))
50730  End With
50740  If Options.OptionsEnabled = 0 Then
50750   cmdOptions.Enabled = False
50760  End If
50770  If Options.OptionsVisible = 0 Then
50780   cmdOptions.Visible = False
50790  End If
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

Private Sub DrawBorder3D(ByVal obj As Object, Index As Integer, BorderWidth As Long)
 Dim r As Rect, DrawObj As Object, tRedraw As Boolean, tDrawwidth As Integer, _
  tScalemode As Long, RectType As Long, RectStyle As Long

 On Error Resume Next
 If BorderWidth <= 0 Then
  Exit Sub
 End If
 RectStyle = BF_RECT
 RectType = Choose(Index, EDGE_RAISED, EDGE_SUNKEN, EDGE_ETCHED, EDGE_BUMP)
 If TypeOf obj Is Form Or TypeOf obj Is PictureBox Then
   With obj
    tScalemode = .ScaleMode
    .ScaleMode = vbPixels
    r.Left = .ScaleLeft
    r.Top = .ScaleTop
    r.Right = .ScaleWidth
    r.Bottom = .ScaleHeight
   End With
   Set DrawObj = obj
  Else
   With obj
    r.Left = .Left - BorderWidth
    r.Top = .Top - BorderWidth
    r.Right = r.Left + .Width + BorderWidth + 1
    r.Bottom = r.Top + .Height + BorderWidth + 1
   End With
   Set DrawObj = obj.Container
 End If
 With DrawObj
  tRedraw = .AutoRedraw
  tDrawwidth = .DrawWidth
  .DrawWidth = BorderWidth
  .AutoRedraw = True
  DrawEdge .hDC, r, RectType, RectStyle
  .AutoRedraw = tRedraw
  .DrawWidth = tDrawwidth
  .Refresh
  .ScaleMode = tScalemode
 End With
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

Private Sub SaveEDoc()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PDFFile As String
50020
50030  If GsDllLoaded = 0 Then
50040   MsgBox LanguageStrings.MessagesMsg08
50050   SetPrinterStop True
50060   frmMain.Visible = True
50070   Unload Me
50080   Exit Sub
50090  End If
50100
50110  PDFFile = Trim$(Create_eDoc)
50120  If PDFFile <> vbNullString Then
50130   If Options.RunProgramAfterSaving = 1 Then
50140    If Options.OnePagePerFile = 1 Then
50150     PDFFile = Replace$(PDFFile, "%d", "1", , , vbTextCompare)
50160    End If
50170    RunProgramAfterSaving GetShortName(PDFFile), GetRunProgramAfterSavingProgramParameters, Options.RunProgramAfterSavingWindowstyle
50180   End If
50190   If chkStartStandardProgram.Value = 1 Then
50200    If Options.OnePagePerFile = 1 Then
50210      OpenDocument Replace$(PDFFile, "%d", "1", , , vbTextCompare)
50220     Else
50230      OpenDocument PDFFile
50240    End If
50250   End If
50260   Unload Me
50270  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinting", "SaveEDoc")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
