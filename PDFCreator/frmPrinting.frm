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
      Caption         =   "Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label lblModifyDate 
      Caption         =   "Modify Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lblCreateFor 
      Caption         =   "Author:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2280
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

Private SaveFilename As String, SaveFilterIndex As Long

Private PSHeader As tPSHeader

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

 ' UnloadDLLComplete GsDllLoaded
 ' GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)

 If GsDllLoaded = 0 Then
  MsgBox LanguageStrings.MessagesMsg08
  SetPrinterStop True
  frmMain.Visible = True
  Unload Me
  Exit Sub
 End If

 Set mail = New clsPDFCreatorMail

 PDFFile = Trim$(Create_eDoc)

 ' UnloadDLLComplete GsDllLoaded

 If Dir(PDFFile) <> "" And Len(Trim$(PDFFile)) > 0 Then
  If mail.Send(PDFFile) <> 0 Then
   MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
  End If
 End If

 Me.Visible = False
 Set mail = Nothing
 Unload Me
End Sub


Private Sub cmdNow_Click(Index As Integer)
 Select Case Index
  Case 0:
   txtCreationDate.Text = CStr(Now)
  Case 1:
   txtModifyDate.Text = CStr(Now)
 End Select
End Sub

Private Sub cmdOptions_Click()
 frmOptions.Show , Me
End Sub

Private Sub cmdSave_Click()
 Dim PDFFile As String

 'UnloadDLLComplete GsDllLoaded
 'GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)

 If GsDllLoaded = 0 Then
  MsgBox LanguageStrings.MessagesMsg08
  SetPrinterStop True
  frmMain.Visible = True
  Unload Me
  Exit Sub
 End If
 If LenB(Dir(CompletePath(App.Path) & "pdfenc.exe")) > 0 Or _
  GhostScriptSecurity = True Then
   SecurityIsPossible = True
 End If
 'GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)

 If LenB(Dir(CompletePath(App.Path) & "pdfenc.exe")) > 0 Or _
  GhostScriptSecurity = True Then
   SecurityIsPossible = True
 End If

 PDFFile = Trim$(Create_eDoc)

 'UnloadDLLComplete GsDllLoaded


 If PDFFile <> vbNullString Then Unload Me
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
 cmdNow(0).Visible = Not Show: cmdNow(1).Visible = Not Show

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF1 Then
  KeyCode = 0
  Call HTMLHelp_ShowTopic("html\welcome.htm")
 End If
End Sub

Private Sub Form_Load()
 Me.KeyPreview = True
 Caption = App.Title & " " & GetProgramReleaseStr ' & " " & LanguageStrings.CommonTitle
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
  lblModifyDate.Caption = .PrintingModifyDate
  lblSubject.Caption = .PrintingSubject
  lblKeywords.Caption = .PrintingKeywords
  chkStartStandardProgram.Caption = .PrintingStartStandardProgram
  cmdWaiting.Caption = .PrintingWaiting
  cmdOptions.Caption = .DialogPrinterOptions
  cmdEMail.Caption = .PrintingEMail
  cmdSave.Caption = .PrintingSave
  cmdNow(0).Caption = .PrintingNow
  cmdNow(1).Caption = .PrintingNow
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
'    txtCreateFor.Text = Options.StandardAuthor
    txtCreateFor.Text = GetSubstFilename(PDFSpoolfile, Options.StandardAuthor, True)
   Else
    txtCreateFor.Text = Trim$(.CreateFor.Comment)
  End If
  If Options.UseCreationDateNow = 1 Then
    txtCreationDate.Text = Now
   Else
    If IsDate(.CreationDate.Comment) = True Then
      txtCreationDate.Text = CStr(CDate(Trim$(.CreationDate.Comment)))
     Else
      txtCreationDate.Text = Trim$(.CreationDate.Comment)
    End If
  End If
  txtModifyDate.Text = txtCreationDate.Text
 End With
 If Options.OptionsEnabled = 0 Then
  cmdOptions.Enabled = False
 End If
 If Options.OptionsVisible = 0 Then
  cmdOptions.Visible = False
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 SetTopMost Me, False, False
 Printing = False
End Sub
