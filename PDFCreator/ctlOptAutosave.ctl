VERSION 5.00
Begin VB.UserControl ctlOptAutosave 
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ScaleHeight     =   5400
   ScaleWidth      =   6765
   ToolboxBitmap   =   "ctlOptAutosave.ctx":0000
   Begin PDFCreator.dmFrame dmFraProgAutosave 
      Height          =   5085
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8969
      Caption         =   "Autosave"
      Caption3D       =   2
      BarColorFrom    =   16744576
      BarColorTo      =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextShaddowColor=   12582912
      Begin VB.TextBox txtAutoSaveDirectoryPreview 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3450
         Width           =   6015
      End
      Begin VB.TextBox txtAutoSaveFilenamePreview 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2145
         Width           =   6015
      End
      Begin VB.CheckBox chkUseAutosave 
         Appearance      =   0  '2D
         Caption         =   "Use Autosave"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6015
      End
      Begin VB.CheckBox chkUseAutosaveDirectory 
         Appearance      =   0  '2D
         Caption         =   "For autosave use this directory"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   5895
      End
      Begin VB.TextBox txtAutosaveDirectory 
         Appearance      =   0  '2D
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   5535
      End
      Begin VB.TextBox txtAutosaveFilename 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "<DateTime>"
         Top             =   1800
         Width           =   3495
      End
      Begin VB.ComboBox cmbAutoSaveFilenameTokens 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "ctlOptAutosave.ctx":0312
         Left            =   3690
         List            =   "ctlOptAutosave.ctx":0314
         Style           =   2  'Dropdown-Liste
         TabIndex        =   7
         Top             =   1785
         Width           =   2460
      End
      Begin VB.ComboBox cmbAutosaveFormat 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdGetAutosaveDirectory 
         Caption         =   "..."
         Height          =   300
         Left            =   5760
         TabIndex        =   11
         Top             =   3120
         Width           =   375
      End
      Begin VB.CheckBox chkAutosaveStartStandardProgram 
         Appearance      =   0  '2D
         Caption         =   "After auto-saving open the document with the default program."
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   4095
         Width           =   5895
      End
      Begin VB.CheckBox chkAutosaveSendEmail 
         Appearance      =   0  '2D
         Caption         =   "Send an email after auto-saving"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   4680
         Width           =   5895
      End
      Begin VB.Label lblAutosaveFilename 
         AutoSize        =   -1  'True
         Caption         =   "Filename"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lblAutosaveformat 
         Caption         =   "Autosaveformat"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblAutosaveFilenameTokens 
         AutoSize        =   -1  'True
         Caption         =   "Add a Filename-Token"
         Height          =   195
         Left            =   3720
         TabIndex        =   5
         Top             =   1560
         Width           =   1605
      End
   End
End
Attribute VB_Name = "ctlOptAutosave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  dmFraProgAutosave.Left = 0
50030  dmFraProgAutosave.Top = 0
50040  UserControl.Height = dmFraProgAutosave.Height
50050  With cmbAutosaveFormat
50060   .Clear
50070   .AddItem "PDF"
50080   .AddItem "PDF/A-1b"
50090   .AddItem "PDF/X"
50100   .AddItem "PNG"
50110   .AddItem "JPEG"
50120   .AddItem "BMP"
50130   .AddItem "PCX"
50140   .AddItem "TIFF"
50150   .AddItem "PS"
50160   .AddItem "EPS"
50170   .AddItem "TXT"
50180   .AddItem "PSD"
50190   .AddItem "PCL"
50200   .AddItem "RAW"
50210   .AddItem "SVG"
50220   .ListIndex = 0
50230  End With
50240  With cmbAutoSaveFilenameTokens
50250   .Clear
50260   .AddItem "<Author>"
50270   .AddItem "<Computername>"
50280   .AddItem "<ClientComputer>"
50290   .AddItem "<DateTime>"
50300   .AddItem "<Title>"
50310   .AddItem "<Username>"
50320   .AddItem "<Counter>"
50330   .AddItem "<REDMON_DOCNAME>"
50340   .AddItem "<REDMON_DOCNAME_FILE>"
50350   .AddItem "<REDMON_DOCNAME_PATH>"
50360   .AddItem "<REDMON_JOB>"
50370   .AddItem "<REDMON_MACHINE>"
50380   .AddItem "<REDMON_PORT>"
50390   .AddItem "<REDMON_PRINTER>"
50400   .AddItem "<REDMON_SESSIONID>"
50410   .AddItem "<REDMON_USER>"
50420   .ListIndex = 0
50430  End With
50440
50450  SetFrames Options.OptionsDesign
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "UserControl_Initialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetFrames(OptionsDesign As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  For Each ctl In UserControl.Controls
50030   If TypeOf ctl Is dmFrame Then
50040    SetFrame ctl, OptionsDesign
50050   End If
50060  Next ctl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "SetFrames")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  dmFraProgAutosave.Width = UserControl.Width
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "UserControl_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetLanguageStrings()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With LanguageStrings
50020   dmFraProgAutosave.Caption = .OptionsProgramAutosaveSymbol
50030   chkUseAutosave.Caption = .OptionsUseAutosave
50040   lblAutosaveformat.Caption = .OptionsAutosaveFormat
50050   lblAutosaveFilename.Caption = .OptionsAutosaveFilename
50060   lblAutosaveFilenameTokens.Caption = .OptionsAutosaveFilenameTokens
50070   chkUseAutosaveDirectory.Caption = .OptionsUseAutosaveDirectory
50080   chkAutosaveStartStandardProgram.Caption = .OptionsAutosaveStartStandardProgram
50090   chkAutosaveSendEmail.Caption = .OptionsSendEmailAfterAutosave
50100  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "SetLanguageStrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetOptions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Options1
50020   chkUseAutosave.value = .UseAutosave
50030
50041   Select Case .AutosaveFormat
         Case 0:
50060     cmbAutosaveFormat.ListIndex = 0  ' PDF
50070    Case 1:
50080     cmbAutosaveFormat.ListIndex = 3  ' PNG
50090    Case 2:
50100     cmbAutosaveFormat.ListIndex = 4  ' JPEG
50110    Case 3:
50120     cmbAutosaveFormat.ListIndex = 5  ' BMP
50130    Case 4:
50140     cmbAutosaveFormat.ListIndex = 6  ' PCX
50150    Case 5:
50160     cmbAutosaveFormat.ListIndex = 7  ' TIFF
50170    Case 6:
50180     cmbAutosaveFormat.ListIndex = 8  ' PS
50190    Case 7:
50200     cmbAutosaveFormat.ListIndex = 9  ' EPS
50210    Case 8:
50220     cmbAutosaveFormat.ListIndex = 10 ' TXT
50230    Case 9:
50240     cmbAutosaveFormat.ListIndex = 1  ' PDF/A
50250    Case 10:
50260     cmbAutosaveFormat.ListIndex = 2 ' PDF/X
50270    Case 11:
50280     cmbAutosaveFormat.ListIndex = 11 ' PSD
50290    Case 12:
50300     cmbAutosaveFormat.ListIndex = 12 ' PCL
50310    Case 13:
50320     cmbAutosaveFormat.ListIndex = 13 ' RAW
50330    Case 14:
50340     cmbAutosaveFormat.ListIndex = 14 ' SVG
50350    Case Else:
50360     cmbAutosaveFormat.ListIndex = 0  ' PDF
50370   End Select
50380
50390   txtAutosaveFilename.Text = .AutosaveFilename
50400
50410   txtAutosaveDirectory.Text = .AutosaveDirectory
50420   chkAutosaveStartStandardProgram.value = .AutosaveStartStandardProgram
50430   chkUseAutosaveDirectory.value = .UseAutosaveDirectory
50440   chkAutosaveSendEmail.value = .SendEmailAfterAutoSaving
50450  End With
50460  If chkUseAutosave.value = 1 Then
50470    ViewAutosave True
50480   Else
50490    ViewAutosave False
50500  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "SetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub GetOptions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 '  Case 0: 'PDF   '  .AddItem "PDF"
50020 '  Case 1: 'PNG   '  .AddItem "PDF/A-1b"
50030 '  Case 2: 'JPEG  '  .AddItem "PDF/X"
50040 '  Case 3: 'BMP   '  .AddItem "PNG"
50050 '  Case 4: 'PCX   '  .AddItem "JPEG"
50060 '  Case 5: 'TIFF  '  .AddItem "BMP"
50070 '  Case 6: 'PS    '  .AddItem "PCX"
50080 '  Case 7: 'EPS   '  .AddItem "TIFF"
50090 '  Case 8: 'TXT   '  .AddItem "PS"
50100 '  Case 9: 'PDFA  '  .AddItem "EPS"
50110 '  Case 10: 'PDFX '  .AddItem "TXT"
50120 '  Case 11: 'PSD  '  .AddItem "PSD"
50130 '  Case 12: 'PCL  '  .AddItem "PCL"
50140 '  Case 13: 'RAW  '  .AddItem "RAW"
50150 '  Case 14: 'SVG  '  .AddItem "SVG"
50160
50170  With Options1
50180   .UseAutosave = Abs(chkUseAutosave.value)
50190   If LenB(CStr(cmbAutosaveFormat.ListIndex)) > 0 Then
50201    Select Case cmbAutosaveFormat.ListIndex
          Case 0:
50220      .AutosaveFormat = 0  ' PDF
50230     Case 1:
50240      .AutosaveFormat = 9  ' PDF/A
50250     Case 2:
50260      .AutosaveFormat = 10 ' PDF/X
50270     Case 3:
50280      .AutosaveFormat = 1  ' PNG
50290     Case 4:
50300      .AutosaveFormat = 2  ' JPEG
50310     Case 5:
50320      .AutosaveFormat = 3  ' BMP
50330     Case 6:
50340      .AutosaveFormat = 4  ' PCX
50350     Case 7:
50360      .AutosaveFormat = 5  ' TIFF
50370     Case 8:
50380      .AutosaveFormat = 6  ' PS
50390     Case 9:
50400      .AutosaveFormat = 7  ' EPS
50410     Case 10:
50420      .AutosaveFormat = 8  ' TXT
50430     Case 11:
50440      .AutosaveFormat = 11 ' PSD
50450     Case 12:
50460      .AutosaveFormat = 12 ' PCL
50470     Case 13:
50480      .AutosaveFormat = 13 ' RAW
50490     Case 14:
50500      .AutosaveFormat = 14 ' SVG
50510    End Select
50520   End If
50530   .AutosaveFilename = Trim$(txtAutosaveFilename.Text)
50540   If LenB(.AutosaveFilename) = 0 Then
50550    .AutosaveFilename = StandardOptions.AutosaveFilename
50560   End If
50570   .UseAutosaveDirectory = Abs(chkUseAutosaveDirectory.value)
50580   .AutosaveDirectory = Trim$(txtAutosaveDirectory.Text)
50590   If LenB(.AutosaveDirectory) = 0 Then
50600    .AutosaveDirectory = StandardOptions.AutosaveDirectory
50610   End If
50620   .AutosaveStartStandardProgram = Abs(chkAutosaveStartStandardProgram.value)
50630   .SendEmailAfterAutoSaving = Abs(chkAutosaveSendEmail.value)
50640  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "GetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewAutosave(Viewit As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblAutosaveformat.Enabled = Viewit
50020  cmbAutosaveFormat.Enabled = Viewit
50030  lblAutosaveFilename.Enabled = Viewit
50040  txtAutosaveFilename.Enabled = Viewit
50050  txtAutoSaveFilenamePreview.Enabled = Viewit
50060  lblAutosaveFilenameTokens.Enabled = Viewit
50070  cmbAutoSaveFilenameTokens.Enabled = Viewit
50080  chkUseAutosaveDirectory.Enabled = Viewit
50090  txtAutoSaveDirectoryPreview.Enabled = Viewit
50100  chkAutosaveStartStandardProgram.Enabled = Viewit
50110  chkAutosaveSendEmail.Enabled = Viewit
50120
50130  If Viewit Then
50140    cmbAutosaveFormat.BackColor = &H80000005
50150    cmbAutoSaveFilenameTokens.BackColor = &H80000005
50160    txtAutosaveFilename.BackColor = &H80000005
50170    txtAutosaveDirectory.BackColor = &H80000005
50180   Else
50190    cmbAutosaveFormat.BackColor = &H8000000F
50200    cmbAutoSaveFilenameTokens.BackColor = &H8000000F
50210    txtAutosaveFilename.BackColor = &H8000000F
50220    txtAutosaveDirectory.BackColor = &H8000000F
50230  End If
50240  If chkUseAutosaveDirectory.value = 1 And Viewit Then
50250    ViewAutosaveDirectory True
50260   Else
50270    ViewAutosaveDirectory False
50280  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "ViewAutosave")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewAutosaveDirectory(Viewit As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtAutosaveDirectory.Enabled = Viewit
50020  txtAutoSaveDirectoryPreview.Enabled = Viewit
50030  cmdGetAutosaveDirectory.Enabled = Viewit
50040  If Viewit = True Then
50050    txtAutosaveDirectory.BackColor = &H80000005
50060   Else
50070    txtAutosaveDirectory.BackColor = &H8000000F
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "ViewAutosaveDirectory")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUseAutosave_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkUseAutosave.value = 1 Then
50020    ViewAutosave True
50030   Else
50040    ViewAutosave False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "chkUseAutosave_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUseAutosaveDirectory_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkUseAutosaveDirectory.value = 1 Then
50020    ViewAutosaveDirectory True
50030   Else
50040    ViewAutosaveDirectory False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "chkUseAutosaveDirectory_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbAutosaveFormat_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtAutosaveFilename.ToolTipText = txtAutosaveFilename.Text
50020  txtAutoSaveFilenamePreview.Text = GetSubstFilename("B:\dummy.dum", txtAutosaveFilename.Text, , True) & _
  GetSaveAutosaveFormatExtension(cmbAutosaveFormat.ListIndex)
50040  If IsValidPath("C:\" & txtAutoSaveFilenamePreview.Text) = False Then
50050    txtAutoSaveFilenamePreview.ForeColor = vbRed
50060   Else
50070    txtAutoSaveFilenamePreview.ForeColor = &H80000008
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "cmbAutosaveFormat_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbAutoSaveFilenameTokens_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtAutosaveFilename.Text = txtAutosaveFilename.Text & cmbAutoSaveFilenameTokens.Text
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "cmbAutoSaveFilenameTokens_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetAutosaveDirectory_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strFolder As String
50020  strFolder = BrowseForFolderFiles(UserControl.Parent.hwnd, LanguageStrings.OptionsAutosaveDirectoryPrompt)
50030  If Len(strFolder) = 0 Then
50040   Exit Sub
50050  End If
50060  txtAutosaveDirectory.Text = CompletePath(strFolder)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "cmdGetAutosaveDirectory_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtAutosaveDirectory_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtAutosaveDirectory.ToolTipText = txtAutosaveDirectory.Text
50020  With txtAutoSaveDirectoryPreview
50030   .Text = GetSubstFilename2(txtAutosaveDirectory.Text)
50040   .ToolTipText = .Text
50050   If IsValidPath(.Text) = False Then
50060     .ForeColor = vbRed
50070    Else
50080     .ForeColor = &H80000008
50090   End If
50100  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "txtAutosaveDirectory_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtAutosaveFilename_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Ext As String
50020  txtAutosaveFilename.ToolTipText = txtAutosaveFilename.Text
50030  With txtAutoSaveFilenamePreview
50040   .Text = GetSubstFilename("B:\dummy.dum", txtAutosaveFilename.Text, , True) & GetSaveAutosaveFormatExtension(cmbAutosaveFormat.ListIndex)
50050   .ToolTipText = .Text
50060   If IsValidPath("C:\" & .Text) = False Then
50070     .ForeColor = vbRed
50080    Else
50090     .ForeColor = &H80000008
50100   End If
50110  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptAutosave", "txtAutosaveFilename_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
