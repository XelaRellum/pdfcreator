VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Transtool"
   ClientHeight    =   6420
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9915
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows-Standard
   WindowState     =   2  'Maximiert
   Begin VB.PictureBox picAbout 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4320
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3720
      Top             =   720
   End
   Begin VB.ComboBox cmbProgramFontsize 
      Appearance      =   0  '2D
      Height          =   315
      Left            =   8160
      TabIndex        =   6
      Text            =   "8"
      Top             =   720
      Width           =   975
   End
   Begin TransTool.XP_ProgressBar xpPgb 
      Height          =   255
      Left            =   9240
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   6956042
      Scrolling       =   8
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3240
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2760
      Top             =   720
   End
   Begin MSComctlLib.ImageList imlTlb 
      Left            =   1560
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":548A
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5824
            Key             =   "save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5BBE
            Key             =   "search"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F58
            Key             =   "empty"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":62F2
            Key             =   "unmark"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Oben ausrichten
      Height          =   630
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlLsv 
      Left            =   960
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":668C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6ADE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2280
      Top             =   720
   End
   Begin VB.ComboBox cmbFonts 
      Appearance      =   0  '2D
      Height          =   315
      Left            =   4920
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   2
      ToolTipText     =   "Font"
      Top             =   720
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6165
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsv 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   1365
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.ComboBox cmbCharset 
      Appearance      =   0  '2D
      Height          =   315
      Left            =   6510
      TabIndex        =   3
      Text            =   "cmbCharset"
      ToolTipText     =   "Charset"
      Top             =   720
      Width           =   1575
   End
   Begin VB.Image imgPaypal 
      Height          =   465
      Left            =   0
      Top             =   840
      Width           =   930
   End
   Begin VB.Menu mnFileMain 
      Caption         =   "&File"
      Begin VB.Menu mnFile 
         Caption         =   "&Open"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnFile 
         Caption         =   "Open template file"
         Index           =   1
      End
      Begin VB.Menu mnFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnFile 
         Caption         =   "&Save"
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu mnFile 
         Caption         =   "Save &as"
         Index           =   4
      End
      Begin VB.Menu mnFile 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnFile 
         Caption         =   "1"
         Index           =   6
      End
      Begin VB.Menu mnFile 
         Caption         =   "2"
         Index           =   7
      End
      Begin VB.Menu mnFile 
         Caption         =   "3"
         Index           =   8
      End
      Begin VB.Menu mnFile 
         Caption         =   "4"
         Index           =   9
      End
      Begin VB.Menu mnFile 
         Caption         =   "5"
         Index           =   10
      End
      Begin VB.Menu mnFile 
         Caption         =   "6"
         Index           =   11
      End
      Begin VB.Menu mnFile 
         Caption         =   "7"
         Index           =   12
      End
      Begin VB.Menu mnFile 
         Caption         =   "8"
         Index           =   13
      End
      Begin VB.Menu mnFile 
         Caption         =   "9"
         Index           =   14
      End
      Begin VB.Menu mnFile 
         Caption         =   "10"
         Index           =   15
      End
      Begin VB.Menu mnFile 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnFile 
         Caption         =   "E&xit"
         Index           =   17
      End
   End
   Begin VB.Menu mnEditMain 
      Caption         =   "&Edit"
      Begin VB.Menu mnEdit 
         Caption         =   "&Search"
         Index           =   0
         Shortcut        =   ^F
      End
      Begin VB.Menu mnEdit 
         Caption         =   "&Unmark all search results"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnEdit 
         Caption         =   "&Go to the next empty item"
         Index           =   3
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnHelpMain 
      Caption         =   "&?"
      Begin VB.Menu mnHelp 
         Caption         =   "&Paypal"
         Index           =   0
      End
      Begin VB.Menu mnHelp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnHelp 
         Caption         =   "&Homepage"
         Index           =   2
      End
      Begin VB.Menu mnHelp 
         Caption         =   "PDFCreator on &Sourceforge"
         Index           =   3
      End
      Begin VB.Menu mnHelp 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnHelp 
         Caption         =   "&About"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mnFileRecentFilesStart = 6

Private TSWidth As Long, mutex As clsMutex, SaveFilename As String, _
 OldString As String, ChangedListitem As Boolean

Public LastAboutTop As Long, LastAboutLeft As Long

Private Sub cmbCharset_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  On Local Error GoTo ErrorHandler
50020  lsv.Font.Charset = cmbCharset.Text
50030  RefreshListview
50040  Exit Sub
ErrorHandler:
50060  If Err.Number = 380 Then
50070   cmbCharset.Text = 0
50080  End If
50090  Err.Clear
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "cmbCharset_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbCharset_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With cmbCharset
50020   .Text = .ItemData(.ListIndex)
50030  End With
50040  lsv.Font.Charset = cmbCharset.Text
50050  RefreshListview
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "cmbCharset_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbCharset_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim allow As String, tStr As String
50020  allow = "0123456789" & Chr$(8) & Chr$(13)
50030  tStr = Chr$(KeyAscii)
50040  If InStr(1, allow, tStr) = 0 Then
50050   KeyAscii = 0
50060  End If
50070  RefreshListview
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "cmbCharset_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbFonts_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lsv.Font.Name = cmbFonts.List(cmbFonts.ListIndex)
50020  RefreshListview
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "cmbFonts_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbProgramFontsize_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ChangeProgramFont
50020  RefreshListview
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "cmbProgramFontsize_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Load()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Set mutex = New clsMutex
50050  InitForm
50060  ShowPaypalMenuimage
50070  RecentfileslistLocation = ApplicationDatapath
50080  ShowRecentFiles
50090  Timer2.Enabled = True
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmMain", "Form_Load")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Timer3
50020   .Enabled = False
50030   .Enabled = True
50040  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Form_Resize")
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
50010  Dim aw As Long
50020  If ChangedListitem = True Then
50030   aw = MsgBox("You have made some changes. Would you like to save this changes before the program exit?", vbYesNoCancel + vbQuestion)
50040   If aw = vbCancel Then
50050    Cancel = 1
50060    Exit Sub
50070   End If
50080   If aw = vbYes Then
50090    SaveLanguageFile
50100   End If
50110  End If
50120  RemovePanelControl cmbProgramFontsize
50130  RemovePanelControl cmbCharset
50140  RemovePanelControl cmbFonts
50150  RemovePanelControl xpPgb
50160  mutex.CloseMutex
50170  Set mutex = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Form_Unload")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  If ColumnHeader.Index = 2 Then
50030   Exit Sub
50040  End If
50050  With lsv
50060   If .SortKey = ColumnHeader.Index - 1 Then
50070     .SortOrder = 1 - .SortOrder
50080    Else
50090     .SortOrder = lvwAscending
50100     .ColumnHeaders(.SortKey + 1).Icon = 0
50110   End If
50120   .SortKey = ColumnHeader.Index - 1
50130   ColumnHeader.Icon = .SortOrder + 1
50140  End With
50150  If lsv.SelectedItem.Index > 0 Then
50160   lsv.SelectedItem.EnsureVisible
50170  End If
50180  Timer4.Enabled = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "lsv_ColumnClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsv_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If lsv.ListItems.Count > 0 Then
50020   OldString = lsv.SelectedItem.Text
50030   frmEdit.txt(0).Text = Replace(lsv.SelectedItem.SubItems(5), "%n", vbCrLf)
50040   frmEdit.txt(1).Text = Replace(lsv.SelectedItem.Text, "%n", vbCrLf)
50050   frmEdit.Show vbModal, Me
50060   If OldString <> lsv.SelectedItem.Text Then
50070    ChangedListitem = True
50080    Caption = "Transtool*"
50090   End If
50100   RefreshStb
50110  End If
50120  EnableEmptyButton
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "lsv_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsv_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyAscii = 13 Then
50020   lsv.StartLabelEdit
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "lsv_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnEdit_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim item As ListItem
50021  Select Case Index
        Case 0
50040    frmSearch.Show vbModeless, Me
50050   Case 1
50060    Unmark
50070   Case 3
50080    GotoEmptyValue
50090  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnEdit_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnExit_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Unload Me
50020  End
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnExit_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnFile_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Files As Collection, Filename As String, c As Long, aw As Long
50021  Select Case Index
        Case 0
50040    OpenLanguageFile
50050   Case 1
50060    OpenTemplateFile
50070    CheckMenuAndButton
50080   Case 3
50090    SaveLanguageFile
50100   Case 4
50110    SaveLanguageFile True
50120   Case mnFileRecentFilesStart To mnFileRecentFilesStart + MaxRecentfiles - 1
50130    OpenRecentFile Index - mnFileRecentFilesStart + 1
50140   Case 17
50150    Unload Me
50160    End
50170  End Select
50180  RefreshStb
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnFile_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnHelp_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 0:
50030    OpenDocument Paypal
50040   Case 2:
50050    OpenDocument Homepage
50060   Case 3:
50070    OpenDocument Sourceforge
50080   Case 5:
50090    frmAbout.Show , Me
50100  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnHelp_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Timer2_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Timer2.Enabled = False
50020  If Len(TemplateInifile) > 0 Then
50030   ReadTemplate TemplateInifile
50040  End If
50050  CheckMenuAndButton
50060  If Len(TranslatedInifile) > 0 Then
50070   ShowLanguageIniFile TranslatedInifile
50080   SaveFilename = TranslatedInifile
50090  End If
50100  RefreshStb
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Timer2_Timer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Timer3_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Timer3.Enabled = False
50020  ResizeFormWithoutFlicker
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Timer3_Timer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Timer4_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Timer4.Enabled = False
50020  If lsv.SelectedItem.Index > 0 Then
50030   lsv.SelectedItem.EnsureVisible
50040  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Timer4_Timer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Button.Key
        Case Is = "open"
50030    OpenLanguageFile
50040   Case Is = "save"
50050    SaveLanguageFile
50060   Case Is = "search"
50070    frmSearch.Show vbModeless, Me
50080   Case Is = "empty"
50090    GotoEmptyValue
50100   Case Is = "unmark"
50110    Unmark
50120  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "tlb_ButtonClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbProgramFontSize_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ChangeProgramFont
50020  RefreshListview
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "cmbProgramFontSize_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbProgramFontSize_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim allow As String, tStr As String
50020  allow = "0123456789" & Chr$(8) & Chr$(13)
50030  tStr = Chr$(KeyAscii)
50040  If InStr(1, allow, tStr) = 0 Then
50050   KeyAscii = 0
50060  End If
50070  RefreshListview
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "cmbProgramFontSize_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Timer1_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ' Create a mutex if possible
50020  If mutex.CheckMutex(TransTool_GUID) = False Then
50030   mutex.CreateMutex TransTool_GUID
50040  End If
50050  DoEvents
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Timer1_Timer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub EnableEmptyButton()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If GetCountEmptyValues > 0 Then
50020    mnEdit(3).Enabled = True
50030    tlb.Buttons("empty").Enabled = True
50040   Else
50050    mnEdit(3).Enabled = False
50060    tlb.Buttons("empty").Enabled = False
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "EnableEmptyButton")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CheckMenuAndButton()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, Enable As Boolean
50020  If lsv.ListItems.Count = 0 Then
50030    Enable = False
50040   Else
50050    Enable = True
50060  End If
50070
50080  With tlb.Buttons
50090   For i = 1 To .Count
50100    If .item(i).Caption <> "-" Then
50110     .item(i).Enabled = Enable
50120    End If
50130   Next i
50140  End With
50150  With mnFile
50160   For i = .LBound To .UBound - 1
50170    If .item(i).Caption <> "-" Then
50180     .item(i).Enabled = Enable
50190    End If
50200   Next i
50210  End With
50220  With mnEdit
50230   For i = .LBound To .UBound
50240    If .item(i).Caption <> "-" Then
50250     .item(i).Enabled = Enable
50260    End If
50270   Next i
50280  End With
50290  If hLItems.Count = 0 Then
50300   mnEdit(1).Enabled = False
50310   tlb.Buttons("unmark").Enabled = False
50320  End If
50330  lsv.Enabled = Enable
50340  mnFile(1).Enabled = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "CheckMenuAndButton")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetCountEmptyValues() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, c As Long
50020  c = 0
50030  For i = 1 To lsv.ListItems.Count
50040   If lsv.ListItems(i).Text = "" Then
50050    c = c + 1
50060   End If
50070  Next i
50080  GetCountEmptyValues = c
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "GetCountEmptyValues")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub GotoEmptyValue()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim item As ListItem
50020  Set item = lsv.FindItem("", lvwText)
50030  item.EnsureVisible
50040  lsv.ListItems(item.Index).Selected = True
50050 ' lsv.StartLabelEdit
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "GotoEmptyValue")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub InitForm()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim cSystem As clsSystem, SMF As Collection, fi As Long, i As Long, tStr As String, _
  tlbB As MSComctlLib.Button, ctl As Control
50060  Set cSystem = New clsSystem: Set SMF = cSystem.GetSystemFont(Me, Menu)
50070  With tlb
50080   Set .ImageList = imlTlb
50090   Set tlbB = tlb.Buttons.Add(, "open", , tbrDefault, "open")
50100   tlbB.ToolTipText = "File open"
50110   tlbB.Description = .ToolTipText
50120   Set tlbB = tlb.Buttons.Add(, "save", , tbrDefault, "save")
50130   tlbB.ToolTipText = "File save"
50140   tlbB.Description = .ToolTipText
50150   tlb.Buttons.Add , , , tbrSeparator
50160   Set tlbB = tlb.Buttons.Add(, "search", , tbrDefault, "search")
50170   tlbB.ToolTipText = "Search"
50180   tlbB.Description = .ToolTipText
50190   Set tlbB = tlb.Buttons.Add(, "empty", , tbrDefault, "empty")
50200   tlbB.ToolTipText = mnEdit(3).Caption
50210   tlbB.Description = .ToolTipText
50220   tlb.Buttons.Add , , , tbrSeparator
50230   Set tlbB = tlb.Buttons.Add(, "unmark", , tbrDefault, "unmark")
50240   tlbB.Enabled = False
50250   tlbB.ToolTipText = mnEdit(1).Caption
50260   tlbB.Description = .ToolTipText
50270  End With
50280  TSWidth = 1200
50290  With lsv.ColumnHeaders
50300   .Clear
50310   .Add , "TranslatedString", "Translated text", 0
50320   .Add , "EmptyCol", "", 0 ' because of the listview-listsubitem-bold bug
50330   .Add , "Line", "", 500, lvwColumnRight
50340   .Add , "TemplateSection", "Section", TSWidth
50350   .Add , "TemplateKey", "Key", 0
50360   .Add , "TemplateString", "Template text", 0
50370  End With
50380  With lsv.ColumnHeaders
50390   .item("TranslatedString").Width = (lsv.Width - lsv.ColumnHeaders("Line").Width - TSWidth) / 3
50400   .item("TemplateKey").Width = .item("TranslatedString").Width
50410   .item("TemplateString").Width = .item("TranslatedString").Width
50420  End With
50430  lsv.ColumnHeaders("TranslatedString").Position = 6
50440  Set lsv.ColumnHeaderIcons = imlLsv
50450
50460  lsv.SortKey = 3
50470  lsv.ColumnHeaders(lsv.SortKey + 1).Icon = 1
50480
50490  With stb.Panels
50500   .Clear
50510   .Add , "Keys", "Keys:"
50520   .Add , "EmptyKeys", "Empty Keys: "
50530   .Add , "Fonts", ""
50540   .Add , "Charset", ""
50550   .Add , "Fontsize", ""
50560   .Add , "Status", ""
50570  End With
50580  With stb
50590   .Panels("Keys").Width = 1500
50600   .Panels("Keys").ToolTipText = "Count of keys"
50610   .Panels("EmptyKeys").Width = 1500
50620   .Panels("EmptyKeys").ToolTipText = "Count of empty keys"
50630   .Panels("Fonts").Width = 3000
50640   .Panels("Fonts").ToolTipText = "Font"
50650   .Panels("Charset").Width = 1500
50660   .Panels("Charset").ToolTipText = "Charset"
50670   .Panels("Fontsize").Width = 1000
50680   .Panels("Fontsize").ToolTipText = "Fontsize"
50690   .Panels("Status").ToolTipText = "Progress"
50700  End With
50710  With xpPgb
50720   .Scrolling = ccScrollingStandard
50730   .ShowText = True
50740   .Color = RGB(&H80, &HFF, &H80)
50750   .Font.Bold = True
50760  End With
50770  With cmbFonts
50780   .Clear
50790   For i = 1 To Screen.FontCount
50800    tStr = Trim$(Screen.Fonts(i))
50810    If Len(tStr) > 0 Then
50820     cmbFonts.AddItem tStr
50830    End If
50840   Next i
50850   If .ListCount > 0 Then
50860     For i = 0 To cmbFonts.ListCount - 1
50870      If SMF.Count > 0 Then
50880       If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
50890        fi = i
50900       End If
50910      End If
50920     Next i
50930    Else
50940    .ListIndex = 0
50950   End If
50960  End With
50970  With cmbCharset
50980   .Clear
50990   .AddItem "0, Western": .ItemData(.NewIndex) = 0
51000   .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
51010   .AddItem "77, Mac": .ItemData(.NewIndex) = 77
51020   .AddItem "161, Greek": .ItemData(.NewIndex) = 161
51030   .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
51040   .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
51050   .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
51060   .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
51070   .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
51080   .AddItem "238, Central European": .ItemData(.NewIndex) = 238
51090   .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
51100   .Text = 0
51110  End With
51120  If fi >= 0 Then
51130   cmbFonts.ListIndex = fi
51140   cmbCharset.Text = SMF(1)(2)
51150   lsv.Font.Name = cmbFonts.List(cmbFonts.ListIndex)
51160   lsv.Font.Charset = cmbCharset.Text
51170   cmbProgramFontsize.Text = SMF(1)(1)
51180  End If
51190  With cmbProgramFontsize
51200   .AddItem "8"
51210   .AddItem "9"
51220   .AddItem "10"
51230   .AddItem "11"
51240   .AddItem "12"
51250   .AddItem "14"
51260   .AddItem "16"
51270   .AddItem "18"
51280   .AddItem "20"
51290   .AddItem "22"
51300   .AddItem "24"
51310   .AddItem "26"
51320   .AddItem "28"
51330   .AddItem "36"
51340   .AddItem "48"
51350   .AddItem "72"
51360  End With
51370  For Each ctl In Controls
51380   If TypeOf ctl Is ComboBox Then
51390    ComboSetListWidth ctl
51400   End If
51410  Next ctl
51420  cmbCharset.ListIndex = 0
51430
51440  SetOptimalComboboxHeigth cmbCharset, Me
51450  SetOptimalComboboxHeigth cmbProgramFontsize, Me
51460
51470  SetPanelControl cmbFonts, stb, "Fonts", True
51480  SetPanelControl cmbCharset, stb, "Charset", True
51490  SetPanelControl cmbProgramFontsize, stb, "Fontsize", True
51500  SetPanelControl xpPgb, stb, "Status", True
51510
51520  Set imgPaypal.Picture = LoadResPicture(1002, vbResBitmap)
51530  ChangedListitem = False
51540  SaveFilename = ""
51550  Timer1.Enabled = True
51560 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
51570 Exit Sub
ErrPtnr_OnError:
51591 Select Case ErrPtnr.OnError("frmMain", "InitForm")
      Case 0: Resume
51610 Case 1: Resume Next
51620 Case 2: Exit Sub
51630 Case 3: End
51640 End Select
51650 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub OpenLanguageFile()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Files As Collection
50020  If OpenFileDialog(Files, , _
  "Languages-INI-Files (*.ini)|*.ini|All Files (*.*)|*.*", "*.ini", App.Path, _
  "Open translated file", OFN_PATHMUSTEXIST, Me.hwnd) > 0 Then
50050   If Files.Count > 0 Then
50060    If Dir(Files.item(1)) <> "" And Len(Trim$(Files.item(1))) > 0 Then
50070     ShowLanguageIniFile Files.item(1)
50080     SaveFilename = Files.item(1)
50090     AddRecentfile SaveFilename
50100     ShowRecentFiles
50110    End If
50120   End If
50130  End If
50140  RefreshStb
50150  EnableEmptyButton
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "OpenLanguageFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub OpenTemplateFile()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Files As Collection
50020  If OpenFileDialog(Files, , _
  "Languages-INI-Files (*.ini)|*.ini|All Files (*.*)|*.*", "*.ini", App.Path, _
  "Open template file", OFN_PATHMUSTEXIST, Me.hwnd) > 0 Then
50050   If Files.Count > 0 Then
50060    If Dir(Files.item(1)) <> "" And Len(Trim$(Files.item(1))) > 0 Then
50070     If CheckTemplate(Files.item(1)) = True Then
50080      ReadTemplate Files.item(1)
50090     End If
50100    End If
50110   End If
50120  End If
50130  RefreshStb
50140  EnableEmptyButton
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "OpenTemplateFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SaveLanguageFile(Optional SaveAs As Boolean = False)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Files As Collection, Filename As String, c As Long, aw As Long, ini As clsINI, _
  i As Long
50030  c = GetCountEmptyValues
50040  If c > 0 Then
50050   If c = 1 Then
50060     aw = MsgBox("There are 1 empty value. Continue?", vbYesNo)
50070    Else
50080     aw = MsgBox("There are " & c & " empty values. Continue?", vbYesNo)
50090   End If
50100   If aw = vbNo Then
50110    Exit Sub
50120   End If
50130  End If
50140  If Len(SaveFilename) = 0 Or SaveAs = True Then
50150   If SaveFileDialog(Filename, , "Languages-INI-Files (*.ini)|*.ini", "*.ini", App.Path, _
   App.EXEName, OFN_OVERWRITEPROMPT + OFN_PATHMUSTEXIST, Me.hwnd) > 0 Then
50170    If Len(Trim$(Filename)) > 0 Then
50180     SaveFilename = Filename
50190    End If
50200   End If
50210  End If
50220  If Len(SaveFilename) > 0 Then
50230   mnFileMain.Enabled = False
50240   mnEditMain.Enabled = False
50250   mnHelpMain.Enabled = False
50260   lsv.Enabled = False
50270   tlb.Enabled = False
50280   Screen.MousePointer = vbHourglass
50290   Set ini = New clsINI
50300   ini.Filename = SaveFilename
50310   ini.CreateIniFile
50320   With xpPgb
50330    .Visible = True
50340    .Min = 0: .Max = lsv.ListItems.Count
50350    For i = 1 To lsv.ListItems.Count
50360     With lsv.ListItems(i)
50370      ini.SaveKey .Text, .ListSubItems(4).Text, .ListSubItems(3).Text
50380     End With
50390     .Value = i
50400    Next i
50410    .Value = 0
50420    .Visible = False
50430   End With
50440   ChangedListitem = False
50450   Caption = "Transtool"
50460   Screen.MousePointer = vbNormal
50470   mnFileMain.Enabled = True
50480   mnEditMain.Enabled = True
50490   mnHelpMain.Enabled = True
50500   lsv.Enabled = True
50510   tlb.Enabled = True
50520  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SaveLanguageFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ShowLanguageIniFile(LanguageIniFilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI, secs As Collection, keys As Collection, i As Long, j As Long, _
  l As Long, c As Long
50030  Screen.MousePointer = vbHourglass
50040  Set ini = New clsINI
50050  ini.Filename = LanguageIniFilename
50060  Set secs = ini.GetAllSectionsFromInifile(, True)
50070  If secs.Count > 0 Then
50080   c = 0
50090   For i = 1 To secs.Count
50100    Set keys = ini.GetAllKeysFromSection(secs.item(i), , , True)
50110    c = c + keys.Count
50120   Next i
50130   With xpPgb
50140    .Min = 0: .Max = c
50150    .Visible = True
50160    c = 0
50170    For i = 1 To secs.Count
50180     Set keys = ini.GetAllKeysFromSection(secs.item(i), , , True)
50190     For j = 1 To keys.Count
50200      c = c + 1
50210      .Value = c
50220      For l = 1 To lsv.ListItems.Count
50230       If UCase$(lsv.ListItems(l).ListSubItems(3)) = UCase$(secs.item(i)) And _
       UCase$(lsv.ListItems(l).ListSubItems(4)) = UCase$(keys.item(j)(0)) Then
50250         lsv.ListItems(l).Text = keys.item(j)(1)
50260         Exit For
50270       End If
50280      Next l
50290     Next j
50300    Next i
50310    .Value = 0
50320    .Visible = False
50330   End With
50340  End If
50350  Screen.MousePointer = vbNormal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "ShowLanguageIniFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ReadTemplate(IniFile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI, secs As Collection, keys As Collection, _
  i As Long, j As Long, lsvItem As ListItem
50030
50040  lsv.ListItems.Clear
50050  Set ini = New clsINI
50060  ini.Filename = IniFile
50070  Set secs = ini.GetAllSectionsFromInifile(, True)
50080  For i = 1 To secs.Count
50090   Set keys = ini.GetAllKeysFromSection(secs(i), , , True)
50100   For j = 1 To keys.Count
50110    Set lsvItem = lsv.ListItems.Add(, , "")
50120    lsvItem.ListSubItems.Add , , ""
50130    lsvItem.ListSubItems.Add , , ""
50140    lsvItem.ListSubItems.Add , , secs(i)
50150    lsvItem.ListSubItems.Add , , keys(j)(0)
50160    lsvItem.ListSubItems.Add , , keys(j)(1)
50170   Next j
50180  Next i
50190  LsvLineNumber
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "ReadTemplate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub LsvLineNumber(Optional Direction = 0)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, MaxZeroStr As String
50020  MaxZeroStr = String(Len(CStr(lsv.ListItems.Count)), "0")
50030  If Direction = 0 Then
50040    For i = 1 To lsv.ListItems.Count
50050     lsv.ListItems(i).ListSubItems(2) = Format(i, MaxZeroStr)
50060    Next i
50070   Else
50080    For i = lsv.ListItems.Count To 1 Step -1
50090     lsv.ListItems(lsv.ListItems.Count - i + 1).ListSubItems(2) = Format(i, MaxZeroStr)
50100    Next i
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "LsvLineNumber")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub RefreshStb()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With stb
50020   .Panels("Keys").Text = "Keys:" & lsv.ListItems.Count
50030   .Panels("EmptyKeys").Text = "Empty Keys:" & GetCountEmptyValues
50040   .Refresh
50050  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "RefreshStb")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ShowPaypalMenuimage()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim h1 As Long, h2 As Long, com As Long
50020  h1 = GetMenu(Me.hwnd): h2 = GetSubMenu(h1, 2)
50030  com = GetMenuItemID(h2, 0)
50040  ModifyMenu h2, com, MF_BYCOMMAND Or MF_BITMAP, com, CLng(imgPaypal.Picture)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "ShowPaypalMenuimage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ResizeFormWithoutFlicker()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tL As Long
50020  If frmMain.WindowState = vbMinimized Then
50030   Exit Sub
50040  End If
50050  If Me.Width < 6000 Then
50060   Me.Width = 6000
50070  End If
50080  If Me.Height < 6000 Then
50090   Me.Height = 6000
50100  End If
50110  With lsv
50120   .Top = tlb.Height
50130   .Left = 0
50140   .Height = Me.ScaleHeight - stb.Height - tlb.Height
50150   .Width = Me.ScaleWidth
50160  End With
50170  With lsv.ColumnHeaders
50180   .item("TranslatedString").Width = (lsv.Width - lsv.ColumnHeaders("Line").Width - TSWidth) / 3 - 120
50190   .item("TemplateKey").Width = .item("TranslatedString").Width
50200   .item("TemplateString").Width = .item("TranslatedString").Width
50210  End With
50220  With stb
50230   tL = Me.ScaleWidth - .Panels("Keys").Width - .Panels("EmptyKeys").Width - _
   .Panels("Fonts").Width - .Panels("Charset").Width - .Panels("Fontsize").Width
50250   If tL < 100 Then
50260    .Panels("Status").Width = 100
50270    Else
50280    .Panels("Status").Width = tL - 99
50290   End If
50300  End With
50310  AdjustControlToPanel xpPgb, stb, "Status", True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "ResizeFormWithoutFlicker")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ShowRecentFiles()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim col As Collection, i As Long, Path As String, File As String
50020  Set col = GetRecentFiles
50030  mnFile(mnFile.Count - 2).Visible = False
50040  For i = 1 To MaxRecentfiles
50050   mnFile(mnFileRecentFilesStart + i - 1).Visible = False
50060  Next i
50070  If RecentFilesCount > 0 Then
50080   For i = 1 To RecentFilesCount
50090    If i <= col.Count Then
50100     SplitPath col(i), , Path, , File
50110     mnFile(mnFileRecentFilesStart + i - 1).Caption = "&" & i & " " & _
     ShortenPath(Me.hDC, CompletePath(Path) & File, 200)
50130     mnFile(mnFileRecentFilesStart + i - 1).Tag = col(i)
50140     mnFile(mnFileRecentFilesStart + i - 1).Visible = True
50150     If mnFile(mnFile.Count - 2).Visible = False Then
50160      mnFile(mnFile.Count - 2).Visible = True
50170     End If
50180    End If
50190   Next i
50200  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "ShowRecentFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub OpenRecentFile(RecentFilenumber As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Filename As String
50020  Filename = GetRecentFile(RecentFilenumber)
50030  If FileExists(Filename) = True Then
50040    ShowLanguageIniFile Filename
50050    SaveFilename = Filename
50060    RefreshStb
50070    AddRecentfile Filename
50080   Else
50090    MsgBox "The file doesn't exists!", vbExclamation
50100    RemoveRecentFile RecentFilenumber
50110  End If
50120  ShowRecentFiles
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "OpenRecentFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ChangeProgramFont()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tL As Long
50020  If Trim$(cmbProgramFontsize.Text) = "" Then
50030   cmbProgramFontsize.Text = 8
50040  End If
50050  tL = CLng(cmbProgramFontsize.Text)
50060  If tL <= 0 Then
50070   tL = 1
50080  End If
50090  If tL > 72 Then
50100   tL = 72
50110  End If
50120  cmbProgramFontsize.Text = tL
50130  lsv.Font.Size = tL
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "ChangeProgramFont")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub RefreshListview()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  For i = 1 To lsv.ColumnHeaders.Count
50030   lsv.ColumnHeaders(i).Width = lsv.ColumnHeaders(i).Width
50040  Next i
50050  lsv.Refresh
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "RefreshListview")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Unmark()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  xpPgb.Visible = True
50020  HighlightListitems frmMain.lsv, hLItems, stb, xpPgb, , lsvColor, False, "Unmarking ..."
50030  xpPgb.Visible = False
50040  mnEdit(1).Enabled = False
50050  tlb.Buttons("unmark").Enabled = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Unmark")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

