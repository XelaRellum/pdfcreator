VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1680
      Top             =   5880
   End
   Begin VB.TextBox txtProgramFontsize 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   5520
      TabIndex        =   4
      Text            =   "8"
      ToolTipText     =   "Fontsize"
      Top             =   5160
      Width           =   975
   End
   Begin VB.ComboBox cmbCharset 
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Text            =   "cmbCharset"
      ToolTipText     =   "Charset"
      Top             =   5160
      Width           =   1935
   End
   Begin VB.ComboBox cmbFonts 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   2
      ToolTipText     =   "Font"
      Top             =   5160
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   120
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   10821
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   255
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu mnFile 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin VB.Menu mnFile 
         Caption         =   "&Save"
         Index           =   1
      End
      Begin VB.Menu mnFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnExit 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TSWidth As Long, mutex As clsMutex

Private Sub cmbCharset_Change()
 On Error GoTo ErrorHandler
 lsv.Font.Charset = cmbCharset.Text
 Exit Sub
ErrorHandler:
 If Err.Number = 380 Then
  cmbCharset.Text = 0
 End If
 Err.Clear
End Sub

Private Sub cmbCharset_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With cmbCharset
50020   .Text = .ItemData(.ListIndex)
50030  End With
50040  lsv.Font.Charset = cmbCharset.Text
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
50050    KeyAscii = 0
50060  End If
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

Private Sub Form_Load()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set mutex = New clsMutex
50020  InitForm
50030  ReadTemplate
50040  RefreshStb
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If frmMain.WindowState = vbMinimized Then
50020   Exit Sub
50030  End If
50040  lsv.Height = Me.ScaleHeight - stb.Height
50050  lsv.Width = Me.ScaleWidth
50060  If Me.Width < 6000 Then
50070   Me.Width = 6000
50080  End If
50090  lsv.ColumnHeaders("TemplateKey").Width = (lsv.Width - TSWidth) / 3 - 120
50100  lsv.ColumnHeaders("EnglishString").Width = (lsv.Width - TSWidth) / 3 - 120
50110  lsv.ColumnHeaders("TranslatedString").Width = (lsv.Width - TSWidth) / 3 - 120
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

Private Sub InitForm()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim cSystem As clsSystem, SMF As Collection, fi As Long, i As Long, tStr As String
50020  Set cSystem = New clsSystem: Set SMF = cSystem.GetSystemFont(Me, Menu)
50030  TSWidth = 1000
50040  With lsv.ColumnHeaders
50050   .Clear
50060   .Add , "TranslatedString", "Translated Text", (lsv.Width - TSWidth) / 3
50070   .Add , "TemplateSection", "Section", TSWidth
50080   .Add , "TemplateKey", "Key", (lsv.Width - TSWidth) / 3
50090   .Add , "EnglishString", "English Text", (lsv.Width - TSWidth) / 3
50100  End With
50110  lsv.ColumnHeaders("TranslatedString").Position = 4
50120  With stb.Panels
50130   .Clear
50140   .Add , "Keys", "Keys:"
50150   .Add , "EmptyKeys", "Empty Keys: "
50160   .Add , "Fonts", ""
50170   .Add , "Charset", ""
50180   .Add , "Fontsize", ""
50190  End With
50200  With stb
50210   .Panels("Keys").Width = 2000
50220   .Panels("EmptyKeys").Width = 2000
50230   .Panels("Fonts").Width = 4000
50240   .Panels("Charset").Width = 2000
50250   .Panels("Fontsize").Width = 1000
50260  End With
50270  SetPanelControl cmbFonts, stb, "Fonts", True
50280  SetPanelControl cmbCharset, stb, "Charset", True
50290  SetPanelControl txtProgramFontsize, stb, "Fontsize", True
50300  With cmbFonts
50310   .Clear
50320   For i = 1 To Screen.FontCount
50330    tStr = Trim$(Screen.Fonts(i))
50340    If Len(tStr) > 0 Then
50350     cmbFonts.AddItem tStr
50360    End If
50370   Next i
50380   If .ListCount > 0 Then
50390     For i = 0 To cmbFonts.ListCount - 1
50400      If SMF.Count > 0 Then
50410       If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
50420        fi = i
50430       End If
50440      End If
50450     Next i
50460    Else
50470    .ListIndex = 0
50480   End If
50490  End With
50500  With cmbCharset
50510   .Clear
50520   .AddItem "0, Western": .ItemData(.NewIndex) = 0
50530   .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
50540   .AddItem "77, Mac": .ItemData(.NewIndex) = 77
50550   .AddItem "161, Greek": .ItemData(.NewIndex) = 161
50560   .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
50570   .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
50580   .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
50590   .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
50600   .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
50610   .AddItem "238, Central European": .ItemData(.NewIndex) = 238
50620   .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
50630   .Text = 0
50640  End With
50650  If fi >= 0 Then
50660   cmbFonts.ListIndex = fi
50670   cmbCharset.Text = SMF(1)(2)
50680   lsv.Font.Name = cmbFonts.List(cmbFonts.ListIndex)
50690   lsv.Font.Charset = cmbCharset.Text
50700   txtProgramFontsize.Text = SMF(1)(1)
50710  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "InitForm")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ReadTemplate()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI, templateInifilename As String, secs As Collection, _
  Keys As Collection, i As Long, j As Long, lsvItem As ListItem
50030  templateInifilename = App.Path & "\english.ini"
50040  Set ini = New clsINI
50050  ini.FileName = App.Path & "\english.ini"
50060  If ini.CheckIniFile = False Then
50070   MsgBox "File 'english.ini' not found! Program exit!"
50080   End
50090  End If
50100  Set secs = ini.GetAllSectionsFromInifile(, True)
50110  For i = 1 To secs.Count
50120   Set Keys = ini.GetAllKeysFromSection(secs.item(i), , , True)
50130   For j = 1 To Keys.Count
50140    Set lsvItem = lsv.ListItems.Add(, , "")
50150    lsvItem.SubItems(1) = secs.item(i)
50160    lsvItem.SubItems(2) = Keys.item(j)(0)
50170    lsvItem.SubItems(3) = Keys.item(j)(1)
50180   Next j
50190  Next i
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

Private Sub Form_Unload(Cancel As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  RemovePanelControl txtProgramFontsize
50020  RemovePanelControl cmbCharset
50030  RemovePanelControl cmbFonts
50040  mutex.CloseMutex
50050  Set mutex = Nothing
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

Private Sub lsv_AfterLabelEdit(Cancel As Integer, NewString As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  NewString = Trim$(NewString)
50020  If Len(NewString) > 0 And Len(lsv.SelectedItem.Text) > 0 Then
50030   RefreshStb
50040  End If
50050  If Len(NewString) > 0 And Len(lsv.SelectedItem.Text) = 0 Then
50060   RefreshStb -1
50070  End If
50080  If Len(NewString) = 0 And Len(lsv.SelectedItem.Text) > 0 Then
50090   RefreshStb 1
50100  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "lsv_AfterLabelEdit")
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
50010  lsv.StartLabelEdit
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
 On Error GoTo ErrorHandler
 Select Case Index
  Case 0:
   With cdlg
    .CancelError = True
    .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    .FileName = ""
    .InitDir = App.Path
    .DialogTitle = App.EXEName
    .Filter = "Languages-INI-Files (*.ini)|*.ini" & "|" & _
     "All Files (*.*)|*.*"
    .ShowOpen
   End With
   If Dir(cdlg.FileName) <> "" And Len(Trim$(cdlg.FileName)) > 0 Then
    ShowLanguageIniFile cdlg.FileName
   End If
  Case 1:
   With cdlg
    .CancelError = True
    .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    .FileName = ""
    .InitDir = App.Path
    .DialogTitle = App.EXEName
    .Filter = "Languages-INI-Files (*.ini)|*.ini"
    .ShowSave
   End With
   If Len(Trim$(cdlg.FileName)) > 0 Then
    SaveLanguageIniFile cdlg.FileName
   End If
  Case 3:
   Unload Me
   End
 End Select
 RefreshStb
 Exit Sub

ErrorHandler:
 If Err.Number = 32755 Then Exit Sub
End Sub

Private Sub ShowLanguageIniFile(LanguageIniFilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI, secs As Collection, Keys As Collection, i As Long, j As Long, L As Long
50020  Me.MousePointer = vbHourglass
50030  Set ini = New clsINI
50040  ini.FileName = LanguageIniFilename
50050  Set secs = ini.GetAllSectionsFromInifile(, True)
50060  For i = 1 To secs.Count
50070   Set Keys = ini.GetAllKeysFromSection(secs.item(i), , , True)
50080   For j = 1 To Keys.Count
50090    For L = 1 To lsv.ListItems.Count
50100     If UCase$(lsv.ListItems(L).SubItems(1)) = UCase$(secs.item(i)) And _
     UCase$(lsv.ListItems(L).SubItems(2)) = UCase$(Keys.item(j)(0)) Then
50120       lsv.ListItems(L).Text = Keys.item(j)(1)
50130       Exit For
50140     End If
50150    Next L
50160   Next j
50170  Next i
50180  Me.MousePointer = vbNormal
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

Private Sub SaveLanguageIniFile(LanguageIniFilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI, i As Long
50020  Set ini = New clsINI
50030  ini.FileName = LanguageIniFilename
50040  ini.CreateIniFile
50050  For i = 1 To lsv.ListItems.Count
50060   With lsv.ListItems(i)
50070    ini.SaveKey .Text, .SubItems(2), .SubItems(1)
50080   End With
50090  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SaveLanguageIniFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetCountEmptyValues()
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

Private Sub RefreshStb(Optional AddNumber As Long = 0)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With stb
50020   .Panels("Keys").Text = "Keys:" & lsv.ListItems.Count
50030   .Panels("EmptyKeys").Text = "Empty Keys:" & GetCountEmptyValues + AddNumber
50040  End With
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

Private Sub txtProgramFontSize_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tL As Long
50020  If Trim$(txtProgramFontsize.Text) = "" Then
50030   txtProgramFontsize.Text = 8
50040  End If
50050  tL = CLng(txtProgramFontsize.Text)
50060  If tL <= 0 Then
50070   tL = 1
50080  End If
50090  If tL > 72 Then
50100   tL = 72
50110  End If
50120  txtProgramFontsize.Text = tL
50130  lsv.Font.Size = tL
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "txtProgramFontSize_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtProgramFontSize_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim allow As String, tStr As String
50020
50030  allow = "0123456789" & Chr$(8) & Chr$(13)
50040
50050  tStr = Chr$(KeyAscii)
50060
50070  If InStr(1, allow, tStr) = 0 Then
50080    KeyAscii = 0
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "txtProgramFontSize_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

