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
 On Local Error GoTo ErrorHandler
 lsv.Font.Charset = cmbCharset.Text
 Exit Sub
ErrorHandler:
 If Err.Number = 380 Then
  cmbCharset.Text = 0
 End If
 Err.Clear
End Sub

Private Sub cmbCharset_Click()
 With cmbCharset
  .Text = .ItemData(.ListIndex)
 End With
 lsv.Font.Charset = cmbCharset.Text
End Sub

Private Sub cmbCharset_KeyPress(KeyAscii As Integer)
 Dim allow As String, tStr As String
 allow = "0123456789" & Chr$(8) & Chr$(13)
 tStr = Chr$(KeyAscii)
 If InStr(1, allow, tStr) = 0 Then
   KeyAscii = 0
 End If
End Sub

Private Sub cmbFonts_Click()
 lsv.Font.Name = cmbFonts.List(cmbFonts.ListIndex)
End Sub

Private Sub Form_Load()
 Set mutex = New clsMutex
 InitForm
 ReadTemplate
 RefreshStb
End Sub

Private Sub Form_Resize()
 If frmMain.WindowState = vbMinimized Then
  Exit Sub
 End If
 lsv.Height = Me.ScaleHeight - stb.Height
 lsv.Width = Me.ScaleWidth
 If Me.Width < 6000 Then
  Me.Width = 6000
 End If
 lsv.ColumnHeaders("TemplateKey").Width = (lsv.Width - TSWidth) / 3 - 120
 lsv.ColumnHeaders("EnglishString").Width = (lsv.Width - TSWidth) / 3 - 120
 lsv.ColumnHeaders("TranslatedString").Width = (lsv.Width - TSWidth) / 3 - 120
End Sub

Private Sub InitForm()
 Dim cSystem As clsSystem, SMF As Collection, fi As Long, i As Long, tStr As String
 Set cSystem = New clsSystem: Set SMF = cSystem.GetSystemFont(Me, Menu)
 TSWidth = 1000
 With lsv.ColumnHeaders
  .Clear
  .Add , "TranslatedString", "Translated Text", (lsv.Width - TSWidth) / 3
  .Add , "TemplateSection", "Section", TSWidth
  .Add , "TemplateKey", "Key", (lsv.Width - TSWidth) / 3
  .Add , "EnglishString", "English Text", (lsv.Width - TSWidth) / 3
 End With
 lsv.ColumnHeaders("TranslatedString").Position = 4
 With stb.Panels
  .Clear
  .Add , "Keys", "Keys:"
  .Add , "EmptyKeys", "Empty Keys: "
  .Add , "Fonts", ""
  .Add , "Charset", ""
  .Add , "Fontsize", ""
 End With
 With stb
  .Panels("Keys").Width = 2000
  .Panels("EmptyKeys").Width = 2000
  .Panels("Fonts").Width = 4000
  .Panels("Charset").Width = 2000
  .Panels("Fontsize").Width = 1000
 End With
 SetPanelControl cmbFonts, stb, "Fonts", True
 SetPanelControl cmbCharset, stb, "Charset", True
 SetPanelControl txtProgramFontsize, stb, "Fontsize", True
 With cmbFonts
  .Clear
  For i = 1 To Screen.FontCount
   tStr = Trim$(Screen.Fonts(i))
   If Len(tStr) > 0 Then
    cmbFonts.AddItem tStr
   End If
  Next i
  If .ListCount > 0 Then
    For i = 0 To cmbFonts.ListCount - 1
     If SMF.Count > 0 Then
      If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
       fi = i
      End If
     End If
    Next i
   Else
   .ListIndex = 0
  End If
 End With
 With cmbCharset
  .Clear
  .AddItem "0, Western": .ItemData(.NewIndex) = 0
  .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
  .AddItem "77, Mac": .ItemData(.NewIndex) = 77
  .AddItem "161, Greek": .ItemData(.NewIndex) = 161
  .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
  .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
  .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
  .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
  .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
  .AddItem "238, Central European": .ItemData(.NewIndex) = 238
  .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
  .Text = 0
 End With
 If fi >= 0 Then
  cmbFonts.ListIndex = fi
  cmbCharset.Text = SMF(1)(2)
  lsv.Font.Name = cmbFonts.List(cmbFonts.ListIndex)
  lsv.Font.Charset = cmbCharset.Text
  txtProgramFontsize.Text = SMF(1)(1)
 End If
End Sub

Private Sub ReadTemplate()
 Dim ini As clsINI, templateInifilename As String, secs As Collection, _
  Keys As Collection, i As Long, j As Long, lsvItem As ListItem
 templateInifilename = App.Path & "\english.ini"
 Set ini = New clsINI
 ini.FileName = App.Path & "\english.ini"
 If ini.CheckIniFile = False Then
  MsgBox "File 'english.ini' not found! Program exit!"
  End
 End If
 Set secs = ini.GetAllSectionsFromInifile(, True)
 For i = 1 To secs.Count
  Set Keys = ini.GetAllKeysFromSection(secs.item(i), , , True)
  For j = 1 To Keys.Count
   Set lsvItem = lsv.ListItems.Add(, , "")
   lsvItem.SubItems(1) = secs.item(i)
   lsvItem.SubItems(2) = Keys.item(j)(0)
   lsvItem.SubItems(3) = Keys.item(j)(1)
  Next j
 Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
 RemovePanelControl txtProgramFontsize
 RemovePanelControl cmbCharset
 RemovePanelControl cmbFonts
 mutex.CloseMutex
 Set mutex = Nothing
End Sub

Private Sub lsv_AfterLabelEdit(Cancel As Integer, NewString As String)
 NewString = Trim$(NewString)
 If Len(NewString) > 0 And Len(lsv.SelectedItem.Text) > 0 Then
  RefreshStb
 End If
 If Len(NewString) > 0 And Len(lsv.SelectedItem.Text) = 0 Then
  RefreshStb -1
 End If
 If Len(NewString) = 0 And Len(lsv.SelectedItem.Text) > 0 Then
  RefreshStb 1
 End If
End Sub

Private Sub lsv_Click()
 lsv.StartLabelEdit
End Sub

Private Sub lsv_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  lsv.StartLabelEdit
 End If
End Sub

Private Sub mnExit_Click(Index As Integer)
 Unload Me
 End
End Sub

Private Sub mnFile_Click(Index As Integer)
 On Local Error GoTo ErrorHandler
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
 Dim ini As clsINI, secs As Collection, Keys As Collection, i As Long, j As Long, L As Long
 Me.MousePointer = vbHourglass
 Set ini = New clsINI
 ini.FileName = LanguageIniFilename
 Set secs = ini.GetAllSectionsFromInifile(, True)
 For i = 1 To secs.Count
  Set Keys = ini.GetAllKeysFromSection(secs.item(i), , , True)
  For j = 1 To Keys.Count
   For L = 1 To lsv.ListItems.Count
    If UCase$(lsv.ListItems(L).SubItems(1)) = UCase$(secs.item(i)) And _
     UCase$(lsv.ListItems(L).SubItems(2)) = UCase$(Keys.item(j)(0)) Then
      lsv.ListItems(L).Text = Keys.item(j)(1)
      Exit For
    End If
   Next L
  Next j
 Next i
 Me.MousePointer = vbNormal
End Sub

Private Sub SaveLanguageIniFile(LanguageIniFilename As String)
 Dim ini As clsINI, i As Long
 Set ini = New clsINI
 ini.FileName = LanguageIniFilename
 ini.CreateIniFile
 For i = 1 To lsv.ListItems.Count
  With lsv.ListItems(i)
   ini.SaveKey .Text, .SubItems(2), .SubItems(1)
  End With
 Next i
End Sub

Private Function GetCountEmptyValues()
 Dim i As Long, c As Long
 c = 0
 For i = 1 To lsv.ListItems.Count
  If lsv.ListItems(i).Text = "" Then
   c = c + 1
  End If
 Next i
 GetCountEmptyValues = c
End Function

Private Sub RefreshStb(Optional AddNumber As Long = 0)
 With stb
  .Panels("Keys").Text = "Keys:" & lsv.ListItems.Count
  .Panels("EmptyKeys").Text = "Empty Keys:" & GetCountEmptyValues + AddNumber
 End With
End Sub

Private Sub Timer1_Timer()
 ' Create a mutex if possible
 If mutex.CheckMutex(TransTool_GUID) = False Then
  mutex.CreateMutex TransTool_GUID
 End If
 DoEvents
End Sub

Private Sub txtProgramFontSize_Change()
 Dim tL As Long
 If Trim$(txtProgramFontsize.Text) = "" Then
  txtProgramFontsize.Text = 8
 End If
 tL = CLng(txtProgramFontsize.Text)
 If tL <= 0 Then
  tL = 1
 End If
 If tL > 72 Then
  tL = 72
 End If
 txtProgramFontsize.Text = tL
 lsv.Font.Size = tL
End Sub

Private Sub txtProgramFontSize_KeyPress(KeyAscii As Integer)
 Dim allow As String, tStr As String

 allow = "0123456789" & Chr$(8) & Chr$(13)

 tStr = Chr$(KeyAscii)

 If InStr(1, allow, tStr) = 0 Then
   KeyAscii = 0
 End If
End Sub

