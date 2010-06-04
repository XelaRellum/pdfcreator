VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlOptLanguages 
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   ScaleHeight     =   6090
   ScaleWidth      =   6690
   ToolboxBitmap   =   "ctlOptLanguages.ctx":0000
   Begin PDFCreator.dmFrame dmFraProgLanguage 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   10398
      Caption         =   "Language"
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
      Begin VB.CommandButton cmdLanguageRemove 
         Height          =   315
         Left            =   4080
         Picture         =   "ctlOptLanguages.ctx":0312
         Style           =   1  'Grafisch
         TabIndex        =   8
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton cmdLanguageRefresh 
         Caption         =   "Refresh List"
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   1545
         Width           =   1575
      End
      Begin VB.CommandButton cmdLanguageInstall 
         Caption         =   "Install"
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Top             =   2025
         Width           =   1575
      End
      Begin VB.ComboBox cmbCurrentLanguage 
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown-Liste
         TabIndex        =   2
         Top             =   630
         Width           =   3795
      End
      Begin MSComctlLib.ListView lsvTranslations 
         Height          =   4230
         Left            =   120
         TabIndex        =   4
         Top             =   1545
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   7461
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lblEnableNotice 
         Caption         =   "You can set these options in the default profile only."
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.Label lblCurrentLanguage 
         AutoSize        =   -1  'True
         Caption         =   "Current language"
         Height          =   195
         Left            =   105
         TabIndex        =   1
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lblLanguagesFromInternet 
         AutoSize        =   -1  'True
         Caption         =   "Load more languages from the internet"
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   1260
         Width           =   2715
      End
   End
End
Attribute VB_Name = "ctlOptLanguages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents dl As clsDownload
Attribute dl.VB_VarHelpID = -1

Public ParentForm As Form

Private LangFiles As Collection, Languages As Collection, FirstStart As Boolean, OldLangListIndex As Long

Private mEnabled As Boolean
Private mControlsEnabled As Boolean

Public Sub SetControlsEnabled(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mControlsEnabled = value
50020  ControlsEnabled = value
50030  dmFraProgLanguage.Enabled = value
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "SetControlsEnabled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Property Let ControlsEnabled(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mEnabled = value
50020
50030  lblCurrentLanguage.Enabled = mEnabled
50040  lblCurrentLanguage.Visible = mEnabled
50050  cmbCurrentLanguage.Enabled = mEnabled
50060  cmbCurrentLanguage.Visible = mEnabled
50070  cmdLanguageRemove.Enabled = mEnabled
50080  cmdLanguageRemove.Visible = mEnabled
50090  lblLanguagesFromInternet.Enabled = mEnabled
50100  lblLanguagesFromInternet.Visible = mEnabled
50110  lsvTranslations.Enabled = mEnabled
50120  lsvTranslations.Visible = mEnabled
50130  cmdLanguageRefresh.Enabled = mEnabled
50140  cmdLanguageRefresh.Visible = mEnabled
50150  cmdLanguageInstall.Enabled = mEnabled
50160  cmdLanguageInstall.Visible = mEnabled
50170  lblEnableNotice.Visible = Not mEnabled
50180  If mControlsEnabled Then
50190    lblEnableNotice.Enabled = Not mEnabled
50200   Else
50210    lblEnableNotice.Enabled = False
50220  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "ControlsEnabled [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get ControlEnabled() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ControlEnabled = mEnabled
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "ControlEnabled [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Sub UserControl_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  mEnabled = True
50030  dmFraProgLanguage.Left = 0
50040  dmFraProgLanguage.Top = 0
50050  UserControl.Height = dmFraProgLanguage.Height
50060
50070  lblEnableNotice.Top = lblCurrentLanguage.Top
50080  lblEnableNotice.Left = lblCurrentLanguage.Left
50090
50100  lsvTranslations.ColumnHeaders.Add , , ""
50110  lsvTranslations.ColumnHeaders.Add , , ""
50120  lsvTranslations.ColumnHeaders(1).Width = 2000
50130  lsvTranslations.ColumnHeaders(2).Width = 1500
50140
50150  ReadAllLanguages LanguagePath, True
50160
50170  mControlsEnabled = True
50180  SetFrames Options.OptionsDesign
50190
50200  SetFont
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "UserControl_Initialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetFont()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Options
50020   SetFontControls UserControl.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50030  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "SetFont")
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
Select Case ErrPtnr.OnError("ctlOptLanguages", "SetFrames")
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
50010  dmFraProgLanguage.Width = UserControl.Width
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "UserControl_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Terminate()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  FirstStart = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "UserControl_Terminate")
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
50020   dmFraProgLanguage.Caption = .OptionsProgramLanguagesSymbol
50030   cmdLanguageInstall.Caption = .OptionsLanguagesInstall
50040   cmdLanguageRefresh.Caption = .OptionsLanguagesRefresh
50050   lblLanguagesFromInternet.Caption = .OptionsLanguagesDownloadMoreLanguages
50060   lsvTranslations.ColumnHeaders(1).Text = .OptionsLanguagesTranslation
50070   lsvTranslations.ColumnHeaders(2).Text = .OptionsLanguagesVersion
50080   lblCurrentLanguage.Caption = .OptionsLanguagesCurrentLanguage
50090   lblEnableNotice.Caption = .OptionsEnableNotice
50100  End With
50110
50120  If FirstStart = False Then
50130   FirstStart = True
50140   ReadAllLanguages LanguagePath, True
50150   SetOptimalComboboxHeigth cmbCurrentLanguage, Me
50160  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "SetLanguageStrings")
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
50010  Dim filename As String, i As Long, lang As String
50020  SplitPath CurrentLanguage, , , , lang
50030  If mEnabled Then
50040   For i = 1 To LangFiles.Count
50050    SplitPath LangFiles(i), , , , filename
50060    If UCase$(lang) = UCase$(filename) Then
50070     If cmbCurrentLanguage.ListIndex <> i - 1 Then
50080      cmbCurrentLanguage.ListIndex = i - 1
50090     End If
50100     Exit For
50110    End If
50120   Next i
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "SetOptions")
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
50010  Dim lang As String
50020  With Options1
50030   If cmbCurrentLanguage.ListIndex < 0 Then
50040     .Language = Options.Language
50050    Else
50060     SplitPath LangFiles(cmbCurrentLanguage.ListIndex + 1), , , , lang
50070     .Language = lang
50080   End If
50090  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "GetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Property Let ScaleMode(value As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UserControl.ScaleMode = value
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "ScaleMode [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get ScaleMode() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ScaleMode = UserControl.ScaleMode
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "ScaleMode [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get hwnd() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  hwnd = UserControl.hwnd
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "hwnd [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Sub cmbCurrentLanguage_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim f As Form
50020  If OldLangListIndex <> cmbCurrentLanguage.ListIndex Then
50030   OldLangListIndex = cmbCurrentLanguage.ListIndex
50040   If InStr(1, LangFiles(cmbCurrentLanguage.ListIndex + 1), LanguagePath, vbTextCompare) = 1 Then
50050     cmdLanguageRemove.Enabled = False
50060    Else
50070     cmdLanguageRemove.Enabled = True
50080   End If
50090   CurrentLanguage = Languages(cmbCurrentLanguage.ListIndex + 1)
50100   LoadLanguage LangFiles(cmbCurrentLanguage.ListIndex + 1)
50110   For Each f In Forms
50120    f.ChangeLanguage
50130   Next
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "cmbCurrentLanguage_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdLanguageInstall_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strInstallPath As String
50020  Const strDownloadPath = "http://www.pdfforge.org/files/translations/"
50030  strInstallPath = CompletePath(GetMyAppData()) & "PDFCreator\Languages"
50040  If Not DirExists(CompletePath(GetMyAppData()) & "PDFCreator") Then
50050   CreateDir CompletePath(GetMyAppData()) & "PDFCreator"
50060  End If
50070
50080  If Not DirExists(GetMyAppData() & "\PDFCreator\Languages") Then
50090   CreateDir GetMyAppData() & "\PDFCreator\Languages"
50100  End If
50110  If lsvTranslations.SelectedItem Is Nothing Then
50120   Exit Sub
50130  End If
50140  InstallInternetLanguageFile lsvTranslations.SelectedItem.Text, lsvTranslations.SelectedItem.SubItems(1), strDownloadPath, strInstallPath
50150  ReadAllLanguages LanguagePath, True
50160  SetOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "cmdLanguageInstall_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdLanguageRefresh_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strLanguages() As String, strFile() As String, i As Long
50020  Const strDownloadURL = "http://www.pdfforge.org/products/pdfcreator/translations/list"
50030  MousePointer = vbHourglass
50040  Set dl = New clsDownload
50050  strLanguages = Split(dl.DownloadString(strDownloadURL), vbLf)
50060  Set dl = Nothing
50070  lsvTranslations.ListItems.Clear
50080  For i = LBound(strLanguages) To UBound(strLanguages)
50090   If ((strLanguages(i) <> vbNullString) And (InStr(1, strLanguages(i), ":"))) Then
50100    strFile = Split(strLanguages(i), ":")
50110    lsvTranslations.ListItems.Add(, , strFile(0)).SubItems(1) = strFile(1)
50120   End If
50130  Next i
50140  MousePointer = vbDefault
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "cmdLanguageRefresh_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdLanguageRemove_Click()
 On Error GoTo ErrorHandler
 Dim oldLanguage As String
 Kill LangFiles(cmbCurrentLanguage.ListIndex + 1)
 If StrComp(LangFiles(cmbCurrentLanguage.ListIndex + 1), oldLanguage, vbTextCompare) <> 0 Then
  Options.Language = "english"
 End If
 ReadAllLanguages LanguagePath, True
 Exit Sub
ErrorHandler:
 MsgBox Err.Description
End Sub

Private Sub ReadAllLanguages(LanguagePath As String, Optional UserPath As Boolean = False)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Languagename As String, ini As clsINI, UserLangFiles As Collection, _
  i As Long, j As Long, found As Boolean, Version As String, filename As String, _
  UserLanguagePath As String
50040
50050  cmbCurrentLanguage.Clear
50060  cmbCurrentLanguage.AddItem "No languages available."
50070  Set LangFiles = GetAllLanguagesFiles(LanguagePath)
50080
50090  If UserPath Then
50100   UserLanguagePath = GetMyAppData() & "\PDFCreator\Languages"
50110   Set UserLangFiles = GetAllLanguagesFiles(UserLanguagePath)
50120   For i = 1 To UserLangFiles.Count
50130    For j = 1 To LangFiles.Count
50140     If GetFilenameFromPath(LangFiles(j)) = GetFilenameFromPath(UserLangFiles(i)) Then
50150      LangFiles.Remove j
50160      Exit For
50170     End If
50180    Next j
50190    LangFiles.Add UserLangFiles(i)
50200   Next i
50210  End If
50220
50230  Set Languages = New Collection
50240  For i = 1 To LangFiles.Count
50250   SplitPath LangFiles(i), , , , filename
50260   Languages.Add filename
50270  Next i
50280
50290  Set ini = New clsINI
50300  For i = 1 To LangFiles.Count
50310   ini.filename = LangFiles.Item(i)
50320   ini.Section = "Common"
50330   Languagename = ini.GetKeyFromSection("Languagename")
50340   Version = ini.GetKeyFromSection("Version")
50350   If Len(Languagename) = 0 Then
50360    Languagename = "No name available."
50370   End If
50380   If IsCompatibleLanguageVersion(Version) Then
50390     If i = 1 Then
50400       cmbCurrentLanguage.List(0) = Languagename
50410      Else
50420       cmbCurrentLanguage.AddItem Languagename
50430     End If
50440    Else
50450     If i = 1 Then
50460       cmbCurrentLanguage.List(0) = Languagename & " [" & Version & "]"
50470      Else
50480       cmbCurrentLanguage.AddItem Languagename & " [" & Version & "]"
50490     End If
50500   End If
50510
50520 '  SplitPath LangFiles.Item(i), , , , filename
50530 '  If UCase$(Options.Language) = UCase$(filename) Then
50540 '   LangListIndex = i - 1
50550 '  End If
50560   DoEvents
50570  Next i
50580  Set ini = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "ReadAllLanguages")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetAllLanguagesFiles(LanguagePath As String) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tColl1 As Collection, tColl2 As Collection, i As Long, tStrf() As String, ini As clsINI, _
  Languagename As String
50030  Set GetAllLanguagesFiles = New Collection
50040  Set tColl1 = GetFiles(LanguagePath, "*.ini", SortedByName)
50050  Set tColl2 = New Collection
50060  For i = 1 To tColl1.Count
50070   tStrf = Split(tColl1(i), "|")
50080   Set ini = New clsINI
50090   ini.filename = tStrf(1)
50100   ini.Section = "Common"
50110   Languagename = ini.GetKeyFromSection("Languagename")
50120   If Len(Languagename) = 0 Then
50130    Languagename = "No name available."
50140   End If
50150   AddSortedStr tColl2, Languagename & "|" & tStrf(1)
50160  Next i
50170  For i = 1 To tColl2.Count
50180   tStrf() = Split(tColl2(i), "|")
50190   GetAllLanguagesFiles.Add tStrf(1)
50200  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "GetAllLanguagesFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function IsCompatibleLanguageVersion(Version As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Byte, delim As String, fVers() As String, fCVers() As String, _
  ProgVersion As String, fPVers() As String
50030  IsCompatibleLanguageVersion = False
50040  delim = "."
50050  ProgVersion = GetProgramRelease
50060  If Len(CompatibleLanguageVersion) = 0 Or Len(Version) = 0 Or Len(ProgVersion) = 0 Then
50070   Exit Function
50080  End If
50090  If InStr(1, CompatibleLanguageVersion, delim) = 0 Or _
    InStr(1, Version, delim) = 0 Or _
    InStr(1, ProgVersion, delim) = 0 Then
50120   Exit Function
50130  End If
50140  fVers = Split(Version, delim)
50150  fCVers = Split(CompatibleLanguageVersion, delim)
50160  fPVers = Split(ProgVersion, delim)
50170  If UBound(fVers) < 2 Or UBound(fCVers) < 2 Or UBound(fPVers) < 2 Then
50180   Exit Function
50190  End If
50200  For i = 0 To 2
50210   If IsNumeric(fVers(i)) = False Or IsNumeric(fCVers(i)) = False Or _
   IsNumeric(fPVers(i)) = False Then
50230    Exit Function
50240   End If
50250  Next i
50260  If CLng(fVers(0)) >= CLng(fCVers(0)) And CLng(fVers(0)) <= CLng(fPVers(0)) Then
50270   If CLng(fVers(1)) >= CLng(fCVers(1)) And CLng(fVers(1)) <= CLng(fPVers(1)) Then
50280    If CLng(fVers(2)) >= CLng(fCVers(2)) And CLng(fVers(2)) <= CLng(fPVers(2)) Then
50290     IsCompatibleLanguageVersion = True
50300    End If
50310   End If
50320  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "IsCompatibleLanguageVersion")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function InstallInternetLanguageFile(File As String, Version As String, DownloadURL As String, ProgramPath As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strLangFile As String, strFile As String
50020  InstallInternetLanguageFile = False
50030  If (File = vbNullString) Or (Version = vbNullString) Or (DownloadURL = vbNullString) Or (ProgramPath = vbNullString) Then
50040   Exit Function
50050  End If
50060  ProgramPath = CompletePath(ProgramPath)
50070  If Right$(DownloadURL, 1) <> "/" Then
50080   DownloadURL = DownloadURL & "/"
50090  End If
50100
50110  Set dl = New clsDownload
50120  strLangFile = dl.DownloadString(DownloadURL & Version & "/" & File)
50130  Set dl = Nothing
50140
50150  If InStr(1, strLangFile, "[Common]", vbTextCompare) = 0 Then
50160   MsgBox LanguageStrings.MessagesMsg37, vbCritical
50170   Exit Function
50180  End If
50190
50200  If Not DirExists(ProgramPath) Then
50210   MsgBox LanguageStrings.MessagesMsg10, vbCritical
50220   Exit Function
50230  End If
50240
50250  strFile = ProgramPath & File
50260
50270  If FileExists(strFile) Then
50280   If MsgBox(LanguageStrings.MessagesMsg05, vbYesNo) = vbNo Then
50290    Exit Function
50300   End If
50310  End If
50320
50330  Open strFile For Output As #1
50340  Print #1, strLangFile
50350  Close #1
50360
50370  MsgBox LanguageStrings.MessagesMsg38, vbInformation
50380
50390  InstallInternetLanguageFile = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptLanguages", "InstallInternetLanguageFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
