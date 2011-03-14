VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "<"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Top             =   3390
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Search"
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   3
      Top             =   3390
      Width           =   1215
   End
   Begin TransTool.XP_ProgressBar xpPgb 
      Height          =   225
      Left            =   3255
      TabIndex        =   13
      Top             =   3885
      Width           =   1380
      _extentx        =   2434
      _extenty        =   397
      font            =   "frmSearch.frx":000C
      brushstyle      =   0
      color           =   65280
   End
   Begin TransTool.dmFrame dmFraSettings 
      Height          =   1275
      Left            =   105
      TabIndex        =   11
      Top             =   1785
      Width           =   4530
      _extentx        =   7990
      _extenty        =   2249
      caption         =   "Settings"
      font            =   "frmSearch.frx":0038
      Begin VB.CheckBox chkWholeWord 
         Appearance      =   0  '2D
         Caption         =   "&Whole word"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   2
         Top             =   780
         Width           =   4215
      End
      Begin VB.CheckBox chkCaseSensitive 
         Appearance      =   0  '2D
         Caption         =   "Case s&ensitive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   1
         Top             =   420
         Width           =   4215
      End
   End
   Begin TransTool.dmFrame dmFraSearch 
      Height          =   1590
      Left            =   105
      TabIndex        =   7
      Top             =   105
      Width           =   4530
      _extentx        =   7990
      _extenty        =   2805
      caption         =   "Search"
      font            =   "frmSearch.frx":0064
      Begin VB.ComboBox cmbSearchtext 
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   1155
         Width           =   4350
      End
      Begin VB.ComboBox cmbColumn 
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   9
         Top             =   525
         Width           =   4350
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Choice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   315
         Width           =   495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Searchtext"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   10
         Top             =   945
         Width           =   765
      End
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Unten ausrichten
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   4080
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Index           =   1
      Left            =   105
      TabIndex        =   6
      Top             =   3390
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   ">"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   4
      Top             =   3390
      Width           =   375
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastFounditem As Long, Backspaced As Boolean

Private Sub cmbColumn_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set hLItems = New Collection
50020  xpPgb.Visible = True
50030  HighlightListitems frmMain.lsv, hLItems, stb, xpPgb, "Status", lsvColor, lsvBold
50040  xpPgb.Visible = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "cmbColumn_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmd_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 0
50030    SearchText
50040   Case 1
50050    EnableControls False
50060    Unload Me
50070  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "cmd_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 0
50030    ShowFoundItems frmMain.lsv, hLItems, LastFounditem - 1, Me
50040    cmdBrowse(1).Enabled = True
50050    If LastFounditem = 1 Then
50060      cmdBrowse(0).Enabled = False
50070     Else
50080      MoveMouseToCommandButton cmdBrowse(0)
50090    End If
50100   Case 1
50110    ShowFoundItems frmMain.lsv, hLItems, LastFounditem + 1, Me
50120    cmdBrowse(0).Enabled = True
50130    If LastFounditem = hLItems.Count Then
50140      cmdBrowse(1).Enabled = False
50150     Else
50160      MoveMouseToCommandButton cmdBrowse(1)
50170    End If
50180  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "cmdBrowse_Click")
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
50010  Dim i As Long
50020  Me.Icon = frmMain.Icon
50030  lsvColor = frmMain.lsv.ForeColor
50040  lsvBold = frmMain.lsv.Font.Bold
50050  Me.Caption = "Search"
50060  With cmbColumn
50070   .AddItem "All"
50080   .ItemData(.NewIndex) = 0
50090   .AddItem "Translated text"
50100   .ItemData(.NewIndex) = 1
50110   .AddItem "English text"
50120   .ItemData(.NewIndex) = 4
50130   .AddItem "Section"
50140   .ItemData(.NewIndex) = 2
50150   .AddItem "Key"
50160   .ItemData(.NewIndex) = 3
50170   .ListIndex = 0
50180  End With
50190  With xpPgb
50200   .Scrolling = ccScrollingStandard
50210   .ShowText = True
50220   .Color = RGB(&H80, &HFF, &H80)
50230   .Font.Bold = True
50240   .Visible = False
50250  End With
50260  With stb.Panels
50270   .Clear
50280   .Add , "Status"
50290   .Add , "FoundIndex"
50300   .Add , "Progress"
50310  End With
50320
50330  With stb
50340   .Panels("Status").ToolTipText = .Panels("Status").key
50350   .Panels("FoundIndex").ToolTipText = .Panels("FoundIndex").key
50360   .Panels("Progress").ToolTipText = .Panels("Progress").key
50370  End With
50380
50390  Set cmbSearchtext.Font = frmMain.lsv.Font
50400
50410  Set hLItems = New Collection
50420  SetLastSearchstrings
50430  ShowAcceleratorsInForm Me, True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "Form_Load")
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
50010  With stb
50020   .Panels("Status").Width = 1500
50030   .Panels("FoundIndex").Width = 500
50040   .Panels("Progress").Width = stb.Width - stb.Panels("Status").Width - 100
50050  End With
50060  With dmFraSearch
50070   .Top = 100
50080   .Left = 100
50090   .Width = ScaleWidth - 200
50100  End With
50110  With dmFraSettings
50120   .Top = dmFraSearch.Top + dmFraSearch.Height + 100
50130   .Left = dmFraSearch.Left
50140   .Width = dmFraSearch.Width
50150  End With
50160  cmbColumn.Width = dmFraSearch.Width - 220
50170  cmbSearchtext.Width = dmFraSearch.Width - 220
50180  With cmd(1)
50190   .Top = dmFraSettings.Top + dmFraSettings.Height + 50
50200   .Left = dmFraSettings.Left
50210  End With
50220  With cmd(0)
50230   .Top = cmd(1).Top
50240   .Left = dmFraSettings.Left + dmFraSettings.Width - cmd(0).Width
50250  End With
50260  With cmdBrowse(0)
50270   .Top = cmd(1).Top
50280   .Left = ScaleWidth / 2 - cmdBrowse(0).Width - 50
50290  End With
50300  With cmdBrowse(1)
50310   .Top = cmdBrowse(0).Top
50320   .Left = ScaleWidth / 2 + 50
50330  End With
50340  Height = cmd(1).Top + cmd(1).Height + stb.Height + (Height - ScaleHeight) + 100
50350  SetPanelControl xpPgb, stb, "Progress", False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "Form_Resize")
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
50010  RemovePanelControl xpPgb
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "Form_Unload")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbSearchtext_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, pos As Long
50020  If Len(cmbSearchtext.Text) = 0 Then
50030    cmd(0).Enabled = False
50040   Else
50050    cmd(0).Enabled = True
50060  End If
50070
50080  With cmbSearchtext
50090   If Backspaced = True Or .Text = "" Then
50100     Backspaced = False
50110    Else
50120     For i = 0 To .ListCount - 1
50130      If InStr(1, .List(i), .Text) <> 0 Then
50140       pos = .SelStart
50150       .Text = .List(i)
50160       .SelStart = pos
50170       .SelLength = Len(.Text) - pos
50180       Exit For
50190      End If
50200     Next i
50210   End If
50220  End With
50230  ComboSetListWidth cmbSearchtext
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "cmbSearchtext_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbSearchtext_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Len(cmbSearchtext.Text) > 0 Then
50020   If KeyAscii = 13 Then
50030    SearchText
50040   End If
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "cmbSearchtext_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbSearchtext_KeyDown(KeyCode As Integer, Shift As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
50020     If cmbSearchtext.Text <> "" Then Backspaced = True
50030   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "cmbSearchtext_KeyDown")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function StringCompare(String1 As String, String2 As String, _
 Optional CaseSensitive As Integer = vbTextCompare, _
 Optional WholeWord As Boolean = False) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  If WholeWord = True Then
50030    If Len(String1) <> Len(String2) Then
50040     StringCompare = False
50050     Exit Function
50060    End If
50070    If CaseSensitive = vbTextCompare Then
50080      If UCase$(String1) = UCase$(String2) Then
50090        StringCompare = True
50100        Exit Function
50110       Else
50120        StringCompare = False
50130        Exit Function
50140      End If
50150     Else
50160      If String1 = String2 Then
50170        StringCompare = True
50180        Exit Function
50190       Else
50200        StringCompare = False
50210        Exit Function
50220      End If
50230    End If
50240   Else
50250    If Len(String1) = 0 And Len(String2) = 0 Then
50260     StringCompare = True
50270     Exit Function
50280    End If
50290    If Len(String1) = 0 And Len(String2) > 0 Then
50300     StringCompare = False
50310     Exit Function
50320    End If
50330    If Len(String1) > 0 And Len(String2) = 0 Then
50340     StringCompare = False
50350     Exit Function
50360    End If
50370    If CaseSensitive = vbTextCompare Then
50380     If InStr(1, String1, String2, CaseSensitive) > 0 Then
50390       StringCompare = True
50400       Exit Function
50410      Else
50420       StringCompare = False
50430       Exit Function
50440     End If
50450    End If
50460  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "StringCompare")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub SearchText()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim CaseSensitive As Boolean, WholeWord As Boolean, i As Long
50020  EnableControls False
50030  If chkCaseSensitive.Value = 1 Then
50040    CaseSensitive = True
50050   Else
50060    CaseSensitive = False
50070  End If
50080  If chkWholeWord.Value = 1 Then
50090    WholeWord = True
50100   Else
50110    WholeWord = False
50120  End If
50130  LastFounditem = 0
50140  stb.Panels("FoundIndex").Text = ""
50150  xpPgb.Visible = True
50160  HighlightListitems frmMain.lsv, hLItems, stb, xpPgb, "Status", lsvColor, False, "Unmarking ..."
50170  xpPgb.Visible = False
50180  Set hLItems = SearchTextInListview(frmMain.lsv, cmbSearchtext.Text, _
  cmbColumn.ItemData(cmbColumn.ListIndex), CaseSensitive, WholeWord)
50200  xpPgb.Visible = True
50210  HighlightListitems frmMain.lsv, hLItems, stb, xpPgb, "Status", vbRed, False
50220  xpPgb.Visible = False
50230  ShowFoundItems frmMain.lsv, hLItems, 1, Me
50240  CheckForBrowse
50250  stb.Panels("Status").Text = "Found: " & CStr(hLItems.Count): stb.Refresh
50260  LastSearchstrings.Add cmbSearchtext.Text
50270  For i = 1 To LastSearchstrings.Count - 1
50280   If LastSearchstrings(i) = cmbSearchtext.Text Then
50290    LastSearchstrings.Remove i
50300    Exit For
50310   End If
50320  Next i
50330  SetLastSearchstrings
50340  If hLItems.Count > 0 Then
50350   frmMain.mnEdit(1).Enabled = True
50360   frmMain.tlb.Buttons("unmark").Enabled = True
50370  End If
50380  EnableControls True
50390  cmbSearchtext.SetFocus
50400  If hLItems.Count > 1 Then
50410    MoveMouseToCommandButton cmdBrowse(1)
50420   Else
50430    cmbSearchtext.SetFocus
50440  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "SearchText")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function SearchTextInListview(lsv As ListView, SearchText As String, _
 Optional Column As Long = 0, Optional CaseSensitive As Boolean = False, Optional WholeWord As Boolean = False) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, j As Long, lIndex(1) As Long, compareMethode As Integer
50020  Set SearchTextInListview = New Collection
50030  stb.Panels("Status").Text = "Searching ...": stb.Refresh
50040  If CaseSensitive = True Then
50050    compareMethode = vbBinaryCompare
50060   Else
50070    compareMethode = vbTextCompare
50080  End If
50090  With xpPgb
50100   .Min = 0
50110   .Max = lsv.ListItems.Count
50120   .Visible = True
50130  End With
50140  For i = 1 To lsv.ListItems.Count
50150   xpPgb.Value = i
50160   For j = 1 To lsv.ColumnHeaders.Count - 1
50170    If Len(lsv.ListItems(i).ListSubItems(j)) > 0 And (Column = 0 Or Column + 1 = j) Then
50180     If StringCompare(lsv.ListItems(i).ListSubItems(j), SearchText, compareMethode, WholeWord) = True Then
50190      lIndex(0) = i
50200      lIndex(1) = j
50210      SearchTextInListview.Add lIndex
50220     End If
50230    End If
50240   Next j
50250   If Len(lsv.ListItems(i).Text) > 0 And (Column = 0 Or Column = 1) Then
50260    If StringCompare(lsv.ListItems(i).Text, SearchText, compareMethode, WholeWord) = True Then
50270     lIndex(0) = i
50280     lIndex(1) = 0
50290     SearchTextInListview.Add lIndex
50300    End If
50310   End If
50320  Next i
50330  With xpPgb
50340   .Min = 0
50350   .Max = 1
50360   .Visible = False
50370  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "SearchTextInListview")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub EnableControls(Enable As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cmd(0).Enabled = Enable
50020  cmd(1).Enabled = Enable
50030  chkWholeWord.Enabled = Enable
50040  chkCaseSensitive.Enabled = Enable
50050  cmbColumn.Enabled = Enable
50060  cmbSearchtext.Enabled = Enable
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "EnableControls")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetLastSearchstrings()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  With cmbSearchtext
50030   .Clear
50040   If LastSearchstrings.Count > 0 Then
50050    For i = LastSearchstrings.Count To 1 Step -1
50060     .AddItem LastSearchstrings(i)
50070    Next i
50080    .Text = .List(0)
50090    .SelStart = 0
50100    .SelLength = Len(.Text)
50110   End If
50120  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "SetLastSearchstrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ShowFoundItems(lsv As ListView, Items As Collection, ItemIndex As Long, Moveform As Form)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim rectList As Rect, Item As ListItem, lsvItemTop As Long, lsvItemLeft As Long, _
  itemHeight As Long, itemWidth As Long, sTPPX As Long, sTPPY As Long, pt As POINTAPI, _
  cRect As Rect
50040  If Items.Count > 0 Then
50050
  lsv.ListItems(Items(ItemIndex)(0)).EnsureVisible: lsv.Refresh
50070   If FormOverListitem(Me, frmMain.lsv, CLng(Items(ItemIndex)(0)), CLng(Items(ItemIndex)(1))) = True Then
50080    MoveFormAwayFromListitem Me, frmMain.lsv, CLng(Items(ItemIndex)(0)), CLng(Items(ItemIndex)(1))
50090   End If
50100   If LastFounditem > 0 Then
50110    If Items(LastFounditem)(1) = 0 Then
50120      lsv.ListItems(Items(LastFounditem)(0)).Bold = False
50130     Else
50140      lsv.ListItems(Items(LastFounditem)(0)).ListSubItems(Items(LastFounditem)(1)).Bold = False
50150    End If
50160   End If
50170   If Items(ItemIndex)(1) = 0 Then
50180     lsv.ListItems(Items(ItemIndex)(0)).Bold = True
50190    Else
50200     lsv.ListItems(Items(ItemIndex)(0)).ListSubItems(Items(ItemIndex)(1)).Bold = True
50210   End If
50220   lsv.Refresh
50230   LastFounditem = ItemIndex
50240   stb.Panels("FoundIndex").Text = CStr(ItemIndex)
50250  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "ShowFoundItems")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CheckForBrowse()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Enable As Boolean
50020  If hLItems.Count > 1 Then
50030    Enable = True
50040   Else
50050    Enable = False
50060  End If
50070  cmdBrowse(0).Enabled = False
50080  cmdBrowse(1).Enabled = Enable
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "CheckForBrowse")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function FormOverListitem(Frm As Form, lsv As ListView, ListItem As Long, Subitem As Long) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tRect As Rect, sTPPX As Long, sTPPY As Long, lsvTop As Long, lsvLeft As Long, _
 sItemTop As Long, sItemLeft As Long, sItemBottom As Long, sItemRight As Long, _
 sItemTopLeftCorner As POINTAPI, sItemTopRightCorner As POINTAPI, _
 sItemBottomLeftCorner As POINTAPI, sItemBottomRightCorner As POINTAPI
50050
50060  FormOverListitem = False
50070  sTPPX = Screen.TwipsPerPixelX: sTPPY = Screen.TwipsPerPixelY
50080
50090  GetWindowRect lsv.hwnd, tRect
50100  lsvTop = tRect.Top: lsvLeft = tRect.Left
50110
50120  tRect.Top = Subitem: tRect.Left = LVIR_LABEL
50130  tRect.Bottom = 0: tRect.Right = 0
50140  Call SendMessage(lsv.hwnd, LVM_GETSUBITEMRECT, ListItem - 1, tRect)
50150
50160  sItemTop = lsvTop + tRect.Top: sItemLeft = lsvLeft + tRect.Left
50170  sItemBottom = lsvTop + tRect.Bottom: sItemRight = lsvLeft + tRect.Right
50180  sItemTopLeftCorner.x = sItemLeft: sItemTopLeftCorner.Y = sItemTop
50190  sItemTopRightCorner.x = sItemRight: sItemTopRightCorner.Y = sItemTop
50200  sItemBottomLeftCorner.x = sItemLeft: sItemBottomLeftCorner.Y = sItemBottom
50210  sItemBottomRightCorner.x = sItemRight: sItemBottomRightCorner.Y = sItemBottom
50220
50230  With tRect
50240   .Top = Frm.Top \ sTPPY: .Left = Frm.Left \ sTPPX
50250   .Bottom = .Top + Frm.Height \ sTPPY: .Right = .Left + Frm.Width \ sTPPX
50260  End With
50270
50280  If PointInRectangle(sItemTopLeftCorner, tRect) = True Or _
    PointInRectangle(sItemTopRightCorner, tRect) = True Or _
    PointInRectangle(sItemBottomLeftCorner, tRect) = True Or _
    PointInRectangle(sItemBottomRightCorner, tRect) = True Then
50320   FormOverListitem = True
50330  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "FormOverListitem")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub MoveFormAwayFromListitem(Frm As Form, lsv As ListView, ListItem As Long, Subitem As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tRect As Rect, sTPPX As Long, sTPPY As Long, lsvTop As Long, _
  sItemTop As Long, sItemBottom As Long, offs As Long
50030
50040  sTPPX = Screen.TwipsPerPixelX: sTPPY = Screen.TwipsPerPixelY
50050
50060  GetWindowRect lsv.hwnd, tRect
50070  lsvTop = tRect.Top
50080
50090  tRect.Top = Subitem: tRect.Left = LVIR_LABEL
50100  tRect.Bottom = 0: tRect.Right = 0
50110  Call SendMessage(lsv.hwnd, LVM_GETSUBITEMRECT, ListItem - 1, tRect)
50120
50130  sItemTop = lsvTop + tRect.Top: sItemBottom = lsvTop + tRect.Bottom
50140
50150  offs = 0
50160  If IsTaskBarOnTop = True Then
50170   If TaskBarAlign = tbaBottom Then
50180    offs = TaskBarHeight
50190   End If
50200  End If
50210
50220  If sItemBottom + Frm.Height \ sTPPY > Screen.Height \ sTPPY - offs Then
50230    Frm.Top = (sItemTop + 1) * sTPPY - Frm.Height
50240   Else
50250    Frm.Top = (sItemBottom + 1) * sTPPY
50260  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "MoveFormAwayFromListitem")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function PointInRectangle(tP As POINTAPI, tRect As Rect) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  PointInRectangle = False
50020  With tRect
50030   If tP.x > .Left And tP.x < .Right And tP.Y > .Top And tP.Y < .Bottom Then
50040     PointInRectangle = True
50050    Else
50060     PointInRectangle = False
50070   End If
50080  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSearch", "PointInRectangle")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
