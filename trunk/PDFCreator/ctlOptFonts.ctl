VERSION 5.00
Begin VB.UserControl ctlOptFonts 
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ScaleHeight     =   4950
   ScaleWidth      =   6765
   ToolboxBitmap   =   "ctlOptFonts.ctx":0000
   Begin PDFCreator.dmFrame dmFraProgFont 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _extentx        =   11245
      _extenty        =   8281
      caption         =   "Programfont"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptFonts.ctx":0312
      Begin VB.ComboBox cmbProgramFontsize 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   5400
         TabIndex        =   6
         Text            =   "8"
         Top             =   600
         Width           =   765
      End
      Begin VB.ComboBox cmbFonts 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   4
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox cmbCharset 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   3000
         TabIndex        =   5
         Text            =   "cmbCharset"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtTest 
         Appearance      =   0  '2D
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   8
         Top             =   1320
         Width           =   6135
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "&Test"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   4095
         Width           =   1755
      End
      Begin VB.CommandButton cmdCancelTest 
         Caption         =   "C&ancel test"
         Height          =   495
         Left            =   2310
         TabIndex        =   10
         Top             =   4095
         Width           =   1755
      End
      Begin VB.Label lblProgfont 
         AutoSize        =   -1  'True
         Caption         =   "Programfont"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblProgcharset 
         AutoSize        =   -1  'True
         Caption         =   "Charset"
         Height          =   195
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblTesttext 
         AutoSize        =   -1  'True
         Caption         =   "Here you can test the font."
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1890
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Left            =   5400
         TabIndex        =   3
         Top             =   360
         Width           =   300
      End
   End
End
Attribute VB_Name = "ctlOptFonts"
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
50020  dmFraProgFont.Left = 0
50030  dmFraProgFont.Top = 0
50040  UserControl.Height = dmFraProgFont.Height
50050
50060  Dim i As Long, fi As Long, tStr As String, SMF As Collection, _
  cSystem As clsSystem
50080  Set cSystem = New clsSystem
50090  Set SMF = cSystem.GetSystemFont(frmMain, Menu)
50100  txtTest.Text = vbNullString
50110  For i = 33 To 255
50120   txtTest.Text = txtTest.Text & Chr$(i)
50130 '  If UnloadForm Then
50140 '   TimerReady = True
50150 '   Exit Sub
50160 '  End If
50170   DoEvents
50180  Next i
50190  With cmbCharset
50200   .Clear
50210   .AddItem "0, Western": .ItemData(.NewIndex) = 0
50220   .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
50230   .AddItem "77, Mac": .ItemData(.NewIndex) = 77
50240   .AddItem "128, Japanese": .ItemData(.NewIndex) = 128
50250   .AddItem "129, Hangeul": .ItemData(.NewIndex) = 129
50260   .AddItem "130, Hangeul (Johab)": .ItemData(.NewIndex) = 130
50270   .AddItem "134, Chinese_GB2312": .ItemData(.NewIndex) = 134
50280   .AddItem "136, Chinese_BIG5": .ItemData(.NewIndex) = 136
50290   .AddItem "161, Greek": .ItemData(.NewIndex) = 161
50300   .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
50310   .AddItem "163, Vietnamese": .ItemData(.NewIndex) = 163
50320   .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
50330   .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
50340   .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
50350   .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
50360   .AddItem "222, Thai": .ItemData(.NewIndex) = 222
50370   .AddItem "238, Central European": .ItemData(.NewIndex) = 238
50380   .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
50390   .Text = 0
50400  End With
50410  With cmbProgramFontsize
50420   .AddItem "8"
50430   .AddItem "9"
50440   .AddItem "10"
50450   .AddItem "11"
50460   .AddItem "12"
50470   .AddItem "14"
50480   .AddItem "16"
50490   .AddItem "18"
50500   .AddItem "20"
50510   .AddItem "22"
50520   .AddItem "24"
50530   .AddItem "26"
50540   .AddItem "28"
50550   .AddItem "36"
50560   .AddItem "48"
50570   .AddItem "72"
50580  End With
50590  cmbProgramFontsize.Text = 8
50600  cmbCharset.Text = cmbCharset.ItemData(0)
50610  cmbCharset.Text = Options.ProgramFontCharset
50620  fi = -1
50630  With cmbFonts
50640   For i = 1 To Screen.FontCount
50650    tStr = Trim$(Screen.Fonts(i))
50660    If LenB(tStr) > 0 Then
50670     cmbFonts.AddItem tStr
50680     If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
50690      fi = i
50700     End If
50710    End If
50720 '   If UnloadForm Then
50730 '    TimerReady = True
50740 '    Exit Sub
50750 '   End If
50760    DoEvents
50770   Next i
50780  End With
50790
50800 ' Form_Resize
50810
50820  cmbProgramFontsize.Width = txtTest.Width - _
  (cmbProgramFontsize.Left - txtTest.Left)
50840
50850  If fi >= 0 Then
50860   cmbFonts.ListIndex = fi
50870   cmbCharset.Text = SMF(1)(2)
50880   cmbProgramFontsize.Text = SMF(1)(1)
50890   txtTest.Font = cmbFonts.List(cmbFonts.ListIndex)
50900   txtTest.Font.Charset = cmbCharset.Text
50910  End If
50920
50930  With cmbCharset
50940   .Top = cmbFonts.Top
50950   .Left = lblProgcharset.Left
50960   .Width = 2295
50970   .SelStart = 0
50980   .SelLength = 0
50990  End With
51000  With cmbProgramFontsize
51010   .Top = cmbFonts.Top
51020   .Left = lblSize.Left
51030   .Width = 765
51040   .SelStart = 0
51050   .SelLength = 0
51060  End With
51070
51080  CorrectCmbCharset
51090
51100  SetFrames Options.OptionsDesign
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "UserControl_Initialize")
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
Select Case ErrPtnr.OnError("ctlOptFonts", "SetFrames")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  For Each ctl In Controls
50030 '  If UnloadForm Then
50040 '   TimerReady = True
50050 '   Exit Sub
50060 '  End If
50070   DoEvents
50080   If TypeOf ctl Is ComboBox Then
50090    ComboSetListWidth ctl
50100   End If
50110  Next ctl
50120
50130  SetOptimalComboboxHeigth cmbCharset, Me
50140  SetOptimalComboboxHeigth cmbProgramFontsize, Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "UserControl_ReadProperties")
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
50010  dmFraProgFont.Width = UserControl.Width
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "UserControl_Resize")
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
50020   dmFraProgFont.Caption = .OptionsProgramFontSymbol
50030   lblProgfont.Caption = .OptionsProgramFont
50040   lblProgcharset.Caption = .OptionsProgramFontcharset
50050   lblSize.Caption = .OptionsProgramFontSize
50060   lblTesttext = .OptionsProgramFontTestdescription
50070   cmdTest.Caption = .OptionsProgramFontTest
50080   cmdCancelTest.Caption = .OptionsProgramFontCancelTest
50090  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "SetLanguageStrings")
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
50010  Dim i As Long
50020  With Options
50030   For i = 0 To cmbFonts.ListCount - 1
50040     If UCase$(cmbFonts.List(i)) = UCase$(.ProgramFont) Then
50050      cmbFonts.ListIndex = i
50060      Exit For
50070     End If
50080   Next i
50090   cmbCharset.Text = .ProgramFontCharset
50100   cmbProgramFontsize.Text = .ProgramFontSize
50110  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "SetOptions")
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
50010  With Options
50020   .ProgramFont = cmbFonts.List(cmbFonts.ListIndex)
50030   If LenB(cmbCharset.Text) > 0 Then
50040    .ProgramFontCharset = cmbCharset.Text
50050   End If
50060   If LenB(cmbProgramFontsize.Text) > 0 Then
50070    .ProgramFontSize = cmbProgramFontsize.Text
50080   End If
50090  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "GetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbCharset_Change()
 On Error GoTo ErrorHandler
 txtTest.Font.Charset = cmbCharset.Text
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
50040  txtTest.Font.Charset = cmbCharset.Text
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "cmbCharset_Click")
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
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "cmbCharset_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbCharset_Validate(Cancel As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, tStr As String
50020  tStr = ""
50030  For i = 1 To Len(cmbCharset.Text)
50040   If InStr("0123456789", Mid(cmbCharset.Text, i, 1)) > 0 Then
50050     tStr = tStr & Mid(cmbCharset.Text, i, 1)
50060    Else
50070     Exit For
50080   End If
50090  Next i
50100  If Len(Trim$(tStr)) = 0 Then
50110    cmbCharset.Text = 0
50120   Else
50130    cmbCharset.Text = tStr
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "cmbCharset_Validate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdTest_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tCharset As Long, tStr As String, tFontSize As Long, tFontname As String, _
  tFontCharset As Long
50030  If InStr(1, cmbCharset.Text, ",", vbTextCompare) > 0 Then
50040    tStr = Trim$(Mid$(cmbCharset.Text, 1, InStr(1, cmbCharset.Text, ",", vbTextCompare) - 1))
50050   Else
50060    tStr = Trim$(cmbCharset.Text)
50070  End If
50080  If Len(tStr) = 0 Then
50090   cmbCharset.Text = 0
50100   Exit Sub
50110  End If
50120  If IsNumeric(tStr) = False Then
50130   cmbCharset.Text = 0
50140   Exit Sub
50150  End If
50160  tCharset = tStr
50170  With cmdTest.Font
50180   tFontname = .Name
50190   tFontSize = .Size
50200   tFontCharset = .Charset
50210  End With
50220  SetFontUserControl cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), CLng(cmbProgramFontsize.Text)
50230  cmbCharset.Text = tCharset
50240  SetFont frmMain, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), cmbProgramFontsize.Text
50250 ' ieb.Refresh
50260  With cmdTest.Font
50270   .Name = tFontname
50280   .Size = tFontSize
50290   .Charset = tFontCharset
50300  End With
50310  With cmdCancelTest
50320   .Font.Name = tFontname
50330   .Font.Size = tFontSize
50340   .Font.Charset = tFontCharset
50350   .Enabled = True
50360  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "cmdTest_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdCancelTest_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Options
50020   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50030   cmbCharset.Text = .ProgramFontCharset
50040   SetFont frmMain, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50050 '  ieb.Refresh
50060  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "cmdCancelTest_Click")
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
50010  txtTest.Font.Name = cmbFonts.List(cmbFonts.ListIndex)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "cmbFonts_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub CorrectCmbCharset()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStrf() As String
50020  If InStr(1, cmbCharset.Text, ",", vbTextCompare) > 0 Then
50030    tStrf = Split(cmbCharset.Text, ",")
50040    If Len(tStrf(0)) = 0 Then
50050      cmbCharset.Text = 0
50060     Else
50070      If IsNumeric(tStrf(0)) = False Then
50080        cmbCharset.Text = 0
50090       Else
50100        cmbCharset.Text = CLng(tStrf(0))
50110      End If
50120    End If
50130   Else
50140    If Len(cmbCharset.Text) = 0 Then
50150      cmbCharset.Text = 0
50160     Else
50170      If IsNumeric(cmbCharset.Text) = False Then
50180        cmbCharset.Text = 0
50190       Else
50200        cmbCharset.Text = CLng(cmbCharset.Text)
50210      End If
50220    End If
50230  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "CorrectCmbCharset")
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
50010  Dim tL As Long
50020 If Trim$(cmbProgramFontsize.Text) = vbNullString Then
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
50130  txtTest.Font.Size = tL
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "cmbProgramFontSize_Change")
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
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "cmbProgramFontSize_KeyPress")
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
50010  Dim tL As Long
50020 If Trim$(cmbProgramFontsize.Text) = vbNullString Then
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
50130  txtTest.Font.Size = tL
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "cmbProgramFontsize_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ComboSetListWidth(oCombo As ComboBox, Optional ByVal nFixWidth As Variant, Optional ByVal nScaleMode As Variant)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With oCombo
50020   If IsMissing(nScaleMode) Or IsMissing(nFixWidth) Then
50030    nScaleMode = .Parent.ScaleMode
50040   End If
50050   If IsMissing(nFixWidth) Then
50060    Dim i As Long, nWidth As Long
50070    nFixWidth = 0
50080    For i = 0 To .ListCount - 1
50090     nWidth = .Parent.TextWidth(.List(i))
50100     If nWidth > nFixWidth Then
50110      nFixWidth = nWidth
50120     End If
50130    Next i
50140    nFixWidth = nFixWidth + .Parent.ScaleX(10, vbPixels, nScaleMode)
50150    If .ListCount > 8 Then
50160     nFixWidth = nFixWidth + .Parent.ScaleX(15, vbPixels, nScaleMode)
50170    End If
50180   End If
50190   SendMessage .hwnd, CB_SETDROPPEDWIDTH, .Parent.ScaleX(nFixWidth, nScaleMode, vbPixels), 0&
50200  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "ComboSetListWidth")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetFontUserControl(ByVal Fontname As String, ByVal Charset As Long, ByVal Fontsize As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control, eB As isExplorerBar, ts As TabStrip, df As dmFrame, f As StdFont
50020
50030  If LenB(Trim$(Fontname)) = 0 Then
50040   Exit Sub
50050  End If
50060
50070  Set f = New StdFont
50080  f.Name = Fontname
50090  f.Size = Fontsize
50100  f.Charset = Charset
50110
50120  For Each ctl In UserControl.Controls
50130   If TypeOf ctl Is Label Or _
     TypeOf ctl Is Form Or _
     TypeOf ctl Is ComboBox Or _
     TypeOf ctl Is CheckBox Or _
     TypeOf ctl Is CommandButton Or _
     TypeOf ctl Is ListView Or _
     TypeOf ctl Is StatusBar Or _
     TypeOf ctl Is TextBox Or _
     TypeOf ctl Is Frame Then
50220    With ctl
50230     .Font = Fontname
50240     If Not (TypeOf ctl Is StatusBar) And Not (TypeOf ctl Is ListView) Then
50250      .Fontsize = Fontsize
50260     End If
50270     .Font.Charset = Charset
50280    End With
50290   End If
50300
50310   If TypeOf ctl Is isExplorerBar Then
50320    Set eB = ctl
50330    eB.Font.Name = Fontname
50340    eB.Font.Size = Fontsize
50350    eB.Font.Charset = Charset
50360   End If
50370   If TypeOf ctl Is TabStrip Then
50380    Set ts = ctl
50390    ts.Font.Name = Fontname
50400    ts.Font.Size = Fontsize
50410    ts.Font.Charset = Charset
50420   End If
50430   If TypeOf ctl Is dmFrame Then
50440    Set df = ctl
50450    df.Font.Name = Fontname
50460    df.Font.Size = Fontsize
50470    df.Font.Charset = Charset
50480    Set df.Font = f
50490   End If
50500  Next ctl
50510
50520  Set f = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "SetFontUserControl")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
