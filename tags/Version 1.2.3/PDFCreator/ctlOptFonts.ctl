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
      Begin VB.Label lblEnableNotice 
         Caption         =   "You can set these options in the default profile only."
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   4800
         Visible         =   0   'False
         Width           =   5895
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

Private mEnabled As Boolean
Private mControlsEnabled As Boolean

Public SetTestFontBack As Boolean

Public Sub SetControlsEnabled(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mControlsEnabled = value
50020  ControlsEnabled = value
50030  dmFraProgFont.Enabled = value
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "SetControlsEnabled")
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
50030  cmbFonts.Enabled = mEnabled
50040  cmbFonts.Visible = mEnabled
50050  cmbCharset.Enabled = mEnabled
50060  cmbCharset.Visible = mEnabled
50070  cmbProgramFontsize.Enabled = mEnabled
50080  cmbProgramFontsize.Visible = mEnabled
50090  lblProgfont.Enabled = mEnabled
50100  lblProgfont.Visible = mEnabled
50110  lblProgcharset.Enabled = mEnabled
50120  lblProgcharset.Visible = mEnabled
50130  lblSize.Enabled = mEnabled
50140  lblSize.Visible = mEnabled
50150  lblTesttext.Enabled = mEnabled
50160  lblTesttext.Visible = mEnabled
50170  txtTest.Enabled = mEnabled
50180  txtTest.Visible = mEnabled
50190  cmdTest.Enabled = mEnabled
50200  cmdTest.Visible = mEnabled
50210  cmdCancelTest.Enabled = mEnabled
50220  cmdCancelTest.Visible = mEnabled
50230
50240  lblEnableNotice.Visible = Not mEnabled
50250  If mControlsEnabled Then
50260    lblEnableNotice.Enabled = Not mEnabled
50270   Else
50280    lblEnableNotice.Enabled = False
50290  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFonts", "ControlsEnabled [LET]")
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
Select Case ErrPtnr.OnError("ctlOptFonts", "ControlEnabled [GET]")
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
50020  Dim i As Long, fi As Long, tStr As String, SMF As Collection, _
  cSystem As clsSystem
50040
50050  dmFraProgFont.Left = 0
50060  dmFraProgFont.Top = 0
50070  UserControl.Height = dmFraProgFont.Height
50080  lblEnableNotice.Top = lblProgfont.Top
50090  lblEnableNotice.Left = lblProgfont.Left
50100
50110  Set cSystem = New clsSystem
50120  Set SMF = cSystem.GetSystemFont(frmMain, Menu)
50130  txtTest.Text = vbNullString
50140  For i = 33 To 255
50150   txtTest.Text = txtTest.Text & Chr$(i)
50160 '  If UnloadForm Then
50170 '   TimerReady = True
50180 '   Exit Sub
50190 '  End If
50200   DoEvents
50210  Next i
50220  With cmbCharset
50230   .Clear
50240   .AddItem "0, Western": .ItemData(.NewIndex) = 0
50250   .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
50260   .AddItem "77, Mac": .ItemData(.NewIndex) = 77
50270   .AddItem "128, Japanese": .ItemData(.NewIndex) = 128
50280   .AddItem "129, Hangeul": .ItemData(.NewIndex) = 129
50290   .AddItem "130, Hangeul (Johab)": .ItemData(.NewIndex) = 130
50300   .AddItem "134, Chinese_GB2312": .ItemData(.NewIndex) = 134
50310   .AddItem "136, Chinese_BIG5": .ItemData(.NewIndex) = 136
50320   .AddItem "161, Greek": .ItemData(.NewIndex) = 161
50330   .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
50340   .AddItem "163, Vietnamese": .ItemData(.NewIndex) = 163
50350   .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
50360   .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
50370   .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
50380   .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
50390   .AddItem "222, Thai": .ItemData(.NewIndex) = 222
50400   .AddItem "238, Central European": .ItemData(.NewIndex) = 238
50410   .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
50420   .Text = 0
50430  End With
50440  With cmbProgramFontsize
50450   .AddItem "8"
50460   .AddItem "9"
50470   .AddItem "10"
50480   .AddItem "11"
50490   .AddItem "12"
50500   .AddItem "14"
50510   .AddItem "16"
50520   .AddItem "18"
50530   .AddItem "20"
50540   .AddItem "22"
50550   .AddItem "24"
50560   .AddItem "26"
50570   .AddItem "28"
50580   .AddItem "36"
50590   .AddItem "48"
50600   .AddItem "72"
50610  End With
50620  cmbProgramFontsize.Text = 8
50630  cmbCharset.Text = cmbCharset.ItemData(0)
50640  cmbCharset.Text = Options.ProgramFontCharset
50650  fi = -1
50660  With cmbFonts
50670   For i = 1 To Screen.FontCount
50680    tStr = Trim$(Screen.Fonts(i))
50690    If LenB(tStr) > 0 Then
50700     cmbFonts.AddItem tStr
50710     If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
50720      fi = i
50730     End If
50740    End If
50750 '   If UnloadForm Then
50760 '    TimerReady = True
50770 '    Exit Sub
50780 '   End If
50790    DoEvents
50800   Next i
50810  End With
50820
50830 ' Form_Resize
50840
50850  cmbProgramFontsize.Width = txtTest.Width - _
  (cmbProgramFontsize.Left - txtTest.Left)
50870
50880  If fi >= 0 Then
50890   cmbFonts.ListIndex = fi
50900   cmbCharset.Text = SMF(1)(2)
50910   cmbProgramFontsize.Text = SMF(1)(1)
50920   txtTest.Font = cmbFonts.List(cmbFonts.ListIndex)
50930   txtTest.Font.Charset = cmbCharset.Text
50940  End If
50950
50960  With cmbCharset
50970   .Top = cmbFonts.Top
50980   .Left = lblProgcharset.Left
50990   .Width = 2295
51000   .SelStart = 0
51010   .SelLength = 0
51020  End With
51030  With cmbProgramFontsize
51040   .Top = cmbFonts.Top
51050   .Left = lblSize.Left
51060   .Width = 765
51070   .SelStart = 0
51080   .SelLength = 0
51090  End With
51100
51110  CorrectCmbCharset
51120
51130  mControlsEnabled = True
51140  SetFrames Options.OptionsDesign
51150
51160  SetFont
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
Select Case ErrPtnr.OnError("ctlOptFonts", "SetFont")
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
50010  With Options1
50020   .ProgramFont = cmbFonts.List(cmbFonts.ListIndex)
50030   If LenB(cmbCharset.Text) > 0 Then
50040    CorrectCmbCharset
50050    .ProgramFontCharset = cmbCharset.Text
50060   End If
50070   If LenB(cmbProgramFontsize.Text) > 0 Then
50080    .ProgramFontSize = cmbProgramFontsize.Text
50090   End If
50100  End With
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
50220  cmbCharset.Text = tCharset
50230
50240  SetTestFontBack = True
50250  SetFontControls UserControl.Controls, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), CLng(cmbProgramFontsize.Text)
50260  SetFontControls frmMain.Controls, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), CLng(cmbProgramFontsize.Text)
50270  SetFontControls frmOptions.Controls, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), CLng(cmbProgramFontsize.Text)
50280
50290  With cmdTest.Font
50300   .Name = tFontname
50310   .Size = tFontSize
50320   .Charset = tFontCharset
50330  End With
50340  With cmdCancelTest
50350   .Font.Name = tFontname
50360   .Font.Size = tFontSize
50370   .Font.Charset = tFontCharset
50380   .Enabled = True
50390  End With
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
50020   cmbCharset.Text = .ProgramFontCharset
50030   SetFontControls frmOptions.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50040   SetFontControls frmMain.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50050   SetFontControls frmOptions.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50060   SetFontControls UserControl.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50070  End With
50080  SetTestFontBack = False
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
50020  If Trim$(cmbProgramFontsize.Text) = vbNullString Then
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
