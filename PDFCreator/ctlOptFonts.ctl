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
      _ExtentX        =   11245
      _ExtentY        =   8281
      Caption         =   "Programfont"
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
 Dim ctl As Control
 dmFraProgFont.Left = 0
 dmFraProgFont.Top = 0
 UserControl.Height = dmFraProgFont.Height

 Dim i As Long, fi As Long, tStr As String, SMF As Collection, _
  cSystem As clsSystem
 Set cSystem = New clsSystem
 Set SMF = cSystem.GetSystemFont(frmMain, Menu)
 txtTest.Text = vbNullString
 For i = 33 To 255
  txtTest.Text = txtTest.Text & Chr$(i)
'  If UnloadForm Then
'   TimerReady = True
'   Exit Sub
'  End If
  DoEvents
 Next i
 With cmbCharset
  .Clear
  .AddItem "0, Western": .ItemData(.NewIndex) = 0
  .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
  .AddItem "77, Mac": .ItemData(.NewIndex) = 77
  .AddItem "128, Japanese": .ItemData(.NewIndex) = 128
  .AddItem "129, Hangeul": .ItemData(.NewIndex) = 129
  .AddItem "130, Hangeul (Johab)": .ItemData(.NewIndex) = 130
  .AddItem "134, Chinese_GB2312": .ItemData(.NewIndex) = 134
  .AddItem "136, Chinese_BIG5": .ItemData(.NewIndex) = 136
  .AddItem "161, Greek": .ItemData(.NewIndex) = 161
  .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
  .AddItem "163, Vietnamese": .ItemData(.NewIndex) = 163
  .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
  .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
  .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
  .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
  .AddItem "222, Thai": .ItemData(.NewIndex) = 222
  .AddItem "238, Central European": .ItemData(.NewIndex) = 238
  .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
  .Text = 0
 End With
 With cmbProgramFontsize
  .AddItem "8"
  .AddItem "9"
  .AddItem "10"
  .AddItem "11"
  .AddItem "12"
  .AddItem "14"
  .AddItem "16"
  .AddItem "18"
  .AddItem "20"
  .AddItem "22"
  .AddItem "24"
  .AddItem "26"
  .AddItem "28"
  .AddItem "36"
  .AddItem "48"
  .AddItem "72"
 End With
 cmbProgramFontsize.Text = 8
 cmbCharset.Text = cmbCharset.ItemData(0)
 cmbCharset.Text = Options.ProgramFontCharset
 fi = -1
 With cmbFonts
  For i = 1 To Screen.FontCount
   tStr = Trim$(Screen.Fonts(i))
   If LenB(tStr) > 0 Then
    cmbFonts.AddItem tStr
    If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
     fi = i
    End If
   End If
'   If UnloadForm Then
'    TimerReady = True
'    Exit Sub
'   End If
   DoEvents
  Next i
 End With

' Form_Resize

 cmbProgramFontsize.Width = txtTest.Width - _
  (cmbProgramFontsize.Left - txtTest.Left)

 If fi >= 0 Then
  cmbFonts.ListIndex = fi
  cmbCharset.Text = SMF(1)(2)
  cmbProgramFontsize.Text = SMF(1)(1)
  txtTest.Font = cmbFonts.List(cmbFonts.ListIndex)
  txtTest.Font.Charset = cmbCharset.Text
 End If

 With cmbCharset
  .Top = cmbFonts.Top
  .Left = lblProgcharset.Left
  .Width = 2295
  .SelStart = 0
  .SelLength = 0
 End With
 With cmbProgramFontsize
  .Top = cmbFonts.Top
  .Left = lblSize.Left
  .Width = 765
  .SelStart = 0
  .SelLength = 0
 End With

 CorrectCmbCharset

 SetFrames Options.OptionsDesign
End Sub

Public Sub SetFrames(OptionsDesign As Long)
 Dim ctl As Control
 For Each ctl In UserControl.Controls
  If TypeOf ctl Is dmFrame Then
   SetFrame ctl, OptionsDesign
  End If
 Next ctl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 Dim ctl As Control
 For Each ctl In Controls
'  If UnloadForm Then
'   TimerReady = True
'   Exit Sub
'  End If
  DoEvents
  If TypeOf ctl Is ComboBox Then
   ComboSetListWidth ctl
  End If
 Next ctl

 SetOptimalComboboxHeigth cmbCharset, Me
 SetOptimalComboboxHeigth cmbProgramFontsize, Me
End Sub

Private Sub UserControl_Resize()
 dmFraProgFont.Width = UserControl.Width
End Sub

Public Sub SetLanguageStrings()
 With LanguageStrings
  dmFraProgFont.Caption = .OptionsProgramFontSymbol
  lblProgfont.Caption = .OptionsProgramFont
  lblProgcharset.Caption = .OptionsProgramFontcharset
  lblSize.Caption = .OptionsProgramFontSize
  lblTesttext = .OptionsProgramFontTestdescription
  cmdTest.Caption = .OptionsProgramFontTest
  cmdCancelTest.Caption = .OptionsProgramFontCancelTest
 End With
End Sub

Public Sub SetOptions()
 Dim i As Long
 With Options
  For i = 0 To cmbFonts.ListCount - 1
    If UCase$(cmbFonts.List(i)) = UCase$(.ProgramFont) Then
     cmbFonts.ListIndex = i
     Exit For
    End If
  Next i
  cmbCharset.Text = .ProgramFontCharset
  cmbProgramFontsize.Text = .ProgramFontSize
 End With
End Sub

Public Sub GetOptions()
 With Options
  .ProgramFont = cmbFonts.List(cmbFonts.ListIndex)
  If LenB(cmbCharset.Text) > 0 Then
   .ProgramFontCharset = cmbCharset.Text
  End If
  If LenB(cmbProgramFontsize.Text) > 0 Then
   .ProgramFontSize = cmbProgramFontsize.Text
  End If
 End With
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
 With cmbCharset
  .Text = .ItemData(.ListIndex)
 End With
 txtTest.Font.Charset = cmbCharset.Text
End Sub

Private Sub cmbCharset_KeyPress(KeyAscii As Integer)
 KeyAscii = AllowedKeypressChars(KeyAscii)
End Sub

Private Sub cmbCharset_Validate(Cancel As Boolean)
 Dim i As Long, tStr As String
 tStr = ""
 For i = 1 To Len(cmbCharset.Text)
  If InStr("0123456789", Mid(cmbCharset.Text, i, 1)) > 0 Then
    tStr = tStr & Mid(cmbCharset.Text, i, 1)
   Else
    Exit For
  End If
 Next i
 If Len(Trim$(tStr)) = 0 Then
   cmbCharset.Text = 0
  Else
   cmbCharset.Text = tStr
 End If
End Sub

Private Sub cmdTest_Click()
 Dim tCharset As Long, tStr As String, tFontSize As Long, tFontname As String, _
  tFontCharset As Long
 If InStr(1, cmbCharset.Text, ",", vbTextCompare) > 0 Then
   tStr = Trim$(Mid$(cmbCharset.Text, 1, InStr(1, cmbCharset.Text, ",", vbTextCompare) - 1))
  Else
   tStr = Trim$(cmbCharset.Text)
 End If
 If Len(tStr) = 0 Then
  cmbCharset.Text = 0
  Exit Sub
 End If
 If IsNumeric(tStr) = False Then
  cmbCharset.Text = 0
  Exit Sub
 End If
 tCharset = tStr
 With cmdTest.Font
  tFontname = .Name
  tFontSize = .Size
  tFontCharset = .Charset
 End With
 SetFontUserControl cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), CLng(cmbProgramFontsize.Text)
 cmbCharset.Text = tCharset
 SetFont frmMain, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), cmbProgramFontsize.Text
' ieb.Refresh
 With cmdTest.Font
  .Name = tFontname
  .Size = tFontSize
  .Charset = tFontCharset
 End With
 With cmdCancelTest
  .Font.Name = tFontname
  .Font.Size = tFontSize
  .Font.Charset = tFontCharset
  .Enabled = True
 End With
End Sub

Private Sub cmdCancelTest_Click()
 With Options
  SetFont frmOptions, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
  cmbCharset.Text = .ProgramFontCharset
  SetFont frmMain, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
 End With
End Sub

Private Sub cmbFonts_Click()
 txtTest.Font.Name = cmbFonts.List(cmbFonts.ListIndex)
End Sub

Public Sub CorrectCmbCharset()
 Dim tStrf() As String
 If InStr(1, cmbCharset.Text, ",", vbTextCompare) > 0 Then
   tStrf = Split(cmbCharset.Text, ",")
   If Len(tStrf(0)) = 0 Then
     cmbCharset.Text = 0
    Else
     If IsNumeric(tStrf(0)) = False Then
       cmbCharset.Text = 0
      Else
       cmbCharset.Text = CLng(tStrf(0))
     End If
   End If
  Else
   If Len(cmbCharset.Text) = 0 Then
     cmbCharset.Text = 0
    Else
     If IsNumeric(cmbCharset.Text) = False Then
       cmbCharset.Text = 0
      Else
       cmbCharset.Text = CLng(cmbCharset.Text)
     End If
   End If
 End If
End Sub


Private Sub cmbProgramFontSize_Change()
 Dim tL As Long
If Trim$(cmbProgramFontsize.Text) = vbNullString Then
  cmbProgramFontsize.Text = 8
 End If
 tL = CLng(cmbProgramFontsize.Text)
 If tL <= 0 Then
  tL = 1
 End If
 If tL > 72 Then
  tL = 72
 End If
 cmbProgramFontsize.Text = tL
 txtTest.Font.Size = tL
End Sub

Private Sub cmbProgramFontSize_KeyPress(KeyAscii As Integer)
 KeyAscii = AllowedKeypressChars(KeyAscii)
End Sub

Private Sub cmbProgramFontsize_Click()
 Dim tL As Long
If Trim$(cmbProgramFontsize.Text) = vbNullString Then
  cmbProgramFontsize.Text = 8
 End If
 tL = CLng(cmbProgramFontsize.Text)
 If tL <= 0 Then
  tL = 1
 End If
 If tL > 72 Then
  tL = 72
 End If
 cmbProgramFontsize.Text = tL
 txtTest.Font.Size = tL
End Sub

Private Sub ComboSetListWidth(oCombo As ComboBox, Optional ByVal nFixWidth As Variant, Optional ByVal nScaleMode As Variant)
 With oCombo
  If IsMissing(nScaleMode) Or IsMissing(nFixWidth) Then
   nScaleMode = .Parent.ScaleMode
  End If
  If IsMissing(nFixWidth) Then
   Dim i As Long, nWidth As Long
   nFixWidth = 0
   For i = 0 To .ListCount - 1
    nWidth = .Parent.TextWidth(.List(i))
    If nWidth > nFixWidth Then
     nFixWidth = nWidth
    End If
   Next i
   nFixWidth = nFixWidth + .Parent.ScaleX(10, vbPixels, nScaleMode)
   If .ListCount > 8 Then
    nFixWidth = nFixWidth + .Parent.ScaleX(15, vbPixels, nScaleMode)
   End If
  End If
  SendMessage .hwnd, CB_SETDROPPEDWIDTH, .Parent.ScaleX(nFixWidth, nScaleMode, vbPixels), 0&
 End With
End Sub

Public Sub SetFontUserControl(ByVal Fontname As String, ByVal Charset As Long, ByVal Fontsize As Long)
 Dim ctl As Control, ts As TabStrip, df As dmFrame, f As StdFont, trv As TreeView

 If LenB(Trim$(Fontname)) = 0 Then
  Exit Sub
 End If

 Set f = New StdFont
 f.Name = Fontname
 f.Size = Fontsize
 f.Charset = Charset

 For Each ctl In UserControl.Controls
  If TypeOf ctl Is Label Or _
     TypeOf ctl Is Form Or _
     TypeOf ctl Is ComboBox Or _
     TypeOf ctl Is CheckBox Or _
     TypeOf ctl Is CommandButton Or _
     TypeOf ctl Is ListView Or _
     TypeOf ctl Is StatusBar Or _
     TypeOf ctl Is TextBox Or _
     TypeOf ctl Is Frame Then
   With ctl
    .Font = Fontname
    If Not (TypeOf ctl Is StatusBar) And Not (TypeOf ctl Is ListView) Then
     .Fontsize = Fontsize
    End If
    .Font.Charset = Charset
   End With
  End If

  If TypeOf ctl Is TreeView Then
   Set trv = ctl
   trv.Font.Name = Fontname
   trv.Font.Size = Fontsize
   trv.Font.Charset = Charset
  End If
  If TypeOf ctl Is TabStrip Then
   Set ts = ctl
   ts.Font.Name = Fontname
   ts.Font.Size = Fontsize
   ts.Font.Charset = Charset
  End If
  If TypeOf ctl Is dmFrame Then
   Set df = ctl
   df.Font.Name = Fontname
   df.Font.Size = Fontsize
   df.Font.Charset = Charset
   Set df.Font = f
  End If
 Next ctl

 Set f = Nothing
End Sub
