VERSION 5.00
Begin VB.UserControl ctlOptGhostscript 
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6660
   ScaleHeight     =   3345
   ScaleWidth      =   6660
   ToolboxBitmap   =   "ctlOptGhostscript.ctx":0000
   Begin PDFCreator.dmFrame dmFraProgGhostscript 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   5530
      Caption         =   "Ghostscript"
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
      Begin VB.ComboBox cmbGhostscript 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown-Liste
         TabIndex        =   2
         Top             =   630
         Width           =   4215
      End
      Begin VB.CommandButton cmdGetgsresourceDirectory 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   5625
         TabIndex        =   19
         Top             =   5490
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtGSresource 
         Appearance      =   0  '2D
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   5490
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetgsfontsDirectory 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   5625
         TabIndex        =   16
         Top             =   4890
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdGetgslibDirectory 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   5625
         TabIndex        =   13
         Top             =   4290
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtGSfonts 
         Appearance      =   0  '2D
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   4890
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txtGSlib 
         Appearance      =   0  '2D
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   4290
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txtGSbin 
         Appearance      =   0  '2D
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3690
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetgsbinDirectory 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   5625
         TabIndex        =   9
         Top             =   3690
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAdditionalGhostscriptSearchpath 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   105
         TabIndex        =   6
         Top             =   2100
         Width           =   6105
      End
      Begin VB.CheckBox chkAddWindowsFontpath 
         Appearance      =   0  '2D
         Caption         =   "Add Windows fontpath"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         TabIndex        =   7
         Top             =   2730
         Width           =   6105
      End
      Begin VB.ComboBox cmbAdditionalGhostscriptParameters 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   105
         TabIndex        =   4
         Top             =   1400
         Width           =   6105
      End
      Begin VB.Label lblEnableNotice 
         Caption         =   "You can set these options in the default profile only."
         Enabled         =   0   'False
         Height          =   255
         Left            =   105
         TabIndex        =   20
         Top             =   6000
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.Label lblGhostscriptResource 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Resource"
         Enabled         =   0   'False
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   5250
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lblGSfonts 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Fonts"
         Enabled         =   0   'False
         Height          =   195
         Left            =   105
         TabIndex        =   14
         Top             =   4650
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblGSlib 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Libraries"
         Enabled         =   0   'False
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   4050
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblGSbin 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Binaries"
         Enabled         =   0   'False
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   3450
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblGhostscriptversion 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscriptversion"
         Height          =   195
         Left            =   105
         TabIndex        =   1
         Top             =   420
         Width           =   1305
      End
      Begin VB.Label lblAdditionalGhostscriptParameters 
         AutoSize        =   -1  'True
         Caption         =   "Additional Ghostscript parameters"
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   1155
         Width           =   2355
      End
      Begin VB.Label lblAdditionalGhostscriptSearchpath 
         AutoSize        =   -1  'True
         Caption         =   "Additional Ghostscript searchpath"
         Height          =   195
         Left            =   105
         TabIndex        =   5
         Top             =   1890
         Width           =   2370
      End
   End
End
Attribute VB_Name = "ctlOptGhostscript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mEnabled As Boolean
Private mControlsEnabled As Boolean

Public Sub SetControlsEnabled(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mControlsEnabled = value
50020  ControlsEnabled = value
50030  dmFraProgGhostscript.Enabled = value
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGhostscript", "SetControlsEnabled")
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
50030  lblGhostscriptversion.Enabled = mEnabled
50040  lblGhostscriptversion.Visible = mEnabled
50050  cmbGhostscript.Enabled = mEnabled
50060  cmbGhostscript.Visible = mEnabled
50070  lblAdditionalGhostscriptParameters.Enabled = mEnabled
50080  lblAdditionalGhostscriptParameters.Visible = mEnabled
50090  cmbAdditionalGhostscriptParameters.Enabled = mEnabled
50100  cmbAdditionalGhostscriptParameters.Visible = mEnabled
50110  lblAdditionalGhostscriptSearchpath.Enabled = mEnabled
50120  lblAdditionalGhostscriptSearchpath.Visible = mEnabled
50130  txtAdditionalGhostscriptSearchpath.Enabled = mEnabled
50140  txtAdditionalGhostscriptSearchpath.Visible = mEnabled
50150  chkAddWindowsFontpath.Enabled = mEnabled
50160  chkAddWindowsFontpath.Visible = mEnabled
50170  lblEnableNotice.Visible = Not mEnabled
50180  If mControlsEnabled Then
50190    lblEnableNotice.Enabled = Not mEnabled
50200   Else
50210    lblEnableNotice.Enabled = False
50220  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGhostscript", "ControlsEnabled [LET]")
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
Select Case ErrPtnr.OnError("ctlOptGhostscript", "ControlEnabled [GET]")
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
50020  dmFraProgGhostscript.Left = 0
50030  dmFraProgGhostscript.Top = 0
50040  UserControl.Height = dmFraProgGhostscript.Height
50050
50060  lblEnableNotice.Top = lblGhostscriptversion.Top
50070  lblEnableNotice.Left = lblGhostscriptversion.Left
50080
50090  mControlsEnabled = True
50100  SetFrames Options.OptionsDesign
50110
50120  SetFont
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGhostscript", "UserControl_Initialize")
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
Select Case ErrPtnr.OnError("ctlOptGhostscript", "SetFont")
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
Select Case ErrPtnr.OnError("ctlOptGhostscript", "SetFrames")
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
50010  dmFraProgGhostscript.Width = UserControl.Width
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGhostscript", "UserControl_Resize")
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
50020   dmFraProgGhostscript.Caption = .OptionsProgramGhostscriptSymbol
50030   lblGhostscriptversion.Caption = .OptionsGhostscriptversion
50040   lblAdditionalGhostscriptParameters.Caption = .OptionsAdditionalGhostscriptParameters
50050   lblAdditionalGhostscriptSearchpath.Caption = .OptionsAdditionalGhostscriptSearchpath
50060   chkAddWindowsFontpath.Caption = .OptionsAddWindowsFontpath
50070   lblGSbin.Caption = .OptionsDirectoriesGSBin
50080   lblGSlib.Caption = .OptionsDirectoriesGSLibraries
50090   lblGSfonts.Caption = .OptionsDirectoriesGSFonts
50100   lblEnableNotice.Caption = .OptionsEnableNotice
50110  End With
50120
50130  SetOptimalComboboxHeigth cmbGhostscript, Me
50140  SetOptimalComboboxHeigth cmbAdditionalGhostscriptParameters, Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGhostscript", "SetLanguageStrings")
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
50010  Dim i As Long, tStr As String, tStr2 As String, tsf() As String, reg As clsRegistry, gsvers As Collection
50020
50030  chkAddWindowsFontpath.value = Options1.AddWindowsFontpath
50040
50050  cmbAdditionalGhostscriptParameters.Clear
50060  cmbAdditionalGhostscriptParameters.AddItem "-dTextAlphaBits=4|-dGraphicsAlphaBits=4|-dDOINTERPOLATE"
50070
50080  cmbAdditionalGhostscriptParameters.Text = Options1.AdditionalGhostscriptParameters
50090  txtAdditionalGhostscriptSearchpath.Text = Options1.AdditionalGhostscriptSearchpath
50100
50110  With txtGSbin
50120   .ToolTipText = .Text
50130  End With
50140  With txtGSlib
50150   .ToolTipText = .Text
50160  End With
50170  With txtGSfonts
50180   .ToolTipText = .Text
50190  End With
50200  tStr2 = CompletePath(UCase$(Trim$(Options1.DirectoryGhostscriptBinaries)))
 cmbGhostscript.Clear: Set reg = New clsRegistry
50220  reg.hkey = HKEY_LOCAL_MACHINE
50230
50240  Set gsvers = GetAllGhostscriptversions
50250
50260  If gsvers.Count = 0 Then
50270    cmbGhostscript.Enabled = False
50280   Else
50290    For i = 1 To gsvers.Count
50300     cmbGhostscript.AddItem gsvers.Item(i)
50310    Next i
50320    cmbGhostscript.ListIndex = cmbGhostscript.ListCount - 1
50330    For i = 0 To cmbGhostscript.ListCount - 1
50340     tStr = ""
50350     If InStr(cmbGhostscript.List(i), ":") Then
50360       reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50370       If tStr2 = CompletePath(UCase$(Trim$(reg.GetRegistryValue("GhostscriptDirectoryBinaries")))) Then
50380        cmbGhostscript.ListIndex = i
50390        Exit For
50400       End If
50410      Else
50420       If InStr(UCase$(cmbGhostscript.List(i)), "AFPL") Then
50430        reg.KeyRoot = "SOFTWARE\AFPL Ghostscript"
50440        If InStr(cmbGhostscript.List(i), " ") > 0 Then
50450         tsf = Split(cmbGhostscript.List(i), " ")
50460         reg.SubKey = tsf(UBound(tsf))
50470         tStr = reg.GetRegistryValue("GS_DLL")
50480         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
50490          cmbGhostscript.ListIndex = i
50500          Exit For
50510         End If
50520        End If
50530       End If
50540       If InStr(UCase$(cmbGhostscript.List(i)), "GNU") Then
50550        reg.KeyRoot = "SOFTWARE\GNU Ghostscript"
50560        If InStr(cmbGhostscript.List(i), " ") > 0 Then
50570         tsf = Split(cmbGhostscript.List(i), " ")
50580         reg.SubKey = tsf(UBound(tsf))
50590         tStr = reg.GetRegistryValue("GS_DLL")
50600         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
50610          cmbGhostscript.ListIndex = i
50620          Exit For
50630         End If
50640        End If
50650       End If
50660       If InStr(UCase$(cmbGhostscript.List(i)), "GPL") Then
50670        reg.KeyRoot = "SOFTWARE\GPL Ghostscript"
50680        If InStr(cmbGhostscript.List(i), " ") > 0 Then
50690         tsf = Split(cmbGhostscript.List(i), " ")
50700         reg.SubKey = tsf(UBound(tsf))
50710         tStr = reg.GetRegistryValue("GS_DLL")
50720         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
50730          cmbGhostscript.ListIndex = i
50740          Exit For
50750         End If
50760        End If
50770       End If
50780     End If
50790    Next i
50800  End If
50810  Set reg = Nothing
50820  With cmbGhostscript
50830   If .ListCount = 0 Then
50840    .Enabled = False
50850    .BackColor = &H8000000F
50860   End If
50870  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGhostscript", "SetOptions")
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
50020   .AdditionalGhostscriptParameters = cmbAdditionalGhostscriptParameters.Text
50030   .AdditionalGhostscriptSearchpath = txtAdditionalGhostscriptSearchpath.Text
50040   .AddWindowsFontpath = chkAddWindowsFontpath.value
50050   .DirectoryGhostscriptBinaries = txtGSbin.Text
50060   .DirectoryGhostscriptFonts = txtGSfonts.Text
50070   .DirectoryGhostscriptLibraries = txtGSlib.Text
50080   .DirectoryGhostscriptResource = txtGSresource.Text
50090  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGhostscript", "GetOptions")
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
Select Case ErrPtnr.OnError("ctlOptGhostscript", "ScaleMode [LET]")
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
Select Case ErrPtnr.OnError("ctlOptGhostscript", "ScaleMode [GET]")
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
Select Case ErrPtnr.OnError("ctlOptGhostscript", "hwnd [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Sub cmbGhostscript_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, gsv As String, tsf() As String, Path As String, tStr As String
50020
50030  gsv = cmbGhostscript.List(cmbGhostscript.ListIndex)
50040  Set reg = New clsRegistry: reg.hkey = HKEY_LOCAL_MACHINE
50050
50060  If InStr(gsv, ":") Then
50070    reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50080    txtGSbin.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50090    txtGSfonts.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50100    txtGSlib.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50110    txtGSresource.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryResource"))
50120    Set reg = Nothing
50130    Exit Sub
50140   Else
50150    If InStr(UCase$(gsv), "AFPL") Then
50160     If InStr(gsv, " ") > 0 Then
50170      tsf = Split(gsv, " ")
50180      reg.KeyRoot = "SOFTWARE\AFPL Ghostscript\" & tsf(UBound(tsf))
50190      tStr = reg.GetRegistryValue("GS_DLL")
50200      SplitPath tStr, , Path
50210      txtGSbin.Text = CompletePath(Path)
50220      If InStrRev(Path, "\") > 0 Then
50230       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50240       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50250       If tsf(UBound(tsf)) <> "8.00" Then
50260        txtGSresource.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
50270       End If
50280      End If
50290     End If
50300    End If
50310    If InStr(UCase$(gsv), "GNU") Then
50320     If InStr(gsv, " ") > 0 Then
50330      tsf = Split(gsv, " ")
50340      reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
50350      tStr = reg.GetRegistryValue("GS_DLL")
50360      SplitPath tStr, , Path
50370      txtGSbin.Text = CompletePath(Path)
50380      If InStrRev(Path, "\") > 0 Then
50390       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50400       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50410       txtGSresource.Text = ""
50420      End If
50430     End If
50440    End If
50450    If InStr(UCase$(gsv), "GPL") Then
50460     If InStr(gsv, " ") > 0 Then
50470      tsf = Split(gsv, " ")
50480      reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
50490      tStr = reg.GetRegistryValue("GS_DLL")
50500      SplitPath tStr, , Path
50510      txtGSbin.Text = CompletePath(Path)
50520      If InStrRev(Path, "\") > 0 Then
50530       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50540       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50550       txtGSresource.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
50560      End If
50570     End If
50580    End If
50590  End If
50600  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGhostscript", "cmbGhostscript_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetgsbinDirectory_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strFolder As String, aw As Long
50020  strFolder = BrowseForFolderFiles(UserControl.hwnd, LanguageStrings.OptionsGhostscriptBinariesDirectoryPrompt)
50030  If Len(strFolder) = 0 Then
50040   Exit Sub
50050  End If
50060  strFolder = CompletePath(strFolder)
50070  If FileExists(strFolder & GsDll) = False Then
50080   MsgBox LanguageStrings.MessagesMsg15
50090   Exit Sub
50100  End If
50110  If UCase$(CompletePath(Options.DirectoryGhostscriptBinaries)) <> UCase$(CompletePath(strFolder)) Then
50120   aw = MsgBox("The program must be restarted!", vbOKCancel)
50130   If aw = vbCancel Then
50140    Exit Sub
50150   End If
50160   txtGSbin.Text = strFolder
50170 '  modOptions2.GetOptions UserControl.Parent, Options
50180   SaveOptions Options
50190   Restart = True
50200   Unload Me
50210  End If
50220  With txtGSbin
50230   .Text = strFolder
50240   .ToolTipText = .Text
50250  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGhostscript", "cmdGetgsbinDirectory_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetgsfontsDirectory_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strFolder As String
50020  strFolder = BrowseForFolderFiles(UserControl.hwnd, LanguageStrings.OptionsGhostscriptFontsDirectoryPrompt)
50030  If Len(strFolder) = 0 Then Exit Sub
50040  strFolder = CompletePath(strFolder)
50050  If LenB(Dir(strFolder & "*.afm", vbNormal)) = 0 And LenB(Dir(strFolder & "*.pfb", vbNormal)) = 0 Then
50060   MsgBox LanguageStrings.MessagesMsg16
50070   Exit Sub
50080  End If
50090  txtGSfonts.Text = strFolder
50100  With txtGSfonts
50110   .ToolTipText = .Text
50120  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGhostscript", "cmdGetgsfontsDirectory_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetgslibDirectory_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strFolder As String
50020  strFolder = BrowseForFolderFiles(UserControl.hwnd, LanguageStrings.OptionsGhostscriptLibrariesDirectoryPrompt)
50030  If Len(strFolder) = 0 Then Exit Sub
50040  strFolder = CompletePath(strFolder)
50050  If LenB(Dir(strFolder & "*.*", vbNormal)) = 0 Then
50060   MsgBox LanguageStrings.MessagesMsg17
50070   Exit Sub
50080  End If
50090  With txtGSlib
50100   .Text = strFolder
50110   .ToolTipText = .Text
50120  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGhostscript", "cmdGetgslibDirectory_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetgsresourceDirectory_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strFolder As String
50020  strFolder = BrowseForFolderFiles(UserControl.hwnd, LanguageStrings.OptionsGhostscriptResourceDirectoryPrompt)
50030  If Len(strFolder) = 0 Then Exit Sub
50040  strFolder = CompletePath(strFolder)
50050  With txtGSresource
50060   .Text = strFolder
50070   .ToolTipText = .Text
50080  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGhostscript", "cmdGetgsresourceDirectory_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Property Get GSBin() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GSBin = txtGSbin.Text
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGhostscript", "GSBin [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
