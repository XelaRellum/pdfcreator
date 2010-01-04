VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlOptActions 
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ScaleHeight     =   4950
   ScaleWidth      =   6765
   ToolboxBitmap   =   "ctlOptActions.ctx":0000
   Begin PDFCreator.dmFrame dmFraProgActions 
      Height          =   4440
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   7832
      Caption         =   "Actions"
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
      Begin PDFCreator.dmFrame dmFraProgActionsRunProgramAfterSaving 
         Height          =   3510
         Left            =   1785
         TabIndex        =   1
         Top             =   735
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6191
         Caption         =   "Run a program/script after saving"
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
         Begin VB.CheckBox chkRunProgramAfterSaving 
            Appearance      =   0  '2D
            Caption         =   "Run a program/script after saving"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   8
            Top             =   420
            Width           =   5805
         End
         Begin VB.ComboBox cmbRunProgramAfterSavingProgramname 
            Height          =   315
            Left            =   210
            TabIndex        =   7
            Top             =   1155
            Width           =   4770
         End
         Begin VB.CommandButton cmdRunProgramAfterSavingPrognameChoice 
            Caption         =   "..."
            Height          =   300
            Left            =   5040
            TabIndex        =   6
            Top             =   1155
            Width           =   435
         End
         Begin VB.TextBox txtRunProgramAfterSavingProgramParameters 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   210
            TabIndex        =   5
            Top             =   1890
            Width           =   5805
         End
         Begin VB.CheckBox chkRunProgramAfterSavingWaitUntilReady 
            Appearance      =   0  '2D
            Caption         =   "Wait until ready"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   4
            Top             =   3150
            Width           =   5805
         End
         Begin VB.ComboBox cmbRunProgramAfterSavingWindowstyle 
            Height          =   315
            Left            =   210
            Style           =   2  'Dropdown-Liste
            TabIndex        =   3
            Top             =   2625
            Width           =   5370
         End
         Begin VB.CommandButton cmdRunProgramAfterSavingPrognameEdit 
            Height          =   300
            Left            =   5520
            Picture         =   "ctlOptActions.ctx":0312
            Style           =   1  'Grafisch
            TabIndex        =   2
            Top             =   1155
            Width           =   435
         End
         Begin VB.Label lblRunProgramAfterSavingProgramname 
            AutoSize        =   -1  'True
            Caption         =   "Program/Script"
            Height          =   195
            Left            =   210
            TabIndex        =   11
            Top             =   945
            Width           =   1065
         End
         Begin VB.Label lblRunProgramAfterSavingProgramParameters 
            AutoSize        =   -1  'True
            Caption         =   "Program parameters"
            Height          =   195
            Left            =   210
            TabIndex        =   10
            Top             =   1680
            Width           =   1410
         End
         Begin VB.Label lblRunProgramAfterSavingWindowstyle 
            AutoSize        =   -1  'True
            Caption         =   "Windowstyle"
            Height          =   195
            Left            =   210
            TabIndex        =   9
            Top             =   2415
            Width           =   900
         End
      End
      Begin PDFCreator.dmFrame dmFraProgActionsRunProgramBeforeSaving 
         Height          =   3510
         Left            =   210
         TabIndex        =   12
         Top             =   735
         Visible         =   0   'False
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6191
         Caption         =   "Run a program/script before saving"
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
         Begin VB.CheckBox chkRunProgramBeforeSaving 
            Appearance      =   0  '2D
            Caption         =   "Run a program/script before saving"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   18
            Top             =   420
            Width           =   5385
         End
         Begin VB.ComboBox cmbRunProgramBeforeSavingProgramname 
            Height          =   315
            Left            =   210
            TabIndex        =   17
            Top             =   1155
            Width           =   4770
         End
         Begin VB.CommandButton cmdRunProgramBeforeSavingPrognameChoice 
            Caption         =   "..."
            Height          =   300
            Left            =   5040
            TabIndex        =   16
            Top             =   1155
            Width           =   435
         End
         Begin VB.TextBox txtRunProgramBeforeSavingProgramParameters 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   210
            TabIndex        =   15
            Top             =   1890
            Width           =   5580
         End
         Begin VB.ComboBox cmbRunProgramBeforeSavingWindowstyle 
            Height          =   315
            Left            =   210
            Style           =   2  'Dropdown-Liste
            TabIndex        =   14
            Top             =   2625
            Width           =   2790
         End
         Begin VB.CommandButton cmdRunProgramBeforeSavingPrognameEdit 
            Height          =   300
            Left            =   5520
            Picture         =   "ctlOptActions.ctx":089C
            Style           =   1  'Grafisch
            TabIndex        =   13
            Top             =   1155
            Width           =   435
         End
         Begin VB.Label lblRunProgramBeforeSavingProgramname 
            AutoSize        =   -1  'True
            Caption         =   "Program/Script"
            Height          =   195
            Left            =   210
            TabIndex        =   21
            Top             =   945
            Width           =   1065
         End
         Begin VB.Label lblRunProgramBeforeSavingProgramParameters 
            AutoSize        =   -1  'True
            Caption         =   "Program parameters"
            Height          =   195
            Left            =   210
            TabIndex        =   20
            Top             =   1680
            Width           =   1410
         End
         Begin VB.Label lblRunProgramBeforeSavingWindowstyle 
            AutoSize        =   -1  'True
            Caption         =   "Windowstyle"
            Height          =   195
            Left            =   210
            TabIndex        =   19
            Top             =   2415
            Width           =   900
         End
      End
      Begin MSComctlLib.TabStrip tbstrProgActions 
         Height          =   3975
         Left            =   60
         TabIndex        =   22
         Top             =   360
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   7011
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "ctlOptActions"
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
50020  Dim i As Long, files As Collection, tsf() As String, Path As String, filename As String, Ext As String
50030  dmFraProgActions.Left = 0
50040  dmFraProgActions.Top = 0
50050  dmFraProgActionsRunProgramBeforeSaving.Left = tbstrProgActions.ClientLeft
50060  dmFraProgActionsRunProgramBeforeSaving.Top = tbstrProgActions.ClientTop + 30
50070  dmFraProgActionsRunProgramAfterSaving.Left = tbstrProgActions.ClientLeft
50080  dmFraProgActionsRunProgramAfterSaving.Top = tbstrProgActions.ClientTop + 30
50090
50100  UserControl.Height = dmFraProgActions.Height
50110  With cmbRunProgramBeforeSavingWindowstyle
50120   .Clear
50130   For i = 0 To 5
50140    .AddItem ""
50150   Next i
50160  End With
50170  With cmbRunProgramAfterSavingWindowstyle
50180   .Clear
50190   For i = 0 To 5
50200    .AddItem ""
50210   Next i
50220  End With
50230  With tbstrProgActions.Tabs
50240   .Clear
50250   .Add
50260   .Add
50270  End With
50280  tbstrProgActions.ZOrder 1
50290  tbstrProgActions.Tabs(2).Selected = True
50300  If Options.RunProgramAfterSaving Then
50310    ViewRunProgramAfterSaving True
50320   Else
50330    ViewRunProgramAfterSaving False
50340  End If
50350  If Options.RunProgramBeforeSaving Then
50360    ViewRunProgramBeforeSaving True
50370   Else
50380    ViewRunProgramBeforeSaving False
50390  End If
50400
50410  Set files = GetFiles(GetPDFCreatorApplicationPath & "Scripts\RunProgramAfterSaving\", "*.*", SortedByName)
50420  For i = 1 To files.Count
50430   tsf = Split(files(i), "|")
50440   SplitPath tsf(1), , Path, filename, , Ext
50450   If UCase$(Ext) <> "TXT" And UCase$(Ext) <> "PDF" And UCase$(Ext) <> "PNG" And _
   UCase$(Ext) <> "JPG" And UCase$(Ext) <> "BMP" And UCase$(Ext) <> "PCX" And _
   UCase$(Ext) <> "TIF" And UCase$(Ext) <> "EPS" And UCase$(Ext) <> "PS" Then
50480    If UCase$(tsf(0)) <> UCase$(GetPDFCreatorApplicationPath & "Scripts\RunProgramAfterSaving\") Then
50490      cmbRunProgramAfterSavingProgramname.AddItem tsf(0)
50500     Else
50510      cmbRunProgramAfterSavingProgramname.AddItem filename
50520    End If
50530   End If
50540  Next i
50550
50560  Set files = GetFiles(GetPDFCreatorApplicationPath & "Scripts\RunProgramBeforeSaving\", "*.*", SortedByName)
50570  For i = 1 To files.Count
50580   tsf = Split(files(i), "|")
50590   SplitPath tsf(1), , Path, filename, , Ext
50600   If UCase$(Ext) <> "TXT" And UCase$(Ext) <> "PDF" And UCase$(Ext) <> "PNG" And _
   UCase$(Ext) <> "JPG" And UCase$(Ext) <> "BMP" And UCase$(Ext) <> "PCX" And _
   UCase$(Ext) <> "TIF" And UCase$(Ext) <> "EPS" And UCase$(Ext) <> "PS" Then
50630    If UCase$(tsf(0)) <> UCase$(GetPDFCreatorApplicationPath & "Scripts\RunProgramBeforeSaving\") Then
50640      cmbRunProgramBeforeSavingProgramname.AddItem tsf(0)
50650     Else
50660      cmbRunProgramBeforeSavingProgramname.AddItem filename
50670    End If
50680   End If
50690  Next i
50700
50710  SetFrames Options.OptionsDesign
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "UserControl_Initialize")
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
Select Case ErrPtnr.OnError("ctlOptActions", "SetFrames")
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
50010  dmFraProgActions.Width = UserControl.Width
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "UserControl_Resize")
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
50020   dmFraProgActions.Caption = .OptionsProgramActionsSymbol
50030
50040   tbstrProgActions.Tabs(1).Caption = LanguageStrings.OptionsProgramRunProgramBeforeSavingCaption
50050   tbstrProgActions.Tabs(2).Caption = LanguageStrings.OptionsProgramRunProgramAfterSavingCaption
50060
50070   dmFraProgActionsRunProgramAfterSaving.Caption = .OptionsProgramRunProgramAfterSavingCaption
50080   dmFraProgActionsRunProgramBeforeSaving.Caption = .OptionsProgramRunProgramBeforeSavingCaption
50090   chkRunProgramBeforeSaving.Caption = .OptionsProgramRunProgramBeforeSavingCaption
50100   lblRunProgramBeforeSavingProgramname.Caption = .OptionsProgramRunProgramBeforeSavingProgram
50110   lblRunProgramBeforeSavingProgramParameters.Caption = .OptionsProgramRunProgramBeforeSavingProgramParameters
50120   lblRunProgramBeforeSavingWindowstyle.Caption = .OptionsProgramRunProgramBeforeSavingWindowstyle
50130   cmbRunProgramBeforeSavingWindowstyle.List(0) = .OptionsProgramRunProgramBeforeSavingWindowstyleHide
50140   cmbRunProgramBeforeSavingWindowstyle.List(1) = .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus
50150   cmbRunProgramBeforeSavingWindowstyle.List(2) = .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus
50160   cmbRunProgramBeforeSavingWindowstyle.List(3) = .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus
50170   cmbRunProgramBeforeSavingWindowstyle.List(4) = .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus
50180   cmbRunProgramBeforeSavingWindowstyle.List(5) = .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus
50190   chkRunProgramAfterSaving.Caption = .OptionsProgramRunProgramAfterSavingCaption
50200   lblRunProgramAfterSavingProgramname.Caption = .OptionsProgramRunProgramAfterSavingProgram
50210   lblRunProgramAfterSavingProgramParameters.Caption = .OptionsProgramRunProgramAfterSavingProgramParameters
50220   chkRunProgramAfterSavingWaitUntilReady.Caption = .OptionsProgramRunProgramAfterSavingWaitUntilReady
50230   lblRunProgramAfterSavingWindowstyle.Caption = .OptionsProgramRunProgramAfterSavingWindowstyle
50240   cmbRunProgramAfterSavingWindowstyle.List(0) = .OptionsProgramRunProgramAfterSavingWindowstyleHide
50250   cmbRunProgramAfterSavingWindowstyle.List(1) = .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus
50260   cmbRunProgramAfterSavingWindowstyle.List(2) = .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus
50270   cmbRunProgramAfterSavingWindowstyle.List(3) = .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus
50280   cmbRunProgramAfterSavingWindowstyle.List(4) = .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus
50290   cmbRunProgramAfterSavingWindowstyle.List(5) = .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus
50300  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "SetLanguageStrings")
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
50020   chkRunProgramBeforeSaving.value = .RunProgramBeforeSaving
50030   cmbRunProgramBeforeSavingProgramname.Text = .RunProgramBeforeSavingProgramname
50040   txtRunProgramBeforeSavingProgramParameters.Text = .RunProgramBeforeSavingProgramParameters
50050   cmbRunProgramBeforeSavingWindowstyle.ListIndex = .RunProgramBeforeSavingWindowstyle
50060
50070   chkRunProgramAfterSaving.value = .RunProgramAfterSaving
50080   cmbRunProgramAfterSavingProgramname.Text = .RunProgramAfterSavingProgramname
50090   txtRunProgramAfterSavingProgramParameters.Text = .RunProgramAfterSavingProgramParameters
50100   chkRunProgramAfterSavingWaitUntilReady.value = .RunProgramAfterSavingWaitUntilReady
50110   cmbRunProgramAfterSavingWindowstyle.ListIndex = .RunProgramAfterSavingWindowstyle
50120  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "SetOptions")
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
50020   .RunProgramAfterSaving = Abs(chkRunProgramAfterSaving.value)
50030   .RunProgramAfterSavingProgramname = cmbRunProgramAfterSavingProgramname.Text
50040   .RunProgramAfterSavingProgramParameters = txtRunProgramAfterSavingProgramParameters.Text
50050   .RunProgramAfterSavingWaitUntilReady = Abs(chkRunProgramAfterSavingWaitUntilReady.value)
50060   If LenB(CStr(cmbRunProgramAfterSavingWindowstyle.ListIndex)) > 0 Then
50070    .RunProgramAfterSavingWindowstyle = cmbRunProgramAfterSavingWindowstyle.ListIndex
50080   End If
50090   .RunProgramBeforeSaving = Abs(chkRunProgramBeforeSaving.value)
50100   .RunProgramBeforeSavingProgramname = cmbRunProgramBeforeSavingProgramname.Text
50110   .RunProgramBeforeSavingProgramParameters = txtRunProgramBeforeSavingProgramParameters.Text
50120   If LenB(CStr(cmbRunProgramBeforeSavingWindowstyle.ListIndex)) > 0 Then
50130    .RunProgramBeforeSavingWindowstyle = cmbRunProgramBeforeSavingWindowstyle.ListIndex
50140   End If
50150  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "GetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub tbstrProgActions_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ViewProgActions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "tbstrProgActions_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkRunProgramBeforeSaving_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkRunProgramBeforeSaving.value = 1 Then
50020    ViewRunProgramBeforeSaving True
50030   Else
50040    ViewRunProgramBeforeSaving False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "chkRunProgramBeforeSaving_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkRunProgramAfterSaving_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkRunProgramAfterSaving.value = 1 Then
50020    ViewRunProgramAfterSaving True
50030   Else
50040    ViewRunProgramAfterSaving False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "chkRunProgramAfterSaving_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewProgActions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case tbstrProgActions.SelectedItem.Index
        Case 1
50030    dmFraProgActionsRunProgramBeforeSaving.Visible = True
50040    dmFraProgActionsRunProgramBeforeSaving.Enabled = True
50050    dmFraProgActionsRunProgramAfterSaving.Visible = False
50060    dmFraProgActionsRunProgramAfterSaving.Enabled = False
50070   Case 2
50080    dmFraProgActionsRunProgramAfterSaving.Visible = True
50090    dmFraProgActionsRunProgramAfterSaving.Enabled = True
50100    dmFraProgActionsRunProgramBeforeSaving.Visible = False
50110    dmFraProgActionsRunProgramBeforeSaving.Enabled = False
50120  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "ViewProgActions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewRunProgramAfterSaving(Viewit As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblRunProgramAfterSavingProgramname.Enabled = Viewit
50020  cmbRunProgramAfterSavingProgramname.Enabled = Viewit
50030  lblRunProgramAfterSavingProgramParameters.Enabled = Viewit
50040  txtRunProgramAfterSavingProgramParameters.Enabled = Viewit
50050  chkRunProgramAfterSavingWaitUntilReady.Enabled = Viewit
50060  lblRunProgramAfterSavingWindowstyle.Enabled = Viewit
50070  cmbRunProgramAfterSavingWindowstyle.Enabled = Viewit
50080  cmdRunProgramAfterSavingPrognameChoice.Enabled = Viewit
50090  cmdRunProgramAfterSavingPrognameEdit.Enabled = Viewit
50100
50110  If Viewit Then
50120    cmbRunProgramAfterSavingProgramname.BackColor = &H80000005
50130    cmbRunProgramAfterSavingWindowstyle.BackColor = &H80000005
50140    txtRunProgramAfterSavingProgramParameters.BackColor = &H80000005
50150   Else
50160    cmbRunProgramAfterSavingProgramname.BackColor = &H8000000F
50170    cmbRunProgramAfterSavingWindowstyle.BackColor = &H8000000F
50180    txtRunProgramAfterSavingProgramParameters.BackColor = &H8000000F
50190  End If
50200
50210  cmbRunProgramAfterSavingProgramname_Change
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "ViewRunProgramAfterSaving")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewRunProgramBeforeSaving(Viewit As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblRunProgramBeforeSavingProgramname.Enabled = Viewit
50020  cmbRunProgramBeforeSavingProgramname.Enabled = Viewit
50030  lblRunProgramBeforeSavingProgramParameters.Enabled = Viewit
50040  txtRunProgramBeforeSavingProgramParameters.Enabled = Viewit
50050  lblRunProgramBeforeSavingWindowstyle.Enabled = Viewit
50060  cmbRunProgramBeforeSavingWindowstyle.Enabled = Viewit
50070  cmdRunProgramBeforeSavingPrognameChoice.Enabled = Viewit
50080  cmdRunProgramBeforeSavingPrognameEdit.Enabled = Viewit
50090
50100  If Viewit Then
50110    cmbRunProgramBeforeSavingProgramname.BackColor = &H80000005
50120    cmbRunProgramBeforeSavingWindowstyle.BackColor = &H80000005
50130    txtRunProgramBeforeSavingProgramParameters.BackColor = &H80000005
50140   Else
50150    cmbRunProgramBeforeSavingProgramname.BackColor = &H8000000F
50160    cmbRunProgramBeforeSavingWindowstyle.BackColor = &H8000000F
50170    txtRunProgramBeforeSavingProgramParameters.BackColor = &H8000000F
50180  End If
50190
50200  cmbRunProgramBeforeSavingProgramname_Change
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "ViewRunProgramBeforeSaving")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdRunProgramAfterSavingPrognameChoice_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim filename As String
50020  filename = BrowseForFolderFiles(UserControl.Parent.hwnd, LanguageStrings.OptionsProgramRunProgramAfterSavingCaption, False)
50030  If LenB(filename) > 0 Then
50040   cmbRunProgramAfterSavingProgramname.Text = filename
50050  End If
50060  If FileExists(filename) = True Then
50070    If IsFileEditable(filename) Then
50080      cmdRunProgramAfterSavingPrognameEdit.Enabled = True
50090     Else
50100      cmdRunProgramAfterSavingPrognameEdit.Enabled = False
50110    End If
50120   Else
50130    cmdRunProgramAfterSavingPrognameEdit.Enabled = False
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "cmdRunProgramAfterSavingPrognameChoice_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdRunProgramAfterSavingPrognameEdit_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Program As String, drv As String
50020  Program = RemoveLeadingAndTrailingQuotes(cmbRunProgramAfterSavingProgramname.Text)
50030  SplitPath Program, drv
50040  If LenB(drv) = 0 Then
50050   Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
50060  End If
50070  If FileExists(Program) = True Then
50080   If IsFileEditable(Program) Then
50090    EditDocument Program
50100   End If
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "cmdRunProgramAfterSavingPrognameEdit_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdRunProgramBeforeSavingPrognameChoice_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim filename As String
50020  filename = BrowseForFolderFiles(UserControl.Parent.hwnd, LanguageStrings.OptionsProgramRunProgramBeforeSavingCaption, False)
50030  If LenB(filename) > 0 Then
50040   cmbRunProgramBeforeSavingProgramname.Text = filename
50050  End If
50060  If FileExists(filename) = True Then
50070    If IsFileEditable(filename) Then
50080      cmdRunProgramBeforeSavingPrognameEdit.Enabled = True
50090     Else
50100      cmdRunProgramBeforeSavingPrognameEdit.Enabled = False
50110    End If
50120   Else
50130    cmdRunProgramBeforeSavingPrognameEdit.Enabled = False
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "cmdRunProgramBeforeSavingPrognameChoice_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdRunProgramBeforeSavingPrognameEdit_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Program As String, drv As String
50020  Program = RemoveLeadingAndTrailingQuotes(cmbRunProgramBeforeSavingProgramname.Text)
50030  SplitPath Program, drv
50040  If LenB(drv) = 0 Then
50050   Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
50060  End If
50070  If FileExists(Program) = True Then
50080   If IsFileEditable(Program) Then
50090    EditDocument Program
50100   End If
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "cmdRunProgramBeforeSavingPrognameEdit_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbRunProgramAfterSavingProgramname_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Program As String, drv As String
50020  Program = RemoveLeadingAndTrailingQuotes(cmbRunProgramAfterSavingProgramname.Text)
50030  SplitPath Program, drv
50040  If LenB(drv) = 0 Then
50050   Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
50060  End If
50070  If FileExists(Program) = True Then
50080    If IsFileEditable(Program) Then
50090      cmdRunProgramAfterSavingPrognameEdit.Enabled = True
50100     Else
50110      cmdRunProgramAfterSavingPrognameEdit.Enabled = False
50120    End If
50130   Else
50140    cmdRunProgramAfterSavingPrognameEdit.Enabled = False
50150  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "cmbRunProgramAfterSavingProgramname_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbRunProgramAfterSavingProgramname_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With cmbRunProgramAfterSavingProgramname
50020   If .ListCount > 0 Then
50030    .Text = "Scripts\RunProgramAfterSaving\" & .List(.ListIndex)
50040   End If
50050  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "cmbRunProgramAfterSavingProgramname_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbRunProgramBeforeSavingProgramname_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Program As String, drv As String
50020  Program = RemoveLeadingAndTrailingQuotes(cmbRunProgramBeforeSavingProgramname.Text)
50030  SplitPath Program, drv
50040  If LenB(drv) = 0 Then
50050   Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
50060  End If
50070  If FileExists(Program) = True Then
50080    If IsFileEditable(Program) Then
50090      cmdRunProgramBeforeSavingPrognameEdit.Enabled = True
50100     Else
50110      cmdRunProgramBeforeSavingPrognameEdit.Enabled = False
50120    End If
50130   Else
50140    cmdRunProgramBeforeSavingPrognameEdit.Enabled = False
50150  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "cmbRunProgramBeforeSavingProgramname_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbRunProgramBeforeSavingProgramname_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With cmbRunProgramBeforeSavingProgramname
50020   If .ListCount > 0 Then
50030    .Text = "Scripts\RunProgramBeforeSaving\" & .List(.ListIndex)
50040   End If
50050  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptActions", "cmbRunProgramBeforeSavingProgramname_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
