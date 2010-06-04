VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlOptSave 
   ClientHeight    =   5865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   ScaleHeight     =   5865
   ScaleWidth      =   6570
   ToolboxBitmap   =   "ctlOptSave.ctx":0000
   Begin PDFCreator.dmFrame dmFraProgSave 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5318
      Caption         =   "Save"
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
      Begin VB.CheckBox chkAllowSpecialGSCharsInFilenames 
         Appearance      =   0  '2D
         Caption         =   "Allow special Ghostscript chars in filenames"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Value           =   1  'Aktiviert
         Width           =   6015
      End
      Begin VB.TextBox txtSavePreview 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   6015
      End
      Begin VB.ComboBox cmbSaveFilenameTokens 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "ctlOptSave.ctx":0312
         Left            =   3720
         List            =   "ctlOptSave.ctx":0314
         Style           =   2  'Dropdown-Liste
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtSaveFilename 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "<Title>"
         Top             =   600
         Width           =   3495
      End
      Begin VB.CheckBox chkSpaces 
         Appearance      =   0  '2D
         Caption         =   "Remove leading and trailing spaces"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Value           =   1  'Aktiviert
         Width           =   6015
      End
      Begin VB.ComboBox cmbStandardSaveFormat 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "ctlOptSave.ctx":0316
         Left            =   120
         List            =   "ctlOptSave.ctx":0318
         Style           =   2  'Dropdown-Liste
         TabIndex        =   8
         Top             =   2100
         Width           =   1935
      End
      Begin VB.Label lblSaveFilenameTokens 
         AutoSize        =   -1  'True
         Caption         =   "Add a Filename-Token"
         Height          =   195
         Left            =   3720
         TabIndex        =   2
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label lblSaveFilename 
         AutoSize        =   -1  'True
         Caption         =   "Filename"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblStandardSaveformat 
         AutoSize        =   -1  'True
         Caption         =   "Standard save format"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1890
         Width           =   1515
      End
   End
   Begin PDFCreator.dmFrame dmFraFilenameSubstitutions 
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4471
      Caption         =   "Filename substitutions"
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
      Begin VB.CommandButton cmdFilenameSubst 
         Caption         =   "&Delete"
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   18
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdFilenameSubst 
         Caption         =   "C&hange"
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   17
         Top             =   1155
         Width           =   1455
      End
      Begin VB.CommandButton cmdFilenameSubstMove 
         Height          =   435
         Index           =   1
         Left            =   120
         Picture         =   "ctlOptSave.ctx":031A
         Style           =   1  'Grafisch
         TabIndex        =   14
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtFilenameSubst 
         Appearance      =   0  '2D
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkFilenameSubst 
         Appearance      =   0  '2D
         Caption         =   "Substitutions only in <Title>"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Value           =   1  'Aktiviert
         Width           =   3255
      End
      Begin VB.TextBox txtFilenameSubst 
         Appearance      =   0  '2D
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdFilenameSubstMove 
         Height          =   435
         Index           =   0
         Left            =   120
         Picture         =   "ctlOptSave.ctx":06A4
         Style           =   1  'Grafisch
         TabIndex        =   13
         Top             =   915
         Width           =   375
      End
      Begin VB.CommandButton cmdFilenameSubst 
         Caption         =   "&Add"
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin MSComctlLib.ListView lsvFilenameSubst 
         Height          =   1335
         Left            =   600
         TabIndex        =   15
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblEqual 
         Caption         =   "="
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   360
         Width           =   135
      End
   End
End
Attribute VB_Name = "ctlOptSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ControlsEnabled As Boolean

Public Sub SetControlsEnabled(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ControlsEnabled = value
50020
50030  lblSaveFilename.Enabled = value
50040  txtSaveFilename.Enabled = value
50050  lblSaveFilenameTokens.Enabled = value
50060  cmbSaveFilenameTokens.Enabled = value
50070  txtSavePreview.Enabled = value
50080  chkSpaces.Enabled = value
50090  lblStandardSaveformat.Enabled = value
50100  cmbStandardSaveFormat.Enabled = value
50110  chkAllowSpecialGSCharsInFilenames.Enabled = value
50120  dmFraProgSave.Enabled = value
50130
50140  txtFilenameSubst(0).Enabled = value
50150  lblEqual.Enabled = value
50160  txtFilenameSubst(1).Enabled = value
50170  If value = False Then
50180   cmdFilenameSubstMove(0).Enabled = value
50190   cmdFilenameSubstMove(1).Enabled = value
50200  End If
50210  lsvFilenameSubst.Enabled = value
50220  If value = False Then
50230    lsvFilenameSubst.ForeColor = &H80000002
50240   Else
50250    lsvFilenameSubst.ForeColor = &H80000008
50260  End If
50270  cmdFilenameSubst(0).Enabled = value
50280  cmdFilenameSubst(1).Enabled = value
50290  cmdFilenameSubst(2).Enabled = value
50300  chkFilenameSubst.Enabled = value
50310  dmFraFilenameSubstitutions.Enabled = value
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "SetControlsEnabled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  dmFraProgSave.Left = 0
50030  dmFraProgSave.Top = 0
50040  dmFraFilenameSubstitutions.Left = dmFraProgSave.Left
50050  dmFraFilenameSubstitutions.Top = dmFraProgSave.Top + dmFraProgSave.Height + 50
50060
50070  UserControl.Height = dmFraFilenameSubstitutions.Top + dmFraFilenameSubstitutions.Height
50080
50090  With cmbStandardSaveFormat
50100   .AddItem "PDF"
50110   .AddItem "PDF/A-1b"
50120   .AddItem "PDF/X"
50130   .AddItem "PNG"
50140   .AddItem "JPEG"
50150   .AddItem "BMP"
50160   .AddItem "PCX"
50170   .AddItem "TIFF"
50180   .AddItem "PS"
50190   .AddItem "EPS"
50200   .AddItem "TXT"
50210   .AddItem "PSD"
50220   .AddItem "PCL"
50230   .AddItem "RAW"
50240   .AddItem "SVG"
50250  End With
50260  With cmbSaveFilenameTokens
50270   .AddItem "<Author>"
50280   .AddItem "<Computername>"
50290   .AddItem "<DateTime>"
50300   .AddItem "<Title>"
50310   .AddItem "<Username>"
50320   .AddItem "<Counter>"
50330   .AddItem "<REDMON_DOCNAME>"
50340   .AddItem "<REDMON_DOCNAME_FILE>"
50350   .AddItem "<REDMON_DOCNAME_PATH>"
50360   .AddItem "<REDMON_JOB>"
50370   .AddItem "<REDMON_MACHINE>"
50380   .AddItem "<REDMON_PORT>"
50390   .AddItem "<REDMON_PRINTER>"
50400   .AddItem "<REDMON_SESSIONID>"
50410   .AddItem "<REDMON_USER>"
50420   .ListIndex = 0
50430  End With
50440
50450  With lsvFilenameSubst
50460   .Appearance = ccFlat
50470   .ColumnHeaders.Clear
50480   .ColumnHeaders.Add , "Str1", "", lsvFilenameSubst.Width / 2 - 140
50490   .ColumnHeaders.Add , "Str2", "", lsvFilenameSubst.Width / 2 - 140
50500   .HideColumnHeaders = True
50510   .GridLines = True
50520   .FullRowSelect = True
50530   .HideSelection = False
50540  End With
50550
50560  cmdFilenameSubst(0).Top = lsvFilenameSubst.Top
50570  cmdFilenameSubst(1).Top = lsvFilenameSubst.Top + (lsvFilenameSubst.Height - cmdFilenameSubst(1).Height) / 2
50580  cmdFilenameSubst(2).Top = lsvFilenameSubst.Top + lsvFilenameSubst.Height - cmdFilenameSubst(2).Height
50590
50600  ControlsEnabled = True
50610  CheckCmdFilenameSubst
50620
50630  SetFrames Options.OptionsDesign
50640
50650  SetFont
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "UserControl_Initialize")
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
Select Case ErrPtnr.OnError("ctlOptSave", "SetFont")
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
Select Case ErrPtnr.OnError("ctlOptSave", "SetFrames")
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
50020   dmFraProgSave.Caption = .OptionsProgramSaveSymbol
50030   lblSaveFilename.Caption = .OptionsSaveFilename
50040   lblSaveFilenameTokens.Caption = .OptionsSaveFilenameTokens
50050   dmFraFilenameSubstitutions.Caption = .OptionsSaveFilenameSubstitutions
50060   chkFilenameSubst.Caption = .OptionsSaveFilenameSubstitutionsTitle
50070   cmdFilenameSubst(0).Caption = .OptionsSaveFilenameAdd
50080   cmdFilenameSubst(1).Caption = .OptionsSaveFilenameChange
50090   cmdFilenameSubst(2).Caption = .OptionsSaveFilenameDelete
50100   chkSpaces.Caption = .OptionsRemoveSpaces
50110   chkAllowSpecialGSCharsInFilenames.Caption = .OptionsAllowSpecialGSCharsInFilenames
50120   lblStandardSaveformat.Caption = .OptionsStandardSaveFormat
50130  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "SetLanguageStrings")
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
50010  Dim lsv As ListView, i As Long, tList() As String, tStrA() As String
50020  With Options1
50030   lsvFilenameSubst.ListItems.Clear
50040   tList = Split(.FilenameSubstitutions, "\")
50050   For i = 0 To UBound(tList)
50060    If InStr(tList(i), "|") <= 0 Then
50070     tList(i) = tList(i) & "|"
50080    End If
50090    If UBound(Split(tList(i), "|")) = 1 Then
50100     tStrA = Split(tList(i), "|")
50110     lsvFilenameSubst.ListItems.Add , , tStrA(0)
50120     lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).SubItems(1) = tStrA(1)
50130    End If
50140   Next i
50150   If lsvFilenameSubst.ListItems.Count > 0 Then
50160    lsvFilenameSubst.ListItems(1).Selected = True
50170    txtFilenameSubst(0).Text = lsvFilenameSubst.ListItems(1).Text
50180    txtFilenameSubst(0).ToolTipText = txtFilenameSubst(0).Text
50190    txtFilenameSubst(1).Text = lsvFilenameSubst.ListItems(1).SubItems(1)
50200    txtFilenameSubst(1).ToolTipText = txtFilenameSubst(1).Text
50210   End If
50220   chkFilenameSubst.value = .FilenameSubstitutionsOnlyInTitle
50230   chkSpaces.value = .RemoveSpaces
50240   txtSaveFilename.Text = .SaveFilename
50250   cmbStandardSaveFormat.ListIndex = .StandardSaveformat
50260   chkAllowSpecialGSCharsInFilenames.value = .AllowSpecialGSCharsInFilenames
50270  End With
50280  If ControlsEnabled Then
50290   CheckCmdFilenameSubst
50300  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "SetOptions")
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
50010  Dim tStr As String, i As Long
50020  With Options1
50030   tStr = ""
50040   For i = 1 To lsvFilenameSubst.ListItems.Count
50050    If i < lsvFilenameSubst.ListItems.Count Then
50060      tStr = tStr & lsvFilenameSubst.ListItems(i).Text & "|" & lsvFilenameSubst.ListItems(i).SubItems(1) & "\"
50070     Else
50080      tStr = tStr & lsvFilenameSubst.ListItems(i).Text & "|" & lsvFilenameSubst.ListItems(i).SubItems(1)
50090    End If
50100   Next i
50110   .FilenameSubstitutions = tStr
50120   .FilenameSubstitutionsOnlyInTitle = Abs(chkFilenameSubst.value)
50130   .RemoveSpaces = Abs(chkSpaces.value)
50140   .SaveFilename = txtSaveFilename.Text
50150   If LenB(CStr(cmbStandardSaveFormat.ListIndex)) > 0 Then
50160    .StandardSaveformat = cmbStandardSaveFormat.ListIndex
50170   End If
50180   .AllowSpecialGSCharsInFilenames = chkAllowSpecialGSCharsInFilenames.value
50190  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "GetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbSaveFilenameTokens_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtSaveFilename.Text = txtSaveFilename.Text & cmbSaveFilenameTokens.Text
50020  txtSavePreview.ToolTipText = txtSavePreview.Text
50030  txtSavePreview.Text = GetSubstFilename("B:\dummy.dum", txtSaveFilename.Text, , True) & _
  GetSaveAutosaveFormatExtension(cmbStandardSaveFormat.ListIndex)
50050  If IsValidPath("C:\" & txtSavePreview.Text) = False Then
50060    txtSavePreview.ForeColor = vbRed
50070   Else
50080    txtSavePreview.ForeColor = &H80000008
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "cmbSaveFilenameTokens_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbStandardSaveFormat_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtSavePreview.ToolTipText = txtSavePreview.Text
50020  txtSavePreview.Text = GetSubstFilename("B:\dummy.dum", txtSaveFilename.Text, , True) & _
  GetSaveAutosaveFormatExtension(cmbStandardSaveFormat.ListIndex)
50040  If IsValidPath("C:\" & txtSavePreview.Text) = False Then
50050    txtSavePreview.ForeColor = vbRed
50060   Else
50070    txtSavePreview.ForeColor = &H80000008
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "cmbStandardSaveFormat_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsvFilenameSubst_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set_txtFilenameSubst
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "lsvFilenameSubst_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


Private Sub AddFilenameSubstitutions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, res As Long
50020  res = CheckFilenameSubstitutions(0)
50031  Select Case res
              Case 0, 2:
50050    lsvFilenameSubst.ListItems.Add , , txtFilenameSubst(0).Text
50060    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).SubItems(1) = txtFilenameSubst(1).Text
50070    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).Selected = True
50080    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).EnsureVisible
50090    Set_txtFilenameSubst
50100 '  Case 2:
50110 '   MsgBox LanguageStrings.MessagesMsg12 & _
    vbCrLf & vbTab & "\ / : * ? < > | """
50130   Case 3:
50140    MsgBox LanguageStrings.MessagesMsg11
50150  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "AddFilenameSubstitutions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ChangeFilenameSubstitutions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, res As Long
50020  res = CheckFilenameSubstitutions(lsvFilenameSubst.SelectedItem.Index)
50031  Select Case res
              Case 0, 2:
50050    lsvFilenameSubst.SelectedItem.Text = txtFilenameSubst(0).Text
50060    lsvFilenameSubst.SelectedItem.SubItems(1) = txtFilenameSubst(1).Text
50070 '  Case 2:
50080 '   MsgBox LanguageStrings.MessagesMsg12 & _
    vbCrLf & vbTab & "\ / : * ? < > | """
50100   Case 3:
50110    MsgBox LanguageStrings.MessagesMsg11
50120  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "ChangeFilenameSubstitutions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DeleteFilenameSubstitutions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim oIndex As Long
50020  With lsvFilenameSubst
50030   oIndex = .SelectedItem.Index
50040   If .ListItems.Count > 0 Then
50050    .ListItems.Remove .SelectedItem.Index
50060    If oIndex > .ListItems.Count Then
50070     oIndex = .ListItems.Count
50080    End If
50090    If .ListItems.Count > 0 Then
50100     .ListItems(oIndex).Selected = True
50110     .ListItems(oIndex).EnsureVisible
50120    End If
50130    Set_txtFilenameSubst
50140   End If
50150  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "DeleteFilenameSubstitutions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub MoveUpFilenameSubstitutions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStrL As String, tStrR As String
50020  If lsvFilenameSubst.SelectedItem.Index = 1 Then
50030   CheckCmdFilenameSubst
50040   Exit Sub
50050  End If
50060  With lsvFilenameSubst
50070   tStrL = .ListItems(.SelectedItem.Index).Text
50080   tStrR = .ListItems(.SelectedItem.Index).SubItems(1)
50090   .ListItems(.SelectedItem.Index).Text = .ListItems(.SelectedItem.Index - 1).Text
50100   .ListItems(.SelectedItem.Index).SubItems(1) = .ListItems(.SelectedItem.Index - 1).SubItems(1)
50110   .ListItems(.SelectedItem.Index - 1).Text = tStrL
50120   .ListItems(.SelectedItem.Index - 1).SubItems(1) = tStrR
50130   .ListItems(.SelectedItem.Index - 1).Selected = True
50140   .ListItems(.SelectedItem.Index).EnsureVisible
50150  End With
50160  Set_txtFilenameSubst
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "MoveUpFilenameSubstitutions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub MoveDownFilenameSubstitutions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStrL As String, tStrR As String
50020  With lsvFilenameSubst
50030   tStrL = .ListItems(.SelectedItem.Index).Text
50040   tStrR = .ListItems(.SelectedItem.Index).SubItems(1)
50050   .ListItems(.SelectedItem.Index).Text = .ListItems(.SelectedItem.Index + 1).Text
50060   .ListItems(.SelectedItem.Index).SubItems(1) = .ListItems(.SelectedItem.Index + 1).SubItems(1)
50070   .ListItems(.SelectedItem.Index + 1).Text = tStrL
50080   .ListItems(.SelectedItem.Index + 1).SubItems(1) = tStrR
50090   .ListItems(.SelectedItem.Index + 1).Selected = True
50100   .ListItems(.SelectedItem.Index).EnsureVisible
50110  End With
50120  Set_txtFilenameSubst
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "MoveDownFilenameSubstitutions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function CheckFilenameSubstitutions(Index As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  CheckFilenameSubstitutions = 0
50030  If Len(txtFilenameSubst(0).Text) = 0 Then
50040   CheckFilenameSubstitutions = 1
50050   Exit Function
50060  End If
50070  If IsForbiddenChars(txtFilenameSubst(0).Text) = True Then
50080   txtFilenameSubst(0).SetFocus
50090   CheckFilenameSubstitutions = 2
50100   Exit Function
50110  End If
50120  If IsForbiddenChars(txtFilenameSubst(1).Text) = True Then
50130   txtFilenameSubst(1).SetFocus
50140   CheckFilenameSubstitutions = 2
50150   Exit Function
50160  End If
50170  If Index = 0 Then
50180    For i = 1 To lsvFilenameSubst.ListItems.Count
50190     If UCase$(txtFilenameSubst(0).Text) = UCase$(lsvFilenameSubst.ListItems(i).Text) Then
50200      CheckFilenameSubstitutions = 3
50210      Exit Function
50220     End If
50230    Next i
50240   Else
50250    For i = 1 To lsvFilenameSubst.ListItems.Count
50260     If UCase$(txtFilenameSubst(0).Text) = UCase$(lsvFilenameSubst.ListItems(i).Text) And _
     Index <> lsvFilenameSubst.SelectedItem.Index Then
50280      CheckFilenameSubstitutions = 3
50290      Exit Function
50300     End If
50310    Next i
50320  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "CheckFilenameSubstitutions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub CheckCmdFilenameSubst()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If lsvFilenameSubst.ListItems.Count > 0 Then
50020    cmdFilenameSubst(1).Enabled = True
50030    cmdFilenameSubst(2).Enabled = True
50040   Else
50050    cmdFilenameSubst(1).Enabled = False
50060    cmdFilenameSubst(2).Enabled = False
50070  End If
50080  If lsvFilenameSubst.ListItems.Count > 1 Then
50090    cmdFilenameSubstMove(0).Enabled = True
50100    cmdFilenameSubstMove(1).Enabled = True
50110   Else
50120    cmdFilenameSubstMove(0).Enabled = False
50130    cmdFilenameSubstMove(1).Enabled = False
50140  End If
50150  If lsvFilenameSubst.ListItems.Count > 0 Then
50160   If lsvFilenameSubst.SelectedItem.Index = 1 Then
50170    cmdFilenameSubstMove(0).Enabled = False
50180   End If
50190   If lsvFilenameSubst.SelectedItem.Index = lsvFilenameSubst.ListItems.Count Then
50200    cmdFilenameSubstMove(1).Enabled = False
50210   End If
50220  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "CheckCmdFilenameSubst")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Set_txtFilenameSubst()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  CheckCmdFilenameSubst
50020  If lsvFilenameSubst.ListItems.Count > 0 Then
50030   txtFilenameSubst(0).Text = lsvFilenameSubst.SelectedItem.Text
50040   txtFilenameSubst(0).ToolTipText = txtFilenameSubst(0).Text
50050   txtFilenameSubst(1).Text = lsvFilenameSubst.SelectedItem.SubItems(1)
50060   txtFilenameSubst(1).ToolTipText = txtFilenameSubst(1).Text
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "Set_txtFilenameSubst")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdFilenameSubst_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 0: ' Add
50030    AddFilenameSubstitutions
50040   Case 1: ' Change
50050    ChangeFilenameSubstitutions
50060   Case 2: ' Delete
50070    DeleteFilenameSubstitutions
50080  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "cmdFilenameSubst_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdFilenameSubstMove_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 0: ' Up
50030    MoveUpFilenameSubstitutions
50040   Case 1: ' Down
50050    MoveDownFilenameSubstitutions
50060  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "cmdFilenameSubstMove_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtSaveFilename_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtSaveFilename.ToolTipText = txtSaveFilename.Text
50020  With txtSavePreview
50030   .Text = GetSubstFilename("C:\test.pdf", txtSaveFilename.Text, , True) & ".pdf"
50040   .ToolTipText = .Text
50050  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "txtSaveFilename_Change")
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
50010  dmFraProgSave.Width = UserControl.Width
50020  dmFraFilenameSubstitutions.Width = dmFraProgSave.Width
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptSave", "UserControl_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
