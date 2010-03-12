VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Options"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.TreeView trvOptions 
      Height          =   7815
      Left            =   0
      TabIndex        =   13
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   13785
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin PDFCreator.dmFrame dmFraDescription 
      Height          =   1065
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1879
      Caption         =   ""
      BarColorFrom    =   723949
      BarColorTo      =   132452
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox picOptions 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   105
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   4
         Top             =   420
         Width           =   615
      End
      Begin VB.Label lblOptions 
         Height          =   615
         Left            =   735
         TabIndex        =   5
         Top             =   420
         Width           =   5655
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   8640
      Width           =   1815
   End
   Begin PDFCreator.dmFrame dmFraProfile 
      Height          =   1065
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   1879
      Caption         =   "Profil"
      BarColorFrom    =   723949
      BarColorTo      =   132452
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdProfileRename 
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         Picture         =   "frmOptions.frx":000C
         Style           =   1  'Grafisch
         TabIndex        =   12
         ToolTipText     =   "Rename profile"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdProfileLoad 
         Height          =   375
         Left            =   8640
         Picture         =   "frmOptions.frx":0400
         Style           =   1  'Grafisch
         TabIndex        =   11
         ToolTipText     =   "Load profile"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdProfileSave 
         Height          =   375
         Left            =   8160
         Picture         =   "frmOptions.frx":07FE
         Style           =   1  'Grafisch
         TabIndex        =   10
         ToolTipText     =   "Save profile"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdProfileDelete 
         Enabled         =   0   'False
         Height          =   375
         Left            =   7680
         Picture         =   "frmOptions.frx":0B93
         Style           =   1  'Grafisch
         TabIndex        =   9
         ToolTipText     =   "Delete profile"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdProfileAdd 
         Height          =   375
         Left            =   6720
         Picture         =   "frmOptions.frx":0F8B
         Style           =   1  'Grafisch
         TabIndex        =   8
         ToolTipText     =   "Add profile"
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox cmbProfile 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   7
         Top             =   480
         Width           =   6375
      End
   End
   Begin MSComctlLib.ImageList imlIeb 
      Left            =   2760
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":137B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":14D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1A6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":2009
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":25A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":293D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":2ED7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":37B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":3D4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":42E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":487F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":4E19
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":53B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":594D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":5EE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":6481
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":6A1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":6FB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":754F
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":7E29
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":8703
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":8C9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":9237
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":97D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":9D6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":A305
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private UnloadForm As Boolean, LoadReady As Boolean, oldLanguage As String, Languages As Collection
 
Private optActionsControl As VBControlExtender, optActions As ctlOptActions
Private optAutosaveControl As VBControlExtender, optAutosave As ctlOptAutosave
'Private optDirectoriesControl As VBControlExtender, optDirectories As ctlOptDirectories
Private optDocumentControl As VBControlExtender, optDocument As ctlOptDocument
Private optFontsControl As VBControlExtender, optFonts As ctlOptFonts
Private optFormatPDFControl As VBControlExtender, optFormatPDF As ctlOptFormatPDF
Private optFormatPSControl As VBControlExtender, optFormatPS As ctlOptFormatPS
Private optFormatEPSControl As VBControlExtender, optFormatEPS As ctlOptFormatEPS
Private optFormatPNGControl As VBControlExtender, optFormatPNG As ctlOptFormatPNG
Private optFormatJPEGControl As VBControlExtender, optFormatJPEG As ctlOptFormatJPEG
Private optFormatBMPControl As VBControlExtender, optFormatBMP As ctlOptFormatBMP
Private optFormatPCXControl As VBControlExtender, optFormatPCX As ctlOptFormatPCX
Private optFormatTIFFControl As VBControlExtender, optFormatTIFF As ctlOptFormatTIFF
Private optFormatTXTControl As VBControlExtender, optFormatTXT As ctlOptFormatTXT
Private optFormatPSDControl As VBControlExtender, optFormatPSD As ctlOptFormatPSD
Private optFormatPCLControl As VBControlExtender, optFormatPCL As ctlOptFormatPCL
Private optFormatRAWControl As VBControlExtender, optFormatRAW As ctlOptFormatRAW
Private optFormatSVGControl As VBControlExtender, optFormatSVG As ctlOptFormatSVG
'Private optFormatXCFControl As VBControlExtender, optFormatXCF As ctlOptFormatXCF
Private optGeneralControl As VBControlExtender, optGeneral As ctlOptGeneral
Private optGhostscriptControl As VBControlExtender, optGhostscript As ctlOptGhostscript
Private optLanguagesControl As VBControlExtender, optLanguages As ctlOptLanguages
Private optPrintControl As VBControlExtender, optPrint As ctlOptPrint
Private optSaveControl As VBControlExtender, optSave As ctlOptSave

Private OldProfile As Long, ProfileOptions() As tOptions, ProfileNames() As String, TempPrinterProfiles As Collection, LastNodeKey As String

Private Sub cmbProfile_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If cmbProfile.ListIndex <> OldProfile Then
50020   If OldProfile <= UBound(ProfileOptions) Then
50030    ProfileOptions(OldProfile) = GetOptionsFromUserControls(ProfileOptions(OldProfile))
50040   End If
50050   OldProfile = cmbProfile.ListIndex
50060   If cmbProfile.ListIndex = 0 Then
50070     optGhostscript.ControlEnabled = True
50080     optLanguages.ControlEnabled = True
50090     cmdProfileRename.Enabled = False
50100     cmdProfileDelete.Enabled = False
50110    Else
50120     optGhostscript.ControlEnabled = False
50130     optLanguages.ControlEnabled = False
50140     cmdProfileRename.Enabled = True
50150     cmdProfileDelete.Enabled = True
50160   End If
50170   Options1 = ProfileOptions(cmbProfile.ListIndex)
50180   Options1.Language = CurrentLanguage
50190   SetOptions
50200  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbProfile_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub AddProfile(ProfileName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim resS As String, i As Long
50020  resS = Trim$(ProfileName)
50030  If LenB(resS) = 0 Then
50040   Exit Sub
50050  End If
50060  With cmbProfile
50070   ReDim Preserve ProfileOptions(.ListCount)
50080   ProfileOptions(.ListCount) = StandardOptions
50090   .AddItem resS
50100   .ListIndex = cmbProfile.ListCount - 1
50110   .Enabled = True
50120   .Visible = True
50130  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "AddProfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub RenameProfile(ProfileName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim resS As String, i As Long, NewPrinterProfiles As Collection, tStr As String, sa(2) As String
50020
50030  resS = Trim$(ProfileName)
50040  If LenB(resS) = 0 Then
50050   Exit Sub
50060  End If
50070
50080  Set NewPrinterProfiles = New Collection
50090  For i = 1 To TempPrinterProfiles.Count
50100   sa(0) = TempPrinterProfiles(i)(0)
50110   sa(1) = TempPrinterProfiles(i)(1)
50120   If i - 1 = cmbProfile.ListIndex Then
50130     sa(2) = resS
50140    Else
50150     sa(2) = TempPrinterProfiles(i)(1)
50160   End If
50170   NewPrinterProfiles.Add sa
50180  Next i
50190  Set TempPrinterProfiles = NewPrinterProfiles
50200
50210  With cmbProfile
50220   .List(.ListIndex) = resS
50230   .Enabled = True
50240   .Visible = True
50250  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "RenameProfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function ProfileExists(ProfileName) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  For i = 0 To cmbProfile.ListCount - 1
50030   If StrComp(cmbProfile.List(i), ProfileName, vbTextCompare) = 0 Then
50040    ProfileExists = True
50050    Exit Function
50060   End If
50070  Next i
50080  ProfileExists = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ProfileExists")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetNextNewProfile(PreFix As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Const MaxCount = 10000
50020  Dim i As Long, j As Long, NewProfile As String
50030  For i = 1 To MaxCount
50040   NewProfile = PreFix & " " & CStr(i)
50050   If Not ProfileExists(NewProfile) Then
50060    GetNextNewProfile = NewProfile
50070    Exit Function
50080   End If
50090  Next i
50100  GetNextNewProfile = PreFix
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "GetNextNewProfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub cmdProfileAdd_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim resS As String, i As Long, res As Long
50020
50030  Dim Profiles As New Collection
50040  For i = 0 To cmbProfile.ListCount - 1
50050   Profiles.Add cmbProfile.List(i)
50060  Next i
50070
50080  Set frmProfile.Profiles = Profiles
50090  frmProfile.dmFrProfile.Caption = LanguageStrings.OptionsProfileAdd
50100  frmProfile.ProfileAction = eProfileAction.AddProfileAction
50110  frmProfile.txtProfile.Text = GetNextNewProfile(Trim$(LanguageStrings.OptionsProfileNewProfile))
50120  frmProfile.txtProfile.SelStart = Len(frmProfile.txtProfile.Text)
50130  frmProfile.Show vbModal, Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdProfileAdd_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdProfileDelete_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim aw As Long, tStr As String, CurrentProfile As String, i As Long
50020  If cmbProfile.ListCount <= 1 Then
50030   Exit Sub
50040  End If
50050  CurrentProfile = cmbProfile.List(cmbProfile.ListIndex)
50060  tStr = ProfileAssociatedPrinters(CurrentProfile)
50070  If LenB(tStr) > 0 Then
50080   MsgBox LanguageStrings.MessagesMsg43 & " (" & tStr & ")"
50090   Exit Sub
50100  End If
50110
50120  aw = MsgBox(Replace(LanguageStrings.MessagesMsg42, "%1", cmbProfile.List(cmbProfile.ListIndex)), vbQuestion Or vbYesNo)
50130  If aw = vbYes Then
50140   For i = cmbProfile.ListIndex + 1 To cmbProfile.ListCount - 1
50150    ProfileOptions(i - 1) = ProfileOptions(i)
50160   Next i
50170   ReDim Preserve ProfileOptions(cmbProfile.ListCount - 2)
50180   cmbProfile.RemoveItem cmbProfile.ListIndex
50190   cmbProfile.ListIndex = 0
50200  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdProfileDelete_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdProfileLoad_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim FilterIndex As Long, files As Collection, dummyOptions As tOptions, tempOptions As tOptions
50020  FilterIndex = OpenFileDialog(files, cmbProfile.List(cmbProfile.ListIndex), _
   "PDFCreator options files (*.ini)|*.ini|All files (*.*)|*.*", "*.ini", GetMyFiles, _
   App.ProductName, OFN_ALLOWMULTISELECT + OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_LONGNAMES + OFN_NODEREFERENCELINKS, Me.hwnd)
50050  If FilterIndex > 0 Then
50060   tempOptions = ReadOptionsINI(dummyOptions, files(1), 0, True, True)
50070   ProfileOptions(cmbProfile.ListIndex) = tempOptions
50080   Options1 = ProfileOptions(cmbProfile.ListIndex)
50090   Options1.Language = CurrentLanguage
50100   SetOptions
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdProfileLoad_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdProfileRename_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim resS As String, i As Long, res As Long
50020
50030  Dim Profiles As New Collection
50040  For i = 0 To cmbProfile.ListCount - 1
50050   Profiles.Add cmbProfile.List(i)
50060  Next i
50070
50080  Set frmProfile.Profiles = Profiles
50090  frmProfile.dmFrProfile.Caption = LanguageStrings.OptionsProfileRenameProfile
50100  frmProfile.ProfileAction = eProfileAction.RenameProfileAction
50110  frmProfile.txtProfile.Text = cmbProfile.List(cmbProfile.ListIndex)
50120  frmProfile.txtProfile.SelStart = Len(frmProfile.txtProfile.Text)
50130  frmProfile.CurrentProfile = cmbProfile.List(cmbProfile.ListIndex)
50140  frmProfile.Show vbModal, Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdProfileRename_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdProfileSave_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim FName As String, res As Long, tempOptions As tOptions
50020  res = SaveFileDialog(FName, cmbProfile.List(cmbProfile.ListIndex), "PDFCreator options files (*.ini)|*.ini|All files (*.*)|*.*", "*.ini", _
  GetMyFiles, App.ProductName, OFN_EXPLORER + OFN_PATHMUSTEXIST + OFN_LONGNAMES + OFN_HIDEREADONLY + OFN_OVERWRITEPROMPT, Me.hwnd)
50040  If res > 0 Then
50050   tempOptions = GetOptionsFromUserControls(ProfileOptions(cmbProfile.ListIndex))   ' Get the current settings settings
50060   SaveOptionsINI tempOptions, FName
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdProfileSave_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyCode = vbKeyF1 Then
50020   KeyCode = 0
50031   Select Case trvOptions.SelectedItem.key
         Case "Program", "ProgramGeneral"
50050     Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50060    Case "ProgramGhostscript"
50070     Call HTMLHelp_ShowTopic("html\ghostscript.htm")
50080    Case "ProgramDocument"
50090     Call HTMLHelp_ShowTopic("html\docproperties.htm")
50100    Case "ProgramSave"
50110     Call HTMLHelp_ShowTopic("html\savesettings.htm")
50120    Case "ProgramAutosave"
50130     Call HTMLHelp_ShowTopic("html\autosave.htm")
50140    Case "ProgramActions"
50150     Call HTMLHelp_ShowTopic("html\actions.htm")
50160    Case "ProgramPrint"
50170     Call HTMLHelp_ShowTopic("html\print.htm")
50180    Case "ProgramFonts"
50190     Call HTMLHelp_ShowTopic("html\fontsettings.htm")
50200    Case "ProgramLanguages"
50210     Call HTMLHelp_ShowTopic("html\changelang.htm")
50220    Case "Formats", "FormatsPDF"
50230     If trvOptions.SelectedItem.key = "FormatsPDF" Then
50241       Select Case optFormatPDF.PDFOptionsIndex
             Case 1
50260         Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50270        Case 2
50280         Call HTMLHelp_ShowTopic("html\pdfcompression.htm")
50290        Case 3
50300         Call HTMLHelp_ShowTopic("html\pdffonts.htm")
50310        Case 4
50320         Call HTMLHelp_ShowTopic("html\pdfcolors.htm")
50330        Case 5
50340         Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
50350        Case 6
50360         Call HTMLHelp_ShowTopic("html\pdfsigning.htm")
50370       Case Else
50380        Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50390       End Select
50400      Else
50410       Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50420     End If
50430    Case "FormatsPNG"
50440     Call HTMLHelp_ShowTopic("html\pngsettings.htm")
50450    Case "FormatsJPEG"
50460     Call HTMLHelp_ShowTopic("html\jpegsettings.htm")
50470    Case "FormatsBMP"
50480     Call HTMLHelp_ShowTopic("html\bmpsettings.htm")
50490    Case "FormatsPCX"
50500     Call HTMLHelp_ShowTopic("html\pcxsettings.htm")
50510    Case "FormatsTIFF"
50520     Call HTMLHelp_ShowTopic("html\tiffsettings.htm")
50530    Case "FormatsPS"
50540     Call HTMLHelp_ShowTopic("html\pssettings.htm")
50550    Case "FormatsEPS"
50560     Call HTMLHelp_ShowTopic("html\epssettings.htm")
50570    Case "FormatsTXT"
50580     Call HTMLHelp_ShowTopic("html\txtsettings.htm")
50590    Case "FormatsPSD"
50600     Call HTMLHelp_ShowTopic("html\psdsettings.htm")
50610    Case "FormatsPCL"
50620     Call HTMLHelp_ShowTopic("html\pclsettings.htm")
50630    Case "FormatsRAW"
50640     Call HTMLHelp_ShowTopic("html\rawsettings.htm")
50650    Case "FormatsSVG"
50660     Call HTMLHelp_ShowTopic("html\svgsettings.htm")
50670    Case Else
50680     Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50690    End Select
50700  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "Form_KeyDown")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub ChangeLanguage()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Not LoadReady Then
50020   Exit Sub
50030  End If
50040
50050  With LanguageStrings
50060   dmFraProfile.Caption = .OptionsProfile
50070   cmdProfileAdd.ToolTipText = .OptionsProfileAdd
50080   cmdProfileDelete.ToolTipText = .OptionsProfileDel
50090   cmdProfileRename.ToolTipText = .OptionsProfileRenameProfile
50100   cmdProfileSave.ToolTipText = .OptionsProfileSaveToDisc
50110   cmdProfileLoad.ToolTipText = .OptionsProfileLoadFromDisc
50120   cmbProfile.List(0) = .OptionsProfileDefaultName
50130
50140   dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50150   Me.Caption = .DialogPrinterOptions
50160   cmdCancel.Caption = .OptionsCancel
50170   cmdReset.Caption = .OptionsReset
50180   cmdSave.Caption = .OptionsSave
50190
50200   trvOptions.Nodes("Program").Text = .OptionsTreeProgram
50210   trvOptions.Nodes("ProgramGeneral").Text = .OptionsProgramGeneralSymbol
50220   trvOptions.Nodes("ProgramGhostscript").Text = .OptionsProgramGhostscriptSymbol
50230   trvOptions.Nodes("ProgramDocument").Text = .OptionsProgramDocumentSymbol
50240   trvOptions.Nodes("ProgramSave").Text = .OptionsProgramSaveSymbol
50250   trvOptions.Nodes("ProgramAutoSave").Text = .OptionsProgramAutosaveSymbol
50260   trvOptions.Nodes("ProgramActions").Text = .OptionsProgramActionsSymbol
50270   trvOptions.Nodes("ProgramPrint").Text = .OptionsProgramPrintSymbol
50280   trvOptions.Nodes("ProgramFonts").Text = .OptionsProgramFontSymbol
50290   trvOptions.Nodes("ProgramLanguages").Text = .OptionsProgramLanguagesSymbol
50300
50310   trvOptions.Nodes("Formats").Text = .OptionsTreeFormats
50320   trvOptions.Nodes("FormatsPDF").Text = .OptionsPDFSymbol
50330   trvOptions.Nodes("FormatsPNG").Text = .OptionsPNGSymbol
50340   trvOptions.Nodes("FormatsJPEG").Text = .OptionsJPEGSymbol
50350   trvOptions.Nodes("FormatsBMP").Text = .OptionsBMPSymbol
50360   trvOptions.Nodes("FormatsPCX").Text = .OptionsPCXSymbol
50370   trvOptions.Nodes("FormatsTIFF").Text = .OptionsTIFFSymbol
50380   trvOptions.Nodes("FormatsPS").Text = .OptionsPSSymbol
50390   trvOptions.Nodes("FormatsEPS").Text = .OptionsEPSSymbol
50400   trvOptions.Nodes("FormatsTXT").Text = .OptionsTXTSymbol
50410   trvOptions.Nodes("FormatsPSD").Text = .OptionsPSDSymbol
50420   trvOptions.Nodes("FormatsPCL").Text = .OptionsPCLSymbol
50430   trvOptions.Nodes("FormatsRAW").Text = .OptionsRAWSymbol
50440   trvOptions.Nodes("FormatsSVG").Text = .OptionsSVGSymbol
50450
50460   lblOptions.Caption = .OptionsProgramLanguagesDescription
50470  End With
50480  optActions.SetLanguageStrings
50490  optAutosave.SetLanguageStrings
50500 ' optDirectories.SetLanguageStrings
50510  optDocument.SetLanguageStrings
50520 ' optFonts.SetLanguageStrings
50530  optFormatPNG.SetLanguageStrings
50540  optFormatJPEG.SetLanguageStrings
50550  optFormatBMP.SetLanguageStrings
50560  optFormatPCX.SetLanguageStrings
50570  optFormatTIFF.SetLanguageStrings
50580  optFormatPDF.SetLanguageStrings
50590  optFormatPS.SetLanguageStrings
50600  optFormatEPS.SetLanguageStrings
50610  optFormatTXT.SetLanguageStrings
50620  optFormatPSD.SetLanguageStrings
50630  optFormatPCL.SetLanguageStrings
50640  optFormatRAW.SetLanguageStrings
50650  optFormatSVG.SetLanguageStrings
50660 ' optFormatXCF.SetLanguageStrings
50670
50680  optGeneral.SetLanguageStrings
50690  optGhostscript.SetLanguageStrings
50700  optLanguages.SetLanguageStrings
50710  optPrint.SetLanguageStrings
50720  optSave.SetLanguageStrings
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ChangeLanguage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdCancel_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim f As Form, LanguagePath As String
50020
50030  If CurrentLanguage <> Options.Language Then
50040   SetLanguage oldLanguage
50050   LanguagePath = CompletePath(GetPDFCreatorApplicationPath) & "Languages\"
50060   LoadLanguage LanguagePath & oldLanguage & ".ini"
50070   For Each f In Forms
50080    f.ChangeLanguage
50090   Next
50100  End If
50110
50120  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdCancel_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdReset_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long
50020  res = MsgBox(LanguageStrings.MessagesMsg03, vbYesNo)
50030  If res = vbYes Then
50040   Options = StandardOptions
50050
50060   optActions.SetOptions
50070   optAutosave.SetOptions
50080 '  optDirectories.SetOptions
50090   optDocument.SetOptions
50100 '  optFonts.SetOptions
50110   optFormatPNG.SetOptions
50120   optFormatJPEG.SetOptions
50130   optFormatBMP.SetOptions
50140   optFormatPCX.SetOptions
50150   optFormatTIFF.SetOptions
50160   optFormatPDF.SetOptions
50170   optFormatPS.SetOptions
50180   optFormatEPS.SetOptions
50190   optFormatTXT.SetOptions
50200   optFormatPSD.SetOptions
50210   optFormatPCL.SetOptions
50220   optFormatRAW.SetOptions
50230   optFormatSVG.SetOptions
50240 '  optFormatXCF.SetOptions
50250
50260   optGeneral.SetOptions
50270   optGhostscript.SetOptions
50280   optLanguages.SetOptions
50290   optPrint.SetOptions
50300   optSave.SetOptions
50310
50320   With Options
50330    SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50340    SetFont frmMain, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50350   End With
50360  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdReset_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetOptionsFromUserControls(DefaultOptions As tOptions) As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Options1 = DefaultOptions
50020  optActions.GetOptions
50030  optAutosave.GetOptions
50040 ' optDirectories.GetOptions
50050  optDocument.GetOptions
50060  optFonts.GetOptions
50070  optFormatPNG.GetOptions
50080  optFormatJPEG.GetOptions
50090  optFormatBMP.GetOptions
50100  optFormatPCX.GetOptions
50110  optFormatTIFF.GetOptions
50120  optFormatPDF.GetOptions
50130  optFormatPS.GetOptions
50140  optFormatEPS.GetOptions
50150  optFormatTXT.GetOptions
50160  optFormatPSD.GetOptions
50170  optFormatPCL.GetOptions
50180  optFormatRAW.GetOptions
50190  optFormatSVG.GetOptions
50200 ' optFormatXCF.GetOptions
50210
50220  optGeneral.GetOptions
50230  optGhostscript.GetOptions
50240  optLanguages.GetOptions
50250  optPrint.GetOptions
50260  optSave.GetOptions
50270
50280  Options1.Counter = Options.Counter
50290  GetOptionsFromUserControls = Options1
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "GetOptionsFromUserControls")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub cmdSave_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tRestart As Boolean, newLanguage As String, Profiles As Collection, i As Long, j As Long, _
  PrinterProfiles As Collection, sa(2) As String, tStr As String
50030
50040  ' Save all Options/Profiles
50050  ProfileOptions(cmbProfile.ListIndex) = GetOptionsFromUserControls(ProfileOptions(cmbProfile.ListIndex)) ' Get the current settings
50060
50070  tRestart = False
50080  If UCase$(Options.DirectoryGhostscriptBinaries) <> UCase$(ProfileOptions(0).DirectoryGhostscriptBinaries) Then
50090   tRestart = True
50100  End If
50110
50120  ' Save default options
50130  Options = ProfileOptions(0)
50140  Options.Language = CurrentLanguage
50150  SaveOptions Options
50160
50170  ' Delete all unnecessary/renamed profiles
50180  Set Profiles = GetProfiles
50190  For i = 1 To Profiles.Count
50200   For j = 1 To cmbProfile.ListCount
50210    If Profiles(i) = cmbProfile.List(j) Then
50220     Exit For
50230    End If
50240   Next j
50250   If j > cmbProfile.ListCount Then
50260    DeleteProfile Profiles(i)
50270   End If
50280  Next i
50290
50300  ' Add all new/renamed profiles
50310  For i = 1 To cmbProfile.ListCount - 1
50320   ProfileOptions(i).LastUpdateCheck = Options.LastUpdateCheck
50330   SaveOptions ProfileOptions(i), cmbProfile.List(i)
50340  Next i
50350  ' Ready profiles saving
50360
50370  Set PrinterProfiles = New Collection
50380  For i = 1 To TempPrinterProfiles.Count
50390   sa(0) = TempPrinterProfiles(i)(0)
50400   sa(1) = TempPrinterProfiles(i)(2)
50410   PrinterProfiles.Add sa
50420  Next i
50430
50440  SavePrinterProfiles PrinterProfiles
50450
50460  SetHelpfile
50470
50480  If IsWin9xMe = False Then
50491   Select Case Options.ProcessPriority
         Case 0: 'Idle
50510     SetProcessPriority Idle
50520    Case 1: 'Normal
50530     SetProcessPriority Normal
50540    Case 2: 'High
50550     SetProcessPriority High
50560    Case 3: 'Realtime
50570     SetProcessPriority RealTime
50580   End Select
50590  End If
50600  If tRestart = True Then
50610   Restart = True
50620  End If
50630  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdSave_Click")
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
50010  Const ControlTop = 100, fraPDFTop = 1360, fraPDFLeft = 2960
50020  Dim pic As New StdPicture, i As Long, tStr As String, gsvers As Collection, _
  fc As Long, reg As clsRegistry, tsf() As String, tStr2 As String, files As Collection, _
  Path As String, filename As String, Ext As String, p As Printer
50050  Dim ctl1 As Control, ctl2 As Control, sa(2) As String
50060  Dim Profiles As Collection, PrinterProfiles As Collection
50070  Dim nodeProgram As Node, nodeFormats As Node
50080
50090  Set TempPrinterProfiles = New Collection
50100
50110  Options1 = Options
50120  CurrentLanguage = Options.Language
50130
50140  UnloadForm = False
50150  Me.Icon = LoadResPicture(2120, vbResIcon)
50160  KeyPreview = True
50170
50180  cmbProfile.Top = (cmdProfileAdd.Height - cmbProfile.Height) / 2 + cmdProfileAdd.Top
50190
50200  oldLanguage = Options.Language
50210
50220  With Screen
50230   .MousePointer = vbHourglass
50240   Move (.Width - Width) / 2, (.Height - Height) / 2
50250  End With
50260
50270  trvOptions.Indentation = 0
50280  trvOptions.LineStyle = tvwTreeLines
50290  Set trvOptions.ImageList = imlIeb
50300  trvOptions.Nodes.Clear
50310  With LanguageStrings
50320   Set nodeProgram = trvOptions.Nodes.Add(, , "Program", .OptionsTreeProgram, 1)
50330   trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramGeneral", .OptionsProgramGeneralSymbol, 1
50340   trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramGhostscript", .OptionsProgramGhostscriptSymbol, 2
50350   trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramDocument", .OptionsProgramDocumentSymbol, 3
50360   trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramSave", .OptionsProgramSaveSymbol, 4
50370   trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramAutosave", .OptionsProgramAutosaveSymbol, 5
50380   trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramActions", .OptionsProgramActionsSymbol, 7
50390   trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramPrint", .OptionsProgramPrintSymbol, 8
50400   trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramFonts", .OptionsProgramFontSymbol, 9
50410   trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramLanguages", .OptionsProgramLanguagesSymbol, 10
50420   Set nodeFormats = trvOptions.Nodes.Add(, , "Formats", .OptionsTreeFormats, 11)
50430   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPDF", .OptionsPDFSymbol, 11
50440   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPNG", .OptionsPNGSymbol, 12
50450   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsJPEG", .OptionsJPEGSymbol, 13
50460   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsBMP", .OptionsBMPSymbol, 14
50470   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPCX", .OptionsPCXSymbol, 15
50480   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsTIFF", .OptionsTIFFSymbol, 16
50490   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPS", .OptionsPSSymbol, 17
50500   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsEPS", .OptionsEPSSymbol, 18
50510   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsTXT", .OptionsTXTSymbol, 21
50520   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPSD", .OptionsPSDSymbol, 22
50530   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPCL", .OptionsPCLSymbol, 23
50540   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsRAW", .OptionsRAWSymbol, 24
50550   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsSVG", .OptionsSVGSymbol, 26
50560  End With
50570  nodeProgram.Expanded = True
50580  nodeFormats.Expanded = True
50590
50600  With LanguageStrings
50610   Set picOptions = LoadResPicture(2101, vbResIcon)
50620   Me.Caption = .DialogPrinterOptions
50630   cmdCancel.Caption = .OptionsCancel
50640   cmdReset.Caption = .OptionsReset
50650   cmdSave.Caption = .OptionsSave
50660  End With
50670
50680  SetFrame dmFraDescription
50690  SetFrame dmFraProfile
50700
50710  ' Add ActionsControl
50720  Set optActionsControl = Controls.Add("PDFCreator.ctlOptActions", "ctlOptActions")
50730  optActionsControl.Width = dmFraDescription.Width
50740  Set optActions = optActionsControl.object
50750  optActions.SetLanguageStrings
50760  optActions.SetOptions
50770  ' Add AutosaveControl
50780  Set optAutosaveControl = Controls.Add("PDFCreator.ctlOptAutosave", "ctlOptAutosave")
50790  optAutosaveControl.Width = dmFraDescription.Width
50800  Set optAutosave = optAutosaveControl.object
50810  optAutosave.SetLanguageStrings
50820  optAutosave.SetOptions
50830  ' Add DirectoriesControl
50840 ' Set optDirectoriesControl = Controls.Add("PDFCreator.ctlOptDirectories", "ctlOptDirectories")
50850 ' optDirectoriesControl.Width = dmFraDescription.Width
50860 ' Set optDirectories = optDirectoriesControl.object
50870 ' optDirectories.SetLanguageStrings
50880 ' optDirectories.SetOptions
50890  ' Add DocumentControl
50900  Set optDocumentControl = Controls.Add("PDFCreator.ctlOptDocument", "ctlOptDocument")
50910  optDocumentControl.Width = dmFraDescription.Width
50920  Set optDocument = optDocumentControl.object
50930  optDocument.SetLanguageStrings
50940  optDocument.SetOptions
50950  ' Add FontsControl
50960  Set optFontsControl = Controls.Add("PDFCreator.ctlOptFonts", "ctlOptFonts")
50970  optFontsControl.Width = dmFraDescription.Width
50980  Set optFonts = optFontsControl.object
50990  optFonts.SetLanguageStrings
51000  optFonts.SetOptions
51010  ' Add FormatPNGControl
51020  Set optFormatPNGControl = Controls.Add("PDFCreator.ctlOptFormatPNG", "ctlOptFormatPNG")
51030  optFormatPNGControl.Width = dmFraDescription.Width
51040  Set optFormatPNG = optFormatPNGControl.object
51050  optFormatPNG.SetLanguageStrings
51060  optFormatPNG.SetOptions
51070  ' Add FormatJPEQControl
51080  Set optFormatJPEGControl = Controls.Add("PDFCreator.ctlOptFormatJPEG", "ctlOptFormatJPEG")
51090  optFormatJPEGControl.Width = dmFraDescription.Width
51100  Set optFormatJPEG = optFormatJPEGControl.object
51110  optFormatJPEG.SetLanguageStrings
51120  optFormatJPEG.SetOptions
51130  ' Add FormatBMPControl
51140  Set optFormatBMPControl = Controls.Add("PDFCreator.ctlOptFormatBMP", "ctlOptFormatBMP")
51150  optFormatBMPControl.Width = dmFraDescription.Width
51160  Set optFormatBMP = optFormatBMPControl.object
51170  optFormatBMP.SetLanguageStrings
51180  optFormatBMP.SetOptions
51190  ' Add FormatPCXControl
51200  Set optFormatPCXControl = Controls.Add("PDFCreator.ctlOptFormatPCX", "ctlOptFormatPCX")
51210  optFormatPCXControl.Width = dmFraDescription.Width
51220  Set optFormatPCX = optFormatPCXControl.object
51230  optFormatPCX.SetLanguageStrings
51240  optFormatPCX.SetOptions
51250  ' Add FormatTIFFControl
51260  Set optFormatTIFFControl = Controls.Add("PDFCreator.ctlOptFormatTIFF", "ctlOptFormatTIFF")
51270  optFormatTIFFControl.Width = dmFraDescription.Width
51280  Set optFormatTIFF = optFormatTIFFControl.object
51290  optFormatTIFF.SetLanguageStrings
51300  optFormatTIFF.SetOptions
51310  ' Add FormatPDFControl
51320  Set optFormatPDFControl = Controls.Add("PDFCreator.ctlOptFormatPDF", "ctlOptFormatPDF")
51330  optFormatPDFControl.Width = dmFraDescription.Width
51340  Set optFormatPDF = optFormatPDFControl.object
51350  optFormatPDF.SetLanguageStrings
51360  optFormatPDF.SetOptions
51370  ' Add FormatPS
51380  Set optFormatPSControl = Controls.Add("PDFCreator.ctlOptFormatPS", "ctlOptFormatPS")
51390  optFormatPSControl.Width = dmFraDescription.Width
51400  Set optFormatPS = optFormatPSControl.object
51410  optFormatPS.SetLanguageStrings
51420  optFormatPS.SetOptions
51430  ' Add FormatEPSControl
51440  Set optFormatEPSControl = Controls.Add("PDFCreator.ctlOptFormatEPS", "ctlOptFormatEPS")
51450  optFormatEPSControl.Width = dmFraDescription.Width
51460  Set optFormatEPS = optFormatEPSControl.object
51470  optFormatEPS.SetLanguageStrings
51480  optFormatEPS.SetOptions
51490  ' Add FormatTXTControl
51500  Set optFormatTXTControl = Controls.Add("PDFCreator.ctlOptFormatTXT", "ctlOptFormatTXT")
51510  optFormatTXTControl.Width = dmFraDescription.Width
51520  Set optFormatTXT = optFormatTXTControl.object
51530  optFormatTXT.SetLanguageStrings
51540  optFormatTXT.SetOptions
51550  ' Add FormatPSDControl
51560  Set optFormatPSDControl = Controls.Add("PDFCreator.ctlOptFormatPSD", "ctlOptFormatPSD")
51570  optFormatPSDControl.Width = dmFraDescription.Width
51580  Set optFormatPSD = optFormatPSDControl.object
51590  optFormatPSD.SetLanguageStrings
51600  optFormatPSD.SetOptions
51610  ' Add FormatPCLControl
51620  Set optFormatPCLControl = Controls.Add("PDFCreator.ctlOptFormatPCL", "ctlOptFormatPCL")
51630  optFormatPCLControl.Width = dmFraDescription.Width
51640  Set optFormatPCL = optFormatPCLControl.object
51650  optFormatPCL.SetLanguageStrings
51660  optFormatPCL.SetOptions
51670  ' Add FormatRAWControl
51680  Set optFormatRAWControl = Controls.Add("PDFCreator.ctlOptFormatRAW", "ctlOptFormatRAW")
51690  optFormatRAWControl.Width = dmFraDescription.Width
51700  Set optFormatRAW = optFormatRAWControl.object
51710  optFormatRAW.SetLanguageStrings
51720  optFormatRAW.SetOptions
51730  ' Add FormatSVGControl
51740  Set optFormatSVGControl = Controls.Add("PDFCreator.ctlOptFormatSVG", "ctlOptFormatSVG")
51750  optFormatSVGControl.Width = dmFraDescription.Width
51760  Set optFormatSVG = optFormatSVGControl.object
51770  optFormatSVG.SetLanguageStrings
51780  optFormatSVG.SetOptions
51790 ' ' Add FormatXCFControl - Doesn't work
51800 ' Set optFormatXCFControl = Controls.Add("PDFCreator.ctlOptFormatXCF", "ctlOptFormatXCF")
51810 ' optFormatXCFControl.Width = dmFraDescription.Width
51820 ' Set optFormatXCF = optFormatXCFControl.object
51830 ' optFormatXCF.SetLanguageStrings
51840 ' optFormatXCF.SetOptions
51850  ' Add GhostscriptControl
51860  Set optGhostscriptControl = Controls.Add("PDFCreator.ctlOptGhostscript", "ctlOptGhostscript")
51870  optGhostscriptControl.Width = dmFraDescription.Width
51880  Set optGhostscript = optGhostscriptControl.object
51890  optGhostscript.SetLanguageStrings
51900  optGhostscript.SetOptions
51910  ' Add LanguagesControl
51920  Set optLanguagesControl = Controls.Add("PDFCreator.ctlOptLanguages", "ctlOptLanguages")
51930  optLanguagesControl.Width = dmFraDescription.Width
51940  Set optLanguages = optLanguagesControl.object
51950  optLanguages.SetLanguageStrings
51960  optLanguages.SetOptions
51970  ' Add PrintControl
51980  Set optPrintControl = Controls.Add("PDFCreator.ctlOptPrint", "ctlOptPrint")
51990  optPrintControl.Width = dmFraDescription.Width
52000  Set optPrint = optPrintControl.object
52010  optPrint.SetLanguageStrings
52020  optPrint.SetOptions
52030  ' Add SaveControl
52040  '
52050  Set optSaveControl = Controls.Add("PDFCreator.ctlOptSave", "ctlOptSave")
52060  optSaveControl.Width = dmFraDescription.Width
52070  Set optSave = optSaveControl.object
52080  optSave.SetLanguageStrings
52090  optSave.SetOptions
52100  ' Add GeneralControl
52110  '
52120  Set optGeneralControl = Controls.Add("PDFCreator.ctlOptGeneral", "ctlOptGeneral")
52130  optGeneralControl.Width = dmFraDescription.Width
52140  Set optGeneral = optGeneralControl.object
52150  optGeneral.SetLanguageStrings
52160 '
52170  optGeneral.SetOptions
52180
52190  dmFraProfile.Caption = LanguageStrings.OptionsProfile
52200  cmbProfile.Clear
52210  cmbProfile.AddItem LanguageStrings.OptionsProfileDefaultName
52220
52230  Set Profiles = GetProfiles
52240  ReDim ProfileNames(Profiles.Count)
52250  ReDim ProfileOptions(Profiles.Count)
52260  ProfileNames(0) = LanguageStrings.OptionsProfileDefaultName
52270  ProfileOptions(0) = Options
52280
52290  With dmFraDescription
52300   .Caption = LanguageStrings.OptionsTreeProgram
52310   .Visible = True
52320
52330   optActionsControl.Top = .Top + .Height + ControlTop
52340   optActionsControl.Left = .Left
52350   optActionsControl.Width = .Width
52360   optAutosaveControl.Top = .Top + .Height + ControlTop
52370   optAutosaveControl.Left = .Left
52380   optAutosaveControl.Width = .Width
52390 '  optDirectoriesControl.Top = .Top + .Height + ControlTop
52400 '  optDirectoriesControl.Left = .Left
52410 '  optDirectoriesControl.Width = .Width
52420   optDocumentControl.Top = .Top + .Height + ControlTop
52430   optDocumentControl.Left = .Left
52440   optDocumentControl.Width = .Width
52450   optFontsControl.Top = .Top + .Height + ControlTop
52460   optFontsControl.Left = .Left
52470   optFontsControl.Width = .Width
52480   optFormatPNGControl.Top = .Top + .Height + ControlTop
52490   optFormatPNGControl.Left = .Left
52500   optFormatPNGControl.Width = .Width
52510   optFormatJPEGControl.Top = .Top + .Height + ControlTop
52520   optFormatJPEGControl.Left = .Left
52530   optFormatJPEGControl.Width = .Width
52540   optFormatBMPControl.Top = .Top + .Height + ControlTop
52550   optFormatBMPControl.Left = .Left
52560   optFormatBMPControl.Width = .Width
52570   optFormatPCXControl.Top = .Top + .Height + ControlTop
52580   optFormatPCXControl.Left = .Left
52590   optFormatPCXControl.Width = .Width
52600   optFormatTIFFControl.Top = .Top + .Height + ControlTop
52610   optFormatTIFFControl.Left = .Left
52620   optFormatTIFFControl.Width = .Width
52630   optFormatPDFControl.Top = .Top + .Height + ControlTop
52640   optFormatPDFControl.Left = .Left
52650   optFormatPDFControl.Width = .Width
52660   optFormatPSControl.Top = .Top + .Height + ControlTop
52670   optFormatPSControl.Left = .Left
52680   optFormatPSControl.Width = .Width
52690   optFormatEPSControl.Top = .Top + .Height + ControlTop
52700   optFormatEPSControl.Left = .Left
52710   optFormatEPSControl.Width = .Width
52720   optFormatTXTControl.Top = .Top + .Height + ControlTop
52730   optFormatTXTControl.Left = .Left
52740   optFormatTXTControl.Width = .Width
52750   optFormatPSDControl.Top = .Top + .Height + ControlTop
52760   optFormatPSDControl.Left = .Left
52770   optFormatPSDControl.Width = .Width
52780   optFormatPCLControl.Top = .Top + .Height + ControlTop
52790   optFormatPCLControl.Left = .Left
52800   optFormatPCLControl.Width = .Width
52810   optFormatRAWControl.Top = .Top + .Height + ControlTop
52820   optFormatRAWControl.Left = .Left
52830   optFormatRAWControl.Width = .Width
52840   optFormatSVGControl.Top = .Top + .Height + ControlTop
52850   optFormatSVGControl.Left = .Left
52860   optFormatSVGControl.Width = .Width
52870 '  optFormatXCFControl.Top = .Top + .Height + ControlTop
52880 '  optFormatXCFControl.Left = .Left
52890 '  optFormatXCFControl.Width = .Width
52900   optGeneralControl.Top = .Top + .Height + ControlTop
52910   optGeneralControl.Left = .Left
52920   optGeneralControl.Width = .Width
52930   optGhostscriptControl.Top = .Top + .Height + ControlTop
52940   optGhostscriptControl.Left = .Left
52950   optGhostscriptControl.Width = .Width
52960   optLanguagesControl.Top = .Top + .Height + ControlTop
52970   optLanguagesControl.Left = .Left
52980   optLanguagesControl.Width = .Width
52990   optPrintControl.Top = .Top + .Height + ControlTop
53000   optPrintControl.Left = .Left
53010   optPrintControl.Width = .Width
53020   optSaveControl.Top = .Top + .Height + ControlTop
53030   optSaveControl.Left = .Left
53040   optSaveControl.Width = .Width
53050
53060   cmdCancel.Left = .Left
53070   cmdReset.Left = .Left + (.Width - cmdReset.Width) / 2
53080   cmdSave.Left = .Left + .Width - cmdSave.Width
53090  End With
53100
53110  For i = 1 To Profiles.Count
53120   cmbProfile.AddItem Profiles(i)
53130   ProfileNames(i) = Profiles(i)
53140   ProfileOptions(i) = ReadOptions(, , Profiles(i))
53150  Next i
53160  SetProfile CurrentPrinterProfile
53170
53180  If cmbProfile.ListIndex = 0 Then
53190    optGhostscript.ControlEnabled = True
53200    optLanguages.ControlEnabled = True
53210    cmdProfileRename.Enabled = False
53220    cmdProfileDelete.Enabled = False
53230   Else
53240    optGhostscript.ControlEnabled = False
53250    optLanguages.ControlEnabled = False
53260    cmdProfileRename.Enabled = True
53270    cmdProfileDelete.Enabled = True
53280  End If
53290
53300  Set PrinterProfiles = GetPrinterProfiles
53310  For i = 1 To PrinterProfiles.Count
53320   sa(0) = PrinterProfiles(i)(0)
53330   sa(1) = PrinterProfiles(i)(1)
53340   sa(2) = PrinterProfiles(i)(1)
53350   TempPrinterProfiles.Add sa
53360  Next i
53370
53380  With LanguageStrings
53390   cmdProfileAdd.ToolTipText = .OptionsProfileAdd
53400   cmdProfileDelete.ToolTipText = .OptionsProfileDel
53410   cmdProfileRename.ToolTipText = .OptionsProfileRenameProfile
53420   cmdProfileSave.ToolTipText = .OptionsProfileSaveToDisc
53430   cmdProfileLoad.ToolTipText = .OptionsProfileLoadFromDisc
53440   cmbProfile.List(0) = .OptionsProfileDefaultName
53450  End With
53460
53470  If ShowOnlyOptions = True Then
53480   FormInTaskbar Me, True, True
53490   Caption = "PDFCreator - " & Caption
53500  End If
53510
53520  ShowAcceleratorsInForm Me, True
53530
53540  Screen.MousePointer = vbNormal
53550
53560  With Options
53570   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
53580  End With
53590
53600  LastNodeKey = ""
53610  trvOptions.Nodes("Program").Selected = True
53620  trvOptions_NodeClick trvOptions.Nodes("Program")
53630  LoadReady = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetOptions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  optActions.SetOptions
50020  optAutosave.SetOptions
50030 ' optDirectories.SetOptions
50040  optDocument.SetOptions
50050  optFonts.SetOptions
50060  optFormatPNG.SetOptions
50070  optFormatJPEG.SetOptions
50080  optFormatBMP.SetOptions
50090  optFormatPCX.SetOptions
50100  optFormatTIFF.SetOptions
50110  optFormatPDF.SetOptions
50120  optFormatPS.SetOptions
50130  optFormatEPS.SetOptions
50140  optFormatTXT.SetOptions
50150  optFormatPSD.SetOptions
50160  optFormatPCL.SetOptions
50170  optFormatRAW.SetOptions
50180  optFormatSVG.SetOptions
50190 ' optFormatXCF.SetOptions
50200  optGhostscript.SetOptions
50210  optLanguages.SetOptions
50220  optPrint.SetOptions
50230  optSave.SetOptions
50240  optGeneral.SetOptions
50250
50260  With Options
50270   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50280  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "SetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UnloadForm = True
50020  Me.Visible = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "Form_QueryUnload")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub trvOptions_NodeClick(ByVal Node As MSComctlLib.Node)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020
50030  If LastNodeKey = Node.key Or (LastNodeKey = "Program" And Node.key = "ProgramGeneral") Or (LastNodeKey = "ProgramGeneral" And Node.key = "Program") Then
50040    Exit Sub
50050   Else
50060    LastNodeKey = Node.key
50070  End If
50080
50090  optActionsControl.Visible = False
50100  optAutosaveControl.Visible = False
50110 ' optDirectoriesControl.Visible = False
50120  optDocumentControl.Visible = False
50130  optFontsControl.Visible = False
50140  optFormatPNGControl.Visible = False
50150  optFormatJPEGControl.Visible = False
50160  optFormatBMPControl.Visible = False
50170  optFormatPCXControl.Visible = False
50180  optFormatTIFFControl.Visible = False
50190  optFormatPDFControl.Visible = False
50200  optFormatPSControl.Visible = False
50210  optFormatEPSControl.Visible = False
50220  optFormatTXTControl.Visible = False
50230  optFormatPSDControl.Visible = False
50240  optFormatPCLControl.Visible = False
50250  optFormatRAWControl.Visible = False
50260  optFormatSVGControl.Visible = False
50270 ' optFormatXCFControl.Visible = False
50280  optGeneralControl.Visible = False
50290  optGhostscriptControl.Visible = False
50300  optLanguagesControl.Visible = False
50310  optPrintControl.Visible = False
50320  optSaveControl.Visible = False
50330
50341  Select Case UCase$(Node.key)
        Case "PROGRAM", "PROGRAMGENERAL"
50360    Set picOptions = LoadResPicture(2101, vbResIcon)
50370    lblOptions.Caption = LanguageStrings.OptionsProgramGeneralDescription
50380    optGeneralControl.Visible = True
50390    dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50400   Case "PROGRAMGHOSTSCRIPT"
50410    Set picOptions = LoadResPicture(2119, vbResIcon)
50420    lblOptions.Caption = LanguageStrings.OptionsProgramGhostscriptDescription
50430    optGhostscriptControl.Visible = True
50440    dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50450   Case "PROGRAMDOCUMENT"
50460    Set picOptions = LoadResPicture(2105, vbResIcon)
50470    lblOptions.Caption = LanguageStrings.OptionsProgramDocumentDescription
50480    optDocumentControl.Visible = True
50490    dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50500   Case "PROGRAMSAVE"
50510    Set picOptions = LoadResPicture(2106, vbResIcon)
50520    lblOptions.Caption = LanguageStrings.OptionsProgramSaveDescription
50530
50540    optSaveControl.Visible = True
50550    dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50560   Case "PROGRAMAUTOSAVE"
50570    Set picOptions = LoadResPicture(2103, vbResIcon)
50580    lblOptions.Caption = LanguageStrings.OptionsProgramAutosaveDescription
50590    optAutosaveControl.Visible = True
50600    dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50610 '  Case "PROGRAMDIRECTORIES"
50620 '   Set picOptions = LoadResPicture(2104, vbResIcon)
50630 '   lblOptions.Caption = LanguageStrings.OptionsProgramDirectoriesDescription
50640 '   optDirectoriesControl.Visible = True
50650 '   dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50660   Case "PROGRAMACTIONS"
50670    Set picOptions = LoadResPicture(2121, vbResIcon)
50680    lblOptions.Caption = LanguageStrings.OptionsProgramActionsDescription
50690    optActionsControl.Visible = True
50700    dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50710   Case "PROGRAMPRINT"
50720    Set picOptions = LoadResPicture(2122, vbResIcon)
50730    lblOptions.Caption = LanguageStrings.OptionsProgramPrintDescription
50740    optPrintControl.Visible = True
50750    dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50760   Case "PROGRAMFONTS"
50770    Set picOptions = LoadResPicture(2102, vbResIcon)
50780    lblOptions.Caption = LanguageStrings.OptionsProgramFontDescription
50790    optFontsControl.Visible = True
50800    dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50810   Case "PROGRAMLANGUAGES"
50820    Set picOptions = LoadResPicture(2123, vbResIcon)
50830    lblOptions.Caption = LanguageStrings.OptionsProgramLanguagesDescription
50840    optLanguagesControl.Visible = True
50850    dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50860   Case "FORMATS", "FORMATSPDF"
50870    Set picOptions = LoadResPicture(2111, vbResIcon)
50880    lblOptions.Caption = LanguageStrings.OptionsPDFDescription
50890    optFormatPDFControl.Visible = True
50900    dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
50910   Case "FORMATSPNG"
50920    Set picOptions = LoadResPicture(2112, vbResIcon)
50930    lblOptions.Caption = LanguageStrings.OptionsPNGDescription
50940    optFormatPNGControl.Visible = True
50950    dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
50960   Case "FORMATSJPEG"
50970    Set picOptions = LoadResPicture(2113, vbResIcon)
50980    lblOptions.Caption = LanguageStrings.OptionsJPEGDescription
50990    optFormatJPEGControl.Visible = True
51000    dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51010   Case "FORMATSBMP"
51020    Set picOptions = LoadResPicture(2114, vbResIcon)
51030    lblOptions.Caption = LanguageStrings.OptionsBMPDescription
51040    optFormatBMPControl.Visible = True
51050    dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51060   Case "FORMATSPCX"
51070    Set picOptions = LoadResPicture(2115, vbResIcon)
51080    lblOptions.Caption = LanguageStrings.OptionsPCXDescription
51090    optFormatPCXControl.Visible = True
51100    dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51110   Case "FORMATSTIFF"
51120    Set picOptions = LoadResPicture(2116, vbResIcon)
51130    lblOptions.Caption = LanguageStrings.OptionsTIFFDescription
51140    optFormatTIFFControl.Visible = True
51150    dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51160   Case "FORMATSPS"
51170    Set picOptions = LoadResPicture(2117, vbResIcon)
51180    lblOptions.Caption = LanguageStrings.OptionsPSDescription
51190    optFormatPSControl.Visible = True
51200    dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51210   Case "FORMATSEPS"
51220    Set picOptions = LoadResPicture(2118, vbResIcon)
51230    lblOptions.Caption = LanguageStrings.OptionsEPSDescription
51240    optFormatEPSControl.Visible = True
51250    dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51260   Case "FORMATSTXT"
51270    Set picOptions = LoadResPicture(2124, vbResIcon)
51280    lblOptions.Caption = LanguageStrings.OptionsTXTDescription
51290    optFormatTXTControl.Visible = True
51300    dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51310   Case "FORMATSPSD"
51320    Set picOptions = LoadResPicture(2125, vbResIcon)
51330    lblOptions.Caption = LanguageStrings.OptionsPSDDescription
51340    optFormatPSDControl.Visible = True
51350    dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51360   Case "FORMATSPCL"
51370    Set picOptions = LoadResPicture(2126, vbResIcon)
51380    lblOptions.Caption = LanguageStrings.OptionsPCLDescription
51390    optFormatPCLControl.Visible = True
51400    dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51410   Case "FORMATSRAW"
51420    Set picOptions = LoadResPicture(2127, vbResIcon)
51430    lblOptions.Caption = LanguageStrings.OptionsRAWDescription
51440    optFormatRAWControl.Visible = True
51450    dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51460   Case "FORMATSSVG"
51470    Set picOptions = LoadResPicture(2129, vbResIcon)
51480    lblOptions.Caption = LanguageStrings.OptionsSVGDescription
51490    optFormatSVGControl.Visible = True
51500    dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51510 '  Case "FORMATSXCF"
51520 '   Set picOptions = LoadResPicture(2128, vbResIcon)
51530 '   lblOptions.Caption = LanguageStrings.OptionsXCFDescription
51540 '   optFormatXCFControl.Visible = True
51550 '   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51560  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "trvOptions_NodeClick")
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
50010  optActions.SetFrames OptionsDesign
50020  optAutosave.SetFrames OptionsDesign
50030 ' optDirectories.SetFrames OptionsDesign
50040  optDocument.SetFrames OptionsDesign
50050  optFonts.SetFrames OptionsDesign
50060  optFormatBMP.SetFrames OptionsDesign
50070  optFormatEPS.SetFrames OptionsDesign
50080  optFormatJPEG.SetFrames OptionsDesign
50090  optFormatPCL.SetFrames OptionsDesign
50100  optFormatPCX.SetFrames OptionsDesign
50110  optFormatPDF.SetFrames OptionsDesign
50120  optFormatPNG.SetFrames OptionsDesign
50130  optFormatPS.SetFrames OptionsDesign
50140  optFormatPSD.SetFrames OptionsDesign
50150  optFormatRAW.SetFrames OptionsDesign
50160  optFormatSVG.SetFrames OptionsDesign
50170  optFormatTIFF.SetFrames OptionsDesign
50180  optFormatTXT.SetFrames OptionsDesign
50190  optGeneral.SetFrames OptionsDesign
50200  optGhostscript.SetFrames OptionsDesign
50210  optLanguages.SetFrames OptionsDesign
50220  optPrint.SetFrames OptionsDesign
50230  optSave.SetFrames OptionsDesign
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "SetFrames")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetProfile(Optional ByVal ProfileName As String = "") ' Empty profilename for default profile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  If cmbProfile.ListIndex < 0 Then
50030    OldProfile = 0
50040   Else
50050    OldProfile = cmbProfile.ListIndex
50060  End If
50070  ProfileName = Trim$(ProfileName)
50080  If LenB(ProfileName) = 0 Then
50090   cmbProfile.ListIndex = 0
50100   Exit Sub
50110  End If
50120  For i = 1 To cmbProfile.ListCount - 1
50130   If LCase$(ProfileName) = LCase$(cmbProfile.List(i)) Then
50140    cmbProfile.ListIndex = i
50150    Exit Sub
50160   End If
50170  Next i
50180  cmbProfile.ListIndex = 0
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "SetProfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
