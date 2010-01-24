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
      _extentx        =   11324
      _extenty        =   1879
      caption         =   ""
      barcolorfrom    =   723949
      barcolorto      =   132452
      font            =   "frmOptions.frx":000C
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
      _extentx        =   16193
      _extenty        =   1879
      caption         =   "Profil"
      barcolorfrom    =   723949
      barcolorto      =   132452
      font            =   "frmOptions.frx":0038
      Begin VB.CommandButton cmdProfileRename 
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         Picture         =   "frmOptions.frx":0064
         Style           =   1  'Grafisch
         TabIndex        =   12
         ToolTipText     =   "Rename profile"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdProfileLoad 
         Height          =   375
         Left            =   8640
         Picture         =   "frmOptions.frx":0458
         Style           =   1  'Grafisch
         TabIndex        =   11
         ToolTipText     =   "Load profile"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdProfileSave 
         Height          =   375
         Left            =   8160
         Picture         =   "frmOptions.frx":0856
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
         Picture         =   "frmOptions.frx":0BEB
         Style           =   1  'Grafisch
         TabIndex        =   9
         ToolTipText     =   "Delete profile"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdProfileAdd 
         Height          =   375
         Left            =   6720
         Picture         =   "frmOptions.frx":0FE3
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
            Picture         =   "frmOptions.frx":13D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":152D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1AC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":2061
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":25FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":2995
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":2F2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":3809
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":3DA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":433D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":48D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":4E71
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":540B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":59A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":5F3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":64D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":6A73
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":700D
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":75A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":7E81
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":875B
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":8CF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":928F
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":9829
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":9DC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":A35D
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
'Private optFontsControl As VBControlExtender, optFonts As ctlOptFonts
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

Private Function ProfileAssociatedPrinter(ProfileName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PrinterProfiles As Collection, p As Variant, i As Long, tStr As String
50020  Set PrinterProfiles = GetPrinterProfiles
50030
50040  For i = 1 To PrinterProfiles.Count
50050   If StrComp(PrinterProfiles(i)(1), ProfileName, vbTextCompare) = 0 Then
50060    If LenB(tStr) = 0 Then
50070      tStr = PrinterProfiles(i)(0)
50080     Else
50090      tStr = tStr & ", " & PrinterProfiles(i)(0)
50100    End If
50110   End If
50120  Next i
50130  ProfileAssociatedPrinter = tStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ProfileAssociatedPrinter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub cmdProfileDelete_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim aw As Long, tStr As String, CurrentProfile As String, i As Long
50020  If cmbProfile.ListCount <= 1 Then
50030   Exit Sub
50040  End If
50050  CurrentProfile = cmbProfile.List(cmbProfile.ListIndex)
50060  tStr = ProfileAssociatedPrinter(CurrentProfile)
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
50030   Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50041   Select Case trvOptions.SelectedItem.key
         Case "Program", "ProgramGeneral"
50060     Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50070    Case "ProgramGhostscript"
50080     Call HTMLHelp_ShowTopic("html\ghostscript.htm")
50090    Case "Program"
50100     Call HTMLHelp_ShowTopic("html\docproperties.htm")
50110    Case "Program"
50120     Call HTMLHelp_ShowTopic("html\savesettings.htm")
50130    Case "Program"
50140     Call HTMLHelp_ShowTopic("html\autosave.htm")
50150    Case "Formats", "FormatsPDF"
50160     If trvOptions.SelectedItem.key = "FormatsPDF" Then
50171       Select Case optFormatPDF.PDFOptionsIndex
             Case 1
50190         Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50200        Case 2
50210         Call HTMLHelp_ShowTopic("html\pdfcompression.htm")
50220        Case 3
50230         Call HTMLHelp_ShowTopic("html\pdffonts.htm")
50240        Case 4
50250         Call HTMLHelp_ShowTopic("html\pdfcolors.htm")
50260        Case 5
50270         Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
50280        Case 6
50290         Call HTMLHelp_ShowTopic("html\pdfsigning.htm")
50300       Case Else
50310        Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50320       End Select
50330      Else
50340       Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50350     End If
50360    Case "FormatsPNG"
50370     Call HTMLHelp_ShowTopic("html\pngsettings.htm")
50380    Case "FormatsJPEG"
50390     Call HTMLHelp_ShowTopic("html\jpegsettings.htm")
50400    Case "FormatsBMP"
50410     Call HTMLHelp_ShowTopic("html\bmpsettings.htm")
50420    Case "FormatsPCX"
50430     Call HTMLHelp_ShowTopic("html\pcxsettings.htm")
50440    Case "FormatsTIFF"
50450     Call HTMLHelp_ShowTopic("html\tiffsettings.htm")
50460    Case "FormatsPS"
50470     Call HTMLHelp_ShowTopic("html\pssettings.htm")
50480    Case "FormatsEPS"
50490     Call HTMLHelp_ShowTopic("html\epssettings.htm")
50500    Case "FormatsTXT"
50510     Call HTMLHelp_ShowTopic("html\txtsettings.htm")
50520    Case "FormatsPSD"
50530     Call HTMLHelp_ShowTopic("html\psdsettings.htm")
50540    Case "FormatsPCL"
50550     Call HTMLHelp_ShowTopic("html\pclsettings.htm")
50560    Case "FormatsRAW"
50570     Call HTMLHelp_ShowTopic("html\rawsettings.htm")
50580    Case "FormatsSVG"
50590     Call HTMLHelp_ShowTopic("html\svgsettings.htm")
50600    Case Else
50610     Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50620    End Select
50630  End If
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
50280   trvOptions.Nodes("ProgramLanguages").Text = .OptionsProgramLanguagesSymbol
50290
50300   trvOptions.Nodes("Formats").Text = .OptionsTreeFormats
50310   trvOptions.Nodes("FormatsPDF").Text = .OptionsPDFSymbol
50320   trvOptions.Nodes("FormatsPNG").Text = .OptionsPNGSymbol
50330   trvOptions.Nodes("FormatsJPEG").Text = .OptionsJPEGSymbol
50340   trvOptions.Nodes("FormatsBMP").Text = .OptionsBMPSymbol
50350   trvOptions.Nodes("FormatsPCX").Text = .OptionsPCXSymbol
50360   trvOptions.Nodes("FormatsTIFF").Text = .OptionsTIFFSymbol
50370   trvOptions.Nodes("FormatsPS").Text = .OptionsPSSymbol
50380   trvOptions.Nodes("FormatsEPS").Text = .OptionsEPSSymbol
50390   trvOptions.Nodes("FormatsTXT").Text = .OptionsTXTSymbol
50400   trvOptions.Nodes("FormatsPSD").Text = .OptionsPSDSymbol
50410   trvOptions.Nodes("FormatsPCL").Text = .OptionsPCLSymbol
50420   trvOptions.Nodes("FormatsRAW").Text = .OptionsRAWSymbol
50430   trvOptions.Nodes("FormatsSVG").Text = .OptionsSVGSymbol
50440
50450   lblOptions.Caption = .OptionsProgramLanguagesDescription
50460  End With
50470  optActions.SetLanguageStrings
50480  optAutosave.SetLanguageStrings
50490 ' optDirectories.SetLanguageStrings
50500  optDocument.SetLanguageStrings
50510 ' optFonts.SetLanguageStrings
50520  optFormatPNG.SetLanguageStrings
50530  optFormatJPEG.SetLanguageStrings
50540  optFormatBMP.SetLanguageStrings
50550  optFormatPCX.SetLanguageStrings
50560  optFormatTIFF.SetLanguageStrings
50570  optFormatPDF.SetLanguageStrings
50580  optFormatPS.SetLanguageStrings
50590  optFormatEPS.SetLanguageStrings
50600  optFormatTXT.SetLanguageStrings
50610  optFormatPSD.SetLanguageStrings
50620  optFormatPCL.SetLanguageStrings
50630  optFormatRAW.SetLanguageStrings
50640  optFormatSVG.SetLanguageStrings
50650 ' optFormatXCF.SetLanguageStrings
50660
50670  optGeneral.SetLanguageStrings
50680  optGhostscript.SetLanguageStrings
50690  optLanguages.SetLanguageStrings
50700  optPrint.SetLanguageStrings
50710  optSave.SetLanguageStrings
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
50060  'optFonts.GetOptions
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
50040  tRestart = False
50050  If UCase$(Options.DirectoryGhostscriptBinaries) <> UCase$(ProfileOptions(0).DirectoryGhostscriptBinaries) Then
50060   tRestart = True
50070  End If
50080
50090  ' Save all Options/Profiles
50100  ProfileOptions(cmbProfile.ListIndex) = GetOptionsFromUserControls(ProfileOptions(cmbProfile.ListIndex)) ' Get the current settings
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
50390   trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramPrint", .OptionsProgramPrintSymbol, 9
50400   trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramLanguages", .OptionsProgramLanguagesSymbol, 10
50410   Set nodeFormats = trvOptions.Nodes.Add(, , "Formats", .OptionsTreeFormats, 11)
50420   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPDF", .OptionsPDFSymbol, 11
50430   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPNG", .OptionsPNGSymbol, 12
50440   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsJPEG", .OptionsJPEGSymbol, 13
50450   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsBMP", .OptionsBMPSymbol, 14
50460   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPCX", .OptionsPCXSymbol, 15
50470   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsTIFF", .OptionsTIFFSymbol, 16
50480   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPS", .OptionsPSSymbol, 17
50490   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsEPS", .OptionsEPSSymbol, 18
50500   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsTXT", .OptionsTXTSymbol, 21
50510   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPSD", .OptionsPSDSymbol, 22
50520   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPCL", .OptionsPCLSymbol, 23
50530   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsRAW", .OptionsRAWSymbol, 24
50540   trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsSVG", .OptionsSVGSymbol, 26
50550  End With
50560  nodeProgram.Expanded = True
50570  nodeFormats.Expanded = True
50580
50590  With LanguageStrings
50600   Set picOptions = LoadResPicture(2101, vbResIcon)
50610   Me.Caption = .DialogPrinterOptions
50620   cmdCancel.Caption = .OptionsCancel
50630   cmdReset.Caption = .OptionsReset
50640   cmdSave.Caption = .OptionsSave
50650  End With
50660
50670  SetFrame dmFraDescription
50680  SetFrame dmFraProfile
50690
50700  ' Add ActionsControl
50710  Set optActionsControl = Controls.Add("PDFCreator.ctlOptActions", "ctlOptActions")
50720  optActionsControl.Width = dmFraDescription.Width
50730  Set optActions = optActionsControl.object
50740  optActions.SetLanguageStrings
50750  optActions.SetOptions
50760  ' Add AutosaveControl
50770  Set optAutosaveControl = Controls.Add("PDFCreator.ctlOptAutosave", "ctlOptAutosave")
50780  optAutosaveControl.Width = dmFraDescription.Width
50790  Set optAutosave = optAutosaveControl.object
50800  optAutosave.SetLanguageStrings
50810  optAutosave.SetOptions
50820  ' Add DirectoriesControl
50830 ' Set optDirectoriesControl = Controls.Add("PDFCreator.ctlOptDirectories", "ctlOptDirectories")
50840 ' optDirectoriesControl.Width = dmFraDescription.Width
50850 ' Set optDirectories = optDirectoriesControl.object
50860 ' optDirectories.SetLanguageStrings
50870 ' optDirectories.SetOptions
50880  ' Add DocumentControl
50890  Set optDocumentControl = Controls.Add("PDFCreator.ctlOptDocument", "ctlOptDocument")
50900  optDocumentControl.Width = dmFraDescription.Width
50910  Set optDocument = optDocumentControl.object
50920  optDocument.SetLanguageStrings
50930  optDocument.SetOptions
50940  ' Add FontsControl
50950 ' Set optFontsControl = Controls.Add("PDFCreator.ctlOptFonts", "ctlOptFonts")
50960 ' optFontsControl.Width = dmFraDescription.Width
50970 ' Set optFonts = optFontsControl.object
50980 ' optFonts.SetLanguageStrings
50990 ' optFonts.SetOptions
51000  ' Add FormatPNGControl
51010  Set optFormatPNGControl = Controls.Add("PDFCreator.ctlOptFormatPNG", "ctlOptFormatPNG")
51020  optFormatPNGControl.Width = dmFraDescription.Width
51030  Set optFormatPNG = optFormatPNGControl.object
51040  optFormatPNG.SetLanguageStrings
51050  optFormatPNG.SetOptions
51060  ' Add FormatJPEQControl
51070  Set optFormatJPEGControl = Controls.Add("PDFCreator.ctlOptFormatJPEG", "ctlOptFormatJPEG")
51080  optFormatJPEGControl.Width = dmFraDescription.Width
51090  Set optFormatJPEG = optFormatJPEGControl.object
51100  optFormatJPEG.SetLanguageStrings
51110  optFormatJPEG.SetOptions
51120  ' Add FormatBMPControl
51130  Set optFormatBMPControl = Controls.Add("PDFCreator.ctlOptFormatBMP", "ctlOptFormatBMP")
51140  optFormatBMPControl.Width = dmFraDescription.Width
51150  Set optFormatBMP = optFormatBMPControl.object
51160  optFormatBMP.SetLanguageStrings
51170  optFormatBMP.SetOptions
51180  ' Add FormatPCXControl
51190  Set optFormatPCXControl = Controls.Add("PDFCreator.ctlOptFormatPCX", "ctlOptFormatPCX")
51200  optFormatPCXControl.Width = dmFraDescription.Width
51210  Set optFormatPCX = optFormatPCXControl.object
51220  optFormatPCX.SetLanguageStrings
51230  optFormatPCX.SetOptions
51240  ' Add FormatTIFFControl
51250  Set optFormatTIFFControl = Controls.Add("PDFCreator.ctlOptFormatTIFF", "ctlOptFormatTIFF")
51260  optFormatTIFFControl.Width = dmFraDescription.Width
51270  Set optFormatTIFF = optFormatTIFFControl.object
51280  optFormatTIFF.SetLanguageStrings
51290  optFormatTIFF.SetOptions
51300  ' Add FormatPDFControl
51310  Set optFormatPDFControl = Controls.Add("PDFCreator.ctlOptFormatPDF", "ctlOptFormatPDF")
51320  optFormatPDFControl.Width = dmFraDescription.Width
51330  Set optFormatPDF = optFormatPDFControl.object
51340  optFormatPDF.SetLanguageStrings
51350  optFormatPDF.SetOptions
51360  ' Add FormatPS
51370  Set optFormatPSControl = Controls.Add("PDFCreator.ctlOptFormatPS", "ctlOptFormatPS")
51380  optFormatPSControl.Width = dmFraDescription.Width
51390  Set optFormatPS = optFormatPSControl.object
51400  optFormatPS.SetLanguageStrings
51410  optFormatPS.SetOptions
51420  ' Add FormatEPSControl
51430  Set optFormatEPSControl = Controls.Add("PDFCreator.ctlOptFormatEPS", "ctlOptFormatEPS")
51440  optFormatEPSControl.Width = dmFraDescription.Width
51450  Set optFormatEPS = optFormatEPSControl.object
51460  optFormatEPS.SetLanguageStrings
51470  optFormatEPS.SetOptions
51480  ' Add FormatTXTControl
51490  Set optFormatTXTControl = Controls.Add("PDFCreator.ctlOptFormatTXT", "ctlOptFormatTXT")
51500  optFormatTXTControl.Width = dmFraDescription.Width
51510  Set optFormatTXT = optFormatTXTControl.object
51520  optFormatTXT.SetLanguageStrings
51530  optFormatTXT.SetOptions
51540  ' Add FormatPSDControl
51550  Set optFormatPSDControl = Controls.Add("PDFCreator.ctlOptFormatPSD", "ctlOptFormatPSD")
51560  optFormatPSDControl.Width = dmFraDescription.Width
51570  Set optFormatPSD = optFormatPSDControl.object
51580  optFormatPSD.SetLanguageStrings
51590  optFormatPSD.SetOptions
51600  ' Add FormatPCLControl
51610  Set optFormatPCLControl = Controls.Add("PDFCreator.ctlOptFormatPCL", "ctlOptFormatPCL")
51620  optFormatPCLControl.Width = dmFraDescription.Width
51630  Set optFormatPCL = optFormatPCLControl.object
51640  optFormatPCL.SetLanguageStrings
51650  optFormatPCL.SetOptions
51660  ' Add FormatRAWControl
51670  Set optFormatRAWControl = Controls.Add("PDFCreator.ctlOptFormatRAW", "ctlOptFormatRAW")
51680  optFormatRAWControl.Width = dmFraDescription.Width
51690  Set optFormatRAW = optFormatRAWControl.object
51700  optFormatRAW.SetLanguageStrings
51710  optFormatRAW.SetOptions
51720  ' Add FormatSVGControl
51730  Set optFormatSVGControl = Controls.Add("PDFCreator.ctlOptFormatSVG", "ctlOptFormatSVG")
51740  optFormatSVGControl.Width = dmFraDescription.Width
51750  Set optFormatSVG = optFormatSVGControl.object
51760  optFormatSVG.SetLanguageStrings
51770  optFormatSVG.SetOptions
51780 ' ' Add FormatXCFControl - Doesn't work
51790 ' Set optFormatXCFControl = Controls.Add("PDFCreator.ctlOptFormatXCF", "ctlOptFormatXCF")
51800 ' optFormatXCFControl.Width = dmFraDescription.Width
51810 ' Set optFormatXCF = optFormatXCFControl.object
51820 ' optFormatXCF.SetLanguageStrings
51830 ' optFormatXCF.SetOptions
51840  ' Add GhostscriptControl
51850  Set optGhostscriptControl = Controls.Add("PDFCreator.ctlOptGhostscript", "ctlOptGhostscript")
51860  optGhostscriptControl.Width = dmFraDescription.Width
51870  Set optGhostscript = optGhostscriptControl.object
51880  optGhostscript.SetLanguageStrings
51890  optGhostscript.SetOptions
51900  ' Add LanguagesControl
51910  Set optLanguagesControl = Controls.Add("PDFCreator.ctlOptLanguages", "ctlOptLanguages")
51920  optLanguagesControl.Width = dmFraDescription.Width
51930  Set optLanguages = optLanguagesControl.object
51940  optLanguages.SetLanguageStrings
51950  optLanguages.SetOptions
51960  ' Add PrintControl
51970  Set optPrintControl = Controls.Add("PDFCreator.ctlOptPrint", "ctlOptPrint")
51980  optPrintControl.Width = dmFraDescription.Width
51990  Set optPrint = optPrintControl.object
52000  optPrint.SetLanguageStrings
52010  optPrint.SetOptions
52020  ' Add SaveControl
52030  '
52040  Set optSaveControl = Controls.Add("PDFCreator.ctlOptSave", "ctlOptSave")
52050  optSaveControl.Width = dmFraDescription.Width
52060  Set optSave = optSaveControl.object
52070  optSave.SetLanguageStrings
52080  optSave.SetOptions
52090  ' Add GeneralControl
52100  '
52110  Set optGeneralControl = Controls.Add("PDFCreator.ctlOptGeneral", "ctlOptGeneral")
52120  optGeneralControl.Width = dmFraDescription.Width
52130  Set optGeneral = optGeneralControl.object
52140  optGeneral.SetLanguageStrings
52150 '
52160  optGeneral.SetOptions
52170
52180  dmFraProfile.Caption = LanguageStrings.OptionsProfile
52190  cmbProfile.Clear
52200  cmbProfile.AddItem LanguageStrings.OptionsProfileDefaultName
52210
52220  Set Profiles = GetProfiles
52230  ReDim ProfileNames(Profiles.Count)
52240  ReDim ProfileOptions(Profiles.Count)
52250  ProfileNames(0) = LanguageStrings.OptionsProfileDefaultName
52260  ProfileOptions(0) = Options
52270
52280  With dmFraDescription
52290   .Caption = LanguageStrings.OptionsTreeProgram
52300   .Visible = True
52310
52320   optActionsControl.Top = .Top + .Height + ControlTop
52330   optActionsControl.Left = .Left
52340   optActionsControl.Width = .Width
52350   optAutosaveControl.Top = .Top + .Height + ControlTop
52360   optAutosaveControl.Left = .Left
52370   optAutosaveControl.Width = .Width
52380 '  optDirectoriesControl.Top = .Top + .Height + ControlTop
52390 '  optDirectoriesControl.Left = .Left
52400 '  optDirectoriesControl.Width = .Width
52410   optDocumentControl.Top = .Top + .Height + ControlTop
52420   optDocumentControl.Left = .Left
52430   optDocumentControl.Width = .Width
52440 '  optFontsControl.Top = .Top + .Height + ControlTop
52450 '  optFontsControl.Left = .Left
52460 '  optFontsControl.Width = .Width
52470   optFormatPNGControl.Top = .Top + .Height + ControlTop
52480   optFormatPNGControl.Left = .Left
52490   optFormatPNGControl.Width = .Width
52500   optFormatJPEGControl.Top = .Top + .Height + ControlTop
52510   optFormatJPEGControl.Left = .Left
52520   optFormatJPEGControl.Width = .Width
52530   optFormatBMPControl.Top = .Top + .Height + ControlTop
52540   optFormatBMPControl.Left = .Left
52550   optFormatBMPControl.Width = .Width
52560   optFormatPCXControl.Top = .Top + .Height + ControlTop
52570   optFormatPCXControl.Left = .Left
52580   optFormatPCXControl.Width = .Width
52590   optFormatTIFFControl.Top = .Top + .Height + ControlTop
52600   optFormatTIFFControl.Left = .Left
52610   optFormatTIFFControl.Width = .Width
52620   optFormatPDFControl.Top = .Top + .Height + ControlTop
52630   optFormatPDFControl.Left = .Left
52640   optFormatPDFControl.Width = .Width
52650   optFormatPSControl.Top = .Top + .Height + ControlTop
52660   optFormatPSControl.Left = .Left
52670   optFormatPSControl.Width = .Width
52680   optFormatEPSControl.Top = .Top + .Height + ControlTop
52690   optFormatEPSControl.Left = .Left
52700   optFormatEPSControl.Width = .Width
52710   optFormatTXTControl.Top = .Top + .Height + ControlTop
52720   optFormatTXTControl.Left = .Left
52730   optFormatTXTControl.Width = .Width
52740   optFormatPSDControl.Top = .Top + .Height + ControlTop
52750   optFormatPSDControl.Left = .Left
52760   optFormatPSDControl.Width = .Width
52770   optFormatPCLControl.Top = .Top + .Height + ControlTop
52780   optFormatPCLControl.Left = .Left
52790   optFormatPCLControl.Width = .Width
52800   optFormatRAWControl.Top = .Top + .Height + ControlTop
52810   optFormatRAWControl.Left = .Left
52820   optFormatRAWControl.Width = .Width
52830   optFormatSVGControl.Top = .Top + .Height + ControlTop
52840   optFormatSVGControl.Left = .Left
52850   optFormatSVGControl.Width = .Width
52860 '  optFormatXCFControl.Top = .Top + .Height + ControlTop
52870 '  optFormatXCFControl.Left = .Left
52880 '  optFormatXCFControl.Width = .Width
52890   optGeneralControl.Top = .Top + .Height + ControlTop
52900   optGeneralControl.Left = .Left
52910   optGeneralControl.Width = .Width
52920   optGhostscriptControl.Top = .Top + .Height + ControlTop
52930   optGhostscriptControl.Left = .Left
52940   optGhostscriptControl.Width = .Width
52950   optLanguagesControl.Top = .Top + .Height + ControlTop
52960   optLanguagesControl.Left = .Left
52970   optLanguagesControl.Width = .Width
52980   optPrintControl.Top = .Top + .Height + ControlTop
52990   optPrintControl.Left = .Left
53000   optPrintControl.Width = .Width
53010   optSaveControl.Top = .Top + .Height + ControlTop
53020   optSaveControl.Left = .Left
53030   optSaveControl.Width = .Width
53040
53050   cmdCancel.Left = .Left
53060   cmdReset.Left = .Left + (.Width - cmdReset.Width) / 2
53070   cmdSave.Left = .Left + .Width - cmdSave.Width
53080  End With
53090
53100  For i = 1 To Profiles.Count
53110   cmbProfile.AddItem Profiles(i)
53120   ProfileNames(i) = Profiles(i)
53130   ProfileOptions(i) = ReadOptions(, , Profiles(i))
53140  Next i
53150  SetProfile CurrentPrinterProfile
53160
53170  If cmbProfile.ListIndex = 0 Then
53180    optGhostscript.ControlEnabled = True
53190    optLanguages.ControlEnabled = True
53200    cmdProfileRename.Enabled = False
53210    cmdProfileDelete.Enabled = False
53220   Else
53230    optGhostscript.ControlEnabled = False
53240    optLanguages.ControlEnabled = False
53250    cmdProfileRename.Enabled = True
53260    cmdProfileDelete.Enabled = True
53270  End If
53280
53290  Set PrinterProfiles = GetPrinterProfiles
53300  For i = 1 To PrinterProfiles.Count
53310   sa(0) = PrinterProfiles(i)(0)
53320   sa(1) = PrinterProfiles(i)(1)
53330   sa(2) = PrinterProfiles(i)(1)
53340   TempPrinterProfiles.Add sa
53350  Next i
53360
53370  With LanguageStrings
53380   cmdProfileAdd.ToolTipText = .OptionsProfileAdd
53390   cmdProfileDelete.ToolTipText = .OptionsProfileDel
53400   cmdProfileRename.ToolTipText = .OptionsProfileRenameProfile
53410   cmdProfileSave.ToolTipText = .OptionsProfileSaveToDisc
53420   cmdProfileLoad.ToolTipText = .OptionsProfileLoadFromDisc
53430   cmbProfile.List(0) = .OptionsProfileDefaultName
53440  End With
53450
53460  If ShowOnlyOptions = True Then
53470   FormInTaskbar Me, True, True
53480   Caption = "PDFCreator - " & Caption
53490  End If
53500
53510  ShowAcceleratorsInForm Me, True
53520
53530  Screen.MousePointer = vbNormal
53540
53550  With Options
53560   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
53570  End With
53580
53590  LastNodeKey = ""
53600  trvOptions.Nodes("Program").Selected = True
53610  trvOptions_NodeClick trvOptions.Nodes("Program")
53620  LoadReady = True
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
50050 ' optFonts.SetOptions
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
50130 ' optFontsControl.Visible = False
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
50760 '  Case "PROGRAMFONTS"
50770 '   Set picOptions = LoadResPicture(2102, vbResIcon)
50780 '   lblOptions.Caption = LanguageStrings.OptionsProgramFontDescription
50790 '   optFontsControl.Visible = True
50800 '   dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50810   Case "PROGRAMLANGUAGE"
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
50050 ' optFonts.SetFrames OptionsDesign
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
