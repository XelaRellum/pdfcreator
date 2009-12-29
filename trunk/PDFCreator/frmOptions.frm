VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
   Begin PDFCreator.isExplorerBar ieb 
      Height          =   7875
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   13891
      FontName        =   "MS Sans Serif"
      FontCharset     =   0
      Begin MSComctlLib.ImageList imlIeb 
         Left            =   0
         Top             =   0
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
               Picture         =   "frmOptions.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":0166
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":0700
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":0C9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":1234
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":15CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":1B68
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":2442
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":29DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":2F76
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":3510
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":3AAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":4044
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":45DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":4B78
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":5112
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":56AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":5C46
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":61E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":6ABA
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":7394
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":792E
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":7EC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":8462
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":89FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":8F96
               Key             =   ""
            EndProperty
         EndProperty
      End
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
      TabIndex        =   7
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
         Picture         =   "frmOptions.frx":9530
         Style           =   1  'Grafisch
         TabIndex        =   13
         ToolTipText     =   "Rename profile"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdProfileLoad 
         Height          =   375
         Left            =   8640
         Picture         =   "frmOptions.frx":9924
         Style           =   1  'Grafisch
         TabIndex        =   12
         ToolTipText     =   "Load profile"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdProfileSave 
         Height          =   375
         Left            =   8160
         Picture         =   "frmOptions.frx":9D22
         Style           =   1  'Grafisch
         TabIndex        =   11
         ToolTipText     =   "Save profile"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdProfileDelete 
         Enabled         =   0   'False
         Height          =   375
         Left            =   7680
         Picture         =   "frmOptions.frx":A0B7
         Style           =   1  'Grafisch
         TabIndex        =   10
         ToolTipText     =   "Delete profile"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdProfileAdd 
         Height          =   375
         Left            =   6720
         Picture         =   "frmOptions.frx":A4AF
         Style           =   1  'Grafisch
         TabIndex        =   9
         ToolTipText     =   "Add profile"
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox cmbProfile 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   8
         Top             =   480
         Width           =   6375
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private UnloadForm As Boolean, LoadReady As Boolean, oldLanguage As String, Languages As Collection
Private LastProgramGroup As String, LastFormatGroup As String
 
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

Private OldProfile As Long, ProfileOptions() As tOptions, ProfileNames() As String, TempPrinterProfiles As Collection

Private Sub cmbProfile_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If cmbProfile.ListIndex <> OldProfile Then
50020   If OldProfile <= UBound(ProfileOptions) Then
50030    ProfileOptions(OldProfile) = GetOptionsFromUserControls
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

Public Sub AddProfile(Profilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim resS As String, i As Long
50020  resS = Trim$(Profilename)
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

Public Sub RenameProfile(Profilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim resS As String, i As Long, NewPrinterProfiles As Collection, tStr As String, sa(2) As String
50020
50030  resS = Trim$(Profilename)
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

Private Function ProfileExists(Profilename) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  For i = 0 To cmbProfile.ListCount - 1
50030   If StrComp(cmbProfile.List(i), Profilename, vbTextCompare) = 0 Then
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

Private Function ProfileAssociatedPrinter(Profilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PrinterProfiles As Collection, p As Variant, i As Long, tStr As String
50020  Set PrinterProfiles = GetPrinterProfiles
50030
50040  For i = 1 To PrinterProfiles.Count
50050   If StrComp(PrinterProfiles(i)(1), Profilename, vbTextCompare) = 0 Then
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
50050   tempOptions = GetOptionsFromUserControls   ' Get the current settings settings
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
50041   Select Case ieb.GetSelectedGroup
          Case 1
50061      Select Case ieb.GetSelectedItem
            Case 1
50080        Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50090       Case 2
50100        Call HTMLHelp_ShowTopic("html\ghostscript.htm")
50110       Case 3
50120        Call HTMLHelp_ShowTopic("html\docproperties.htm")
50130       Case 4
50140        Call HTMLHelp_ShowTopic("html\savesettings.htm")
50150       Case 5
50160        Call HTMLHelp_ShowTopic("html\autosave.htm")
50170       Case Else
50180        Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50190      End Select
50200     Case 2
50211      Select Case ieb.GetSelectedItem
            Case 1
50231        Select Case optFormatPDF.PDFOptionsIndex
              Case 1
50250          Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50260         Case 2
50270          Call HTMLHelp_ShowTopic("html\pdfcompression.htm")
50280         Case 3
50290          Call HTMLHelp_ShowTopic("html\pdffonts.htm")
50300         Case 4
50310          Call HTMLHelp_ShowTopic("html\pdfcolors.htm")
50320         Case 5
50330          Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
50340         Case 6
50350          Call HTMLHelp_ShowTopic("html\pdfsigning.htm")
50360         Case Else
50370          Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50380        End Select
50390       Case 2
50400        Call HTMLHelp_ShowTopic("html\pngsettings.htm")
50410       Case 3
50420        Call HTMLHelp_ShowTopic("html\jpegsettings.htm")
50430       Case 4
50440        Call HTMLHelp_ShowTopic("html\bmpsettings.htm")
50450       Case 5
50460        Call HTMLHelp_ShowTopic("html\pcxsettings.htm")
50470       Case 6
50480        Call HTMLHelp_ShowTopic("html\tiffsettings.htm")
50490       Case 7
50500        Call HTMLHelp_ShowTopic("html\pssettings.htm")
50510       Case 8
50520        Call HTMLHelp_ShowTopic("html\epssettings.htm")
50530       Case Else
50540        Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50550      End Select
50560    End Select
50570  End If
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
50200   ieb.DisableUpdates True
50210   ieb.SetGroupCaption "Program", .OptionsTreeProgram
50220   ieb.SetItemText "Program", "General", .OptionsProgramGeneralSymbol
50230   ieb.SetItemText "Program", "Ghostscript", .OptionsProgramGhostscriptSymbol
50240   ieb.SetItemText "Program", "Document", .OptionsProgramDocumentSymbol
50250   ieb.SetItemText "Program", "Save", .OptionsProgramSaveSymbol
50260   ieb.SetItemText "Program", "AutoSave", .OptionsProgramAutosaveSymbol
50270 '  ieb.SetItemText "Program", "Directories", .OptionsProgramDirectoriesSymbol
50280
50290   ieb.SetItemText "Program", "Actions", .OptionsProgramActionsSymbol
50300   ieb.SetItemText "Program", "Print", .OptionsProgramPrintSymbol
50310  ' ieb.SetItemText "Program", "Fonts", .OptionsProgramFontSymbol
50320   ieb.SetItemText "Program", "Language", .OptionsProgramLanguagesSymbol
50330
50340   ieb.SetGroupCaption "Formats", .OptionsTreeFormats
50350   ieb.SetItemText "Formats", "PDF", .OptionsPDFSymbol
50360   ieb.SetItemText "Formats", "PNG", .OptionsPNGSymbol
50370   ieb.SetItemText "Formats", "JPEG", .OptionsJPEGSymbol
50380   ieb.SetItemText "Formats", "BMP", .OptionsBMPSymbol
50390   ieb.SetItemText "Formats", "PCX", .OptionsPCXSymbol
50400   ieb.SetItemText "Formats", "TIFF", .OptionsTIFFSymbol
50410   ieb.SetItemText "Formats", "PS", .OptionsPSSymbol
50420   ieb.SetItemText "Formats", "EPS", .OptionsEPSSymbol
50430   ieb.SetItemText "Formats", "TXT", .OptionsTXTSymbol
50440   ieb.SetItemText "Formats", "PSD", .OptionsPSDSymbol
50450   ieb.SetItemText "Formats", "PCL", .OptionsPCLSymbol
50460   ieb.SetItemText "Formats", "RAW", .OptionsRAWSymbol
50470   ieb.SetItemText "Formats", "SVG", .OptionsSVGSymbol
50480 '  ieb.SetItemText "Formats", "XCF", .OptionsXCFSymbol
50490   ieb.DisableUpdates False
50500
50510   lblOptions.Caption = .OptionsProgramLanguagesDescription
50520  End With
50530  optActions.SetLanguageStrings
50540  optAutosave.SetLanguageStrings
50550 ' optDirectories.SetLanguageStrings
50560  optDocument.SetLanguageStrings
50570 ' optFonts.SetLanguageStrings
50580  optFormatPNG.SetLanguageStrings
50590  optFormatJPEG.SetLanguageStrings
50600  optFormatBMP.SetLanguageStrings
50610  optFormatPCX.SetLanguageStrings
50620  optFormatTIFF.SetLanguageStrings
50630  optFormatPDF.SetLanguageStrings
50640  optFormatPS.SetLanguageStrings
50650  optFormatEPS.SetLanguageStrings
50660  optFormatTXT.SetLanguageStrings
50670  optFormatPSD.SetLanguageStrings
50680  optFormatPCL.SetLanguageStrings
50690  optFormatRAW.SetLanguageStrings
50700  optFormatSVG.SetLanguageStrings
50710 ' optFormatXCF.SetLanguageStrings
50720
50730  optGeneral.SetLanguageStrings
50740  optGhostscript.SetLanguageStrings
50750  optLanguages.SetLanguageStrings
50760  optPrint.SetLanguageStrings
50770  optSave.SetLanguageStrings
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
50350    ieb.Refresh
50360   End With
50370  End If
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

Private Function GetOptionsFromUserControls() As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Options1 = StandardOptions
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
50280  GetOptionsFromUserControls = Options1
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
  PrinterProfiles As Collection, sa(2) As String
50030
50040  tRestart = False
50050  If UCase$(Options.DirectoryGhostscriptBinaries) <> UCase$(ProfileOptions(0).DirectoryGhostscriptBinaries) Then
50060   tRestart = True
50070  End If
50080
50090  ' Save all Options/Profiles
50100  ProfileOptions(cmbProfile.ListIndex) = GetOptionsFromUserControls ' Get the current settings
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
50320   SaveOptions ProfileOptions(i), cmbProfile.List(i)
50330  Next i
50340  ' Ready profiles saving
50350
50360  Set PrinterProfiles = New Collection
50370  For i = 1 To TempPrinterProfiles.Count
50380   sa(0) = TempPrinterProfiles(i)(0)
50390   sa(1) = TempPrinterProfiles(i)(2)
50400   PrinterProfiles.Add sa
50410  Next i
50420
50430  SavePrinterProfiles PrinterProfiles
50440
50450  SetHelpfile
50460
50470  If IsWin9xMe = False Then
50481   Select Case Options.ProcessPriority
         Case 0: 'Idle
50500     SetProcessPriority Idle
50510    Case 1: 'Normal
50520     SetProcessPriority Normal
50530    Case 2: 'High
50540     SetProcessPriority High
50550    Case 3: 'Realtime
50560     SetProcessPriority RealTime
50570   End Select
50580  End If
50590  If tRestart = True Then
50600   Restart = True
50610  End If
50620  Unload Me
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
50070
50080  Set TempPrinterProfiles = New Collection
50090
50100  Options1 = Options
50110  CurrentLanguage = Options.Language
50120
50130  UnloadForm = False
50140  Me.Icon = LoadResPicture(2120, vbResIcon)
50150  KeyPreview = True
50160
50170  cmbProfile.Top = (cmdProfileAdd.Height - cmbProfile.Height) / 2 + cmdProfileAdd.Top
50180
50190  oldLanguage = Options.Language
50200
50210  With Screen
50220   .MousePointer = vbHourglass
50230   Move (.Width - Width) / 2, (.Height - Height) / 2
50240  End With
50250
50260  LastProgramGroup = ""
50270  LastFormatGroup = ""
50280
50290  ieb.DisableUpdates True
50300  ieb.ClearStructure
50310  ieb.SetImageList imlIeb
50320  With LanguageStrings
50330   ieb.AddGroup "Program", .OptionsTreeProgram, 0
50340   ieb.AddItem "Program", "General", .OptionsProgramGeneralSymbol, 1
50350   ieb.AddItem "Program", "Ghostscript", .OptionsProgramGhostscriptSymbol, 2
50360   ieb.AddItem "Program", "Document", .OptionsProgramDocumentSymbol, 3
50370   ieb.AddItem "Program", "Save", .OptionsProgramSaveSymbol, 4
50380   ieb.AddItem "Program", "AutoSave", .OptionsProgramAutosaveSymbol, 5
50390 '  ieb.AddItem "Program", "Directories", .OptionsProgramDirectoriesSymbol, 6
50400   ieb.AddItem "Program", "Actions", .OptionsProgramActionsSymbol, 7
50410   ieb.AddItem "Program", "Print", .OptionsProgramPrintSymbol, 8
50420 '  ieb.AddItem "Program", "Fonts", .OptionsProgramFontSymbol, 9
50430   ieb.AddItem "Program", "Language", .OptionsProgramLanguagesSymbol, 10
50440
50450   ieb.AddGroup "Formats", .OptionsTreeFormats, 0
50460   ieb.AddItem "Formats", "PDF", .OptionsPDFSymbol, 11
50470   ieb.AddItem "Formats", "PNG", .OptionsPNGSymbol, 12
50480   ieb.AddItem "Formats", "JPEG", .OptionsJPEGSymbol, 13
50490   ieb.AddItem "Formats", "BMP", .OptionsBMPSymbol, 14
50500   ieb.AddItem "Formats", "PCX", .OptionsPCXSymbol, 15
50510   ieb.AddItem "Formats", "TIFF", .OptionsTIFFSymbol, 16
50520   ieb.AddItem "Formats", "PS", .OptionsPSSymbol, 17
50530   ieb.AddItem "Formats", "EPS", .OptionsEPSSymbol, 18
50540   ieb.AddItem "Formats", "TXT", .OptionsTXTSymbol, 21
50550   ieb.AddItem "Formats", "PSD", .OptionsPSDSymbol, 22
50560   ieb.AddItem "Formats", "PCL", .OptionsPCLSymbol, 23
50570   ieb.AddItem "Formats", "RAW", .OptionsRAWSymbol, 24
50580 '  ieb.AddItem "Formats", "XCF", .OptionsXCFSymbol, 25
50590   ieb.AddItem "Formats", "SVG", .OptionsSVGSymbol, 26
50600   ieb.ExpandGroup "Formats", False
50610
50620   ieb.DisableUpdates False
50630
50640   Set picOptions = LoadResPicture(2101, vbResIcon)
50650
50660   Me.Caption = .DialogPrinterOptions
50670   cmdCancel.Caption = .OptionsCancel
50680   cmdReset.Caption = .OptionsReset
50690   cmdSave.Caption = .OptionsSave
50700  End With
50710
50720  SetFrame dmFraDescription
50730  SetFrame dmFraProfile
50740
50750  ' Add ActionsControl
50760  Set optActionsControl = Controls.Add("PDFCreator.ctlOptActions", "ctlOptActions")
50770  optActionsControl.Width = dmFraDescription.Width
50780  Set optActions = optActionsControl.object
50790  optActions.SetLanguageStrings
50800  optActions.SetOptions
50810  ' Add AutosaveControl
50820  Set optAutosaveControl = Controls.Add("PDFCreator.ctlOptAutosave", "ctlOptAutosave")
50830  optAutosaveControl.Width = dmFraDescription.Width
50840  Set optAutosave = optAutosaveControl.object
50850  optAutosave.SetLanguageStrings
50860  optAutosave.SetOptions
50870  ' Add DirectoriesControl
50880 ' Set optDirectoriesControl = Controls.Add("PDFCreator.ctlOptDirectories", "ctlOptDirectories")
50890 ' optDirectoriesControl.Width = dmFraDescription.Width
50900 ' Set optDirectories = optDirectoriesControl.object
50910 ' optDirectories.SetLanguageStrings
50920 ' optDirectories.SetOptions
50930  ' Add DocumentControl
50940  Set optDocumentControl = Controls.Add("PDFCreator.ctlOptDocument", "ctlOptDocument")
50950  optDocumentControl.Width = dmFraDescription.Width
50960  Set optDocument = optDocumentControl.object
50970  optDocument.SetLanguageStrings
50980  optDocument.SetOptions
50990  ' Add FontsControl
51000 ' Set optFontsControl = Controls.Add("PDFCreator.ctlOptFonts", "ctlOptFonts")
51010 ' optFontsControl.Width = dmFraDescription.Width
51020 ' Set optFonts = optFontsControl.object
51030 ' optFonts.SetLanguageStrings
51040 ' optFonts.SetOptions
51050  ' Add FormatPNGControl
51060  Set optFormatPNGControl = Controls.Add("PDFCreator.ctlOptFormatPNG", "ctlOptFormatPNG")
51070  optFormatPNGControl.Width = dmFraDescription.Width
51080  Set optFormatPNG = optFormatPNGControl.object
51090  optFormatPNG.SetLanguageStrings
51100  optFormatPNG.SetOptions
51110  ' Add FormatJPEQControl
51120  Set optFormatJPEGControl = Controls.Add("PDFCreator.ctlOptFormatJPEG", "ctlOptFormatJPEG")
51130  optFormatJPEGControl.Width = dmFraDescription.Width
51140  Set optFormatJPEG = optFormatJPEGControl.object
51150  optFormatJPEG.SetLanguageStrings
51160  optFormatJPEG.SetOptions
51170  ' Add FormatBMPControl
51180  Set optFormatBMPControl = Controls.Add("PDFCreator.ctlOptFormatBMP", "ctlOptFormatBMP")
51190  optFormatBMPControl.Width = dmFraDescription.Width
51200  Set optFormatBMP = optFormatBMPControl.object
51210  optFormatBMP.SetLanguageStrings
51220  optFormatBMP.SetOptions
51230  ' Add FormatPCXControl
51240  Set optFormatPCXControl = Controls.Add("PDFCreator.ctlOptFormatPCX", "ctlOptFormatPCX")
51250  optFormatPCXControl.Width = dmFraDescription.Width
51260  Set optFormatPCX = optFormatPCXControl.object
51270  optFormatPCX.SetLanguageStrings
51280  optFormatPCX.SetOptions
51290  ' Add FormatTIFFControl
51300  Set optFormatTIFFControl = Controls.Add("PDFCreator.ctlOptFormatTIFF", "ctlOptFormatTIFF")
51310  optFormatTIFFControl.Width = dmFraDescription.Width
51320  Set optFormatTIFF = optFormatTIFFControl.object
51330  optFormatTIFF.SetLanguageStrings
51340  optFormatTIFF.SetOptions
51350  ' Add FormatPDFControl
51360  Set optFormatPDFControl = Controls.Add("PDFCreator.ctlOptFormatPDF", "ctlOptFormatPDF")
51370  optFormatPDFControl.Width = dmFraDescription.Width
51380  Set optFormatPDF = optFormatPDFControl.object
51390  optFormatPDF.SetLanguageStrings
51400  optFormatPDF.SetOptions
51410  ' Add FormatPS
51420  Set optFormatPSControl = Controls.Add("PDFCreator.ctlOptFormatPS", "ctlOptFormatPS")
51430  optFormatPSControl.Width = dmFraDescription.Width
51440  Set optFormatPS = optFormatPSControl.object
51450  optFormatPS.SetLanguageStrings
51460  optFormatPS.SetOptions
51470  ' Add FormatEPSControl
51480  Set optFormatEPSControl = Controls.Add("PDFCreator.ctlOptFormatEPS", "ctlOptFormatEPS")
51490  optFormatEPSControl.Width = dmFraDescription.Width
51500  Set optFormatEPS = optFormatEPSControl.object
51510  optFormatEPS.SetLanguageStrings
51520  optFormatEPS.SetOptions
51530  ' Add FormatTXTControl
51540  Set optFormatTXTControl = Controls.Add("PDFCreator.ctlOptFormatTXT", "ctlOptFormatTXT")
51550  optFormatTXTControl.Width = dmFraDescription.Width
51560  Set optFormatTXT = optFormatTXTControl.object
51570  optFormatTXT.SetLanguageStrings
51580  optFormatTXT.SetOptions
51590  ' Add FormatPSDControl
51600  Set optFormatPSDControl = Controls.Add("PDFCreator.ctlOptFormatPSD", "ctlOptFormatPSD")
51610  optFormatPSDControl.Width = dmFraDescription.Width
51620  Set optFormatPSD = optFormatPSDControl.object
51630  optFormatPSD.SetLanguageStrings
51640  optFormatPSD.SetOptions
51650  ' Add FormatPCLControl
51660  Set optFormatPCLControl = Controls.Add("PDFCreator.ctlOptFormatPCL", "ctlOptFormatPCL")
51670  optFormatPCLControl.Width = dmFraDescription.Width
51680  Set optFormatPCL = optFormatPCLControl.object
51690  optFormatPCL.SetLanguageStrings
51700  optFormatPCL.SetOptions
51710  ' Add FormatRAWControl
51720  Set optFormatRAWControl = Controls.Add("PDFCreator.ctlOptFormatRAW", "ctlOptFormatRAW")
51730  optFormatRAWControl.Width = dmFraDescription.Width
51740  Set optFormatRAW = optFormatRAWControl.object
51750  optFormatRAW.SetLanguageStrings
51760  optFormatRAW.SetOptions
51770  ' Add FormatSVGControl
51780  Set optFormatSVGControl = Controls.Add("PDFCreator.ctlOptFormatSVG", "ctlOptFormatSVG")
51790  optFormatSVGControl.Width = dmFraDescription.Width
51800  Set optFormatSVG = optFormatSVGControl.object
51810  optFormatSVG.SetLanguageStrings
51820  optFormatSVG.SetOptions
51830 ' ' Add FormatXCFControl - Doesn't work
51840 ' Set optFormatXCFControl = Controls.Add("PDFCreator.ctlOptFormatXCF", "ctlOptFormatXCF")
51850 ' optFormatXCFControl.Width = dmFraDescription.Width
51860 ' Set optFormatXCF = optFormatXCFControl.object
51870 ' optFormatXCF.SetLanguageStrings
51880 ' optFormatXCF.SetOptions
51890  ' Add GhostscriptControl
51900  Set optGhostscriptControl = Controls.Add("PDFCreator.ctlOptGhostscript", "ctlOptGhostscript")
51910  optGhostscriptControl.Width = dmFraDescription.Width
51920  Set optGhostscript = optGhostscriptControl.object
51930  optGhostscript.SetLanguageStrings
51940  optGhostscript.SetOptions
51950  ' Add LanguagesControl
51960  Set optLanguagesControl = Controls.Add("PDFCreator.ctlOptLanguages", "ctlOptLanguages")
51970  optLanguagesControl.Width = dmFraDescription.Width
51980  Set optLanguages = optLanguagesControl.object
51990  optLanguages.SetLanguageStrings
52000  optLanguages.SetOptions
52010  ' Add PrintControl
52020  Set optPrintControl = Controls.Add("PDFCreator.ctlOptPrint", "ctlOptPrint")
52030  optPrintControl.Width = dmFraDescription.Width
52040  Set optPrint = optPrintControl.object
52050  optPrint.SetLanguageStrings
52060  optPrint.SetOptions
52070  ' Add SaveControl
52080  '
52090  Set optSaveControl = Controls.Add("PDFCreator.ctlOptSave", "ctlOptSave")
52100  optSaveControl.Width = dmFraDescription.Width
52110  Set optSave = optSaveControl.object
52120  optSave.SetLanguageStrings
52130  optSave.SetOptions
52140  ' Add GeneralControl
52150  '
52160  Set optGeneralControl = Controls.Add("PDFCreator.ctlOptGeneral", "ctlOptGeneral")
52170  optGeneralControl.Width = dmFraDescription.Width
52180  Set optGeneral = optGeneralControl.object
52190  optGeneral.SetLanguageStrings
52200 '
52210  optGeneral.SetOptions
52220
52230  dmFraProfile.Caption = LanguageStrings.OptionsProfile
52240  cmbProfile.Clear
52250  cmbProfile.AddItem LanguageStrings.OptionsProfileDefaultName
52260
52270  Set Profiles = GetProfiles
52280  ReDim ProfileNames(Profiles.Count)
52290  ReDim ProfileOptions(Profiles.Count)
52300  ProfileNames(0) = LanguageStrings.OptionsProfileDefaultName
52310  ProfileOptions(0) = Options
52320
52330  With dmFraDescription
52340   .Caption = LanguageStrings.OptionsTreeProgram
52350   .Visible = True
52360
52370   optActionsControl.Top = .Top + .Height + ControlTop
52380   optActionsControl.Left = .Left
52390   optActionsControl.Width = .Width
52400   optAutosaveControl.Top = .Top + .Height + ControlTop
52410   optAutosaveControl.Left = .Left
52420   optAutosaveControl.Width = .Width
52430 '  optDirectoriesControl.Top = .Top + .Height + ControlTop
52440 '  optDirectoriesControl.Left = .Left
52450 '  optDirectoriesControl.Width = .Width
52460   optDocumentControl.Top = .Top + .Height + ControlTop
52470   optDocumentControl.Left = .Left
52480   optDocumentControl.Width = .Width
52490 '  optFontsControl.Top = .Top + .Height + ControlTop
52500 '  optFontsControl.Left = .Left
52510 '  optFontsControl.Width = .Width
52520   optFormatPNGControl.Top = .Top + .Height + ControlTop
52530   optFormatPNGControl.Left = .Left
52540   optFormatPNGControl.Width = .Width
52550   optFormatJPEGControl.Top = .Top + .Height + ControlTop
52560   optFormatJPEGControl.Left = .Left
52570   optFormatJPEGControl.Width = .Width
52580   optFormatBMPControl.Top = .Top + .Height + ControlTop
52590   optFormatBMPControl.Left = .Left
52600   optFormatBMPControl.Width = .Width
52610   optFormatPCXControl.Top = .Top + .Height + ControlTop
52620   optFormatPCXControl.Left = .Left
52630   optFormatPCXControl.Width = .Width
52640   optFormatTIFFControl.Top = .Top + .Height + ControlTop
52650   optFormatTIFFControl.Left = .Left
52660   optFormatTIFFControl.Width = .Width
52670   optFormatPDFControl.Top = .Top + .Height + ControlTop
52680   optFormatPDFControl.Left = .Left
52690   optFormatPDFControl.Width = .Width
52700   optFormatPSControl.Top = .Top + .Height + ControlTop
52710   optFormatPSControl.Left = .Left
52720   optFormatPSControl.Width = .Width
52730   optFormatEPSControl.Top = .Top + .Height + ControlTop
52740   optFormatEPSControl.Left = .Left
52750   optFormatEPSControl.Width = .Width
52760   optFormatTXTControl.Top = .Top + .Height + ControlTop
52770   optFormatTXTControl.Left = .Left
52780   optFormatTXTControl.Width = .Width
52790   optFormatPSDControl.Top = .Top + .Height + ControlTop
52800   optFormatPSDControl.Left = .Left
52810   optFormatPSDControl.Width = .Width
52820   optFormatPCLControl.Top = .Top + .Height + ControlTop
52830   optFormatPCLControl.Left = .Left
52840   optFormatPCLControl.Width = .Width
52850   optFormatRAWControl.Top = .Top + .Height + ControlTop
52860   optFormatRAWControl.Left = .Left
52870   optFormatRAWControl.Width = .Width
52880   optFormatSVGControl.Top = .Top + .Height + ControlTop
52890   optFormatSVGControl.Left = .Left
52900   optFormatSVGControl.Width = .Width
52910 '  optFormatXCFControl.Top = .Top + .Height + ControlTop
52920 '  optFormatXCFControl.Left = .Left
52930 '  optFormatXCFControl.Width = .Width
52940   optGeneralControl.Top = .Top + .Height + ControlTop
52950   optGeneralControl.Left = .Left
52960   optGeneralControl.Width = .Width
52970   optGhostscriptControl.Top = .Top + .Height + ControlTop
52980   optGhostscriptControl.Left = .Left
52990   optGhostscriptControl.Width = .Width
53000   optLanguagesControl.Top = .Top + .Height + ControlTop
53010   optLanguagesControl.Left = .Left
53020   optLanguagesControl.Width = .Width
53030   optPrintControl.Top = .Top + .Height + ControlTop
53040   optPrintControl.Left = .Left
53050   optPrintControl.Width = .Width
53060   optSaveControl.Top = .Top + .Height + ControlTop
53070   optSaveControl.Left = .Left
53080   optSaveControl.Width = .Width
53090
53100   cmdCancel.Left = .Left
53110   cmdReset.Left = .Left + (.Width - cmdReset.Width) / 2
53120   cmdSave.Left = .Left + .Width - cmdSave.Width
53130  End With
53140
53150  For i = 1 To Profiles.Count
53160   cmbProfile.AddItem Profiles(i)
53170   ProfileNames(i) = Profiles(i)
53180   ProfileOptions(i) = ReadOptions(, , Profiles(i))
53190  Next i
53200  SetProfile CurrentPrinterProfile
53210
53220  If cmbProfile.ListIndex = 0 Then
53230    optGhostscript.ControlEnabled = True
53240    optLanguages.ControlEnabled = True
53250    cmdProfileRename.Enabled = False
53260    cmdProfileDelete.Enabled = False
53270   Else
53280    optGhostscript.ControlEnabled = False
53290    optLanguages.ControlEnabled = False
53300    cmdProfileRename.Enabled = True
53310    cmdProfileDelete.Enabled = True
53320  End If
53330
53340  Set PrinterProfiles = GetPrinterProfiles
53350  For i = 1 To PrinterProfiles.Count
53360   sa(0) = PrinterProfiles(i)(0)
53370   sa(1) = PrinterProfiles(i)(1)
53380   sa(2) = PrinterProfiles(i)(1)
53390   TempPrinterProfiles.Add sa
53400  Next i
53410
53420  With LanguageStrings
53430   cmdProfileAdd.ToolTipText = .OptionsProfileAdd
53440   cmdProfileDelete.ToolTipText = .OptionsProfileDel
53450   cmdProfileRename.ToolTipText = .OptionsProfileRenameProfile
53460   cmdProfileSave.ToolTipText = .OptionsProfileSaveToDisc
53470   cmdProfileLoad.ToolTipText = .OptionsProfileLoadFromDisc
53480   cmbProfile.List(0) = .OptionsProfileDefaultName
53490  End With
53500
53510  If ShowOnlyOptions = True Then
53520   FormInTaskbar Me, True, True
53530   Caption = "PDFCreator - " & Caption
53540  End If
53550
53560  ShowAcceleratorsInForm Me, True
53570
53580  Screen.MousePointer = vbNormal
53590
53600  With Options
53610   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
53620  End With
53630  ieb.Refresh
53640  ieb_ItemClick "Program", "General"
53650  LoadReady = True
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

Private Sub ieb_GroupClick(ByVal Group As Long, bExpanded As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ieb.DisableUpdates True
50021  Select Case Group
        Case 1:
50040    If bExpanded Then
50050      ieb.ExpandGroup 2, False
50060     Else
50070      ieb.ExpandGroup 2, True
50080    End If
50090    If LastFormatGroup = "" Then
50100      ieb_ItemClick "Program", "PDF"
50110     Else
50120      ieb_ItemClick "Program", LastProgramGroup
50130    End If
50140   Case 2:
50150    If bExpanded Then
50160      ieb.ExpandGroup 1, False
50170     Else
50180      ieb.ExpandGroup 1, True
50190    End If
50200    If LastFormatGroup = "" Then
50210      ieb_ItemClick "Formats", "PDF"
50220     Else
50230      ieb_ItemClick "Formats", LastFormatGroup
50240    End If
50250  End Select
50260  ieb.DisableUpdates False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ieb_GroupClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ieb_ItemClick(sGroup As String, sItemKey As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020
50030  optActionsControl.Visible = False
50040  optAutosaveControl.Visible = False
50050 ' optDirectoriesControl.Visible = False
50060  optDocumentControl.Visible = False
50070 ' optFontsControl.Visible = False
50080  optFormatPNGControl.Visible = False
50090  optFormatJPEGControl.Visible = False
50100  optFormatBMPControl.Visible = False
50110  optFormatPCXControl.Visible = False
50120  optFormatTIFFControl.Visible = False
50130  optFormatPDFControl.Visible = False
50140  optFormatPSControl.Visible = False
50150  optFormatEPSControl.Visible = False
50160  optFormatTXTControl.Visible = False
50170  optFormatPSDControl.Visible = False
50180  optFormatPCLControl.Visible = False
50190  optFormatRAWControl.Visible = False
50200  optFormatSVGControl.Visible = False
50210 ' optFormatXCFControl.Visible = False
50220  optGeneralControl.Visible = False
50230  optGhostscriptControl.Visible = False
50240  optLanguagesControl.Visible = False
50250  optPrintControl.Visible = False
50260  optSaveControl.Visible = False
50270
50281  Select Case UCase$(sGroup)
        Case "PROGRAM"
50300    LastProgramGroup = sItemKey
50311    Select Case UCase$(sItemKey)
          Case "GENERAL"
50330      Set picOptions = LoadResPicture(2101, vbResIcon)
50340      lblOptions.Caption = LanguageStrings.OptionsProgramGeneralDescription
50350      optGeneralControl.Visible = True
50360      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50370     Case "GHOSTSCRIPT"
50380      Set picOptions = LoadResPicture(2119, vbResIcon)
50390      lblOptions.Caption = LanguageStrings.OptionsProgramGhostscriptDescription
50400      optGhostscriptControl.Visible = True
50410      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50420     Case "DOCUMENT"
50430      Set picOptions = LoadResPicture(2105, vbResIcon)
50440      lblOptions.Caption = LanguageStrings.OptionsProgramDocumentDescription
50450      optDocumentControl.Visible = True
50460      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50470     Case "SAVE"
50480      Set picOptions = LoadResPicture(2106, vbResIcon)
50490      lblOptions.Caption = LanguageStrings.OptionsProgramSaveDescription
50500      optSaveControl.Visible = True
50510      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50520     Case "AUTOSAVE"
50530      Set picOptions = LoadResPicture(2103, vbResIcon)
50540      lblOptions.Caption = LanguageStrings.OptionsProgramAutosaveDescription
50550      optAutosaveControl.Visible = True
50560      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50570 '    Case "DIRECTORIES"
50580 '     Set picOptions = LoadResPicture(2104, vbResIcon)
50590 '     lblOptions.Caption = LanguageStrings.OptionsProgramDirectoriesDescription
50600 '     optDirectoriesControl.Visible = True
50610 '     dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50620     Case "ACTIONS"
50630      Set picOptions = LoadResPicture(2121, vbResIcon)
50640      lblOptions.Caption = LanguageStrings.OptionsProgramActionsDescription
50650      optActionsControl.Visible = True
50660      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50670     Case "PRINT"
50680      Set picOptions = LoadResPicture(2122, vbResIcon)
50690      lblOptions.Caption = LanguageStrings.OptionsProgramPrintDescription
50700      optPrintControl.Visible = True
50710      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50720 '    Case "FONTS"
50730 '     Set picOptions = LoadResPicture(2102, vbResIcon)
50740 '     lblOptions.Caption = LanguageStrings.OptionsProgramFontDescription
50750 '     optFontsControl.Visible = True
50760 '     dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50770     Case "LANGUAGE"
50780      Set picOptions = LoadResPicture(2123, vbResIcon)
50790      lblOptions.Caption = LanguageStrings.OptionsProgramLanguagesDescription
50800      optLanguagesControl.Visible = True
50810      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50820    End Select
50830   Case "FORMATS"
50840    LastFormatGroup = sItemKey
50851    Select Case UCase$(sItemKey)
          Case "PDF"
50870      Set picOptions = LoadResPicture(2111, vbResIcon)
50880      lblOptions.Caption = LanguageStrings.OptionsPDFDescription
50890      optFormatPDFControl.Visible = True
50900      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
50910      optLanguagesControl.Visible = True
50920     Case "PNG"
50930      Set picOptions = LoadResPicture(2112, vbResIcon)
50940      lblOptions.Caption = LanguageStrings.OptionsPNGDescription
50950      optFormatPNGControl.Visible = True
50960      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
50970     Case "JPEG"
50980      Set picOptions = LoadResPicture(2113, vbResIcon)
50990      lblOptions.Caption = LanguageStrings.OptionsJPEGDescription
51000      optFormatJPEGControl.Visible = True
51010      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51020     Case "BMP"
51030      Set picOptions = LoadResPicture(2114, vbResIcon)
51040      lblOptions.Caption = LanguageStrings.OptionsBMPDescription
51050      optFormatBMPControl.Visible = True
51060      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51070     Case "PCX"
51080      Set picOptions = LoadResPicture(2115, vbResIcon)
51090      lblOptions.Caption = LanguageStrings.OptionsPCXDescription
51100      optFormatPCXControl.Visible = True
51110      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51120     Case "TIFF"
51130      Set picOptions = LoadResPicture(2116, vbResIcon)
51140      lblOptions.Caption = LanguageStrings.OptionsTIFFDescription
51150      optFormatTIFFControl.Visible = True
51160      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51170     Case "PS"
51180      Set picOptions = LoadResPicture(2117, vbResIcon)
51190      lblOptions.Caption = LanguageStrings.OptionsPSDescription
51200      optFormatPSControl.Visible = True
51210      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51220     Case "EPS"
51230      Set picOptions = LoadResPicture(2118, vbResIcon)
51240      lblOptions.Caption = LanguageStrings.OptionsEPSDescription
51250      optFormatEPSControl.Visible = True
51260      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51270     Case "TXT"
51280      Set picOptions = LoadResPicture(2124, vbResIcon)
51290      lblOptions.Caption = LanguageStrings.OptionsTXTDescription
51300      optFormatTXTControl.Visible = True
51310      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51320     Case "PSD"
51330      Set picOptions = LoadResPicture(2125, vbResIcon)
51340      lblOptions.Caption = LanguageStrings.OptionsPSDDescription
51350      optFormatPSDControl.Visible = True
51360      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51370     Case "PCL"
51380      Set picOptions = LoadResPicture(2126, vbResIcon)
51390      lblOptions.Caption = LanguageStrings.OptionsPCLDescription
51400      optFormatPCLControl.Visible = True
51410      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51420     Case "RAW"
51430      Set picOptions = LoadResPicture(2127, vbResIcon)
51440      lblOptions.Caption = LanguageStrings.OptionsRAWDescription
51450      optFormatRAWControl.Visible = True
51460      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51470     Case "SVG"
51480      Set picOptions = LoadResPicture(2129, vbResIcon)
51490      lblOptions.Caption = LanguageStrings.OptionsSVGDescription
51500      optFormatSVGControl.Visible = True
51510      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51520 '    Case "XCF"
51530 '     Set picOptions = LoadResPicture(2128, vbResIcon)
51540 '     lblOptions.Caption = LanguageStrings.OptionsXCFDescription
51550 '     optFormatXCFControl.Visible = True
51560 '     dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51570    End Select
51580  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ieb_ItemClick")
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

Private Sub SetProfile(Optional ByVal Profilename As String = "") ' Empty profilename for default profile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  If cmbProfile.ListIndex < 0 Then
50030    OldProfile = 0
50040   Else
50050    OldProfile = cmbProfile.ListIndex
50060  End If
50070  Profilename = Trim$(Profilename)
50080  If LenB(Profilename) = 0 Then
50090   cmbProfile.ListIndex = 0
50100   Exit Sub
50110  End If
50120  For i = 1 To cmbProfile.ListCount - 1
50130   If LCase$(Profilename) = LCase$(cmbProfile.List(i)) Then
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
