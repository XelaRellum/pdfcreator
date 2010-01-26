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
 If cmbProfile.ListIndex <> OldProfile Then
  If OldProfile <= UBound(ProfileOptions) Then
   ProfileOptions(OldProfile) = GetOptionsFromUserControls(ProfileOptions(OldProfile))
  End If
  OldProfile = cmbProfile.ListIndex
  If cmbProfile.ListIndex = 0 Then
    optGhostscript.ControlEnabled = True
    optLanguages.ControlEnabled = True
    cmdProfileRename.Enabled = False
    cmdProfileDelete.Enabled = False
   Else
    optGhostscript.ControlEnabled = False
    optLanguages.ControlEnabled = False
    cmdProfileRename.Enabled = True
    cmdProfileDelete.Enabled = True
  End If
  Options1 = ProfileOptions(cmbProfile.ListIndex)
  Options1.Language = CurrentLanguage
  SetOptions
 End If
End Sub

Public Sub AddProfile(ProfileName As String)
 Dim resS As String, i As Long
 resS = Trim$(ProfileName)
 If LenB(resS) = 0 Then
  Exit Sub
 End If
 With cmbProfile
  ReDim Preserve ProfileOptions(.ListCount)
  ProfileOptions(.ListCount) = StandardOptions
  .AddItem resS
  .ListIndex = cmbProfile.ListCount - 1
  .Enabled = True
  .Visible = True
 End With
End Sub

Public Sub RenameProfile(ProfileName As String)
 Dim resS As String, i As Long, NewPrinterProfiles As Collection, tStr As String, sa(2) As String

 resS = Trim$(ProfileName)
 If LenB(resS) = 0 Then
  Exit Sub
 End If

 Set NewPrinterProfiles = New Collection
 For i = 1 To TempPrinterProfiles.Count
  sa(0) = TempPrinterProfiles(i)(0)
  sa(1) = TempPrinterProfiles(i)(1)
  If i - 1 = cmbProfile.ListIndex Then
    sa(2) = resS
   Else
    sa(2) = TempPrinterProfiles(i)(1)
  End If
  NewPrinterProfiles.Add sa
 Next i
 Set TempPrinterProfiles = NewPrinterProfiles

 With cmbProfile
  .List(.ListIndex) = resS
  .Enabled = True
  .Visible = True
 End With
End Sub

Private Function ProfileExists(ProfileName) As Boolean
 Dim i As Long
 For i = 0 To cmbProfile.ListCount - 1
  If StrComp(cmbProfile.List(i), ProfileName, vbTextCompare) = 0 Then
   ProfileExists = True
   Exit Function
  End If
 Next i
 ProfileExists = False
End Function

Private Function GetNextNewProfile(PreFix As String) As String
 Const MaxCount = 10000
 Dim i As Long, j As Long, NewProfile As String
 For i = 1 To MaxCount
  NewProfile = PreFix & " " & CStr(i)
  If Not ProfileExists(NewProfile) Then
   GetNextNewProfile = NewProfile
   Exit Function
  End If
 Next i
 GetNextNewProfile = PreFix
End Function

Private Sub cmdProfileAdd_Click()
 Dim resS As String, i As Long, res As Long

 Dim Profiles As New Collection
 For i = 0 To cmbProfile.ListCount - 1
  Profiles.Add cmbProfile.List(i)
 Next i

 Set frmProfile.Profiles = Profiles
 frmProfile.dmFrProfile.Caption = LanguageStrings.OptionsProfileAdd
 frmProfile.ProfileAction = eProfileAction.AddProfileAction
 frmProfile.txtProfile.Text = GetNextNewProfile(Trim$(LanguageStrings.OptionsProfileNewProfile))
 frmProfile.txtProfile.SelStart = Len(frmProfile.txtProfile.Text)
 frmProfile.Show vbModal, Me
End Sub

Private Function ProfileAssociatedPrinter(ProfileName As String)
 Dim PrinterProfiles As Collection, p As Variant, i As Long, tStr As String
 Set PrinterProfiles = GetPrinterProfiles

 For i = 1 To PrinterProfiles.Count
  If StrComp(PrinterProfiles(i)(1), ProfileName, vbTextCompare) = 0 Then
   If LenB(tStr) = 0 Then
     tStr = PrinterProfiles(i)(0)
    Else
     tStr = tStr & ", " & PrinterProfiles(i)(0)
   End If
  End If
 Next i
 ProfileAssociatedPrinter = tStr
End Function

Private Sub cmdProfileDelete_Click()
 Dim aw As Long, tStr As String, CurrentProfile As String, i As Long
 If cmbProfile.ListCount <= 1 Then
  Exit Sub
 End If
 CurrentProfile = cmbProfile.List(cmbProfile.ListIndex)
 tStr = ProfileAssociatedPrinter(CurrentProfile)
 If LenB(tStr) > 0 Then
  MsgBox LanguageStrings.MessagesMsg43 & " (" & tStr & ")"
  Exit Sub
 End If

 aw = MsgBox(Replace(LanguageStrings.MessagesMsg42, "%1", cmbProfile.List(cmbProfile.ListIndex)), vbQuestion Or vbYesNo)
 If aw = vbYes Then
  For i = cmbProfile.ListIndex + 1 To cmbProfile.ListCount - 1
   ProfileOptions(i - 1) = ProfileOptions(i)
  Next i
  ReDim Preserve ProfileOptions(cmbProfile.ListCount - 2)
  cmbProfile.RemoveItem cmbProfile.ListIndex
  cmbProfile.ListIndex = 0
 End If
End Sub

Private Sub cmdProfileLoad_Click()
 Dim FilterIndex As Long, files As Collection, dummyOptions As tOptions, tempOptions As tOptions
 FilterIndex = OpenFileDialog(files, cmbProfile.List(cmbProfile.ListIndex), _
   "PDFCreator options files (*.ini)|*.ini|All files (*.*)|*.*", "*.ini", GetMyFiles, _
   App.ProductName, OFN_ALLOWMULTISELECT + OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_LONGNAMES + OFN_NODEREFERENCELINKS, Me.hwnd)
 If FilterIndex > 0 Then
  tempOptions = ReadOptionsINI(dummyOptions, files(1), 0, True, True)
  ProfileOptions(cmbProfile.ListIndex) = tempOptions
  Options1 = ProfileOptions(cmbProfile.ListIndex)
  Options1.Language = CurrentLanguage
  SetOptions
 End If
End Sub

Private Sub cmdProfileRename_Click()
 Dim resS As String, i As Long, res As Long

 Dim Profiles As New Collection
 For i = 0 To cmbProfile.ListCount - 1
  Profiles.Add cmbProfile.List(i)
 Next i

 Set frmProfile.Profiles = Profiles
 frmProfile.dmFrProfile.Caption = LanguageStrings.OptionsProfileRenameProfile
 frmProfile.ProfileAction = eProfileAction.RenameProfileAction
 frmProfile.txtProfile.Text = cmbProfile.List(cmbProfile.ListIndex)
 frmProfile.txtProfile.SelStart = Len(frmProfile.txtProfile.Text)
 frmProfile.CurrentProfile = cmbProfile.List(cmbProfile.ListIndex)
 frmProfile.Show vbModal, Me
End Sub

Private Sub cmdProfileSave_Click()
 Dim FName As String, res As Long, tempOptions As tOptions
 res = SaveFileDialog(FName, cmbProfile.List(cmbProfile.ListIndex), "PDFCreator options files (*.ini)|*.ini|All files (*.*)|*.*", "*.ini", _
  GetMyFiles, App.ProductName, OFN_EXPLORER + OFN_PATHMUSTEXIST + OFN_LONGNAMES + OFN_HIDEREADONLY + OFN_OVERWRITEPROMPT, Me.hwnd)
 If res > 0 Then
  tempOptions = GetOptionsFromUserControls(ProfileOptions(cmbProfile.ListIndex))   ' Get the current settings settings
  SaveOptionsINI tempOptions, FName
 End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF1 Then
  KeyCode = 0
  Select Case trvOptions.SelectedItem.key
   Case "Program", "ProgramGeneral"
    Call HTMLHelp_ShowTopic("html\generalsettings.htm")
   Case "ProgramGhostscript"
    Call HTMLHelp_ShowTopic("html\ghostscript.htm")
   Case "ProgramDocument"
    Call HTMLHelp_ShowTopic("html\docproperties.htm")
   Case "ProgramSave"
    Call HTMLHelp_ShowTopic("html\savesettings.htm")
   Case "ProgramAutosave"
    Call HTMLHelp_ShowTopic("html\autosave.htm")
   Case "ProgramActions"
    Call HTMLHelp_ShowTopic("html\actions.htm")
   Case "ProgramPrint"
    Call HTMLHelp_ShowTopic("html\print.htm")
   Case "ProgramFonts"
    Call HTMLHelp_ShowTopic("html\fontsettings.htm")
   Case "ProgramLanguages"
    Call HTMLHelp_ShowTopic("html\changelang.htm")
   Case "Formats", "FormatsPDF"
    If trvOptions.SelectedItem.key = "FormatsPDF" Then
      Select Case optFormatPDF.PDFOptionsIndex
       Case 1
        Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
       Case 2
        Call HTMLHelp_ShowTopic("html\pdfcompression.htm")
       Case 3
        Call HTMLHelp_ShowTopic("html\pdffonts.htm")
       Case 4
        Call HTMLHelp_ShowTopic("html\pdfcolors.htm")
       Case 5
        Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
       Case 6
        Call HTMLHelp_ShowTopic("html\pdfsigning.htm")
      Case Else
       Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
      End Select
     Else
      Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
    End If
   Case "FormatsPNG"
    Call HTMLHelp_ShowTopic("html\pngsettings.htm")
   Case "FormatsJPEG"
    Call HTMLHelp_ShowTopic("html\jpegsettings.htm")
   Case "FormatsBMP"
    Call HTMLHelp_ShowTopic("html\bmpsettings.htm")
   Case "FormatsPCX"
    Call HTMLHelp_ShowTopic("html\pcxsettings.htm")
   Case "FormatsTIFF"
    Call HTMLHelp_ShowTopic("html\tiffsettings.htm")
   Case "FormatsPS"
    Call HTMLHelp_ShowTopic("html\pssettings.htm")
   Case "FormatsEPS"
    Call HTMLHelp_ShowTopic("html\epssettings.htm")
   Case "FormatsTXT"
    Call HTMLHelp_ShowTopic("html\txtsettings.htm")
   Case "FormatsPSD"
    Call HTMLHelp_ShowTopic("html\psdsettings.htm")
   Case "FormatsPCL"
    Call HTMLHelp_ShowTopic("html\pclsettings.htm")
   Case "FormatsRAW"
    Call HTMLHelp_ShowTopic("html\rawsettings.htm")
   Case "FormatsSVG"
    Call HTMLHelp_ShowTopic("html\svgsettings.htm")
   Case Else
    Call HTMLHelp_ShowTopic("html\generalsettings.htm")
   End Select
 End If
End Sub

Public Sub ChangeLanguage()
 If Not LoadReady Then
  Exit Sub
 End If

 With LanguageStrings
  dmFraProfile.Caption = .OptionsProfile
  cmdProfileAdd.ToolTipText = .OptionsProfileAdd
  cmdProfileDelete.ToolTipText = .OptionsProfileDel
  cmdProfileRename.ToolTipText = .OptionsProfileRenameProfile
  cmdProfileSave.ToolTipText = .OptionsProfileSaveToDisc
  cmdProfileLoad.ToolTipText = .OptionsProfileLoadFromDisc
  cmbProfile.List(0) = .OptionsProfileDefaultName

  dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
  Me.Caption = .DialogPrinterOptions
  cmdCancel.Caption = .OptionsCancel
  cmdReset.Caption = .OptionsReset
  cmdSave.Caption = .OptionsSave

  trvOptions.Nodes("Program").Text = .OptionsTreeProgram
  trvOptions.Nodes("ProgramGeneral").Text = .OptionsProgramGeneralSymbol
  trvOptions.Nodes("ProgramGhostscript").Text = .OptionsProgramGhostscriptSymbol
  trvOptions.Nodes("ProgramDocument").Text = .OptionsProgramDocumentSymbol
  trvOptions.Nodes("ProgramSave").Text = .OptionsProgramSaveSymbol
  trvOptions.Nodes("ProgramAutoSave").Text = .OptionsProgramAutosaveSymbol
  trvOptions.Nodes("ProgramActions").Text = .OptionsProgramActionsSymbol
  trvOptions.Nodes("ProgramPrint").Text = .OptionsProgramPrintSymbol
  trvOptions.Nodes("ProgramFonts").Text = .OptionsProgramFontSymbol
  trvOptions.Nodes("ProgramLanguages").Text = .OptionsProgramLanguagesSymbol

  trvOptions.Nodes("Formats").Text = .OptionsTreeFormats
  trvOptions.Nodes("FormatsPDF").Text = .OptionsPDFSymbol
  trvOptions.Nodes("FormatsPNG").Text = .OptionsPNGSymbol
  trvOptions.Nodes("FormatsJPEG").Text = .OptionsJPEGSymbol
  trvOptions.Nodes("FormatsBMP").Text = .OptionsBMPSymbol
  trvOptions.Nodes("FormatsPCX").Text = .OptionsPCXSymbol
  trvOptions.Nodes("FormatsTIFF").Text = .OptionsTIFFSymbol
  trvOptions.Nodes("FormatsPS").Text = .OptionsPSSymbol
  trvOptions.Nodes("FormatsEPS").Text = .OptionsEPSSymbol
  trvOptions.Nodes("FormatsTXT").Text = .OptionsTXTSymbol
  trvOptions.Nodes("FormatsPSD").Text = .OptionsPSDSymbol
  trvOptions.Nodes("FormatsPCL").Text = .OptionsPCLSymbol
  trvOptions.Nodes("FormatsRAW").Text = .OptionsRAWSymbol
  trvOptions.Nodes("FormatsSVG").Text = .OptionsSVGSymbol

  lblOptions.Caption = .OptionsProgramLanguagesDescription
 End With
 optActions.SetLanguageStrings
 optAutosave.SetLanguageStrings
' optDirectories.SetLanguageStrings
 optDocument.SetLanguageStrings
' optFonts.SetLanguageStrings
 optFormatPNG.SetLanguageStrings
 optFormatJPEG.SetLanguageStrings
 optFormatBMP.SetLanguageStrings
 optFormatPCX.SetLanguageStrings
 optFormatTIFF.SetLanguageStrings
 optFormatPDF.SetLanguageStrings
 optFormatPS.SetLanguageStrings
 optFormatEPS.SetLanguageStrings
 optFormatTXT.SetLanguageStrings
 optFormatPSD.SetLanguageStrings
 optFormatPCL.SetLanguageStrings
 optFormatRAW.SetLanguageStrings
 optFormatSVG.SetLanguageStrings
' optFormatXCF.SetLanguageStrings

 optGeneral.SetLanguageStrings
 optGhostscript.SetLanguageStrings
 optLanguages.SetLanguageStrings
 optPrint.SetLanguageStrings
 optSave.SetLanguageStrings
End Sub

Private Sub cmdCancel_Click()
 Dim f As Form, LanguagePath As String

 If CurrentLanguage <> Options.Language Then
  SetLanguage oldLanguage
  LanguagePath = CompletePath(GetPDFCreatorApplicationPath) & "Languages\"
  LoadLanguage LanguagePath & oldLanguage & ".ini"
  For Each f In Forms
   f.ChangeLanguage
  Next
 End If

 Unload Me
End Sub

Private Sub cmdReset_Click()
 Dim res As Long
 res = MsgBox(LanguageStrings.MessagesMsg03, vbYesNo)
 If res = vbYes Then
  Options = StandardOptions

  optActions.SetOptions
  optAutosave.SetOptions
'  optDirectories.SetOptions
  optDocument.SetOptions
'  optFonts.SetOptions
  optFormatPNG.SetOptions
  optFormatJPEG.SetOptions
  optFormatBMP.SetOptions
  optFormatPCX.SetOptions
  optFormatTIFF.SetOptions
  optFormatPDF.SetOptions
  optFormatPS.SetOptions
  optFormatEPS.SetOptions
  optFormatTXT.SetOptions
  optFormatPSD.SetOptions
  optFormatPCL.SetOptions
  optFormatRAW.SetOptions
  optFormatSVG.SetOptions
'  optFormatXCF.SetOptions

  optGeneral.SetOptions
  optGhostscript.SetOptions
  optLanguages.SetOptions
  optPrint.SetOptions
  optSave.SetOptions

  With Options
   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
   SetFont frmMain, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
  End With
 End If
End Sub

Private Function GetOptionsFromUserControls(DefaultOptions As tOptions) As tOptions
 Options1 = DefaultOptions
 optActions.GetOptions
 optAutosave.GetOptions
' optDirectories.GetOptions
 optDocument.GetOptions
 optFonts.GetOptions
 optFormatPNG.GetOptions
 optFormatJPEG.GetOptions
 optFormatBMP.GetOptions
 optFormatPCX.GetOptions
 optFormatTIFF.GetOptions
 optFormatPDF.GetOptions
 optFormatPS.GetOptions
 optFormatEPS.GetOptions
 optFormatTXT.GetOptions
 optFormatPSD.GetOptions
 optFormatPCL.GetOptions
 optFormatRAW.GetOptions
 optFormatSVG.GetOptions
' optFormatXCF.GetOptions

 optGeneral.GetOptions
 optGhostscript.GetOptions
 optLanguages.GetOptions
 optPrint.GetOptions
 optSave.GetOptions

 Options1.Counter = Options.Counter
 GetOptionsFromUserControls = Options1
End Function

Private Sub cmdSave_Click()
 Dim tRestart As Boolean, newLanguage As String, Profiles As Collection, i As Long, j As Long, _
  PrinterProfiles As Collection, sa(2) As String, tStr As String

 tRestart = False
 If UCase$(Options.DirectoryGhostscriptBinaries) <> UCase$(ProfileOptions(0).DirectoryGhostscriptBinaries) Then
  tRestart = True
 End If

 ' Save all Options/Profiles
 ProfileOptions(cmbProfile.ListIndex) = GetOptionsFromUserControls(ProfileOptions(cmbProfile.ListIndex)) ' Get the current settings

 ' Save default options
 Options = ProfileOptions(0)
 Options.Language = CurrentLanguage
 SaveOptions Options

 ' Delete all unnecessary/renamed profiles
 Set Profiles = GetProfiles
 For i = 1 To Profiles.Count
  For j = 1 To cmbProfile.ListCount
   If Profiles(i) = cmbProfile.List(j) Then
    Exit For
   End If
  Next j
  If j > cmbProfile.ListCount Then
   DeleteProfile Profiles(i)
  End If
 Next i

 ' Add all new/renamed profiles
 For i = 1 To cmbProfile.ListCount - 1
  ProfileOptions(i).LastUpdateCheck = Options.LastUpdateCheck
  SaveOptions ProfileOptions(i), cmbProfile.List(i)
 Next i
 ' Ready profiles saving

 Set PrinterProfiles = New Collection
 For i = 1 To TempPrinterProfiles.Count
  sa(0) = TempPrinterProfiles(i)(0)
  sa(1) = TempPrinterProfiles(i)(2)
  PrinterProfiles.Add sa
 Next i

 SavePrinterProfiles PrinterProfiles

 SetHelpfile

 If IsWin9xMe = False Then
  Select Case Options.ProcessPriority
   Case 0: 'Idle
    SetProcessPriority Idle
   Case 1: 'Normal
    SetProcessPriority Normal
   Case 2: 'High
    SetProcessPriority High
   Case 3: 'Realtime
    SetProcessPriority RealTime
  End Select
 End If
 If tRestart = True Then
  Restart = True
 End If
 Unload Me
End Sub

Private Sub Form_Load()
 Const ControlTop = 100, fraPDFTop = 1360, fraPDFLeft = 2960
 Dim pic As New StdPicture, i As Long, tStr As String, gsvers As Collection, _
  fc As Long, reg As clsRegistry, tsf() As String, tStr2 As String, files As Collection, _
  Path As String, filename As String, Ext As String, p As Printer
 Dim ctl1 As Control, ctl2 As Control, sa(2) As String
 Dim Profiles As Collection, PrinterProfiles As Collection
 Dim nodeProgram As Node, nodeFormats As Node

 Set TempPrinterProfiles = New Collection

 Options1 = Options
 CurrentLanguage = Options.Language

 UnloadForm = False
 Me.Icon = LoadResPicture(2120, vbResIcon)
 KeyPreview = True

 cmbProfile.Top = (cmdProfileAdd.Height - cmbProfile.Height) / 2 + cmdProfileAdd.Top

 oldLanguage = Options.Language

 With Screen
  .MousePointer = vbHourglass
  Move (.Width - Width) / 2, (.Height - Height) / 2
 End With

 trvOptions.Indentation = 0
 trvOptions.LineStyle = tvwTreeLines
 Set trvOptions.ImageList = imlIeb
 trvOptions.Nodes.Clear
 With LanguageStrings
  Set nodeProgram = trvOptions.Nodes.Add(, , "Program", .OptionsTreeProgram, 1)
  trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramGeneral", .OptionsProgramGeneralSymbol, 1
  trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramGhostscript", .OptionsProgramGhostscriptSymbol, 2
  trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramDocument", .OptionsProgramDocumentSymbol, 3
  trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramSave", .OptionsProgramSaveSymbol, 4
  trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramAutosave", .OptionsProgramAutosaveSymbol, 5
  trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramActions", .OptionsProgramActionsSymbol, 7
  trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramPrint", .OptionsProgramPrintSymbol, 8
  trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramFonts", .OptionsProgramFontSymbol, 9
  trvOptions.Nodes.Add nodeProgram, tvwChild, "ProgramLanguages", .OptionsProgramLanguagesSymbol, 10
  Set nodeFormats = trvOptions.Nodes.Add(, , "Formats", .OptionsTreeFormats, 11)
  trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPDF", .OptionsPDFSymbol, 11
  trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPNG", .OptionsPNGSymbol, 12
  trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsJPEG", .OptionsJPEGSymbol, 13
  trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsBMP", .OptionsBMPSymbol, 14
  trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPCX", .OptionsPCXSymbol, 15
  trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsTIFF", .OptionsTIFFSymbol, 16
  trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPS", .OptionsPSSymbol, 17
  trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsEPS", .OptionsEPSSymbol, 18
  trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsTXT", .OptionsTXTSymbol, 21
  trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPSD", .OptionsPSDSymbol, 22
  trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsPCL", .OptionsPCLSymbol, 23
  trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsRAW", .OptionsRAWSymbol, 24
  trvOptions.Nodes.Add nodeFormats, tvwChild, "FormatsSVG", .OptionsSVGSymbol, 26
 End With
 nodeProgram.Expanded = True
 nodeFormats.Expanded = True

 With LanguageStrings
  Set picOptions = LoadResPicture(2101, vbResIcon)
  Me.Caption = .DialogPrinterOptions
  cmdCancel.Caption = .OptionsCancel
  cmdReset.Caption = .OptionsReset
  cmdSave.Caption = .OptionsSave
 End With

 SetFrame dmFraDescription
 SetFrame dmFraProfile

 ' Add ActionsControl
 Set optActionsControl = Controls.Add("PDFCreator.ctlOptActions", "ctlOptActions")
 optActionsControl.Width = dmFraDescription.Width
 Set optActions = optActionsControl.object
 optActions.SetLanguageStrings
 optActions.SetOptions
 ' Add AutosaveControl
 Set optAutosaveControl = Controls.Add("PDFCreator.ctlOptAutosave", "ctlOptAutosave")
 optAutosaveControl.Width = dmFraDescription.Width
 Set optAutosave = optAutosaveControl.object
 optAutosave.SetLanguageStrings
 optAutosave.SetOptions
 ' Add DirectoriesControl
' Set optDirectoriesControl = Controls.Add("PDFCreator.ctlOptDirectories", "ctlOptDirectories")
' optDirectoriesControl.Width = dmFraDescription.Width
' Set optDirectories = optDirectoriesControl.object
' optDirectories.SetLanguageStrings
' optDirectories.SetOptions
 ' Add DocumentControl
 Set optDocumentControl = Controls.Add("PDFCreator.ctlOptDocument", "ctlOptDocument")
 optDocumentControl.Width = dmFraDescription.Width
 Set optDocument = optDocumentControl.object
 optDocument.SetLanguageStrings
 optDocument.SetOptions
 ' Add FontsControl
 Set optFontsControl = Controls.Add("PDFCreator.ctlOptFonts", "ctlOptFonts")
 optFontsControl.Width = dmFraDescription.Width
 Set optFonts = optFontsControl.object
 optFonts.SetLanguageStrings
 optFonts.SetOptions
 ' Add FormatPNGControl
 Set optFormatPNGControl = Controls.Add("PDFCreator.ctlOptFormatPNG", "ctlOptFormatPNG")
 optFormatPNGControl.Width = dmFraDescription.Width
 Set optFormatPNG = optFormatPNGControl.object
 optFormatPNG.SetLanguageStrings
 optFormatPNG.SetOptions
 ' Add FormatJPEQControl
 Set optFormatJPEGControl = Controls.Add("PDFCreator.ctlOptFormatJPEG", "ctlOptFormatJPEG")
 optFormatJPEGControl.Width = dmFraDescription.Width
 Set optFormatJPEG = optFormatJPEGControl.object
 optFormatJPEG.SetLanguageStrings
 optFormatJPEG.SetOptions
 ' Add FormatBMPControl
 Set optFormatBMPControl = Controls.Add("PDFCreator.ctlOptFormatBMP", "ctlOptFormatBMP")
 optFormatBMPControl.Width = dmFraDescription.Width
 Set optFormatBMP = optFormatBMPControl.object
 optFormatBMP.SetLanguageStrings
 optFormatBMP.SetOptions
 ' Add FormatPCXControl
 Set optFormatPCXControl = Controls.Add("PDFCreator.ctlOptFormatPCX", "ctlOptFormatPCX")
 optFormatPCXControl.Width = dmFraDescription.Width
 Set optFormatPCX = optFormatPCXControl.object
 optFormatPCX.SetLanguageStrings
 optFormatPCX.SetOptions
 ' Add FormatTIFFControl
 Set optFormatTIFFControl = Controls.Add("PDFCreator.ctlOptFormatTIFF", "ctlOptFormatTIFF")
 optFormatTIFFControl.Width = dmFraDescription.Width
 Set optFormatTIFF = optFormatTIFFControl.object
 optFormatTIFF.SetLanguageStrings
 optFormatTIFF.SetOptions
 ' Add FormatPDFControl
 Set optFormatPDFControl = Controls.Add("PDFCreator.ctlOptFormatPDF", "ctlOptFormatPDF")
 optFormatPDFControl.Width = dmFraDescription.Width
 Set optFormatPDF = optFormatPDFControl.object
 optFormatPDF.SetLanguageStrings
 optFormatPDF.SetOptions
 ' Add FormatPS
 Set optFormatPSControl = Controls.Add("PDFCreator.ctlOptFormatPS", "ctlOptFormatPS")
 optFormatPSControl.Width = dmFraDescription.Width
 Set optFormatPS = optFormatPSControl.object
 optFormatPS.SetLanguageStrings
 optFormatPS.SetOptions
 ' Add FormatEPSControl
 Set optFormatEPSControl = Controls.Add("PDFCreator.ctlOptFormatEPS", "ctlOptFormatEPS")
 optFormatEPSControl.Width = dmFraDescription.Width
 Set optFormatEPS = optFormatEPSControl.object
 optFormatEPS.SetLanguageStrings
 optFormatEPS.SetOptions
 ' Add FormatTXTControl
 Set optFormatTXTControl = Controls.Add("PDFCreator.ctlOptFormatTXT", "ctlOptFormatTXT")
 optFormatTXTControl.Width = dmFraDescription.Width
 Set optFormatTXT = optFormatTXTControl.object
 optFormatTXT.SetLanguageStrings
 optFormatTXT.SetOptions
 ' Add FormatPSDControl
 Set optFormatPSDControl = Controls.Add("PDFCreator.ctlOptFormatPSD", "ctlOptFormatPSD")
 optFormatPSDControl.Width = dmFraDescription.Width
 Set optFormatPSD = optFormatPSDControl.object
 optFormatPSD.SetLanguageStrings
 optFormatPSD.SetOptions
 ' Add FormatPCLControl
 Set optFormatPCLControl = Controls.Add("PDFCreator.ctlOptFormatPCL", "ctlOptFormatPCL")
 optFormatPCLControl.Width = dmFraDescription.Width
 Set optFormatPCL = optFormatPCLControl.object
 optFormatPCL.SetLanguageStrings
 optFormatPCL.SetOptions
 ' Add FormatRAWControl
 Set optFormatRAWControl = Controls.Add("PDFCreator.ctlOptFormatRAW", "ctlOptFormatRAW")
 optFormatRAWControl.Width = dmFraDescription.Width
 Set optFormatRAW = optFormatRAWControl.object
 optFormatRAW.SetLanguageStrings
 optFormatRAW.SetOptions
 ' Add FormatSVGControl
 Set optFormatSVGControl = Controls.Add("PDFCreator.ctlOptFormatSVG", "ctlOptFormatSVG")
 optFormatSVGControl.Width = dmFraDescription.Width
 Set optFormatSVG = optFormatSVGControl.object
 optFormatSVG.SetLanguageStrings
 optFormatSVG.SetOptions
' ' Add FormatXCFControl - Doesn't work
' Set optFormatXCFControl = Controls.Add("PDFCreator.ctlOptFormatXCF", "ctlOptFormatXCF")
' optFormatXCFControl.Width = dmFraDescription.Width
' Set optFormatXCF = optFormatXCFControl.object
' optFormatXCF.SetLanguageStrings
' optFormatXCF.SetOptions
 ' Add GhostscriptControl
 Set optGhostscriptControl = Controls.Add("PDFCreator.ctlOptGhostscript", "ctlOptGhostscript")
 optGhostscriptControl.Width = dmFraDescription.Width
 Set optGhostscript = optGhostscriptControl.object
 optGhostscript.SetLanguageStrings
 optGhostscript.SetOptions
 ' Add LanguagesControl
 Set optLanguagesControl = Controls.Add("PDFCreator.ctlOptLanguages", "ctlOptLanguages")
 optLanguagesControl.Width = dmFraDescription.Width
 Set optLanguages = optLanguagesControl.object
 optLanguages.SetLanguageStrings
 optLanguages.SetOptions
 ' Add PrintControl
 Set optPrintControl = Controls.Add("PDFCreator.ctlOptPrint", "ctlOptPrint")
 optPrintControl.Width = dmFraDescription.Width
 Set optPrint = optPrintControl.object
 optPrint.SetLanguageStrings
 optPrint.SetOptions
 ' Add SaveControl
 '
 Set optSaveControl = Controls.Add("PDFCreator.ctlOptSave", "ctlOptSave")
 optSaveControl.Width = dmFraDescription.Width
 Set optSave = optSaveControl.object
 optSave.SetLanguageStrings
 optSave.SetOptions
 ' Add GeneralControl
 '
 Set optGeneralControl = Controls.Add("PDFCreator.ctlOptGeneral", "ctlOptGeneral")
 optGeneralControl.Width = dmFraDescription.Width
 Set optGeneral = optGeneralControl.object
 optGeneral.SetLanguageStrings
'
 optGeneral.SetOptions

 dmFraProfile.Caption = LanguageStrings.OptionsProfile
 cmbProfile.Clear
 cmbProfile.AddItem LanguageStrings.OptionsProfileDefaultName

 Set Profiles = GetProfiles
 ReDim ProfileNames(Profiles.Count)
 ReDim ProfileOptions(Profiles.Count)
 ProfileNames(0) = LanguageStrings.OptionsProfileDefaultName
 ProfileOptions(0) = Options

 With dmFraDescription
  .Caption = LanguageStrings.OptionsTreeProgram
  .Visible = True

  optActionsControl.Top = .Top + .Height + ControlTop
  optActionsControl.Left = .Left
  optActionsControl.Width = .Width
  optAutosaveControl.Top = .Top + .Height + ControlTop
  optAutosaveControl.Left = .Left
  optAutosaveControl.Width = .Width
'  optDirectoriesControl.Top = .Top + .Height + ControlTop
'  optDirectoriesControl.Left = .Left
'  optDirectoriesControl.Width = .Width
  optDocumentControl.Top = .Top + .Height + ControlTop
  optDocumentControl.Left = .Left
  optDocumentControl.Width = .Width
  optFontsControl.Top = .Top + .Height + ControlTop
  optFontsControl.Left = .Left
  optFontsControl.Width = .Width
  optFormatPNGControl.Top = .Top + .Height + ControlTop
  optFormatPNGControl.Left = .Left
  optFormatPNGControl.Width = .Width
  optFormatJPEGControl.Top = .Top + .Height + ControlTop
  optFormatJPEGControl.Left = .Left
  optFormatJPEGControl.Width = .Width
  optFormatBMPControl.Top = .Top + .Height + ControlTop
  optFormatBMPControl.Left = .Left
  optFormatBMPControl.Width = .Width
  optFormatPCXControl.Top = .Top + .Height + ControlTop
  optFormatPCXControl.Left = .Left
  optFormatPCXControl.Width = .Width
  optFormatTIFFControl.Top = .Top + .Height + ControlTop
  optFormatTIFFControl.Left = .Left
  optFormatTIFFControl.Width = .Width
  optFormatPDFControl.Top = .Top + .Height + ControlTop
  optFormatPDFControl.Left = .Left
  optFormatPDFControl.Width = .Width
  optFormatPSControl.Top = .Top + .Height + ControlTop
  optFormatPSControl.Left = .Left
  optFormatPSControl.Width = .Width
  optFormatEPSControl.Top = .Top + .Height + ControlTop
  optFormatEPSControl.Left = .Left
  optFormatEPSControl.Width = .Width
  optFormatTXTControl.Top = .Top + .Height + ControlTop
  optFormatTXTControl.Left = .Left
  optFormatTXTControl.Width = .Width
  optFormatPSDControl.Top = .Top + .Height + ControlTop
  optFormatPSDControl.Left = .Left
  optFormatPSDControl.Width = .Width
  optFormatPCLControl.Top = .Top + .Height + ControlTop
  optFormatPCLControl.Left = .Left
  optFormatPCLControl.Width = .Width
  optFormatRAWControl.Top = .Top + .Height + ControlTop
  optFormatRAWControl.Left = .Left
  optFormatRAWControl.Width = .Width
  optFormatSVGControl.Top = .Top + .Height + ControlTop
  optFormatSVGControl.Left = .Left
  optFormatSVGControl.Width = .Width
'  optFormatXCFControl.Top = .Top + .Height + ControlTop
'  optFormatXCFControl.Left = .Left
'  optFormatXCFControl.Width = .Width
  optGeneralControl.Top = .Top + .Height + ControlTop
  optGeneralControl.Left = .Left
  optGeneralControl.Width = .Width
  optGhostscriptControl.Top = .Top + .Height + ControlTop
  optGhostscriptControl.Left = .Left
  optGhostscriptControl.Width = .Width
  optLanguagesControl.Top = .Top + .Height + ControlTop
  optLanguagesControl.Left = .Left
  optLanguagesControl.Width = .Width
  optPrintControl.Top = .Top + .Height + ControlTop
  optPrintControl.Left = .Left
  optPrintControl.Width = .Width
  optSaveControl.Top = .Top + .Height + ControlTop
  optSaveControl.Left = .Left
  optSaveControl.Width = .Width

  cmdCancel.Left = .Left
  cmdReset.Left = .Left + (.Width - cmdReset.Width) / 2
  cmdSave.Left = .Left + .Width - cmdSave.Width
 End With

 For i = 1 To Profiles.Count
  cmbProfile.AddItem Profiles(i)
  ProfileNames(i) = Profiles(i)
  ProfileOptions(i) = ReadOptions(, , Profiles(i))
 Next i
 SetProfile CurrentPrinterProfile

 If cmbProfile.ListIndex = 0 Then
   optGhostscript.ControlEnabled = True
   optLanguages.ControlEnabled = True
   cmdProfileRename.Enabled = False
   cmdProfileDelete.Enabled = False
  Else
   optGhostscript.ControlEnabled = False
   optLanguages.ControlEnabled = False
   cmdProfileRename.Enabled = True
   cmdProfileDelete.Enabled = True
 End If

 Set PrinterProfiles = GetPrinterProfiles
 For i = 1 To PrinterProfiles.Count
  sa(0) = PrinterProfiles(i)(0)
  sa(1) = PrinterProfiles(i)(1)
  sa(2) = PrinterProfiles(i)(1)
  TempPrinterProfiles.Add sa
 Next i

 With LanguageStrings
  cmdProfileAdd.ToolTipText = .OptionsProfileAdd
  cmdProfileDelete.ToolTipText = .OptionsProfileDel
  cmdProfileRename.ToolTipText = .OptionsProfileRenameProfile
  cmdProfileSave.ToolTipText = .OptionsProfileSaveToDisc
  cmdProfileLoad.ToolTipText = .OptionsProfileLoadFromDisc
  cmbProfile.List(0) = .OptionsProfileDefaultName
 End With

 If ShowOnlyOptions = True Then
  FormInTaskbar Me, True, True
  Caption = "PDFCreator - " & Caption
 End If

 ShowAcceleratorsInForm Me, True

 Screen.MousePointer = vbNormal

 With Options
  SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
 End With

 LastNodeKey = ""
 trvOptions.Nodes("Program").Selected = True
 trvOptions_NodeClick trvOptions.Nodes("Program")
 LoadReady = True
End Sub

Private Sub SetOptions()
 optActions.SetOptions
 optAutosave.SetOptions
' optDirectories.SetOptions
 optDocument.SetOptions
 optFonts.SetOptions
 optFormatPNG.SetOptions
 optFormatJPEG.SetOptions
 optFormatBMP.SetOptions
 optFormatPCX.SetOptions
 optFormatTIFF.SetOptions
 optFormatPDF.SetOptions
 optFormatPS.SetOptions
 optFormatEPS.SetOptions
 optFormatTXT.SetOptions
 optFormatPSD.SetOptions
 optFormatPCL.SetOptions
 optFormatRAW.SetOptions
 optFormatSVG.SetOptions
' optFormatXCF.SetOptions
 optGhostscript.SetOptions
 optLanguages.SetOptions
 optPrint.SetOptions
 optSave.SetOptions
 optGeneral.SetOptions

 With Options
  SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
 End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 UnloadForm = True
 Me.Visible = False
End Sub

Private Sub trvOptions_NodeClick(ByVal Node As MSComctlLib.Node)
 Dim ctl As Control

 If LastNodeKey = Node.key Or (LastNodeKey = "Program" And Node.key = "ProgramGeneral") Or (LastNodeKey = "ProgramGeneral" And Node.key = "Program") Then
   Exit Sub
  Else
   LastNodeKey = Node.key
 End If

 optActionsControl.Visible = False
 optAutosaveControl.Visible = False
' optDirectoriesControl.Visible = False
 optDocumentControl.Visible = False
 optFontsControl.Visible = False
 optFormatPNGControl.Visible = False
 optFormatJPEGControl.Visible = False
 optFormatBMPControl.Visible = False
 optFormatPCXControl.Visible = False
 optFormatTIFFControl.Visible = False
 optFormatPDFControl.Visible = False
 optFormatPSControl.Visible = False
 optFormatEPSControl.Visible = False
 optFormatTXTControl.Visible = False
 optFormatPSDControl.Visible = False
 optFormatPCLControl.Visible = False
 optFormatRAWControl.Visible = False
 optFormatSVGControl.Visible = False
' optFormatXCFControl.Visible = False
 optGeneralControl.Visible = False
 optGhostscriptControl.Visible = False
 optLanguagesControl.Visible = False
 optPrintControl.Visible = False
 optSaveControl.Visible = False

 Select Case UCase$(Node.key)
  Case "PROGRAM", "PROGRAMGENERAL"
   Set picOptions = LoadResPicture(2101, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsProgramGeneralDescription
   optGeneralControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
  Case "PROGRAMGHOSTSCRIPT"
   Set picOptions = LoadResPicture(2119, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsProgramGhostscriptDescription
   optGhostscriptControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
  Case "PROGRAMDOCUMENT"
   Set picOptions = LoadResPicture(2105, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsProgramDocumentDescription
   optDocumentControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
  Case "PROGRAMSAVE"
   Set picOptions = LoadResPicture(2106, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsProgramSaveDescription

   optSaveControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
  Case "PROGRAMAUTOSAVE"
   Set picOptions = LoadResPicture(2103, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsProgramAutosaveDescription
   optAutosaveControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
'  Case "PROGRAMDIRECTORIES"
'   Set picOptions = LoadResPicture(2104, vbResIcon)
'   lblOptions.Caption = LanguageStrings.OptionsProgramDirectoriesDescription
'   optDirectoriesControl.Visible = True
'   dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
  Case "PROGRAMACTIONS"
   Set picOptions = LoadResPicture(2121, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsProgramActionsDescription
   optActionsControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
  Case "PROGRAMPRINT"
   Set picOptions = LoadResPicture(2122, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsProgramPrintDescription
   optPrintControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
  Case "PROGRAMFONTS"
   Set picOptions = LoadResPicture(2102, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsProgramFontDescription
   optFontsControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
  Case "PROGRAMLANGUAGES"
   Set picOptions = LoadResPicture(2123, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsProgramLanguagesDescription
   optLanguagesControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
  Case "FORMATS", "FORMATSPDF"
   Set picOptions = LoadResPicture(2111, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsPDFDescription
   optFormatPDFControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
  Case "FORMATSPNG"
   Set picOptions = LoadResPicture(2112, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsPNGDescription
   optFormatPNGControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
  Case "FORMATSJPEG"
   Set picOptions = LoadResPicture(2113, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsJPEGDescription
   optFormatJPEGControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
  Case "FORMATSBMP"
   Set picOptions = LoadResPicture(2114, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsBMPDescription
   optFormatBMPControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
  Case "FORMATSPCX"
   Set picOptions = LoadResPicture(2115, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsPCXDescription
   optFormatPCXControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
  Case "FORMATSTIFF"
   Set picOptions = LoadResPicture(2116, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsTIFFDescription
   optFormatTIFFControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
  Case "FORMATSPS"
   Set picOptions = LoadResPicture(2117, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsPSDescription
   optFormatPSControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
  Case "FORMATSEPS"
   Set picOptions = LoadResPicture(2118, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsEPSDescription
   optFormatEPSControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
  Case "FORMATSTXT"
   Set picOptions = LoadResPicture(2124, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsTXTDescription
   optFormatTXTControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
  Case "FORMATSPSD"
   Set picOptions = LoadResPicture(2125, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsPSDDescription
   optFormatPSDControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
  Case "FORMATSPCL"
   Set picOptions = LoadResPicture(2126, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsPCLDescription
   optFormatPCLControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
  Case "FORMATSRAW"
   Set picOptions = LoadResPicture(2127, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsRAWDescription
   optFormatRAWControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
  Case "FORMATSSVG"
   Set picOptions = LoadResPicture(2129, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsSVGDescription
   optFormatSVGControl.Visible = True
   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
'  Case "FORMATSXCF"
'   Set picOptions = LoadResPicture(2128, vbResIcon)
'   lblOptions.Caption = LanguageStrings.OptionsXCFDescription
'   optFormatXCFControl.Visible = True
'   dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
 End Select
End Sub

Public Sub SetFrames(OptionsDesign As Long)
 optActions.SetFrames OptionsDesign
 optAutosave.SetFrames OptionsDesign
' optDirectories.SetFrames OptionsDesign
 optDocument.SetFrames OptionsDesign
 optFonts.SetFrames OptionsDesign
 optFormatBMP.SetFrames OptionsDesign
 optFormatEPS.SetFrames OptionsDesign
 optFormatJPEG.SetFrames OptionsDesign
 optFormatPCL.SetFrames OptionsDesign
 optFormatPCX.SetFrames OptionsDesign
 optFormatPDF.SetFrames OptionsDesign
 optFormatPNG.SetFrames OptionsDesign
 optFormatPS.SetFrames OptionsDesign
 optFormatPSD.SetFrames OptionsDesign
 optFormatRAW.SetFrames OptionsDesign
 optFormatSVG.SetFrames OptionsDesign
 optFormatTIFF.SetFrames OptionsDesign
 optFormatTXT.SetFrames OptionsDesign
 optGeneral.SetFrames OptionsDesign
 optGhostscript.SetFrames OptionsDesign
 optLanguages.SetFrames OptionsDesign
 optPrint.SetFrames OptionsDesign
 optSave.SetFrames OptionsDesign
End Sub

Private Sub SetProfile(Optional ByVal ProfileName As String = "") ' Empty profilename for default profile
 Dim i As Long
 If cmbProfile.ListIndex < 0 Then
   OldProfile = 0
  Else
   OldProfile = cmbProfile.ListIndex
 End If
 ProfileName = Trim$(ProfileName)
 If LenB(ProfileName) = 0 Then
  cmbProfile.ListIndex = 0
  Exit Sub
 End If
 For i = 1 To cmbProfile.ListCount - 1
  If LCase$(ProfileName) = LCase$(cmbProfile.List(i)) Then
   cmbProfile.ListIndex = i
   Exit Sub
  End If
 Next i
 cmbProfile.ListIndex = 0
End Sub
