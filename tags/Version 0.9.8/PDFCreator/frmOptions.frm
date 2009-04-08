VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Options"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin PDFCreator.dmFrame dmFraDescription 
      Height          =   1065
      Left            =   2760
      TabIndex        =   3
      Top             =   105
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
      Top             =   7320
      Width           =   1815
   End
   Begin PDFCreator.isExplorerBar ieb 
      Align           =   3  'Links ausrichten
      Height          =   7935
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   13996
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
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   7320
      Width           =   1815
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
Private optDirectoriesControl As VBControlExtender, optDirectories As ctlOptDirectories
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyCode = vbKeyF1 Then
50020   KeyCode = 0
50030     Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50041    Select Case ieb.GetSelectedGroup
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
50170       Case 6
50180        Call HTMLHelp_ShowTopic("html\directories.htm")
50190       Case 7
50200        Call HTMLHelp_ShowTopic("html\fontsetting.htm")
50210       Case Else
50220        Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50230      End Select
50240     Case 2
50251      Select Case ieb.GetSelectedItem
            Case 1
50271        Select Case optFormatPDF.PDFOptionsIndex
              Case 1
50290          Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50300         Case 2
50310          Call HTMLHelp_ShowTopic("html\pdfcompression.htm")
50320         Case 3
50330          Call HTMLHelp_ShowTopic("html\pdffonts.htm")
50340         Case 4
50350          Call HTMLHelp_ShowTopic("html\pdfcolors.htm")
50360         Case 5
50370          Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
50380         Case 6
50390          Call HTMLHelp_ShowTopic("html\pdfsigning.htm")
50400         Case Else
50410          Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50420        End Select
50430       Case 2
50440        Call HTMLHelp_ShowTopic("html\pngsettings.htm")
50450       Case 3
50460        Call HTMLHelp_ShowTopic("html\jpegsettings.htm")
50470       Case 4
50480        Call HTMLHelp_ShowTopic("html\bmpsettings.htm")
50490       Case 5
50500        Call HTMLHelp_ShowTopic("html\pcxsettings.htm")
50510       Case 6
50520        Call HTMLHelp_ShowTopic("html\tiffsettings.htm")
50530       Case 7
50540        Call HTMLHelp_ShowTopic("html\pssettings.htm")
50550       Case 8
50560        Call HTMLHelp_ShowTopic("html\epssettings.htm")
50570       Case Else
50580        Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50590      End Select
50600    End Select
50610  End If
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
50060   Me.Caption = .DialogPrinterOptions
50070   cmdCancel.Caption = .OptionsCancel
50080   cmdReset.Caption = .OptionsReset
50090   cmdSave.Caption = .OptionsSave
50100
50110   ieb.DisableUpdates True
50120   ieb.SetGroupCaption "Program", .OptionsTreeProgram
50130   ieb.SetItemText "Program", "General", .OptionsProgramGeneralSymbol
50140   ieb.SetItemText "Program", "Ghostscript", .OptionsProgramGhostscriptSymbol
50150   ieb.SetItemText "Program", "Document", .OptionsProgramDocumentSymbol
50160   ieb.SetItemText "Program", "Save", .OptionsProgramSaveSymbol
50170   ieb.SetItemText "Program", "AutoSave", .OptionsProgramAutosaveSymbol
50180   ieb.SetItemText "Program", "Directories", .OptionsProgramDirectoriesSymbol
50190
50200   ieb.SetItemText "Program", "Actions", .OptionsProgramActionsSymbol
50210   ieb.SetItemText "Program", "Print", .OptionsProgramPrintSymbol
50220   ieb.SetItemText "Program", "Fonts", .OptionsProgramFontSymbol
50230   ieb.SetItemText "Program", "Language", .OptionsProgramLanguagesSymbol
50240
50250   ieb.SetGroupCaption "Formats", .OptionsTreeFormats
50260   ieb.SetItemText "Formats", "PDF", .OptionsPDFSymbol
50270   ieb.SetItemText "Formats", "PNG", .OptionsPNGSymbol
50280   ieb.SetItemText "Formats", "JPEG", .OptionsJPEGSymbol
50290   ieb.SetItemText "Formats", "BMP", .OptionsBMPSymbol
50300   ieb.SetItemText "Formats", "PCX", .OptionsPCXSymbol
50310   ieb.SetItemText "Formats", "TIFF", .OptionsTIFFSymbol
50320   ieb.SetItemText "Formats", "PS", .OptionsPSSymbol
50330   ieb.SetItemText "Formats", "EPS", .OptionsEPSSymbol
50340   ieb.SetItemText "Formats", "TXT", .OptionsTXTSymbol
50350   ieb.SetItemText "Formats", "PSD", .OptionsPSDSymbol
50360   ieb.SetItemText "Formats", "PCL", .OptionsPCLSymbol
50370   ieb.SetItemText "Formats", "RAW", .OptionsRAWSymbol
50380   ieb.SetItemText "Formats", "SVG", .OptionsSVGSymbol
50390 '  ieb.SetItemText "Formats", "XCF", .OptionsXCFSymbol
50400   ieb.DisableUpdates False
50410
50420   lblOptions.Caption = .OptionsProgramLanguagesDescription
50430  End With
50440  optActions.SetLanguageStrings
50450  optAutosave.SetLanguageStrings
50460  optDirectories.SetLanguageStrings
50470  optDocument.SetLanguageStrings
50480  optFonts.SetLanguageStrings
50490  optFormatPNG.SetLanguageStrings
50500  optFormatJPEG.SetLanguageStrings
50510  optFormatBMP.SetLanguageStrings
50520  optFormatPCX.SetLanguageStrings
50530  optFormatTIFF.SetLanguageStrings
50540  optFormatPDF.SetLanguageStrings
50550  optFormatPS.SetLanguageStrings
50560  optFormatEPS.SetLanguageStrings
50570  optFormatTXT.SetLanguageStrings
50580  optFormatPSD.SetLanguageStrings
50590  optFormatPCL.SetLanguageStrings
50600  optFormatRAW.SetLanguageStrings
50610  optFormatSVG.SetLanguageStrings
50620 ' optFormatXCF.SetLanguageStrings
50630
50640  optGeneral.SetLanguageStrings
50650  optGhostscript.SetLanguageStrings
50660  optLanguages.SetLanguageStrings
50670  optPrint.SetLanguageStrings
50680  optSave.SetLanguageStrings
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
50030  If oldLanguage <> Options.Language Then
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
50080   optDirectories.SetOptions
50090   optDocument.SetOptions
50100   optFonts.SetOptions
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

Private Sub cmdSave_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tRestart As Boolean, tOpt As tOptions, newLanguage As String
50020  tRestart = False
50030  If UCase$(Options.DirectoryGhostscriptBinaries) <> UCase$(optGhostscript.GSBin) Then
50040   tRestart = True
50050  End If
50060
50070  tOpt = Options
50080
50090  optActions.GetOptions
50100  optAutosave.GetOptions
50110  optDirectories.GetOptions
50120  optDocument.GetOptions
50130  optFonts.GetOptions
50140  optFormatPNG.GetOptions
50150  optFormatJPEG.GetOptions
50160  optFormatBMP.GetOptions
50170  optFormatPCX.GetOptions
50180  optFormatTIFF.GetOptions
50190  optFormatPDF.GetOptions
50200  optFormatPS.GetOptions
50210  optFormatEPS.GetOptions
50220  optFormatTXT.GetOptions
50230  optFormatPSD.GetOptions
50240  optFormatPCL.GetOptions
50250  optFormatRAW.GetOptions
50260  optFormatSVG.GetOptions
50270 ' optFormatXCF.GetOptions
50280
50290  optGeneral.GetOptions
50300  optGhostscript.GetOptions
50310  optLanguages.GetOptions
50320  optPrint.GetOptions
50330  optSave.GetOptions
50340
50350 ' newLanguage = Options.Language
50360 ' Options.Language = newLanguage
50370 ' Options.StampFontname = tOpt.StampFontname
50380 ' Options.StampFontsize = tOpt.StampFontsize
50390  SaveOptions Options
50400
50410  SetHelpfile
50420
50430  If IsWin9xMe = False Then
50441   Select Case Options.ProcessPriority
         Case 0: 'Idle
50460     SetProcessPriority Idle
50470    Case 1: 'Normal
50480     SetProcessPriority Normal
50490    Case 2: 'High
50500     SetProcessPriority High
50510    Case 3: 'Realtime
50520     SetProcessPriority RealTime
50530   End Select
50540  End If
50550  If tRestart = True Then
50560   Restart = True
50570  End If
50580  Unload Me
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
50050  Dim ctl1 As Control, ctl2 As Control
50060
50070  UnloadForm = False
50080  Me.Icon = LoadResPicture(2120, vbResIcon)
50090  KeyPreview = True
50100
50110  oldLanguage = Options.Language
50120
50130  With Screen
50140   .MousePointer = vbHourglass
50150   Move (.Width - Width) / 2, (.Height - Height) / 2
50160  End With
50170
50180  LastProgramGroup = ""
50190  LastFormatGroup = ""
50200
50210  ieb.DisableUpdates True
50220  ieb.ClearStructure
50230  ieb.SetImageList imlIeb
50240  With LanguageStrings
50250   ieb.AddGroup "Program", .OptionsTreeProgram, 0
50260   ieb.AddItem "Program", "General", .OptionsProgramGeneralSymbol, 1
50270   ieb.AddItem "Program", "Ghostscript", .OptionsProgramGhostscriptSymbol, 2
50280   ieb.AddItem "Program", "Document", .OptionsProgramDocumentSymbol, 3
50290   ieb.AddItem "Program", "Save", .OptionsProgramSaveSymbol, 4
50300   ieb.AddItem "Program", "AutoSave", .OptionsProgramAutosaveSymbol, 5
50310   ieb.AddItem "Program", "Directories", .OptionsProgramDirectoriesSymbol, 6
50320   ieb.AddItem "Program", "Actions", .OptionsProgramActionsSymbol, 7
50330   ieb.AddItem "Program", "Print", .OptionsProgramPrintSymbol, 8
50340   ieb.AddItem "Program", "Fonts", .OptionsProgramFontSymbol, 9
50350   ieb.AddItem "Program", "Language", .OptionsProgramLanguagesSymbol, 10
50360
50370   ieb.AddGroup "Formats", .OptionsTreeFormats, 0
50380   ieb.AddItem "Formats", "PDF", .OptionsPDFSymbol, 11
50390   ieb.AddItem "Formats", "PNG", .OptionsPNGSymbol, 12
50400   ieb.AddItem "Formats", "JPEG", .OptionsJPEGSymbol, 13
50410   ieb.AddItem "Formats", "BMP", .OptionsBMPSymbol, 14
50420   ieb.AddItem "Formats", "PCX", .OptionsPCXSymbol, 15
50430   ieb.AddItem "Formats", "TIFF", .OptionsTIFFSymbol, 16
50440   ieb.AddItem "Formats", "PS", .OptionsPSSymbol, 17
50450   ieb.AddItem "Formats", "EPS", .OptionsEPSSymbol, 18
50460   ieb.AddItem "Formats", "TXT", .OptionsTXTSymbol, 21
50470   ieb.AddItem "Formats", "PSD", .OptionsPSDSymbol, 22
50480   ieb.AddItem "Formats", "PCL", .OptionsPCLSymbol, 23
50490   ieb.AddItem "Formats", "RAW", .OptionsRAWSymbol, 24
50500 '  ieb.AddItem "Formats", "XCF", .OptionsXCFSymbol, 25
50510   ieb.AddItem "Formats", "SVG", .OptionsSVGSymbol, 26
50520   ieb.ExpandGroup "Formats", False
50530
50540   ieb.DisableUpdates False
50550
50560   Set picOptions = LoadResPicture(2101, vbResIcon)
50570
50580   Me.Caption = .DialogPrinterOptions
50590   cmdCancel.Caption = .OptionsCancel
50600   cmdReset.Caption = .OptionsReset
50610   cmdSave.Caption = .OptionsSave
50620  End With
50630
50640  SetFrame dmFraDescription
50650
50660  ' Add ActionsControl
50670  Set optActionsControl = Controls.Add("PDFCreator.ctlOptActions", "ctlOptActions")
50680  optActionsControl.Width = dmFraDescription.Width
50690  Set optActions = optActionsControl.object
50700  optActions.SetLanguageStrings
50710  optActions.SetOptions
50720  ' Add AutosaveControl
50730  Set optAutosaveControl = Controls.Add("PDFCreator.ctlOptAutosave", "ctlOptAutosave")
50740  optAutosaveControl.Width = dmFraDescription.Width
50750  Set optAutosave = optAutosaveControl.object
50760  optAutosave.SetLanguageStrings
50770  optAutosave.SetOptions
50780  ' Add DirectoriesControl
50790  Set optDirectoriesControl = Controls.Add("PDFCreator.ctlOptDirectories", "ctlOptDirectories")
50800  optDirectoriesControl.Width = dmFraDescription.Width
50810  Set optDirectories = optDirectoriesControl.object
50820  optDirectories.SetLanguageStrings
50830  optDirectories.SetOptions
50840  ' Add DocumentControl
50850  Set optDocumentControl = Controls.Add("PDFCreator.ctlOptDocument", "ctlOptDocument")
50860  optDocumentControl.Width = dmFraDescription.Width
50870  Set optDocument = optDocumentControl.object
50880  optDocument.SetLanguageStrings
50890  optDocument.SetOptions
50900  ' Add FontsControl
50910  Set optFontsControl = Controls.Add("PDFCreator.ctlOptFonts", "ctlOptFonts")
50920  optFontsControl.Width = dmFraDescription.Width
50930  Set optFonts = optFontsControl.object
50940  optFonts.SetLanguageStrings
50950  optFonts.SetOptions
50960  ' Add FormatPNGControl
50970  Set optFormatPNGControl = Controls.Add("PDFCreator.ctlOptFormatPNG", "ctlOptFormatPNG")
50980  optFormatPNGControl.Width = dmFraDescription.Width
50990  Set optFormatPNG = optFormatPNGControl.object
51000  optFormatPNG.SetLanguageStrings
51010  optFormatPNG.SetOptions
51020  ' Add FormatJPEQControl
51030  Set optFormatJPEGControl = Controls.Add("PDFCreator.ctlOptFormatJPEG", "ctlOptFormatJPEG")
51040  optFormatJPEGControl.Width = dmFraDescription.Width
51050  Set optFormatJPEG = optFormatJPEGControl.object
51060  optFormatJPEG.SetLanguageStrings
51070  optFormatJPEG.SetOptions
51080  ' Add FormatBMPControl
51090  Set optFormatBMPControl = Controls.Add("PDFCreator.ctlOptFormatBMP", "ctlOptFormatBMP")
51100  optFormatBMPControl.Width = dmFraDescription.Width
51110  Set optFormatBMP = optFormatBMPControl.object
51120  optFormatBMP.SetLanguageStrings
51130  optFormatBMP.SetOptions
51140  ' Add FormatPCXControl
51150  Set optFormatPCXControl = Controls.Add("PDFCreator.ctlOptFormatPCX", "ctlOptFormatPCX")
51160  optFormatPCXControl.Width = dmFraDescription.Width
51170  Set optFormatPCX = optFormatPCXControl.object
51180  optFormatPCX.SetLanguageStrings
51190  optFormatPCX.SetOptions
51200  ' Add FormatTIFFControl
51210  Set optFormatTIFFControl = Controls.Add("PDFCreator.ctlOptFormatTIFF", "ctlOptFormatTIFF")
51220  optFormatTIFFControl.Width = dmFraDescription.Width
51230  Set optFormatTIFF = optFormatTIFFControl.object
51240  optFormatTIFF.SetLanguageStrings
51250  optFormatTIFF.SetOptions
51260  ' Add FormatPDFControl
51270  Set optFormatPDFControl = Controls.Add("PDFCreator.ctlOptFormatPDF", "ctlOptFormatPDF")
51280  optFormatPDFControl.Width = dmFraDescription.Width
51290  Set optFormatPDF = optFormatPDFControl.object
51300  optFormatPDF.SetLanguageStrings
51310  optFormatPDF.SetOptions
51320  ' Add FormatPS
51330  Set optFormatPSControl = Controls.Add("PDFCreator.ctlOptFormatPS", "ctlOptFormatPS")
51340  optFormatPSControl.Width = dmFraDescription.Width
51350  Set optFormatPS = optFormatPSControl.object
51360  optFormatPS.SetLanguageStrings
51370  optFormatPS.SetOptions
51380  ' Add FormatEPSControl
51390  Set optFormatEPSControl = Controls.Add("PDFCreator.ctlOptFormatEPS", "ctlOptFormatEPS")
51400  optFormatEPSControl.Width = dmFraDescription.Width
51410  Set optFormatEPS = optFormatEPSControl.object
51420  optFormatEPS.SetLanguageStrings
51430  optFormatEPS.SetOptions
51440  ' Add FormatTXTControl
51450  Set optFormatTXTControl = Controls.Add("PDFCreator.ctlOptFormatTXT", "ctlOptFormatTXT")
51460  optFormatTXTControl.Width = dmFraDescription.Width
51470  Set optFormatTXT = optFormatTXTControl.object
51480  optFormatTXT.SetLanguageStrings
51490  optFormatTXT.SetOptions
51500  ' Add FormatPSDControl
51510  Set optFormatPSDControl = Controls.Add("PDFCreator.ctlOptFormatPSD", "ctlOptFormatPSD")
51520  optFormatPSDControl.Width = dmFraDescription.Width
51530  Set optFormatPSD = optFormatPSDControl.object
51540  optFormatPSD.SetLanguageStrings
51550  optFormatPSD.SetOptions
51560  ' Add FormatPCLControl
51570  Set optFormatPCLControl = Controls.Add("PDFCreator.ctlOptFormatPCL", "ctlOptFormatPCL")
51580  optFormatPCLControl.Width = dmFraDescription.Width
51590  Set optFormatPCL = optFormatPCLControl.object
51600  optFormatPCL.SetLanguageStrings
51610  optFormatPCL.SetOptions
51620  ' Add FormatRAWControl
51630  Set optFormatRAWControl = Controls.Add("PDFCreator.ctlOptFormatRAW", "ctlOptFormatRAW")
51640  optFormatRAWControl.Width = dmFraDescription.Width
51650  Set optFormatRAW = optFormatRAWControl.object
51660  optFormatRAW.SetLanguageStrings
51670  optFormatRAW.SetOptions
51680  ' Add FormatSVGControl
51690  Set optFormatSVGControl = Controls.Add("PDFCreator.ctlOptFormatSVG", "ctlOptFormatSVG")
51700  optFormatSVGControl.Width = dmFraDescription.Width
51710  Set optFormatSVG = optFormatSVGControl.object
51720  optFormatSVG.SetLanguageStrings
51730  optFormatSVG.SetOptions
51740 ' ' Add FormatXCFControl - Doesn't work
51750 ' Set optFormatXCFControl = Controls.Add("PDFCreator.ctlOptFormatXCF", "ctlOptFormatXCF")
51760 ' optFormatXCFControl.Width = dmFraDescription.Width
51770 ' Set optFormatXCF = optFormatXCFControl.object
51780 ' optFormatXCF.SetLanguageStrings
51790 ' optFormatXCF.SetOptions
51800  ' Add GhostscriptControl
51810  Set optGhostscriptControl = Controls.Add("PDFCreator.ctlOptGhostscript", "ctlOptGhostscript")
51820  optGhostscriptControl.Width = dmFraDescription.Width
51830  Set optGhostscript = optGhostscriptControl.object
51840  optGhostscript.SetLanguageStrings
51850  optGhostscript.SetOptions
51860  ' Add LanguagesControl
51870  Set optLanguagesControl = Controls.Add("PDFCreator.ctlOptLanguages", "ctlOptLanguages")
51880  optLanguagesControl.Width = dmFraDescription.Width
51890  Set optLanguages = optLanguagesControl.object
51900  optLanguages.SetLanguageStrings
51910  optLanguages.SetOptions
51920  ' Add PrintControl
51930  Set optPrintControl = Controls.Add("PDFCreator.ctlOptPrint", "ctlOptPrint")
51940  optPrintControl.Width = dmFraDescription.Width
51950  Set optPrint = optPrintControl.object
51960  optPrint.SetLanguageStrings
51970  optPrint.SetOptions
51980  ' Add SaveControl
51990  Set optSaveControl = Controls.Add("PDFCreator.ctlOptSave", "ctlOptSave")
52000  optSaveControl.Width = dmFraDescription.Width
52010  Set optSave = optSaveControl.object
52020  optSave.SetLanguageStrings
52030  optSave.SetOptions
52040  ' Add GeneralControl
52050  Set optGeneralControl = Controls.Add("PDFCreator.ctlOptGeneral", "ctlOptGeneral")
52060  optGeneralControl.Width = dmFraDescription.Width
52070  Set optGeneral = optGeneralControl.object
52080  optGeneral.SetLanguageStrings
52090  optGeneral.SetOptions
52100
52110  With dmFraDescription
52120   .Caption = LanguageStrings.OptionsTreeProgram
52130   .Visible = True
52140
52150   optActionsControl.Top = .Top + .Height + ControlTop
52160   optActionsControl.Left = .Left
52170   optActionsControl.Width = .Width
52180   optAutosaveControl.Top = .Top + .Height + ControlTop
52190   optAutosaveControl.Left = .Left
52200   optAutosaveControl.Width = .Width
52210   optDirectoriesControl.Top = .Top + .Height + ControlTop
52220   optDirectoriesControl.Left = .Left
52230   optDirectoriesControl.Width = .Width
52240   optDocumentControl.Top = .Top + .Height + ControlTop
52250   optDocumentControl.Left = .Left
52260   optDocumentControl.Width = .Width
52270   optFontsControl.Top = .Top + .Height + ControlTop
52280   optFontsControl.Left = .Left
52290   optFontsControl.Width = .Width
52300   optFormatPNGControl.Top = .Top + .Height + ControlTop
52310   optFormatPNGControl.Left = .Left
52320   optFormatPNGControl.Width = .Width
52330   optFormatJPEGControl.Top = .Top + .Height + ControlTop
52340   optFormatJPEGControl.Left = .Left
52350   optFormatJPEGControl.Width = .Width
52360   optFormatBMPControl.Top = .Top + .Height + ControlTop
52370   optFormatBMPControl.Left = .Left
52380   optFormatBMPControl.Width = .Width
52390   optFormatPCXControl.Top = .Top + .Height + ControlTop
52400   optFormatPCXControl.Left = .Left
52410   optFormatPCXControl.Width = .Width
52420   optFormatTIFFControl.Top = .Top + .Height + ControlTop
52430   optFormatTIFFControl.Left = .Left
52440   optFormatTIFFControl.Width = .Width
52450   optFormatPDFControl.Top = .Top + .Height + ControlTop
52460   optFormatPDFControl.Left = .Left
52470   optFormatPDFControl.Width = .Width
52480   optFormatPSControl.Top = .Top + .Height + ControlTop
52490   optFormatPSControl.Left = .Left
52500   optFormatPSControl.Width = .Width
52510   optFormatEPSControl.Top = .Top + .Height + ControlTop
52520   optFormatEPSControl.Left = .Left
52530   optFormatEPSControl.Width = .Width
52540   optFormatTXTControl.Top = .Top + .Height + ControlTop
52550   optFormatTXTControl.Left = .Left
52560   optFormatTXTControl.Width = .Width
52570   optFormatPSDControl.Top = .Top + .Height + ControlTop
52580   optFormatPSDControl.Left = .Left
52590   optFormatPSDControl.Width = .Width
52600   optFormatPCLControl.Top = .Top + .Height + ControlTop
52610   optFormatPCLControl.Left = .Left
52620   optFormatPCLControl.Width = .Width
52630   optFormatRAWControl.Top = .Top + .Height + ControlTop
52640   optFormatRAWControl.Left = .Left
52650   optFormatRAWControl.Width = .Width
52660   optFormatSVGControl.Top = .Top + .Height + ControlTop
52670   optFormatSVGControl.Left = .Left
52680   optFormatSVGControl.Width = .Width
52690 '  optFormatXCFControl.Top = .Top + .Height + ControlTop
52700 '  optFormatXCFControl.Left = .Left
52710 '  optFormatXCFControl.Width = .Width
52720   optGeneralControl.Top = .Top + .Height + ControlTop
52730   optGeneralControl.Left = .Left
52740   optGeneralControl.Width = .Width
52750   optGhostscriptControl.Top = .Top + .Height + ControlTop
52760   optGhostscriptControl.Left = .Left
52770   optGhostscriptControl.Width = .Width
52780   optLanguagesControl.Top = .Top + .Height + ControlTop
52790   optLanguagesControl.Left = .Left
52800   optLanguagesControl.Width = .Width
52810   optPrintControl.Top = .Top + .Height + ControlTop
52820   optPrintControl.Left = .Left
52830   optPrintControl.Width = .Width
52840   optSaveControl.Top = .Top + .Height + ControlTop
52850   optSaveControl.Left = .Left
52860   optSaveControl.Width = .Width
52870
52880   cmdCancel.Left = .Left
52890   cmdReset.Left = .Left + (.Width - cmdReset.Width) / 2
52900   cmdSave.Left = .Left + .Width - cmdSave.Width
52910  End With
52920
52930  If ShowOnlyOptions = True Then
52940   FormInTaskbar Me, True, True
52950   Caption = "PDFCreator - " & Caption
52960  End If
52970
52980  ShowAcceleratorsInForm Me, True
52990
53000  Screen.MousePointer = vbNormal
53010
53020  With Options
53030   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
53040  End With
53050  ieb.Refresh
53060  ieb_ItemClick "Program", "General"
53070  LoadReady = True
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
50050  optDirectoriesControl.Visible = False
50060  optDocumentControl.Visible = False
50070  optFontsControl.Visible = False
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
50570     Case "DIRECTORIES"
50580      Set picOptions = LoadResPicture(2104, vbResIcon)
50590      lblOptions.Caption = LanguageStrings.OptionsProgramDirectoriesDescription
50600      optDirectoriesControl.Visible = True
50610      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
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
50720     Case "FONTS"
50730      Set picOptions = LoadResPicture(2102, vbResIcon)
50740      lblOptions.Caption = LanguageStrings.OptionsProgramFontDescription
50750      optFontsControl.Visible = True
50760      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
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
50910     Case "PNG"
50920      Set picOptions = LoadResPicture(2112, vbResIcon)
50930      lblOptions.Caption = LanguageStrings.OptionsPNGDescription
50940      optFormatPNGControl.Visible = True
50950      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
50960     Case "JPEG"
50970      Set picOptions = LoadResPicture(2113, vbResIcon)
50980      lblOptions.Caption = LanguageStrings.OptionsJPEGDescription
50990      optFormatJPEGControl.Visible = True
51000      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51010     Case "BMP"
51020      Set picOptions = LoadResPicture(2114, vbResIcon)
51030      lblOptions.Caption = LanguageStrings.OptionsBMPDescription
51040      optFormatBMPControl.Visible = True
51050      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51060     Case "PCX"
51070      Set picOptions = LoadResPicture(2115, vbResIcon)
51080      lblOptions.Caption = LanguageStrings.OptionsPCXDescription
51090      optFormatPCXControl.Visible = True
51100      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51110     Case "TIFF"
51120      Set picOptions = LoadResPicture(2116, vbResIcon)
51130      lblOptions.Caption = LanguageStrings.OptionsTIFFDescription
51140      optFormatTIFFControl.Visible = True
51150      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51160     Case "PS"
51170      Set picOptions = LoadResPicture(2117, vbResIcon)
51180      lblOptions.Caption = LanguageStrings.OptionsPSDescription
51190      optFormatPSControl.Visible = True
51200      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51210     Case "EPS"
51220      Set picOptions = LoadResPicture(2118, vbResIcon)
51230      lblOptions.Caption = LanguageStrings.OptionsEPSDescription
51240      optFormatEPSControl.Visible = True
51250      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51260     Case "TXT"
51270      Set picOptions = LoadResPicture(2124, vbResIcon)
51280      lblOptions.Caption = LanguageStrings.OptionsTXTDescription
51290      optFormatTXTControl.Visible = True
51300      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51310     Case "PSD"
51320      Set picOptions = LoadResPicture(2125, vbResIcon)
51330      lblOptions.Caption = LanguageStrings.OptionsPSDDescription
51340      optFormatPSDControl.Visible = True
51350      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51360     Case "PCL"
51370      Set picOptions = LoadResPicture(2126, vbResIcon)
51380      lblOptions.Caption = LanguageStrings.OptionsPCLDescription
51390      optFormatPCLControl.Visible = True
51400      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51410     Case "RAW"
51420      Set picOptions = LoadResPicture(2127, vbResIcon)
51430      lblOptions.Caption = LanguageStrings.OptionsRAWDescription
51440      optFormatRAWControl.Visible = True
51450      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51460     Case "SVG"
51470      Set picOptions = LoadResPicture(2129, vbResIcon)
51480      lblOptions.Caption = LanguageStrings.OptionsSVGDescription
51490      optFormatSVGControl.Visible = True
51500      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51510 '    Case "XCF"
51520 '     Set picOptions = LoadResPicture(2128, vbResIcon)
51530 '     lblOptions.Caption = LanguageStrings.OptionsXCFDescription
51540 '     optFormatXCFControl.Visible = True
51550 '     dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51560    End Select
51570  End Select
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
50030  optDirectories.SetFrames OptionsDesign
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
