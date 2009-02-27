VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
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
            NumListImages   =   25
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
50380 '  ieb.SetItemText "Formats", "XCF", .OptionsXCFSymbol
50390   ieb.DisableUpdates False
50400
50410   lblOptions.Caption = .OptionsProgramLanguagesDescription
50420  End With
50430  optActions.SetLanguageStrings
50440  optAutosave.SetLanguageStrings
50450  optDirectories.SetLanguageStrings
50460  optDocument.SetLanguageStrings
50470  optFonts.SetLanguageStrings
50480  optFormatPNG.SetLanguageStrings
50490  optFormatJPEG.SetLanguageStrings
50500  optFormatBMP.SetLanguageStrings
50510  optFormatPCX.SetLanguageStrings
50520  optFormatTIFF.SetLanguageStrings
50530  optFormatPDF.SetLanguageStrings
50540  optFormatPS.SetLanguageStrings
50550  optFormatEPS.SetLanguageStrings
50560  optFormatTXT.SetLanguageStrings
50570  optFormatPSD.SetLanguageStrings
50580  optFormatPCL.SetLanguageStrings
50590  optFormatRAW.SetLanguageStrings
50600 ' optFormatXCF.SetLanguageStrings
50610  optGeneral.SetLanguageStrings
50620  optGhostscript.SetLanguageStrings
50630  optLanguages.SetLanguageStrings
50640  optPrint.SetLanguageStrings
50650  optSave.SetLanguageStrings
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
50230 '  optFormatXCF.SetOptions
50240   optGeneral.SetOptions
50250   optGhostscript.SetOptions
50260   optLanguages.SetOptions
50270   optPrint.SetOptions
50280   optSave.SetOptions
50290
50300   With Options
50310    SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50320    SetFont frmMain, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50330    ieb.Refresh
50340   End With
50350  End If
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
50260 ' optFormatXCF.GetOptions
50270  optGeneral.GetOptions
50280  optGhostscript.GetOptions
50290  optLanguages.GetOptions
50300  optPrint.GetOptions
50310  optSave.GetOptions
50320
50330 ' newLanguage = Options.Language
50340 ' Options.Language = newLanguage
50350 ' Options.StampFontname = tOpt.StampFontname
50360 ' Options.StampFontsize = tOpt.StampFontsize
50370  SaveOptions Options
50380
50390  SetHelpfile
50400
50410  If IsWin9xMe = False Then
50421   Select Case Options.ProcessPriority
         Case 0: 'Idle
50440     SetProcessPriority Idle
50450    Case 1: 'Normal
50460     SetProcessPriority Normal
50470    Case 2: 'High
50480     SetProcessPriority High
50490    Case 3: 'Realtime
50500     SetProcessPriority RealTime
50510   End Select
50520  End If
50530  If tRestart = True Then
50540   Restart = True
50550  End If
50560  Unload Me
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
50510   ieb.ExpandGroup "Formats", False
50520
50530   ieb.DisableUpdates False
50540
50550   Set picOptions = LoadResPicture(2101, vbResIcon)
50560
50570   Me.Caption = .DialogPrinterOptions
50580   cmdCancel.Caption = .OptionsCancel
50590   cmdReset.Caption = .OptionsReset
50600   cmdSave.Caption = .OptionsSave
50610  End With
50620
50630  SetFrame dmFraDescription
50640
50650  ' Add ActionsControl
50660  Set optActionsControl = Controls.Add("PDFCreator.ctlOptActions", "ctlOptActions")
50670  optActionsControl.Width = dmFraDescription.Width
50680  Set optActions = optActionsControl.object
50690  optActions.SetLanguageStrings
50700  optActions.SetOptions
50710  ' Add AutosaveControl
50720  Set optAutosaveControl = Controls.Add("PDFCreator.ctlOptAutosave", "ctlOptAutosave")
50730  optAutosaveControl.Width = dmFraDescription.Width
50740  Set optAutosave = optAutosaveControl.object
50750  optAutosave.SetLanguageStrings
50760  optAutosave.SetOptions
50770  ' Add DirectoriesControl
50780  Set optDirectoriesControl = Controls.Add("PDFCreator.ctlOptDirectories", "ctlOptDirectories")
50790  optDirectoriesControl.Width = dmFraDescription.Width
50800  Set optDirectories = optDirectoriesControl.object
50810  optDirectories.SetLanguageStrings
50820  optDirectories.SetOptions
50830  ' Add DocumentControl
50840  Set optDocumentControl = Controls.Add("PDFCreator.ctlOptDocument", "ctlOptDocument")
50850  optDocumentControl.Width = dmFraDescription.Width
50860  Set optDocument = optDocumentControl.object
50870  optDocument.SetLanguageStrings
50880  optDocument.SetOptions
50890  ' Add FontsControl
50900  Set optFontsControl = Controls.Add("PDFCreator.ctlOptFonts", "ctlOptFonts")
50910  optFontsControl.Width = dmFraDescription.Width
50920  Set optFonts = optFontsControl.object
50930  optFonts.SetLanguageStrings
50940  optFonts.SetOptions
50950  ' Add FormatPNGControl
50960  Set optFormatPNGControl = Controls.Add("PDFCreator.ctlOptFormatPNG", "ctlOptFormatPNG")
50970  optFormatPNGControl.Width = dmFraDescription.Width
50980  Set optFormatPNG = optFormatPNGControl.object
50990  optFormatPNG.SetLanguageStrings
51000  optFormatPNG.SetOptions
51010  ' Add FormatJPEQControl
51020  Set optFormatJPEGControl = Controls.Add("PDFCreator.ctlOptFormatJPEG", "ctlOptFormatJPEG")
51030  optFormatJPEGControl.Width = dmFraDescription.Width
51040  Set optFormatJPEG = optFormatJPEGControl.object
51050  optFormatJPEG.SetLanguageStrings
51060  optFormatJPEG.SetOptions
51070  ' Add FormatBMPControl
51080  Set optFormatBMPControl = Controls.Add("PDFCreator.ctlOptFormatBMP", "ctlOptFormatBMP")
51090  optFormatBMPControl.Width = dmFraDescription.Width
51100  Set optFormatBMP = optFormatBMPControl.object
51110  optFormatBMP.SetLanguageStrings
51120  optFormatBMP.SetOptions
51130  ' Add FormatPCXControl
51140  Set optFormatPCXControl = Controls.Add("PDFCreator.ctlOptFormatPCX", "ctlOptFormatPCX")
51150  optFormatPCXControl.Width = dmFraDescription.Width
51160  Set optFormatPCX = optFormatPCXControl.object
51170  optFormatPCX.SetLanguageStrings
51180  optFormatPCX.SetOptions
51190  ' Add FormatTIFFControl
51200  Set optFormatTIFFControl = Controls.Add("PDFCreator.ctlOptFormatTIFF", "ctlOptFormatTIFF")
51210  optFormatTIFFControl.Width = dmFraDescription.Width
51220  Set optFormatTIFF = optFormatTIFFControl.object
51230  optFormatTIFF.SetLanguageStrings
51240  optFormatTIFF.SetOptions
51250  ' Add FormatPDFControl
51260  Set optFormatPDFControl = Controls.Add("PDFCreator.ctlOptFormatPDF", "ctlOptFormatPDF")
51270  optFormatPDFControl.Width = dmFraDescription.Width
51280  Set optFormatPDF = optFormatPDFControl.object
51290  optFormatPDF.SetLanguageStrings
51300  optFormatPDF.SetOptions
51310  ' Add FormatPS
51320  Set optFormatPSControl = Controls.Add("PDFCreator.ctlOptFormatPS", "ctlOptFormatPS")
51330  optFormatPSControl.Width = dmFraDescription.Width
51340  Set optFormatPS = optFormatPSControl.object
51350  optFormatPS.SetLanguageStrings
51360  optFormatPS.SetOptions
51370  ' Add FormatEPSControl
51380  Set optFormatEPSControl = Controls.Add("PDFCreator.ctlOptFormatEPS", "ctlOptFormatEPS")
51390  optFormatEPSControl.Width = dmFraDescription.Width
51400  Set optFormatEPS = optFormatEPSControl.object
51410  optFormatEPS.SetLanguageStrings
51420  optFormatEPS.SetOptions
51430  ' Add FormatTXTControl
51440  Set optFormatTXTControl = Controls.Add("PDFCreator.ctlOptFormatTXT", "ctlOptFormatTXT")
51450  optFormatTXTControl.Width = dmFraDescription.Width
51460  Set optFormatTXT = optFormatTXTControl.object
51470  optFormatTXT.SetLanguageStrings
51480  optFormatTXT.SetOptions
51490  ' Add FormatPSDControl
51500  Set optFormatPSDControl = Controls.Add("PDFCreator.ctlOptFormatPSD", "ctlOptFormatPSD")
51510  optFormatPSDControl.Width = dmFraDescription.Width
51520  Set optFormatPSD = optFormatPSDControl.object
51530  optFormatPSD.SetLanguageStrings
51540  optFormatPSD.SetOptions
51550  ' Add FormatPCLControl
51560  Set optFormatPCLControl = Controls.Add("PDFCreator.ctlOptFormatPCL", "ctlOptFormatPCL")
51570  optFormatPCLControl.Width = dmFraDescription.Width
51580  Set optFormatPCL = optFormatPCLControl.object
51590  optFormatPCL.SetLanguageStrings
51600  optFormatPCL.SetOptions
51610  ' Add FormatRAWControl
51620  Set optFormatRAWControl = Controls.Add("PDFCreator.ctlOptFormatRAW", "ctlOptFormatRAW")
51630  optFormatRAWControl.Width = dmFraDescription.Width
51640  Set optFormatRAW = optFormatRAWControl.object
51650  optFormatRAW.SetLanguageStrings
51660  optFormatRAW.SetOptions
51670 ' ' Add FormatXCFControl - Doesn't work
51680 ' Set optFormatXCFControl = Controls.Add("PDFCreator.ctlOptFormatXCF", "ctlOptFormatXCF")
51690 ' optFormatXCFControl.Width = dmFraDescription.Width
51700 ' Set optFormatXCF = optFormatXCFControl.object
51710 ' optFormatXCF.SetLanguageStrings
51720 ' optFormatXCF.SetOptions
51730  ' Add GhostscriptControl
51740  Set optGhostscriptControl = Controls.Add("PDFCreator.ctlOptGhostscript", "ctlOptGhostscript")
51750  optGhostscriptControl.Width = dmFraDescription.Width
51760  Set optGhostscript = optGhostscriptControl.object
51770  optGhostscript.SetLanguageStrings
51780  optGhostscript.SetOptions
51790  ' Add LanguagesControl
51800  Set optLanguagesControl = Controls.Add("PDFCreator.ctlOptLanguages", "ctlOptLanguages")
51810  optLanguagesControl.Width = dmFraDescription.Width
51820  Set optLanguages = optLanguagesControl.object
51830  optLanguages.SetLanguageStrings
51840  optLanguages.SetOptions
51850  ' Add PrintControl
51860  Set optPrintControl = Controls.Add("PDFCreator.ctlOptPrint", "ctlOptPrint")
51870  optPrintControl.Width = dmFraDescription.Width
51880  Set optPrint = optPrintControl.object
51890  optPrint.SetLanguageStrings
51900  optPrint.SetOptions
51910  ' Add SaveControl
51920  Set optSaveControl = Controls.Add("PDFCreator.ctlOptSave", "ctlOptSave")
51930  optSaveControl.Width = dmFraDescription.Width
51940  Set optSave = optSaveControl.object
51950  optSave.SetLanguageStrings
51960  optSave.SetOptions
51970  ' Add GeneralControl
51980  Set optGeneralControl = Controls.Add("PDFCreator.ctlOptGeneral", "ctlOptGeneral")
51990  optGeneralControl.Width = dmFraDescription.Width
52000  Set optGeneral = optGeneralControl.object
52010  optGeneral.SetLanguageStrings
52020  optGeneral.SetOptions
52030
52040  With dmFraDescription
52050   .Caption = LanguageStrings.OptionsTreeProgram
52060   .Visible = True
52070
52080   optActionsControl.Top = .Top + .Height + ControlTop
52090   optActionsControl.Left = .Left
52100   optActionsControl.Width = .Width
52110   optAutosaveControl.Top = .Top + .Height + ControlTop
52120   optAutosaveControl.Left = .Left
52130   optAutosaveControl.Width = .Width
52140   optDirectoriesControl.Top = .Top + .Height + ControlTop
52150   optDirectoriesControl.Left = .Left
52160   optDirectoriesControl.Width = .Width
52170   optDocumentControl.Top = .Top + .Height + ControlTop
52180   optDocumentControl.Left = .Left
52190   optDocumentControl.Width = .Width
52200   optFontsControl.Top = .Top + .Height + ControlTop
52210   optFontsControl.Left = .Left
52220   optFontsControl.Width = .Width
52230   optFormatPNGControl.Top = .Top + .Height + ControlTop
52240   optFormatPNGControl.Left = .Left
52250   optFormatPNGControl.Width = .Width
52260   optFormatJPEGControl.Top = .Top + .Height + ControlTop
52270   optFormatJPEGControl.Left = .Left
52280   optFormatJPEGControl.Width = .Width
52290   optFormatBMPControl.Top = .Top + .Height + ControlTop
52300   optFormatBMPControl.Left = .Left
52310   optFormatBMPControl.Width = .Width
52320   optFormatPCXControl.Top = .Top + .Height + ControlTop
52330   optFormatPCXControl.Left = .Left
52340   optFormatPCXControl.Width = .Width
52350   optFormatTIFFControl.Top = .Top + .Height + ControlTop
52360   optFormatTIFFControl.Left = .Left
52370   optFormatTIFFControl.Width = .Width
52380   optFormatPDFControl.Top = .Top + .Height + ControlTop
52390   optFormatPDFControl.Left = .Left
52400   optFormatPDFControl.Width = .Width
52410   optFormatPSControl.Top = .Top + .Height + ControlTop
52420   optFormatPSControl.Left = .Left
52430   optFormatPSControl.Width = .Width
52440   optFormatEPSControl.Top = .Top + .Height + ControlTop
52450   optFormatEPSControl.Left = .Left
52460   optFormatEPSControl.Width = .Width
52470   optFormatTXTControl.Top = .Top + .Height + ControlTop
52480   optFormatTXTControl.Left = .Left
52490   optFormatTXTControl.Width = .Width
52500   optFormatPSDControl.Top = .Top + .Height + ControlTop
52510   optFormatPSDControl.Left = .Left
52520   optFormatPSDControl.Width = .Width
52530   optFormatPCLControl.Top = .Top + .Height + ControlTop
52540   optFormatPCLControl.Left = .Left
52550   optFormatPCLControl.Width = .Width
52560   optFormatRAWControl.Top = .Top + .Height + ControlTop
52570   optFormatRAWControl.Left = .Left
52580   optFormatRAWControl.Width = .Width
52590 '  optFormatXCFControl.Top = .Top + .Height + ControlTop
52600 '  optFormatXCFControl.Left = .Left
52610 '  optFormatXCFControl.Width = .Width
52620   optGeneralControl.Top = .Top + .Height + ControlTop
52630   optGeneralControl.Left = .Left
52640   optGeneralControl.Width = .Width
52650   optGhostscriptControl.Top = .Top + .Height + ControlTop
52660   optGhostscriptControl.Left = .Left
52670   optGhostscriptControl.Width = .Width
52680   optLanguagesControl.Top = .Top + .Height + ControlTop
52690   optLanguagesControl.Left = .Left
52700   optLanguagesControl.Width = .Width
52710   optPrintControl.Top = .Top + .Height + ControlTop
52720   optPrintControl.Left = .Left
52730   optPrintControl.Width = .Width
52740   optSaveControl.Top = .Top + .Height + ControlTop
52750   optSaveControl.Left = .Left
52760   optSaveControl.Width = .Width
52770
52780   cmdCancel.Left = .Left
52790   cmdReset.Left = .Left + (.Width - cmdReset.Width) / 2
52800   cmdSave.Left = .Left + .Width - cmdSave.Width
52810  End With
52820
52830  If ShowOnlyOptions = True Then
52840   FormInTaskbar Me, True, True
52850   Caption = "PDFCreator - " & Caption
52860  End If
52870
52880  ShowAcceleratorsInForm Me, True
52890
52900  Screen.MousePointer = vbNormal
52910
52920  With Options
52930   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
52940  End With
52950  ieb.Refresh
52960  ieb_ItemClick "Program", "General"
52970  LoadReady = True
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
50200 ' optFormatXCFControl.Visible = False
50210  optGeneralControl.Visible = False
50220  optGhostscriptControl.Visible = False
50230  optLanguagesControl.Visible = False
50240  optPrintControl.Visible = False
50250  optSaveControl.Visible = False
50260
50271  Select Case UCase$(sGroup)
        Case "PROGRAM"
50290    LastProgramGroup = sItemKey
50301    Select Case UCase$(sItemKey)
          Case "GENERAL"
50320      Set picOptions = LoadResPicture(2101, vbResIcon)
50330      lblOptions.Caption = LanguageStrings.OptionsProgramGeneralDescription
50340      optGeneralControl.Visible = True
50350      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50360     Case "GHOSTSCRIPT"
50370      Set picOptions = LoadResPicture(2119, vbResIcon)
50380      lblOptions.Caption = LanguageStrings.OptionsProgramGhostscriptDescription
50390      optGhostscriptControl.Visible = True
50400      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50410     Case "DOCUMENT"
50420      Set picOptions = LoadResPicture(2105, vbResIcon)
50430      lblOptions.Caption = LanguageStrings.OptionsProgramDocumentDescription
50440      optDocumentControl.Visible = True
50450      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50460     Case "SAVE"
50470      Set picOptions = LoadResPicture(2106, vbResIcon)
50480      lblOptions.Caption = LanguageStrings.OptionsProgramSaveDescription
50490      optSaveControl.Visible = True
50500      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50510     Case "AUTOSAVE"
50520      Set picOptions = LoadResPicture(2103, vbResIcon)
50530      lblOptions.Caption = LanguageStrings.OptionsProgramAutosaveDescription
50540      optAutosaveControl.Visible = True
50550      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50560     Case "DIRECTORIES"
50570      Set picOptions = LoadResPicture(2104, vbResIcon)
50580      lblOptions.Caption = LanguageStrings.OptionsProgramDirectoriesDescription
50590      optDirectoriesControl.Visible = True
50600      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50610     Case "ACTIONS"
50620      Set picOptions = LoadResPicture(2121, vbResIcon)
50630      lblOptions.Caption = LanguageStrings.OptionsProgramActionsDescription
50640      optActionsControl.Visible = True
50650      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50660     Case "PRINT"
50670      Set picOptions = LoadResPicture(2122, vbResIcon)
50680      lblOptions.Caption = LanguageStrings.OptionsProgramPrintDescription
50690      optPrintControl.Visible = True
50700      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50710     Case "FONTS"
50720      Set picOptions = LoadResPicture(2102, vbResIcon)
50730      lblOptions.Caption = LanguageStrings.OptionsProgramFontDescription
50740      optFontsControl.Visible = True
50750      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50760     Case "LANGUAGE"
50770      Set picOptions = LoadResPicture(2123, vbResIcon)
50780      lblOptions.Caption = LanguageStrings.OptionsProgramLanguagesDescription
50790      optLanguagesControl.Visible = True
50800      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50810    End Select
50820   Case "FORMATS"
50830    LastFormatGroup = sItemKey
50841    Select Case UCase$(sItemKey)
          Case "PDF"
50860      Set picOptions = LoadResPicture(2111, vbResIcon)
50870      lblOptions.Caption = LanguageStrings.OptionsPDFDescription
50880      optFormatPDFControl.Visible = True
50890      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
50900     Case "PNG"
50910      Set picOptions = LoadResPicture(2112, vbResIcon)
50920      lblOptions.Caption = LanguageStrings.OptionsPNGDescription
50930      optFormatPNGControl.Visible = True
50940      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
50950     Case "JPEG"
50960      Set picOptions = LoadResPicture(2113, vbResIcon)
50970      lblOptions.Caption = LanguageStrings.OptionsJPEGDescription
50980      optFormatJPEGControl.Visible = True
50990      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51000     Case "BMP"
51010      Set picOptions = LoadResPicture(2114, vbResIcon)
51020      lblOptions.Caption = LanguageStrings.OptionsBMPDescription
51030      optFormatBMPControl.Visible = True
51040      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51050     Case "PCX"
51060      Set picOptions = LoadResPicture(2115, vbResIcon)
51070      lblOptions.Caption = LanguageStrings.OptionsPCXDescription
51080      optFormatPCXControl.Visible = True
51090      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51100     Case "TIFF"
51110      Set picOptions = LoadResPicture(2116, vbResIcon)
51120      lblOptions.Caption = LanguageStrings.OptionsTIFFDescription
51130      optFormatTIFFControl.Visible = True
51140      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51150     Case "PS"
51160      Set picOptions = LoadResPicture(2117, vbResIcon)
51170      lblOptions.Caption = LanguageStrings.OptionsPSDescription
51180      optFormatPSControl.Visible = True
51190      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51200     Case "EPS"
51210      Set picOptions = LoadResPicture(2118, vbResIcon)
51220      lblOptions.Caption = LanguageStrings.OptionsEPSDescription
51230      optFormatEPSControl.Visible = True
51240      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51250     Case "TXT"
51260      Set picOptions = LoadResPicture(2124, vbResIcon)
51270      lblOptions.Caption = LanguageStrings.OptionsTXTDescription
51280      optFormatTXTControl.Visible = True
51290      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51300     Case "PSD"
51310      Set picOptions = LoadResPicture(2125, vbResIcon)
51320      lblOptions.Caption = LanguageStrings.OptionsPSDDescription
51330      optFormatPSDControl.Visible = True
51340      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51350     Case "PCL"
51360      Set picOptions = LoadResPicture(2126, vbResIcon)
51370      lblOptions.Caption = LanguageStrings.OptionsPCLDescription
51380      optFormatPCLControl.Visible = True
51390      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51400     Case "RAW"
51410      Set picOptions = LoadResPicture(2127, vbResIcon)
51420      lblOptions.Caption = LanguageStrings.OptionsRAWDescription
51430      optFormatRAWControl.Visible = True
51440      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51450 '    Case "XCF"
51460 '     Set picOptions = LoadResPicture(2128, vbResIcon)
51470 '     lblOptions.Caption = LanguageStrings.OptionsXCFDescription
51480 '     optFormatXCFControl.Visible = True
51490 '     dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51500    End Select
51510  End Select
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
50160  optFormatTIFF.SetFrames OptionsDesign
50170  optFormatTXT.SetFrames OptionsDesign
50180  optGeneral.SetFrames OptionsDesign
50190  optGhostscript.SetFrames OptionsDesign
50200  optLanguages.SetFrames OptionsDesign
50210  optPrint.SetFrames OptionsDesign
50220  optSave.SetFrames OptionsDesign
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
