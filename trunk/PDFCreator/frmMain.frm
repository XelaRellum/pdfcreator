VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "PDFCreator"
   ClientHeight    =   2970
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9510
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows-Standard
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   960
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   2160
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Unten ausrichten
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   2700
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsv 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2778
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.Menu mnPrinterMain 
      Caption         =   "Printer"
      Begin VB.Menu mnPrinter 
         Caption         =   "Printer stop "
         Index           =   0
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Options"
         Index           =   2
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Logging"
         Index           =   4
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Logfile"
         Index           =   5
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Close"
         Index           =   7
      End
   End
   Begin VB.Menu mnDocumentMain 
      Caption         =   "Document"
      Begin VB.Menu mnDocument 
         Caption         =   "Print"
         Index           =   0
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Add"
         Index           =   2
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Delete"
         Index           =   3
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Top"
         Index           =   5
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Up"
         Index           =   6
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Down"
         Index           =   7
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Bottom"
         Index           =   8
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Combine"
         Index           =   10
      End
   End
   Begin VB.Menu mnViewMain 
      Caption         =   "View"
      Begin VB.Menu mnView 
         Caption         =   "Statusbar"
         Checked         =   -1  'True
         Index           =   0
      End
   End
   Begin VB.Menu mnLanguageMain 
      Caption         =   "Language"
      Begin VB.Menu mnLanguage 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnHelpMain 
      Caption         =   "?"
      Begin VB.Menu mnHelp 
         Caption         =   "Info"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'"On Error Resume Next" Functions -> Don't add the ErrorHandler
' modGeneral: makePath
' modPDFCreator: ClearCache
' modLvwListItem: LvwInsertListItemStore
' frmPrinting: Create_eDoc
' frmSave: SetNode

Const TimerIntervall = 500

Private LanguagePath As String, Languagefile As String, mutex As clsMutex, _
 Printjobs As Collection

Private Sub Form_Load()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, stdio As clsStdIO, cinStr As String, Tempfile As String, _
  InFile As String, OutFile As String, Ext As String, IFIsPS As Boolean
50030 '##############################################
50040 'Performance Tools
50050  Dim LastStop As Currency
50060  LastStop = ExactTimer_Value()
50070 '##############################################
50080
50090  If CheckPath(GetMyAppData) = True Then
50100    If Len(Dir(GetMyAppData & "PDFcreator", vbDirectory)) = 0 Then
50110     MakePath GetMyAppData & "PDFcreator"
50120    End If
50130    PDFCreatorINIFile = GetMyAppData & "PDFcreator\PDFCreator.ini"
50140   Else
50150    PDFCreatorINIFile = App.Path & "\PDFCreator.ini"
50160  End If
50170
50180  Options = ReadOptions
50190
50200  IfLoggingWriteLogfile "PDFCreator Program Start"
50210
50220  ' The program has commandswitches
50230  ' -IPTRUE : Install Printer
50240  ' -IPFALSE: UnInstall Printer
50250  ' -NSTRUE: No Start
50260  ' -ULTRUE: Unload all PDFCreator programs
50270  ' -PPDFCREATORPRINTER: The printer call the program
50280  ' -IF: Inputfile
50290  ' -OF: Outputfile
50300  ' -CLEARCACHE: Clear the temp. cache
50310
50320  ' Check Installprinter
50330
50340  Select Case UCase$(CommandSwitch("IP", True))
  Case "TRUE":
50360    Monitorname = "PDFCreator": Portname = "PDFCreator:": Drivername = "PDFCreator": PrinterName = "PDFCreator"
50370    InstallCompletePrinter
50380   Case "FALSE":
50390    Monitorname = "PDFCreator": Portname = "PDFCreator:": Drivername = "PDFCreator": PrinterName = "PDFCreator"
50400    UnInstallCompletePrinter
50410  End Select
50420
50430  ' Initialize unload running program
50440  If UCase$(CommandSwitch("UL", True)) = "TRUE" Then
50450   fn = FreeFile
50460   Open App.Path & "\Unload.tmp" For Output As #fn
50470   Close #fn
50480  End If
50490
50500  If UCase$(CommandSwitch("CLEAR", True)) = "CACHE" Then
50510   ClearCache
50520  End If
50530
50540  ' NS: If NS=True Then end the program here
50550  ' It is necessary for uninstall.
50560  If UCase$(CommandSwitch("NS", True)) = "TRUE" Then
50570   End
50580  End If
50590
50600  CreatePDFCreatorTempfolder
50610
50620  If IsWin9xMe = False Then
50630   Select Case Options.ProcessPriority
   Case 0: 'Idle
50650     SetProcessPriority Idle
50660    Case 1: 'Normal
50670     SetProcessPriority Normal
50680    Case 2: 'High
50690     SetProcessPriority High
50700    Case 3: 'Realtime
50710     SetProcessPriority RealTime
50720   End Select
50730  End If
50740
50750  LanguagePath = App.Path & "\Languages\"
50760  ReadAllLanguages LanguagePath
50770  Languagefile = LanguagePath & Options.Language & ".ini"
50780  LoadLanguage Languagefile
50790
50800  GsDllLoaded = LoadDLL(Options.DirectoryGhostscriptBinaries & "\gsdll32.dll")
50810
50820  If GsDllLoaded = 0 Then
50830   MsgBox LanguageStrings.MessagesMsg08
50840  End If
50850
50860  IFIsPS = False
50870  InFile = UCase$(CommandSwitch("IF", True))
50880  If Len(InFile) > 0 Then
50890   If Dir(InFile) <> "" Then
50900     If Len(UCase$(CommandSwitch("OF", True))) > 0 Then
50910       If CheckIfPSFile(InFile) = True Then
50920        If GsDllLoaded = 0 Then
50930         End
50940        End If
50950        OutFile = CommandSwitch("OF", True)
50960        SplitPath OutFile, , , , , Ext
50970        Select Case UCase$(Ext)
        Case "PDF"
50990          CallGScript InFile, OutFile, Options, PDFWriter
51000         Case "PNG"
51010          CallGScript PDFSpoolfile, OutFile, Options, PNGWriter
51020         Case "JPG"
51030          CallGScript PDFSpoolfile, OutFile, Options, JPEGWriter
51040         Case "BMP"
51050          CallGScript PDFSpoolfile, OutFile, Options, BMPWriter
51060         Case "PCX"
51070          CallGScript PDFSpoolfile, OutFile, Options, PCXWriter
51080         Case "TIF"
51090          CallGScript PDFSpoolfile, OutFile, Options, TIFFWriter
51100         Case "PS"
51110          CallGScript PDFSpoolfile, OutFile, Options, PSWriter
51120         Case "EPS"
51130          CallGScript PDFSpoolfile, OutFile, Options, EPSWriter
51140        End Select
51150       End If
51160       End
51170      Else
51180       If CheckIfPSFile(CommandSwitch("IF", True)) Then
51190         Tempfile = GetTempFile(GetPDFCreatorTempfolder, "~PD")
51200         FileCopy CommandSwitch("IF", True), Tempfile
51210         IFIsPS = True
51220        Else
51230         MsgBox LanguageStrings.MessagesMsg06
51240       End If
51250       DoEvents
51260     End If
51270    Else
51280     MsgBox LanguageStrings.MessagesMsg14
51290   End If
51300  End If
51310
51320  ' Printer has started the program
51330  If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
51340   CheckAutosaveAndPrint
51350  End If
51360
51370  ' Create a mutex; if mutex exists then exit
51380  Set mutex = New clsMutex
51390  If mutex.CheckMutex(PDFCreator_GUID) = False Then
51400    mutex.CreateMutex PDFCreator_GUID
51410   Else
51420    End
51430  End If
51440
51450  InitProgram
51460
51470  If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Or _
  (Len(CommandSwitch("IF", True)) > 0 And IFIsPS = True) Then
51490   If lsv.ListItems.Count <= 1 Then
51500    Me.Visible = False
51510   End If
51520  End If
51530
51540 '##############################################
51550 'MsgBox "Programmstart: " & ExactTimer_Value() - LastStop & " Sekunden"
51560 'LastStop = ExactTimer_Value()
51570 '##############################################
51580  'IfLoggingWriteLogfile "PDFCreator started in " & ExactTimer_Value() - LastStop & " seconds"
51590
51600  ' Only for the first time set Interval to 10 ms
51610  Timer1.Interval = 10
51620  Timer1.Enabled = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Me.WindowState = vbMinimized Then
50020   Exit Sub
50030  End If
50040  If Me.Height < 3000 Then
50050   Me.Height = 3000
50060   Exit Sub
50070  End If
50080  If Me.Width < 3000 Then
50090   Me.Width = 3000
50100   Exit Sub
50110  End If
50120  With lsv
50130   .Top = 0: .Left = 0
50140   .Width = Me.Width - 125
50150   .Height = Me.ScaleHeight - Abs(stb.Visible) * stb.Height
50160  End With
50170  stb.Panels("Status").Width = Me.Width - 150 - stb.Panels("Percent").Width
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Form_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Unload(Cancel As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  TerminateProgram
50020  End
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Form_Unload")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub InitProgram()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim FileName As String, Tempfile As String
50020
50030  Printing = False
50040  FileName = CommandSwitch("F", True)
50050
50060  If Len(Dir(FileName)) > 0 And Len(Trim$(FileName)) > 0 Then
50070   If FileLen(FileName) > 0 Then
50080    Tempfile = GetTempFile(GetPDFCreatorTempfolder, "~PD")
50090    FileCopy FileName, Tempfile
50100   End If
50110  End If
50120
50130  Set Printjobs = New Collection
50140
50150  stb.Panels.Clear
50160  stb.Panels.Add , "Status", ""
50170  stb.Panels.Add , "Percent", ""
50180  stb.Panels("Percent").Width = 1000
50190
50200  With lsv
50210   .View = lvwReport
50220   .FullRowSelect = True
50230   .HideSelection = False
50240   .ColumnHeaders.Clear
50250   .ColumnHeaders.Add , "Documenttitle", "Documenttitle", 2000
50260   .ColumnHeaders.Add , "Status", "Status", 1500
50270   .ColumnHeaders.Add , "Date", "Created on", 1700
50280   .ColumnHeaders.Add , "Size", "Size", 1500, lvwColumnRight
50290   .ColumnHeaders.Add , "Filename", "Filename", lsv.Width - 3500
50300  End With
50310
50320
50330  With Options
50340   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50350  End With
50360
50370  SetLanguageMenu
50380  If Options.Logging = 1 Then
50390    mnPrinter(4).Checked = True
50400   Else
50410    mnPrinter(4).Checked = False
50420  End If
50430
50440  CheckPrintJobs
50450  DoEvents
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "InitProgram")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub TerminateProgram()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Timer1.Enabled = False
50020  Set Printjobs = Nothing
50030  mutex.CloseMutex
50040  Set mutex = Nothing
50050  IfLoggingWriteLogfile "PDFCreator Program End"
50060  UnLoadDLL GsDllLoaded
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "TerminateProgram")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetAllLanguagesFiles(LanguagePath As String) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Languagefile As String
50020  Set GetAllLanguagesFiles = New Collection
50030  Languagefile = Dir(LanguagePath & "*.ini")
50040  Do While Len(Languagefile) > 0
50050    GetAllLanguagesFiles.Add LanguagePath & Languagefile
50060    Languagefile = Dir()
50070   DoEvents
50080  Loop
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "GetAllLanguagesFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub ReadAllLanguages(LanguagePath As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Languagename As String, ini As clsINI, LangFiles As Collection, i As Long
50020  mnLanguage(0).Caption = "No languages available."
50030
50040  Set LangFiles = GetAllLanguagesFiles(LanguagePath)
50050  Set ini = New clsINI
50060  For i = 1 To LangFiles.Count
50070   ini.FileName = LangFiles.item(i)
50080   ini.Section = "Common"
50090   Languagename = ini.GetKeyFromSection("Languagename")
50100   If Len(Languagename) = 0 Then
50110    Languagename = "No name available."
50120   End If
50130   Load mnLanguage(mnLanguage.Count)
50140   mnLanguage(mnLanguage.Count - 1).Caption = Languagename
50150   mnLanguage(mnLanguage.Count - 1).Tag = LangFiles.item(i)
50160   DoEvents
50170  Next i
50180
50190  If mnLanguage.Count > 1 Then
50200   mnLanguage(0).Caption = "No languages available."
50210   mnLanguage(0).Visible = False
50220  End If
50230  Set ini = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "ReadAllLanguages")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetLanguageMenu()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, Version As String, reg As clsRegistry
50020
50030  For i = mnLanguage.LBound To mnLanguage.UBound
50040   If UCase$(Languagefile) = UCase$(mnLanguage.item(i).Tag) Then
50050     mnLanguage.item(i).Checked = True
50060    Else
50070     mnLanguage.item(i).Checked = False
50080   End If
50090  Next i
50100
50110  With LanguageStrings
50120   Set reg = New clsRegistry
50130   With reg
50140    .hkey = HKEY_LOCAL_MACHINE
50150    .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50160    Version = .GetRegistryValue("ApplicationVersion")
50170   End With
50180   Set reg = Nothing
50190   Caption = App.Title & " " & GetProgramRelease & " " & .CommonTitle
50200
50210   mnPrinterMain.Caption = .DialogPrinter
50220   mnPrinter(0).Caption = .DialogPrinterPrinterStop
50230   mnPrinter(2).Caption = .DialogPrinterOptions
50240   mnPrinter(4).Caption = .DialogPrinterLogging
50250   mnPrinter(5).Caption = .DialogPrinterLogfile
50260   mnPrinter(7).Caption = .DialogPrinterClose
50270
50280   mnDocumentMain.Caption = .DialogDocument
50290   mnDocument(0).Caption = .DialogDocumentPrint
50300   mnDocument(2).Caption = .DialogDocumentAdd
50310   mnDocument(3).Caption = .DialogDocumentDelete
50320   mnDocument(5).Caption = .DialogDocumentTop
50330   mnDocument(6).Caption = .DialogDocumentUp
50340   mnDocument(7).Caption = .DialogDocumentDown
50350   mnDocument(8).Caption = .DialogDocumentBottom
50360   mnDocument(10).Caption = .DialogDocumentCombine
50370
50380   mnViewMain.Caption = .DialogView
50390   mnView(0).Caption = .DialogViewStatusbar
50400
50410   mnLanguageMain.Caption = .DialogLanguage
50420
50430   mnHelp(0).Caption = .DialogInfo
50440
50450   lsv.ColumnHeaders("Date").Text = .ListDate
50460   lsv.ColumnHeaders("Documenttitle").Text = .ListDocumenttitle
50470   lsv.ColumnHeaders("Filename").Text = .ListFilename
50480   lsv.ColumnHeaders("Size").Text = .ListSize
50490   lsv.ColumnHeaders("Status").Text = .ListStatus
50500  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SetLanguageMenu")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsv_KeyUp(KeyCode As Integer, Shift As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  If KeyCode = 46 Then
50030   For i = 1 To lsv.ListItems.Count
50040    If lsv.ListItems(i).Selected = True Then
50050     Kill lsv.ListItems(i).SubItems(4)
50060    End If
50070    DoEvents
50080   Next i
50090   LvwRemoveSelectedItems lsv, True
50100  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "lsv_KeyUp")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 If Button = 2 Then
50020   SetDocumentMenu
50030   PopupMenu mnDocumentMain, , x, y
50040 End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "lsv_MouseUp")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsv_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tFilename As String, i As Long, aLen As Double, tLen As Double
50020
50030  If data.GetFormat(vbCFFiles) Then
50040   If data.Files.Count = 1 Then
50050     If CheckIfPSFile(data.Files.item(1)) Then
50060      tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
50070      FileCopy data.Files.item(1), tFilename
50080     End If
50090     DoEvents
50100    Else
50110     aLen = 0
50120     For i = 1 To data.Files.Count
50130      aLen = aLen + FileLen(data.Files.item(i))
50140     Next i
50150     For i = 1 To data.Files.Count
50160      If CheckIfPSFile(data.Files.item(i)) Then
50170       tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
50180       DoEvents
50190       FileCopy data.Files.item(i), tFilename
50200      End If
50210      tLen = tLen + FileLen(data.Files.item(i))
50220      stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
50230      DoEvents
50240     Next i
50250   End If
50260  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "lsv_OLEDragDrop")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnDocument_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tFilename As String, cFiles As Collection, sFiles() As String, _
  i As Long, j As Long, aLen As Double, tLen As Double, aw As Long
50030  Timer1.Enabled = False
50040  Screen.MousePointer = vbHourglass
50050  DoEvents
50060  Select Case Index
  Case 0:
50080    For j = 1 To LvwGetCountSelectedItems(lsv, True)
50090     DoEvents
50100     For i = lsv.ListItems.Count To 1 Step -1
50110      If lsv.ListItems(i).Selected = True Then
50120       lsv.ListItems(i).SubItems(1) = LanguageStrings.ListPrinting
50130       LvwListItemToTop lsv, i, True
50140       Exit For
50150      End If
50160     Next i
50170    Next j
50180    SetPrinterStop False
50190    mnPrinter(0).Checked = False
50200   Case 2: ' Add
50210    DoEvents
50220    With cdlg
50230     .InitDir = GetSpecialFolder(ssfPERSONAL)
50240     .Flags = cdlOFNFileMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNLongNames
50250     .DialogTitle = LanguageStrings.ListAddPostscriptFile
50260     .Filter = LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps|"
50270     .FileName = vbNullString
50280     .ShowOpen
50290     If Len(.FileName) > 0 Then
50300      sFiles = Split(.FileName, Chr$(0))
50310      If UBound(sFiles) = 0 Then
50320        If CheckIfPSFile(sFiles(0)) Then
50330          tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
50340          Kill tFilename
50350          FileCopy sFiles(0), tFilename
50360         Else
50370          MsgBox LanguageStrings.MessagesMsg06 & vbCrLf & vbCrLf & sFiles(0), vbOKOnly Or vbExclamation
50380        End If
50390        DoEvents
50400       Else
50410        aLen = 0
50420        For i = 1 To UBound(sFiles)
50430         aLen = aLen + FileLen(sFiles(i))
50440        Next i
50450        For i = 1 To UBound(sFiles)
50460         If CheckIfPSFile(sFiles(0)) Then
50470           tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PA")
50480           Kill tFilename
50490           DoEvents
50500           FileCopy sFiles(i), tFilename
50510          Else
50520           aw = MsgBox(LanguageStrings.MessagesMsg06 & vbCrLf & vbCrLf & sFiles(i), vbOKCancel Or vbExclamation)
50530           If aw = vbCancel Then
50540            Screen.MousePointer = vbNormal
50550            Exit Sub
50560           End If
50570         End If
50580         tLen = tLen + FileLen(sFiles(i))
50590         stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
50600         DoEvents
50610        Next i
50620      End If
50630     End If
50640    End With
50650    stb.Panels("Percent").Text = vbNullString
50660   Case 3: ' Delete
50670    For i = 1 To lsv.ListItems.Count
50680     If lsv.ListItems(i).Selected = True Then
50690      Kill lsv.ListItems(i).SubItems(4)
50700     End If
50710     DoEvents
50720    Next i
50730    LvwRemoveSelectedItems lsv, True
50740   Case 5: ' Top
50750    For j = 1 To LvwGetCountSelectedItems(lsv, True)
50760     For i = lsv.ListItems.Count To 1 Step -1
50770      If lsv.ListItems(i).Selected = True Then
50780       LvwListItemToTop lsv, i, True
50790       Exit For
50800      End If
50810     Next i
50820    Next j
50830   Case 6: ' Up
50840    LvwListItemUp lsv, , True
50850   Case 7: ' Down
50860    LvwListItemDown lsv, , True
50870   Case 8: ' Bottom
50880    For j = 1 To LvwGetCountSelectedItems(lsv, True)
50890     For i = 1 To lsv.ListItems.Count
50900      If lsv.ListItems(i).Selected = True Then
50910       LvwListItemToBottom lsv, i, True
50920       Exit For
50930      End If
50940     Next i
50950    Next j
50960   Case 10: ' Combine
50970    Set cFiles = New Collection
50980    For i = 1 To lsv.ListItems.Count
50990     If lsv.ListItems(i).Selected = True Then
51000      cFiles.Add lsv.ListItems(i).SubItems(4)
51010     End If
51020    Next i
51030    tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PC")
51040    Kill tFilename
51050    If cFiles.Count > 1 Then
51060     CombineFiles tFilename, cFiles, stb
51070    End If
51080    Set cFiles = Nothing
51090  End Select
51100  Screen.MousePointer = vbNormal
51110  Timer1.Enabled = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnDocument_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnDocumentMain_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetDocumentMenu
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnDocumentMain_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnHelp_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Select Case Index
  Case 0:
50030    frmInfo.Show vbModal, Me
50040  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnHelp_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnPrinter_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Select Case Index
  Case 0:
50030    If mnPrinter(Index).Checked = False Then
50040      SetPrinterStop True
50050      mnPrinter(Index).Checked = True
50060     Else
50070      SetPrinterStop False
50080      mnPrinter(Index).Checked = False
50090    End If
50100   Case 2:
50110    frmOptions.Show , Me
50120   Case 4:
50130    If mnPrinter(Index).Checked = False Then
50140      SetLogging True
50150      mnPrinter(Index).Checked = True
50160     Else
50170      SetLogging False
50180      mnPrinter(Index).Checked = False
50190    End If
50200   Case 5:
50210    frmLog.Show , Me
50220   Case 7:
50230    End
50240  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnPrinter_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnLanguage_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim File As String
50020  Screen.MousePointer = vbHourglass
50030  LoadLanguage mnLanguage(Index).Tag
50040  Languagefile = mnLanguage(Index).Tag
50050  SetLanguageMenu
50060  SplitPath Languagefile, , , , File
50070  SetLanguage File
50080  Me.Refresh
50090  Screen.MousePointer = vbNormal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnLanguage_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnView_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Select Case Index
  Case 0:
50030    stb.Visible = Not stb.Visible
50040    mnView(0).Checked = Not mnView(0).Checked
50050    Form_Resize
50060  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnView_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Timer1_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Timer1.Enabled = False
50020  DoEvents
50030  If Len(Dir(App.Path & "\Unload.tmp")) > 0 Then
50040   End
50050  End If
50060  CheckPrintJobs
50070  CheckForPrinting
50080  If lsv.ListItems.Count = 0 And UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
50090   End
50100  End If
50110  If lsv.ListItems.Count = 1 Then
50120   lsv.ListItems(1).Selected = True
50130  End If
50140  DoEvents
50150  Timer1.Interval = TimerIntervall
50160  Timer1.Enabled = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Timer1_Timer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CheckForPrinting()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If lsv.ListItems.Count > 0 Then
50020   If mnPrinter(0).Checked = True Then
50030     lsv.ListItems(1).SubItems(1) = LanguageStrings.ListWaiting
50040    Else
50050     lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting
50060     PDFSpoolfile = lsv.ListItems(1).SubItems(4)
50070     If PrinterStop = False Then
50080      If IsFormLoaded(frmPrinting) = False Then
50090       frmPrinting.Show , Me
50100       If Me.Visible = True Then
50110        Me.Show
50120       End If
50130      End If
50140     End If
50150     If PrinterStop = False Then
50160       mnPrinter(0).Checked = False
50170      Else
50180       mnPrinter(0).Checked = True
50190     End If
50200   End If
50210  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "CheckForPrinting")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CheckPrintJobs()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Temppath As String, LItem As ListItem, tColl As Collection, _
  tFile() As String, i As Long, j As Long, kB As Long, MB As Long, GB As Long
50030  kB = 1024: MB = kB * 1024: GB = MB * 1024
50040  Set tColl = New Collection
50050  Temppath = GetPDFCreatorTempfolder
50060  Set tColl = GetFiles(GetPDFCreatorTempfolder, "~PD*.tmp")
50070  If tColl.Count = 0 And lsv.ListItems.Count > 0 Then
50080   lsv.ListItems.Clear
50090  End If
50100  For i = 1 To tColl.Count
50110   tFile = Split(tColl.item(i), "|")
50120   For j = 1 To lsv.ListItems.Count
50130    If UCase$(tFile(1)) = UCase$(lsv.ListItems(j).SubItems(4)) Then
50140     Exit For
50150    End If
50160   Next j
50170   If j > lsv.ListItems.Count Then
50180     SetTopMost Me, True, True
50190     SetTopMost Me, False, True
50200     Set LItem = lsv.ListItems.Add(, , GetPDFTitle(tFile(1)))
50210     LItem.SubItems(1) = LanguageStrings.ListWaiting
50220     LItem.SubItems(2) = tFile(3)
50230     If CLng(tFile(2)) > GB Then
50240       LItem.SubItems(3) = Format$(CDbl(tFile(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
50250      Else
50260       If CLng(tFile(2)) > MB Then
50270         LItem.SubItems(3) = Format$(CDbl(tFile(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
50280        Else
50290         If CLng(tFile(2)) > kB Then
50300           LItem.SubItems(3) = Format$(CDbl(tFile(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
50310          Else
50320           LItem.SubItems(3) = Format$(tFile(2), "0 " & LanguageStrings.ListBytes)
50330         End If
50340      End If
50350     End If
50360     LItem.SubItems(4) = tFile(1)
50370     DoEvents
50380    Else
50390 '
50400   End If
50410  Next i
50420  i = 0
50430  Do Until i + 1 >= lsv.ListItems.Count
50440   i = i + 1
50450   For j = 1 To tColl.Count
50460    tFile = Split(tColl.item(j), "|")
50470    If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
50480     Exit For
50490    End If
50500   Next j
50510   If j > tColl.Count Then
50520    lsv.ListItems.Remove i
50530   End If
50540   DoEvents
50550  Loop
50560  If lsv.ListItems.Count = 1 Then
50570    stb.Panels("Status").Text = "Status: " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg01
50580   Else
50590    stb.Panels("Status").Text = "Status: " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg02
50600  End If
50610  Set tColl = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "CheckPrintJobs")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetDocumentMenu()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim c As Long
50020  If lsv.ListItems.Count = 0 Then
50030    mnDocument(0).Enabled = False
50040    mnDocument(3).Enabled = False
50050    mnDocument(5).Enabled = False
50060    mnDocument(6).Enabled = False
50070    mnDocument(7).Enabled = False
50080    mnDocument(8).Enabled = False
50090    mnDocument(10).Enabled = False
50100    Exit Sub
50110   Else
50120    If lsv.ListItems.Count = 1 Then
50130     mnDocument(0).Enabled = True
50140     mnDocument(3).Enabled = True
50150     mnDocument(5).Enabled = False
50160     mnDocument(6).Enabled = False
50170     mnDocument(7).Enabled = False
50180     mnDocument(8).Enabled = False
50190     mnDocument(10).Enabled = False
50200     Exit Sub
50210    End If
50220  End If
50230  mnDocument(0).Enabled = True
50240  mnDocument(3).Enabled = True
50250  mnDocument(5).Enabled = True
50260  mnDocument(6).Enabled = True
50270  mnDocument(7).Enabled = True
50280  mnDocument(8).Enabled = True
50290  mnDocument(10).Enabled = False
50300  c = LvwGetCountSelectedItems(lsv, True)
50310  If c > 1 Then
50320   mnDocument(10).Enabled = True
50330  End If
50340  If lsv.SelectedItem.Index = 1 And c <= 1 Then
50350   mnDocument(5).Enabled = False
50360   mnDocument(6).Enabled = False
50370   mnDocument(7).Enabled = True
50380   mnDocument(8).Enabled = True
50390  End If
50400  If lsv.SelectedItem.Index = lsv.ListItems.Count And c <= 1 Then
50410   mnDocument(5).Enabled = True
50420   mnDocument(6).Enabled = True
50430   mnDocument(7).Enabled = False
50440   mnDocument(8).Enabled = False
50450  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SetDocumentMenu")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CheckAutosaveAndPrint()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tColl As Collection, i As Long, tFile() As String, Pathname As String
50020
50030  If Options.UseAutosave = 1 Then
50040   Set tColl = GetFiles(GetPDFCreatorTempfolder, "~P*.tmp")
50050   For i = 1 To tColl.Count
50060    tFile = Split(tColl.item(i), "|")
50070    SplitPath GetAutosaveFilename(tFile(1)), , Pathname
50080    If Len(Dir(Pathname, vbDirectory)) = 0 Then
50090      If Options.UseAutosaveDirectory = 1 Then
50100        IfLoggingWriteLogfile "Error: AutoSaveDirectory not found."
50110       Else
50120        IfLoggingWriteLogfile "Error: LastSaveDirectory not found."
50130      End If
50140     Else
50150      CallGScript tFile(1), GetAutosaveFilename(tFile(1)), Options, Options.AutosaveFormat
50160      Kill tFile(1)
50170    End If
50180   Next i
50190   End
50200  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "CheckAutosaveAndPrint")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
