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

Const TimerIntervall = 500

Private LanguagePath As String, Languagefile As String, mutex As clsMutex, _
 Printjobs As Collection

Private Sub Form_Load()
 Dim fn As Long, stdio As clsStdIO, cinStr As String, Tempfile As String

'##############################################
'Performance Tools
 Dim LastStop As Currency, ct As Integer
 LastStop = ExactTimer_Value()
'##############################################

 PDFCreatorINIFile = App.Path & "\PDFCreator.ini"
 Options = ReadOptions
 IfLoggingWriteLogfile "PDFCreator Program Start"
 
 ' The program has commandswitches
 ' -IPTRUE : Install Printer
 ' -IPFALSE: UnInstall Printer
 ' -NSTRUE: No Start
 ' -ULTRUE: Unload all PDFCreator programs
 ' -PPDFCREATORPRINTER: The printer call the program
 
 
 ' Check Installprinter
 Select Case UCase$(CommandSwitch("IP", True))
  Case "TRUE":
   Monitorname = "PDFCreator": Portname = "PDFCreator:": Drivername = "PDFCreator": PrinterName = "PDFCreator"
   InstallCompletePrinter
  Case "FALSE":
   Monitorname = "PDFCreator": Portname = "PDFCreator:": Drivername = "PDFCreator": PrinterName = "PDFCreator"
   UnInstallCompletePrinter
 End Select
 
 ' Initialize unload running program
 If UCase$(CommandSwitch("UL", True)) = "TRUE" Then
  fn = FreeFile
  Open App.Path & "\Unload.tmp" For Output As #fn
  Close #fn
 End If
 
 ' NS: If NS=True Then end the program here
 ' It is necessary for uninstall.
 If UCase$(CommandSwitch("NS", True)) = "TRUE" Then
  End
 End If
 
 CreatePDFCreatorTempfolder
 
 Set stdio = New clsStdIO
 cinStr = stdio.StdIn
 If Len(cinStr) > 0 Then
  Tempfile = GetTempFile(GetTempPath & "PDFCreator\", "~PD")
  fn = FreeFile
  Open Tempfile For Output As #fn
  Print #fn, cinStr
  Close #fn
 End If
 Set stdio = Nothing
 
 ' Printer has started the program
 If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
  CheckAutosaveAndPrint
 End If
 InitProgram
 
 If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
  If lsv.ListItems.Count <= 1 Then
   Me.Visible = False
  End If
 End If
'##############################################
'MsgBox "Programmstart: " & ExactTimer_Value() - LastStop & " Sekunden"
'LastStop = ExactTimer_Value()
'##############################################
 'IfLoggingWriteLogfile "PDFCreator started in " & ExactTimer_Value() - LastStop & " seconds"
End Sub

Private Sub Form_Resize()
 If Me.WindowState = vbMinimized Then
  Exit Sub
 End If
 If Me.Height < 3000 Then
  Me.Height = 3000
  Exit Sub
 End If
 If Me.Width < 3000 Then
  Me.Width = 3000
  Exit Sub
 End If
 With lsv
  .Top = 0: .Left = 0
  .Width = Me.Width - 125
  .Height = Me.ScaleHeight - Abs(stb.Visible) * stb.Height
 End With
 stb.Panels("Status").Width = Me.Width - 150 - stb.Panels("Percent").Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
 TerminateProgram
 End
End Sub

Private Sub InitProgram()
 Dim Filename As String, Tempfile As String, res As Long
 
 Printing = False
 Filename = CommandSwitch("F", True)
 
 If Dir(Filename) <> "" And Len(Trim$(Filename)) > 0 Then
  If FileLen(Filename) > 0 Then
   Tempfile = GetTempFile(GetTempPath & "PDFCreator\", "~PD")
   FileCopy Filename, Tempfile
  End If
 End If
 
 Set Printjobs = New Collection
 Set mutex = New clsMutex
 
 If mutex.CheckMutex(PDFCreator_GUID) = False Then
   res = mutex.CreateMutex(PDFCreator_GUID)
  Else
   End
 End If
 
 stb.Panels.Clear
 stb.Panels.Add , "Status", ""
 stb.Panels.Add , "Percent", ""
 stb.Panels("Percent").Width = 1000
 
 With lsv
  .View = lvwReport
  .FullRowSelect = True
  .HideSelection = False
  .ColumnHeaders.Clear
  .ColumnHeaders.Add , "Documenttitle", "Documenttitle", 2000
  .ColumnHeaders.Add , "Status", "Status", 1500
  .ColumnHeaders.Add , "Date", "Created on", 1700
  .ColumnHeaders.Add , "Size", "Size", 1500, lvwColumnRight
  .ColumnHeaders.Add , "Filename", "Filename", lsv.Width - 3500
 End With
 
 LanguagePath = App.Path & "\Languages\"
 ReadAllLanguages LanguagePath

 With Options
  SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
 End With
 
 Languagefile = LanguagePath & Options.Language & ".ini"
 
 LoadLanguage Languagefile
 SetLanguageMenu
 If Options.Logging = 1 Then
   mnPrinter(4).Checked = True
  Else
   mnPrinter(4).Checked = False
 End If
 
 CheckPrintJobs
 DoEvents

 ' Only for the first time set Interval to 10 ms
 Timer1.Interval = 10
 Timer1.Enabled = True
End Sub

Private Sub TerminateProgram()
 Timer1.Enabled = False
 Set Printjobs = Nothing
 mutex.CloseMutex
 Set mutex = Nothing
 IfLoggingWriteLogfile "PDFCreator Program End"
End Sub

Private Function GetAllLanguagesFiles(LanguagePath As String) As Collection
 Dim Languagefile As String
 Set GetAllLanguagesFiles = New Collection
 Languagefile = Dir(LanguagePath & "*.ini")
 Do While Languagefile <> ""
   GetAllLanguagesFiles.Add LanguagePath & Languagefile
   Languagefile = Dir()
  DoEvents
 Loop
End Function

Private Sub ReadAllLanguages(LanguagePath As String)
 Dim Languagename As String, ini As clsINI, LangFiles As Collection, i As Long
 mnLanguage(0).Caption = "No languages available."
  
 Set LangFiles = GetAllLanguagesFiles(LanguagePath)
 Set ini = New clsINI
 For i = 1 To LangFiles.Count
  ini.Filename = LangFiles.item(i)
  ini.Section = "Common"
  Languagename = ini.GetKeyFromSection("Languagename")
  If Languagename = "" Then
   Languagename = "No name available."
  End If
  Load mnLanguage(mnLanguage.Count)
  mnLanguage(mnLanguage.Count - 1).Caption = Languagename
  mnLanguage(mnLanguage.Count - 1).Tag = LangFiles.item(i)
  DoEvents
 Next i
 
 If mnLanguage.Count > 1 Then
  mnLanguage(0).Caption = "No languages available."
  mnLanguage(0).Visible = False
 End If
 Set ini = Nothing
End Sub

Private Sub SetLanguageMenu()
 Dim i As Long
 
 For i = mnLanguage.LBound To mnLanguage.UBound
  If UCase$(Languagefile) = UCase$(mnLanguage.item(i).Tag) Then
    mnLanguage.item(i).Checked = True
   Else
    mnLanguage.item(i).Checked = False
  End If
 Next i
 
 With LanguageStrings
  Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & _
   " " & .CommonTitle
  
  mnPrinterMain.Caption = .DialogPrinter
  mnPrinter(0).Caption = .DialogPrinterPrinterStop
  mnPrinter(2).Caption = .DialogPrinterOptions
  mnPrinter(4).Caption = .DialogPrinterLogging
  mnPrinter(5).Caption = .DialogPrinterLogfile
  mnPrinter(7).Caption = .DialogPrinterClose
  
  mnDocumentMain.Caption = .DialogDocument
  mnDocument(0).Caption = .DialogDocumentPrint
  mnDocument(2).Caption = .DialogDocumentAdd
  mnDocument(3).Caption = .DialogDocumentDelete
  mnDocument(5).Caption = .DialogDocumentTop
  mnDocument(6).Caption = .DialogDocumentUp
  mnDocument(7).Caption = .DialogDocumentDown
  mnDocument(8).Caption = .DialogDocumentBottom
  mnDocument(10).Caption = .DialogDocumentCombine
  
  mnViewMain.Caption = .DialogView
  mnView(0).Caption = .DialogViewStatusbar
  
  mnLanguageMain.Caption = .DialogLanguage
  
  mnHelp(0).Caption = .DialogInfo
  
  lsv.ColumnHeaders("Date").Text = .ListDate
  lsv.ColumnHeaders("Documenttitle").Text = .ListDocumenttitle
  lsv.ColumnHeaders("Filename").Text = .ListFilename
  lsv.ColumnHeaders("Size").Text = .ListSize
  lsv.ColumnHeaders("Status").Text = .ListStatus
 End With
End Sub

Private Sub lsv_KeyUp(KeyCode As Integer, Shift As Integer)
 Dim i As Long
 If KeyCode = 46 Then
  For i = 1 To lsv.ListItems.Count
   If lsv.ListItems(i).Selected = True Then
    Kill lsv.ListItems(i).SubItems(4)
   End If
   DoEvents
  Next i
  LvwRemoveSelectedItems lsv, True
 End If
End Sub

Private Sub lsv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
  SetDocumentMenu
  PopupMenu mnDocumentMain, , x, y
End If
End Sub

Private Sub lsv_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim tFilename As String, i As Long, aLen As Double, tLen As Double
  
If data.GetFormat(vbCFFiles) Then
     If data.Files.Count = 1 Then
       If CheckIfPSFile(data.Files.item(1)) Then
        tFilename = GetTempFile(GetTempPath & "PDFCreator\", "~PA")
        FileCopy data.Files.item(1), tFilename
       End If
       DoEvents
      Else
       aLen = 0
       For i = 1 To data.Files.Count
        aLen = aLen + FileLen(data.Files.item(i))
       Next i
       For i = 1 To data.Files.Count
        If CheckIfPSFile(data.Files.item(i)) Then
         tFilename = GetTempFile(GetTempPath & "PDFCreator\", "~PA")
         DoEvents
         FileCopy data.Files.item(i), tFilename
        End If
         tLen = tLen + FileLen(data.Files.item(i))
         stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
         DoEvents
       Next i
     End If
End If
End Sub

Private Sub mnDocument_Click(Index As Integer)
 On Local Error Resume Next
 Dim tFilename As String, cFiles As Collection, sFiles() As String, _
  i As Long, j As Long, aLen As Double, tLen As Double, _
  aw As Long
 Timer1.Enabled = False
 Screen.MousePointer = vbHourglass
 DoEvents
 Select Case Index
  Case 0:
   For j = 1 To LvwGetCountSelectedItems(lsv, True)
    DoEvents
    For i = lsv.ListItems.Count To 1 Step -1
     If lsv.ListItems(i).Selected = True Then
      lsv.ListItems(i).SubItems(1) = LanguageStrings.ListPrinting
      LvwListItemToTop lsv, i, True
      Exit For
     End If
    Next i
   Next j
   SetPrinterStop False
   mnPrinter(0).Checked = False
  Case 2: ' Add
   DoEvents
   With cdlg
    .InitDir = GetSpecialFolder(ssfPERSONAL)
    .Flags = cdlOFNFileMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNLongNames
    .DialogTitle = LanguageStrings.ListAddPostscriptFile
    .Filter = LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps|"
    .Filename = ""
    .ShowOpen
    If .Filename <> "" Then
     sFiles = Split(.Filename, Chr$(0))
     If UBound(sFiles) = 0 Then
       If CheckIfPSFile(sFiles(0)) Then
         tFilename = GetTempFile(GetTempPath & "PDFCreator\", "~PA")
         Kill tFilename
         FileCopy sFiles(0), tFilename
        Else
         MsgBox LanguageStrings.MessagesMsg06 & vbCrLf & vbCrLf & sFiles(0), vbOKOnly Or vbExclamation
       End If
       DoEvents
      Else
       aLen = 0
       For i = 1 To UBound(sFiles)
        aLen = aLen + FileLen(sFiles(i))
       Next i
       For i = 1 To UBound(sFiles)
        If CheckIfPSFile(sFiles(0)) Then
          tFilename = GetTempFile(GetTempPath & "PDFCreator\", "~PA")
          Kill tFilename
          DoEvents
          FileCopy sFiles(i), tFilename
         Else
          aw = MsgBox(LanguageStrings.MessagesMsg06 & vbCrLf & vbCrLf & sFiles(i), vbOKCancel Or vbExclamation)
          If aw = vbCancel Then
           Screen.MousePointer = vbNormal
           Exit Sub
          End If
        End If
        tLen = tLen + FileLen(sFiles(i))
        stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
        DoEvents
       Next i
     End If
    End If
   End With
   stb.Panels("Percent").Text = ""
  Case 3: ' Delete
   For i = 1 To lsv.ListItems.Count
    If lsv.ListItems(i).Selected = True Then
     Kill lsv.ListItems(i).SubItems(4)
    End If
    DoEvents
   Next i
   LvwRemoveSelectedItems lsv, True
  Case 5: ' Top
   For j = 1 To LvwGetCountSelectedItems(lsv, True)
    For i = lsv.ListItems.Count To 1 Step -1
     If lsv.ListItems(i).Selected = True Then
      LvwListItemToTop lsv, i, True
      Exit For
     End If
    Next i
   Next j
  Case 6: ' Up
   LvwListItemUp lsv, , True
  Case 7: ' Down
   LvwListItemDown lsv, , True
  Case 8: ' Bottom
   For j = 1 To LvwGetCountSelectedItems(lsv, True)
    For i = 1 To lsv.ListItems.Count
     If lsv.ListItems(i).Selected = True Then
      LvwListItemToBottom lsv, i, True
      Exit For
     End If
    Next i
   Next j
  Case 10: ' Combine
   Set cFiles = New Collection
   For i = 1 To lsv.ListItems.Count
    If lsv.ListItems(i).Selected = True Then
     cFiles.Add lsv.ListItems(i).SubItems(4)
    End If
   Next i
   tFilename = GetTempFile(GetTempPath & "PDFCreator\", "~PC")
   Kill tFilename
   If cFiles.Count > 1 Then
    CombineFiles tFilename, cFiles, stb
   End If
   Set cFiles = Nothing
 End Select
 Screen.MousePointer = vbNormal
 Timer1.Enabled = True
End Sub

Private Sub mnDocumentMain_Click()
 SetDocumentMenu
End Sub

Private Sub mnHelp_Click(Index As Integer)
 Select Case Index
  Case 0:
   frmInfo.Show vbModal, Me
 End Select
End Sub

Private Sub mnPrinter_Click(Index As Integer)
 Select Case Index
  Case 0:
   If mnPrinter(Index).Checked = False Then
     SetPrinterStop True
     mnPrinter(Index).Checked = True
    Else
     SetPrinterStop False
     mnPrinter(Index).Checked = False
   End If
  Case 2:
   frmOptions.Show , Me
  Case 4:
   If mnPrinter(Index).Checked = False Then
     SetLogging True
     mnPrinter(Index).Checked = True
    Else
     SetLogging False
     mnPrinter(Index).Checked = False
   End If
  Case 5:
   frmLog.Show , Me
  Case 7:
   End
 End Select
End Sub

Private Sub mnLanguage_Click(Index As Integer)
 Dim File As String
 Screen.MousePointer = vbHourglass
 LoadLanguage mnLanguage(Index).Tag
 Languagefile = mnLanguage(Index).Tag
 SetLanguageMenu
 SplitPath Languagefile, , , , File
 SetLanguage File
 Me.Refresh
 Screen.MousePointer = vbNormal
End Sub

Private Sub mnView_Click(Index As Integer)
 Select Case Index
  Case 0:
   stb.Visible = Not stb.Visible
   mnView(0).Checked = Not mnView(0).Checked
   Form_Resize
 End Select
End Sub

Private Sub Timer1_Timer()
 Timer1.Enabled = False
 DoEvents
 If Dir(App.Path & "\Unload.tmp") <> "" Then
  End
 End If
 CheckPrintJobs
 CheckForPrinting
 If lsv.ListItems.Count = 0 And UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
  End
 End If
 If lsv.ListItems.Count = 1 Then
  lsv.ListItems(1).Selected = True
 End If
 DoEvents
 Timer1.Interval = TimerIntervall
 Timer1.Enabled = True
End Sub

Private Sub CheckForPrinting()
 If lsv.ListItems.Count > 0 Then
  If mnPrinter(0).Checked = True Then
    lsv.ListItems(1).SubItems(1) = LanguageStrings.ListWaiting
   Else
    lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting
    PDFSpoolfile = lsv.ListItems(1).SubItems(4)
    If PrinterStop = False Then
     If IsFormLoaded(frmPrinting) = False Then
      frmPrinting.Show , Me
      If Me.Visible = True Then
       Me.Show
      End If
     End If
    End If
    If PrinterStop = False Then
      mnPrinter(0).Checked = False
     Else
      mnPrinter(0).Checked = True
    End If
  End If
 End If
End Sub

Private Sub CheckPrintJobs()
 On Local Error Resume Next
 Dim Temppath As String, LItem As ListItem, tColl As Collection, _
  tFile() As String, i As Long, j As Long, kB As Long, MB As Long, GB As Long
 kB = 1024: MB = kB * 1024: GB = MB * 1024
 Set tColl = New Collection
 Temppath = GetTempPath
 Set tColl = GetFiles(Temppath & "PDFCreator\", "~P*.tmp")
 If tColl.Count = 0 And lsv.ListItems.Count > 0 Then
  lsv.ListItems.Clear
 End If
 For i = 1 To tColl.Count
  tFile = Split(tColl.item(i), "|")
  For j = 1 To lsv.ListItems.Count
   If UCase$(tFile(1)) = UCase$(lsv.ListItems(j).SubItems(4)) Then
    Exit For
   End If
  Next j
  If j > lsv.ListItems.Count Then
    SetTopMost Me, True, True
    SetTopMost Me, False, True
    Set LItem = lsv.ListItems.Add(, , GetPDFTitle(tFile(1)))
    LItem.SubItems(1) = LanguageStrings.ListWaiting
    LItem.SubItems(2) = tFile(3)
    If CLng(tFile(2)) > GB Then
      LItem.SubItems(3) = Format$(CDbl(tFile(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
     Else
      If CLng(tFile(2)) > MB Then
        LItem.SubItems(3) = Format$(CDbl(tFile(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
       Else
        If CLng(tFile(2)) > kB Then
          LItem.SubItems(3) = Format$(CDbl(tFile(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
         Else
          LItem.SubItems(3) = Format$(tFile(2), "0 " & LanguageStrings.ListBytes)
        End If
     End If
    End If
    LItem.SubItems(4) = tFile(1)
    DoEvents
   Else
'
  End If
 Next i
 i = 0
 Do Until i + 1 >= lsv.ListItems.Count
  i = i + 1
  For j = 1 To tColl.Count
   tFile = Split(tColl.item(j), "|")
   If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
    Exit For
   End If
  Next j
  If j > tColl.Count Then
   lsv.ListItems.Remove i
  End If
  DoEvents
 Loop
 If lsv.ListItems.Count = 1 Then
   stb.Panels("Status").Text = "Status: " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg01
  Else
   stb.Panels("Status").Text = "Status: " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg02
 End If
 Set tColl = Nothing
End Sub

Private Sub SetDocumentMenu()
 Dim c As Long
 If lsv.ListItems.Count = 0 Then
   mnDocument(0).Enabled = False
   mnDocument(3).Enabled = False
   mnDocument(5).Enabled = False
   mnDocument(6).Enabled = False
   mnDocument(7).Enabled = False
   mnDocument(8).Enabled = False
   mnDocument(10).Enabled = False
   Exit Sub
  Else
   If lsv.ListItems.Count = 1 Then
    mnDocument(0).Enabled = True
    mnDocument(3).Enabled = True
    mnDocument(5).Enabled = False
    mnDocument(6).Enabled = False
    mnDocument(7).Enabled = False
    mnDocument(8).Enabled = False
    mnDocument(10).Enabled = False
    Exit Sub
   End If
 End If
 mnDocument(0).Enabled = True
 mnDocument(3).Enabled = True
 mnDocument(5).Enabled = True
 mnDocument(6).Enabled = True
 mnDocument(7).Enabled = True
 mnDocument(8).Enabled = True
 mnDocument(10).Enabled = False
 c = LvwGetCountSelectedItems(lsv, True)
 If c > 1 Then
  mnDocument(10).Enabled = True
 End If
 If lsv.SelectedItem.Index = 1 And c <= 1 Then
  mnDocument(5).Enabled = False
  mnDocument(6).Enabled = False
  mnDocument(7).Enabled = True
  mnDocument(8).Enabled = True
 End If
 If lsv.SelectedItem.Index = lsv.ListItems.Count And c <= 1 Then
  mnDocument(5).Enabled = True
  mnDocument(6).Enabled = True
  mnDocument(7).Enabled = False
  mnDocument(8).Enabled = False
 End If
End Sub

Private Sub CheckAutosaveAndPrint()
 On Local Error Resume Next
 Dim tColl As Collection, i As Long, tFile() As String, Pathname As String
 If Options.UseAutosave = 1 Then
  Set tColl = GetFiles(GetTempPath & "PDFCreator\", "~P*.tmp")
  For i = 1 To tColl.Count
   tFile = Split(tColl.item(i), "|")
   SplitPath GetAutosaveFilename(tFile(1)), , Pathname
   If Dir(Pathname, vbDirectory) = "" Then
     If Options.UseAutosaveDirectory = 1 Then
       IfLoggingWriteLogfile "Error: AutoSaveDirectory not found."
      Else
       IfLoggingWriteLogfile "Error: LastSaveDirectory not found."
     End If
    Else
     CallGScript tFile(1), GetAutosaveFilename(tFile(1)), Options, Options.AutosaveFormat
     Kill tFile(1)
   End If
  Next i
  End
 End If
End Sub
