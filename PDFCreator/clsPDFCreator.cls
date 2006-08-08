VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDFCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event eReady()
Public Event eError()

Private Const ErrDescr_2 = "Unknown Event!"
Private Const ErrDescr_1 = "Test Error Event."
Private Const ErrDescr1 = "The ActiveX-Server has not been started! Please use function ""cStart()"" to start the ActiveX-Server!"
Private Const ErrDescr2 = "An instance of the PDFCreator is already running. It is not allowed to have more than one instance running!"
Private Const ErrDescr3 = "Unknown Options Name '%1' in '%2'!"
Private Const ErrDescr4 = "Ghostscript Error: '%1'!"
Private Const ErrDescr5 = "The file '%1' could not be found!"
Private Const ErrDescr6 = "No valid path '%1' was given!"
Private Const ErrDescr7 = "The file '%1' is in use!"
Private Const ErrDescr8 = "Function '%1': Options1 is not a 'clsPDFCreatorOptions' type!"
Private Const ErrDescr9 = "No valid postscript file '%1' was given!"

Private mForm As frmMain, mError As clsPDFCreatorError, mProgramIsRunning As Boolean
Private WithEvents eActionTimer As Timer
Attribute eActionTimer.VB_VarHelpID = -1

' --- Class subroutines - begin ---
Private Sub Class_Initialize()
 Set mError = New clsPDFCreatorError
 InstanceCounter = InstanceCounter + 1
 ProgramIsVisible = True
End Sub

Private Sub Class_Terminate()
 InstanceCounter = InstanceCounter - 1
 If InstanceCounter = 0 Then
  cCloseRunningSession
 End If
 Set mError = Nothing
 Set mForm = Nothing
 Set eActionTimer = Nothing
End Sub
' --- Class subroutines - end ---

' --- Public properties - begin ---
Public Property Get cError() As clsPDFCreatorError
 Set cError = mError
End Property

Public Property Get cErrorDetail(ByVal PropertyName As String) As Variant
 Select Case UCase$(PropertyName)
  Case "DESCRIPTION": cErrorDetail = mError.Description
  Case "NUMBER": cErrorDetail = mError.Number
  Case Else:
   mError.Number = 3
   mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in cErrorDetail")
   RaiseEvent eError
 End Select
End Property

Public Sub cErrorClear()
 mError.Number = 0
 mError.Description = ""
End Sub

Public Property Get cOutputFilename() As String
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 cOutputFilename = ConvertedOutputFilename
End Property

Public Property Get cPDFCreatorApplicationPath() As String
 cPDFCreatorApplicationPath = GetPDFCreatorApplicationPath
End Property

Public Property Get cIsLogfileDialogDisplayed() As Boolean
 Dim f As Form
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 cIsLogfileDialogDisplayed = False
 For Each f In Forms
  If UCase$(f.Name) = "FRMLOG" Then
   cIsLogfileDialogDisplayed = True
  End If
 Next f
End Property

Public Property Get cIsOptionsDialogDisplayed() As Boolean
 Dim f As Form
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 cIsOptionsDialogDisplayed = False
 For Each f In Forms
  If UCase$(f.Name) = "FRMOPTIONS" Then
   cIsOptionsDialogDisplayed = True
  End If
 Next f
End Property

Public Property Get cProgramRelease(Optional ByVal WithBeta As Boolean = True) As String
 cProgramRelease = GetProgramRelease(WithBeta)
End Property

Public Property Get cProgramIsRunning() As Boolean
 cProgramIsRunning = ProgramIsRunning(PDFCreator_GUID)
End Property

Public Property Get cWindowsVersion() As String
 cWindowsVersion = GetWinVersionStr
End Property

Public Property Get cVisible() As Boolean
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 cVisible = ProgramIsVisible
End Property

Public Property Let cVisible(ByVal Value As Boolean)
 Dim f As Form
 ProgramIsVisible = Value
 If Not mForm Is Nothing Then
  For Each f In Forms
   If Not f Is Nothing Then
    If UCase$(f.Name) <> "FRMSYSTRAY" Then
     f.Visible = ProgramIsVisible
    End If
   End If
  Next f
  If ProgramIsVisible Then
    If mForm.WindowState = vbMinimized Then
      FormInTaskbar mForm, False, False, False
      mForm.SystrayEnter
     Else
      mForm.SysTrayLeave
      FormInTaskbar mForm, True, False, True
    End If
   Else
    If mForm.WindowState = vbMinimized Then
     mForm.SysTrayLeave
    End If
  End If
 End If
End Property

Public Property Get cInstalledAsServer() As Boolean
 cInstalledAsServer = InstalledAsServer
End Property

Public Property Get cPrinterStop() As Boolean
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 cPrinterStop = PrinterStop
End Property

Public Property Let cPrinterStop(ByVal Value As Boolean)
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 If Not mForm Is Nothing Then
  If Value Then
    If Not PrinterStop Then
     mForm.SetMenuPrinterStop
    End If
   Else
    If PrinterStop Then
     mForm.SetMenuPrinterStop
    End If
  End If
 End If
End Property

Public Property Get cOptionsNames() As Collection
 Set cOptionsNames = New Collection
 With cOptionsNames
  .Add "AdditionalGhostscriptParameters"
  .Add "AdditionalGhostscriptSearchpath"
  .Add "AddWindowsFontpath"
  .Add "AutosaveDirectory"
  .Add "AutosaveFilename"
  .Add "AutosaveFormat"
  .Add "AutosaveStartStandardProgram"
  .Add "BitmapResolution"
  .Add "BMPColorscount"
  .Add "ClientComputerResolveIPAddress"
  .Add "DeviceHeightPoints"
  .Add "DeviceWidthPoints"
  .Add "DirectoryGhostscriptBinaries"
  .Add "DirectoryGhostscriptFonts"
  .Add "DirectoryGhostscriptLibraries"
  .Add "DirectoryGhostscriptResource"
  .Add "DisableEmail"
  .Add "DontUseDocumentSettings"
  .Add "EPSLanguageLevel"
  .Add "FilenameSubstitutions"
  .Add "FilenameSubstitutionsOnlyInTitle"
  .Add "JPEGColorscount"
  .Add "JPEGQuality"
  .Add "Language"
  .Add "LastSaveDirectory"
  .Add "Logging"
  .Add "LogLines"
  .Add "NoConfirmMessageSwitchingDefaultprinter"
  .Add "NoProcessingAtStartup"
  .Add "NoPSCheck"
  .Add "OnePagePerFile"
  .Add "OptionsDesign"
  .Add "OptionsEnabled"
  .Add "OptionsVisible"
  .Add "Papersize"
  .Add "PCXColorscount"
  .Add "PDFAllowAssembly"
  .Add "PDFAllowDegradedPrinting"
  .Add "PDFAllowFillIn"
  .Add "PDFAllowScreenReaders"
  .Add "PDFColorsCMYKToRGB"
  .Add "PDFColorsColorModel"
  .Add "PDFColorsPreserveHalftone"
  .Add "PDFColorsPreserveOverprint"
  .Add "PDFColorsPreserveTransfer"
  .Add "PDFCompressionColorCompression"
  .Add "PDFCompressionColorCompressionChoice"
  .Add "PDFCompressionColorCompressionJPEGHighFactor"
  .Add "PDFCompressionColorCompressionJPEGLowFactor"
  .Add "PDFCompressionColorCompressionJPEGMaximumFactor"
  .Add "PDFCompressionColorCompressionJPEGMediumFactor"
  .Add "PDFCompressionColorCompressionJPEGMinimumFactor"
  .Add "PDFCompressionColorResample"
  .Add "PDFCompressionColorResampleChoice"
  .Add "PDFCompressionColorResolution"
  .Add "PDFCompressionGreyCompression"
  .Add "PDFCompressionGreyCompressionChoice"
  .Add "PDFCompressionGreyCompressionJPEGHighFactor"
  .Add "PDFCompressionGreyCompressionJPEGLowFactor"
  .Add "PDFCompressionGreyCompressionJPEGMaximumFactor"
  .Add "PDFCompressionGreyCompressionJPEGMediumFactor"
  .Add "PDFCompressionGreyCompressionJPEGMinimumFactor"
  .Add "PDFCompressionGreyResample"
  .Add "PDFCompressionGreyResampleChoice"
  .Add "PDFCompressionGreyResolution"
  .Add "PDFCompressionMonoCompression"
  .Add "PDFCompressionMonoCompressionChoice"
  .Add "PDFCompressionMonoResample"
  .Add "PDFCompressionMonoResampleChoice"
  .Add "PDFCompressionMonoResolution"
  .Add "PDFCompressionTextCompression"
  .Add "PDFDisallowCopy"
  .Add "PDFDisallowModifyAnnotations"
  .Add "PDFDisallowModifyContents"
  .Add "PDFDisallowPrinting"
  .Add "PDFEncryptor"
  .Add "PDFFontsEmbedAll"
  .Add "PDFFontsSubSetFonts"
  .Add "PDFFontsSubSetFontsPercent"
  .Add "PDFGeneralASCII85"
  .Add "PDFGeneralAutorotate"
  .Add "PDFGeneralCompatibility"
  .Add "PDFGeneralOverprint"
  .Add "PDFGeneralResolution"
  .Add "PDFHighEncryption"
  .Add "PDFLowEncryption"
  .Add "PDFOptimize"
  .Add "PDFOwnerPass"
  .Add "PDFOwnerPasswordString"
  .Add "PDFUserPass"
  .Add "PDFUserPasswordString"
  .Add "PDFUseSecurity"
  .Add "PNGColorscount"
  .Add "PrintAfterSaving"
  .Add "PrintAfterSavingDuplex"
  .Add "PrintAfterSavingNoCancel"
  .Add "PrintAfterSavingPrinter"
  .Add "PrintAfterSavingQueryUser"
  .Add "PrintAfterSavingTumble"
  .Add "PrinterStop"
  .Add "PrinterTemppath"
  .Add "ProcessPriority"
  .Add "ProgramFont"
  .Add "ProgramFontCharset"
  .Add "ProgramFontSize"
  .Add "PSLanguageLevel"
  .Add "RemoveAllKnownFileExtensions"
  .Add "RemoveSpaces"
  .Add "RunProgramAfterSaving"
  .Add "RunProgramAfterSavingProgramname"
  .Add "RunProgramAfterSavingProgramParameters"
  .Add "RunProgramAfterSavingWaitUntilReady"
  .Add "RunProgramAfterSavingWindowstyle"
  .Add "RunProgramBeforeSaving"
  .Add "RunProgramBeforeSavingProgramname"
  .Add "RunProgramBeforeSavingProgramParameters"
  .Add "RunProgramBeforeSavingWindowstyle"
  .Add "SaveFilename"
  .Add "SendEmailAfterAutoSaving"
  .Add "SendMailMethod"
  .Add "ShowAnimation"
  .Add "StampFontColor"
  .Add "StampFontname"
  .Add "StampFontsize"
  .Add "StampOutlineFontthickness"
  .Add "StampString"
  .Add "StampUseOutlineFont"
  .Add "StandardAuthor"
  .Add "StandardCreationdate"
  .Add "StandardDateformat"
  .Add "StandardKeywords"
  .Add "StandardMailDomain"
  .Add "StandardModifydate"
  .Add "StandardSaveformat"
  .Add "StandardSubject"
  .Add "StandardTitle"
  .Add "StartStandardProgram"
  .Add "TIFFColorscount"
  .Add "Toolbars"
  .Add "UseAutosave"
  .Add "UseAutosaveDirectory"
  .Add "UseCreationDateNow"
  .Add "UseCustomPaperSize"
  .Add "UseFixPapersize"
  .Add "UseStandardAuthor"
 End With
End Property

Public Property Get cOption(ByVal PropertyName As String) As Variant
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 cOption = cGetOptionFromOptions(PropertyName, Options)
End Property

Public Property Let cOption(ByVal PropertyName As String, ByVal Value As Variant)
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 cLetOption PropertyName, Value
End Property

Public Property Get cOptions() As clsPDFCreatorOptions
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 Set cOptions = cGetOptions(Options)
End Property

Public Property Set cOptions(ByVal Options1 As clsPDFCreatorOptions)
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 Options = cLetOptions(Options1)
End Property

Public Property Get cStandardOption(ByVal PropertyName As String) As Variant
 cStandardOption = cGetOptionFromOptions(PropertyName, StandardOptions)
End Property

Public Property Get cStandardOptions() As clsPDFCreatorOptions
 cStandardOptions = cGetOptions(StandardOptions)
End Property

Public Property Get cPostscriptInfo(ByVal PostscriptFilename As String, ByVal PropertyName As String) As String
 Dim PSHeader As tPSHeader
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 If FileExists(PostscriptFilename) = False Then
  mError.Number = 5
  mError.Description = Replace$(ErrDescr5, "%1", PostscriptFilename)
  RaiseEvent eError
  Exit Property
 End If
 If FileInUse(PostscriptFilename) Then
  mError.Number = 7
  mError.Description = Replace$(ErrDescr7, "%1", PostscriptFilename)
  RaiseEvent eError
  Exit Property
 End If
 PSHeader = GetPSHeader(PostscriptFilename)
 With PSHeader
  Select Case True
   Case ("BoundingBox") = PropertyName
    cPostscriptInfo = .BoundingBox.Comment
   Case ("CreateFor") = PropertyName
    cPostscriptInfo = .CreateFor.Comment
   Case ("CreationDate") = PropertyName
    cPostscriptInfo = .CreationDate.Comment
   Case ("Creator") = PropertyName
    cPostscriptInfo = .Creator.Comment
   Case ("DocumentData") = PropertyName
    cPostscriptInfo = .DocumentData.Comment
   Case ("DocumentNeededResources") = PropertyName
    cPostscriptInfo = .DocumentNeededResources.Comment
   Case ("DocumentSuppliedResources") = PropertyName
    cPostscriptInfo = .DocumentSuppliedResources.Comment
   Case ("LanguageLevel") = PropertyName
    cPostscriptInfo = .LanguageLevel.Comment
   Case ("Orientation") = PropertyName
    cPostscriptInfo = .Orientation.Comment
   Case ("PageOrder") = PropertyName
    cPostscriptInfo = .PageOrder.Comment
   Case ("Pages") = PropertyName
    cPostscriptInfo = .Pages.Comment
   Case ("TargetDevice") = PropertyName
    cPostscriptInfo = .TargetDevice.Comment
   Case ("Title") = PropertyName
    cPostscriptInfo = .Title.Comment
   Case Else
    mError.Number = 3
    mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in cPostscriptInfo")
    RaiseEvent eError
  End Select
 End With
End Property

Public Property Get cPrintjobInfos(ByVal PrintjobFilename As String) As clsPDFCreatorInfoSpoolFile
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 cPrintjobInfos = cGetInfoSpoolfile(ReadInfoSpoolfile(PrintjobFilename))
End Property

Public Property Get cPrintjobInfo(ByVal PrintjobFilename As String, ByVal PropertyName As String) As String
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 cPrintjobInfo = cGetPrintjobInfoFromPrintjobInfos(PropertyName, ReadInfoSpoolfile(PrintjobFilename), "cPrintjobInfo")
End Property

Public Property Get cCountOfPrintjobs() As Long
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 If Not mForm Is Nothing Then
  cCountOfPrintjobs = mForm.lsv.ListItems.Count
 End If
End Property

Public Property Get cPrintjobFilename(ByVal JobNumber As Long) As String
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 If Not mForm Is Nothing Then
  If JobNumber > 0 And JobNumber <= mForm.lsv.ListItems.Count Then
   cPrintjobFilename = mForm.lsv.ListItems(JobNumber).SubItems(4)
  End If
 End If
End Property

Public Property Get cDefaultPrinter() As String
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 cDefaultPrinter = Printer.DeviceName
End Property

Public Property Let cDefaultPrinter(ByVal Printername As String)
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Property
 End If
 SetDefaultprinterInProg Printername
End Property

Public Property Get cWindowState() As Long
 cWindowState = ProgramWindowState
End Property

Public Property Let cWindowState(ByVal Value As Long) ' Value must be '0 = normal' or '1 = minimized' or '2 = maximized'
 If Value = vbNormal Or Value = vbMinimized Or Value = vbMaximized Then
  ProgramWindowState = Value
  If Not mForm Is Nothing Then
   If mForm.WindowState = vbMinimized And ProgramWindowState <> vbMinimized Then
    mForm.ShowFrmMain
   End If
   mForm.WindowState = Value
  End If
 End If
End Property

Public Property Get cIsConverted() As Boolean
 cIsConverted = IsConverted
End Property

Public Property Let cIsConverted(ByVal Value As Boolean)
 IsConverted = Value
End Property
' --- Public properties - end ---

' --- Public functions/Subroutines - begin ---
Public Sub cAddPrintjob(ByVal PostscriptFilename As String)
 Dim i As Long, Ext As String, tFilename As String
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 If IsValidPath(PostscriptFilename) = False Then
  mError.Number = 6
  mError.Description = Replace$(ErrDescr6, "%1", PostscriptFilename)
  RaiseEvent eError
  Exit Sub
 End If
 If FileExists(PostscriptFilename) = False Then
  mError.Number = 5
  mError.Description = Replace$(ErrDescr5, "%1", PostscriptFilename)
  RaiseEvent eError
  Exit Sub
 End If
 If IsPostscriptFile(PostscriptFilename) = True Then
   tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
   FileCopy PostscriptFilename, tFilename
  Else
   mError.Number = 9
   mError.Description = Replace$(ErrDescr9, "%1", PostscriptFilename)
   RaiseEvent eError
   Exit Sub
 End If
End Sub

Public Sub cDeletePrintjob(ByVal JobNumber As Long)
 Dim i As Long
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 If Not mForm Is Nothing Then
  If JobNumber > 0 And JobNumber <= mForm.lsv.ListItems.Count Then
   For i = 1 To mForm.lsv.ListItems.Count
    mForm.lsv.ListItems(i).Selected = False
   Next i
   mForm.lsv.ListItems(JobNumber).Selected = True
   mForm.DocumentDelete
  End If
 End If
End Sub

Public Sub cMovePrintjobBottom(ByVal JobNumber As Long)
 Dim i As Long
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 If Not mForm Is Nothing Then
  If JobNumber > 0 And JobNumber < mForm.lsv.ListItems.Count Then
   For i = 1 To mForm.lsv.ListItems.Count
    mForm.lsv.ListItems(i).Selected = False
   Next i
   mForm.lsv.ListItems(JobNumber).Selected = True
   mForm.DocumentBottom
  End If
 End If
End Sub

Public Sub cMovePrintjobTop(ByVal JobNumber As Long)
 Dim i As Long
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 If Not mForm Is Nothing Then
  If JobNumber > 1 And JobNumber <= mForm.lsv.ListItems.Count Then
   For i = 1 To mForm.lsv.ListItems.Count
    mForm.lsv.ListItems(i).Selected = False
   Next i
   mForm.lsv.ListItems(JobNumber).Selected = True
   mForm.DocumentTop
  End If
 End If
End Sub

Public Sub cMovePrintjobUp(ByVal JobNumber As Long)
 Dim i As Long
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 If Not mForm Is Nothing Then
  If JobNumber > 1 And JobNumber <= mForm.lsv.ListItems.Count Then
   For i = 1 To mForm.lsv.ListItems.Count
    mForm.lsv.ListItems(i).Selected = False
   Next i
   mForm.lsv.ListItems(JobNumber).Selected = True
   mForm.DocumentUp
  End If
 End If
End Sub

Public Sub cMovePrintjobDown(ByVal JobNumber As Long)
 Dim i As Long
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 If Not mForm Is Nothing Then
  If JobNumber > 0 And JobNumber < mForm.lsv.ListItems.Count Then
   For i = 1 To mForm.lsv.ListItems.Count
    mForm.lsv.ListItems(i).Selected = False
   Next i
   mForm.lsv.ListItems(JobNumber).Selected = True
   mForm.DocumentDown
  End If
 End If
End Sub

Public Function cClose() As Boolean
 If ProgramIsStarted And InstanceCounter = 1 Then
  cClose = cCloseRunningSession
 End If
End Function

Public Function cCloseRunningSession() As Boolean
 On Error Resume Next
 Dim f As Form
 cVisible = False
 cCloseRunningSession = False
 For Each f In Forms
  If Not f Is Nothing Then
   If UCase$(f.Name) <> "FRMMAIN" Then
    Unload f
    cCloseRunningSession = True
   End If
  End If
 Next f
 If Not mForm Is Nothing Then
  With mForm
   .Timer1 = False
   .Timer2 = False
   .Timer3 = False
  End With
  Unload mForm
  Set mForm = Nothing
  ProgramIsStarted = False
  cCloseRunningSession = True
 End If
End Function

Public Function cStart(Optional ByVal Params As String, Optional ByVal ForceInitialize As Boolean = False) As Boolean
 If cProgramIsRunning And Not ForceInitialize Then
  With mError
   .Number = 2
   .Description = ErrDescr2
  End With
  RaiseEvent eError
  cStart = False
  Exit Function
 End If
 mProgramIsRunning = False
 StartProgram Params
 Set mForm = frmMain
 Set eActionTimer = mForm.Timer3
 eActionTimer.Enabled = True
 cStart = True
End Function

Public Sub cClearCache()
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 ClearCache
End Sub

Public Sub cClearLogfile()
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 ClearLogfile
End Sub

Public Sub cConvertPostscriptfile(ByVal InputFilename As String, ByVal OutputFilename As String)
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 ConvertPostscriptFile InputFilename, OutputFilename
End Sub

Public Sub cTestEvent(ByVal EventName As String)
 Select Case UCase$(EventName)
  Case "ERROR"
   With mError
    .Number = -1
    .Description = ErrDescr_1
   End With
   RaiseEvent eError
  Case "READY"
   RaiseEvent eReady
  Case Else
   With mError
    .Number = -2
    .Description = ErrDescr_2
   End With
   RaiseEvent eError
 End Select
End Sub

Public Sub cShowLogfileDialog(ByVal Value As Boolean)
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 If Value Then
   frmLog.Show
  Else
   If Not frmLog Is Nothing Then
    Unload frmLog
   End If
  End If
End Sub

Public Sub cShowOptionsDialog(ByVal Value As Boolean)
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 If Value Then
   frmOptions.Show
  Else
   If Not frmOptions Is Nothing Then
    Unload frmOptions
   End If
  End If
End Sub

Public Sub cSendMail(ByVal OutputFilename As String, ByVal Recipients As String)
 Dim mail As clsPDFCreatorMail
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 Set mail = New clsPDFCreatorMail
 If mail.Send(OutputFilename, Options.StandardSubject, Options.SendMailMethod, Recipients) <> 0 Then
  MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
 End If
 Set mail = Nothing
End Sub

Public Function cIsPrintable(ByVal Filename As String) As Boolean
 cIsPrintable = IsPrintable(Filename)
End Function

Public Sub cCombineAll()
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 If Not mForm Is Nothing Then
  mForm.CombineAll
 End If
End Sub

Public Function cGetPDFCreatorPrinters() As Collection
 Set cGetPDFCreatorPrinters = GetPDFCreatorPrinters
End Function

Public Function cGetLogfile() As String
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Function
 End If
 cGetLogfile = ReadLogfile
End Function

Public Sub cWriteToLogfile(ByVal LogStr As String)
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 IfLoggingWriteLogfile LogStr
End Sub

Public Sub cPrintFile(ByVal Filename As String)
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 PrintFile Filename
End Sub

Public Sub cPrintPDFCreatorTestpage()
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
  Else
 End If
 PrintTestpage mForm
End Sub

Public Sub cPrintPrinterTestpage(Optional ByVal Printername As String)
 Dim tPrintername As String
 If LenB(Printername) = 0 Then
   tPrintername = cDefaultPrinter
  Else
   tPrintername = Printername
 End If
 ShellAndWait 0, "open", "RUNDLL32.EXE", "PRINTUI.DLL,PrintUIEntry /n """ & tPrintername & """ /k", GetPDFCreatorApplicationPath, wNormal, WCNone
End Sub

Public Function cReadOptions() As clsPDFCreatorOptions
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Function
 End If
 Set cReadOptions = cGetOptions(ReadOptions)
End Function

Public Sub cSaveOptions(Optional ByVal Options1 As Variant) ' Options1 must be a clsPDFCreatorOptions type
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 If Not IsMissing(Options1) Then
   If TypeOf Options1 Is clsPDFCreatorOptions Then
     SaveOptions cLetOptions(Options1)
    Else
     mError.Number = 8
     mError.Description = Replace$(ErrDescr8, "%1", "cSaveOptions")
     RaiseEvent eError
     Exit Sub
   End If
  Else
   SaveOptions Options
 End If
End Sub

Public Function cReadOptionsFromFile(ByVal INIFilename As String) As clsPDFCreatorOptions
 Dim dummyOptions As tOptions
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Function
 End If
 If IsValidPath(INIFilename) = False Then
  mError.Number = 6
  mError.Description = Replace$(ErrDescr6, "%1", INIFilename)
  RaiseEvent eError
  Exit Function
 End If
 If FileExists(INIFilename) = False Then
  mError.Number = 5
  mError.Description = Replace$(ErrDescr5, "%1", INIFilename)
  RaiseEvent eError
  Exit Function
 End If
 Set cReadOptionsFromFile = cGetOptions(ReadOptionsINI(dummyOptions, INIFilename, True, True))
End Function

Public Sub cSaveOptionsToFile(ByVal INIFilename As String, Optional ByVal Options1 As Variant) ' Options1 must be a clsPDFCreatorOptions type
 If CheckForRunningProgramErrors Then
  RaiseEvent eError
  Exit Sub
 End If
 If Not IsMissing(Options1) Then
   If TypeOf Options1 Is clsPDFCreatorOptions Then
     SaveOptionsINI cLetOptions(Options1), INIFilename
    Else
     mError.Number = 8
     mError.Description = Replace$(ErrDescr8, "%1", "cSaveOptionsToFile")
     RaiseEvent eError
     Exit Sub
   End If
  Else
   SaveOptionsINI Options, INIFilename
 End If
End Sub
' --- Public functions/Subroutines - end ---

' --- Private functions/Subroutines - end ---
Private Sub eActionTimer_Timer()
 If ReadyConverting Then
  RaiseEvent eReady
  ReadyConverting = False
 End If
 If GhostscriptError < 0 Then
  mError.Number = 4
  mError.Description = ErrDescr4
  RaiseEvent eError
  GhostscriptError = 0
 End If
 DoEvents
End Sub

Private Function CheckForRunningProgramErrors() As Boolean
 CheckForRunningProgramErrors = False
 If ProgramIsStarted = False Then
  With mError
   .Number = 1
   .Description = ErrDescr1
  End With
  CheckForRunningProgramErrors = True
  Exit Function
 End If
 If mProgramIsRunning Then
  With mError
   .Number = 2
   .Description = ErrDescr2
  End With
  CheckForRunningProgramErrors = True
  Exit Function
 End If
End Function

Private Function cGetPrintjobInfoFromPrintjobInfos(PropertyName As String, sInfo As InfoSpoolFile, FunctionName As String) As Variant
 Select Case UCase$(PropertyName)
  Case "COMPUTER": cGetPrintjobInfoFromPrintjobInfos = sInfo.Computer
  Case "CREATED": cGetPrintjobInfoFromPrintjobInfos = sInfo.Created
  Case "REDMON_DOCNAME": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_DOCNAME
  Case "REDMON_FILENAME": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_FILENAME
  Case "REDMON_JOB": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_JOB
  Case "REDMON_MACHINE": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_MACHINE
  Case "REDMON_PORT": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_PORT
  Case "REDMON_PRINTER": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_PRINTER
  Case "REDMON_SESSIONID": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_SESSIONID
  Case "REDMON_USER": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_USER
  Case "SPOOLERACCOUNT": cGetPrintjobInfoFromPrintjobInfos = sInfo.SpoolerAccount
  Case "SPOOLFILENAME": cGetPrintjobInfoFromPrintjobInfos = sInfo.SpoolFilename
  Case Else:
   mError.Number = 3
   mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in " & FunctionName)
   RaiseEvent eError
 End Select
End Function

Private Function cGetOptionFromOptions(PropertyName As String, Options As tOptions) As Variant
 Select Case UCase$(PropertyName)
  Case "ADDITIONALGHOSTSCRIPTPARAMETERS": cGetOptionFromOptions = Options.AdditionalGhostscriptParameters
  Case "ADDITIONALGHOSTSCRIPTSEARCHPATH": cGetOptionFromOptions = Options.AdditionalGhostscriptSearchpath
  Case "ADDWINDOWSFONTPATH": cGetOptionFromOptions = Options.AddWindowsFontpath
  Case "AUTOSAVEDIRECTORY": cGetOptionFromOptions = Options.AutosaveDirectory
  Case "AUTOSAVEFILENAME": cGetOptionFromOptions = Options.AutosaveFilename
  Case "AUTOSAVEFORMAT": cGetOptionFromOptions = Options.AutosaveFormat
  Case "AUTOSAVESTARTSTANDARDPROGRAM": cGetOptionFromOptions = Options.AutosaveStartStandardProgram
  Case "BITMAPRESOLUTION": cGetOptionFromOptions = Options.BitmapResolution
  Case "BMPCOLORSCOUNT": cGetOptionFromOptions = Options.BMPColorscount
  Case "CLIENTCOMPUTERRESOLVEIPADDRESS": cGetOptionFromOptions = Options.ClientComputerResolveIPAddress
  Case "DEVICEHEIGHTPOINTS": cGetOptionFromOptions = Options.DeviceHeightPoints
  Case "DEVICEWIDTHPOINTS": cGetOptionFromOptions = Options.DeviceWidthPoints
  Case "DIRECTORYGHOSTSCRIPTBINARIES": cGetOptionFromOptions = Options.DirectoryGhostscriptBinaries
  Case "DIRECTORYGHOSTSCRIPTFONTS": cGetOptionFromOptions = Options.DirectoryGhostscriptFonts
  Case "DIRECTORYGHOSTSCRIPTLIBRARIES": cGetOptionFromOptions = Options.DirectoryGhostscriptLibraries
  Case "DIRECTORYGHOSTSCRIPTRESOURCE": cGetOptionFromOptions = Options.DirectoryGhostscriptResource
  Case "DISABLEEMAIL": cGetOptionFromOptions = Options.DisableEmail
  Case "DONTUSEDOCUMENTSETTINGS": cGetOptionFromOptions = Options.DontUseDocumentSettings
  Case "EPSLANGUAGELEVEL": cGetOptionFromOptions = Options.EPSLanguageLevel
  Case "FILENAMESUBSTITUTIONS": cGetOptionFromOptions = Options.FilenameSubstitutions
  Case "FILENAMESUBSTITUTIONSONLYINTITLE": cGetOptionFromOptions = Options.FilenameSubstitutionsOnlyInTitle
  Case "JPEGCOLORSCOUNT": cGetOptionFromOptions = Options.JPEGColorscount
  Case "JPEGQUALITY": cGetOptionFromOptions = Options.JPEGQuality
  Case "LANGUAGE": cGetOptionFromOptions = Options.Language
  Case "LASTSAVEDIRECTORY": cGetOptionFromOptions = Options.LastSaveDirectory
  Case "LOGGING": cGetOptionFromOptions = Options.Logging
  Case "LOGLINES": cGetOptionFromOptions = Options.LogLines
  Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": cGetOptionFromOptions = Options.NoConfirmMessageSwitchingDefaultprinter
  Case "NOPROCESSINGATSTARTUP": cGetOptionFromOptions = Options.NoProcessingAtStartup
  Case "NOPSCHECK": cGetOptionFromOptions = Options.NoPSCheck
  Case "ONEPAGEPERFILE": cGetOptionFromOptions = Options.OnePagePerFile
  Case "OPTIONSDESIGN": cGetOptionFromOptions = Options.OptionsDesign
  Case "OPTIONSENABLED": cGetOptionFromOptions = Options.OptionsEnabled
  Case "OPTIONSVISIBLE": cGetOptionFromOptions = Options.OptionsVisible
  Case "PAPERSIZE": cGetOptionFromOptions = Options.Papersize
  Case "PCXCOLORSCOUNT": cGetOptionFromOptions = Options.PCXColorscount
  Case "PDFALLOWASSEMBLY": cGetOptionFromOptions = Options.PDFAllowAssembly
  Case "PDFALLOWDEGRADEDPRINTING": cGetOptionFromOptions = Options.PDFAllowDegradedPrinting
  Case "PDFALLOWFILLIN": cGetOptionFromOptions = Options.PDFAllowFillIn
  Case "PDFALLOWSCREENREADERS": cGetOptionFromOptions = Options.PDFAllowScreenReaders
  Case "PDFCOLORSCMYKTORGB": cGetOptionFromOptions = Options.PDFColorsCMYKToRGB
  Case "PDFCOLORSCOLORMODEL": cGetOptionFromOptions = Options.PDFColorsColorModel
  Case "PDFCOLORSPRESERVEHALFTONE": cGetOptionFromOptions = Options.PDFColorsPreserveHalftone
  Case "PDFCOLORSPRESERVEOVERPRINT": cGetOptionFromOptions = Options.PDFColorsPreserveOverprint
  Case "PDFCOLORSPRESERVETRANSFER": cGetOptionFromOptions = Options.PDFColorsPreserveTransfer
  Case "PDFCOMPRESSIONCOLORCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionColorCompression
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionColorCompressionChoice
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGHighFactor
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGLowFactor
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMaximumFactor
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMediumFactor
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMinimumFactor
  Case "PDFCOMPRESSIONCOLORRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionColorResample
  Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionColorResampleChoice
  Case "PDFCOMPRESSIONCOLORRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionColorResolution
  Case "PDFCOMPRESSIONGREYCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionGreyCompression
  Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionChoice
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGHighFactor
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGLowFactor
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMaximumFactor
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMediumFactor
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMinimumFactor
  Case "PDFCOMPRESSIONGREYRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionGreyResample
  Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyResampleChoice
  Case "PDFCOMPRESSIONGREYRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionGreyResolution
  Case "PDFCOMPRESSIONMONOCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionMonoCompression
  Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoCompressionChoice
  Case "PDFCOMPRESSIONMONORESAMPLE": cGetOptionFromOptions = Options.PDFCompressionMonoResample
  Case "PDFCOMPRESSIONMONORESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoResampleChoice
  Case "PDFCOMPRESSIONMONORESOLUTION": cGetOptionFromOptions = Options.PDFCompressionMonoResolution
  Case "PDFCOMPRESSIONTEXTCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionTextCompression
  Case "PDFDISALLOWCOPY": cGetOptionFromOptions = Options.PDFDisallowCopy
  Case "PDFDISALLOWMODIFYANNOTATIONS": cGetOptionFromOptions = Options.PDFDisallowModifyAnnotations
  Case "PDFDISALLOWMODIFYCONTENTS": cGetOptionFromOptions = Options.PDFDisallowModifyContents
  Case "PDFDISALLOWPRINTING": cGetOptionFromOptions = Options.PDFDisallowPrinting
  Case "PDFENCRYPTOR": cGetOptionFromOptions = Options.PDFEncryptor
  Case "PDFFONTSEMBEDALL": cGetOptionFromOptions = Options.PDFFontsEmbedAll
  Case "PDFFONTSSUBSETFONTS": cGetOptionFromOptions = Options.PDFFontsSubSetFonts
  Case "PDFFONTSSUBSETFONTSPERCENT": cGetOptionFromOptions = Options.PDFFontsSubSetFontsPercent
  Case "PDFGENERALASCII85": cGetOptionFromOptions = Options.PDFGeneralASCII85
  Case "PDFGENERALAUTOROTATE": cGetOptionFromOptions = Options.PDFGeneralAutorotate
  Case "PDFGENERALCOMPATIBILITY": cGetOptionFromOptions = Options.PDFGeneralCompatibility
  Case "PDFGENERALOVERPRINT": cGetOptionFromOptions = Options.PDFGeneralOverprint
  Case "PDFGENERALRESOLUTION": cGetOptionFromOptions = Options.PDFGeneralResolution
  Case "PDFHIGHENCRYPTION": cGetOptionFromOptions = Options.PDFHighEncryption
  Case "PDFLOWENCRYPTION": cGetOptionFromOptions = Options.PDFLowEncryption
  Case "PDFOPTIMIZE": cGetOptionFromOptions = Options.PDFOptimize
  Case "PDFOWNERPASS": cGetOptionFromOptions = Options.PDFOwnerPass
  Case "PDFOWNERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFOwnerPasswordString
  Case "PDFUSERPASS": cGetOptionFromOptions = Options.PDFUserPass
  Case "PDFUSERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFUserPasswordString
  Case "PDFUSESECURITY": cGetOptionFromOptions = Options.PDFUseSecurity
  Case "PNGCOLORSCOUNT": cGetOptionFromOptions = Options.PNGColorscount
  Case "PRINTAFTERSAVING": cGetOptionFromOptions = Options.PrintAfterSaving
  Case "PRINTAFTERSAVINGDUPLEX": cGetOptionFromOptions = Options.PrintAfterSavingDuplex
  Case "PRINTAFTERSAVINGNOCANCEL": cGetOptionFromOptions = Options.PrintAfterSavingNoCancel
  Case "PRINTAFTERSAVINGPRINTER": cGetOptionFromOptions = Options.PrintAfterSavingPrinter
  Case "PRINTAFTERSAVINGQUERYUSER": cGetOptionFromOptions = Options.PrintAfterSavingQueryUser
  Case "PRINTAFTERSAVINGTUMBLE": cGetOptionFromOptions = Options.PrintAfterSavingTumble
  Case "PRINTERSTOP": cGetOptionFromOptions = Options.PrinterStop
  Case "PRINTERTEMPPATH": cGetOptionFromOptions = Options.PrinterTemppath
  Case "PROCESSPRIORITY": cGetOptionFromOptions = Options.ProcessPriority
  Case "PROGRAMFONT": cGetOptionFromOptions = Options.ProgramFont
  Case "PROGRAMFONTCHARSET": cGetOptionFromOptions = Options.ProgramFontCharset
  Case "PROGRAMFONTSIZE": cGetOptionFromOptions = Options.ProgramFontSize
  Case "PSLANGUAGELEVEL": cGetOptionFromOptions = Options.PSLanguageLevel
  Case "REMOVEALLKNOWNFILEEXTENSIONS": cGetOptionFromOptions = Options.RemoveAllKnownFileExtensions
  Case "REMOVESPACES": cGetOptionFromOptions = Options.RemoveSpaces
  Case "RUNPROGRAMAFTERSAVING": cGetOptionFromOptions = Options.RunProgramAfterSaving
  Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramname
  Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramParameters
  Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": cGetOptionFromOptions = Options.RunProgramAfterSavingWaitUntilReady
  Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramAfterSavingWindowstyle
  Case "RUNPROGRAMBEFORESAVING": cGetOptionFromOptions = Options.RunProgramBeforeSaving
  Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramname
  Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramParameters
  Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramBeforeSavingWindowstyle
  Case "SAVEFILENAME": cGetOptionFromOptions = Options.SaveFilename
  Case "SENDEMAILAFTERAUTOSAVING": cGetOptionFromOptions = Options.SendEmailAfterAutoSaving
  Case "SENDMAILMETHOD": cGetOptionFromOptions = Options.SendMailMethod
  Case "SHOWANIMATION": cGetOptionFromOptions = Options.ShowAnimation
  Case "STAMPFONTCOLOR": cGetOptionFromOptions = Options.StampFontColor
  Case "STAMPFONTNAME": cGetOptionFromOptions = Options.StampFontname
  Case "STAMPFONTSIZE": cGetOptionFromOptions = Options.StampFontsize
  Case "STAMPOUTLINEFONTTHICKNESS": cGetOptionFromOptions = Options.StampOutlineFontthickness
  Case "STAMPSTRING": cGetOptionFromOptions = Options.StampString
  Case "STAMPUSEOUTLINEFONT": cGetOptionFromOptions = Options.StampUseOutlineFont
  Case "STANDARDAUTHOR": cGetOptionFromOptions = Options.StandardAuthor
  Case "STANDARDCREATIONDATE": cGetOptionFromOptions = Options.StandardCreationdate
  Case "STANDARDDATEFORMAT": cGetOptionFromOptions = Options.StandardDateformat
  Case "STANDARDKEYWORDS": cGetOptionFromOptions = Options.StandardKeywords
  Case "STANDARDMAILDOMAIN": cGetOptionFromOptions = Options.StandardMailDomain
  Case "STANDARDMODIFYDATE": cGetOptionFromOptions = Options.StandardModifydate
  Case "STANDARDSAVEFORMAT": cGetOptionFromOptions = Options.StandardSaveformat
  Case "STANDARDSUBJECT": cGetOptionFromOptions = Options.StandardSubject
  Case "STANDARDTITLE": cGetOptionFromOptions = Options.StandardTitle
  Case "STARTSTANDARDPROGRAM": cGetOptionFromOptions = Options.StartStandardProgram
  Case "TIFFCOLORSCOUNT": cGetOptionFromOptions = Options.TIFFColorscount
  Case "TOOLBARS": cGetOptionFromOptions = Options.Toolbars
  Case "USEAUTOSAVE": cGetOptionFromOptions = Options.UseAutosave
  Case "USEAUTOSAVEDIRECTORY": cGetOptionFromOptions = Options.UseAutosaveDirectory
  Case "USECREATIONDATENOW": cGetOptionFromOptions = Options.UseCreationDateNow
  Case "USECUSTOMPAPERSIZE": cGetOptionFromOptions = Options.UseCustomPaperSize
  Case "USEFIXPAPERSIZE": cGetOptionFromOptions = Options.UseFixPapersize
  Case "USESTANDARDAUTHOR": cGetOptionFromOptions = Options.UseStandardAuthor
 End Select
End Function

Private Function cGetInfoSpoolfile(sInfos As InfoSpoolFile) As clsPDFCreatorInfoSpoolFile
 With cGetInfoSpoolfile
  .Computer = sInfos.Computer
  .Created = sInfos.Created
  .REDMON_DOCNAME = sInfos.REDMON_DOCNAME
  .REDMON_FILENAME = sInfos.REDMON_FILENAME
  .REDMON_JOB = sInfos.REDMON_JOB
  .REDMON_MACHINE = sInfos.REDMON_MACHINE
  .REDMON_PORT = sInfos.REDMON_PORT
  .REDMON_PRINTER = sInfos.REDMON_PRINTER
  .REDMON_SESSIONID = sInfos.REDMON_SESSIONID
  .REDMON_USER = sInfos.REDMON_USER
  .SpoolerAccount = sInfos.SpoolerAccount
  .SpoolFilename = sInfos.SpoolFilename
 End With
End Function

Private Function cGetOptions(sOptions As tOptions) As clsPDFCreatorOptions
 Set cGetOptions = New clsPDFCreatorOptions
 With cGetOptions
  .AdditionalGhostscriptParameters = sOptions.AdditionalGhostscriptParameters
  .AdditionalGhostscriptSearchpath = sOptions.AdditionalGhostscriptSearchpath
  .AddWindowsFontpath = sOptions.AddWindowsFontpath
  .AutosaveDirectory = sOptions.AutosaveDirectory
  .AutosaveFilename = sOptions.AutosaveFilename
  .AutosaveFormat = sOptions.AutosaveFormat
  .AutosaveStartStandardProgram = sOptions.AutosaveStartStandardProgram
  .BitmapResolution = sOptions.BitmapResolution
  .BMPColorscount = sOptions.BMPColorscount
  .ClientComputerResolveIPAddress = sOptions.ClientComputerResolveIPAddress
  .DeviceHeightPoints = sOptions.DeviceHeightPoints
  .DeviceWidthPoints = sOptions.DeviceWidthPoints
  .DirectoryGhostscriptBinaries = sOptions.DirectoryGhostscriptBinaries
  .DirectoryGhostscriptFonts = sOptions.DirectoryGhostscriptFonts
  .DirectoryGhostscriptLibraries = sOptions.DirectoryGhostscriptLibraries
  .DirectoryGhostscriptResource = sOptions.DirectoryGhostscriptResource
  .DisableEmail = sOptions.DisableEmail
  .DontUseDocumentSettings = sOptions.DontUseDocumentSettings
  .EPSLanguageLevel = sOptions.EPSLanguageLevel
  .FilenameSubstitutions = sOptions.FilenameSubstitutions
  .FilenameSubstitutionsOnlyInTitle = sOptions.FilenameSubstitutionsOnlyInTitle
  .JPEGColorscount = sOptions.JPEGColorscount
  .JPEGQuality = sOptions.JPEGQuality
  .Language = sOptions.Language
  .LastSaveDirectory = sOptions.LastSaveDirectory
  .Logging = sOptions.Logging
  .LogLines = sOptions.LogLines
  .NoConfirmMessageSwitchingDefaultprinter = sOptions.NoConfirmMessageSwitchingDefaultprinter
  .NoProcessingAtStartup = sOptions.NoProcessingAtStartup
  .NoPSCheck = sOptions.NoPSCheck
  .OnePagePerFile = sOptions.OnePagePerFile
  .OptionsDesign = sOptions.OptionsDesign
  .OptionsEnabled = sOptions.OptionsEnabled
  .OptionsVisible = sOptions.OptionsVisible
  .Papersize = sOptions.Papersize
  .PCXColorscount = sOptions.PCXColorscount
  .PDFAllowAssembly = sOptions.PDFAllowAssembly
  .PDFAllowDegradedPrinting = sOptions.PDFAllowDegradedPrinting
  .PDFAllowFillIn = sOptions.PDFAllowFillIn
  .PDFAllowScreenReaders = sOptions.PDFAllowScreenReaders
  .PDFColorsCMYKToRGB = sOptions.PDFColorsCMYKToRGB
  .PDFColorsColorModel = sOptions.PDFColorsColorModel
  .PDFColorsPreserveHalftone = sOptions.PDFColorsPreserveHalftone
  .PDFColorsPreserveOverprint = sOptions.PDFColorsPreserveOverprint
  .PDFColorsPreserveTransfer = sOptions.PDFColorsPreserveTransfer
  .PDFCompressionColorCompression = sOptions.PDFCompressionColorCompression
  .PDFCompressionColorCompressionChoice = sOptions.PDFCompressionColorCompressionChoice
  .PDFCompressionColorCompressionJPEGHighFactor = sOptions.PDFCompressionColorCompressionJPEGHighFactor
  .PDFCompressionColorCompressionJPEGLowFactor = sOptions.PDFCompressionColorCompressionJPEGLowFactor
  .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
  .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
  .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
  .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
  .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
  .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
  .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
  .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
  .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
  .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
  .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
  .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
  .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
  .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
  .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
  .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
  .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
  .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
  .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
  .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
  .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
  .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
  .PDFDisallowCopy = sOptions.PDFDisallowCopy
  .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
  .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
  .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
  .PDFEncryptor = sOptions.PDFEncryptor
  .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
  .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
  .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
  .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
  .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
  .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
  .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
  .PDFGeneralResolution = sOptions.PDFGeneralResolution
  .PDFHighEncryption = sOptions.PDFHighEncryption
  .PDFLowEncryption = sOptions.PDFLowEncryption
  .PDFOptimize = sOptions.PDFOptimize
  .PDFOwnerPass = sOptions.PDFOwnerPass
  .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
  .PDFUserPass = sOptions.PDFUserPass
  .PDFUserPasswordString = sOptions.PDFUserPasswordString
  .PDFUseSecurity = sOptions.PDFUseSecurity
  .PNGColorscount = sOptions.PNGColorscount
  .PrintAfterSaving = sOptions.PrintAfterSaving
  .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
  .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
  .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
  .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
  .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
  .PrinterStop = sOptions.PrinterStop
  .PrinterTemppath = sOptions.PrinterTemppath
  .ProcessPriority = sOptions.ProcessPriority
  .ProgramFont = sOptions.ProgramFont
  .ProgramFontCharset = sOptions.ProgramFontCharset
  .ProgramFontSize = sOptions.ProgramFontSize
  .PSLanguageLevel = sOptions.PSLanguageLevel
  .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
  .RemoveSpaces = sOptions.RemoveSpaces
  .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
  .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
  .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
  .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
  .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
  .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
  .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
  .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
  .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
  .SaveFilename = sOptions.SaveFilename
  .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
  .SendMailMethod = sOptions.SendMailMethod
  .ShowAnimation = sOptions.ShowAnimation
  .StampFontColor = sOptions.StampFontColor
  .StampFontname = sOptions.StampFontname
  .StampFontsize = sOptions.StampFontsize
  .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
  .StampString = sOptions.StampString
  .StampUseOutlineFont = sOptions.StampUseOutlineFont
  .StandardAuthor = sOptions.StandardAuthor
  .StandardCreationdate = sOptions.StandardCreationdate
  .StandardDateformat = sOptions.StandardDateformat
  .StandardKeywords = sOptions.StandardKeywords
  .StandardMailDomain = sOptions.StandardMailDomain
  .StandardModifydate = sOptions.StandardModifydate
  .StandardSaveformat = sOptions.StandardSaveformat
  .StandardSubject = sOptions.StandardSubject
  .StandardTitle = sOptions.StandardTitle
  .StartStandardProgram = sOptions.StartStandardProgram
  .TIFFColorscount = sOptions.TIFFColorscount
  .Toolbars = sOptions.Toolbars
  .UseAutosave = sOptions.UseAutosave
  .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
  .UseCreationDateNow = sOptions.UseCreationDateNow
  .UseCustomPaperSize = sOptions.UseCustomPaperSize
  .UseFixPapersize = sOptions.UseFixPapersize
  .UseStandardAuthor = sOptions.UseStandardAuthor
 End With
End Function

Private Sub cLetOption(PropertyName As String, Value As Variant)
 Select Case UCase$(PropertyName)
  Case "ADDITIONALGHOSTSCRIPTPARAMETERS": Options.AdditionalGhostscriptParameters = Value
  Case "ADDITIONALGHOSTSCRIPTSEARCHPATH": Options.AdditionalGhostscriptSearchpath = Value
  Case "ADDWINDOWSFONTPATH": Options.AddWindowsFontpath = Value
  Case "AUTOSAVEDIRECTORY": Options.AutosaveDirectory = Value
  Case "AUTOSAVEFILENAME": Options.AutosaveFilename = Value
  Case "AUTOSAVEFORMAT": Options.AutosaveFormat = Value
  Case "AUTOSAVESTARTSTANDARDPROGRAM": Options.AutosaveStartStandardProgram = Value
  Case "BITMAPRESOLUTION": Options.BitmapResolution = Value
  Case "BMPCOLORSCOUNT": Options.BMPColorscount = Value
  Case "CLIENTCOMPUTERRESOLVEIPADDRESS": Options.ClientComputerResolveIPAddress = Value
  Case "DEVICEHEIGHTPOINTS": Options.DeviceHeightPoints = Value
  Case "DEVICEWIDTHPOINTS": Options.DeviceWidthPoints = Value
  Case "DIRECTORYGHOSTSCRIPTBINARIES": Options.DirectoryGhostscriptBinaries = Value
  Case "DIRECTORYGHOSTSCRIPTFONTS": Options.DirectoryGhostscriptFonts = Value
  Case "DIRECTORYGHOSTSCRIPTLIBRARIES": Options.DirectoryGhostscriptLibraries = Value
  Case "DIRECTORYGHOSTSCRIPTRESOURCE": Options.DirectoryGhostscriptResource = Value
  Case "DISABLEEMAIL": Options.DisableEmail = Value
  Case "DONTUSEDOCUMENTSETTINGS": Options.DontUseDocumentSettings = Value
  Case "EPSLANGUAGELEVEL": Options.EPSLanguageLevel = Value
  Case "FILENAMESUBSTITUTIONS": Options.FilenameSubstitutions = Value
  Case "FILENAMESUBSTITUTIONSONLYINTITLE": Options.FilenameSubstitutionsOnlyInTitle = Value
  Case "JPEGCOLORSCOUNT": Options.JPEGColorscount = Value
  Case "JPEGQUALITY": Options.JPEGQuality = Value
  Case "LANGUAGE": Options.Language = Value
  Case "LASTSAVEDIRECTORY": Options.LastSaveDirectory = Value
  Case "LOGGING": Options.Logging = Value
  Case "LOGLINES": Options.LogLines = Value
  Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": Options.NoConfirmMessageSwitchingDefaultprinter = Value
  Case "NOPROCESSINGATSTARTUP": Options.NoProcessingAtStartup = Value
  Case "NOPSCHECK": Options.NoPSCheck = Value
  Case "ONEPAGEPERFILE": Options.OnePagePerFile = Value
  Case "OPTIONSDESIGN": Options.OptionsDesign = Value
  Case "OPTIONSENABLED": Options.OptionsEnabled = Value
  Case "OPTIONSVISIBLE": Options.OptionsVisible = Value
  Case "PAPERSIZE": Options.Papersize = Value
  Case "PCXCOLORSCOUNT": Options.PCXColorscount = Value
  Case "PDFALLOWASSEMBLY": Options.PDFAllowAssembly = Value
  Case "PDFALLOWDEGRADEDPRINTING": Options.PDFAllowDegradedPrinting = Value
  Case "PDFALLOWFILLIN": Options.PDFAllowFillIn = Value
  Case "PDFALLOWSCREENREADERS": Options.PDFAllowScreenReaders = Value
  Case "PDFCOLORSCMYKTORGB": Options.PDFColorsCMYKToRGB = Value
  Case "PDFCOLORSCOLORMODEL": Options.PDFColorsColorModel = Value
  Case "PDFCOLORSPRESERVEHALFTONE": Options.PDFColorsPreserveHalftone = Value
  Case "PDFCOLORSPRESERVEOVERPRINT": Options.PDFColorsPreserveOverprint = Value
  Case "PDFCOLORSPRESERVETRANSFER": Options.PDFColorsPreserveTransfer = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSION": Options.PDFCompressionColorCompression = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": Options.PDFCompressionColorCompressionChoice = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": Options.PDFCompressionColorCompressionJPEGHighFactor = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": Options.PDFCompressionColorCompressionJPEGLowFactor = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": Options.PDFCompressionColorCompressionJPEGMaximumFactor = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": Options.PDFCompressionColorCompressionJPEGMediumFactor = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": Options.PDFCompressionColorCompressionJPEGMinimumFactor = Value
  Case "PDFCOMPRESSIONCOLORRESAMPLE": Options.PDFCompressionColorResample = Value
  Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": Options.PDFCompressionColorResampleChoice = Value
  Case "PDFCOMPRESSIONCOLORRESOLUTION": Options.PDFCompressionColorResolution = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSION": Options.PDFCompressionGreyCompression = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": Options.PDFCompressionGreyCompressionChoice = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": Options.PDFCompressionGreyCompressionJPEGHighFactor = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": Options.PDFCompressionGreyCompressionJPEGLowFactor = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMaximumFactor = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMediumFactor = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMinimumFactor = Value
  Case "PDFCOMPRESSIONGREYRESAMPLE": Options.PDFCompressionGreyResample = Value
  Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": Options.PDFCompressionGreyResampleChoice = Value
  Case "PDFCOMPRESSIONGREYRESOLUTION": Options.PDFCompressionGreyResolution = Value
  Case "PDFCOMPRESSIONMONOCOMPRESSION": Options.PDFCompressionMonoCompression = Value
  Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": Options.PDFCompressionMonoCompressionChoice = Value
  Case "PDFCOMPRESSIONMONORESAMPLE": Options.PDFCompressionMonoResample = Value
  Case "PDFCOMPRESSIONMONORESAMPLECHOICE": Options.PDFCompressionMonoResampleChoice = Value
  Case "PDFCOMPRESSIONMONORESOLUTION": Options.PDFCompressionMonoResolution = Value
  Case "PDFCOMPRESSIONTEXTCOMPRESSION": Options.PDFCompressionTextCompression = Value
  Case "PDFDISALLOWCOPY": Options.PDFDisallowCopy = Value
  Case "PDFDISALLOWMODIFYANNOTATIONS": Options.PDFDisallowModifyAnnotations = Value
  Case "PDFDISALLOWMODIFYCONTENTS": Options.PDFDisallowModifyContents = Value
  Case "PDFDISALLOWPRINTING": Options.PDFDisallowPrinting = Value
  Case "PDFENCRYPTOR": Options.PDFEncryptor = Value
  Case "PDFFONTSEMBEDALL": Options.PDFFontsEmbedAll = Value
  Case "PDFFONTSSUBSETFONTS": Options.PDFFontsSubSetFonts = Value
  Case "PDFFONTSSUBSETFONTSPERCENT": Options.PDFFontsSubSetFontsPercent = Value
  Case "PDFGENERALASCII85": Options.PDFGeneralASCII85 = Value
  Case "PDFGENERALAUTOROTATE": Options.PDFGeneralAutorotate = Value
  Case "PDFGENERALCOMPATIBILITY": Options.PDFGeneralCompatibility = Value
  Case "PDFGENERALOVERPRINT": Options.PDFGeneralOverprint = Value
  Case "PDFGENERALRESOLUTION": Options.PDFGeneralResolution = Value
  Case "PDFHIGHENCRYPTION": Options.PDFHighEncryption = Value
  Case "PDFLOWENCRYPTION": Options.PDFLowEncryption = Value
  Case "PDFOPTIMIZE": Options.PDFOptimize = Value
  Case "PDFOWNERPASS": Options.PDFOwnerPass = Value
  Case "PDFOWNERPASSWORDSTRING": Options.PDFOwnerPasswordString = Value
  Case "PDFUSERPASS": Options.PDFUserPass = Value
  Case "PDFUSERPASSWORDSTRING": Options.PDFUserPasswordString = Value
  Case "PDFUSESECURITY": Options.PDFUseSecurity = Value
  Case "PNGCOLORSCOUNT": Options.PNGColorscount = Value
  Case "PRINTAFTERSAVING": Options.PrintAfterSaving = Value
  Case "PRINTAFTERSAVINGDUPLEX": Options.PrintAfterSavingDuplex = Value
  Case "PRINTAFTERSAVINGNOCANCEL": Options.PrintAfterSavingNoCancel = Value
  Case "PRINTAFTERSAVINGPRINTER": Options.PrintAfterSavingPrinter = Value
  Case "PRINTAFTERSAVINGQUERYUSER": Options.PrintAfterSavingQueryUser = Value
  Case "PRINTAFTERSAVINGTUMBLE": Options.PrintAfterSavingTumble = Value
  Case "PRINTERSTOP": Options.PrinterStop = Value
  Case "PRINTERTEMPPATH": Options.PrinterTemppath = Value
  Case "PROCESSPRIORITY": Options.ProcessPriority = Value
  Case "PROGRAMFONT": Options.ProgramFont = Value
  Case "PROGRAMFONTCHARSET": Options.ProgramFontCharset = Value
  Case "PROGRAMFONTSIZE": Options.ProgramFontSize = Value
  Case "PSLANGUAGELEVEL": Options.PSLanguageLevel = Value
  Case "REMOVEALLKNOWNFILEEXTENSIONS": Options.RemoveAllKnownFileExtensions = Value
  Case "REMOVESPACES": Options.RemoveSpaces = Value
  Case "RUNPROGRAMAFTERSAVING": Options.RunProgramAfterSaving = Value
  Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": Options.RunProgramAfterSavingProgramname = Value
  Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": Options.RunProgramAfterSavingProgramParameters = Value
  Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": Options.RunProgramAfterSavingWaitUntilReady = Value
  Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": Options.RunProgramAfterSavingWindowstyle = Value
  Case "RUNPROGRAMBEFORESAVING": Options.RunProgramBeforeSaving = Value
  Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": Options.RunProgramBeforeSavingProgramname = Value
  Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": Options.RunProgramBeforeSavingProgramParameters = Value
  Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": Options.RunProgramBeforeSavingWindowstyle = Value
  Case "SAVEFILENAME": Options.SaveFilename = Value
  Case "SENDEMAILAFTERAUTOSAVING": Options.SendEmailAfterAutoSaving = Value
  Case "SENDMAILMETHOD": Options.SendMailMethod = Value
  Case "SHOWANIMATION": Options.ShowAnimation = Value
  Case "STAMPFONTCOLOR": Options.StampFontColor = Value
  Case "STAMPFONTNAME": Options.StampFontname = Value
  Case "STAMPFONTSIZE": Options.StampFontsize = Value
  Case "STAMPOUTLINEFONTTHICKNESS": Options.StampOutlineFontthickness = Value
  Case "STAMPSTRING": Options.StampString = Value
  Case "STAMPUSEOUTLINEFONT": Options.StampUseOutlineFont = Value
  Case "STANDARDAUTHOR": Options.StandardAuthor = Value
  Case "STANDARDCREATIONDATE": Options.StandardCreationdate = Value
  Case "STANDARDDATEFORMAT": Options.StandardDateformat = Value
  Case "STANDARDKEYWORDS": Options.StandardKeywords = Value
  Case "STANDARDMAILDOMAIN": Options.StandardMailDomain = Value
  Case "STANDARDMODIFYDATE": Options.StandardModifydate = Value
  Case "STANDARDSAVEFORMAT": Options.StandardSaveformat = Value
  Case "STANDARDSUBJECT": Options.StandardSubject = Value
  Case "STANDARDTITLE": Options.StandardTitle = Value
  Case "STARTSTANDARDPROGRAM": Options.StartStandardProgram = Value
  Case "TIFFCOLORSCOUNT": Options.TIFFColorscount = Value
  Case "TOOLBARS": Options.Toolbars = Value
  Case "USEAUTOSAVE": Options.UseAutosave = Value
  Case "USEAUTOSAVEDIRECTORY": Options.UseAutosaveDirectory = Value
  Case "USECREATIONDATENOW": Options.UseCreationDateNow = Value
  Case "USECUSTOMPAPERSIZE": Options.UseCustomPaperSize = Value
  Case "USEFIXPAPERSIZE": Options.UseFixPapersize = Value
  Case "USESTANDARDAUTHOR": Options.UseStandardAuthor = Value
  Case Else:
   mError.Number = 3
   mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in cLetOption")
   RaiseEvent eError
 End Select
End Sub

Private Function cLetOptions(sOptions As Variant) As tOptions
 With cLetOptions
  .AdditionalGhostscriptParameters = sOptions.AdditionalGhostscriptParameters
  .AdditionalGhostscriptSearchpath = sOptions.AdditionalGhostscriptSearchpath
  .AddWindowsFontpath = sOptions.AddWindowsFontpath
  .AutosaveDirectory = sOptions.AutosaveDirectory
  .AutosaveFilename = sOptions.AutosaveFilename
  .AutosaveFormat = sOptions.AutosaveFormat
  .AutosaveStartStandardProgram = sOptions.AutosaveStartStandardProgram
  .BitmapResolution = sOptions.BitmapResolution
  .BMPColorscount = sOptions.BMPColorscount
  .ClientComputerResolveIPAddress = sOptions.ClientComputerResolveIPAddress
  .DeviceHeightPoints = sOptions.DeviceHeightPoints
  .DeviceWidthPoints = sOptions.DeviceWidthPoints
  .DirectoryGhostscriptBinaries = sOptions.DirectoryGhostscriptBinaries
  .DirectoryGhostscriptFonts = sOptions.DirectoryGhostscriptFonts
  .DirectoryGhostscriptLibraries = sOptions.DirectoryGhostscriptLibraries
  .DirectoryGhostscriptResource = sOptions.DirectoryGhostscriptResource
  .DisableEmail = sOptions.DisableEmail
  .DontUseDocumentSettings = sOptions.DontUseDocumentSettings
  .EPSLanguageLevel = sOptions.EPSLanguageLevel
  .FilenameSubstitutions = sOptions.FilenameSubstitutions
  .FilenameSubstitutionsOnlyInTitle = sOptions.FilenameSubstitutionsOnlyInTitle
  .JPEGColorscount = sOptions.JPEGColorscount
  .JPEGQuality = sOptions.JPEGQuality
  .Language = sOptions.Language
  .LastSaveDirectory = sOptions.LastSaveDirectory
  .Logging = sOptions.Logging
  .LogLines = sOptions.LogLines
  .NoConfirmMessageSwitchingDefaultprinter = sOptions.NoConfirmMessageSwitchingDefaultprinter
  .NoProcessingAtStartup = sOptions.NoProcessingAtStartup
  .NoPSCheck = sOptions.NoPSCheck
  .OnePagePerFile = sOptions.OnePagePerFile
  .OptionsDesign = sOptions.OptionsDesign
  .OptionsEnabled = sOptions.OptionsEnabled
  .OptionsVisible = sOptions.OptionsVisible
  .Papersize = sOptions.Papersize
  .PCXColorscount = sOptions.PCXColorscount
  .PDFAllowAssembly = sOptions.PDFAllowAssembly
  .PDFAllowDegradedPrinting = sOptions.PDFAllowDegradedPrinting
  .PDFAllowFillIn = sOptions.PDFAllowFillIn
  .PDFAllowScreenReaders = sOptions.PDFAllowScreenReaders
  .PDFColorsCMYKToRGB = sOptions.PDFColorsCMYKToRGB
  .PDFColorsColorModel = sOptions.PDFColorsColorModel
  .PDFColorsPreserveHalftone = sOptions.PDFColorsPreserveHalftone
  .PDFColorsPreserveOverprint = sOptions.PDFColorsPreserveOverprint
  .PDFColorsPreserveTransfer = sOptions.PDFColorsPreserveTransfer
  .PDFCompressionColorCompression = sOptions.PDFCompressionColorCompression
  .PDFCompressionColorCompressionChoice = sOptions.PDFCompressionColorCompressionChoice
  .PDFCompressionColorCompressionJPEGHighFactor = sOptions.PDFCompressionColorCompressionJPEGHighFactor
  .PDFCompressionColorCompressionJPEGLowFactor = sOptions.PDFCompressionColorCompressionJPEGLowFactor
  .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
  .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
  .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
  .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
  .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
  .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
  .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
  .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
  .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
  .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
  .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
  .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
  .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
  .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
  .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
  .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
  .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
  .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
  .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
  .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
  .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
  .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
  .PDFDisallowCopy = sOptions.PDFDisallowCopy
  .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
  .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
  .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
  .PDFEncryptor = sOptions.PDFEncryptor
  .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
  .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
  .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
  .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
  .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
  .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
  .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
  .PDFGeneralResolution = sOptions.PDFGeneralResolution
  .PDFHighEncryption = sOptions.PDFHighEncryption
  .PDFLowEncryption = sOptions.PDFLowEncryption
  .PDFOptimize = sOptions.PDFOptimize
  .PDFOwnerPass = sOptions.PDFOwnerPass
  .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
  .PDFUserPass = sOptions.PDFUserPass
  .PDFUserPasswordString = sOptions.PDFUserPasswordString
  .PDFUseSecurity = sOptions.PDFUseSecurity
  .PNGColorscount = sOptions.PNGColorscount
  .PrintAfterSaving = sOptions.PrintAfterSaving
  .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
  .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
  .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
  .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
  .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
  .PrinterStop = sOptions.PrinterStop
  .PrinterTemppath = sOptions.PrinterTemppath
  .ProcessPriority = sOptions.ProcessPriority
  .ProgramFont = sOptions.ProgramFont
  .ProgramFontCharset = sOptions.ProgramFontCharset
  .ProgramFontSize = sOptions.ProgramFontSize
  .PSLanguageLevel = sOptions.PSLanguageLevel
  .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
  .RemoveSpaces = sOptions.RemoveSpaces
  .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
  .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
  .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
  .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
  .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
  .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
  .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
  .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
  .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
  .SaveFilename = sOptions.SaveFilename
  .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
  .SendMailMethod = sOptions.SendMailMethod
  .ShowAnimation = sOptions.ShowAnimation
  .StampFontColor = sOptions.StampFontColor
  .StampFontname = sOptions.StampFontname
  .StampFontsize = sOptions.StampFontsize
  .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
  .StampString = sOptions.StampString
  .StampUseOutlineFont = sOptions.StampUseOutlineFont
  .StandardAuthor = sOptions.StandardAuthor
  .StandardCreationdate = sOptions.StandardCreationdate
  .StandardDateformat = sOptions.StandardDateformat
  .StandardKeywords = sOptions.StandardKeywords
  .StandardMailDomain = sOptions.StandardMailDomain
  .StandardModifydate = sOptions.StandardModifydate
  .StandardSaveformat = sOptions.StandardSaveformat
  .StandardSubject = sOptions.StandardSubject
  .StandardTitle = sOptions.StandardTitle
  .StartStandardProgram = sOptions.StartStandardProgram
  .TIFFColorscount = sOptions.TIFFColorscount
  .Toolbars = sOptions.Toolbars
  .UseAutosave = sOptions.UseAutosave
  .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
  .UseCreationDateNow = sOptions.UseCreationDateNow
  .UseCustomPaperSize = sOptions.UseCustomPaperSize
  .UseFixPapersize = sOptions.UseFixPapersize
  .UseStandardAuthor = sOptions.UseStandardAuthor
 End With
End Function
' --- Private functions/Subroutines - end ---

