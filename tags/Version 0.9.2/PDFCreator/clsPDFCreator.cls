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

Private Const ErrDescr_2 = "Unknown event!"
Private Const ErrDescr_1 = "Test error event."
Private Const ErrDescr1 = "The activex-server is not started! Please use the function cStart() to start the activex-server!"
Private Const ErrDescr2 = "There is a running instance of PDFCreator. Only one instance is allowed!"
Private Const ErrDescr3 = "Unknown option name '%1' in '%2'!"
Private Const ErrDescr4 = "Ghostscript error: '%1'!"
Private Const ErrDescr5 = "File '%1' doesn't exist!"
Private Const ErrDescr6 = "Not a valid path or filename '%1'!"
Private Const ErrDescr7 = "File '%1' is in use!"
Private Const ErrDescr8 = "Function '%1': Options1 is not a 'clsPDFCreatorOptions' type!"

Private mForm As frmMain, mError As clsPDFCreatorError, mProgramIsRunning As Boolean
Private WithEvents eActionTimer As Timer
Attribute eActionTimer.VB_VarHelpID = -1

' --- Class subroutines - begin ---
Private Sub Class_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set mError = New clsPDFCreatorError
50020  InstanceCounter = InstanceCounter + 1
50030  ProgramIsVisible = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "Class_Initialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Class_Terminate()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  InstanceCounter = InstanceCounter - 1
50020  If InstanceCounter = 0 Then
50030   cCloseRunningSession
50040  End If
50050  Set mError = Nothing
50060  Set mForm = Nothing
50070  Set eActionTimer = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "Class_Terminate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
' --- Class subroutines - end ---

' --- Public properties - begin ---
Public Property Get cError() As clsPDFCreatorError
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set cError = mError
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cError [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cErrorDetail(ByVal PropertyName As String) As Variant
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case UCase$(PropertyName)
        Case "DESCRIPTION": cErrorDetail = mError.Description
50030   Case "NUMBER": cErrorDetail = mError.Number
50040   Case Else:
50050    mError.Number = 3
50060    mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in cErrorDetail")
50070    RaiseEvent eError
50080  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cErrorDetail [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Sub cErrorClear()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mError.Number = 0
50020  mError.Description = ""
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cErrorClear")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Property Get cOutputFilename() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  cOutputFilename = ConvertedOutputFilename
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cOutputFilename [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cPDFCreatorApplicationPath() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cPDFCreatorApplicationPath = GetPDFCreatorApplicationPath
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPDFCreatorApplicationPath [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cIsLogfileDialogDisplayed() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim f As Form
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Property
50050  End If
50060  cIsLogfileDialogDisplayed = False
50070  For Each f In Forms
50080   If UCase$(f.Name) = "FRMLOG" Then
50090    cIsLogfileDialogDisplayed = True
50100   End If
50110  Next f
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cIsLogfileDialogDisplayed [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cIsOptionsDialogDisplayed() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim f As Form
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Property
50050  End If
50060  cIsOptionsDialogDisplayed = False
50070  For Each f In Forms
50080   If UCase$(f.Name) = "FRMOPTIONS" Then
50090    cIsOptionsDialogDisplayed = True
50100   End If
50110  Next f
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cIsOptionsDialogDisplayed [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cProgramRelease(Optional ByVal WithBeta As Boolean = True) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cProgramRelease = GetProgramRelease(WithBeta)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cProgramRelease [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cProgramIsRunning() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cProgramIsRunning = ProgramIsRunning(PDFCreator_GUID)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cProgramIsRunning [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cWindowsVersion() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cWindowsVersion = GetWinVersionStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cWindowsVersion [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cVisible() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  cVisible = ProgramIsVisible
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cVisible [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let cVisible(ByVal Value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim f As Form
50020  ProgramIsVisible = Value
50030  If Not mForm Is Nothing Then
50040   For Each f In Forms
50050    If Not f Is Nothing Then
50060     If UCase$(f.Name) <> "FRMSYSTRAY" Then
50070      f.Visible = ProgramIsVisible
50080     End If
50090    End If
50100   Next f
50110   If ProgramIsVisible Then
50120     If mForm.WindowState = vbMinimized Then
50130       FormInTaskbar mForm, False, False, False
50140       mForm.SystrayEnter
50150      Else
50160       mForm.SysTrayLeave
50170       FormInTaskbar mForm, True, False, True
50180     End If
50190    Else
50200     If mForm.WindowState = vbMinimized Then
50210      mForm.SysTrayLeave
50220     End If
50230   End If
50240  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cVisible [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cInstalledAsServer() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cInstalledAsServer = InstalledAsServer
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cInstalledAsServer [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cPrinterStop() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  cPrinterStop = PrinterStop
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPrinterStop [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let cPrinterStop(ByVal Value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  If Not mForm Is Nothing Then
50060   If Value Then
50070     If Not PrinterStop Then
50080      mForm.SetMenuPrinterStop
50090     End If
50100    Else
50110     If PrinterStop Then
50120      mForm.SetMenuPrinterStop
50130     End If
50140   End If
50150  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPrinterStop [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cOptionsNames() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set cOptionsNames = New Collection
50020  With cOptionsNames
50030   .Add "AdditionalGhostscriptParameters"
50040   .Add "AdditionalGhostscriptSearchpath"
50050   .Add "AddWindowsFontpath"
50060   .Add "AutosaveDirectory"
50070   .Add "AutosaveFilename"
50080   .Add "AutosaveFormat"
50090   .Add "AutosaveStartStandardProgram"
50100   .Add "BitmapResolution"
50110   .Add "BMPColorscount"
50120   .Add "ClientComputerResolveIPAddress"
50130   .Add "DeviceHeightPoints"
50140   .Add "DeviceWidthPoints"
50150   .Add "DirectoryGhostscriptBinaries"
50160   .Add "DirectoryGhostscriptFonts"
50170   .Add "DirectoryGhostscriptLibraries"
50180   .Add "DirectoryGhostscriptResource"
50190   .Add "DisableEmail"
50200   .Add "DontUseDocumentSettings"
50210   .Add "EPSLanguageLevel"
50220   .Add "FilenameSubstitutions"
50230   .Add "FilenameSubstitutionsOnlyInTitle"
50240   .Add "JPEGColorscount"
50250   .Add "JPEGQuality"
50260   .Add "Language"
50270   .Add "LastSaveDirectory"
50280   .Add "Logging"
50290   .Add "LogLines"
50300   .Add "NoConfirmMessageSwitchingDefaultprinter"
50310   .Add "NoProcessingAtStartup"
50320   .Add "NoPSCheck"
50330   .Add "OnePagePerFile"
50340   .Add "OptionsDesign"
50350   .Add "OptionsEnabled"
50360   .Add "OptionsVisible"
50370   .Add "Papersize"
50380   .Add "PCXColorscount"
50390   .Add "PDFAllowAssembly"
50400   .Add "PDFAllowDegradedPrinting"
50410   .Add "PDFAllowFillIn"
50420   .Add "PDFAllowScreenReaders"
50430   .Add "PDFColorsCMYKToRGB"
50440   .Add "PDFColorsColorModel"
50450   .Add "PDFColorsPreserveHalftone"
50460   .Add "PDFColorsPreserveOverprint"
50470   .Add "PDFColorsPreserveTransfer"
50480   .Add "PDFCompressionColorCompression"
50490   .Add "PDFCompressionColorCompressionChoice"
50500   .Add "PDFCompressionColorCompressionJPEGHighFactor"
50510   .Add "PDFCompressionColorCompressionJPEGLowFactor"
50520   .Add "PDFCompressionColorCompressionJPEGMaximumFactor"
50530   .Add "PDFCompressionColorCompressionJPEGMediumFactor"
50540   .Add "PDFCompressionColorCompressionJPEGMinimumFactor"
50550   .Add "PDFCompressionColorResample"
50560   .Add "PDFCompressionColorResampleChoice"
50570   .Add "PDFCompressionColorResolution"
50580   .Add "PDFCompressionGreyCompression"
50590   .Add "PDFCompressionGreyCompressionChoice"
50600   .Add "PDFCompressionGreyCompressionJPEGHighFactor"
50610   .Add "PDFCompressionGreyCompressionJPEGLowFactor"
50620   .Add "PDFCompressionGreyCompressionJPEGMaximumFactor"
50630   .Add "PDFCompressionGreyCompressionJPEGMediumFactor"
50640   .Add "PDFCompressionGreyCompressionJPEGMinimumFactor"
50650   .Add "PDFCompressionGreyResample"
50660   .Add "PDFCompressionGreyResampleChoice"
50670   .Add "PDFCompressionGreyResolution"
50680   .Add "PDFCompressionMonoCompression"
50690   .Add "PDFCompressionMonoCompressionChoice"
50700   .Add "PDFCompressionMonoResample"
50710   .Add "PDFCompressionMonoResampleChoice"
50720   .Add "PDFCompressionMonoResolution"
50730   .Add "PDFCompressionTextCompression"
50740   .Add "PDFDisallowCopy"
50750   .Add "PDFDisallowModifyAnnotations"
50760   .Add "PDFDisallowModifyContents"
50770   .Add "PDFDisallowPrinting"
50780   .Add "PDFEncryptor"
50790   .Add "PDFFontsEmbedAll"
50800   .Add "PDFFontsSubSetFonts"
50810   .Add "PDFFontsSubSetFontsPercent"
50820   .Add "PDFGeneralASCII85"
50830   .Add "PDFGeneralAutorotate"
50840   .Add "PDFGeneralCompatibility"
50850   .Add "PDFGeneralOverprint"
50860   .Add "PDFGeneralResolution"
50870   .Add "PDFHighEncryption"
50880   .Add "PDFLowEncryption"
50890   .Add "PDFOptimize"
50900   .Add "PDFOwnerPass"
50910   .Add "PDFOwnerPasswordString"
50920   .Add "PDFUserPass"
50930   .Add "PDFUserPasswordString"
50940   .Add "PDFUseSecurity"
50950   .Add "PNGColorscount"
50960   .Add "PrintAfterSaving"
50970   .Add "PrintAfterSavingDuplex"
50980   .Add "PrintAfterSavingNoCancel"
50990   .Add "PrintAfterSavingPrinter"
51000   .Add "PrintAfterSavingQueryUser"
51010   .Add "PrintAfterSavingTumble"
51020   .Add "PrinterStop"
51030   .Add "PrinterTemppath"
51040   .Add "ProcessPriority"
51050   .Add "ProgramFont"
51060   .Add "ProgramFontCharset"
51070   .Add "ProgramFontSize"
51080   .Add "PSLanguageLevel"
51090   .Add "RemoveAllKnownFileExtensions"
51100   .Add "RemoveSpaces"
51110   .Add "RunProgramAfterSaving"
51120   .Add "RunProgramAfterSavingProgramname"
51130   .Add "RunProgramAfterSavingProgramParameters"
51140   .Add "RunProgramAfterSavingWaitUntilReady"
51150   .Add "RunProgramAfterSavingWindowstyle"
51160   .Add "RunProgramBeforeSaving"
51170   .Add "RunProgramBeforeSavingProgramname"
51180   .Add "RunProgramBeforeSavingProgramParameters"
51190   .Add "RunProgramBeforeSavingWindowstyle"
51200   .Add "SaveFilename"
51210   .Add "SendEmailAfterAutoSaving"
51220   .Add "SendMailMethod"
51230   .Add "ShowAnimation"
51240   .Add "StampFontColor"
51250   .Add "StampFontname"
51260   .Add "StampFontsize"
51270   .Add "StampOutlineFontthickness"
51280   .Add "StampString"
51290   .Add "StampUseOutlineFont"
51300   .Add "StandardAuthor"
51310   .Add "StandardCreationdate"
51320   .Add "StandardDateformat"
51330   .Add "StandardKeywords"
51340   .Add "StandardMailDomain"
51350   .Add "StandardModifydate"
51360   .Add "StandardSaveformat"
51370   .Add "StandardSubject"
51380   .Add "StandardTitle"
51390   .Add "StartStandardProgram"
51400   .Add "TIFFColorscount"
51410   .Add "Toolbars"
51420   .Add "UseAutosave"
51430   .Add "UseAutosaveDirectory"
51440   .Add "UseCreationDateNow"
51450   .Add "UseStandardAuthor"
51460  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cOptionsNames [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cOption(ByVal PropertyName As String) As Variant
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  cOption = cGetOptionFromOptions(PropertyName, Options)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cOption [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let cOption(ByVal PropertyName As String, ByVal Value As Variant)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  cLetOption PropertyName, Value
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cOption [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cOptions() As clsPDFCreatorOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  Set cOptions = cGetOptions(Options)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cOptions [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Set cOptions(ByVal Options1 As clsPDFCreatorOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  Options = cLetOptions(Options1)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cOptions [SET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cStandardOption(ByVal PropertyName As String) As Variant
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cStandardOption = cGetOptionFromOptions(PropertyName, StandardOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cStandardOption [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cStandardOptions() As clsPDFCreatorOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cStandardOptions = cGetOptions(StandardOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cStandardOptions [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cPostscriptInfo(ByVal PostscriptFilename As String, ByVal PropertyName As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PSHeader As tPSHeader
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Property
50050  End If
50060  If FileExists(PostscriptFilename) = False Then
50070   mError.Number = 5
50080   mError.Description = Replace$(ErrDescr5, "%1", PostscriptFilename)
50090   RaiseEvent eError
50100   Exit Property
50110  End If
50120  If FileInUse(PostscriptFilename) Then
50130   mError.Number = 7
50140   mError.Description = Replace$(ErrDescr7, "%1", PostscriptFilename)
50150   RaiseEvent eError
50160   Exit Property
50170  End If
50180  PSHeader = GetPSHeader(PostscriptFilename)
50190  With PSHeader
50201   Select Case True
         Case ("BoundingBox") = PropertyName
50220     cPostscriptInfo = .BoundingBox.Comment
50230    Case ("CreateFor") = PropertyName
50240     cPostscriptInfo = .CreateFor.Comment
50250    Case ("CreationDate") = PropertyName
50260     cPostscriptInfo = .CreationDate.Comment
50270    Case ("Creator") = PropertyName
50280     cPostscriptInfo = .Creator.Comment
50290    Case ("DocumentData") = PropertyName
50300     cPostscriptInfo = .DocumentData.Comment
50310    Case ("DocumentNeededResources") = PropertyName
50320     cPostscriptInfo = .DocumentNeededResources.Comment
50330    Case ("DocumentSuppliedResources") = PropertyName
50340     cPostscriptInfo = .DocumentSuppliedResources.Comment
50350    Case ("LanguageLevel") = PropertyName
50360     cPostscriptInfo = .LanguageLevel.Comment
50370    Case ("Orientation") = PropertyName
50380     cPostscriptInfo = .Orientation.Comment
50390    Case ("PageOrder") = PropertyName
50400     cPostscriptInfo = .PageOrder.Comment
50410    Case ("Pages") = PropertyName
50420     cPostscriptInfo = .Pages.Comment
50430    Case ("TargetDevice") = PropertyName
50440     cPostscriptInfo = .TargetDevice.Comment
50450    Case ("Title") = PropertyName
50460     cPostscriptInfo = .Title.Comment
50470    Case Else
50480     mError.Number = 3
50490     mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in cPostscriptInfo")
50500     RaiseEvent eError
50510   End Select
50520  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPostscriptInfo [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cPrintjobInfos(ByVal PrintjobFilename As String) As clsPDFCreatorInfoSpoolFile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  cPrintjobInfos = cGetInfoSpoolfile(ReadInfoSpoolfile(PrintjobFilename))
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPrintjobInfos [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cPrintjobInfo(ByVal PrintjobFilename As String, ByVal PropertyName As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  cPrintjobInfo = cGetPrintjobInfoFromPrintjobInfos(PropertyName, ReadInfoSpoolfile(PrintjobFilename), "cPrintjobInfo")
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPrintjobInfo [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cCountOfPrintjobs() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  If Not mForm Is Nothing Then
50060   cCountOfPrintjobs = mForm.lsv.ListItems.Count
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cCountOfPrintjobs [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cPrintjobFilename(ByVal JobNumber As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  If Not mForm Is Nothing Then
50060   If JobNumber > 0 And JobNumber <= mForm.lsv.ListItems.Count Then
50070    cPrintjobFilename = mForm.lsv.ListItems(JobNumber).SubItems(4)
50080   End If
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPrintjobFilename [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cDefaultPrinter() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  cDefaultPrinter = Printer.DeviceName
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cDefaultPrinter [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let cDefaultPrinter(ByVal Printername As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  SetDefaultprinterInProg Printername
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cDefaultPrinter [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cWindowState() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cWindowState = ProgramWindowState
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cWindowState [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let cWindowState(ByVal Value As Long) ' Value must be '0 = normal' or '1 = minimized' or '2 = maximized'
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Value = vbNormal Or Value = vbMinimized Or Value = vbMaximized Then
50020   ProgramWindowState = Value
50030   If Not mForm Is Nothing Then
50040    If mForm.WindowState = vbMinimized And ProgramWindowState <> vbMinimized Then
50050     mForm.ShowFrmMain
50060    End If
50070    mForm.WindowState = Value
50080   End If
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cWindowState [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get cIsConverted() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cIsConverted = IsConverted
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cIsConverted [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let cIsConverted(ByVal Value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  IsConverted = Value
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cIsConverted [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
' --- Public properties - end ---

' --- Public functions/Subroutines - begin ---
Public Sub cDeletePrintjob(ByVal JobNumber As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Sub
50050  End If
50060  If Not mForm Is Nothing Then
50070   If JobNumber > 0 And JobNumber <= mForm.lsv.ListItems.Count Then
50080    For i = 1 To mForm.lsv.ListItems.Count
50090     mForm.lsv.ListItems(i).Selected = False
50100    Next i
50110    mForm.lsv.ListItems(JobNumber).Selected = True
50120    mForm.DocumentDelete
50130   End If
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cDeletePrintjob")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub cMovePrintjobBottom(ByVal JobNumber As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Sub
50050  End If
50060  If Not mForm Is Nothing Then
50070   If JobNumber > 0 And JobNumber < mForm.lsv.ListItems.Count Then
50080    For i = 1 To mForm.lsv.ListItems.Count
50090     mForm.lsv.ListItems(i).Selected = False
50100    Next i
50110    mForm.lsv.ListItems(JobNumber).Selected = True
50120    mForm.DocumentBottom
50130   End If
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cMovePrintjobBottom")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub cMovePrintjobTop(ByVal JobNumber As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Sub
50050  End If
50060  If Not mForm Is Nothing Then
50070   If JobNumber > 1 And JobNumber <= mForm.lsv.ListItems.Count Then
50080    For i = 1 To mForm.lsv.ListItems.Count
50090     mForm.lsv.ListItems(i).Selected = False
50100    Next i
50110    mForm.lsv.ListItems(JobNumber).Selected = True
50120    mForm.DocumentTop
50130   End If
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cMovePrintjobTop")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub cMovePrintjobUp(ByVal JobNumber As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Sub
50050  End If
50060  If Not mForm Is Nothing Then
50070   If JobNumber > 1 And JobNumber <= mForm.lsv.ListItems.Count Then
50080    For i = 1 To mForm.lsv.ListItems.Count
50090     mForm.lsv.ListItems(i).Selected = False
50100    Next i
50110    mForm.lsv.ListItems(JobNumber).Selected = True
50120    mForm.DocumentUp
50130   End If
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cMovePrintjobUp")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub cMovePrintjobDown(ByVal JobNumber As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Sub
50050  End If
50060  If Not mForm Is Nothing Then
50070   If JobNumber > 0 And JobNumber < mForm.lsv.ListItems.Count Then
50080    For i = 1 To mForm.lsv.ListItems.Count
50090     mForm.lsv.ListItems(i).Selected = False
50100    Next i
50110    mForm.lsv.ListItems(JobNumber).Selected = True
50120    mForm.DocumentDown
50130   End If
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cMovePrintjobDown")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function cClose() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If ProgramIsStarted And InstanceCounter = 1 Then
50020   cClose = cCloseRunningSession
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cClose")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If cProgramIsRunning And Not ForceInitialize Then
50020   With mError
50030    .Number = 2
50040    .Description = ErrDescr2
50050   End With
50060   RaiseEvent eError
50070   cStart = False
50080   Exit Function
50090  End If
50100  mProgramIsRunning = False
50110  StartProgram Params
50120  Set mForm = frmMain
50130  Set eActionTimer = mForm.Timer3
50140  eActionTimer.Enabled = True
50150  cStart = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cStart")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub cClearCache()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  ClearCache
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cClearCache")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub cClearLogfile()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  ClearLogfile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cClearLogfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub cConvertPostscriptfile(ByVal InputFilename As String, ByVal OutputFilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  ConvertPostscriptFile InputFilename, OutputFilename
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cConvertPostscriptfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub cTestEvent(ByVal EventName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case UCase$(EventName)
        Case "ERROR"
50030    With mError
50040     .Number = -1
50050     .Description = ErrDescr_1
50060    End With
50070    RaiseEvent eError
50080   Case "READY"
50090    RaiseEvent eReady
50100   Case Else
50110    With mError
50120     .Number = -2
50130     .Description = ErrDescr_2
50140    End With
50150    RaiseEvent eError
50160  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cTestEvent")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub cShowLogfileDialog(ByVal Value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  If Value Then
50060    frmLog.Show
50070   Else
50080    If Not frmLog Is Nothing Then
50090     Unload frmLog
50100    End If
50110   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cShowLogfileDialog")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub cShowOptionsDialog(ByVal Value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  If Value Then
50060    frmOptions.Show
50070   Else
50080    If Not frmOptions Is Nothing Then
50090     Unload frmOptions
50100    End If
50110   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cShowOptionsDialog")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub cSendMail(ByVal OutputFilename As String, ByVal Recipients As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim mail As clsPDFCreatorMail
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Sub
50050  End If
50060  Set mail = New clsPDFCreatorMail
50070  If mail.Send(OutputFilename, Options.StandardSubject, Options.SendMailMethod, Recipients) <> 0 Then
50080   MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
50090  End If
50100  Set mail = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cSendMail")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function cIsPrintable(ByVal Filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cIsPrintable = IsPrintable(Filename)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cIsPrintable")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub cCombineAll()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  If Not mForm Is Nothing Then
50060   mForm.CombineAll
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cCombineAll")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function cGetPDFCreatorPrinters() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set cGetPDFCreatorPrinters = GetPDFCreatorPrinters
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cGetPDFCreatorPrinters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function cGetLogfile() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Function
50040  End If
50050  cGetLogfile = ReadLogfile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cGetLogfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub cWriteToLogfile(ByVal LogStr As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  IfLoggingWriteLogfile LogStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cWriteToLogfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub cPrintFile(ByVal Filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  PrintFile Filename
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPrintFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub cPrintPDFCreatorTestpage()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040   Else
50050  End If
50060  PrintTestpage mForm
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPrintPDFCreatorTestpage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub cPrintPrinterTestpage(Optional ByVal Printername As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tPrintername As String
50020  If LenB(Printername) = 0 Then
50030    tPrintername = cDefaultPrinter
50040   Else
50050    tPrintername = Printername
50060  End If
50070  ShellAndWait 0, "open", "RUNDLL32.EXE", "PRINTUI.DLL,PrintUIEntry /n """ & tPrintername & """ /k", GetPDFCreatorApplicationPath, wNormal, WCNone
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPrintPrinterTestpage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function cReadOptions() As clsPDFCreatorOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Function
50040  End If
50050  Set cReadOptions = cGetOptions(ReadOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cReadOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub cSaveOptions(Optional ByVal Options1 As Variant) ' Options1 must be a clsPDFCreatorOptions type
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  If Not IsMissing(Options1) Then
50060    If TypeOf Options1 Is clsPDFCreatorOptions Then
50070      SaveOptions cLetOptions(Options1)
50080     Else
50090      mError.Number = 8
50100      mError.Description = Replace$(ErrDescr8, "%1", "cSaveOptions")
50110      RaiseEvent eError
50120      Exit Sub
50130    End If
50140   Else
50150    SaveOptions Options
50160  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cSaveOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function cReadOptionsFromFile(ByVal INIFilename As String) As clsPDFCreatorOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim dummyOptions As tOptions
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Function
50050  End If
50060  If FileExists(INIFilename) = False Then
50070   mError.Number = 5
50080   mError.Description = Replace$(ErrDescr5, "%1", INIFilename)
50090   RaiseEvent eError
50100   Exit Function
50110  End If
50120  If IsValidPath(INIFilename) = False Then
50130   mError.Number = 6
50140   mError.Description = Replace$(ErrDescr6, "%1", INIFilename)
50150   RaiseEvent eError
50160   Exit Function
50170  End If
50180  Set cReadOptionsFromFile = cGetOptions(ReadOptionsINI(dummyOptions, INIFilename, True, True))
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cReadOptionsFromFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub cSaveOptionsToFile(ByVal INIFilename As String, Optional ByVal Options1 As Variant) ' Options1 must be a clsPDFCreatorOptions type
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  If Not IsMissing(Options1) Then
50060    If TypeOf Options1 Is clsPDFCreatorOptions Then
50070      SaveOptionsINI cLetOptions(Options1), INIFilename
50080     Else
50090      mError.Number = 8
50100      mError.Description = Replace$(ErrDescr8, "%1", "cSaveOptionsToFile")
50110      RaiseEvent eError
50120      Exit Sub
50130    End If
50140   Else
50150    SaveOptionsINI Options, INIFilename
50160  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cSaveOptionsToFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
' --- Public functions/Subroutines - end ---

' --- Private functions/Subroutines - end ---
Private Sub eActionTimer_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If ReadyConverting Then
50020   RaiseEvent eReady
50030   ReadyConverting = False
50040  End If
50050  If GhostscriptError < 0 Then
50060   mError.Number = 4
50070   mError.Description = ErrDescr4
50080   RaiseEvent eError
50090   GhostscriptError = 0
50100  End If
50110  DoEvents
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "eActionTimer_Timer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function CheckForRunningProgramErrors() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  CheckForRunningProgramErrors = False
50020  If ProgramIsStarted = False Then
50030   With mError
50040    .Number = 1
50050    .Description = ErrDescr1
50060   End With
50070   CheckForRunningProgramErrors = True
50080   Exit Function
50090  End If
50100  If mProgramIsRunning Then
50110   With mError
50120    .Number = 2
50130    .Description = ErrDescr2
50140   End With
50150   CheckForRunningProgramErrors = True
50160   Exit Function
50170  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "CheckForRunningProgramErrors")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function cGetPrintjobInfoFromPrintjobInfos(PropertyName As String, sInfo As InfoSpoolFile, FunctionName As String) As Variant
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case UCase$(PropertyName)
        Case "COMPUTER": cGetPrintjobInfoFromPrintjobInfos = sInfo.Computer
50030   Case "CREATED": cGetPrintjobInfoFromPrintjobInfos = sInfo.Created
50040   Case "REDMON_DOCNAME": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_DOCNAME
50050   Case "REDMON_FILENAME": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_FILENAME
50060   Case "REDMON_JOB": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_JOB
50070   Case "REDMON_MACHINE": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_MACHINE
50080   Case "REDMON_PORT": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_PORT
50090   Case "REDMON_PRINTER": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_PRINTER
50100   Case "REDMON_SESSIONID": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_SESSIONID
50110   Case "REDMON_USER": cGetPrintjobInfoFromPrintjobInfos = sInfo.REDMON_USER
50120   Case "SPOOLERACCOUNT": cGetPrintjobInfoFromPrintjobInfos = sInfo.SpoolerAccount
50130   Case "SPOOLFILENAME": cGetPrintjobInfoFromPrintjobInfos = sInfo.SpoolFilename
50140   Case Else:
50150    mError.Number = 3
50160    mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in " & FunctionName)
50170    RaiseEvent eError
50180  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cGetPrintjobInfoFromPrintjobInfos")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function cGetOptionFromOptions(PropertyName As String, Options As tOptions) As Variant
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case UCase$(PropertyName)
        Case "ADDITIONALGHOSTSCRIPTPARAMETERS": cGetOptionFromOptions = Options.AdditionalGhostscriptParameters
50030   Case "ADDITIONALGHOSTSCRIPTSEARCHPATH": cGetOptionFromOptions = Options.AdditionalGhostscriptSearchpath
50040   Case "ADDWINDOWSFONTPATH": cGetOptionFromOptions = Options.AddWindowsFontpath
50050   Case "AUTOSAVEDIRECTORY": cGetOptionFromOptions = Options.AutosaveDirectory
50060   Case "AUTOSAVEFILENAME": cGetOptionFromOptions = Options.AutosaveFilename
50070   Case "AUTOSAVEFORMAT": cGetOptionFromOptions = Options.AutosaveFormat
50080   Case "AUTOSAVESTARTSTANDARDPROGRAM": cGetOptionFromOptions = Options.AutosaveStartStandardProgram
50090   Case "BITMAPRESOLUTION": cGetOptionFromOptions = Options.BitmapResolution
50100   Case "BMPCOLORSCOUNT": cGetOptionFromOptions = Options.BMPColorscount
50110   Case "CLIENTCOMPUTERRESOLVEIPADDRESS": cGetOptionFromOptions = Options.ClientComputerResolveIPAddress
50120   Case "DEVICEHEIGHTPOINTS": cGetOptionFromOptions = Options.DeviceHeightPoints
50130   Case "DEVICEWIDTHPOINTS": cGetOptionFromOptions = Options.DeviceWidthPoints
50140   Case "DIRECTORYGHOSTSCRIPTBINARIES": cGetOptionFromOptions = Options.DirectoryGhostscriptBinaries
50150   Case "DIRECTORYGHOSTSCRIPTFONTS": cGetOptionFromOptions = Options.DirectoryGhostscriptFonts
50160   Case "DIRECTORYGHOSTSCRIPTLIBRARIES": cGetOptionFromOptions = Options.DirectoryGhostscriptLibraries
50170   Case "DIRECTORYGHOSTSCRIPTRESOURCE": cGetOptionFromOptions = Options.DirectoryGhostscriptResource
50180   Case "DISABLEEMAIL": cGetOptionFromOptions = Options.DisableEmail
50190   Case "DONTUSEDOCUMENTSETTINGS": cGetOptionFromOptions = Options.DontUseDocumentSettings
50200   Case "EPSLANGUAGELEVEL": cGetOptionFromOptions = Options.EPSLanguageLevel
50210   Case "FILENAMESUBSTITUTIONS": cGetOptionFromOptions = Options.FilenameSubstitutions
50220   Case "FILENAMESUBSTITUTIONSONLYINTITLE": cGetOptionFromOptions = Options.FilenameSubstitutionsOnlyInTitle
50230   Case "JPEGCOLORSCOUNT": cGetOptionFromOptions = Options.JPEGColorscount
50240   Case "JPEGQUALITY": cGetOptionFromOptions = Options.JPEGQuality
50250   Case "LANGUAGE": cGetOptionFromOptions = Options.Language
50260   Case "LASTSAVEDIRECTORY": cGetOptionFromOptions = Options.LastSaveDirectory
50270   Case "LOGGING": cGetOptionFromOptions = Options.Logging
50280   Case "LOGLINES": cGetOptionFromOptions = Options.LogLines
50290   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": cGetOptionFromOptions = Options.NoConfirmMessageSwitchingDefaultprinter
50300   Case "NOPROCESSINGATSTARTUP": cGetOptionFromOptions = Options.NoProcessingAtStartup
50310   Case "NOPSCHECK": cGetOptionFromOptions = Options.NoPSCheck
50320   Case "ONEPAGEPERFILE": cGetOptionFromOptions = Options.OnePagePerFile
50330   Case "OPTIONSDESIGN": cGetOptionFromOptions = Options.OptionsDesign
50340   Case "OPTIONSENABLED": cGetOptionFromOptions = Options.OptionsEnabled
50350   Case "OPTIONSVISIBLE": cGetOptionFromOptions = Options.OptionsVisible
50360   Case "PAPERSIZE": cGetOptionFromOptions = Options.Papersize
50370   Case "PCXCOLORSCOUNT": cGetOptionFromOptions = Options.PCXColorscount
50380   Case "PDFALLOWASSEMBLY": cGetOptionFromOptions = Options.PDFAllowAssembly
50390   Case "PDFALLOWDEGRADEDPRINTING": cGetOptionFromOptions = Options.PDFAllowDegradedPrinting
50400   Case "PDFALLOWFILLIN": cGetOptionFromOptions = Options.PDFAllowFillIn
50410   Case "PDFALLOWSCREENREADERS": cGetOptionFromOptions = Options.PDFAllowScreenReaders
50420   Case "PDFCOLORSCMYKTORGB": cGetOptionFromOptions = Options.PDFColorsCMYKToRGB
50430   Case "PDFCOLORSCOLORMODEL": cGetOptionFromOptions = Options.PDFColorsColorModel
50440   Case "PDFCOLORSPRESERVEHALFTONE": cGetOptionFromOptions = Options.PDFColorsPreserveHalftone
50450   Case "PDFCOLORSPRESERVEOVERPRINT": cGetOptionFromOptions = Options.PDFColorsPreserveOverprint
50460   Case "PDFCOLORSPRESERVETRANSFER": cGetOptionFromOptions = Options.PDFColorsPreserveTransfer
50470   Case "PDFCOMPRESSIONCOLORCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionColorCompression
50480   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionColorCompressionChoice
50490   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGHighFactor
50500   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGLowFactor
50510   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMaximumFactor
50520   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMediumFactor
50530   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMinimumFactor
50540   Case "PDFCOMPRESSIONCOLORRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionColorResample
50550   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionColorResampleChoice
50560   Case "PDFCOMPRESSIONCOLORRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionColorResolution
50570   Case "PDFCOMPRESSIONGREYCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionGreyCompression
50580   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionChoice
50590   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGHighFactor
50600   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGLowFactor
50610   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMaximumFactor
50620   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMediumFactor
50630   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMinimumFactor
50640   Case "PDFCOMPRESSIONGREYRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionGreyResample
50650   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyResampleChoice
50660   Case "PDFCOMPRESSIONGREYRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionGreyResolution
50670   Case "PDFCOMPRESSIONMONOCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionMonoCompression
50680   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoCompressionChoice
50690   Case "PDFCOMPRESSIONMONORESAMPLE": cGetOptionFromOptions = Options.PDFCompressionMonoResample
50700   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoResampleChoice
50710   Case "PDFCOMPRESSIONMONORESOLUTION": cGetOptionFromOptions = Options.PDFCompressionMonoResolution
50720   Case "PDFCOMPRESSIONTEXTCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionTextCompression
50730   Case "PDFDISALLOWCOPY": cGetOptionFromOptions = Options.PDFDisallowCopy
50740   Case "PDFDISALLOWMODIFYANNOTATIONS": cGetOptionFromOptions = Options.PDFDisallowModifyAnnotations
50750   Case "PDFDISALLOWMODIFYCONTENTS": cGetOptionFromOptions = Options.PDFDisallowModifyContents
50760   Case "PDFDISALLOWPRINTING": cGetOptionFromOptions = Options.PDFDisallowPrinting
50770   Case "PDFENCRYPTOR": cGetOptionFromOptions = Options.PDFEncryptor
50780   Case "PDFFONTSEMBEDALL": cGetOptionFromOptions = Options.PDFFontsEmbedAll
50790   Case "PDFFONTSSUBSETFONTS": cGetOptionFromOptions = Options.PDFFontsSubSetFonts
50800   Case "PDFFONTSSUBSETFONTSPERCENT": cGetOptionFromOptions = Options.PDFFontsSubSetFontsPercent
50810   Case "PDFGENERALASCII85": cGetOptionFromOptions = Options.PDFGeneralASCII85
50820   Case "PDFGENERALAUTOROTATE": cGetOptionFromOptions = Options.PDFGeneralAutorotate
50830   Case "PDFGENERALCOMPATIBILITY": cGetOptionFromOptions = Options.PDFGeneralCompatibility
50840   Case "PDFGENERALOVERPRINT": cGetOptionFromOptions = Options.PDFGeneralOverprint
50850   Case "PDFGENERALRESOLUTION": cGetOptionFromOptions = Options.PDFGeneralResolution
50860   Case "PDFHIGHENCRYPTION": cGetOptionFromOptions = Options.PDFHighEncryption
50870   Case "PDFLOWENCRYPTION": cGetOptionFromOptions = Options.PDFLowEncryption
50880   Case "PDFOPTIMIZE": cGetOptionFromOptions = Options.PDFOptimize
50890   Case "PDFOWNERPASS": cGetOptionFromOptions = Options.PDFOwnerPass
50900   Case "PDFOWNERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFOwnerPasswordString
50910   Case "PDFUSERPASS": cGetOptionFromOptions = Options.PDFUserPass
50920   Case "PDFUSERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFUserPasswordString
50930   Case "PDFUSESECURITY": cGetOptionFromOptions = Options.PDFUseSecurity
50940   Case "PNGCOLORSCOUNT": cGetOptionFromOptions = Options.PNGColorscount
50950   Case "PRINTAFTERSAVING": cGetOptionFromOptions = Options.PrintAfterSaving
50960   Case "PRINTAFTERSAVINGDUPLEX": cGetOptionFromOptions = Options.PrintAfterSavingDuplex
50970   Case "PRINTAFTERSAVINGNOCANCEL": cGetOptionFromOptions = Options.PrintAfterSavingNoCancel
50980   Case "PRINTAFTERSAVINGPRINTER": cGetOptionFromOptions = Options.PrintAfterSavingPrinter
50990   Case "PRINTAFTERSAVINGQUERYUSER": cGetOptionFromOptions = Options.PrintAfterSavingQueryUser
51000   Case "PRINTAFTERSAVINGTUMBLE": cGetOptionFromOptions = Options.PrintAfterSavingTumble
51010   Case "PRINTERSTOP": cGetOptionFromOptions = Options.PrinterStop
51020   Case "PRINTERTEMPPATH": cGetOptionFromOptions = Options.PrinterTemppath
51030   Case "PROCESSPRIORITY": cGetOptionFromOptions = Options.ProcessPriority
51040   Case "PROGRAMFONT": cGetOptionFromOptions = Options.ProgramFont
51050   Case "PROGRAMFONTCHARSET": cGetOptionFromOptions = Options.ProgramFontCharset
51060   Case "PROGRAMFONTSIZE": cGetOptionFromOptions = Options.ProgramFontSize
51070   Case "PSLANGUAGELEVEL": cGetOptionFromOptions = Options.PSLanguageLevel
51080   Case "REMOVEALLKNOWNFILEEXTENSIONS": cGetOptionFromOptions = Options.RemoveAllKnownFileExtensions
51090   Case "REMOVESPACES": cGetOptionFromOptions = Options.RemoveSpaces
51100   Case "RUNPROGRAMAFTERSAVING": cGetOptionFromOptions = Options.RunProgramAfterSaving
51110   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramname
51120   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramParameters
51130   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": cGetOptionFromOptions = Options.RunProgramAfterSavingWaitUntilReady
51140   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramAfterSavingWindowstyle
51150   Case "RUNPROGRAMBEFORESAVING": cGetOptionFromOptions = Options.RunProgramBeforeSaving
51160   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramname
51170   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramParameters
51180   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramBeforeSavingWindowstyle
51190   Case "SAVEFILENAME": cGetOptionFromOptions = Options.SaveFilename
51200   Case "SENDEMAILAFTERAUTOSAVING": cGetOptionFromOptions = Options.SendEmailAfterAutoSaving
51210   Case "SENDMAILMETHOD": cGetOptionFromOptions = Options.SendMailMethod
51220   Case "SHOWANIMATION": cGetOptionFromOptions = Options.ShowAnimation
51230   Case "STAMPFONTCOLOR": cGetOptionFromOptions = Options.StampFontColor
51240   Case "STAMPFONTNAME": cGetOptionFromOptions = Options.StampFontname
51250   Case "STAMPFONTSIZE": cGetOptionFromOptions = Options.StampFontsize
51260   Case "STAMPOUTLINEFONTTHICKNESS": cGetOptionFromOptions = Options.StampOutlineFontthickness
51270   Case "STAMPSTRING": cGetOptionFromOptions = Options.StampString
51280   Case "STAMPUSEOUTLINEFONT": cGetOptionFromOptions = Options.StampUseOutlineFont
51290   Case "STANDARDAUTHOR": cGetOptionFromOptions = Options.StandardAuthor
51300   Case "STANDARDCREATIONDATE": cGetOptionFromOptions = Options.StandardCreationdate
51310   Case "STANDARDDATEFORMAT": cGetOptionFromOptions = Options.StandardDateformat
51320   Case "STANDARDKEYWORDS": cGetOptionFromOptions = Options.StandardKeywords
51330   Case "STANDARDMAILDOMAIN": cGetOptionFromOptions = Options.StandardMailDomain
51340   Case "STANDARDMODIFYDATE": cGetOptionFromOptions = Options.StandardModifydate
51350   Case "STANDARDSAVEFORMAT": cGetOptionFromOptions = Options.StandardSaveformat
51360   Case "STANDARDSUBJECT": cGetOptionFromOptions = Options.StandardSubject
51370   Case "STANDARDTITLE": cGetOptionFromOptions = Options.StandardTitle
51380   Case "STARTSTANDARDPROGRAM": cGetOptionFromOptions = Options.StartStandardProgram
51390   Case "TIFFCOLORSCOUNT": cGetOptionFromOptions = Options.TIFFColorscount
51400   Case "TOOLBARS": cGetOptionFromOptions = Options.Toolbars
51410   Case "USEAUTOSAVE": cGetOptionFromOptions = Options.UseAutosave
51420   Case "USEAUTOSAVEDIRECTORY": cGetOptionFromOptions = Options.UseAutosaveDirectory
51430   Case "USECREATIONDATENOW": cGetOptionFromOptions = Options.UseCreationDateNow
51440   Case "USESTANDARDAUTHOR": cGetOptionFromOptions = Options.UseStandardAuthor
51450  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cGetOptionFromOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function cGetInfoSpoolfile(sInfos As InfoSpoolFile) As clsPDFCreatorInfoSpoolFile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With cGetInfoSpoolfile
50020   .Computer = sInfos.Computer
50030   .Created = sInfos.Created
50040   .REDMON_DOCNAME = sInfos.REDMON_DOCNAME
50050   .REDMON_FILENAME = sInfos.REDMON_FILENAME
50060   .REDMON_JOB = sInfos.REDMON_JOB
50070   .REDMON_MACHINE = sInfos.REDMON_MACHINE
50080   .REDMON_PORT = sInfos.REDMON_PORT
50090   .REDMON_PRINTER = sInfos.REDMON_PRINTER
50100   .REDMON_SESSIONID = sInfos.REDMON_SESSIONID
50110   .REDMON_USER = sInfos.REDMON_USER
50120   .SpoolerAccount = sInfos.SpoolerAccount
50130   .SpoolFilename = sInfos.SpoolFilename
50140  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cGetInfoSpoolfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function cGetOptions(sOptions As tOptions) As clsPDFCreatorOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set cGetOptions = New clsPDFCreatorOptions
50020  With cGetOptions
50030   .AdditionalGhostscriptParameters = sOptions.AdditionalGhostscriptParameters
50040   .AdditionalGhostscriptSearchpath = sOptions.AdditionalGhostscriptSearchpath
50050   .AddWindowsFontpath = sOptions.AddWindowsFontpath
50060   .AutosaveDirectory = sOptions.AutosaveDirectory
50070   .AutosaveFilename = sOptions.AutosaveFilename
50080   .AutosaveFormat = sOptions.AutosaveFormat
50090   .AutosaveStartStandardProgram = sOptions.AutosaveStartStandardProgram
50100   .BitmapResolution = sOptions.BitmapResolution
50110   .BMPColorscount = sOptions.BMPColorscount
50120   .ClientComputerResolveIPAddress = sOptions.ClientComputerResolveIPAddress
50130   .DeviceHeightPoints = sOptions.DeviceHeightPoints
50140   .DeviceWidthPoints = sOptions.DeviceWidthPoints
50150   .DirectoryGhostscriptBinaries = sOptions.DirectoryGhostscriptBinaries
50160   .DirectoryGhostscriptFonts = sOptions.DirectoryGhostscriptFonts
50170   .DirectoryGhostscriptLibraries = sOptions.DirectoryGhostscriptLibraries
50180   .DirectoryGhostscriptResource = sOptions.DirectoryGhostscriptResource
50190   .DisableEmail = sOptions.DisableEmail
50200   .DontUseDocumentSettings = sOptions.DontUseDocumentSettings
50210   .EPSLanguageLevel = sOptions.EPSLanguageLevel
50220   .FilenameSubstitutions = sOptions.FilenameSubstitutions
50230   .FilenameSubstitutionsOnlyInTitle = sOptions.FilenameSubstitutionsOnlyInTitle
50240   .JPEGColorscount = sOptions.JPEGColorscount
50250   .JPEGQuality = sOptions.JPEGQuality
50260   .Language = sOptions.Language
50270   .LastSaveDirectory = sOptions.LastSaveDirectory
50280   .Logging = sOptions.Logging
50290   .LogLines = sOptions.LogLines
50300   .NoConfirmMessageSwitchingDefaultprinter = sOptions.NoConfirmMessageSwitchingDefaultprinter
50310   .NoProcessingAtStartup = sOptions.NoProcessingAtStartup
50320   .NoPSCheck = sOptions.NoPSCheck
50330   .OnePagePerFile = sOptions.OnePagePerFile
50340   .OptionsDesign = sOptions.OptionsDesign
50350   .OptionsEnabled = sOptions.OptionsEnabled
50360   .OptionsVisible = sOptions.OptionsVisible
50370   .Papersize = sOptions.Papersize
50380   .PCXColorscount = sOptions.PCXColorscount
50390   .PDFAllowAssembly = sOptions.PDFAllowAssembly
50400   .PDFAllowDegradedPrinting = sOptions.PDFAllowDegradedPrinting
50410   .PDFAllowFillIn = sOptions.PDFAllowFillIn
50420   .PDFAllowScreenReaders = sOptions.PDFAllowScreenReaders
50430   .PDFColorsCMYKToRGB = sOptions.PDFColorsCMYKToRGB
50440   .PDFColorsColorModel = sOptions.PDFColorsColorModel
50450   .PDFColorsPreserveHalftone = sOptions.PDFColorsPreserveHalftone
50460   .PDFColorsPreserveOverprint = sOptions.PDFColorsPreserveOverprint
50470   .PDFColorsPreserveTransfer = sOptions.PDFColorsPreserveTransfer
50480   .PDFCompressionColorCompression = sOptions.PDFCompressionColorCompression
50490   .PDFCompressionColorCompressionChoice = sOptions.PDFCompressionColorCompressionChoice
50500   .PDFCompressionColorCompressionJPEGHighFactor = sOptions.PDFCompressionColorCompressionJPEGHighFactor
50510   .PDFCompressionColorCompressionJPEGLowFactor = sOptions.PDFCompressionColorCompressionJPEGLowFactor
50520   .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
50530   .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
50540   .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
50550   .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
50560   .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
50570   .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
50580   .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
50590   .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
50600   .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
50610   .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
50620   .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
50630   .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
50640   .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
50650   .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
50660   .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
50670   .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
50680   .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
50690   .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
50700   .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
50710   .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
50720   .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
50730   .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
50740   .PDFDisallowCopy = sOptions.PDFDisallowCopy
50750   .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
50760   .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
50770   .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
50780   .PDFEncryptor = sOptions.PDFEncryptor
50790   .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
50800   .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
50810   .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
50820   .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
50830   .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
50840   .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
50850   .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
50860   .PDFGeneralResolution = sOptions.PDFGeneralResolution
50870   .PDFHighEncryption = sOptions.PDFHighEncryption
50880   .PDFLowEncryption = sOptions.PDFLowEncryption
50890   .PDFOptimize = sOptions.PDFOptimize
50900   .PDFOwnerPass = sOptions.PDFOwnerPass
50910   .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
50920   .PDFUserPass = sOptions.PDFUserPass
50930   .PDFUserPasswordString = sOptions.PDFUserPasswordString
50940   .PDFUseSecurity = sOptions.PDFUseSecurity
50950   .PNGColorscount = sOptions.PNGColorscount
50960   .PrintAfterSaving = sOptions.PrintAfterSaving
50970   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
50980   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
50990   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
51000   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51010   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51020   .PrinterStop = sOptions.PrinterStop
51030   .PrinterTemppath = sOptions.PrinterTemppath
51040   .ProcessPriority = sOptions.ProcessPriority
51050   .ProgramFont = sOptions.ProgramFont
51060   .ProgramFontCharset = sOptions.ProgramFontCharset
51070   .ProgramFontSize = sOptions.ProgramFontSize
51080   .PSLanguageLevel = sOptions.PSLanguageLevel
51090   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51100   .RemoveSpaces = sOptions.RemoveSpaces
51110   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51120   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51130   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51140   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51150   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51160   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51170   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51180   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51190   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51200   .SaveFilename = sOptions.SaveFilename
51210   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51220   .SendMailMethod = sOptions.SendMailMethod
51230   .ShowAnimation = sOptions.ShowAnimation
51240   .StampFontColor = sOptions.StampFontColor
51250   .StampFontname = sOptions.StampFontname
51260   .StampFontsize = sOptions.StampFontsize
51270   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51280   .StampString = sOptions.StampString
51290   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51300   .StandardAuthor = sOptions.StandardAuthor
51310   .StandardCreationdate = sOptions.StandardCreationdate
51320   .StandardDateformat = sOptions.StandardDateformat
51330   .StandardKeywords = sOptions.StandardKeywords
51340   .StandardMailDomain = sOptions.StandardMailDomain
51350   .StandardModifydate = sOptions.StandardModifydate
51360   .StandardSaveformat = sOptions.StandardSaveformat
51370   .StandardSubject = sOptions.StandardSubject
51380   .StandardTitle = sOptions.StandardTitle
51390   .StartStandardProgram = sOptions.StartStandardProgram
51400   .TIFFColorscount = sOptions.TIFFColorscount
51410   .Toolbars = sOptions.Toolbars
51420   .UseAutosave = sOptions.UseAutosave
51430   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51440   .UseCreationDateNow = sOptions.UseCreationDateNow
51450   .UseStandardAuthor = sOptions.UseStandardAuthor
51460  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cGetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub cLetOption(PropertyName As String, Value As Variant)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case UCase$(PropertyName)
        Case "ADDITIONALGHOSTSCRIPTPARAMETERS": Options.AdditionalGhostscriptParameters = Value
50030   Case "ADDITIONALGHOSTSCRIPTSEARCHPATH": Options.AdditionalGhostscriptSearchpath = Value
50040   Case "ADDWINDOWSFONTPATH": Options.AddWindowsFontpath = Value
50050   Case "AUTOSAVEDIRECTORY": Options.AutosaveDirectory = Value
50060   Case "AUTOSAVEFILENAME": Options.AutosaveFilename = Value
50070   Case "AUTOSAVEFORMAT": Options.AutosaveFormat = Value
50080   Case "AUTOSAVESTARTSTANDARDPROGRAM": Options.AutosaveStartStandardProgram = Value
50090   Case "BITMAPRESOLUTION": Options.BitmapResolution = Value
50100   Case "BMPCOLORSCOUNT": Options.BMPColorscount = Value
50110   Case "CLIENTCOMPUTERRESOLVEIPADDRESS": Options.ClientComputerResolveIPAddress = Value
50120   Case "DEVICEHEIGHTPOINTS": Options.DeviceHeightPoints = Value
50130   Case "DEVICEWIDTHPOINTS": Options.DeviceWidthPoints = Value
50140   Case "DIRECTORYGHOSTSCRIPTBINARIES": Options.DirectoryGhostscriptBinaries = Value
50150   Case "DIRECTORYGHOSTSCRIPTFONTS": Options.DirectoryGhostscriptFonts = Value
50160   Case "DIRECTORYGHOSTSCRIPTLIBRARIES": Options.DirectoryGhostscriptLibraries = Value
50170   Case "DIRECTORYGHOSTSCRIPTRESOURCE": Options.DirectoryGhostscriptResource = Value
50180   Case "DISABLEEMAIL": Options.DisableEmail = Value
50190   Case "DONTUSEDOCUMENTSETTINGS": Options.DontUseDocumentSettings = Value
50200   Case "EPSLANGUAGELEVEL": Options.EPSLanguageLevel = Value
50210   Case "FILENAMESUBSTITUTIONS": Options.FilenameSubstitutions = Value
50220   Case "FILENAMESUBSTITUTIONSONLYINTITLE": Options.FilenameSubstitutionsOnlyInTitle = Value
50230   Case "JPEGCOLORSCOUNT": Options.JPEGColorscount = Value
50240   Case "JPEGQUALITY": Options.JPEGQuality = Value
50250   Case "LANGUAGE": Options.Language = Value
50260   Case "LASTSAVEDIRECTORY": Options.LastSaveDirectory = Value
50270   Case "LOGGING": Options.Logging = Value
50280   Case "LOGLINES": Options.LogLines = Value
50290   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": Options.NoConfirmMessageSwitchingDefaultprinter = Value
50300   Case "NOPROCESSINGATSTARTUP": Options.NoProcessingAtStartup = Value
50310   Case "NOPSCHECK": Options.NoPSCheck = Value
50320   Case "ONEPAGEPERFILE": Options.OnePagePerFile = Value
50330   Case "OPTIONSDESIGN": Options.OptionsDesign = Value
50340   Case "OPTIONSENABLED": Options.OptionsEnabled = Value
50350   Case "OPTIONSVISIBLE": Options.OptionsVisible = Value
50360   Case "PAPERSIZE": Options.Papersize = Value
50370   Case "PCXCOLORSCOUNT": Options.PCXColorscount = Value
50380   Case "PDFALLOWASSEMBLY": Options.PDFAllowAssembly = Value
50390   Case "PDFALLOWDEGRADEDPRINTING": Options.PDFAllowDegradedPrinting = Value
50400   Case "PDFALLOWFILLIN": Options.PDFAllowFillIn = Value
50410   Case "PDFALLOWSCREENREADERS": Options.PDFAllowScreenReaders = Value
50420   Case "PDFCOLORSCMYKTORGB": Options.PDFColorsCMYKToRGB = Value
50430   Case "PDFCOLORSCOLORMODEL": Options.PDFColorsColorModel = Value
50440   Case "PDFCOLORSPRESERVEHALFTONE": Options.PDFColorsPreserveHalftone = Value
50450   Case "PDFCOLORSPRESERVEOVERPRINT": Options.PDFColorsPreserveOverprint = Value
50460   Case "PDFCOLORSPRESERVETRANSFER": Options.PDFColorsPreserveTransfer = Value
50470   Case "PDFCOMPRESSIONCOLORCOMPRESSION": Options.PDFCompressionColorCompression = Value
50480   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": Options.PDFCompressionColorCompressionChoice = Value
50490   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": Options.PDFCompressionColorCompressionJPEGHighFactor = Value
50500   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": Options.PDFCompressionColorCompressionJPEGLowFactor = Value
50510   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": Options.PDFCompressionColorCompressionJPEGMaximumFactor = Value
50520   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": Options.PDFCompressionColorCompressionJPEGMediumFactor = Value
50530   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": Options.PDFCompressionColorCompressionJPEGMinimumFactor = Value
50540   Case "PDFCOMPRESSIONCOLORRESAMPLE": Options.PDFCompressionColorResample = Value
50550   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": Options.PDFCompressionColorResampleChoice = Value
50560   Case "PDFCOMPRESSIONCOLORRESOLUTION": Options.PDFCompressionColorResolution = Value
50570   Case "PDFCOMPRESSIONGREYCOMPRESSION": Options.PDFCompressionGreyCompression = Value
50580   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": Options.PDFCompressionGreyCompressionChoice = Value
50590   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": Options.PDFCompressionGreyCompressionJPEGHighFactor = Value
50600   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": Options.PDFCompressionGreyCompressionJPEGLowFactor = Value
50610   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMaximumFactor = Value
50620   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMediumFactor = Value
50630   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMinimumFactor = Value
50640   Case "PDFCOMPRESSIONGREYRESAMPLE": Options.PDFCompressionGreyResample = Value
50650   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": Options.PDFCompressionGreyResampleChoice = Value
50660   Case "PDFCOMPRESSIONGREYRESOLUTION": Options.PDFCompressionGreyResolution = Value
50670   Case "PDFCOMPRESSIONMONOCOMPRESSION": Options.PDFCompressionMonoCompression = Value
50680   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": Options.PDFCompressionMonoCompressionChoice = Value
50690   Case "PDFCOMPRESSIONMONORESAMPLE": Options.PDFCompressionMonoResample = Value
50700   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": Options.PDFCompressionMonoResampleChoice = Value
50710   Case "PDFCOMPRESSIONMONORESOLUTION": Options.PDFCompressionMonoResolution = Value
50720   Case "PDFCOMPRESSIONTEXTCOMPRESSION": Options.PDFCompressionTextCompression = Value
50730   Case "PDFDISALLOWCOPY": Options.PDFDisallowCopy = Value
50740   Case "PDFDISALLOWMODIFYANNOTATIONS": Options.PDFDisallowModifyAnnotations = Value
50750   Case "PDFDISALLOWMODIFYCONTENTS": Options.PDFDisallowModifyContents = Value
50760   Case "PDFDISALLOWPRINTING": Options.PDFDisallowPrinting = Value
50770   Case "PDFENCRYPTOR": Options.PDFEncryptor = Value
50780   Case "PDFFONTSEMBEDALL": Options.PDFFontsEmbedAll = Value
50790   Case "PDFFONTSSUBSETFONTS": Options.PDFFontsSubSetFonts = Value
50800   Case "PDFFONTSSUBSETFONTSPERCENT": Options.PDFFontsSubSetFontsPercent = Value
50810   Case "PDFGENERALASCII85": Options.PDFGeneralASCII85 = Value
50820   Case "PDFGENERALAUTOROTATE": Options.PDFGeneralAutorotate = Value
50830   Case "PDFGENERALCOMPATIBILITY": Options.PDFGeneralCompatibility = Value
50840   Case "PDFGENERALOVERPRINT": Options.PDFGeneralOverprint = Value
50850   Case "PDFGENERALRESOLUTION": Options.PDFGeneralResolution = Value
50860   Case "PDFHIGHENCRYPTION": Options.PDFHighEncryption = Value
50870   Case "PDFLOWENCRYPTION": Options.PDFLowEncryption = Value
50880   Case "PDFOPTIMIZE": Options.PDFOptimize = Value
50890   Case "PDFOWNERPASS": Options.PDFOwnerPass = Value
50900   Case "PDFOWNERPASSWORDSTRING": Options.PDFOwnerPasswordString = Value
50910   Case "PDFUSERPASS": Options.PDFUserPass = Value
50920   Case "PDFUSERPASSWORDSTRING": Options.PDFUserPasswordString = Value
50930   Case "PDFUSESECURITY": Options.PDFUseSecurity = Value
50940   Case "PNGCOLORSCOUNT": Options.PNGColorscount = Value
50950   Case "PRINTAFTERSAVING": Options.PrintAfterSaving = Value
50960   Case "PRINTAFTERSAVINGDUPLEX": Options.PrintAfterSavingDuplex = Value
50970   Case "PRINTAFTERSAVINGNOCANCEL": Options.PrintAfterSavingNoCancel = Value
50980   Case "PRINTAFTERSAVINGPRINTER": Options.PrintAfterSavingPrinter = Value
50990   Case "PRINTAFTERSAVINGQUERYUSER": Options.PrintAfterSavingQueryUser = Value
51000   Case "PRINTAFTERSAVINGTUMBLE": Options.PrintAfterSavingTumble = Value
51010   Case "PRINTERSTOP": Options.PrinterStop = Value
51020   Case "PRINTERTEMPPATH": Options.PrinterTemppath = Value
51030   Case "PROCESSPRIORITY": Options.ProcessPriority = Value
51040   Case "PROGRAMFONT": Options.ProgramFont = Value
51050   Case "PROGRAMFONTCHARSET": Options.ProgramFontCharset = Value
51060   Case "PROGRAMFONTSIZE": Options.ProgramFontSize = Value
51070   Case "PSLANGUAGELEVEL": Options.PSLanguageLevel = Value
51080   Case "REMOVEALLKNOWNFILEEXTENSIONS": Options.RemoveAllKnownFileExtensions = Value
51090   Case "REMOVESPACES": Options.RemoveSpaces = Value
51100   Case "RUNPROGRAMAFTERSAVING": Options.RunProgramAfterSaving = Value
51110   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": Options.RunProgramAfterSavingProgramname = Value
51120   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": Options.RunProgramAfterSavingProgramParameters = Value
51130   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": Options.RunProgramAfterSavingWaitUntilReady = Value
51140   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": Options.RunProgramAfterSavingWindowstyle = Value
51150   Case "RUNPROGRAMBEFORESAVING": Options.RunProgramBeforeSaving = Value
51160   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": Options.RunProgramBeforeSavingProgramname = Value
51170   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": Options.RunProgramBeforeSavingProgramParameters = Value
51180   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": Options.RunProgramBeforeSavingWindowstyle = Value
51190   Case "SAVEFILENAME": Options.SaveFilename = Value
51200   Case "SENDEMAILAFTERAUTOSAVING": Options.SendEmailAfterAutoSaving = Value
51210   Case "SENDMAILMETHOD": Options.SendMailMethod = Value
51220   Case "SHOWANIMATION": Options.ShowAnimation = Value
51230   Case "STAMPFONTCOLOR": Options.StampFontColor = Value
51240   Case "STAMPFONTNAME": Options.StampFontname = Value
51250   Case "STAMPFONTSIZE": Options.StampFontsize = Value
51260   Case "STAMPOUTLINEFONTTHICKNESS": Options.StampOutlineFontthickness = Value
51270   Case "STAMPSTRING": Options.StampString = Value
51280   Case "STAMPUSEOUTLINEFONT": Options.StampUseOutlineFont = Value
51290   Case "STANDARDAUTHOR": Options.StandardAuthor = Value
51300   Case "STANDARDCREATIONDATE": Options.StandardCreationdate = Value
51310   Case "STANDARDDATEFORMAT": Options.StandardDateformat = Value
51320   Case "STANDARDKEYWORDS": Options.StandardKeywords = Value
51330   Case "STANDARDMAILDOMAIN": Options.StandardMailDomain = Value
51340   Case "STANDARDMODIFYDATE": Options.StandardModifydate = Value
51350   Case "STANDARDSAVEFORMAT": Options.StandardSaveformat = Value
51360   Case "STANDARDSUBJECT": Options.StandardSubject = Value
51370   Case "STANDARDTITLE": Options.StandardTitle = Value
51380   Case "STARTSTANDARDPROGRAM": Options.StartStandardProgram = Value
51390   Case "TIFFCOLORSCOUNT": Options.TIFFColorscount = Value
51400   Case "TOOLBARS": Options.Toolbars = Value
51410   Case "USEAUTOSAVE": Options.UseAutosave = Value
51420   Case "USEAUTOSAVEDIRECTORY": Options.UseAutosaveDirectory = Value
51430   Case "USECREATIONDATENOW": Options.UseCreationDateNow = Value
51440   Case "USESTANDARDAUTHOR": Options.UseStandardAuthor = Value
51450   Case Else:
51460    mError.Number = 3
51470    mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in cLetOption")
51480    RaiseEvent eError
51490  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cLetOption")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function cLetOptions(sOptions As Variant) As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With cLetOptions
50020   .AdditionalGhostscriptParameters = sOptions.AdditionalGhostscriptParameters
50030   .AdditionalGhostscriptSearchpath = sOptions.AdditionalGhostscriptSearchpath
50040   .AddWindowsFontpath = sOptions.AddWindowsFontpath
50050   .AutosaveDirectory = sOptions.AutosaveDirectory
50060   .AutosaveFilename = sOptions.AutosaveFilename
50070   .AutosaveFormat = sOptions.AutosaveFormat
50080   .AutosaveStartStandardProgram = sOptions.AutosaveStartStandardProgram
50090   .BitmapResolution = sOptions.BitmapResolution
50100   .BMPColorscount = sOptions.BMPColorscount
50110   .ClientComputerResolveIPAddress = sOptions.ClientComputerResolveIPAddress
50120   .DeviceHeightPoints = sOptions.DeviceHeightPoints
50130   .DeviceWidthPoints = sOptions.DeviceWidthPoints
50140   .DirectoryGhostscriptBinaries = sOptions.DirectoryGhostscriptBinaries
50150   .DirectoryGhostscriptFonts = sOptions.DirectoryGhostscriptFonts
50160   .DirectoryGhostscriptLibraries = sOptions.DirectoryGhostscriptLibraries
50170   .DirectoryGhostscriptResource = sOptions.DirectoryGhostscriptResource
50180   .DisableEmail = sOptions.DisableEmail
50190   .DontUseDocumentSettings = sOptions.DontUseDocumentSettings
50200   .EPSLanguageLevel = sOptions.EPSLanguageLevel
50210   .FilenameSubstitutions = sOptions.FilenameSubstitutions
50220   .FilenameSubstitutionsOnlyInTitle = sOptions.FilenameSubstitutionsOnlyInTitle
50230   .JPEGColorscount = sOptions.JPEGColorscount
50240   .JPEGQuality = sOptions.JPEGQuality
50250   .Language = sOptions.Language
50260   .LastSaveDirectory = sOptions.LastSaveDirectory
50270   .Logging = sOptions.Logging
50280   .LogLines = sOptions.LogLines
50290   .NoConfirmMessageSwitchingDefaultprinter = sOptions.NoConfirmMessageSwitchingDefaultprinter
50300   .NoProcessingAtStartup = sOptions.NoProcessingAtStartup
50310   .NoPSCheck = sOptions.NoPSCheck
50320   .OnePagePerFile = sOptions.OnePagePerFile
50330   .OptionsDesign = sOptions.OptionsDesign
50340   .OptionsEnabled = sOptions.OptionsEnabled
50350   .OptionsVisible = sOptions.OptionsVisible
50360   .Papersize = sOptions.Papersize
50370   .PCXColorscount = sOptions.PCXColorscount
50380   .PDFAllowAssembly = sOptions.PDFAllowAssembly
50390   .PDFAllowDegradedPrinting = sOptions.PDFAllowDegradedPrinting
50400   .PDFAllowFillIn = sOptions.PDFAllowFillIn
50410   .PDFAllowScreenReaders = sOptions.PDFAllowScreenReaders
50420   .PDFColorsCMYKToRGB = sOptions.PDFColorsCMYKToRGB
50430   .PDFColorsColorModel = sOptions.PDFColorsColorModel
50440   .PDFColorsPreserveHalftone = sOptions.PDFColorsPreserveHalftone
50450   .PDFColorsPreserveOverprint = sOptions.PDFColorsPreserveOverprint
50460   .PDFColorsPreserveTransfer = sOptions.PDFColorsPreserveTransfer
50470   .PDFCompressionColorCompression = sOptions.PDFCompressionColorCompression
50480   .PDFCompressionColorCompressionChoice = sOptions.PDFCompressionColorCompressionChoice
50490   .PDFCompressionColorCompressionJPEGHighFactor = sOptions.PDFCompressionColorCompressionJPEGHighFactor
50500   .PDFCompressionColorCompressionJPEGLowFactor = sOptions.PDFCompressionColorCompressionJPEGLowFactor
50510   .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
50520   .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
50530   .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
50540   .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
50550   .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
50560   .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
50570   .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
50580   .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
50590   .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
50600   .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
50610   .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
50620   .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
50630   .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
50640   .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
50650   .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
50660   .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
50670   .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
50680   .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
50690   .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
50700   .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
50710   .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
50720   .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
50730   .PDFDisallowCopy = sOptions.PDFDisallowCopy
50740   .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
50750   .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
50760   .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
50770   .PDFEncryptor = sOptions.PDFEncryptor
50780   .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
50790   .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
50800   .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
50810   .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
50820   .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
50830   .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
50840   .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
50850   .PDFGeneralResolution = sOptions.PDFGeneralResolution
50860   .PDFHighEncryption = sOptions.PDFHighEncryption
50870   .PDFLowEncryption = sOptions.PDFLowEncryption
50880   .PDFOptimize = sOptions.PDFOptimize
50890   .PDFOwnerPass = sOptions.PDFOwnerPass
50900   .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
50910   .PDFUserPass = sOptions.PDFUserPass
50920   .PDFUserPasswordString = sOptions.PDFUserPasswordString
50930   .PDFUseSecurity = sOptions.PDFUseSecurity
50940   .PNGColorscount = sOptions.PNGColorscount
50950   .PrintAfterSaving = sOptions.PrintAfterSaving
50960   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
50970   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
50980   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
50990   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51000   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51010   .PrinterStop = sOptions.PrinterStop
51020   .PrinterTemppath = sOptions.PrinterTemppath
51030   .ProcessPriority = sOptions.ProcessPriority
51040   .ProgramFont = sOptions.ProgramFont
51050   .ProgramFontCharset = sOptions.ProgramFontCharset
51060   .ProgramFontSize = sOptions.ProgramFontSize
51070   .PSLanguageLevel = sOptions.PSLanguageLevel
51080   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51090   .RemoveSpaces = sOptions.RemoveSpaces
51100   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51110   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51120   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51130   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51140   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51150   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51160   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51170   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51180   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51190   .SaveFilename = sOptions.SaveFilename
51200   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51210   .SendMailMethod = sOptions.SendMailMethod
51220   .ShowAnimation = sOptions.ShowAnimation
51230   .StampFontColor = sOptions.StampFontColor
51240   .StampFontname = sOptions.StampFontname
51250   .StampFontsize = sOptions.StampFontsize
51260   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51270   .StampString = sOptions.StampString
51280   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51290   .StandardAuthor = sOptions.StandardAuthor
51300   .StandardCreationdate = sOptions.StandardCreationdate
51310   .StandardDateformat = sOptions.StandardDateformat
51320   .StandardKeywords = sOptions.StandardKeywords
51330   .StandardMailDomain = sOptions.StandardMailDomain
51340   .StandardModifydate = sOptions.StandardModifydate
51350   .StandardSaveformat = sOptions.StandardSaveformat
51360   .StandardSubject = sOptions.StandardSubject
51370   .StandardTitle = sOptions.StandardTitle
51380   .StartStandardProgram = sOptions.StartStandardProgram
51390   .TIFFColorscount = sOptions.TIFFColorscount
51400   .Toolbars = sOptions.Toolbars
51410   .UseAutosave = sOptions.UseAutosave
51420   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51430   .UseCreationDateNow = sOptions.UseCreationDateNow
51440   .UseStandardAuthor = sOptions.UseStandardAuthor
51450  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cLetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
' --- Private functions/Subroutines - end ---

