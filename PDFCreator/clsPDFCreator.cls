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
Private Const ErrDescr10 = "Empty array 'Arguments'!"

Private mForm As frmMain, mError As clsPDFCreatorError, mProgramIsRunning As Boolean
Private WithEvents eActionTimer As Timer
Attribute eActionTimer.VB_VarHelpID = -1
Private In_eActionTimer As Boolean

' --- Class subroutines - begin ---
Private Sub Class_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set mError = New clsPDFCreatorError
50020  InstanceCounter = InstanceCounter + 1
50030  If InstanceCounter = 1 Then
50040   ProgramIsVisible = True
50050  End If
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

Public Property Get cGhostscriptVersion() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim G As tGhostscriptRevision
50020  G = GetGhostscriptRevision
50030  cGhostscriptVersion = G.strProduct & vbCrLf & _
  G.strCopyright & vbCrLf & _
  CStr(G.intRevision) & vbCrLf & _
  CStr(G.intRevisionDate)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cGhostscriptVersion [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

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

Public Property Let cVisible(ByVal value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim f As Form
50020  ProgramIsVisible = value
50030
50040  If cProgramIsRunning Then
50050   Set mForm = frmMain
50060  End If
50070
50080  If Not mForm Is Nothing Then
50090   For Each f In Forms
50100    If Not f Is Nothing Then
50110     If UCase$(f.Name) <> "FRMSYSTRAY" Then
50120      f.Visible = ProgramIsVisible
50130     End If
50140    End If
50150   Next f
50160   If ProgramIsVisible Then
50170     If mForm.WindowState = vbMinimized Then
50180       FormInTaskbar mForm, False, False, False
50190       mForm.SystrayEnter
50200      Else
50210       mForm.SysTrayLeave
50220       FormInTaskbar mForm, True, False, True
50230     End If
50240    Else
50250     If mForm.WindowState = vbMinimized Then
50260      mForm.SysTrayLeave
50270     End If
50280   End If
50290  End If
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

Public Property Let cPrinterStop(ByVal value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  If Not mForm Is Nothing Then
50060   If value Then
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
50130   .Add "Counter"
50140   .Add "DeviceHeightPoints"
50150   .Add "DeviceWidthPoints"
50160   .Add "DirectoryGhostscriptBinaries"
50170   .Add "DirectoryGhostscriptFonts"
50180   .Add "DirectoryGhostscriptLibraries"
50190   .Add "DirectoryGhostscriptResource"
50200   .Add "DisableEmail"
50210   .Add "DontUseDocumentSettings"
50220   .Add "EPSLanguageLevel"
50230   .Add "FilenameSubstitutions"
50240   .Add "FilenameSubstitutionsOnlyInTitle"
50250   .Add "JPEGColorscount"
50260   .Add "JPEGQuality"
50270   .Add "Language"
50280   .Add "LastSaveDirectory"
50290   .Add "Logging"
50300   .Add "LogLines"
50310   .Add "NoConfirmMessageSwitchingDefaultprinter"
50320   .Add "NoProcessingAtStartup"
50330   .Add "NoPSCheck"
50340   .Add "OnePagePerFile"
50350   .Add "OptionsDesign"
50360   .Add "OptionsEnabled"
50370   .Add "OptionsVisible"
50380   .Add "Papersize"
50390   .Add "PCLColorsCount"
50400   .Add "PCXColorscount"
50410   .Add "PDFAllowAssembly"
50420   .Add "PDFAllowDegradedPrinting"
50430   .Add "PDFAllowFillIn"
50440   .Add "PDFAllowScreenReaders"
50450   .Add "PDFColorsCMYKToRGB"
50460   .Add "PDFColorsColorModel"
50470   .Add "PDFColorsPreserveHalftone"
50480   .Add "PDFColorsPreserveOverprint"
50490   .Add "PDFColorsPreserveTransfer"
50500   .Add "PDFCompressionColorCompression"
50510   .Add "PDFCompressionColorCompressionChoice"
50520   .Add "PDFCompressionColorCompressionJPEGHighFactor"
50530   .Add "PDFCompressionColorCompressionJPEGLowFactor"
50540   .Add "PDFCompressionColorCompressionJPEGMaximumFactor"
50550   .Add "PDFCompressionColorCompressionJPEGMediumFactor"
50560   .Add "PDFCompressionColorCompressionJPEGMinimumFactor"
50570   .Add "PDFCompressionColorResample"
50580   .Add "PDFCompressionColorResampleChoice"
50590   .Add "PDFCompressionColorResolution"
50600   .Add "PDFCompressionGreyCompression"
50610   .Add "PDFCompressionGreyCompressionChoice"
50620   .Add "PDFCompressionGreyCompressionJPEGHighFactor"
50630   .Add "PDFCompressionGreyCompressionJPEGLowFactor"
50640   .Add "PDFCompressionGreyCompressionJPEGMaximumFactor"
50650   .Add "PDFCompressionGreyCompressionJPEGMediumFactor"
50660   .Add "PDFCompressionGreyCompressionJPEGMinimumFactor"
50670   .Add "PDFCompressionGreyResample"
50680   .Add "PDFCompressionGreyResampleChoice"
50690   .Add "PDFCompressionGreyResolution"
50700   .Add "PDFCompressionMonoCompression"
50710   .Add "PDFCompressionMonoCompressionChoice"
50720   .Add "PDFCompressionMonoResample"
50730   .Add "PDFCompressionMonoResampleChoice"
50740   .Add "PDFCompressionMonoResolution"
50750   .Add "PDFCompressionTextCompression"
50760   .Add "PDFDisallowCopy"
50770   .Add "PDFDisallowModifyAnnotations"
50780   .Add "PDFDisallowModifyContents"
50790   .Add "PDFDisallowPrinting"
50800   .Add "PDFEncryptor"
50810   .Add "PDFFontsEmbedAll"
50820   .Add "PDFFontsSubSetFonts"
50830   .Add "PDFFontsSubSetFontsPercent"
50840   .Add "PDFGeneralASCII85"
50850   .Add "PDFGeneralAutorotate"
50860   .Add "PDFGeneralCompatibility"
50870   .Add "PDFGeneralDefault"
50880   .Add "PDFGeneralOverprint"
50890   .Add "PDFGeneralResolution"
50900   .Add "PDFHighEncryption"
50910   .Add "PDFLowEncryption"
50920   .Add "PDFOptimize"
50930   .Add "PDFOwnerPass"
50940   .Add "PDFOwnerPasswordString"
50950   .Add "PDFSigningMultiSignature"
50960   .Add "PDFSigningPFXFile"
50970   .Add "PDFSigningPFXFilePassword"
50980   .Add "PDFSigningSignatureContact"
50990   .Add "PDFSigningSignatureLeftX"
51000   .Add "PDFSigningSignatureLeftY"
51010   .Add "PDFSigningSignatureLocation"
51020   .Add "PDFSigningSignatureReason"
51030   .Add "PDFSigningSignatureRightX"
51040   .Add "PDFSigningSignatureRightY"
51050   .Add "PDFSigningSignatureVisible"
51060   .Add "PDFSigningSignPDF"
51070   .Add "PDFUpdateMetadata"
51080   .Add "PDFUserPass"
51090   .Add "PDFUserPasswordString"
51100   .Add "PDFUseSecurity"
51110   .Add "PNGColorscount"
51120   .Add "PrintAfterSaving"
51130   .Add "PrintAfterSavingDuplex"
51140   .Add "PrintAfterSavingNoCancel"
51150   .Add "PrintAfterSavingPrinter"
51160   .Add "PrintAfterSavingQueryUser"
51170   .Add "PrintAfterSavingTumble"
51180   .Add "PrinterStop"
51190   .Add "PrinterTemppath"
51200   .Add "ProcessPriority"
51210   .Add "ProgramFont"
51220   .Add "ProgramFontCharset"
51230   .Add "ProgramFontSize"
51240   .Add "PSDColorsCount"
51250   .Add "PSLanguageLevel"
51260   .Add "RAWColorsCount"
51270   .Add "RemoveAllKnownFileExtensions"
51280   .Add "RemoveSpaces"
51290   .Add "RunProgramAfterSaving"
51300   .Add "RunProgramAfterSavingProgramname"
51310   .Add "RunProgramAfterSavingProgramParameters"
51320   .Add "RunProgramAfterSavingWaitUntilReady"
51330   .Add "RunProgramAfterSavingWindowstyle"
51340   .Add "RunProgramBeforeSaving"
51350   .Add "RunProgramBeforeSavingProgramname"
51360   .Add "RunProgramBeforeSavingProgramParameters"
51370   .Add "RunProgramBeforeSavingWindowstyle"
51380   .Add "SaveFilename"
51390   .Add "SendEmailAfterAutoSaving"
51400   .Add "SendMailMethod"
51410   .Add "ShowAnimation"
51420   .Add "StampFontColor"
51430   .Add "StampFontname"
51440   .Add "StampFontsize"
51450   .Add "StampOutlineFontthickness"
51460   .Add "StampString"
51470   .Add "StampUseOutlineFont"
51480   .Add "StandardAuthor"
51490   .Add "StandardCreationdate"
51500   .Add "StandardDateformat"
51510   .Add "StandardKeywords"
51520   .Add "StandardMailDomain"
51530   .Add "StandardModifydate"
51540   .Add "StandardSaveformat"
51550   .Add "StandardSubject"
51560   .Add "StandardTitle"
51570   .Add "StartStandardProgram"
51580   .Add "TIFFColorscount"
51590   .Add "Toolbars"
51600   .Add "UseAutosave"
51610   .Add "UseAutosaveDirectory"
51620   .Add "UseCreationDateNow"
51630   .Add "UseCustomPaperSize"
51640   .Add "UseFixPapersize"
51650   .Add "UseStandardAuthor"
51660   .Add "XCFColorsCount"
51670  End With
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
50010  cOption = cGetOptionFromOptions(PropertyName, Options)
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

Public Property Let cOption(ByVal PropertyName As String, ByVal value As Variant)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  cLetOption PropertyName, value
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
50010  Set cOptions = cGetOptions(Options)
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
50010  Set cStandardOptions = cGetOptions(StandardOptions)
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
50460     cPostscriptInfo = .title.Comment
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
50050  Set cPrintjobInfos = cGetInfoSpoolfile(ReadInfoSpoolfile(PrintjobFilename))
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

Public Property Get cStopURLPrinting() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  cStopURLPrinting = StopURLPrinting
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cStopURLPrinting [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let cStopURLPrinting(ByVal value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  StopURLPrinting = value
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cStopURLPrinting [LET]")
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

Public Property Let cWindowState(ByVal value As Long) ' Value must be '0 = normal' or '1 = minimized' or '2 = maximized'
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If value = vbNormal Or value = vbMinimized Or value = vbMaximized Then
50020   ProgramWindowState = value
50030   If Not mForm Is Nothing Then
50040    If mForm.WindowState = vbMinimized And ProgramWindowState <> vbMinimized Then
50050     mForm.ShowFrmMain
50060    End If
50070    mForm.WindowState = value
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

Public Property Let cIsConverted(ByVal value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  IsConverted = value
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

Public Property Get cInstanceCounter() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cInstanceCounter = InstanceCounter
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cInstanceCounter [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
' --- Public properties - end ---

' --- Public functions/Subroutines - begin ---
Public Sub cAddPrintjob(ByVal PostscriptFilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, Ext As String, tFilename As String
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Sub
50050  End If
50060  If IsValidPath(PostscriptFilename) = False Then
50070   mError.Number = 6
50080   mError.Description = Replace$(ErrDescr6, "%1", PostscriptFilename)
50090   RaiseEvent eError
50100   Exit Sub
50110  End If
50120  If FileExists(PostscriptFilename) = False Then
50130   mError.Number = 5
50140   mError.Description = Replace$(ErrDescr5, "%1", PostscriptFilename)
50150   RaiseEvent eError
50160   Exit Sub
50170  End If
50180  If IsPostscriptFile(PostscriptFilename) = True Then
50190    tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
50200    FileCopy PostscriptFilename, tFilename
50210   Else
50220    mError.Number = 9
50230    mError.Description = Replace$(ErrDescr9, "%1", PostscriptFilename)
50240    RaiseEvent eError
50250    Exit Sub
50260  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cAddPrintjob")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

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
 On Error Resume Next
 Dim f As Form
 Do
  DoEvents
 Loop Until ProgramIsStarted = True
 cVisible = False
 cClose = False
 Set eActionTimer = Nothing
 Do
  DoEvents
 Loop Until In_eActionTimer = False
 For Each f In Forms
  If Not f Is Nothing Then
   If UCase$(f.Name) <> "FRMMAIN" Then
    Unload f
    cClose = True
   End If
  End If
 Next f
 If Not mForm Is Nothing Then
  With mForm
   .Timer1.Enabled = False
   .Timer2.Enabled = False
   Do
    DoEvents
   Loop Until .InTimer1 = False
   Do
    DoEvents
   Loop Until .InTimer2 = False
  End With
  Unload mForm
  Set mForm = Nothing
  Do
   DoEvents
  Loop Until Forms.Count = 0
  ProgramIsStarted = False
  cClose = True
 End If
 Set mError = Nothing
End Function

Public Function cCloseRunningSession() As Boolean ' This function is deprecated. Use "cClose()"!
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cCloseRunningSession = cClose
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cCloseRunningSession")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50160  cVisible = ProgramIsVisible
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

Public Sub cConvertPostscriptfile(ByVal InputFilename As String, ByVal OutputFilename As String) ' This function is deprecated! Use "cConvertFile()" instead.
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  cConvertFile InputFilename, OutputFilename
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

Public Sub cConvertFile(ByVal InputFilename As String, ByVal OutputFilename As String, Optional ByVal SubFormat As String = "") ' Allowed values for 'SubFormat': "PDF/A-1b" for "PDF/A-1b"files; "PDF/X" for "PDF/X"-files
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  ConvertFile InputFilename, OutputFilename, SubFormat
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cConvertFile")
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

Public Sub cShowLogfileDialog(ByVal value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  If value Then
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

Public Sub cShowOptionsDialog(ByVal value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  If value Then
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

Public Function cIsPrintable(ByVal filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cIsPrintable = IsPrintable(filename)
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

Public Sub cPrintFile(ByVal filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  PrintFile filename
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

Public Sub cPrintURL(ByVal URL As String, Optional ByVal TimeBetweenLoadAndPrint As Long = 0)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  PrintURL URL, TimeBetweenLoadAndPrint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPrintURL")
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
50060  If IsValidPath(INIFilename) = False Then
50070   mError.Number = 6
50080   mError.Description = Replace$(ErrDescr6, "%1", INIFilename)
50090   RaiseEvent eError
50100   Exit Function
50110  End If
50120  If FileExists(INIFilename) = False Then
50130   mError.Number = 5
50140   mError.Description = Replace$(ErrDescr5, "%1", INIFilename)
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

Public Function cGhostscriptRun(Arguments() As Variant) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim gsArgs() As String, i As Long, c As Collection
50020  If Not IsArray(Arguments) Then
50030   mError.Number = 10
50040   mError.Description = ErrDescr10
50050   RaiseEvent eError
50060   Exit Function
50070  End If
50080  Set c = New Collection
50090  For i = LBound(Arguments) To UBound(Arguments)
50100   If LenB(Trim$(CStr(Arguments(i)))) > 0 Then
50110    c.Add Trim$(CStr(Arguments(i)))
50120   End If
50130  Next i
50140  If c.Count = 0 Then
50150   mError.Number = 10
50160   mError.Description = ErrDescr10
50170   RaiseEvent eError
50180   Exit Function
50190  End If
50200  c.Add "", , 1
50210  ReDim gsArgs(c.Count - 1)
50220  For i = 0 To c.Count - 1
50230   gsArgs(i) = c(i + 1)
50240  Next i
50250  cGhostscriptRun = CallGS(gsArgs)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cGhostscriptRun")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' --- Private functions/Subroutines - end ---
Private Sub eActionTimer_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  In_eActionTimer = True
50020  If ReadyConverting Then
50030   RaiseEvent eReady
50040   ReadyConverting = False
50050  End If
50060  If GhostscriptError < 0 Then
50070   mError.Number = 4
50080   mError.Description = ErrDescr4
50090   RaiseEvent eError
50100   GhostscriptError = 0
50110  End If
50120  DoEvents
50130  In_eActionTimer = False
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
50120   Case "COUNTER": cGetOptionFromOptions = Options.Counter
50130   Case "DEVICEHEIGHTPOINTS": cGetOptionFromOptions = Options.DeviceHeightPoints
50140   Case "DEVICEWIDTHPOINTS": cGetOptionFromOptions = Options.DeviceWidthPoints
50150   Case "DIRECTORYGHOSTSCRIPTBINARIES": cGetOptionFromOptions = Options.DirectoryGhostscriptBinaries
50160   Case "DIRECTORYGHOSTSCRIPTFONTS": cGetOptionFromOptions = Options.DirectoryGhostscriptFonts
50170   Case "DIRECTORYGHOSTSCRIPTLIBRARIES": cGetOptionFromOptions = Options.DirectoryGhostscriptLibraries
50180   Case "DIRECTORYGHOSTSCRIPTRESOURCE": cGetOptionFromOptions = Options.DirectoryGhostscriptResource
50190   Case "DISABLEEMAIL": cGetOptionFromOptions = Options.DisableEmail
50200   Case "DONTUSEDOCUMENTSETTINGS": cGetOptionFromOptions = Options.DontUseDocumentSettings
50210   Case "EPSLANGUAGELEVEL": cGetOptionFromOptions = Options.EPSLanguageLevel
50220   Case "FILENAMESUBSTITUTIONS": cGetOptionFromOptions = Options.FilenameSubstitutions
50230   Case "FILENAMESUBSTITUTIONSONLYINTITLE": cGetOptionFromOptions = Options.FilenameSubstitutionsOnlyInTitle
50240   Case "JPEGCOLORSCOUNT": cGetOptionFromOptions = Options.JPEGColorscount
50250   Case "JPEGQUALITY": cGetOptionFromOptions = Options.JPEGQuality
50260   Case "LANGUAGE": cGetOptionFromOptions = Options.Language
50270   Case "LASTSAVEDIRECTORY": cGetOptionFromOptions = Options.LastSaveDirectory
50280   Case "LOGGING": cGetOptionFromOptions = Options.Logging
50290   Case "LOGLINES": cGetOptionFromOptions = Options.LogLines
50300   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": cGetOptionFromOptions = Options.NoConfirmMessageSwitchingDefaultprinter
50310   Case "NOPROCESSINGATSTARTUP": cGetOptionFromOptions = Options.NoProcessingAtStartup
50320   Case "NOPSCHECK": cGetOptionFromOptions = Options.NoPSCheck
50330   Case "ONEPAGEPERFILE": cGetOptionFromOptions = Options.OnePagePerFile
50340   Case "OPTIONSDESIGN": cGetOptionFromOptions = Options.OptionsDesign
50350   Case "OPTIONSENABLED": cGetOptionFromOptions = Options.OptionsEnabled
50360   Case "OPTIONSVISIBLE": cGetOptionFromOptions = Options.OptionsVisible
50370   Case "PAPERSIZE": cGetOptionFromOptions = Options.Papersize
50380   Case "PCLCOLORSCOUNT": cGetOptionFromOptions = Options.PCLColorsCount
50390   Case "PCXCOLORSCOUNT": cGetOptionFromOptions = Options.PCXColorscount
50400   Case "PDFALLOWASSEMBLY": cGetOptionFromOptions = Options.PDFAllowAssembly
50410   Case "PDFALLOWDEGRADEDPRINTING": cGetOptionFromOptions = Options.PDFAllowDegradedPrinting
50420   Case "PDFALLOWFILLIN": cGetOptionFromOptions = Options.PDFAllowFillIn
50430   Case "PDFALLOWSCREENREADERS": cGetOptionFromOptions = Options.PDFAllowScreenReaders
50440   Case "PDFCOLORSCMYKTORGB": cGetOptionFromOptions = Options.PDFColorsCMYKToRGB
50450   Case "PDFCOLORSCOLORMODEL": cGetOptionFromOptions = Options.PDFColorsColorModel
50460   Case "PDFCOLORSPRESERVEHALFTONE": cGetOptionFromOptions = Options.PDFColorsPreserveHalftone
50470   Case "PDFCOLORSPRESERVEOVERPRINT": cGetOptionFromOptions = Options.PDFColorsPreserveOverprint
50480   Case "PDFCOLORSPRESERVETRANSFER": cGetOptionFromOptions = Options.PDFColorsPreserveTransfer
50490   Case "PDFCOMPRESSIONCOLORCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionColorCompression
50500   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionColorCompressionChoice
50510   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGHighFactor
50520   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGLowFactor
50530   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMaximumFactor
50540   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMediumFactor
50550   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMinimumFactor
50560   Case "PDFCOMPRESSIONCOLORRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionColorResample
50570   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionColorResampleChoice
50580   Case "PDFCOMPRESSIONCOLORRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionColorResolution
50590   Case "PDFCOMPRESSIONGREYCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionGreyCompression
50600   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionChoice
50610   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGHighFactor
50620   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGLowFactor
50630   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMaximumFactor
50640   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMediumFactor
50650   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMinimumFactor
50660   Case "PDFCOMPRESSIONGREYRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionGreyResample
50670   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyResampleChoice
50680   Case "PDFCOMPRESSIONGREYRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionGreyResolution
50690   Case "PDFCOMPRESSIONMONOCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionMonoCompression
50700   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoCompressionChoice
50710   Case "PDFCOMPRESSIONMONORESAMPLE": cGetOptionFromOptions = Options.PDFCompressionMonoResample
50720   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoResampleChoice
50730   Case "PDFCOMPRESSIONMONORESOLUTION": cGetOptionFromOptions = Options.PDFCompressionMonoResolution
50740   Case "PDFCOMPRESSIONTEXTCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionTextCompression
50750   Case "PDFDISALLOWCOPY": cGetOptionFromOptions = Options.PDFDisallowCopy
50760   Case "PDFDISALLOWMODIFYANNOTATIONS": cGetOptionFromOptions = Options.PDFDisallowModifyAnnotations
50770   Case "PDFDISALLOWMODIFYCONTENTS": cGetOptionFromOptions = Options.PDFDisallowModifyContents
50780   Case "PDFDISALLOWPRINTING": cGetOptionFromOptions = Options.PDFDisallowPrinting
50790   Case "PDFENCRYPTOR": cGetOptionFromOptions = Options.PDFEncryptor
50800   Case "PDFFONTSEMBEDALL": cGetOptionFromOptions = Options.PDFFontsEmbedAll
50810   Case "PDFFONTSSUBSETFONTS": cGetOptionFromOptions = Options.PDFFontsSubSetFonts
50820   Case "PDFFONTSSUBSETFONTSPERCENT": cGetOptionFromOptions = Options.PDFFontsSubSetFontsPercent
50830   Case "PDFGENERALASCII85": cGetOptionFromOptions = Options.PDFGeneralASCII85
50840   Case "PDFGENERALAUTOROTATE": cGetOptionFromOptions = Options.PDFGeneralAutorotate
50850   Case "PDFGENERALCOMPATIBILITY": cGetOptionFromOptions = Options.PDFGeneralCompatibility
50860   Case "PDFGENERALDEFAULT": cGetOptionFromOptions = Options.PDFGeneralDefault
50870   Case "PDFGENERALOVERPRINT": cGetOptionFromOptions = Options.PDFGeneralOverprint
50880   Case "PDFGENERALRESOLUTION": cGetOptionFromOptions = Options.PDFGeneralResolution
50890   Case "PDFHIGHENCRYPTION": cGetOptionFromOptions = Options.PDFHighEncryption
50900   Case "PDFLOWENCRYPTION": cGetOptionFromOptions = Options.PDFLowEncryption
50910   Case "PDFOPTIMIZE": cGetOptionFromOptions = Options.PDFOptimize
50920   Case "PDFOWNERPASS": cGetOptionFromOptions = Options.PDFOwnerPass
50930   Case "PDFOWNERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFOwnerPasswordString
50940   Case "PDFSIGNINGMULTISIGNATURE": cGetOptionFromOptions = Options.PDFSigningMultiSignature
50950   Case "PDFSIGNINGPFXFILE": cGetOptionFromOptions = Options.PDFSigningPFXFile
50960   Case "PDFSIGNINGPFXFILEPASSWORD": cGetOptionFromOptions = Options.PDFSigningPFXFilePassword
50970   Case "PDFSIGNINGSIGNATURECONTACT": cGetOptionFromOptions = Options.PDFSigningSignatureContact
50980   Case "PDFSIGNINGSIGNATURELEFTX": cGetOptionFromOptions = Options.PDFSigningSignatureLeftX
50990   Case "PDFSIGNINGSIGNATURELEFTY": cGetOptionFromOptions = Options.PDFSigningSignatureLeftY
51000   Case "PDFSIGNINGSIGNATURELOCATION": cGetOptionFromOptions = Options.PDFSigningSignatureLocation
51010   Case "PDFSIGNINGSIGNATUREREASON": cGetOptionFromOptions = Options.PDFSigningSignatureReason
51020   Case "PDFSIGNINGSIGNATURERIGHTX": cGetOptionFromOptions = Options.PDFSigningSignatureRightX
51030   Case "PDFSIGNINGSIGNATURERIGHTY": cGetOptionFromOptions = Options.PDFSigningSignatureRightY
51040   Case "PDFSIGNINGSIGNATUREVISIBLE": cGetOptionFromOptions = Options.PDFSigningSignatureVisible
51050   Case "PDFSIGNINGSIGNPDF": cGetOptionFromOptions = Options.PDFSigningSignPDF
51060   Case "PDFUPDATEMETADATA": cGetOptionFromOptions = Options.PDFUpdateMetadata
51070   Case "PDFUSERPASS": cGetOptionFromOptions = Options.PDFUserPass
51080   Case "PDFUSERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFUserPasswordString
51090   Case "PDFUSESECURITY": cGetOptionFromOptions = Options.PDFUseSecurity
51100   Case "PNGCOLORSCOUNT": cGetOptionFromOptions = Options.PNGColorscount
51110   Case "PRINTAFTERSAVING": cGetOptionFromOptions = Options.PrintAfterSaving
51120   Case "PRINTAFTERSAVINGDUPLEX": cGetOptionFromOptions = Options.PrintAfterSavingDuplex
51130   Case "PRINTAFTERSAVINGNOCANCEL": cGetOptionFromOptions = Options.PrintAfterSavingNoCancel
51140   Case "PRINTAFTERSAVINGPRINTER": cGetOptionFromOptions = Options.PrintAfterSavingPrinter
51150   Case "PRINTAFTERSAVINGQUERYUSER": cGetOptionFromOptions = Options.PrintAfterSavingQueryUser
51160   Case "PRINTAFTERSAVINGTUMBLE": cGetOptionFromOptions = Options.PrintAfterSavingTumble
51170   Case "PRINTERSTOP": cGetOptionFromOptions = Options.PrinterStop
51180   Case "PRINTERTEMPPATH": cGetOptionFromOptions = Options.PrinterTemppath
51190   Case "PROCESSPRIORITY": cGetOptionFromOptions = Options.ProcessPriority
51200   Case "PROGRAMFONT": cGetOptionFromOptions = Options.ProgramFont
51210   Case "PROGRAMFONTCHARSET": cGetOptionFromOptions = Options.ProgramFontCharset
51220   Case "PROGRAMFONTSIZE": cGetOptionFromOptions = Options.ProgramFontSize
51230   Case "PSDCOLORSCOUNT": cGetOptionFromOptions = Options.PSDColorsCount
51240   Case "PSLANGUAGELEVEL": cGetOptionFromOptions = Options.PSLanguageLevel
51250   Case "RAWCOLORSCOUNT": cGetOptionFromOptions = Options.RAWColorsCount
51260   Case "REMOVEALLKNOWNFILEEXTENSIONS": cGetOptionFromOptions = Options.RemoveAllKnownFileExtensions
51270   Case "REMOVESPACES": cGetOptionFromOptions = Options.RemoveSpaces
51280   Case "RUNPROGRAMAFTERSAVING": cGetOptionFromOptions = Options.RunProgramAfterSaving
51290   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramname
51300   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramParameters
51310   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": cGetOptionFromOptions = Options.RunProgramAfterSavingWaitUntilReady
51320   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramAfterSavingWindowstyle
51330   Case "RUNPROGRAMBEFORESAVING": cGetOptionFromOptions = Options.RunProgramBeforeSaving
51340   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramname
51350   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramParameters
51360   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramBeforeSavingWindowstyle
51370   Case "SAVEFILENAME": cGetOptionFromOptions = Options.SaveFilename
51380   Case "SENDEMAILAFTERAUTOSAVING": cGetOptionFromOptions = Options.SendEmailAfterAutoSaving
51390   Case "SENDMAILMETHOD": cGetOptionFromOptions = Options.SendMailMethod
51400   Case "SHOWANIMATION": cGetOptionFromOptions = Options.ShowAnimation
51410   Case "STAMPFONTCOLOR": cGetOptionFromOptions = Options.StampFontColor
51420   Case "STAMPFONTNAME": cGetOptionFromOptions = Options.StampFontname
51430   Case "STAMPFONTSIZE": cGetOptionFromOptions = Options.StampFontsize
51440   Case "STAMPOUTLINEFONTTHICKNESS": cGetOptionFromOptions = Options.StampOutlineFontthickness
51450   Case "STAMPSTRING": cGetOptionFromOptions = Options.StampString
51460   Case "STAMPUSEOUTLINEFONT": cGetOptionFromOptions = Options.StampUseOutlineFont
51470   Case "STANDARDAUTHOR": cGetOptionFromOptions = Options.StandardAuthor
51480   Case "STANDARDCREATIONDATE": cGetOptionFromOptions = Options.StandardCreationdate
51490   Case "STANDARDDATEFORMAT": cGetOptionFromOptions = Options.StandardDateformat
51500   Case "STANDARDKEYWORDS": cGetOptionFromOptions = Options.StandardKeywords
51510   Case "STANDARDMAILDOMAIN": cGetOptionFromOptions = Options.StandardMailDomain
51520   Case "STANDARDMODIFYDATE": cGetOptionFromOptions = Options.StandardModifydate
51530   Case "STANDARDSAVEFORMAT": cGetOptionFromOptions = Options.StandardSaveformat
51540   Case "STANDARDSUBJECT": cGetOptionFromOptions = Options.StandardSubject
51550   Case "STANDARDTITLE": cGetOptionFromOptions = Options.StandardTitle
51560   Case "STARTSTANDARDPROGRAM": cGetOptionFromOptions = Options.StartStandardProgram
51570   Case "TIFFCOLORSCOUNT": cGetOptionFromOptions = Options.TIFFColorscount
51580   Case "TOOLBARS": cGetOptionFromOptions = Options.Toolbars
51590   Case "USEAUTOSAVE": cGetOptionFromOptions = Options.UseAutosave
51600   Case "USEAUTOSAVEDIRECTORY": cGetOptionFromOptions = Options.UseAutosaveDirectory
51610   Case "USECREATIONDATENOW": cGetOptionFromOptions = Options.UseCreationDateNow
51620   Case "USECUSTOMPAPERSIZE": cGetOptionFromOptions = Options.UseCustomPaperSize
51630   Case "USEFIXPAPERSIZE": cGetOptionFromOptions = Options.UseFixPapersize
51640   Case "USESTANDARDAUTHOR": cGetOptionFromOptions = Options.UseStandardAuthor
51650   Case "XCFCOLORSCOUNT": cGetOptionFromOptions = Options.XCFColorsCount
51660  End Select
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
50010  Set cGetInfoSpoolfile = New clsPDFCreatorInfoSpoolFile
50020  With cGetInfoSpoolfile
50030   .Computer = sInfos.Computer
50040   .Created = sInfos.Created
50050   .REDMON_DOCNAME = sInfos.REDMON_DOCNAME
50060   .REDMON_FILENAME = sInfos.REDMON_FILENAME
50070   .REDMON_JOB = sInfos.REDMON_JOB
50080   .REDMON_MACHINE = sInfos.REDMON_MACHINE
50090   .REDMON_PORT = sInfos.REDMON_PORT
50100   .REDMON_PRINTER = sInfos.REDMON_PRINTER
50110   .REDMON_SESSIONID = sInfos.REDMON_SESSIONID
50120   .REDMON_USER = sInfos.REDMON_USER
50130   .SpoolerAccount = sInfos.SpoolerAccount
50140   .SpoolFilename = sInfos.SpoolFilename
50150  End With
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
50130   .Counter = sOptions.Counter
50140   .DeviceHeightPoints = sOptions.DeviceHeightPoints
50150   .DeviceWidthPoints = sOptions.DeviceWidthPoints
50160   .DirectoryGhostscriptBinaries = sOptions.DirectoryGhostscriptBinaries
50170   .DirectoryGhostscriptFonts = sOptions.DirectoryGhostscriptFonts
50180   .DirectoryGhostscriptLibraries = sOptions.DirectoryGhostscriptLibraries
50190   .DirectoryGhostscriptResource = sOptions.DirectoryGhostscriptResource
50200   .DisableEmail = sOptions.DisableEmail
50210   .DontUseDocumentSettings = sOptions.DontUseDocumentSettings
50220   .EPSLanguageLevel = sOptions.EPSLanguageLevel
50230   .FilenameSubstitutions = sOptions.FilenameSubstitutions
50240   .FilenameSubstitutionsOnlyInTitle = sOptions.FilenameSubstitutionsOnlyInTitle
50250   .JPEGColorscount = sOptions.JPEGColorscount
50260   .JPEGQuality = sOptions.JPEGQuality
50270   .Language = sOptions.Language
50280   .LastSaveDirectory = sOptions.LastSaveDirectory
50290   .Logging = sOptions.Logging
50300   .LogLines = sOptions.LogLines
50310   .NoConfirmMessageSwitchingDefaultprinter = sOptions.NoConfirmMessageSwitchingDefaultprinter
50320   .NoProcessingAtStartup = sOptions.NoProcessingAtStartup
50330   .NoPSCheck = sOptions.NoPSCheck
50340   .OnePagePerFile = sOptions.OnePagePerFile
50350   .OptionsDesign = sOptions.OptionsDesign
50360   .OptionsEnabled = sOptions.OptionsEnabled
50370   .OptionsVisible = sOptions.OptionsVisible
50380   .Papersize = sOptions.Papersize
50390   .PCLColorsCount = sOptions.PCLColorsCount
50400   .PCXColorscount = sOptions.PCXColorscount
50410   .PDFAllowAssembly = sOptions.PDFAllowAssembly
50420   .PDFAllowDegradedPrinting = sOptions.PDFAllowDegradedPrinting
50430   .PDFAllowFillIn = sOptions.PDFAllowFillIn
50440   .PDFAllowScreenReaders = sOptions.PDFAllowScreenReaders
50450   .PDFColorsCMYKToRGB = sOptions.PDFColorsCMYKToRGB
50460   .PDFColorsColorModel = sOptions.PDFColorsColorModel
50470   .PDFColorsPreserveHalftone = sOptions.PDFColorsPreserveHalftone
50480   .PDFColorsPreserveOverprint = sOptions.PDFColorsPreserveOverprint
50490   .PDFColorsPreserveTransfer = sOptions.PDFColorsPreserveTransfer
50500   .PDFCompressionColorCompression = sOptions.PDFCompressionColorCompression
50510   .PDFCompressionColorCompressionChoice = sOptions.PDFCompressionColorCompressionChoice
50520   .PDFCompressionColorCompressionJPEGHighFactor = sOptions.PDFCompressionColorCompressionJPEGHighFactor
50530   .PDFCompressionColorCompressionJPEGLowFactor = sOptions.PDFCompressionColorCompressionJPEGLowFactor
50540   .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
50550   .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
50560   .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
50570   .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
50580   .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
50590   .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
50600   .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
50610   .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
50620   .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
50630   .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
50640   .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
50650   .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
50660   .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
50670   .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
50680   .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
50690   .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
50700   .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
50710   .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
50720   .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
50730   .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
50740   .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
50750   .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
50760   .PDFDisallowCopy = sOptions.PDFDisallowCopy
50770   .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
50780   .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
50790   .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
50800   .PDFEncryptor = sOptions.PDFEncryptor
50810   .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
50820   .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
50830   .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
50840   .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
50850   .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
50860   .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
50870   .PDFGeneralDefault = sOptions.PDFGeneralDefault
50880   .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
50890   .PDFGeneralResolution = sOptions.PDFGeneralResolution
50900   .PDFHighEncryption = sOptions.PDFHighEncryption
50910   .PDFLowEncryption = sOptions.PDFLowEncryption
50920   .PDFOptimize = sOptions.PDFOptimize
50930   .PDFOwnerPass = sOptions.PDFOwnerPass
50940   .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
50950   .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
50960   .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
50970   .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
50980   .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
50990   .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
51000   .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
51010   .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
51020   .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
51030   .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
51040   .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
51050   .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
51060   .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
51070   .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
51080   .PDFUserPass = sOptions.PDFUserPass
51090   .PDFUserPasswordString = sOptions.PDFUserPasswordString
51100   .PDFUseSecurity = sOptions.PDFUseSecurity
51110   .PNGColorscount = sOptions.PNGColorscount
51120   .PrintAfterSaving = sOptions.PrintAfterSaving
51130   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
51140   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
51150   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
51160   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51170   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51180   .PrinterStop = sOptions.PrinterStop
51190   .PrinterTemppath = sOptions.PrinterTemppath
51200   .ProcessPriority = sOptions.ProcessPriority
51210   .ProgramFont = sOptions.ProgramFont
51220   .ProgramFontCharset = sOptions.ProgramFontCharset
51230   .ProgramFontSize = sOptions.ProgramFontSize
51240   .PSDColorsCount = sOptions.PSDColorsCount
51250   .PSLanguageLevel = sOptions.PSLanguageLevel
51260   .RAWColorsCount = sOptions.RAWColorsCount
51270   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51280   .RemoveSpaces = sOptions.RemoveSpaces
51290   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51300   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51310   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51320   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51330   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51340   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51350   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51360   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51370   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51380   .SaveFilename = sOptions.SaveFilename
51390   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51400   .SendMailMethod = sOptions.SendMailMethod
51410   .ShowAnimation = sOptions.ShowAnimation
51420   .StampFontColor = sOptions.StampFontColor
51430   .StampFontname = sOptions.StampFontname
51440   .StampFontsize = sOptions.StampFontsize
51450   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51460   .StampString = sOptions.StampString
51470   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51480   .StandardAuthor = sOptions.StandardAuthor
51490   .StandardCreationdate = sOptions.StandardCreationdate
51500   .StandardDateformat = sOptions.StandardDateformat
51510   .StandardKeywords = sOptions.StandardKeywords
51520   .StandardMailDomain = sOptions.StandardMailDomain
51530   .StandardModifydate = sOptions.StandardModifydate
51540   .StandardSaveformat = sOptions.StandardSaveformat
51550   .StandardSubject = sOptions.StandardSubject
51560   .StandardTitle = sOptions.StandardTitle
51570   .StartStandardProgram = sOptions.StartStandardProgram
51580   .TIFFColorscount = sOptions.TIFFColorscount
51590   .Toolbars = sOptions.Toolbars
51600   .UseAutosave = sOptions.UseAutosave
51610   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51620   .UseCreationDateNow = sOptions.UseCreationDateNow
51630   .UseCustomPaperSize = sOptions.UseCustomPaperSize
51640   .UseFixPapersize = sOptions.UseFixPapersize
51650   .UseStandardAuthor = sOptions.UseStandardAuthor
51660   .XCFColorsCount = sOptions.XCFColorsCount
51670  End With
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

Private Sub cLetOption(PropertyName As String, value As Variant)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case UCase$(PropertyName)
        Case "ADDITIONALGHOSTSCRIPTPARAMETERS": Options.AdditionalGhostscriptParameters = value
50030   Case "ADDITIONALGHOSTSCRIPTSEARCHPATH": Options.AdditionalGhostscriptSearchpath = value
50040   Case "ADDWINDOWSFONTPATH": Options.AddWindowsFontpath = value
50050   Case "AUTOSAVEDIRECTORY": Options.AutosaveDirectory = value
50060   Case "AUTOSAVEFILENAME": Options.AutosaveFilename = value
50070   Case "AUTOSAVEFORMAT": Options.AutosaveFormat = value
50080   Case "AUTOSAVESTARTSTANDARDPROGRAM": Options.AutosaveStartStandardProgram = value
50090   Case "BITMAPRESOLUTION": Options.BitmapResolution = value
50100   Case "BMPCOLORSCOUNT": Options.BMPColorscount = value
50110   Case "CLIENTCOMPUTERRESOLVEIPADDRESS": Options.ClientComputerResolveIPAddress = value
50120   Case "COUNTER": Options.Counter = value
50130   Case "DEVICEHEIGHTPOINTS": Options.DeviceHeightPoints = value
50140   Case "DEVICEWIDTHPOINTS": Options.DeviceWidthPoints = value
50150   Case "DIRECTORYGHOSTSCRIPTBINARIES": Options.DirectoryGhostscriptBinaries = value
50160   Case "DIRECTORYGHOSTSCRIPTFONTS": Options.DirectoryGhostscriptFonts = value
50170   Case "DIRECTORYGHOSTSCRIPTLIBRARIES": Options.DirectoryGhostscriptLibraries = value
50180   Case "DIRECTORYGHOSTSCRIPTRESOURCE": Options.DirectoryGhostscriptResource = value
50190   Case "DISABLEEMAIL": Options.DisableEmail = value
50200   Case "DONTUSEDOCUMENTSETTINGS": Options.DontUseDocumentSettings = value
50210   Case "EPSLANGUAGELEVEL": Options.EPSLanguageLevel = value
50220   Case "FILENAMESUBSTITUTIONS": Options.FilenameSubstitutions = value
50230   Case "FILENAMESUBSTITUTIONSONLYINTITLE": Options.FilenameSubstitutionsOnlyInTitle = value
50240   Case "JPEGCOLORSCOUNT": Options.JPEGColorscount = value
50250   Case "JPEGQUALITY": Options.JPEGQuality = value
50260   Case "LANGUAGE": Options.Language = value
50270   Case "LASTSAVEDIRECTORY": Options.LastSaveDirectory = value
50280   Case "LOGGING": Options.Logging = value
50290   Case "LOGLINES": Options.LogLines = value
50300   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": Options.NoConfirmMessageSwitchingDefaultprinter = value
50310   Case "NOPROCESSINGATSTARTUP": Options.NoProcessingAtStartup = value
50320   Case "NOPSCHECK": Options.NoPSCheck = value
50330   Case "ONEPAGEPERFILE": Options.OnePagePerFile = value
50340   Case "OPTIONSDESIGN": Options.OptionsDesign = value
50350   Case "OPTIONSENABLED": Options.OptionsEnabled = value
50360   Case "OPTIONSVISIBLE": Options.OptionsVisible = value
50370   Case "PAPERSIZE": Options.Papersize = value
50380   Case "PCLCOLORSCOUNT": Options.PCLColorsCount = value
50390   Case "PCXCOLORSCOUNT": Options.PCXColorscount = value
50400   Case "PDFALLOWASSEMBLY": Options.PDFAllowAssembly = value
50410   Case "PDFALLOWDEGRADEDPRINTING": Options.PDFAllowDegradedPrinting = value
50420   Case "PDFALLOWFILLIN": Options.PDFAllowFillIn = value
50430   Case "PDFALLOWSCREENREADERS": Options.PDFAllowScreenReaders = value
50440   Case "PDFCOLORSCMYKTORGB": Options.PDFColorsCMYKToRGB = value
50450   Case "PDFCOLORSCOLORMODEL": Options.PDFColorsColorModel = value
50460   Case "PDFCOLORSPRESERVEHALFTONE": Options.PDFColorsPreserveHalftone = value
50470   Case "PDFCOLORSPRESERVEOVERPRINT": Options.PDFColorsPreserveOverprint = value
50480   Case "PDFCOLORSPRESERVETRANSFER": Options.PDFColorsPreserveTransfer = value
50490   Case "PDFCOMPRESSIONCOLORCOMPRESSION": Options.PDFCompressionColorCompression = value
50500   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": Options.PDFCompressionColorCompressionChoice = value
50510   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": Options.PDFCompressionColorCompressionJPEGHighFactor = value
50520   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": Options.PDFCompressionColorCompressionJPEGLowFactor = value
50530   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": Options.PDFCompressionColorCompressionJPEGMaximumFactor = value
50540   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": Options.PDFCompressionColorCompressionJPEGMediumFactor = value
50550   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": Options.PDFCompressionColorCompressionJPEGMinimumFactor = value
50560   Case "PDFCOMPRESSIONCOLORRESAMPLE": Options.PDFCompressionColorResample = value
50570   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": Options.PDFCompressionColorResampleChoice = value
50580   Case "PDFCOMPRESSIONCOLORRESOLUTION": Options.PDFCompressionColorResolution = value
50590   Case "PDFCOMPRESSIONGREYCOMPRESSION": Options.PDFCompressionGreyCompression = value
50600   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": Options.PDFCompressionGreyCompressionChoice = value
50610   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": Options.PDFCompressionGreyCompressionJPEGHighFactor = value
50620   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": Options.PDFCompressionGreyCompressionJPEGLowFactor = value
50630   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMaximumFactor = value
50640   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMediumFactor = value
50650   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMinimumFactor = value
50660   Case "PDFCOMPRESSIONGREYRESAMPLE": Options.PDFCompressionGreyResample = value
50670   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": Options.PDFCompressionGreyResampleChoice = value
50680   Case "PDFCOMPRESSIONGREYRESOLUTION": Options.PDFCompressionGreyResolution = value
50690   Case "PDFCOMPRESSIONMONOCOMPRESSION": Options.PDFCompressionMonoCompression = value
50700   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": Options.PDFCompressionMonoCompressionChoice = value
50710   Case "PDFCOMPRESSIONMONORESAMPLE": Options.PDFCompressionMonoResample = value
50720   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": Options.PDFCompressionMonoResampleChoice = value
50730   Case "PDFCOMPRESSIONMONORESOLUTION": Options.PDFCompressionMonoResolution = value
50740   Case "PDFCOMPRESSIONTEXTCOMPRESSION": Options.PDFCompressionTextCompression = value
50750   Case "PDFDISALLOWCOPY": Options.PDFDisallowCopy = value
50760   Case "PDFDISALLOWMODIFYANNOTATIONS": Options.PDFDisallowModifyAnnotations = value
50770   Case "PDFDISALLOWMODIFYCONTENTS": Options.PDFDisallowModifyContents = value
50780   Case "PDFDISALLOWPRINTING": Options.PDFDisallowPrinting = value
50790   Case "PDFENCRYPTOR": Options.PDFEncryptor = value
50800   Case "PDFFONTSEMBEDALL": Options.PDFFontsEmbedAll = value
50810   Case "PDFFONTSSUBSETFONTS": Options.PDFFontsSubSetFonts = value
50820   Case "PDFFONTSSUBSETFONTSPERCENT": Options.PDFFontsSubSetFontsPercent = value
50830   Case "PDFGENERALASCII85": Options.PDFGeneralASCII85 = value
50840   Case "PDFGENERALAUTOROTATE": Options.PDFGeneralAutorotate = value
50850   Case "PDFGENERALCOMPATIBILITY": Options.PDFGeneralCompatibility = value
50860   Case "PDFGENERALDEFAULT": Options.PDFGeneralDefault = value
50870   Case "PDFGENERALOVERPRINT": Options.PDFGeneralOverprint = value
50880   Case "PDFGENERALRESOLUTION": Options.PDFGeneralResolution = value
50890   Case "PDFHIGHENCRYPTION": Options.PDFHighEncryption = value
50900   Case "PDFLOWENCRYPTION": Options.PDFLowEncryption = value
50910   Case "PDFOPTIMIZE": Options.PDFOptimize = value
50920   Case "PDFOWNERPASS": Options.PDFOwnerPass = value
50930   Case "PDFOWNERPASSWORDSTRING": Options.PDFOwnerPasswordString = value
50940   Case "PDFSIGNINGMULTISIGNATURE": Options.PDFSigningMultiSignature = value
50950   Case "PDFSIGNINGPFXFILE": Options.PDFSigningPFXFile = value
50960   Case "PDFSIGNINGPFXFILEPASSWORD": Options.PDFSigningPFXFilePassword = value
50970   Case "PDFSIGNINGSIGNATURECONTACT": Options.PDFSigningSignatureContact = value
50980   Case "PDFSIGNINGSIGNATURELEFTX": Options.PDFSigningSignatureLeftX = value
50990   Case "PDFSIGNINGSIGNATURELEFTY": Options.PDFSigningSignatureLeftY = value
51000   Case "PDFSIGNINGSIGNATURELOCATION": Options.PDFSigningSignatureLocation = value
51010   Case "PDFSIGNINGSIGNATUREREASON": Options.PDFSigningSignatureReason = value
51020   Case "PDFSIGNINGSIGNATURERIGHTX": Options.PDFSigningSignatureRightX = value
51030   Case "PDFSIGNINGSIGNATURERIGHTY": Options.PDFSigningSignatureRightY = value
51040   Case "PDFSIGNINGSIGNATUREVISIBLE": Options.PDFSigningSignatureVisible = value
51050   Case "PDFSIGNINGSIGNPDF": Options.PDFSigningSignPDF = value
51060   Case "PDFUPDATEMETADATA": Options.PDFUpdateMetadata = value
51070   Case "PDFUSERPASS": Options.PDFUserPass = value
51080   Case "PDFUSERPASSWORDSTRING": Options.PDFUserPasswordString = value
51090   Case "PDFUSESECURITY": Options.PDFUseSecurity = value
51100   Case "PNGCOLORSCOUNT": Options.PNGColorscount = value
51110   Case "PRINTAFTERSAVING": Options.PrintAfterSaving = value
51120   Case "PRINTAFTERSAVINGDUPLEX": Options.PrintAfterSavingDuplex = value
51130   Case "PRINTAFTERSAVINGNOCANCEL": Options.PrintAfterSavingNoCancel = value
51140   Case "PRINTAFTERSAVINGPRINTER": Options.PrintAfterSavingPrinter = value
51150   Case "PRINTAFTERSAVINGQUERYUSER": Options.PrintAfterSavingQueryUser = value
51160   Case "PRINTAFTERSAVINGTUMBLE": Options.PrintAfterSavingTumble = value
51170   Case "PRINTERSTOP": Options.PrinterStop = value
51180   Case "PRINTERTEMPPATH": Options.PrinterTemppath = value
51190   Case "PROCESSPRIORITY": Options.ProcessPriority = value
51200   Case "PROGRAMFONT": Options.ProgramFont = value
51210   Case "PROGRAMFONTCHARSET": Options.ProgramFontCharset = value
51220   Case "PROGRAMFONTSIZE": Options.ProgramFontSize = value
51230   Case "PSDCOLORSCOUNT": Options.PSDColorsCount = value
51240   Case "PSLANGUAGELEVEL": Options.PSLanguageLevel = value
51250   Case "RAWCOLORSCOUNT": Options.RAWColorsCount = value
51260   Case "REMOVEALLKNOWNFILEEXTENSIONS": Options.RemoveAllKnownFileExtensions = value
51270   Case "REMOVESPACES": Options.RemoveSpaces = value
51280   Case "RUNPROGRAMAFTERSAVING": Options.RunProgramAfterSaving = value
51290   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": Options.RunProgramAfterSavingProgramname = value
51300   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": Options.RunProgramAfterSavingProgramParameters = value
51310   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": Options.RunProgramAfterSavingWaitUntilReady = value
51320   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": Options.RunProgramAfterSavingWindowstyle = value
51330   Case "RUNPROGRAMBEFORESAVING": Options.RunProgramBeforeSaving = value
51340   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": Options.RunProgramBeforeSavingProgramname = value
51350   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": Options.RunProgramBeforeSavingProgramParameters = value
51360   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": Options.RunProgramBeforeSavingWindowstyle = value
51370   Case "SAVEFILENAME": Options.SaveFilename = value
51380   Case "SENDEMAILAFTERAUTOSAVING": Options.SendEmailAfterAutoSaving = value
51390   Case "SENDMAILMETHOD": Options.SendMailMethod = value
51400   Case "SHOWANIMATION": Options.ShowAnimation = value
51410   Case "STAMPFONTCOLOR": Options.StampFontColor = value
51420   Case "STAMPFONTNAME": Options.StampFontname = value
51430   Case "STAMPFONTSIZE": Options.StampFontsize = value
51440   Case "STAMPOUTLINEFONTTHICKNESS": Options.StampOutlineFontthickness = value
51450   Case "STAMPSTRING": Options.StampString = value
51460   Case "STAMPUSEOUTLINEFONT": Options.StampUseOutlineFont = value
51470   Case "STANDARDAUTHOR": Options.StandardAuthor = value
51480   Case "STANDARDCREATIONDATE": Options.StandardCreationdate = value
51490   Case "STANDARDDATEFORMAT": Options.StandardDateformat = value
51500   Case "STANDARDKEYWORDS": Options.StandardKeywords = value
51510   Case "STANDARDMAILDOMAIN": Options.StandardMailDomain = value
51520   Case "STANDARDMODIFYDATE": Options.StandardModifydate = value
51530   Case "STANDARDSAVEFORMAT": Options.StandardSaveformat = value
51540   Case "STANDARDSUBJECT": Options.StandardSubject = value
51550   Case "STANDARDTITLE": Options.StandardTitle = value
51560   Case "STARTSTANDARDPROGRAM": Options.StartStandardProgram = value
51570   Case "TIFFCOLORSCOUNT": Options.TIFFColorscount = value
51580   Case "TOOLBARS": Options.Toolbars = value
51590   Case "USEAUTOSAVE": Options.UseAutosave = value
51600   Case "USEAUTOSAVEDIRECTORY": Options.UseAutosaveDirectory = value
51610   Case "USECREATIONDATENOW": Options.UseCreationDateNow = value
51620   Case "USECUSTOMPAPERSIZE": Options.UseCustomPaperSize = value
51630   Case "USEFIXPAPERSIZE": Options.UseFixPapersize = value
51640   Case "USESTANDARDAUTHOR": Options.UseStandardAuthor = value
51650   Case "XCFCOLORSCOUNT": Options.XCFColorsCount = value
51660   Case Else:
51670    mError.Number = 3
51680    mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in cLetOption")
51690    RaiseEvent eError
51700  End Select
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
50120   .Counter = sOptions.Counter
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
50380   .PCLColorsCount = sOptions.PCLColorsCount
50390   .PCXColorscount = sOptions.PCXColorscount
50400   .PDFAllowAssembly = sOptions.PDFAllowAssembly
50410   .PDFAllowDegradedPrinting = sOptions.PDFAllowDegradedPrinting
50420   .PDFAllowFillIn = sOptions.PDFAllowFillIn
50430   .PDFAllowScreenReaders = sOptions.PDFAllowScreenReaders
50440   .PDFColorsCMYKToRGB = sOptions.PDFColorsCMYKToRGB
50450   .PDFColorsColorModel = sOptions.PDFColorsColorModel
50460   .PDFColorsPreserveHalftone = sOptions.PDFColorsPreserveHalftone
50470   .PDFColorsPreserveOverprint = sOptions.PDFColorsPreserveOverprint
50480   .PDFColorsPreserveTransfer = sOptions.PDFColorsPreserveTransfer
50490   .PDFCompressionColorCompression = sOptions.PDFCompressionColorCompression
50500   .PDFCompressionColorCompressionChoice = sOptions.PDFCompressionColorCompressionChoice
50510   .PDFCompressionColorCompressionJPEGHighFactor = sOptions.PDFCompressionColorCompressionJPEGHighFactor
50520   .PDFCompressionColorCompressionJPEGLowFactor = sOptions.PDFCompressionColorCompressionJPEGLowFactor
50530   .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
50540   .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
50550   .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
50560   .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
50570   .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
50580   .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
50590   .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
50600   .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
50610   .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
50620   .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
50630   .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
50640   .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
50650   .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
50660   .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
50670   .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
50680   .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
50690   .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
50700   .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
50710   .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
50720   .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
50730   .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
50740   .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
50750   .PDFDisallowCopy = sOptions.PDFDisallowCopy
50760   .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
50770   .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
50780   .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
50790   .PDFEncryptor = sOptions.PDFEncryptor
50800   .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
50810   .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
50820   .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
50830   .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
50840   .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
50850   .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
50860   .PDFGeneralDefault = sOptions.PDFGeneralDefault
50870   .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
50880   .PDFGeneralResolution = sOptions.PDFGeneralResolution
50890   .PDFHighEncryption = sOptions.PDFHighEncryption
50900   .PDFLowEncryption = sOptions.PDFLowEncryption
50910   .PDFOptimize = sOptions.PDFOptimize
50920   .PDFOwnerPass = sOptions.PDFOwnerPass
50930   .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
50940   .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
50950   .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
50960   .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
50970   .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
50980   .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
50990   .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
51000   .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
51010   .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
51020   .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
51030   .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
51040   .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
51050   .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
51060   .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
51070   .PDFUserPass = sOptions.PDFUserPass
51080   .PDFUserPasswordString = sOptions.PDFUserPasswordString
51090   .PDFUseSecurity = sOptions.PDFUseSecurity
51100   .PNGColorscount = sOptions.PNGColorscount
51110   .PrintAfterSaving = sOptions.PrintAfterSaving
51120   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
51130   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
51140   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
51150   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51160   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51170   .PrinterStop = sOptions.PrinterStop
51180   .PrinterTemppath = sOptions.PrinterTemppath
51190   .ProcessPriority = sOptions.ProcessPriority
51200   .ProgramFont = sOptions.ProgramFont
51210   .ProgramFontCharset = sOptions.ProgramFontCharset
51220   .ProgramFontSize = sOptions.ProgramFontSize
51230   .PSDColorsCount = sOptions.PSDColorsCount
51240   .PSLanguageLevel = sOptions.PSLanguageLevel
51250   .RAWColorsCount = sOptions.RAWColorsCount
51260   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51270   .RemoveSpaces = sOptions.RemoveSpaces
51280   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51290   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51300   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51310   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51320   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51330   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51340   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51350   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51360   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51370   .SaveFilename = sOptions.SaveFilename
51380   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51390   .SendMailMethod = sOptions.SendMailMethod
51400   .ShowAnimation = sOptions.ShowAnimation
51410   .StampFontColor = sOptions.StampFontColor
51420   .StampFontname = sOptions.StampFontname
51430   .StampFontsize = sOptions.StampFontsize
51440   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51450   .StampString = sOptions.StampString
51460   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51470   .StandardAuthor = sOptions.StandardAuthor
51480   .StandardCreationdate = sOptions.StandardCreationdate
51490   .StandardDateformat = sOptions.StandardDateformat
51500   .StandardKeywords = sOptions.StandardKeywords
51510   .StandardMailDomain = sOptions.StandardMailDomain
51520   .StandardModifydate = sOptions.StandardModifydate
51530   .StandardSaveformat = sOptions.StandardSaveformat
51540   .StandardSubject = sOptions.StandardSubject
51550   .StandardTitle = sOptions.StandardTitle
51560   .StartStandardProgram = sOptions.StartStandardProgram
51570   .TIFFColorscount = sOptions.TIFFColorscount
51580   .Toolbars = sOptions.Toolbars
51590   .UseAutosave = sOptions.UseAutosave
51600   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51610   .UseCreationDateNow = sOptions.UseCreationDateNow
51620   .UseCustomPaperSize = sOptions.UseCustomPaperSize
51630   .UseFixPapersize = sOptions.UseFixPapersize
51640   .UseStandardAuthor = sOptions.UseStandardAuthor
51650   .XCFColorsCount = sOptions.XCFColorsCount
51660  End With
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

Public Function GetPDFCreatorPrinters() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Printers As Collection, reg As clsRegistry, SubKeys As Collection, _
  i As Long, j As Long
50030  Set GetPDFCreatorPrinters = New Collection
50040  Set Printers = GetAvailablePrinters2
50050  Set reg = New clsRegistry
50060  Set SubKeys = reg.EnumRegistryKeys(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Print\Monitors\PDFCreator\Ports")
50070  For i = 1 To Printers.Count
50080   For j = 1 To SubKeys.Count
50090    If SubKeys(j) = Printers(i)(1) Then
50100     GetPDFCreatorPrinters.Add Printers(i)(0)
50110     Exit Function
50120     End If
50130    Next j
50140  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "GetPDFCreatorPrinters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

