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
Private Const ErrDescr9 = "No valid file '%1' was given!"
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
50100   .Add "BMPColorscount"
50110   .Add "BMPResolution"
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
50270   .Add "JPEGResolution"
50280   .Add "Language"
50290   .Add "LastSaveDirectory"
50300   .Add "Logging"
50310   .Add "LogLines"
50320   .Add "NoConfirmMessageSwitchingDefaultprinter"
50330   .Add "NoProcessingAtStartup"
50340   .Add "NoPSCheck"
50350   .Add "OnePagePerFile"
50360   .Add "OptionsDesign"
50370   .Add "OptionsEnabled"
50380   .Add "OptionsVisible"
50390   .Add "Papersize"
50400   .Add "PCLColorsCount"
50410   .Add "PCLResolution"
50420   .Add "PCXColorscount"
50430   .Add "PCXResolution"
50440   .Add "PDFAllowAssembly"
50450   .Add "PDFAllowDegradedPrinting"
50460   .Add "PDFAllowFillIn"
50470   .Add "PDFAllowScreenReaders"
50480   .Add "PDFColorsCMYKToRGB"
50490   .Add "PDFColorsColorModel"
50500   .Add "PDFColorsPreserveHalftone"
50510   .Add "PDFColorsPreserveOverprint"
50520   .Add "PDFColorsPreserveTransfer"
50530   .Add "PDFCompressionColorCompression"
50540   .Add "PDFCompressionColorCompressionChoice"
50550   .Add "PDFCompressionColorCompressionJPEGHighFactor"
50560   .Add "PDFCompressionColorCompressionJPEGLowFactor"
50570   .Add "PDFCompressionColorCompressionJPEGMaximumFactor"
50580   .Add "PDFCompressionColorCompressionJPEGMediumFactor"
50590   .Add "PDFCompressionColorCompressionJPEGMinimumFactor"
50600   .Add "PDFCompressionColorResample"
50610   .Add "PDFCompressionColorResampleChoice"
50620   .Add "PDFCompressionColorResolution"
50630   .Add "PDFCompressionGreyCompression"
50640   .Add "PDFCompressionGreyCompressionChoice"
50650   .Add "PDFCompressionGreyCompressionJPEGHighFactor"
50660   .Add "PDFCompressionGreyCompressionJPEGLowFactor"
50670   .Add "PDFCompressionGreyCompressionJPEGMaximumFactor"
50680   .Add "PDFCompressionGreyCompressionJPEGMediumFactor"
50690   .Add "PDFCompressionGreyCompressionJPEGMinimumFactor"
50700   .Add "PDFCompressionGreyResample"
50710   .Add "PDFCompressionGreyResampleChoice"
50720   .Add "PDFCompressionGreyResolution"
50730   .Add "PDFCompressionMonoCompression"
50740   .Add "PDFCompressionMonoCompressionChoice"
50750   .Add "PDFCompressionMonoResample"
50760   .Add "PDFCompressionMonoResampleChoice"
50770   .Add "PDFCompressionMonoResolution"
50780   .Add "PDFCompressionTextCompression"
50790   .Add "PDFDisallowCopy"
50800   .Add "PDFDisallowModifyAnnotations"
50810   .Add "PDFDisallowModifyContents"
50820   .Add "PDFDisallowPrinting"
50830   .Add "PDFEncryptor"
50840   .Add "PDFFontsEmbedAll"
50850   .Add "PDFFontsSubSetFonts"
50860   .Add "PDFFontsSubSetFontsPercent"
50870   .Add "PDFGeneralASCII85"
50880   .Add "PDFGeneralAutorotate"
50890   .Add "PDFGeneralCompatibility"
50900   .Add "PDFGeneralDefault"
50910   .Add "PDFGeneralOverprint"
50920   .Add "PDFGeneralResolution"
50930   .Add "PDFHighEncryption"
50940   .Add "PDFLowEncryption"
50950   .Add "PDFOptimize"
50960   .Add "PDFOwnerPass"
50970   .Add "PDFOwnerPasswordString"
50980   .Add "PDFSigningMultiSignature"
50990   .Add "PDFSigningPFXFile"
51000   .Add "PDFSigningPFXFilePassword"
51010   .Add "PDFSigningSignatureContact"
51020   .Add "PDFSigningSignatureLeftX"
51030   .Add "PDFSigningSignatureLeftY"
51040   .Add "PDFSigningSignatureLocation"
51050   .Add "PDFSigningSignatureReason"
51060   .Add "PDFSigningSignatureRightX"
51070   .Add "PDFSigningSignatureRightY"
51080   .Add "PDFSigningSignatureVisible"
51090   .Add "PDFSigningSignPDF"
51100   .Add "PDFUpdateMetadata"
51110   .Add "PDFUserPass"
51120   .Add "PDFUserPasswordString"
51130   .Add "PDFUseSecurity"
51140   .Add "PNGColorscount"
51150   .Add "PNGResolution"
51160   .Add "PrintAfterSaving"
51170   .Add "PrintAfterSavingDuplex"
51180   .Add "PrintAfterSavingNoCancel"
51190   .Add "PrintAfterSavingPrinter"
51200   .Add "PrintAfterSavingQueryUser"
51210   .Add "PrintAfterSavingTumble"
51220   .Add "PrinterStop"
51230   .Add "PrinterTemppath"
51240   .Add "ProcessPriority"
51250   .Add "ProgramFont"
51260   .Add "ProgramFontCharset"
51270   .Add "ProgramFontSize"
51280   .Add "PSDColorsCount"
51290   .Add "PSDResolution"
51300   .Add "PSLanguageLevel"
51310   .Add "RAWColorsCount"
51320   .Add "RAWResolution"
51330   .Add "RemoveAllKnownFileExtensions"
51340   .Add "RemoveSpaces"
51350   .Add "RunProgramAfterSaving"
51360   .Add "RunProgramAfterSavingProgramname"
51370   .Add "RunProgramAfterSavingProgramParameters"
51380   .Add "RunProgramAfterSavingWaitUntilReady"
51390   .Add "RunProgramAfterSavingWindowstyle"
51400   .Add "RunProgramBeforeSaving"
51410   .Add "RunProgramBeforeSavingProgramname"
51420   .Add "RunProgramBeforeSavingProgramParameters"
51430   .Add "RunProgramBeforeSavingWindowstyle"
51440   .Add "SaveFilename"
51450   .Add "SendEmailAfterAutoSaving"
51460   .Add "SendMailMethod"
51470   .Add "ShowAnimation"
51480   .Add "StampFontColor"
51490   .Add "StampFontname"
51500   .Add "StampFontsize"
51510   .Add "StampOutlineFontthickness"
51520   .Add "StampString"
51530   .Add "StampUseOutlineFont"
51540   .Add "StandardAuthor"
51550   .Add "StandardCreationdate"
51560   .Add "StandardDateformat"
51570   .Add "StandardKeywords"
51580   .Add "StandardMailDomain"
51590   .Add "StandardModifydate"
51600   .Add "StandardSaveformat"
51610   .Add "StandardSubject"
51620   .Add "StandardTitle"
51630   .Add "StartStandardProgram"
51640   .Add "TIFFColorscount"
51650   .Add "TIFFResolution"
51660   .Add "Toolbars"
51670   .Add "UseAutosave"
51680   .Add "UseAutosaveDirectory"
51690   .Add "UseCreationDateNow"
51700   .Add "UseCustomPaperSize"
51710   .Add "UseFixPapersize"
51720   .Add "UseStandardAuthor"
51730  End With
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
Public Sub cAddPrintjob(ByVal Filename As String) ' Supported file types: Postscript, JPEG, BMP
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, Ext As String, tFilename As String, ivgf As Boolean
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Sub
50050  End If
50060  If IsValidPath(Filename) = False Then
50070   mError.Number = 6
50080   mError.Description = Replace$(ErrDescr6, "%1", Filename)
50090   RaiseEvent eError
50100   Exit Sub
50110  End If
50120  If FileExists(Filename) = False Then
50130   mError.Number = 5
50140   mError.Description = Replace$(ErrDescr5, "%1", Filename)
50150   RaiseEvent eError
50160   Exit Sub
50170  End If
50180  ivgf = IsValidGraphicFile(Filename)
50190  If IsPostscriptFile(Filename) Or ivgf Then
50200    tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
50210    If ivgf Then
50220      Image2PS Filename, tFilename
50230     Else
50240      FileCopy Filename, tFilename
50250    End If
50260   Else
50270    mError.Number = 9
50280    mError.Description = Replace$(ErrDescr9, "%1", Filename)
50290    RaiseEvent eError
50300    Exit Sub
50310  End If
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
 
 eActionTimer.Enabled = False
 DoEvents
 Set eActionTimer = Nothing
 
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
50130  Set eActionTimer = mForm.Timer3: eActionTimer.Interval = 100
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
50090   Case "BMPCOLORSCOUNT": cGetOptionFromOptions = Options.BMPColorscount
50100   Case "BMPRESOLUTION": cGetOptionFromOptions = Options.BMPResolution
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
50260   Case "JPEGRESOLUTION": cGetOptionFromOptions = Options.JPEGResolution
50270   Case "LANGUAGE": cGetOptionFromOptions = Options.Language
50280   Case "LASTSAVEDIRECTORY": cGetOptionFromOptions = Options.LastSaveDirectory
50290   Case "LOGGING": cGetOptionFromOptions = Options.Logging
50300   Case "LOGLINES": cGetOptionFromOptions = Options.LogLines
50310   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": cGetOptionFromOptions = Options.NoConfirmMessageSwitchingDefaultprinter
50320   Case "NOPROCESSINGATSTARTUP": cGetOptionFromOptions = Options.NoProcessingAtStartup
50330   Case "NOPSCHECK": cGetOptionFromOptions = Options.NoPSCheck
50340   Case "ONEPAGEPERFILE": cGetOptionFromOptions = Options.OnePagePerFile
50350   Case "OPTIONSDESIGN": cGetOptionFromOptions = Options.OptionsDesign
50360   Case "OPTIONSENABLED": cGetOptionFromOptions = Options.OptionsEnabled
50370   Case "OPTIONSVISIBLE": cGetOptionFromOptions = Options.OptionsVisible
50380   Case "PAPERSIZE": cGetOptionFromOptions = Options.Papersize
50390   Case "PCLCOLORSCOUNT": cGetOptionFromOptions = Options.PCLColorsCount
50400   Case "PCLRESOLUTION": cGetOptionFromOptions = Options.PCLResolution
50410   Case "PCXCOLORSCOUNT": cGetOptionFromOptions = Options.PCXColorscount
50420   Case "PCXRESOLUTION": cGetOptionFromOptions = Options.PCXResolution
50430   Case "PDFALLOWASSEMBLY": cGetOptionFromOptions = Options.PDFAllowAssembly
50440   Case "PDFALLOWDEGRADEDPRINTING": cGetOptionFromOptions = Options.PDFAllowDegradedPrinting
50450   Case "PDFALLOWFILLIN": cGetOptionFromOptions = Options.PDFAllowFillIn
50460   Case "PDFALLOWSCREENREADERS": cGetOptionFromOptions = Options.PDFAllowScreenReaders
50470   Case "PDFCOLORSCMYKTORGB": cGetOptionFromOptions = Options.PDFColorsCMYKToRGB
50480   Case "PDFCOLORSCOLORMODEL": cGetOptionFromOptions = Options.PDFColorsColorModel
50490   Case "PDFCOLORSPRESERVEHALFTONE": cGetOptionFromOptions = Options.PDFColorsPreserveHalftone
50500   Case "PDFCOLORSPRESERVEOVERPRINT": cGetOptionFromOptions = Options.PDFColorsPreserveOverprint
50510   Case "PDFCOLORSPRESERVETRANSFER": cGetOptionFromOptions = Options.PDFColorsPreserveTransfer
50520   Case "PDFCOMPRESSIONCOLORCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionColorCompression
50530   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionColorCompressionChoice
50540   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGHighFactor
50550   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGLowFactor
50560   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMaximumFactor
50570   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMediumFactor
50580   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMinimumFactor
50590   Case "PDFCOMPRESSIONCOLORRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionColorResample
50600   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionColorResampleChoice
50610   Case "PDFCOMPRESSIONCOLORRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionColorResolution
50620   Case "PDFCOMPRESSIONGREYCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionGreyCompression
50630   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionChoice
50640   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGHighFactor
50650   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGLowFactor
50660   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMaximumFactor
50670   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMediumFactor
50680   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMinimumFactor
50690   Case "PDFCOMPRESSIONGREYRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionGreyResample
50700   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyResampleChoice
50710   Case "PDFCOMPRESSIONGREYRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionGreyResolution
50720   Case "PDFCOMPRESSIONMONOCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionMonoCompression
50730   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoCompressionChoice
50740   Case "PDFCOMPRESSIONMONORESAMPLE": cGetOptionFromOptions = Options.PDFCompressionMonoResample
50750   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoResampleChoice
50760   Case "PDFCOMPRESSIONMONORESOLUTION": cGetOptionFromOptions = Options.PDFCompressionMonoResolution
50770   Case "PDFCOMPRESSIONTEXTCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionTextCompression
50780   Case "PDFDISALLOWCOPY": cGetOptionFromOptions = Options.PDFDisallowCopy
50790   Case "PDFDISALLOWMODIFYANNOTATIONS": cGetOptionFromOptions = Options.PDFDisallowModifyAnnotations
50800   Case "PDFDISALLOWMODIFYCONTENTS": cGetOptionFromOptions = Options.PDFDisallowModifyContents
50810   Case "PDFDISALLOWPRINTING": cGetOptionFromOptions = Options.PDFDisallowPrinting
50820   Case "PDFENCRYPTOR": cGetOptionFromOptions = Options.PDFEncryptor
50830   Case "PDFFONTSEMBEDALL": cGetOptionFromOptions = Options.PDFFontsEmbedAll
50840   Case "PDFFONTSSUBSETFONTS": cGetOptionFromOptions = Options.PDFFontsSubSetFonts
50850   Case "PDFFONTSSUBSETFONTSPERCENT": cGetOptionFromOptions = Options.PDFFontsSubSetFontsPercent
50860   Case "PDFGENERALASCII85": cGetOptionFromOptions = Options.PDFGeneralASCII85
50870   Case "PDFGENERALAUTOROTATE": cGetOptionFromOptions = Options.PDFGeneralAutorotate
50880   Case "PDFGENERALCOMPATIBILITY": cGetOptionFromOptions = Options.PDFGeneralCompatibility
50890   Case "PDFGENERALDEFAULT": cGetOptionFromOptions = Options.PDFGeneralDefault
50900   Case "PDFGENERALOVERPRINT": cGetOptionFromOptions = Options.PDFGeneralOverprint
50910   Case "PDFGENERALRESOLUTION": cGetOptionFromOptions = Options.PDFGeneralResolution
50920   Case "PDFHIGHENCRYPTION": cGetOptionFromOptions = Options.PDFHighEncryption
50930   Case "PDFLOWENCRYPTION": cGetOptionFromOptions = Options.PDFLowEncryption
50940   Case "PDFOPTIMIZE": cGetOptionFromOptions = Options.PDFOptimize
50950   Case "PDFOWNERPASS": cGetOptionFromOptions = Options.PDFOwnerPass
50960   Case "PDFOWNERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFOwnerPasswordString
50970   Case "PDFSIGNINGMULTISIGNATURE": cGetOptionFromOptions = Options.PDFSigningMultiSignature
50980   Case "PDFSIGNINGPFXFILE": cGetOptionFromOptions = Options.PDFSigningPFXFile
50990   Case "PDFSIGNINGPFXFILEPASSWORD": cGetOptionFromOptions = Options.PDFSigningPFXFilePassword
51000   Case "PDFSIGNINGSIGNATURECONTACT": cGetOptionFromOptions = Options.PDFSigningSignatureContact
51010   Case "PDFSIGNINGSIGNATURELEFTX": cGetOptionFromOptions = Options.PDFSigningSignatureLeftX
51020   Case "PDFSIGNINGSIGNATURELEFTY": cGetOptionFromOptions = Options.PDFSigningSignatureLeftY
51030   Case "PDFSIGNINGSIGNATURELOCATION": cGetOptionFromOptions = Options.PDFSigningSignatureLocation
51040   Case "PDFSIGNINGSIGNATUREREASON": cGetOptionFromOptions = Options.PDFSigningSignatureReason
51050   Case "PDFSIGNINGSIGNATURERIGHTX": cGetOptionFromOptions = Options.PDFSigningSignatureRightX
51060   Case "PDFSIGNINGSIGNATURERIGHTY": cGetOptionFromOptions = Options.PDFSigningSignatureRightY
51070   Case "PDFSIGNINGSIGNATUREVISIBLE": cGetOptionFromOptions = Options.PDFSigningSignatureVisible
51080   Case "PDFSIGNINGSIGNPDF": cGetOptionFromOptions = Options.PDFSigningSignPDF
51090   Case "PDFUPDATEMETADATA": cGetOptionFromOptions = Options.PDFUpdateMetadata
51100   Case "PDFUSERPASS": cGetOptionFromOptions = Options.PDFUserPass
51110   Case "PDFUSERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFUserPasswordString
51120   Case "PDFUSESECURITY": cGetOptionFromOptions = Options.PDFUseSecurity
51130   Case "PNGCOLORSCOUNT": cGetOptionFromOptions = Options.PNGColorscount
51140   Case "PNGRESOLUTION": cGetOptionFromOptions = Options.PNGResolution
51150   Case "PRINTAFTERSAVING": cGetOptionFromOptions = Options.PrintAfterSaving
51160   Case "PRINTAFTERSAVINGDUPLEX": cGetOptionFromOptions = Options.PrintAfterSavingDuplex
51170   Case "PRINTAFTERSAVINGNOCANCEL": cGetOptionFromOptions = Options.PrintAfterSavingNoCancel
51180   Case "PRINTAFTERSAVINGPRINTER": cGetOptionFromOptions = Options.PrintAfterSavingPrinter
51190   Case "PRINTAFTERSAVINGQUERYUSER": cGetOptionFromOptions = Options.PrintAfterSavingQueryUser
51200   Case "PRINTAFTERSAVINGTUMBLE": cGetOptionFromOptions = Options.PrintAfterSavingTumble
51210   Case "PRINTERSTOP": cGetOptionFromOptions = Options.PrinterStop
51220   Case "PRINTERTEMPPATH": cGetOptionFromOptions = Options.PrinterTemppath
51230   Case "PROCESSPRIORITY": cGetOptionFromOptions = Options.ProcessPriority
51240   Case "PROGRAMFONT": cGetOptionFromOptions = Options.ProgramFont
51250   Case "PROGRAMFONTCHARSET": cGetOptionFromOptions = Options.ProgramFontCharset
51260   Case "PROGRAMFONTSIZE": cGetOptionFromOptions = Options.ProgramFontSize
51270   Case "PSDCOLORSCOUNT": cGetOptionFromOptions = Options.PSDColorsCount
51280   Case "PSDRESOLUTION": cGetOptionFromOptions = Options.PSDResolution
51290   Case "PSLANGUAGELEVEL": cGetOptionFromOptions = Options.PSLanguageLevel
51300   Case "RAWCOLORSCOUNT": cGetOptionFromOptions = Options.RAWColorsCount
51310   Case "RAWRESOLUTION": cGetOptionFromOptions = Options.RAWResolution
51320   Case "REMOVEALLKNOWNFILEEXTENSIONS": cGetOptionFromOptions = Options.RemoveAllKnownFileExtensions
51330   Case "REMOVESPACES": cGetOptionFromOptions = Options.RemoveSpaces
51340   Case "RUNPROGRAMAFTERSAVING": cGetOptionFromOptions = Options.RunProgramAfterSaving
51350   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramname
51360   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramParameters
51370   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": cGetOptionFromOptions = Options.RunProgramAfterSavingWaitUntilReady
51380   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramAfterSavingWindowstyle
51390   Case "RUNPROGRAMBEFORESAVING": cGetOptionFromOptions = Options.RunProgramBeforeSaving
51400   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramname
51410   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramParameters
51420   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramBeforeSavingWindowstyle
51430   Case "SAVEFILENAME": cGetOptionFromOptions = Options.SaveFilename
51440   Case "SENDEMAILAFTERAUTOSAVING": cGetOptionFromOptions = Options.SendEmailAfterAutoSaving
51450   Case "SENDMAILMETHOD": cGetOptionFromOptions = Options.SendMailMethod
51460   Case "SHOWANIMATION": cGetOptionFromOptions = Options.ShowAnimation
51470   Case "STAMPFONTCOLOR": cGetOptionFromOptions = Options.StampFontColor
51480   Case "STAMPFONTNAME": cGetOptionFromOptions = Options.StampFontname
51490   Case "STAMPFONTSIZE": cGetOptionFromOptions = Options.StampFontsize
51500   Case "STAMPOUTLINEFONTTHICKNESS": cGetOptionFromOptions = Options.StampOutlineFontthickness
51510   Case "STAMPSTRING": cGetOptionFromOptions = Options.StampString
51520   Case "STAMPUSEOUTLINEFONT": cGetOptionFromOptions = Options.StampUseOutlineFont
51530   Case "STANDARDAUTHOR": cGetOptionFromOptions = Options.StandardAuthor
51540   Case "STANDARDCREATIONDATE": cGetOptionFromOptions = Options.StandardCreationdate
51550   Case "STANDARDDATEFORMAT": cGetOptionFromOptions = Options.StandardDateformat
51560   Case "STANDARDKEYWORDS": cGetOptionFromOptions = Options.StandardKeywords
51570   Case "STANDARDMAILDOMAIN": cGetOptionFromOptions = Options.StandardMailDomain
51580   Case "STANDARDMODIFYDATE": cGetOptionFromOptions = Options.StandardModifydate
51590   Case "STANDARDSAVEFORMAT": cGetOptionFromOptions = Options.StandardSaveformat
51600   Case "STANDARDSUBJECT": cGetOptionFromOptions = Options.StandardSubject
51610   Case "STANDARDTITLE": cGetOptionFromOptions = Options.StandardTitle
51620   Case "STARTSTANDARDPROGRAM": cGetOptionFromOptions = Options.StartStandardProgram
51630   Case "TIFFCOLORSCOUNT": cGetOptionFromOptions = Options.TIFFColorscount
51640   Case "TIFFRESOLUTION": cGetOptionFromOptions = Options.TIFFResolution
51650   Case "TOOLBARS": cGetOptionFromOptions = Options.Toolbars
51660   Case "USEAUTOSAVE": cGetOptionFromOptions = Options.UseAutosave
51670   Case "USEAUTOSAVEDIRECTORY": cGetOptionFromOptions = Options.UseAutosaveDirectory
51680   Case "USECREATIONDATENOW": cGetOptionFromOptions = Options.UseCreationDateNow
51690   Case "USECUSTOMPAPERSIZE": cGetOptionFromOptions = Options.UseCustomPaperSize
51700   Case "USEFIXPAPERSIZE": cGetOptionFromOptions = Options.UseFixPapersize
51710   Case "USESTANDARDAUTHOR": cGetOptionFromOptions = Options.UseStandardAuthor
51720  End Select
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
50100   .BMPColorscount = sOptions.BMPColorscount
50110   .BMPResolution = sOptions.BMPResolution
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
50270   .JPEGResolution = sOptions.JPEGResolution
50280   .Language = sOptions.Language
50290   .LastSaveDirectory = sOptions.LastSaveDirectory
50300   .Logging = sOptions.Logging
50310   .LogLines = sOptions.LogLines
50320   .NoConfirmMessageSwitchingDefaultprinter = sOptions.NoConfirmMessageSwitchingDefaultprinter
50330   .NoProcessingAtStartup = sOptions.NoProcessingAtStartup
50340   .NoPSCheck = sOptions.NoPSCheck
50350   .OnePagePerFile = sOptions.OnePagePerFile
50360   .OptionsDesign = sOptions.OptionsDesign
50370   .OptionsEnabled = sOptions.OptionsEnabled
50380   .OptionsVisible = sOptions.OptionsVisible
50390   .Papersize = sOptions.Papersize
50400   .PCLColorsCount = sOptions.PCLColorsCount
50410   .PCLResolution = sOptions.PCLResolution
50420   .PCXColorscount = sOptions.PCXColorscount
50430   .PCXResolution = sOptions.PCXResolution
50440   .PDFAllowAssembly = sOptions.PDFAllowAssembly
50450   .PDFAllowDegradedPrinting = sOptions.PDFAllowDegradedPrinting
50460   .PDFAllowFillIn = sOptions.PDFAllowFillIn
50470   .PDFAllowScreenReaders = sOptions.PDFAllowScreenReaders
50480   .PDFColorsCMYKToRGB = sOptions.PDFColorsCMYKToRGB
50490   .PDFColorsColorModel = sOptions.PDFColorsColorModel
50500   .PDFColorsPreserveHalftone = sOptions.PDFColorsPreserveHalftone
50510   .PDFColorsPreserveOverprint = sOptions.PDFColorsPreserveOverprint
50520   .PDFColorsPreserveTransfer = sOptions.PDFColorsPreserveTransfer
50530   .PDFCompressionColorCompression = sOptions.PDFCompressionColorCompression
50540   .PDFCompressionColorCompressionChoice = sOptions.PDFCompressionColorCompressionChoice
50550   .PDFCompressionColorCompressionJPEGHighFactor = sOptions.PDFCompressionColorCompressionJPEGHighFactor
50560   .PDFCompressionColorCompressionJPEGLowFactor = sOptions.PDFCompressionColorCompressionJPEGLowFactor
50570   .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
50580   .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
50590   .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
50600   .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
50610   .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
50620   .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
50630   .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
50640   .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
50650   .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
50660   .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
50670   .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
50680   .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
50690   .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
50700   .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
50710   .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
50720   .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
50730   .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
50740   .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
50750   .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
50760   .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
50770   .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
50780   .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
50790   .PDFDisallowCopy = sOptions.PDFDisallowCopy
50800   .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
50810   .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
50820   .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
50830   .PDFEncryptor = sOptions.PDFEncryptor
50840   .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
50850   .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
50860   .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
50870   .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
50880   .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
50890   .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
50900   .PDFGeneralDefault = sOptions.PDFGeneralDefault
50910   .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
50920   .PDFGeneralResolution = sOptions.PDFGeneralResolution
50930   .PDFHighEncryption = sOptions.PDFHighEncryption
50940   .PDFLowEncryption = sOptions.PDFLowEncryption
50950   .PDFOptimize = sOptions.PDFOptimize
50960   .PDFOwnerPass = sOptions.PDFOwnerPass
50970   .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
50980   .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
50990   .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
51000   .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
51010   .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
51020   .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
51030   .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
51040   .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
51050   .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
51060   .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
51070   .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
51080   .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
51090   .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
51100   .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
51110   .PDFUserPass = sOptions.PDFUserPass
51120   .PDFUserPasswordString = sOptions.PDFUserPasswordString
51130   .PDFUseSecurity = sOptions.PDFUseSecurity
51140   .PNGColorscount = sOptions.PNGColorscount
51150   .PNGResolution = sOptions.PNGResolution
51160   .PrintAfterSaving = sOptions.PrintAfterSaving
51170   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
51180   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
51190   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
51200   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51210   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51220   .PrinterStop = sOptions.PrinterStop
51230   .PrinterTemppath = sOptions.PrinterTemppath
51240   .ProcessPriority = sOptions.ProcessPriority
51250   .ProgramFont = sOptions.ProgramFont
51260   .ProgramFontCharset = sOptions.ProgramFontCharset
51270   .ProgramFontSize = sOptions.ProgramFontSize
51280   .PSDColorsCount = sOptions.PSDColorsCount
51290   .PSDResolution = sOptions.PSDResolution
51300   .PSLanguageLevel = sOptions.PSLanguageLevel
51310   .RAWColorsCount = sOptions.RAWColorsCount
51320   .RAWResolution = sOptions.RAWResolution
51330   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51340   .RemoveSpaces = sOptions.RemoveSpaces
51350   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51360   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51370   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51380   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51390   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51400   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51410   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51420   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51430   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51440   .SaveFilename = sOptions.SaveFilename
51450   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51460   .SendMailMethod = sOptions.SendMailMethod
51470   .ShowAnimation = sOptions.ShowAnimation
51480   .StampFontColor = sOptions.StampFontColor
51490   .StampFontname = sOptions.StampFontname
51500   .StampFontsize = sOptions.StampFontsize
51510   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51520   .StampString = sOptions.StampString
51530   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51540   .StandardAuthor = sOptions.StandardAuthor
51550   .StandardCreationdate = sOptions.StandardCreationdate
51560   .StandardDateformat = sOptions.StandardDateformat
51570   .StandardKeywords = sOptions.StandardKeywords
51580   .StandardMailDomain = sOptions.StandardMailDomain
51590   .StandardModifydate = sOptions.StandardModifydate
51600   .StandardSaveformat = sOptions.StandardSaveformat
51610   .StandardSubject = sOptions.StandardSubject
51620   .StandardTitle = sOptions.StandardTitle
51630   .StartStandardProgram = sOptions.StartStandardProgram
51640   .TIFFColorscount = sOptions.TIFFColorscount
51650   .TIFFResolution = sOptions.TIFFResolution
51660   .Toolbars = sOptions.Toolbars
51670   .UseAutosave = sOptions.UseAutosave
51680   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51690   .UseCreationDateNow = sOptions.UseCreationDateNow
51700   .UseCustomPaperSize = sOptions.UseCustomPaperSize
51710   .UseFixPapersize = sOptions.UseFixPapersize
51720   .UseStandardAuthor = sOptions.UseStandardAuthor
51730  End With
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
50090   Case "BMPCOLORSCOUNT": Options.BMPColorscount = value
50100   Case "BMPRESOLUTION": Options.BMPResolution = value
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
50260   Case "JPEGRESOLUTION": Options.JPEGResolution = value
50270   Case "LANGUAGE": Options.Language = value
50280   Case "LASTSAVEDIRECTORY": Options.LastSaveDirectory = value
50290   Case "LOGGING": Options.Logging = value
50300   Case "LOGLINES": Options.LogLines = value
50310   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": Options.NoConfirmMessageSwitchingDefaultprinter = value
50320   Case "NOPROCESSINGATSTARTUP": Options.NoProcessingAtStartup = value
50330   Case "NOPSCHECK": Options.NoPSCheck = value
50340   Case "ONEPAGEPERFILE": Options.OnePagePerFile = value
50350   Case "OPTIONSDESIGN": Options.OptionsDesign = value
50360   Case "OPTIONSENABLED": Options.OptionsEnabled = value
50370   Case "OPTIONSVISIBLE": Options.OptionsVisible = value
50380   Case "PAPERSIZE": Options.Papersize = value
50390   Case "PCLCOLORSCOUNT": Options.PCLColorsCount = value
50400   Case "PCLRESOLUTION": Options.PCLResolution = value
50410   Case "PCXCOLORSCOUNT": Options.PCXColorscount = value
50420   Case "PCXRESOLUTION": Options.PCXResolution = value
50430   Case "PDFALLOWASSEMBLY": Options.PDFAllowAssembly = value
50440   Case "PDFALLOWDEGRADEDPRINTING": Options.PDFAllowDegradedPrinting = value
50450   Case "PDFALLOWFILLIN": Options.PDFAllowFillIn = value
50460   Case "PDFALLOWSCREENREADERS": Options.PDFAllowScreenReaders = value
50470   Case "PDFCOLORSCMYKTORGB": Options.PDFColorsCMYKToRGB = value
50480   Case "PDFCOLORSCOLORMODEL": Options.PDFColorsColorModel = value
50490   Case "PDFCOLORSPRESERVEHALFTONE": Options.PDFColorsPreserveHalftone = value
50500   Case "PDFCOLORSPRESERVEOVERPRINT": Options.PDFColorsPreserveOverprint = value
50510   Case "PDFCOLORSPRESERVETRANSFER": Options.PDFColorsPreserveTransfer = value
50520   Case "PDFCOMPRESSIONCOLORCOMPRESSION": Options.PDFCompressionColorCompression = value
50530   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": Options.PDFCompressionColorCompressionChoice = value
50540   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": Options.PDFCompressionColorCompressionJPEGHighFactor = value
50550   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": Options.PDFCompressionColorCompressionJPEGLowFactor = value
50560   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": Options.PDFCompressionColorCompressionJPEGMaximumFactor = value
50570   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": Options.PDFCompressionColorCompressionJPEGMediumFactor = value
50580   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": Options.PDFCompressionColorCompressionJPEGMinimumFactor = value
50590   Case "PDFCOMPRESSIONCOLORRESAMPLE": Options.PDFCompressionColorResample = value
50600   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": Options.PDFCompressionColorResampleChoice = value
50610   Case "PDFCOMPRESSIONCOLORRESOLUTION": Options.PDFCompressionColorResolution = value
50620   Case "PDFCOMPRESSIONGREYCOMPRESSION": Options.PDFCompressionGreyCompression = value
50630   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": Options.PDFCompressionGreyCompressionChoice = value
50640   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": Options.PDFCompressionGreyCompressionJPEGHighFactor = value
50650   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": Options.PDFCompressionGreyCompressionJPEGLowFactor = value
50660   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMaximumFactor = value
50670   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMediumFactor = value
50680   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMinimumFactor = value
50690   Case "PDFCOMPRESSIONGREYRESAMPLE": Options.PDFCompressionGreyResample = value
50700   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": Options.PDFCompressionGreyResampleChoice = value
50710   Case "PDFCOMPRESSIONGREYRESOLUTION": Options.PDFCompressionGreyResolution = value
50720   Case "PDFCOMPRESSIONMONOCOMPRESSION": Options.PDFCompressionMonoCompression = value
50730   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": Options.PDFCompressionMonoCompressionChoice = value
50740   Case "PDFCOMPRESSIONMONORESAMPLE": Options.PDFCompressionMonoResample = value
50750   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": Options.PDFCompressionMonoResampleChoice = value
50760   Case "PDFCOMPRESSIONMONORESOLUTION": Options.PDFCompressionMonoResolution = value
50770   Case "PDFCOMPRESSIONTEXTCOMPRESSION": Options.PDFCompressionTextCompression = value
50780   Case "PDFDISALLOWCOPY": Options.PDFDisallowCopy = value
50790   Case "PDFDISALLOWMODIFYANNOTATIONS": Options.PDFDisallowModifyAnnotations = value
50800   Case "PDFDISALLOWMODIFYCONTENTS": Options.PDFDisallowModifyContents = value
50810   Case "PDFDISALLOWPRINTING": Options.PDFDisallowPrinting = value
50820   Case "PDFENCRYPTOR": Options.PDFEncryptor = value
50830   Case "PDFFONTSEMBEDALL": Options.PDFFontsEmbedAll = value
50840   Case "PDFFONTSSUBSETFONTS": Options.PDFFontsSubSetFonts = value
50850   Case "PDFFONTSSUBSETFONTSPERCENT": Options.PDFFontsSubSetFontsPercent = value
50860   Case "PDFGENERALASCII85": Options.PDFGeneralASCII85 = value
50870   Case "PDFGENERALAUTOROTATE": Options.PDFGeneralAutorotate = value
50880   Case "PDFGENERALCOMPATIBILITY": Options.PDFGeneralCompatibility = value
50890   Case "PDFGENERALDEFAULT": Options.PDFGeneralDefault = value
50900   Case "PDFGENERALOVERPRINT": Options.PDFGeneralOverprint = value
50910   Case "PDFGENERALRESOLUTION": Options.PDFGeneralResolution = value
50920   Case "PDFHIGHENCRYPTION": Options.PDFHighEncryption = value
50930   Case "PDFLOWENCRYPTION": Options.PDFLowEncryption = value
50940   Case "PDFOPTIMIZE": Options.PDFOptimize = value
50950   Case "PDFOWNERPASS": Options.PDFOwnerPass = value
50960   Case "PDFOWNERPASSWORDSTRING": Options.PDFOwnerPasswordString = value
50970   Case "PDFSIGNINGMULTISIGNATURE": Options.PDFSigningMultiSignature = value
50980   Case "PDFSIGNINGPFXFILE": Options.PDFSigningPFXFile = value
50990   Case "PDFSIGNINGPFXFILEPASSWORD": Options.PDFSigningPFXFilePassword = value
51000   Case "PDFSIGNINGSIGNATURECONTACT": Options.PDFSigningSignatureContact = value
51010   Case "PDFSIGNINGSIGNATURELEFTX": Options.PDFSigningSignatureLeftX = value
51020   Case "PDFSIGNINGSIGNATURELEFTY": Options.PDFSigningSignatureLeftY = value
51030   Case "PDFSIGNINGSIGNATURELOCATION": Options.PDFSigningSignatureLocation = value
51040   Case "PDFSIGNINGSIGNATUREREASON": Options.PDFSigningSignatureReason = value
51050   Case "PDFSIGNINGSIGNATURERIGHTX": Options.PDFSigningSignatureRightX = value
51060   Case "PDFSIGNINGSIGNATURERIGHTY": Options.PDFSigningSignatureRightY = value
51070   Case "PDFSIGNINGSIGNATUREVISIBLE": Options.PDFSigningSignatureVisible = value
51080   Case "PDFSIGNINGSIGNPDF": Options.PDFSigningSignPDF = value
51090   Case "PDFUPDATEMETADATA": Options.PDFUpdateMetadata = value
51100   Case "PDFUSERPASS": Options.PDFUserPass = value
51110   Case "PDFUSERPASSWORDSTRING": Options.PDFUserPasswordString = value
51120   Case "PDFUSESECURITY": Options.PDFUseSecurity = value
51130   Case "PNGCOLORSCOUNT": Options.PNGColorscount = value
51140   Case "PNGRESOLUTION": Options.PNGResolution = value
51150   Case "PRINTAFTERSAVING": Options.PrintAfterSaving = value
51160   Case "PRINTAFTERSAVINGDUPLEX": Options.PrintAfterSavingDuplex = value
51170   Case "PRINTAFTERSAVINGNOCANCEL": Options.PrintAfterSavingNoCancel = value
51180   Case "PRINTAFTERSAVINGPRINTER": Options.PrintAfterSavingPrinter = value
51190   Case "PRINTAFTERSAVINGQUERYUSER": Options.PrintAfterSavingQueryUser = value
51200   Case "PRINTAFTERSAVINGTUMBLE": Options.PrintAfterSavingTumble = value
51210   Case "PRINTERSTOP": Options.PrinterStop = value
51220   Case "PRINTERTEMPPATH": Options.PrinterTemppath = value
51230   Case "PROCESSPRIORITY": Options.ProcessPriority = value
51240   Case "PROGRAMFONT": Options.ProgramFont = value
51250   Case "PROGRAMFONTCHARSET": Options.ProgramFontCharset = value
51260   Case "PROGRAMFONTSIZE": Options.ProgramFontSize = value
51270   Case "PSDCOLORSCOUNT": Options.PSDColorsCount = value
51280   Case "PSDRESOLUTION": Options.PSDResolution = value
51290   Case "PSLANGUAGELEVEL": Options.PSLanguageLevel = value
51300   Case "RAWCOLORSCOUNT": Options.RAWColorsCount = value
51310   Case "RAWRESOLUTION": Options.RAWResolution = value
51320   Case "REMOVEALLKNOWNFILEEXTENSIONS": Options.RemoveAllKnownFileExtensions = value
51330   Case "REMOVESPACES": Options.RemoveSpaces = value
51340   Case "RUNPROGRAMAFTERSAVING": Options.RunProgramAfterSaving = value
51350   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": Options.RunProgramAfterSavingProgramname = value
51360   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": Options.RunProgramAfterSavingProgramParameters = value
51370   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": Options.RunProgramAfterSavingWaitUntilReady = value
51380   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": Options.RunProgramAfterSavingWindowstyle = value
51390   Case "RUNPROGRAMBEFORESAVING": Options.RunProgramBeforeSaving = value
51400   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": Options.RunProgramBeforeSavingProgramname = value
51410   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": Options.RunProgramBeforeSavingProgramParameters = value
51420   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": Options.RunProgramBeforeSavingWindowstyle = value
51430   Case "SAVEFILENAME": Options.SaveFilename = value
51440   Case "SENDEMAILAFTERAUTOSAVING": Options.SendEmailAfterAutoSaving = value
51450   Case "SENDMAILMETHOD": Options.SendMailMethod = value
51460   Case "SHOWANIMATION": Options.ShowAnimation = value
51470   Case "STAMPFONTCOLOR": Options.StampFontColor = value
51480   Case "STAMPFONTNAME": Options.StampFontname = value
51490   Case "STAMPFONTSIZE": Options.StampFontsize = value
51500   Case "STAMPOUTLINEFONTTHICKNESS": Options.StampOutlineFontthickness = value
51510   Case "STAMPSTRING": Options.StampString = value
51520   Case "STAMPUSEOUTLINEFONT": Options.StampUseOutlineFont = value
51530   Case "STANDARDAUTHOR": Options.StandardAuthor = value
51540   Case "STANDARDCREATIONDATE": Options.StandardCreationdate = value
51550   Case "STANDARDDATEFORMAT": Options.StandardDateformat = value
51560   Case "STANDARDKEYWORDS": Options.StandardKeywords = value
51570   Case "STANDARDMAILDOMAIN": Options.StandardMailDomain = value
51580   Case "STANDARDMODIFYDATE": Options.StandardModifydate = value
51590   Case "STANDARDSAVEFORMAT": Options.StandardSaveformat = value
51600   Case "STANDARDSUBJECT": Options.StandardSubject = value
51610   Case "STANDARDTITLE": Options.StandardTitle = value
51620   Case "STARTSTANDARDPROGRAM": Options.StartStandardProgram = value
51630   Case "TIFFCOLORSCOUNT": Options.TIFFColorscount = value
51640   Case "TIFFRESOLUTION": Options.TIFFResolution = value
51650   Case "TOOLBARS": Options.Toolbars = value
51660   Case "USEAUTOSAVE": Options.UseAutosave = value
51670   Case "USEAUTOSAVEDIRECTORY": Options.UseAutosaveDirectory = value
51680   Case "USECREATIONDATENOW": Options.UseCreationDateNow = value
51690   Case "USECUSTOMPAPERSIZE": Options.UseCustomPaperSize = value
51700   Case "USEFIXPAPERSIZE": Options.UseFixPapersize = value
51710   Case "USESTANDARDAUTHOR": Options.UseStandardAuthor = value
51720   Case Else:
51730    mError.Number = 3
51740    mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in cLetOption")
51750    RaiseEvent eError
51760  End Select
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
50090   .BMPColorscount = sOptions.BMPColorscount
50100   .BMPResolution = sOptions.BMPResolution
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
50260   .JPEGResolution = sOptions.JPEGResolution
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
50400   .PCLResolution = sOptions.PCLResolution
50410   .PCXColorscount = sOptions.PCXColorscount
50420   .PCXResolution = sOptions.PCXResolution
50430   .PDFAllowAssembly = sOptions.PDFAllowAssembly
50440   .PDFAllowDegradedPrinting = sOptions.PDFAllowDegradedPrinting
50450   .PDFAllowFillIn = sOptions.PDFAllowFillIn
50460   .PDFAllowScreenReaders = sOptions.PDFAllowScreenReaders
50470   .PDFColorsCMYKToRGB = sOptions.PDFColorsCMYKToRGB
50480   .PDFColorsColorModel = sOptions.PDFColorsColorModel
50490   .PDFColorsPreserveHalftone = sOptions.PDFColorsPreserveHalftone
50500   .PDFColorsPreserveOverprint = sOptions.PDFColorsPreserveOverprint
50510   .PDFColorsPreserveTransfer = sOptions.PDFColorsPreserveTransfer
50520   .PDFCompressionColorCompression = sOptions.PDFCompressionColorCompression
50530   .PDFCompressionColorCompressionChoice = sOptions.PDFCompressionColorCompressionChoice
50540   .PDFCompressionColorCompressionJPEGHighFactor = sOptions.PDFCompressionColorCompressionJPEGHighFactor
50550   .PDFCompressionColorCompressionJPEGLowFactor = sOptions.PDFCompressionColorCompressionJPEGLowFactor
50560   .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
50570   .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
50580   .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
50590   .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
50600   .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
50610   .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
50620   .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
50630   .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
50640   .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
50650   .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
50660   .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
50670   .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
50680   .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
50690   .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
50700   .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
50710   .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
50720   .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
50730   .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
50740   .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
50750   .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
50760   .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
50770   .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
50780   .PDFDisallowCopy = sOptions.PDFDisallowCopy
50790   .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
50800   .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
50810   .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
50820   .PDFEncryptor = sOptions.PDFEncryptor
50830   .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
50840   .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
50850   .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
50860   .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
50870   .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
50880   .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
50890   .PDFGeneralDefault = sOptions.PDFGeneralDefault
50900   .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
50910   .PDFGeneralResolution = sOptions.PDFGeneralResolution
50920   .PDFHighEncryption = sOptions.PDFHighEncryption
50930   .PDFLowEncryption = sOptions.PDFLowEncryption
50940   .PDFOptimize = sOptions.PDFOptimize
50950   .PDFOwnerPass = sOptions.PDFOwnerPass
50960   .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
50970   .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
50980   .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
50990   .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
51000   .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
51010   .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
51020   .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
51030   .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
51040   .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
51050   .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
51060   .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
51070   .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
51080   .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
51090   .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
51100   .PDFUserPass = sOptions.PDFUserPass
51110   .PDFUserPasswordString = sOptions.PDFUserPasswordString
51120   .PDFUseSecurity = sOptions.PDFUseSecurity
51130   .PNGColorscount = sOptions.PNGColorscount
51140   .PNGResolution = sOptions.PNGResolution
51150   .PrintAfterSaving = sOptions.PrintAfterSaving
51160   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
51170   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
51180   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
51190   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51200   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51210   .PrinterStop = sOptions.PrinterStop
51220   .PrinterTemppath = sOptions.PrinterTemppath
51230   .ProcessPriority = sOptions.ProcessPriority
51240   .ProgramFont = sOptions.ProgramFont
51250   .ProgramFontCharset = sOptions.ProgramFontCharset
51260   .ProgramFontSize = sOptions.ProgramFontSize
51270   .PSDColorsCount = sOptions.PSDColorsCount
51280   .PSDResolution = sOptions.PSDResolution
51290   .PSLanguageLevel = sOptions.PSLanguageLevel
51300   .RAWColorsCount = sOptions.RAWColorsCount
51310   .RAWResolution = sOptions.RAWResolution
51320   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51330   .RemoveSpaces = sOptions.RemoveSpaces
51340   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51350   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51360   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51370   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51380   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51390   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51400   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51410   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51420   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51430   .SaveFilename = sOptions.SaveFilename
51440   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51450   .SendMailMethod = sOptions.SendMailMethod
51460   .ShowAnimation = sOptions.ShowAnimation
51470   .StampFontColor = sOptions.StampFontColor
51480   .StampFontname = sOptions.StampFontname
51490   .StampFontsize = sOptions.StampFontsize
51500   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51510   .StampString = sOptions.StampString
51520   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51530   .StandardAuthor = sOptions.StandardAuthor
51540   .StandardCreationdate = sOptions.StandardCreationdate
51550   .StandardDateformat = sOptions.StandardDateformat
51560   .StandardKeywords = sOptions.StandardKeywords
51570   .StandardMailDomain = sOptions.StandardMailDomain
51580   .StandardModifydate = sOptions.StandardModifydate
51590   .StandardSaveformat = sOptions.StandardSaveformat
51600   .StandardSubject = sOptions.StandardSubject
51610   .StandardTitle = sOptions.StandardTitle
51620   .StartStandardProgram = sOptions.StartStandardProgram
51630   .TIFFColorscount = sOptions.TIFFColorscount
51640   .TIFFResolution = sOptions.TIFFResolution
51650   .Toolbars = sOptions.Toolbars
51660   .UseAutosave = sOptions.UseAutosave
51670   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51680   .UseCreationDateNow = sOptions.UseCreationDateNow
51690   .UseCustomPaperSize = sOptions.UseCustomPaperSize
51700   .UseFixPapersize = sOptions.UseFixPapersize
51710   .UseStandardAuthor = sOptions.UseStandardAuthor
51720  End With
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
