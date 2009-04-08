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
50060   .Add "AllowSpecialGSCharsInFilenames"
50070   .Add "AutosaveDirectory"
50080   .Add "AutosaveFilename"
50090   .Add "AutosaveFormat"
50100   .Add "AutosaveStartStandardProgram"
50110   .Add "BMPColorscount"
50120   .Add "BMPResolution"
50130   .Add "ClientComputerResolveIPAddress"
50140   .Add "Counter"
50150   .Add "DeviceHeightPoints"
50160   .Add "DeviceWidthPoints"
50170   .Add "DirectoryGhostscriptBinaries"
50180   .Add "DirectoryGhostscriptFonts"
50190   .Add "DirectoryGhostscriptLibraries"
50200   .Add "DirectoryGhostscriptResource"
50210   .Add "DisableEmail"
50220   .Add "DontUseDocumentSettings"
50230   .Add "EPSLanguageLevel"
50240   .Add "FilenameSubstitutions"
50250   .Add "FilenameSubstitutionsOnlyInTitle"
50260   .Add "JPEGColorscount"
50270   .Add "JPEGQuality"
50280   .Add "JPEGResolution"
50290   .Add "Language"
50300   .Add "LastSaveDirectory"
50310   .Add "LastUpdateCheck"
50320   .Add "Logging"
50330   .Add "LogLines"
50340   .Add "NoConfirmMessageSwitchingDefaultprinter"
50350   .Add "NoProcessingAtStartup"
50360   .Add "NoPSCheck"
50370   .Add "OnePagePerFile"
50380   .Add "OptionsDesign"
50390   .Add "OptionsEnabled"
50400   .Add "OptionsVisible"
50410   .Add "Papersize"
50420   .Add "PCLColorsCount"
50430   .Add "PCLResolution"
50440   .Add "PCXColorscount"
50450   .Add "PCXResolution"
50460   .Add "PDFAllowAssembly"
50470   .Add "PDFAllowDegradedPrinting"
50480   .Add "PDFAllowFillIn"
50490   .Add "PDFAllowScreenReaders"
50500   .Add "PDFColorsCMYKToRGB"
50510   .Add "PDFColorsColorModel"
50520   .Add "PDFColorsPreserveHalftone"
50530   .Add "PDFColorsPreserveOverprint"
50540   .Add "PDFColorsPreserveTransfer"
50550   .Add "PDFCompressionColorCompression"
50560   .Add "PDFCompressionColorCompressionChoice"
50570   .Add "PDFCompressionColorCompressionJPEGHighFactor"
50580   .Add "PDFCompressionColorCompressionJPEGLowFactor"
50590   .Add "PDFCompressionColorCompressionJPEGMaximumFactor"
50600   .Add "PDFCompressionColorCompressionJPEGMediumFactor"
50610   .Add "PDFCompressionColorCompressionJPEGMinimumFactor"
50620   .Add "PDFCompressionColorResample"
50630   .Add "PDFCompressionColorResampleChoice"
50640   .Add "PDFCompressionColorResolution"
50650   .Add "PDFCompressionGreyCompression"
50660   .Add "PDFCompressionGreyCompressionChoice"
50670   .Add "PDFCompressionGreyCompressionJPEGHighFactor"
50680   .Add "PDFCompressionGreyCompressionJPEGLowFactor"
50690   .Add "PDFCompressionGreyCompressionJPEGMaximumFactor"
50700   .Add "PDFCompressionGreyCompressionJPEGMediumFactor"
50710   .Add "PDFCompressionGreyCompressionJPEGMinimumFactor"
50720   .Add "PDFCompressionGreyResample"
50730   .Add "PDFCompressionGreyResampleChoice"
50740   .Add "PDFCompressionGreyResolution"
50750   .Add "PDFCompressionMonoCompression"
50760   .Add "PDFCompressionMonoCompressionChoice"
50770   .Add "PDFCompressionMonoResample"
50780   .Add "PDFCompressionMonoResampleChoice"
50790   .Add "PDFCompressionMonoResolution"
50800   .Add "PDFCompressionTextCompression"
50810   .Add "PDFDisallowCopy"
50820   .Add "PDFDisallowModifyAnnotations"
50830   .Add "PDFDisallowModifyContents"
50840   .Add "PDFDisallowPrinting"
50850   .Add "PDFEncryptor"
50860   .Add "PDFFontsEmbedAll"
50870   .Add "PDFFontsSubSetFonts"
50880   .Add "PDFFontsSubSetFontsPercent"
50890   .Add "PDFGeneralASCII85"
50900   .Add "PDFGeneralAutorotate"
50910   .Add "PDFGeneralCompatibility"
50920   .Add "PDFGeneralDefault"
50930   .Add "PDFGeneralOverprint"
50940   .Add "PDFGeneralResolution"
50950   .Add "PDFHighEncryption"
50960   .Add "PDFLowEncryption"
50970   .Add "PDFOptimize"
50980   .Add "PDFOwnerPass"
50990   .Add "PDFOwnerPasswordString"
51000   .Add "PDFSigningMultiSignature"
51010   .Add "PDFSigningPFXFile"
51020   .Add "PDFSigningPFXFilePassword"
51030   .Add "PDFSigningSignatureContact"
51040   .Add "PDFSigningSignatureLeftX"
51050   .Add "PDFSigningSignatureLeftY"
51060   .Add "PDFSigningSignatureLocation"
51070   .Add "PDFSigningSignatureReason"
51080   .Add "PDFSigningSignatureRightX"
51090   .Add "PDFSigningSignatureRightY"
51100   .Add "PDFSigningSignatureVisible"
51110   .Add "PDFSigningSignPDF"
51120   .Add "PDFUpdateMetadata"
51130   .Add "PDFUserPass"
51140   .Add "PDFUserPasswordString"
51150   .Add "PDFUseSecurity"
51160   .Add "PNGColorscount"
51170   .Add "PNGResolution"
51180   .Add "PrintAfterSaving"
51190   .Add "PrintAfterSavingDuplex"
51200   .Add "PrintAfterSavingNoCancel"
51210   .Add "PrintAfterSavingPrinter"
51220   .Add "PrintAfterSavingQueryUser"
51230   .Add "PrintAfterSavingTumble"
51240   .Add "PrinterStop"
51250   .Add "PrinterTemppath"
51260   .Add "ProcessPriority"
51270   .Add "ProgramFont"
51280   .Add "ProgramFontCharset"
51290   .Add "ProgramFontSize"
51300   .Add "PSDColorsCount"
51310   .Add "PSDResolution"
51320   .Add "PSLanguageLevel"
51330   .Add "RAWColorsCount"
51340   .Add "RAWResolution"
51350   .Add "RemoveAllKnownFileExtensions"
51360   .Add "RemoveSpaces"
51370   .Add "RunProgramAfterSaving"
51380   .Add "RunProgramAfterSavingProgramname"
51390   .Add "RunProgramAfterSavingProgramParameters"
51400   .Add "RunProgramAfterSavingWaitUntilReady"
51410   .Add "RunProgramAfterSavingWindowstyle"
51420   .Add "RunProgramBeforeSaving"
51430   .Add "RunProgramBeforeSavingProgramname"
51440   .Add "RunProgramBeforeSavingProgramParameters"
51450   .Add "RunProgramBeforeSavingWindowstyle"
51460   .Add "SaveFilename"
51470   .Add "SendEmailAfterAutoSaving"
51480   .Add "SendMailMethod"
51490   .Add "ShowAnimation"
51500   .Add "StampFontColor"
51510   .Add "StampFontname"
51520   .Add "StampFontsize"
51530   .Add "StampOutlineFontthickness"
51540   .Add "StampString"
51550   .Add "StampUseOutlineFont"
51560   .Add "StandardAuthor"
51570   .Add "StandardCreationdate"
51580   .Add "StandardDateformat"
51590   .Add "StandardKeywords"
51600   .Add "StandardMailDomain"
51610   .Add "StandardModifydate"
51620   .Add "StandardSaveformat"
51630   .Add "StandardSubject"
51640   .Add "StandardTitle"
51650   .Add "StartStandardProgram"
51660   .Add "SVGResolution"
51670   .Add "TIFFColorscount"
51680   .Add "TIFFResolution"
51690   .Add "Toolbars"
51700   .Add "UpdateInterval"
51710   .Add "UseAutosave"
51720   .Add "UseAutosaveDirectory"
51730   .Add "UseCreationDateNow"
51740   .Add "UseCustomPaperSize"
51750   .Add "UseFixPapersize"
51760   .Add "UseStandardAuthor"
51770  End With
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
Public Sub cAddPrintjob(ByVal filename As String) ' Supported file types: Postscript, JPEG, BMP
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, Ext As String, tFilename As String, ivgf As Boolean
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Sub
50050  End If
50060  If IsValidPath(filename) = False Then
50070   mError.Number = 6
50080   mError.Description = Replace$(ErrDescr6, "%1", filename)
50090   RaiseEvent eError
50100   Exit Sub
50110  End If
50120  If FileExists(filename) = False Then
50130   mError.Number = 5
50140   mError.Description = Replace$(ErrDescr5, "%1", filename)
50150   RaiseEvent eError
50160   Exit Sub
50170  End If
50180  ivgf = IsValidGraphicFile(filename)
50190  If IsPostscriptFile(filename) Or ivgf Then
50200    tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
50210    If ivgf Then
50220      Image2PS filename, tFilename
50230     Else
50240      FileCopy filename, tFilename
50250    End If
50260   Else
50270    mError.Number = 9
50280    mError.Description = Replace$(ErrDescr9, "%1", filename)
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
 ShutDown = True
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
50060   mForm.DocumentCombineAll
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
50050   Case "ALLOWSPECIALGSCHARSINFILENAMES": cGetOptionFromOptions = Options.AllowSpecialGSCharsInFilenames
50060   Case "AUTOSAVEDIRECTORY": cGetOptionFromOptions = Options.AutosaveDirectory
50070   Case "AUTOSAVEFILENAME": cGetOptionFromOptions = Options.AutosaveFilename
50080   Case "AUTOSAVEFORMAT": cGetOptionFromOptions = Options.AutosaveFormat
50090   Case "AUTOSAVESTARTSTANDARDPROGRAM": cGetOptionFromOptions = Options.AutosaveStartStandardProgram
50100   Case "BMPCOLORSCOUNT": cGetOptionFromOptions = Options.BMPColorscount
50110   Case "BMPRESOLUTION": cGetOptionFromOptions = Options.BMPResolution
50120   Case "CLIENTCOMPUTERRESOLVEIPADDRESS": cGetOptionFromOptions = Options.ClientComputerResolveIPAddress
50130   Case "COUNTER": cGetOptionFromOptions = Options.Counter
50140   Case "DEVICEHEIGHTPOINTS": cGetOptionFromOptions = Options.DeviceHeightPoints
50150   Case "DEVICEWIDTHPOINTS": cGetOptionFromOptions = Options.DeviceWidthPoints
50160   Case "DIRECTORYGHOSTSCRIPTBINARIES": cGetOptionFromOptions = Options.DirectoryGhostscriptBinaries
50170   Case "DIRECTORYGHOSTSCRIPTFONTS": cGetOptionFromOptions = Options.DirectoryGhostscriptFonts
50180   Case "DIRECTORYGHOSTSCRIPTLIBRARIES": cGetOptionFromOptions = Options.DirectoryGhostscriptLibraries
50190   Case "DIRECTORYGHOSTSCRIPTRESOURCE": cGetOptionFromOptions = Options.DirectoryGhostscriptResource
50200   Case "DISABLEEMAIL": cGetOptionFromOptions = Options.DisableEmail
50210   Case "DONTUSEDOCUMENTSETTINGS": cGetOptionFromOptions = Options.DontUseDocumentSettings
50220   Case "EPSLANGUAGELEVEL": cGetOptionFromOptions = Options.EPSLanguageLevel
50230   Case "FILENAMESUBSTITUTIONS": cGetOptionFromOptions = Options.FilenameSubstitutions
50240   Case "FILENAMESUBSTITUTIONSONLYINTITLE": cGetOptionFromOptions = Options.FilenameSubstitutionsOnlyInTitle
50250   Case "JPEGCOLORSCOUNT": cGetOptionFromOptions = Options.JPEGColorscount
50260   Case "JPEGQUALITY": cGetOptionFromOptions = Options.JPEGQuality
50270   Case "JPEGRESOLUTION": cGetOptionFromOptions = Options.JPEGResolution
50280   Case "LANGUAGE": cGetOptionFromOptions = Options.Language
50290   Case "LASTSAVEDIRECTORY": cGetOptionFromOptions = Options.LastSaveDirectory
50300   Case "LASTUPDATECHECK": cGetOptionFromOptions = Options.LastUpdateCheck
50310   Case "LOGGING": cGetOptionFromOptions = Options.Logging
50320   Case "LOGLINES": cGetOptionFromOptions = Options.LogLines
50330   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": cGetOptionFromOptions = Options.NoConfirmMessageSwitchingDefaultprinter
50340   Case "NOPROCESSINGATSTARTUP": cGetOptionFromOptions = Options.NoProcessingAtStartup
50350   Case "NOPSCHECK": cGetOptionFromOptions = Options.NoPSCheck
50360   Case "ONEPAGEPERFILE": cGetOptionFromOptions = Options.OnePagePerFile
50370   Case "OPTIONSDESIGN": cGetOptionFromOptions = Options.OptionsDesign
50380   Case "OPTIONSENABLED": cGetOptionFromOptions = Options.OptionsEnabled
50390   Case "OPTIONSVISIBLE": cGetOptionFromOptions = Options.OptionsVisible
50400   Case "PAPERSIZE": cGetOptionFromOptions = Options.Papersize
50410   Case "PCLCOLORSCOUNT": cGetOptionFromOptions = Options.PCLColorsCount
50420   Case "PCLRESOLUTION": cGetOptionFromOptions = Options.PCLResolution
50430   Case "PCXCOLORSCOUNT": cGetOptionFromOptions = Options.PCXColorscount
50440   Case "PCXRESOLUTION": cGetOptionFromOptions = Options.PCXResolution
50450   Case "PDFALLOWASSEMBLY": cGetOptionFromOptions = Options.PDFAllowAssembly
50460   Case "PDFALLOWDEGRADEDPRINTING": cGetOptionFromOptions = Options.PDFAllowDegradedPrinting
50470   Case "PDFALLOWFILLIN": cGetOptionFromOptions = Options.PDFAllowFillIn
50480   Case "PDFALLOWSCREENREADERS": cGetOptionFromOptions = Options.PDFAllowScreenReaders
50490   Case "PDFCOLORSCMYKTORGB": cGetOptionFromOptions = Options.PDFColorsCMYKToRGB
50500   Case "PDFCOLORSCOLORMODEL": cGetOptionFromOptions = Options.PDFColorsColorModel
50510   Case "PDFCOLORSPRESERVEHALFTONE": cGetOptionFromOptions = Options.PDFColorsPreserveHalftone
50520   Case "PDFCOLORSPRESERVEOVERPRINT": cGetOptionFromOptions = Options.PDFColorsPreserveOverprint
50530   Case "PDFCOLORSPRESERVETRANSFER": cGetOptionFromOptions = Options.PDFColorsPreserveTransfer
50540   Case "PDFCOMPRESSIONCOLORCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionColorCompression
50550   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionColorCompressionChoice
50560   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGHighFactor
50570   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGLowFactor
50580   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMaximumFactor
50590   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMediumFactor
50600   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMinimumFactor
50610   Case "PDFCOMPRESSIONCOLORRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionColorResample
50620   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionColorResampleChoice
50630   Case "PDFCOMPRESSIONCOLORRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionColorResolution
50640   Case "PDFCOMPRESSIONGREYCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionGreyCompression
50650   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionChoice
50660   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGHighFactor
50670   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGLowFactor
50680   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMaximumFactor
50690   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMediumFactor
50700   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMinimumFactor
50710   Case "PDFCOMPRESSIONGREYRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionGreyResample
50720   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyResampleChoice
50730   Case "PDFCOMPRESSIONGREYRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionGreyResolution
50740   Case "PDFCOMPRESSIONMONOCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionMonoCompression
50750   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoCompressionChoice
50760   Case "PDFCOMPRESSIONMONORESAMPLE": cGetOptionFromOptions = Options.PDFCompressionMonoResample
50770   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoResampleChoice
50780   Case "PDFCOMPRESSIONMONORESOLUTION": cGetOptionFromOptions = Options.PDFCompressionMonoResolution
50790   Case "PDFCOMPRESSIONTEXTCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionTextCompression
50800   Case "PDFDISALLOWCOPY": cGetOptionFromOptions = Options.PDFDisallowCopy
50810   Case "PDFDISALLOWMODIFYANNOTATIONS": cGetOptionFromOptions = Options.PDFDisallowModifyAnnotations
50820   Case "PDFDISALLOWMODIFYCONTENTS": cGetOptionFromOptions = Options.PDFDisallowModifyContents
50830   Case "PDFDISALLOWPRINTING": cGetOptionFromOptions = Options.PDFDisallowPrinting
50840   Case "PDFENCRYPTOR": cGetOptionFromOptions = Options.PDFEncryptor
50850   Case "PDFFONTSEMBEDALL": cGetOptionFromOptions = Options.PDFFontsEmbedAll
50860   Case "PDFFONTSSUBSETFONTS": cGetOptionFromOptions = Options.PDFFontsSubSetFonts
50870   Case "PDFFONTSSUBSETFONTSPERCENT": cGetOptionFromOptions = Options.PDFFontsSubSetFontsPercent
50880   Case "PDFGENERALASCII85": cGetOptionFromOptions = Options.PDFGeneralASCII85
50890   Case "PDFGENERALAUTOROTATE": cGetOptionFromOptions = Options.PDFGeneralAutorotate
50900   Case "PDFGENERALCOMPATIBILITY": cGetOptionFromOptions = Options.PDFGeneralCompatibility
50910   Case "PDFGENERALDEFAULT": cGetOptionFromOptions = Options.PDFGeneralDefault
50920   Case "PDFGENERALOVERPRINT": cGetOptionFromOptions = Options.PDFGeneralOverprint
50930   Case "PDFGENERALRESOLUTION": cGetOptionFromOptions = Options.PDFGeneralResolution
50940   Case "PDFHIGHENCRYPTION": cGetOptionFromOptions = Options.PDFHighEncryption
50950   Case "PDFLOWENCRYPTION": cGetOptionFromOptions = Options.PDFLowEncryption
50960   Case "PDFOPTIMIZE": cGetOptionFromOptions = Options.PDFOptimize
50970   Case "PDFOWNERPASS": cGetOptionFromOptions = Options.PDFOwnerPass
50980   Case "PDFOWNERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFOwnerPasswordString
50990   Case "PDFSIGNINGMULTISIGNATURE": cGetOptionFromOptions = Options.PDFSigningMultiSignature
51000   Case "PDFSIGNINGPFXFILE": cGetOptionFromOptions = Options.PDFSigningPFXFile
51010   Case "PDFSIGNINGPFXFILEPASSWORD": cGetOptionFromOptions = Options.PDFSigningPFXFilePassword
51020   Case "PDFSIGNINGSIGNATURECONTACT": cGetOptionFromOptions = Options.PDFSigningSignatureContact
51030   Case "PDFSIGNINGSIGNATURELEFTX": cGetOptionFromOptions = Options.PDFSigningSignatureLeftX
51040   Case "PDFSIGNINGSIGNATURELEFTY": cGetOptionFromOptions = Options.PDFSigningSignatureLeftY
51050   Case "PDFSIGNINGSIGNATURELOCATION": cGetOptionFromOptions = Options.PDFSigningSignatureLocation
51060   Case "PDFSIGNINGSIGNATUREREASON": cGetOptionFromOptions = Options.PDFSigningSignatureReason
51070   Case "PDFSIGNINGSIGNATURERIGHTX": cGetOptionFromOptions = Options.PDFSigningSignatureRightX
51080   Case "PDFSIGNINGSIGNATURERIGHTY": cGetOptionFromOptions = Options.PDFSigningSignatureRightY
51090   Case "PDFSIGNINGSIGNATUREVISIBLE": cGetOptionFromOptions = Options.PDFSigningSignatureVisible
51100   Case "PDFSIGNINGSIGNPDF": cGetOptionFromOptions = Options.PDFSigningSignPDF
51110   Case "PDFUPDATEMETADATA": cGetOptionFromOptions = Options.PDFUpdateMetadata
51120   Case "PDFUSERPASS": cGetOptionFromOptions = Options.PDFUserPass
51130   Case "PDFUSERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFUserPasswordString
51140   Case "PDFUSESECURITY": cGetOptionFromOptions = Options.PDFUseSecurity
51150   Case "PNGCOLORSCOUNT": cGetOptionFromOptions = Options.PNGColorscount
51160   Case "PNGRESOLUTION": cGetOptionFromOptions = Options.PNGResolution
51170   Case "PRINTAFTERSAVING": cGetOptionFromOptions = Options.PrintAfterSaving
51180   Case "PRINTAFTERSAVINGDUPLEX": cGetOptionFromOptions = Options.PrintAfterSavingDuplex
51190   Case "PRINTAFTERSAVINGNOCANCEL": cGetOptionFromOptions = Options.PrintAfterSavingNoCancel
51200   Case "PRINTAFTERSAVINGPRINTER": cGetOptionFromOptions = Options.PrintAfterSavingPrinter
51210   Case "PRINTAFTERSAVINGQUERYUSER": cGetOptionFromOptions = Options.PrintAfterSavingQueryUser
51220   Case "PRINTAFTERSAVINGTUMBLE": cGetOptionFromOptions = Options.PrintAfterSavingTumble
51230   Case "PRINTERSTOP": cGetOptionFromOptions = Options.PrinterStop
51240   Case "PRINTERTEMPPATH": cGetOptionFromOptions = Options.PrinterTemppath
51250   Case "PROCESSPRIORITY": cGetOptionFromOptions = Options.ProcessPriority
51260   Case "PROGRAMFONT": cGetOptionFromOptions = Options.ProgramFont
51270   Case "PROGRAMFONTCHARSET": cGetOptionFromOptions = Options.ProgramFontCharset
51280   Case "PROGRAMFONTSIZE": cGetOptionFromOptions = Options.ProgramFontSize
51290   Case "PSDCOLORSCOUNT": cGetOptionFromOptions = Options.PSDColorsCount
51300   Case "PSDRESOLUTION": cGetOptionFromOptions = Options.PSDResolution
51310   Case "PSLANGUAGELEVEL": cGetOptionFromOptions = Options.PSLanguageLevel
51320   Case "RAWCOLORSCOUNT": cGetOptionFromOptions = Options.RAWColorsCount
51330   Case "RAWRESOLUTION": cGetOptionFromOptions = Options.RAWResolution
51340   Case "REMOVEALLKNOWNFILEEXTENSIONS": cGetOptionFromOptions = Options.RemoveAllKnownFileExtensions
51350   Case "REMOVESPACES": cGetOptionFromOptions = Options.RemoveSpaces
51360   Case "RUNPROGRAMAFTERSAVING": cGetOptionFromOptions = Options.RunProgramAfterSaving
51370   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramname
51380   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramParameters
51390   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": cGetOptionFromOptions = Options.RunProgramAfterSavingWaitUntilReady
51400   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramAfterSavingWindowstyle
51410   Case "RUNPROGRAMBEFORESAVING": cGetOptionFromOptions = Options.RunProgramBeforeSaving
51420   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramname
51430   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramParameters
51440   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramBeforeSavingWindowstyle
51450   Case "SAVEFILENAME": cGetOptionFromOptions = Options.SaveFilename
51460   Case "SENDEMAILAFTERAUTOSAVING": cGetOptionFromOptions = Options.SendEmailAfterAutoSaving
51470   Case "SENDMAILMETHOD": cGetOptionFromOptions = Options.SendMailMethod
51480   Case "SHOWANIMATION": cGetOptionFromOptions = Options.ShowAnimation
51490   Case "STAMPFONTCOLOR": cGetOptionFromOptions = Options.StampFontColor
51500   Case "STAMPFONTNAME": cGetOptionFromOptions = Options.StampFontname
51510   Case "STAMPFONTSIZE": cGetOptionFromOptions = Options.StampFontsize
51520   Case "STAMPOUTLINEFONTTHICKNESS": cGetOptionFromOptions = Options.StampOutlineFontthickness
51530   Case "STAMPSTRING": cGetOptionFromOptions = Options.StampString
51540   Case "STAMPUSEOUTLINEFONT": cGetOptionFromOptions = Options.StampUseOutlineFont
51550   Case "STANDARDAUTHOR": cGetOptionFromOptions = Options.StandardAuthor
51560   Case "STANDARDCREATIONDATE": cGetOptionFromOptions = Options.StandardCreationdate
51570   Case "STANDARDDATEFORMAT": cGetOptionFromOptions = Options.StandardDateformat
51580   Case "STANDARDKEYWORDS": cGetOptionFromOptions = Options.StandardKeywords
51590   Case "STANDARDMAILDOMAIN": cGetOptionFromOptions = Options.StandardMailDomain
51600   Case "STANDARDMODIFYDATE": cGetOptionFromOptions = Options.StandardModifydate
51610   Case "STANDARDSAVEFORMAT": cGetOptionFromOptions = Options.StandardSaveformat
51620   Case "STANDARDSUBJECT": cGetOptionFromOptions = Options.StandardSubject
51630   Case "STANDARDTITLE": cGetOptionFromOptions = Options.StandardTitle
51640   Case "STARTSTANDARDPROGRAM": cGetOptionFromOptions = Options.StartStandardProgram
51650   Case "SVGRESOLUTION": cGetOptionFromOptions = Options.SVGResolution
51660   Case "TIFFCOLORSCOUNT": cGetOptionFromOptions = Options.TIFFColorscount
51670   Case "TIFFRESOLUTION": cGetOptionFromOptions = Options.TIFFResolution
51680   Case "TOOLBARS": cGetOptionFromOptions = Options.Toolbars
51690   Case "UPDATEINTERVAL": cGetOptionFromOptions = Options.UpdateInterval
51700   Case "USEAUTOSAVE": cGetOptionFromOptions = Options.UseAutosave
51710   Case "USEAUTOSAVEDIRECTORY": cGetOptionFromOptions = Options.UseAutosaveDirectory
51720   Case "USECREATIONDATENOW": cGetOptionFromOptions = Options.UseCreationDateNow
51730   Case "USECUSTOMPAPERSIZE": cGetOptionFromOptions = Options.UseCustomPaperSize
51740   Case "USEFIXPAPERSIZE": cGetOptionFromOptions = Options.UseFixPapersize
51750   Case "USESTANDARDAUTHOR": cGetOptionFromOptions = Options.UseStandardAuthor
51760  End Select
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
50060   .AllowSpecialGSCharsInFilenames = sOptions.AllowSpecialGSCharsInFilenames
50070   .AutosaveDirectory = sOptions.AutosaveDirectory
50080   .AutosaveFilename = sOptions.AutosaveFilename
50090   .AutosaveFormat = sOptions.AutosaveFormat
50100   .AutosaveStartStandardProgram = sOptions.AutosaveStartStandardProgram
50110   .BMPColorscount = sOptions.BMPColorscount
50120   .BMPResolution = sOptions.BMPResolution
50130   .ClientComputerResolveIPAddress = sOptions.ClientComputerResolveIPAddress
50140   .Counter = sOptions.Counter
50150   .DeviceHeightPoints = sOptions.DeviceHeightPoints
50160   .DeviceWidthPoints = sOptions.DeviceWidthPoints
50170   .DirectoryGhostscriptBinaries = sOptions.DirectoryGhostscriptBinaries
50180   .DirectoryGhostscriptFonts = sOptions.DirectoryGhostscriptFonts
50190   .DirectoryGhostscriptLibraries = sOptions.DirectoryGhostscriptLibraries
50200   .DirectoryGhostscriptResource = sOptions.DirectoryGhostscriptResource
50210   .DisableEmail = sOptions.DisableEmail
50220   .DontUseDocumentSettings = sOptions.DontUseDocumentSettings
50230   .EPSLanguageLevel = sOptions.EPSLanguageLevel
50240   .FilenameSubstitutions = sOptions.FilenameSubstitutions
50250   .FilenameSubstitutionsOnlyInTitle = sOptions.FilenameSubstitutionsOnlyInTitle
50260   .JPEGColorscount = sOptions.JPEGColorscount
50270   .JPEGQuality = sOptions.JPEGQuality
50280   .JPEGResolution = sOptions.JPEGResolution
50290   .Language = sOptions.Language
50300   .LastSaveDirectory = sOptions.LastSaveDirectory
50310   .LastUpdateCheck = sOptions.LastUpdateCheck
50320   .Logging = sOptions.Logging
50330   .LogLines = sOptions.LogLines
50340   .NoConfirmMessageSwitchingDefaultprinter = sOptions.NoConfirmMessageSwitchingDefaultprinter
50350   .NoProcessingAtStartup = sOptions.NoProcessingAtStartup
50360   .NoPSCheck = sOptions.NoPSCheck
50370   .OnePagePerFile = sOptions.OnePagePerFile
50380   .OptionsDesign = sOptions.OptionsDesign
50390   .OptionsEnabled = sOptions.OptionsEnabled
50400   .OptionsVisible = sOptions.OptionsVisible
50410   .Papersize = sOptions.Papersize
50420   .PCLColorsCount = sOptions.PCLColorsCount
50430   .PCLResolution = sOptions.PCLResolution
50440   .PCXColorscount = sOptions.PCXColorscount
50450   .PCXResolution = sOptions.PCXResolution
50460   .PDFAllowAssembly = sOptions.PDFAllowAssembly
50470   .PDFAllowDegradedPrinting = sOptions.PDFAllowDegradedPrinting
50480   .PDFAllowFillIn = sOptions.PDFAllowFillIn
50490   .PDFAllowScreenReaders = sOptions.PDFAllowScreenReaders
50500   .PDFColorsCMYKToRGB = sOptions.PDFColorsCMYKToRGB
50510   .PDFColorsColorModel = sOptions.PDFColorsColorModel
50520   .PDFColorsPreserveHalftone = sOptions.PDFColorsPreserveHalftone
50530   .PDFColorsPreserveOverprint = sOptions.PDFColorsPreserveOverprint
50540   .PDFColorsPreserveTransfer = sOptions.PDFColorsPreserveTransfer
50550   .PDFCompressionColorCompression = sOptions.PDFCompressionColorCompression
50560   .PDFCompressionColorCompressionChoice = sOptions.PDFCompressionColorCompressionChoice
50570   .PDFCompressionColorCompressionJPEGHighFactor = sOptions.PDFCompressionColorCompressionJPEGHighFactor
50580   .PDFCompressionColorCompressionJPEGLowFactor = sOptions.PDFCompressionColorCompressionJPEGLowFactor
50590   .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
50600   .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
50610   .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
50620   .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
50630   .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
50640   .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
50650   .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
50660   .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
50670   .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
50680   .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
50690   .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
50700   .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
50710   .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
50720   .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
50730   .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
50740   .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
50750   .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
50760   .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
50770   .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
50780   .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
50790   .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
50800   .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
50810   .PDFDisallowCopy = sOptions.PDFDisallowCopy
50820   .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
50830   .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
50840   .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
50850   .PDFEncryptor = sOptions.PDFEncryptor
50860   .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
50870   .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
50880   .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
50890   .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
50900   .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
50910   .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
50920   .PDFGeneralDefault = sOptions.PDFGeneralDefault
50930   .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
50940   .PDFGeneralResolution = sOptions.PDFGeneralResolution
50950   .PDFHighEncryption = sOptions.PDFHighEncryption
50960   .PDFLowEncryption = sOptions.PDFLowEncryption
50970   .PDFOptimize = sOptions.PDFOptimize
50980   .PDFOwnerPass = sOptions.PDFOwnerPass
50990   .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
51000   .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
51010   .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
51020   .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
51030   .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
51040   .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
51050   .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
51060   .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
51070   .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
51080   .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
51090   .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
51100   .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
51110   .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
51120   .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
51130   .PDFUserPass = sOptions.PDFUserPass
51140   .PDFUserPasswordString = sOptions.PDFUserPasswordString
51150   .PDFUseSecurity = sOptions.PDFUseSecurity
51160   .PNGColorscount = sOptions.PNGColorscount
51170   .PNGResolution = sOptions.PNGResolution
51180   .PrintAfterSaving = sOptions.PrintAfterSaving
51190   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
51200   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
51210   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
51220   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51230   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51240   .PrinterStop = sOptions.PrinterStop
51250   .PrinterTemppath = sOptions.PrinterTemppath
51260   .ProcessPriority = sOptions.ProcessPriority
51270   .ProgramFont = sOptions.ProgramFont
51280   .ProgramFontCharset = sOptions.ProgramFontCharset
51290   .ProgramFontSize = sOptions.ProgramFontSize
51300   .PSDColorsCount = sOptions.PSDColorsCount
51310   .PSDResolution = sOptions.PSDResolution
51320   .PSLanguageLevel = sOptions.PSLanguageLevel
51330   .RAWColorsCount = sOptions.RAWColorsCount
51340   .RAWResolution = sOptions.RAWResolution
51350   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51360   .RemoveSpaces = sOptions.RemoveSpaces
51370   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51380   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51390   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51400   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51410   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51420   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51430   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51440   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51450   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51460   .SaveFilename = sOptions.SaveFilename
51470   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51480   .SendMailMethod = sOptions.SendMailMethod
51490   .ShowAnimation = sOptions.ShowAnimation
51500   .StampFontColor = sOptions.StampFontColor
51510   .StampFontname = sOptions.StampFontname
51520   .StampFontsize = sOptions.StampFontsize
51530   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51540   .StampString = sOptions.StampString
51550   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51560   .StandardAuthor = sOptions.StandardAuthor
51570   .StandardCreationdate = sOptions.StandardCreationdate
51580   .StandardDateformat = sOptions.StandardDateformat
51590   .StandardKeywords = sOptions.StandardKeywords
51600   .StandardMailDomain = sOptions.StandardMailDomain
51610   .StandardModifydate = sOptions.StandardModifydate
51620   .StandardSaveformat = sOptions.StandardSaveformat
51630   .StandardSubject = sOptions.StandardSubject
51640   .StandardTitle = sOptions.StandardTitle
51650   .StartStandardProgram = sOptions.StartStandardProgram
51660   .SVGResolution = sOptions.SVGResolution
51670   .TIFFColorscount = sOptions.TIFFColorscount
51680   .TIFFResolution = sOptions.TIFFResolution
51690   .Toolbars = sOptions.Toolbars
51700   .UpdateInterval = sOptions.UpdateInterval
51710   .UseAutosave = sOptions.UseAutosave
51720   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51730   .UseCreationDateNow = sOptions.UseCreationDateNow
51740   .UseCustomPaperSize = sOptions.UseCustomPaperSize
51750   .UseFixPapersize = sOptions.UseFixPapersize
51760   .UseStandardAuthor = sOptions.UseStandardAuthor
51770  End With
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
50050   Case "ALLOWSPECIALGSCHARSINFILENAMES": Options.AllowSpecialGSCharsInFilenames = value
50060   Case "AUTOSAVEDIRECTORY": Options.AutosaveDirectory = value
50070   Case "AUTOSAVEFILENAME": Options.AutosaveFilename = value
50080   Case "AUTOSAVEFORMAT": Options.AutosaveFormat = value
50090   Case "AUTOSAVESTARTSTANDARDPROGRAM": Options.AutosaveStartStandardProgram = value
50100   Case "BMPCOLORSCOUNT": Options.BMPColorscount = value
50110   Case "BMPRESOLUTION": Options.BMPResolution = value
50120   Case "CLIENTCOMPUTERRESOLVEIPADDRESS": Options.ClientComputerResolveIPAddress = value
50130   Case "COUNTER": Options.Counter = value
50140   Case "DEVICEHEIGHTPOINTS": Options.DeviceHeightPoints = value
50150   Case "DEVICEWIDTHPOINTS": Options.DeviceWidthPoints = value
50160   Case "DIRECTORYGHOSTSCRIPTBINARIES": Options.DirectoryGhostscriptBinaries = value
50170   Case "DIRECTORYGHOSTSCRIPTFONTS": Options.DirectoryGhostscriptFonts = value
50180   Case "DIRECTORYGHOSTSCRIPTLIBRARIES": Options.DirectoryGhostscriptLibraries = value
50190   Case "DIRECTORYGHOSTSCRIPTRESOURCE": Options.DirectoryGhostscriptResource = value
50200   Case "DISABLEEMAIL": Options.DisableEmail = value
50210   Case "DONTUSEDOCUMENTSETTINGS": Options.DontUseDocumentSettings = value
50220   Case "EPSLANGUAGELEVEL": Options.EPSLanguageLevel = value
50230   Case "FILENAMESUBSTITUTIONS": Options.FilenameSubstitutions = value
50240   Case "FILENAMESUBSTITUTIONSONLYINTITLE": Options.FilenameSubstitutionsOnlyInTitle = value
50250   Case "JPEGCOLORSCOUNT": Options.JPEGColorscount = value
50260   Case "JPEGQUALITY": Options.JPEGQuality = value
50270   Case "JPEGRESOLUTION": Options.JPEGResolution = value
50280   Case "LANGUAGE": Options.Language = value
50290   Case "LASTSAVEDIRECTORY": Options.LastSaveDirectory = value
50300   Case "LASTUPDATECHECK": Options.LastUpdateCheck = value
50310   Case "LOGGING": Options.Logging = value
50320   Case "LOGLINES": Options.LogLines = value
50330   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": Options.NoConfirmMessageSwitchingDefaultprinter = value
50340   Case "NOPROCESSINGATSTARTUP": Options.NoProcessingAtStartup = value
50350   Case "NOPSCHECK": Options.NoPSCheck = value
50360   Case "ONEPAGEPERFILE": Options.OnePagePerFile = value
50370   Case "OPTIONSDESIGN": Options.OptionsDesign = value
50380   Case "OPTIONSENABLED": Options.OptionsEnabled = value
50390   Case "OPTIONSVISIBLE": Options.OptionsVisible = value
50400   Case "PAPERSIZE": Options.Papersize = value
50410   Case "PCLCOLORSCOUNT": Options.PCLColorsCount = value
50420   Case "PCLRESOLUTION": Options.PCLResolution = value
50430   Case "PCXCOLORSCOUNT": Options.PCXColorscount = value
50440   Case "PCXRESOLUTION": Options.PCXResolution = value
50450   Case "PDFALLOWASSEMBLY": Options.PDFAllowAssembly = value
50460   Case "PDFALLOWDEGRADEDPRINTING": Options.PDFAllowDegradedPrinting = value
50470   Case "PDFALLOWFILLIN": Options.PDFAllowFillIn = value
50480   Case "PDFALLOWSCREENREADERS": Options.PDFAllowScreenReaders = value
50490   Case "PDFCOLORSCMYKTORGB": Options.PDFColorsCMYKToRGB = value
50500   Case "PDFCOLORSCOLORMODEL": Options.PDFColorsColorModel = value
50510   Case "PDFCOLORSPRESERVEHALFTONE": Options.PDFColorsPreserveHalftone = value
50520   Case "PDFCOLORSPRESERVEOVERPRINT": Options.PDFColorsPreserveOverprint = value
50530   Case "PDFCOLORSPRESERVETRANSFER": Options.PDFColorsPreserveTransfer = value
50540   Case "PDFCOMPRESSIONCOLORCOMPRESSION": Options.PDFCompressionColorCompression = value
50550   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": Options.PDFCompressionColorCompressionChoice = value
50560   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": Options.PDFCompressionColorCompressionJPEGHighFactor = value
50570   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": Options.PDFCompressionColorCompressionJPEGLowFactor = value
50580   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": Options.PDFCompressionColorCompressionJPEGMaximumFactor = value
50590   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": Options.PDFCompressionColorCompressionJPEGMediumFactor = value
50600   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": Options.PDFCompressionColorCompressionJPEGMinimumFactor = value
50610   Case "PDFCOMPRESSIONCOLORRESAMPLE": Options.PDFCompressionColorResample = value
50620   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": Options.PDFCompressionColorResampleChoice = value
50630   Case "PDFCOMPRESSIONCOLORRESOLUTION": Options.PDFCompressionColorResolution = value
50640   Case "PDFCOMPRESSIONGREYCOMPRESSION": Options.PDFCompressionGreyCompression = value
50650   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": Options.PDFCompressionGreyCompressionChoice = value
50660   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": Options.PDFCompressionGreyCompressionJPEGHighFactor = value
50670   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": Options.PDFCompressionGreyCompressionJPEGLowFactor = value
50680   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMaximumFactor = value
50690   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMediumFactor = value
50700   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": Options.PDFCompressionGreyCompressionJPEGMinimumFactor = value
50710   Case "PDFCOMPRESSIONGREYRESAMPLE": Options.PDFCompressionGreyResample = value
50720   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": Options.PDFCompressionGreyResampleChoice = value
50730   Case "PDFCOMPRESSIONGREYRESOLUTION": Options.PDFCompressionGreyResolution = value
50740   Case "PDFCOMPRESSIONMONOCOMPRESSION": Options.PDFCompressionMonoCompression = value
50750   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": Options.PDFCompressionMonoCompressionChoice = value
50760   Case "PDFCOMPRESSIONMONORESAMPLE": Options.PDFCompressionMonoResample = value
50770   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": Options.PDFCompressionMonoResampleChoice = value
50780   Case "PDFCOMPRESSIONMONORESOLUTION": Options.PDFCompressionMonoResolution = value
50790   Case "PDFCOMPRESSIONTEXTCOMPRESSION": Options.PDFCompressionTextCompression = value
50800   Case "PDFDISALLOWCOPY": Options.PDFDisallowCopy = value
50810   Case "PDFDISALLOWMODIFYANNOTATIONS": Options.PDFDisallowModifyAnnotations = value
50820   Case "PDFDISALLOWMODIFYCONTENTS": Options.PDFDisallowModifyContents = value
50830   Case "PDFDISALLOWPRINTING": Options.PDFDisallowPrinting = value
50840   Case "PDFENCRYPTOR": Options.PDFEncryptor = value
50850   Case "PDFFONTSEMBEDALL": Options.PDFFontsEmbedAll = value
50860   Case "PDFFONTSSUBSETFONTS": Options.PDFFontsSubSetFonts = value
50870   Case "PDFFONTSSUBSETFONTSPERCENT": Options.PDFFontsSubSetFontsPercent = value
50880   Case "PDFGENERALASCII85": Options.PDFGeneralASCII85 = value
50890   Case "PDFGENERALAUTOROTATE": Options.PDFGeneralAutorotate = value
50900   Case "PDFGENERALCOMPATIBILITY": Options.PDFGeneralCompatibility = value
50910   Case "PDFGENERALDEFAULT": Options.PDFGeneralDefault = value
50920   Case "PDFGENERALOVERPRINT": Options.PDFGeneralOverprint = value
50930   Case "PDFGENERALRESOLUTION": Options.PDFGeneralResolution = value
50940   Case "PDFHIGHENCRYPTION": Options.PDFHighEncryption = value
50950   Case "PDFLOWENCRYPTION": Options.PDFLowEncryption = value
50960   Case "PDFOPTIMIZE": Options.PDFOptimize = value
50970   Case "PDFOWNERPASS": Options.PDFOwnerPass = value
50980   Case "PDFOWNERPASSWORDSTRING": Options.PDFOwnerPasswordString = value
50990   Case "PDFSIGNINGMULTISIGNATURE": Options.PDFSigningMultiSignature = value
51000   Case "PDFSIGNINGPFXFILE": Options.PDFSigningPFXFile = value
51010   Case "PDFSIGNINGPFXFILEPASSWORD": Options.PDFSigningPFXFilePassword = value
51020   Case "PDFSIGNINGSIGNATURECONTACT": Options.PDFSigningSignatureContact = value
51030   Case "PDFSIGNINGSIGNATURELEFTX": Options.PDFSigningSignatureLeftX = value
51040   Case "PDFSIGNINGSIGNATURELEFTY": Options.PDFSigningSignatureLeftY = value
51050   Case "PDFSIGNINGSIGNATURELOCATION": Options.PDFSigningSignatureLocation = value
51060   Case "PDFSIGNINGSIGNATUREREASON": Options.PDFSigningSignatureReason = value
51070   Case "PDFSIGNINGSIGNATURERIGHTX": Options.PDFSigningSignatureRightX = value
51080   Case "PDFSIGNINGSIGNATURERIGHTY": Options.PDFSigningSignatureRightY = value
51090   Case "PDFSIGNINGSIGNATUREVISIBLE": Options.PDFSigningSignatureVisible = value
51100   Case "PDFSIGNINGSIGNPDF": Options.PDFSigningSignPDF = value
51110   Case "PDFUPDATEMETADATA": Options.PDFUpdateMetadata = value
51120   Case "PDFUSERPASS": Options.PDFUserPass = value
51130   Case "PDFUSERPASSWORDSTRING": Options.PDFUserPasswordString = value
51140   Case "PDFUSESECURITY": Options.PDFUseSecurity = value
51150   Case "PNGCOLORSCOUNT": Options.PNGColorscount = value
51160   Case "PNGRESOLUTION": Options.PNGResolution = value
51170   Case "PRINTAFTERSAVING": Options.PrintAfterSaving = value
51180   Case "PRINTAFTERSAVINGDUPLEX": Options.PrintAfterSavingDuplex = value
51190   Case "PRINTAFTERSAVINGNOCANCEL": Options.PrintAfterSavingNoCancel = value
51200   Case "PRINTAFTERSAVINGPRINTER": Options.PrintAfterSavingPrinter = value
51210   Case "PRINTAFTERSAVINGQUERYUSER": Options.PrintAfterSavingQueryUser = value
51220   Case "PRINTAFTERSAVINGTUMBLE": Options.PrintAfterSavingTumble = value
51230   Case "PRINTERSTOP": Options.PrinterStop = value
51240   Case "PRINTERTEMPPATH": Options.PrinterTemppath = value
51250   Case "PROCESSPRIORITY": Options.ProcessPriority = value
51260   Case "PROGRAMFONT": Options.ProgramFont = value
51270   Case "PROGRAMFONTCHARSET": Options.ProgramFontCharset = value
51280   Case "PROGRAMFONTSIZE": Options.ProgramFontSize = value
51290   Case "PSDCOLORSCOUNT": Options.PSDColorsCount = value
51300   Case "PSDRESOLUTION": Options.PSDResolution = value
51310   Case "PSLANGUAGELEVEL": Options.PSLanguageLevel = value
51320   Case "RAWCOLORSCOUNT": Options.RAWColorsCount = value
51330   Case "RAWRESOLUTION": Options.RAWResolution = value
51340   Case "REMOVEALLKNOWNFILEEXTENSIONS": Options.RemoveAllKnownFileExtensions = value
51350   Case "REMOVESPACES": Options.RemoveSpaces = value
51360   Case "RUNPROGRAMAFTERSAVING": Options.RunProgramAfterSaving = value
51370   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": Options.RunProgramAfterSavingProgramname = value
51380   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": Options.RunProgramAfterSavingProgramParameters = value
51390   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": Options.RunProgramAfterSavingWaitUntilReady = value
51400   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": Options.RunProgramAfterSavingWindowstyle = value
51410   Case "RUNPROGRAMBEFORESAVING": Options.RunProgramBeforeSaving = value
51420   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": Options.RunProgramBeforeSavingProgramname = value
51430   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": Options.RunProgramBeforeSavingProgramParameters = value
51440   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": Options.RunProgramBeforeSavingWindowstyle = value
51450   Case "SAVEFILENAME": Options.SaveFilename = value
51460   Case "SENDEMAILAFTERAUTOSAVING": Options.SendEmailAfterAutoSaving = value
51470   Case "SENDMAILMETHOD": Options.SendMailMethod = value
51480   Case "SHOWANIMATION": Options.ShowAnimation = value
51490   Case "STAMPFONTCOLOR": Options.StampFontColor = value
51500   Case "STAMPFONTNAME": Options.StampFontname = value
51510   Case "STAMPFONTSIZE": Options.StampFontsize = value
51520   Case "STAMPOUTLINEFONTTHICKNESS": Options.StampOutlineFontthickness = value
51530   Case "STAMPSTRING": Options.StampString = value
51540   Case "STAMPUSEOUTLINEFONT": Options.StampUseOutlineFont = value
51550   Case "STANDARDAUTHOR": Options.StandardAuthor = value
51560   Case "STANDARDCREATIONDATE": Options.StandardCreationdate = value
51570   Case "STANDARDDATEFORMAT": Options.StandardDateformat = value
51580   Case "STANDARDKEYWORDS": Options.StandardKeywords = value
51590   Case "STANDARDMAILDOMAIN": Options.StandardMailDomain = value
51600   Case "STANDARDMODIFYDATE": Options.StandardModifydate = value
51610   Case "STANDARDSAVEFORMAT": Options.StandardSaveformat = value
51620   Case "STANDARDSUBJECT": Options.StandardSubject = value
51630   Case "STANDARDTITLE": Options.StandardTitle = value
51640   Case "STARTSTANDARDPROGRAM": Options.StartStandardProgram = value
51650   Case "SVGRESOLUTION": Options.SVGResolution = value
51660   Case "TIFFCOLORSCOUNT": Options.TIFFColorscount = value
51670   Case "TIFFRESOLUTION": Options.TIFFResolution = value
51680   Case "TOOLBARS": Options.Toolbars = value
51690   Case "UPDATEINTERVAL": Options.UpdateInterval = value
51700   Case "USEAUTOSAVE": Options.UseAutosave = value
51710   Case "USEAUTOSAVEDIRECTORY": Options.UseAutosaveDirectory = value
51720   Case "USECREATIONDATENOW": Options.UseCreationDateNow = value
51730   Case "USECUSTOMPAPERSIZE": Options.UseCustomPaperSize = value
51740   Case "USEFIXPAPERSIZE": Options.UseFixPapersize = value
51750   Case "USESTANDARDAUTHOR": Options.UseStandardAuthor = value
51760   Case Else:
51770    mError.Number = 3
51780    mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in cLetOption")
51790    RaiseEvent eError
51800  End Select
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
50050   .AllowSpecialGSCharsInFilenames = sOptions.AllowSpecialGSCharsInFilenames
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
50300   .LastUpdateCheck = sOptions.LastUpdateCheck
50310   .Logging = sOptions.Logging
50320   .LogLines = sOptions.LogLines
50330   .NoConfirmMessageSwitchingDefaultprinter = sOptions.NoConfirmMessageSwitchingDefaultprinter
50340   .NoProcessingAtStartup = sOptions.NoProcessingAtStartup
50350   .NoPSCheck = sOptions.NoPSCheck
50360   .OnePagePerFile = sOptions.OnePagePerFile
50370   .OptionsDesign = sOptions.OptionsDesign
50380   .OptionsEnabled = sOptions.OptionsEnabled
50390   .OptionsVisible = sOptions.OptionsVisible
50400   .Papersize = sOptions.Papersize
50410   .PCLColorsCount = sOptions.PCLColorsCount
50420   .PCLResolution = sOptions.PCLResolution
50430   .PCXColorscount = sOptions.PCXColorscount
50440   .PCXResolution = sOptions.PCXResolution
50450   .PDFAllowAssembly = sOptions.PDFAllowAssembly
50460   .PDFAllowDegradedPrinting = sOptions.PDFAllowDegradedPrinting
50470   .PDFAllowFillIn = sOptions.PDFAllowFillIn
50480   .PDFAllowScreenReaders = sOptions.PDFAllowScreenReaders
50490   .PDFColorsCMYKToRGB = sOptions.PDFColorsCMYKToRGB
50500   .PDFColorsColorModel = sOptions.PDFColorsColorModel
50510   .PDFColorsPreserveHalftone = sOptions.PDFColorsPreserveHalftone
50520   .PDFColorsPreserveOverprint = sOptions.PDFColorsPreserveOverprint
50530   .PDFColorsPreserveTransfer = sOptions.PDFColorsPreserveTransfer
50540   .PDFCompressionColorCompression = sOptions.PDFCompressionColorCompression
50550   .PDFCompressionColorCompressionChoice = sOptions.PDFCompressionColorCompressionChoice
50560   .PDFCompressionColorCompressionJPEGHighFactor = sOptions.PDFCompressionColorCompressionJPEGHighFactor
50570   .PDFCompressionColorCompressionJPEGLowFactor = sOptions.PDFCompressionColorCompressionJPEGLowFactor
50580   .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
50590   .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
50600   .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
50610   .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
50620   .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
50630   .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
50640   .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
50650   .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
50660   .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
50670   .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
50680   .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
50690   .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
50700   .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
50710   .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
50720   .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
50730   .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
50740   .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
50750   .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
50760   .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
50770   .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
50780   .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
50790   .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
50800   .PDFDisallowCopy = sOptions.PDFDisallowCopy
50810   .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
50820   .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
50830   .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
50840   .PDFEncryptor = sOptions.PDFEncryptor
50850   .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
50860   .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
50870   .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
50880   .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
50890   .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
50900   .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
50910   .PDFGeneralDefault = sOptions.PDFGeneralDefault
50920   .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
50930   .PDFGeneralResolution = sOptions.PDFGeneralResolution
50940   .PDFHighEncryption = sOptions.PDFHighEncryption
50950   .PDFLowEncryption = sOptions.PDFLowEncryption
50960   .PDFOptimize = sOptions.PDFOptimize
50970   .PDFOwnerPass = sOptions.PDFOwnerPass
50980   .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
50990   .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
51000   .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
51010   .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
51020   .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
51030   .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
51040   .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
51050   .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
51060   .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
51070   .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
51080   .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
51090   .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
51100   .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
51110   .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
51120   .PDFUserPass = sOptions.PDFUserPass
51130   .PDFUserPasswordString = sOptions.PDFUserPasswordString
51140   .PDFUseSecurity = sOptions.PDFUseSecurity
51150   .PNGColorscount = sOptions.PNGColorscount
51160   .PNGResolution = sOptions.PNGResolution
51170   .PrintAfterSaving = sOptions.PrintAfterSaving
51180   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
51190   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
51200   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
51210   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51220   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51230   .PrinterStop = sOptions.PrinterStop
51240   .PrinterTemppath = sOptions.PrinterTemppath
51250   .ProcessPriority = sOptions.ProcessPriority
51260   .ProgramFont = sOptions.ProgramFont
51270   .ProgramFontCharset = sOptions.ProgramFontCharset
51280   .ProgramFontSize = sOptions.ProgramFontSize
51290   .PSDColorsCount = sOptions.PSDColorsCount
51300   .PSDResolution = sOptions.PSDResolution
51310   .PSLanguageLevel = sOptions.PSLanguageLevel
51320   .RAWColorsCount = sOptions.RAWColorsCount
51330   .RAWResolution = sOptions.RAWResolution
51340   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51350   .RemoveSpaces = sOptions.RemoveSpaces
51360   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51370   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51380   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51390   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51400   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51410   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51420   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51430   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51440   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51450   .SaveFilename = sOptions.SaveFilename
51460   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51470   .SendMailMethod = sOptions.SendMailMethod
51480   .ShowAnimation = sOptions.ShowAnimation
51490   .StampFontColor = sOptions.StampFontColor
51500   .StampFontname = sOptions.StampFontname
51510   .StampFontsize = sOptions.StampFontsize
51520   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51530   .StampString = sOptions.StampString
51540   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51550   .StandardAuthor = sOptions.StandardAuthor
51560   .StandardCreationdate = sOptions.StandardCreationdate
51570   .StandardDateformat = sOptions.StandardDateformat
51580   .StandardKeywords = sOptions.StandardKeywords
51590   .StandardMailDomain = sOptions.StandardMailDomain
51600   .StandardModifydate = sOptions.StandardModifydate
51610   .StandardSaveformat = sOptions.StandardSaveformat
51620   .StandardSubject = sOptions.StandardSubject
51630   .StandardTitle = sOptions.StandardTitle
51640   .StartStandardProgram = sOptions.StartStandardProgram
51650   .SVGResolution = sOptions.SVGResolution
51660   .TIFFColorscount = sOptions.TIFFColorscount
51670   .TIFFResolution = sOptions.TIFFResolution
51680   .Toolbars = sOptions.Toolbars
51690   .UpdateInterval = sOptions.UpdateInterval
51700   .UseAutosave = sOptions.UseAutosave
51710   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51720   .UseCreationDateNow = sOptions.UseCreationDateNow
51730   .UseCustomPaperSize = sOptions.UseCustomPaperSize
51740   .UseFixPapersize = sOptions.UseFixPapersize
51750   .UseStandardAuthor = sOptions.UseStandardAuthor
51760  End With
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

