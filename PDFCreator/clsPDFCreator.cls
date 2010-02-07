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
Private Const ErrDescr11 = "The profile '%1' doesn't exists!"

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
Public Property Get cIsClosed() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Forms.Count = 0 Then
50020    cIsClosed = True
50030   Else
50040    cIsClosed = False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cIsClosed [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

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
50590   .Add "PDFCompressionColorCompressionJPEGManualFactor"
50600   .Add "PDFCompressionColorCompressionJPEGMaximumFactor"
50610   .Add "PDFCompressionColorCompressionJPEGMediumFactor"
50620   .Add "PDFCompressionColorCompressionJPEGMinimumFactor"
50630   .Add "PDFCompressionColorResample"
50640   .Add "PDFCompressionColorResampleChoice"
50650   .Add "PDFCompressionColorResolution"
50660   .Add "PDFCompressionGreyCompression"
50670   .Add "PDFCompressionGreyCompressionChoice"
50680   .Add "PDFCompressionGreyCompressionJPEGHighFactor"
50690   .Add "PDFCompressionGreyCompressionJPEGLowFactor"
50700   .Add "PDFCompressionGreyCompressionJPEGManualFactor"
50710   .Add "PDFCompressionGreyCompressionJPEGMaximumFactor"
50720   .Add "PDFCompressionGreyCompressionJPEGMediumFactor"
50730   .Add "PDFCompressionGreyCompressionJPEGMinimumFactor"
50740   .Add "PDFCompressionGreyResample"
50750   .Add "PDFCompressionGreyResampleChoice"
50760   .Add "PDFCompressionGreyResolution"
50770   .Add "PDFCompressionMonoCompression"
50780   .Add "PDFCompressionMonoCompressionChoice"
50790   .Add "PDFCompressionMonoResample"
50800   .Add "PDFCompressionMonoResampleChoice"
50810   .Add "PDFCompressionMonoResolution"
50820   .Add "PDFCompressionTextCompression"
50830   .Add "PDFDisallowCopy"
50840   .Add "PDFDisallowModifyAnnotations"
50850   .Add "PDFDisallowModifyContents"
50860   .Add "PDFDisallowPrinting"
50870   .Add "PDFEncryptor"
50880   .Add "PDFFontsEmbedAll"
50890   .Add "PDFFontsSubSetFonts"
50900   .Add "PDFFontsSubSetFontsPercent"
50910   .Add "PDFGeneralASCII85"
50920   .Add "PDFGeneralAutorotate"
50930   .Add "PDFGeneralCompatibility"
50940   .Add "PDFGeneralDefault"
50950   .Add "PDFGeneralOverprint"
50960   .Add "PDFGeneralResolution"
50970   .Add "PDFHighEncryption"
50980   .Add "PDFLowEncryption"
50990   .Add "PDFOptimize"
51000   .Add "PDFOwnerPass"
51010   .Add "PDFOwnerPasswordString"
51020   .Add "PDFSigningMultiSignature"
51030   .Add "PDFSigningPFXFile"
51040   .Add "PDFSigningPFXFilePassword"
51050   .Add "PDFSigningSignatureContact"
51060   .Add "PDFSigningSignatureLeftX"
51070   .Add "PDFSigningSignatureLeftY"
51080   .Add "PDFSigningSignatureLocation"
51090   .Add "PDFSigningSignatureOnPage"
51100   .Add "PDFSigningSignatureReason"
51110   .Add "PDFSigningSignatureRightX"
51120   .Add "PDFSigningSignatureRightY"
51130   .Add "PDFSigningSignatureVisible"
51140   .Add "PDFSigningSignPDF"
51150   .Add "PDFUpdateMetadata"
51160   .Add "PDFUserPass"
51170   .Add "PDFUserPasswordString"
51180   .Add "PDFUseSecurity"
51190   .Add "PNGColorscount"
51200   .Add "PNGResolution"
51210   .Add "PrintAfterSaving"
51220   .Add "PrintAfterSavingBitsPerPixel"
51230   .Add "PrintAfterSavingDuplex"
51240   .Add "PrintAfterSavingMaxResolution"
51250   .Add "PrintAfterSavingMaxResolutionEnabled"
51260   .Add "PrintAfterSavingNoCancel"
51270   .Add "PrintAfterSavingPrinter"
51280   .Add "PrintAfterSavingQueryUser"
51290   .Add "PrintAfterSavingTumble"
51300   .Add "PrinterStop"
51310   .Add "PrinterTemppath"
51320   .Add "ProcessPriority"
51330   .Add "ProgramFont"
51340   .Add "ProgramFontCharset"
51350   .Add "ProgramFontSize"
51360   .Add "PSDColorsCount"
51370   .Add "PSDResolution"
51380   .Add "PSLanguageLevel"
51390   .Add "RAWColorsCount"
51400   .Add "RAWResolution"
51410   .Add "RemoveAllKnownFileExtensions"
51420   .Add "RemoveSpaces"
51430   .Add "RunProgramAfterSaving"
51440   .Add "RunProgramAfterSavingProgramname"
51450   .Add "RunProgramAfterSavingProgramParameters"
51460   .Add "RunProgramAfterSavingWaitUntilReady"
51470   .Add "RunProgramAfterSavingWindowstyle"
51480   .Add "RunProgramBeforeSaving"
51490   .Add "RunProgramBeforeSavingProgramname"
51500   .Add "RunProgramBeforeSavingProgramParameters"
51510   .Add "RunProgramBeforeSavingWindowstyle"
51520   .Add "SaveFilename"
51530   .Add "SendEmailAfterAutoSaving"
51540   .Add "SendMailMethod"
51550   .Add "ShowAnimation"
51560   .Add "StampFontColor"
51570   .Add "StampFontname"
51580   .Add "StampFontsize"
51590   .Add "StampOutlineFontthickness"
51600   .Add "StampString"
51610   .Add "StampUseOutlineFont"
51620   .Add "StandardAuthor"
51630   .Add "StandardCreationdate"
51640   .Add "StandardDateformat"
51650   .Add "StandardKeywords"
51660   .Add "StandardMailDomain"
51670   .Add "StandardModifydate"
51680   .Add "StandardSaveformat"
51690   .Add "StandardSubject"
51700   .Add "StandardTitle"
51710   .Add "StartStandardProgram"
51720   .Add "SVGResolution"
51730   .Add "TIFFColorscount"
51740   .Add "TIFFResolution"
51750   .Add "Toolbars"
51760   .Add "UpdateInterval"
51770   .Add "UseAutosave"
51780   .Add "UseAutosaveDirectory"
51790   .Add "UseCreationDateNow"
51800   .Add "UseCustomPaperSize"
51810   .Add "UseFixPapersize"
51820   .Add "UseStandardAuthor"
51830  End With
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
50050  cLetOption PropertyName, value, Options, ""
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

Public Property Get cOptionProfile(ByVal ProfileName As String, ByVal PropertyName As String) As Variant
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  If LenB(Trim(ProfileName)) > 0 Then
50060   If ProfileExists(ProfileName) = False Then
50070    mError.Number = 11
50080    mError.Description = Replace$(ErrDescr11, "%1", ProfileName)
50090    RaiseEvent eError
50100    Exit Property
50110   End If
50120  End If
50130  cOptionProfile = cGetOptionFromOptions(PropertyName, ReadOptions(, , ProfileName))
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cOptionProfile [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let cOptionProfile(ByVal ProfileName As String, ByVal PropertyName As String, ByVal value As Variant)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  If LenB(Trim(ProfileName)) > 0 Then
50060   If ProfileExists(ProfileName) = False Then
50070    mError.Number = 11
50080    mError.Description = Replace$(ErrDescr11, "%1", ProfileName)
50090    RaiseEvent eError
50100    Exit Property
50110   End If
50120  End If
50130  cLetOption PropertyName, value, Options, ProfileName
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cOptionProfile [LET]")
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

Public Property Get cOptionsProfile(ByVal ProfileName As String) As clsPDFCreatorOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  If LenB(Trim(ProfileName)) > 0 Then
50060   If ProfileExists(ProfileName) = False Then
50070    mError.Number = 11
50080    mError.Description = Replace$(ErrDescr11, "%1", ProfileName)
50090    RaiseEvent eError
50100    Exit Property
50110   End If
50120  End If
50130  If LenB(ProfileName) = 0 Then
50140    Set cOptions = cGetOptions(Options)
50150   Else
50160    Set cOptions = cGetOptions(ReadOptions(, , ProfileName))
50170  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cOptionsProfile [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Set cOptionsProfile(ByVal ProfileName As String, ByVal Options1 As clsPDFCreatorOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  If LenB(Trim(ProfileName)) > 0 Then
50060   If ProfileExists(ProfileName) = False Then
50070    mError.Number = 11
50080    mError.Description = Replace$(ErrDescr11, "%1", ProfileName)
50090    RaiseEvent eError
50100    Exit Property
50110   End If
50120  End If
50130  If LenB(ProfileName) = 0 Then
50140    Options = cLetOptions(Options1)
50150   Else
50160    SaveOptions cLetOptions(Options1), ProfileName
50170  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cOptionsProfile [SET]")
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ShutDown = True
50020  Dim f As Form, i As Long
50030
50040  i = 0
50050  Do
50060   i = i + 1
50070   Sleep 100
50080   DoEvents
50090  Loop Until ProgramIsStarted = True
50100
50110  cVisible = False
50120  cClose = False
50130
50140 'MsgBox "1"
50150 '50140  eActionTimer.Enabled = False
50160 '50150  i = 0
50170 '50160  Do
50180 '50170   i = i + 1
50190 '50180   Sleep 10
50200 '50190   DoEvents
50210 '50200  Loop Until In_eActionTimer = False
50220  Set eActionTimer = Nothing
50230 'MsgBox "2"
50240
50250  For Each f In Forms
50260   If Not f Is Nothing Then
50270    If UCase$(f.Name) <> "FRMMAIN" Then
50280     Unload f
50290     cClose = True
50300    End If
50310   End If
50320  Next f
50330
50340  If Not mForm Is Nothing Then
50350   Unload mForm
50360   Set mForm = Nothing
50370  End If
50380  ProgramIsStarted = False
50390  cClose = True
50400
50410  Set mError = Nothing
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
50100
50110  'While Forms.Count > 0
50120  ' DoEvents
50130  ' Sleep 100
50140  'Wend
50150
50160  mProgramIsRunning = False
50170  StartProgram Params
50180  Set mForm = frmMain
50190  Set eActionTimer = mForm.Timer3: eActionTimer.Interval = 10
50200  eActionTimer.Enabled = True
50210  cStart = True
50220  cVisible = ProgramIsVisible
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

Public Function cReadOptions(Optional ByVal ProfileName As String = "") As clsPDFCreatorOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Function
50040  End If
50050  If LenB(Trim(ProfileName)) > 0 Then
50060   If ProfileExists(ProfileName) = False Then
50070    mError.Number = 11
50080    mError.Description = Replace$(ErrDescr11, "%1", ProfileName)
50090    RaiseEvent eError
50100    Exit Function
50110   End If
50120  End If
50130  Set cReadOptions = cGetOptions(ReadOptions(, , ProfileName))
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

Public Sub cSaveOptions(Optional ByVal Options1 As Variant, Optional ByVal ProfileName As String = "") ' Options1 must be a clsPDFCreatorOptions type
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Sub
50040  End If
50050  If Not IsMissing(Options1) Then
50060    If TypeOf Options1 Is clsPDFCreatorOptions Then
50070      SaveOptions cLetOptions(Options1), Trim$(ProfileName)
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
50180  Set cReadOptionsFromFile = cGetOptions(ReadOptionsINI(dummyOptions, INIFilename, 0, True, True))
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
50010  If In_eActionTimer Then
50020   Exit Sub
50030  End If
50040  In_eActionTimer = True
50050  If ReadyConverting Then
50060   RaiseEvent eReady
50070   ReadyConverting = False
50080  End If
50090  If GhostscriptError < 0 Then
50100   mError.Number = 4
50110   mError.Description = Replace$(ErrDescr4, "%1", CStr(GhostscriptError))
50120   RaiseEvent eError
50130   GhostscriptError = 0
50140  End If
50150 'DoEvents
50160  In_eActionTimer = False
50170  If ShutDown = True Then
50180   If Not eActionTimer Is Nothing Then
50190    eActionTimer.Enabled = False
50200   End If
50210  End If
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
50580   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMANUALFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGManualFactor
50590   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMaximumFactor
50600   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMediumFactor
50610   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMinimumFactor
50620   Case "PDFCOMPRESSIONCOLORRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionColorResample
50630   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionColorResampleChoice
50640   Case "PDFCOMPRESSIONCOLORRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionColorResolution
50650   Case "PDFCOMPRESSIONGREYCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionGreyCompression
50660   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionChoice
50670   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGHighFactor
50680   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGLowFactor
50690   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMANUALFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGManualFactor
50700   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMaximumFactor
50710   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMediumFactor
50720   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMinimumFactor
50730   Case "PDFCOMPRESSIONGREYRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionGreyResample
50740   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyResampleChoice
50750   Case "PDFCOMPRESSIONGREYRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionGreyResolution
50760   Case "PDFCOMPRESSIONMONOCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionMonoCompression
50770   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoCompressionChoice
50780   Case "PDFCOMPRESSIONMONORESAMPLE": cGetOptionFromOptions = Options.PDFCompressionMonoResample
50790   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoResampleChoice
50800   Case "PDFCOMPRESSIONMONORESOLUTION": cGetOptionFromOptions = Options.PDFCompressionMonoResolution
50810   Case "PDFCOMPRESSIONTEXTCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionTextCompression
50820   Case "PDFDISALLOWCOPY": cGetOptionFromOptions = Options.PDFDisallowCopy
50830   Case "PDFDISALLOWMODIFYANNOTATIONS": cGetOptionFromOptions = Options.PDFDisallowModifyAnnotations
50840   Case "PDFDISALLOWMODIFYCONTENTS": cGetOptionFromOptions = Options.PDFDisallowModifyContents
50850   Case "PDFDISALLOWPRINTING": cGetOptionFromOptions = Options.PDFDisallowPrinting
50860   Case "PDFENCRYPTOR": cGetOptionFromOptions = Options.PDFEncryptor
50870   Case "PDFFONTSEMBEDALL": cGetOptionFromOptions = Options.PDFFontsEmbedAll
50880   Case "PDFFONTSSUBSETFONTS": cGetOptionFromOptions = Options.PDFFontsSubSetFonts
50890   Case "PDFFONTSSUBSETFONTSPERCENT": cGetOptionFromOptions = Options.PDFFontsSubSetFontsPercent
50900   Case "PDFGENERALASCII85": cGetOptionFromOptions = Options.PDFGeneralASCII85
50910   Case "PDFGENERALAUTOROTATE": cGetOptionFromOptions = Options.PDFGeneralAutorotate
50920   Case "PDFGENERALCOMPATIBILITY": cGetOptionFromOptions = Options.PDFGeneralCompatibility
50930   Case "PDFGENERALDEFAULT": cGetOptionFromOptions = Options.PDFGeneralDefault
50940   Case "PDFGENERALOVERPRINT": cGetOptionFromOptions = Options.PDFGeneralOverprint
50950   Case "PDFGENERALRESOLUTION": cGetOptionFromOptions = Options.PDFGeneralResolution
50960   Case "PDFHIGHENCRYPTION": cGetOptionFromOptions = Options.PDFHighEncryption
50970   Case "PDFLOWENCRYPTION": cGetOptionFromOptions = Options.PDFLowEncryption
50980   Case "PDFOPTIMIZE": cGetOptionFromOptions = Options.PDFOptimize
50990   Case "PDFOWNERPASS": cGetOptionFromOptions = Options.PDFOwnerPass
51000   Case "PDFOWNERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFOwnerPasswordString
51010   Case "PDFSIGNINGMULTISIGNATURE": cGetOptionFromOptions = Options.PDFSigningMultiSignature
51020   Case "PDFSIGNINGPFXFILE": cGetOptionFromOptions = Options.PDFSigningPFXFile
51030   Case "PDFSIGNINGPFXFILEPASSWORD": cGetOptionFromOptions = Options.PDFSigningPFXFilePassword
51040   Case "PDFSIGNINGSIGNATURECONTACT": cGetOptionFromOptions = Options.PDFSigningSignatureContact
51050   Case "PDFSIGNINGSIGNATURELEFTX": cGetOptionFromOptions = Options.PDFSigningSignatureLeftX
51060   Case "PDFSIGNINGSIGNATURELEFTY": cGetOptionFromOptions = Options.PDFSigningSignatureLeftY
51070   Case "PDFSIGNINGSIGNATURELOCATION": cGetOptionFromOptions = Options.PDFSigningSignatureLocation
51080   Case "PDFSIGNINGSIGNATUREONPAGE": cGetOptionFromOptions = Options.PDFSigningSignatureOnPage
51090   Case "PDFSIGNINGSIGNATUREREASON": cGetOptionFromOptions = Options.PDFSigningSignatureReason
51100   Case "PDFSIGNINGSIGNATURERIGHTX": cGetOptionFromOptions = Options.PDFSigningSignatureRightX
51110   Case "PDFSIGNINGSIGNATURERIGHTY": cGetOptionFromOptions = Options.PDFSigningSignatureRightY
51120   Case "PDFSIGNINGSIGNATUREVISIBLE": cGetOptionFromOptions = Options.PDFSigningSignatureVisible
51130   Case "PDFSIGNINGSIGNPDF": cGetOptionFromOptions = Options.PDFSigningSignPDF
51140   Case "PDFUPDATEMETADATA": cGetOptionFromOptions = Options.PDFUpdateMetadata
51150   Case "PDFUSERPASS": cGetOptionFromOptions = Options.PDFUserPass
51160   Case "PDFUSERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFUserPasswordString
51170   Case "PDFUSESECURITY": cGetOptionFromOptions = Options.PDFUseSecurity
51180   Case "PNGCOLORSCOUNT": cGetOptionFromOptions = Options.PNGColorscount
51190   Case "PNGRESOLUTION": cGetOptionFromOptions = Options.PNGResolution
51200   Case "PRINTAFTERSAVING": cGetOptionFromOptions = Options.PrintAfterSaving
51210   Case "PRINTAFTERSAVINGBITSPERPIXEL": cGetOptionFromOptions = Options.PrintAfterSavingBitsPerPixel
51220   Case "PRINTAFTERSAVINGDUPLEX": cGetOptionFromOptions = Options.PrintAfterSavingDuplex
51230   Case "PRINTAFTERSAVINGMAXRESOLUTION": cGetOptionFromOptions = Options.PrintAfterSavingMaxResolution
51240   Case "PRINTAFTERSAVINGMAXRESOLUTIONENABLED": cGetOptionFromOptions = Options.PrintAfterSavingMaxResolutionEnabled
51250   Case "PRINTAFTERSAVINGNOCANCEL": cGetOptionFromOptions = Options.PrintAfterSavingNoCancel
51260   Case "PRINTAFTERSAVINGPRINTER": cGetOptionFromOptions = Options.PrintAfterSavingPrinter
51270   Case "PRINTAFTERSAVINGQUERYUSER": cGetOptionFromOptions = Options.PrintAfterSavingQueryUser
51280   Case "PRINTAFTERSAVINGTUMBLE": cGetOptionFromOptions = Options.PrintAfterSavingTumble
51290   Case "PRINTERSTOP": cGetOptionFromOptions = Options.PrinterStop
51300   Case "PRINTERTEMPPATH": cGetOptionFromOptions = Options.PrinterTemppath
51310   Case "PROCESSPRIORITY": cGetOptionFromOptions = Options.ProcessPriority
51320   Case "PROGRAMFONT": cGetOptionFromOptions = Options.ProgramFont
51330   Case "PROGRAMFONTCHARSET": cGetOptionFromOptions = Options.ProgramFontCharset
51340   Case "PROGRAMFONTSIZE": cGetOptionFromOptions = Options.ProgramFontSize
51350   Case "PSDCOLORSCOUNT": cGetOptionFromOptions = Options.PSDColorsCount
51360   Case "PSDRESOLUTION": cGetOptionFromOptions = Options.PSDResolution
51370   Case "PSLANGUAGELEVEL": cGetOptionFromOptions = Options.PSLanguageLevel
51380   Case "RAWCOLORSCOUNT": cGetOptionFromOptions = Options.RAWColorsCount
51390   Case "RAWRESOLUTION": cGetOptionFromOptions = Options.RAWResolution
51400   Case "REMOVEALLKNOWNFILEEXTENSIONS": cGetOptionFromOptions = Options.RemoveAllKnownFileExtensions
51410   Case "REMOVESPACES": cGetOptionFromOptions = Options.RemoveSpaces
51420   Case "RUNPROGRAMAFTERSAVING": cGetOptionFromOptions = Options.RunProgramAfterSaving
51430   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramname
51440   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramParameters
51450   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": cGetOptionFromOptions = Options.RunProgramAfterSavingWaitUntilReady
51460   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramAfterSavingWindowstyle
51470   Case "RUNPROGRAMBEFORESAVING": cGetOptionFromOptions = Options.RunProgramBeforeSaving
51480   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramname
51490   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramParameters
51500   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramBeforeSavingWindowstyle
51510   Case "SAVEFILENAME": cGetOptionFromOptions = Options.SaveFilename
51520   Case "SENDEMAILAFTERAUTOSAVING": cGetOptionFromOptions = Options.SendEmailAfterAutoSaving
51530   Case "SENDMAILMETHOD": cGetOptionFromOptions = Options.SendMailMethod
51540   Case "SHOWANIMATION": cGetOptionFromOptions = Options.ShowAnimation
51550   Case "STAMPFONTCOLOR": cGetOptionFromOptions = Options.StampFontColor
51560   Case "STAMPFONTNAME": cGetOptionFromOptions = Options.StampFontname
51570   Case "STAMPFONTSIZE": cGetOptionFromOptions = Options.StampFontsize
51580   Case "STAMPOUTLINEFONTTHICKNESS": cGetOptionFromOptions = Options.StampOutlineFontthickness
51590   Case "STAMPSTRING": cGetOptionFromOptions = Options.StampString
51600   Case "STAMPUSEOUTLINEFONT": cGetOptionFromOptions = Options.StampUseOutlineFont
51610   Case "STANDARDAUTHOR": cGetOptionFromOptions = Options.StandardAuthor
51620   Case "STANDARDCREATIONDATE": cGetOptionFromOptions = Options.StandardCreationdate
51630   Case "STANDARDDATEFORMAT": cGetOptionFromOptions = Options.StandardDateformat
51640   Case "STANDARDKEYWORDS": cGetOptionFromOptions = Options.StandardKeywords
51650   Case "STANDARDMAILDOMAIN": cGetOptionFromOptions = Options.StandardMailDomain
51660   Case "STANDARDMODIFYDATE": cGetOptionFromOptions = Options.StandardModifydate
51670   Case "STANDARDSAVEFORMAT": cGetOptionFromOptions = Options.StandardSaveformat
51680   Case "STANDARDSUBJECT": cGetOptionFromOptions = Options.StandardSubject
51690   Case "STANDARDTITLE": cGetOptionFromOptions = Options.StandardTitle
51700   Case "STARTSTANDARDPROGRAM": cGetOptionFromOptions = Options.StartStandardProgram
51710   Case "SVGRESOLUTION": cGetOptionFromOptions = Options.SVGResolution
51720   Case "TIFFCOLORSCOUNT": cGetOptionFromOptions = Options.TIFFColorscount
51730   Case "TIFFRESOLUTION": cGetOptionFromOptions = Options.TIFFResolution
51740   Case "TOOLBARS": cGetOptionFromOptions = Options.Toolbars
51750   Case "UPDATEINTERVAL": cGetOptionFromOptions = Options.UpdateInterval
51760   Case "USEAUTOSAVE": cGetOptionFromOptions = Options.UseAutosave
51770   Case "USEAUTOSAVEDIRECTORY": cGetOptionFromOptions = Options.UseAutosaveDirectory
51780   Case "USECREATIONDATENOW": cGetOptionFromOptions = Options.UseCreationDateNow
51790   Case "USECUSTOMPAPERSIZE": cGetOptionFromOptions = Options.UseCustomPaperSize
51800   Case "USEFIXPAPERSIZE": cGetOptionFromOptions = Options.UseFixPapersize
51810   Case "USESTANDARDAUTHOR": cGetOptionFromOptions = Options.UseStandardAuthor
51820  End Select
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
50590   .PDFCompressionColorCompressionJPEGManualFactor = sOptions.PDFCompressionColorCompressionJPEGManualFactor
50600   .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
50610   .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
50620   .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
50630   .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
50640   .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
50650   .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
50660   .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
50670   .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
50680   .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
50690   .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
50700   .PDFCompressionGreyCompressionJPEGManualFactor = sOptions.PDFCompressionGreyCompressionJPEGManualFactor
50710   .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
50720   .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
50730   .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
50740   .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
50750   .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
50760   .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
50770   .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
50780   .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
50790   .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
50800   .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
50810   .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
50820   .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
50830   .PDFDisallowCopy = sOptions.PDFDisallowCopy
50840   .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
50850   .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
50860   .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
50870   .PDFEncryptor = sOptions.PDFEncryptor
50880   .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
50890   .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
50900   .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
50910   .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
50920   .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
50930   .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
50940   .PDFGeneralDefault = sOptions.PDFGeneralDefault
50950   .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
50960   .PDFGeneralResolution = sOptions.PDFGeneralResolution
50970   .PDFHighEncryption = sOptions.PDFHighEncryption
50980   .PDFLowEncryption = sOptions.PDFLowEncryption
50990   .PDFOptimize = sOptions.PDFOptimize
51000   .PDFOwnerPass = sOptions.PDFOwnerPass
51010   .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
51020   .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
51030   .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
51040   .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
51050   .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
51060   .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
51070   .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
51080   .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
51090   .PDFSigningSignatureOnPage = sOptions.PDFSigningSignatureOnPage
51100   .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
51110   .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
51120   .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
51130   .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
51140   .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
51150   .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
51160   .PDFUserPass = sOptions.PDFUserPass
51170   .PDFUserPasswordString = sOptions.PDFUserPasswordString
51180   .PDFUseSecurity = sOptions.PDFUseSecurity
51190   .PNGColorscount = sOptions.PNGColorscount
51200   .PNGResolution = sOptions.PNGResolution
51210   .PrintAfterSaving = sOptions.PrintAfterSaving
51220   .PrintAfterSavingBitsPerPixel = sOptions.PrintAfterSavingBitsPerPixel
51230   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
51240   .PrintAfterSavingMaxResolution = sOptions.PrintAfterSavingMaxResolution
51250   .PrintAfterSavingMaxResolutionEnabled = sOptions.PrintAfterSavingMaxResolutionEnabled
51260   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
51270   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
51280   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51290   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51300   .PrinterStop = sOptions.PrinterStop
51310   .PrinterTemppath = sOptions.PrinterTemppath
51320   .ProcessPriority = sOptions.ProcessPriority
51330   .ProgramFont = sOptions.ProgramFont
51340   .ProgramFontCharset = sOptions.ProgramFontCharset
51350   .ProgramFontSize = sOptions.ProgramFontSize
51360   .PSDColorsCount = sOptions.PSDColorsCount
51370   .PSDResolution = sOptions.PSDResolution
51380   .PSLanguageLevel = sOptions.PSLanguageLevel
51390   .RAWColorsCount = sOptions.RAWColorsCount
51400   .RAWResolution = sOptions.RAWResolution
51410   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51420   .RemoveSpaces = sOptions.RemoveSpaces
51430   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51440   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51450   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51460   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51470   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51480   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51490   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51500   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51510   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51520   .SaveFilename = sOptions.SaveFilename
51530   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51540   .SendMailMethod = sOptions.SendMailMethod
51550   .ShowAnimation = sOptions.ShowAnimation
51560   .StampFontColor = sOptions.StampFontColor
51570   .StampFontname = sOptions.StampFontname
51580   .StampFontsize = sOptions.StampFontsize
51590   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51600   .StampString = sOptions.StampString
51610   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51620   .StandardAuthor = sOptions.StandardAuthor
51630   .StandardCreationdate = sOptions.StandardCreationdate
51640   .StandardDateformat = sOptions.StandardDateformat
51650   .StandardKeywords = sOptions.StandardKeywords
51660   .StandardMailDomain = sOptions.StandardMailDomain
51670   .StandardModifydate = sOptions.StandardModifydate
51680   .StandardSaveformat = sOptions.StandardSaveformat
51690   .StandardSubject = sOptions.StandardSubject
51700   .StandardTitle = sOptions.StandardTitle
51710   .StartStandardProgram = sOptions.StartStandardProgram
51720   .SVGResolution = sOptions.SVGResolution
51730   .TIFFColorscount = sOptions.TIFFColorscount
51740   .TIFFResolution = sOptions.TIFFResolution
51750   .Toolbars = sOptions.Toolbars
51760   .UpdateInterval = sOptions.UpdateInterval
51770   .UseAutosave = sOptions.UseAutosave
51780   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51790   .UseCreationDateNow = sOptions.UseCreationDateNow
51800   .UseCustomPaperSize = sOptions.UseCustomPaperSize
51810   .UseFixPapersize = sOptions.UseFixPapersize
51820   .UseStandardAuthor = sOptions.UseStandardAuthor
51830  End With
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

Private Sub cLetOption(PropertyName As String, value As Variant, Options1 As tOptions, ProfileName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case UCase$(PropertyName)
        Case "ADDITIONALGHOSTSCRIPTPARAMETERS": Options1.AdditionalGhostscriptParameters = value
50030   Case "ADDITIONALGHOSTSCRIPTSEARCHPATH": Options1.AdditionalGhostscriptSearchpath = value
50040   Case "ADDWINDOWSFONTPATH": Options1.AddWindowsFontpath = value
50050   Case "ALLOWSPECIALGSCHARSINFILENAMES": Options1.AllowSpecialGSCharsInFilenames = value
50060   Case "AUTOSAVEDIRECTORY": Options1.AutosaveDirectory = value
50070   Case "AUTOSAVEFILENAME": Options1.AutosaveFilename = value
50080   Case "AUTOSAVEFORMAT": Options1.AutosaveFormat = value
50090   Case "AUTOSAVESTARTSTANDARDPROGRAM": Options1.AutosaveStartStandardProgram = value
50100   Case "BMPCOLORSCOUNT": Options1.BMPColorscount = value
50110   Case "BMPRESOLUTION": Options1.BMPResolution = value
50120   Case "CLIENTCOMPUTERRESOLVEIPADDRESS": Options1.ClientComputerResolveIPAddress = value
50130   Case "COUNTER": Options1.Counter = value
50140   Case "DEVICEHEIGHTPOINTS": Options1.DeviceHeightPoints = value
50150   Case "DEVICEWIDTHPOINTS": Options1.DeviceWidthPoints = value
50160   Case "DIRECTORYGHOSTSCRIPTBINARIES": Options1.DirectoryGhostscriptBinaries = value
50170   Case "DIRECTORYGHOSTSCRIPTFONTS": Options1.DirectoryGhostscriptFonts = value
50180   Case "DIRECTORYGHOSTSCRIPTLIBRARIES": Options1.DirectoryGhostscriptLibraries = value
50190   Case "DIRECTORYGHOSTSCRIPTRESOURCE": Options1.DirectoryGhostscriptResource = value
50200   Case "DISABLEEMAIL": Options1.DisableEmail = value
50210   Case "DONTUSEDOCUMENTSETTINGS": Options1.DontUseDocumentSettings = value
50220   Case "EPSLANGUAGELEVEL": Options1.EPSLanguageLevel = value
50230   Case "FILENAMESUBSTITUTIONS": Options1.FilenameSubstitutions = value
50240   Case "FILENAMESUBSTITUTIONSONLYINTITLE": Options1.FilenameSubstitutionsOnlyInTitle = value
50250   Case "JPEGCOLORSCOUNT": Options1.JPEGColorscount = value
50260   Case "JPEGQUALITY": Options1.JPEGQuality = value
50270   Case "JPEGRESOLUTION": Options1.JPEGResolution = value
50280   Case "LANGUAGE": Options1.Language = value
50290   Case "LASTSAVEDIRECTORY": Options1.LastSaveDirectory = value
50300   Case "LASTUPDATECHECK": Options1.LastUpdateCheck = value
50310   Case "LOGGING": Options1.Logging = value
50320   Case "LOGLINES": Options1.LogLines = value
50330   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": Options1.NoConfirmMessageSwitchingDefaultprinter = value
50340   Case "NOPROCESSINGATSTARTUP": Options1.NoProcessingAtStartup = value
50350   Case "NOPSCHECK": Options1.NoPSCheck = value
50360   Case "ONEPAGEPERFILE": Options1.OnePagePerFile = value
50370   Case "OPTIONSDESIGN": Options1.OptionsDesign = value
50380   Case "OPTIONSENABLED": Options1.OptionsEnabled = value
50390   Case "OPTIONSVISIBLE": Options1.OptionsVisible = value
50400   Case "PAPERSIZE": Options1.Papersize = value
50410   Case "PCLCOLORSCOUNT": Options1.PCLColorsCount = value
50420   Case "PCLRESOLUTION": Options1.PCLResolution = value
50430   Case "PCXCOLORSCOUNT": Options1.PCXColorscount = value
50440   Case "PCXRESOLUTION": Options1.PCXResolution = value
50450   Case "PDFALLOWASSEMBLY": Options1.PDFAllowAssembly = value
50460   Case "PDFALLOWDEGRADEDPRINTING": Options1.PDFAllowDegradedPrinting = value
50470   Case "PDFALLOWFILLIN": Options1.PDFAllowFillIn = value
50480   Case "PDFALLOWSCREENREADERS": Options1.PDFAllowScreenReaders = value
50490   Case "PDFCOLORSCMYKTORGB": Options1.PDFColorsCMYKToRGB = value
50500   Case "PDFCOLORSCOLORMODEL": Options1.PDFColorsColorModel = value
50510   Case "PDFCOLORSPRESERVEHALFTONE": Options1.PDFColorsPreserveHalftone = value
50520   Case "PDFCOLORSPRESERVEOVERPRINT": Options1.PDFColorsPreserveOverprint = value
50530   Case "PDFCOLORSPRESERVETRANSFER": Options1.PDFColorsPreserveTransfer = value
50540   Case "PDFCOMPRESSIONCOLORCOMPRESSION": Options1.PDFCompressionColorCompression = value
50550   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": Options1.PDFCompressionColorCompressionChoice = value
50560   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": Options1.PDFCompressionColorCompressionJPEGHighFactor = value
50570   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": Options1.PDFCompressionColorCompressionJPEGLowFactor = value
50580   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMANUALFACTOR": Options1.PDFCompressionColorCompressionJPEGManualFactor = value
50590   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": Options1.PDFCompressionColorCompressionJPEGMaximumFactor = value
50600   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": Options1.PDFCompressionColorCompressionJPEGMediumFactor = value
50610   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": Options1.PDFCompressionColorCompressionJPEGMinimumFactor = value
50620   Case "PDFCOMPRESSIONCOLORRESAMPLE": Options1.PDFCompressionColorResample = value
50630   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": Options1.PDFCompressionColorResampleChoice = value
50640   Case "PDFCOMPRESSIONCOLORRESOLUTION": Options1.PDFCompressionColorResolution = value
50650   Case "PDFCOMPRESSIONGREYCOMPRESSION": Options1.PDFCompressionGreyCompression = value
50660   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": Options1.PDFCompressionGreyCompressionChoice = value
50670   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": Options1.PDFCompressionGreyCompressionJPEGHighFactor = value
50680   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": Options1.PDFCompressionGreyCompressionJPEGLowFactor = value
50690   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMANUALFACTOR": Options1.PDFCompressionGreyCompressionJPEGManualFactor = value
50700   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": Options1.PDFCompressionGreyCompressionJPEGMaximumFactor = value
50710   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": Options1.PDFCompressionGreyCompressionJPEGMediumFactor = value
50720   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": Options1.PDFCompressionGreyCompressionJPEGMinimumFactor = value
50730   Case "PDFCOMPRESSIONGREYRESAMPLE": Options1.PDFCompressionGreyResample = value
50740   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": Options1.PDFCompressionGreyResampleChoice = value
50750   Case "PDFCOMPRESSIONGREYRESOLUTION": Options1.PDFCompressionGreyResolution = value
50760   Case "PDFCOMPRESSIONMONOCOMPRESSION": Options1.PDFCompressionMonoCompression = value
50770   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": Options1.PDFCompressionMonoCompressionChoice = value
50780   Case "PDFCOMPRESSIONMONORESAMPLE": Options1.PDFCompressionMonoResample = value
50790   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": Options1.PDFCompressionMonoResampleChoice = value
50800   Case "PDFCOMPRESSIONMONORESOLUTION": Options1.PDFCompressionMonoResolution = value
50810   Case "PDFCOMPRESSIONTEXTCOMPRESSION": Options1.PDFCompressionTextCompression = value
50820   Case "PDFDISALLOWCOPY": Options1.PDFDisallowCopy = value
50830   Case "PDFDISALLOWMODIFYANNOTATIONS": Options1.PDFDisallowModifyAnnotations = value
50840   Case "PDFDISALLOWMODIFYCONTENTS": Options1.PDFDisallowModifyContents = value
50850   Case "PDFDISALLOWPRINTING": Options1.PDFDisallowPrinting = value
50860   Case "PDFENCRYPTOR": Options1.PDFEncryptor = value
50870   Case "PDFFONTSEMBEDALL": Options1.PDFFontsEmbedAll = value
50880   Case "PDFFONTSSUBSETFONTS": Options1.PDFFontsSubSetFonts = value
50890   Case "PDFFONTSSUBSETFONTSPERCENT": Options1.PDFFontsSubSetFontsPercent = value
50900   Case "PDFGENERALASCII85": Options1.PDFGeneralASCII85 = value
50910   Case "PDFGENERALAUTOROTATE": Options1.PDFGeneralAutorotate = value
50920   Case "PDFGENERALCOMPATIBILITY": Options1.PDFGeneralCompatibility = value
50930   Case "PDFGENERALDEFAULT": Options1.PDFGeneralDefault = value
50940   Case "PDFGENERALOVERPRINT": Options1.PDFGeneralOverprint = value
50950   Case "PDFGENERALRESOLUTION": Options1.PDFGeneralResolution = value
50960   Case "PDFHIGHENCRYPTION": Options1.PDFHighEncryption = value
50970   Case "PDFLOWENCRYPTION": Options1.PDFLowEncryption = value
50980   Case "PDFOPTIMIZE": Options1.PDFOptimize = value
50990   Case "PDFOWNERPASS": Options1.PDFOwnerPass = value
51000   Case "PDFOWNERPASSWORDSTRING": Options1.PDFOwnerPasswordString = value
51010   Case "PDFSIGNINGMULTISIGNATURE": Options1.PDFSigningMultiSignature = value
51020   Case "PDFSIGNINGPFXFILE": Options1.PDFSigningPFXFile = value
51030   Case "PDFSIGNINGPFXFILEPASSWORD": Options1.PDFSigningPFXFilePassword = value
51040   Case "PDFSIGNINGSIGNATURECONTACT": Options1.PDFSigningSignatureContact = value
51050   Case "PDFSIGNINGSIGNATURELEFTX": Options1.PDFSigningSignatureLeftX = value
51060   Case "PDFSIGNINGSIGNATURELEFTY": Options1.PDFSigningSignatureLeftY = value
51070   Case "PDFSIGNINGSIGNATURELOCATION": Options1.PDFSigningSignatureLocation = value
51080   Case "PDFSIGNINGSIGNATUREONPAGE": Options1.PDFSigningSignatureOnPage = value
51090   Case "PDFSIGNINGSIGNATUREREASON": Options1.PDFSigningSignatureReason = value
51100   Case "PDFSIGNINGSIGNATURERIGHTX": Options1.PDFSigningSignatureRightX = value
51110   Case "PDFSIGNINGSIGNATURERIGHTY": Options1.PDFSigningSignatureRightY = value
51120   Case "PDFSIGNINGSIGNATUREVISIBLE": Options1.PDFSigningSignatureVisible = value
51130   Case "PDFSIGNINGSIGNPDF": Options1.PDFSigningSignPDF = value
51140   Case "PDFUPDATEMETADATA": Options1.PDFUpdateMetadata = value
51150   Case "PDFUSERPASS": Options1.PDFUserPass = value
51160   Case "PDFUSERPASSWORDSTRING": Options1.PDFUserPasswordString = value
51170   Case "PDFUSESECURITY": Options1.PDFUseSecurity = value
51180   Case "PNGCOLORSCOUNT": Options1.PNGColorscount = value
51190   Case "PNGRESOLUTION": Options1.PNGResolution = value
51200   Case "PRINTAFTERSAVING": Options1.PrintAfterSaving = value
51210   Case "PRINTAFTERSAVINGBITSPERPIXEL": Options1.PrintAfterSavingBitsPerPixel = value
51220   Case "PRINTAFTERSAVINGDUPLEX": Options1.PrintAfterSavingDuplex = value
51230   Case "PRINTAFTERSAVINGMAXRESOLUTION": Options1.PrintAfterSavingMaxResolution = value
51240   Case "PRINTAFTERSAVINGMAXRESOLUTIONENABLED": Options1.PrintAfterSavingMaxResolutionEnabled = value
51250   Case "PRINTAFTERSAVINGNOCANCEL": Options1.PrintAfterSavingNoCancel = value
51260   Case "PRINTAFTERSAVINGPRINTER": Options1.PrintAfterSavingPrinter = value
51270   Case "PRINTAFTERSAVINGQUERYUSER": Options1.PrintAfterSavingQueryUser = value
51280   Case "PRINTAFTERSAVINGTUMBLE": Options1.PrintAfterSavingTumble = value
51290   Case "PRINTERSTOP": Options1.PrinterStop = value
51300   Case "PRINTERTEMPPATH": Options1.PrinterTemppath = value
51310   Case "PROCESSPRIORITY": Options1.ProcessPriority = value
51320   Case "PROGRAMFONT": Options1.ProgramFont = value
51330   Case "PROGRAMFONTCHARSET": Options1.ProgramFontCharset = value
51340   Case "PROGRAMFONTSIZE": Options1.ProgramFontSize = value
51350   Case "PSDCOLORSCOUNT": Options1.PSDColorsCount = value
51360   Case "PSDRESOLUTION": Options1.PSDResolution = value
51370   Case "PSLANGUAGELEVEL": Options1.PSLanguageLevel = value
51380   Case "RAWCOLORSCOUNT": Options1.RAWColorsCount = value
51390   Case "RAWRESOLUTION": Options1.RAWResolution = value
51400   Case "REMOVEALLKNOWNFILEEXTENSIONS": Options1.RemoveAllKnownFileExtensions = value
51410   Case "REMOVESPACES": Options1.RemoveSpaces = value
51420   Case "RUNPROGRAMAFTERSAVING": Options1.RunProgramAfterSaving = value
51430   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": Options1.RunProgramAfterSavingProgramname = value
51440   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": Options1.RunProgramAfterSavingProgramParameters = value
51450   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": Options1.RunProgramAfterSavingWaitUntilReady = value
51460   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": Options1.RunProgramAfterSavingWindowstyle = value
51470   Case "RUNPROGRAMBEFORESAVING": Options1.RunProgramBeforeSaving = value
51480   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": Options1.RunProgramBeforeSavingProgramname = value
51490   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": Options1.RunProgramBeforeSavingProgramParameters = value
51500   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": Options1.RunProgramBeforeSavingWindowstyle = value
51510   Case "SAVEFILENAME": Options1.SaveFilename = value
51520   Case "SENDEMAILAFTERAUTOSAVING": Options1.SendEmailAfterAutoSaving = value
51530   Case "SENDMAILMETHOD": Options1.SendMailMethod = value
51540   Case "SHOWANIMATION": Options1.ShowAnimation = value
51550   Case "STAMPFONTCOLOR": Options1.StampFontColor = value
51560   Case "STAMPFONTNAME": Options1.StampFontname = value
51570   Case "STAMPFONTSIZE": Options1.StampFontsize = value
51580   Case "STAMPOUTLINEFONTTHICKNESS": Options1.StampOutlineFontthickness = value
51590   Case "STAMPSTRING": Options1.StampString = value
51600   Case "STAMPUSEOUTLINEFONT": Options1.StampUseOutlineFont = value
51610   Case "STANDARDAUTHOR": Options1.StandardAuthor = value
51620   Case "STANDARDCREATIONDATE": Options1.StandardCreationdate = value
51630   Case "STANDARDDATEFORMAT": Options1.StandardDateformat = value
51640   Case "STANDARDKEYWORDS": Options1.StandardKeywords = value
51650   Case "STANDARDMAILDOMAIN": Options1.StandardMailDomain = value
51660   Case "STANDARDMODIFYDATE": Options1.StandardModifydate = value
51670   Case "STANDARDSAVEFORMAT": Options1.StandardSaveformat = value
51680   Case "STANDARDSUBJECT": Options1.StandardSubject = value
51690   Case "STANDARDTITLE": Options1.StandardTitle = value
51700   Case "STARTSTANDARDPROGRAM": Options1.StartStandardProgram = value
51710   Case "SVGRESOLUTION": Options1.SVGResolution = value
51720   Case "TIFFCOLORSCOUNT": Options1.TIFFColorscount = value
51730   Case "TIFFRESOLUTION": Options1.TIFFResolution = value
51740   Case "TOOLBARS": Options1.Toolbars = value
51750   Case "UPDATEINTERVAL": Options1.UpdateInterval = value
51760   Case "USEAUTOSAVE": Options1.UseAutosave = value
51770   Case "USEAUTOSAVEDIRECTORY": Options1.UseAutosaveDirectory = value
51780   Case "USECREATIONDATENOW": Options1.UseCreationDateNow = value
51790   Case "USECUSTOMPAPERSIZE": Options1.UseCustomPaperSize = value
51800   Case "USEFIXPAPERSIZE": Options1.UseFixPapersize = value
51810   Case "USESTANDARDAUTHOR": Options1.UseStandardAuthor = value
51820   Case Else:
51830    mError.Number = 3
51840    mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in cLetOption")
51850    RaiseEvent eError
51860  End Select
51870  If ProfileName = "" Then
51880    Options = Options1
51890   Else
51900    SaveOptions Options1, ProfileName
51910  End If
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
50580   .PDFCompressionColorCompressionJPEGManualFactor = sOptions.PDFCompressionColorCompressionJPEGManualFactor
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
50690   .PDFCompressionGreyCompressionJPEGManualFactor = sOptions.PDFCompressionGreyCompressionJPEGManualFactor
50700   .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
50710   .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
50720   .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
50730   .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
50740   .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
50750   .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
50760   .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
50770   .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
50780   .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
50790   .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
50800   .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
50810   .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
50820   .PDFDisallowCopy = sOptions.PDFDisallowCopy
50830   .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
50840   .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
50850   .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
50860   .PDFEncryptor = sOptions.PDFEncryptor
50870   .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
50880   .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
50890   .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
50900   .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
50910   .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
50920   .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
50930   .PDFGeneralDefault = sOptions.PDFGeneralDefault
50940   .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
50950   .PDFGeneralResolution = sOptions.PDFGeneralResolution
50960   .PDFHighEncryption = sOptions.PDFHighEncryption
50970   .PDFLowEncryption = sOptions.PDFLowEncryption
50980   .PDFOptimize = sOptions.PDFOptimize
50990   .PDFOwnerPass = sOptions.PDFOwnerPass
51000   .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
51010   .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
51020   .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
51030   .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
51040   .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
51050   .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
51060   .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
51070   .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
51080   .PDFSigningSignatureOnPage = sOptions.PDFSigningSignatureOnPage
51090   .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
51100   .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
51110   .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
51120   .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
51130   .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
51140   .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
51150   .PDFUserPass = sOptions.PDFUserPass
51160   .PDFUserPasswordString = sOptions.PDFUserPasswordString
51170   .PDFUseSecurity = sOptions.PDFUseSecurity
51180   .PNGColorscount = sOptions.PNGColorscount
51190   .PNGResolution = sOptions.PNGResolution
51200   .PrintAfterSaving = sOptions.PrintAfterSaving
51210   .PrintAfterSavingBitsPerPixel = sOptions.PrintAfterSavingBitsPerPixel
51220   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
51230   .PrintAfterSavingMaxResolution = sOptions.PrintAfterSavingMaxResolution
51240   .PrintAfterSavingMaxResolutionEnabled = sOptions.PrintAfterSavingMaxResolutionEnabled
51250   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
51260   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
51270   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51280   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51290   .PrinterStop = sOptions.PrinterStop
51300   .PrinterTemppath = sOptions.PrinterTemppath
51310   .ProcessPriority = sOptions.ProcessPriority
51320   .ProgramFont = sOptions.ProgramFont
51330   .ProgramFontCharset = sOptions.ProgramFontCharset
51340   .ProgramFontSize = sOptions.ProgramFontSize
51350   .PSDColorsCount = sOptions.PSDColorsCount
51360   .PSDResolution = sOptions.PSDResolution
51370   .PSLanguageLevel = sOptions.PSLanguageLevel
51380   .RAWColorsCount = sOptions.RAWColorsCount
51390   .RAWResolution = sOptions.RAWResolution
51400   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51410   .RemoveSpaces = sOptions.RemoveSpaces
51420   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51430   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51440   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51450   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51460   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51470   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51480   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51490   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51500   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51510   .SaveFilename = sOptions.SaveFilename
51520   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51530   .SendMailMethod = sOptions.SendMailMethod
51540   .ShowAnimation = sOptions.ShowAnimation
51550   .StampFontColor = sOptions.StampFontColor
51560   .StampFontname = sOptions.StampFontname
51570   .StampFontsize = sOptions.StampFontsize
51580   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51590   .StampString = sOptions.StampString
51600   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51610   .StandardAuthor = sOptions.StandardAuthor
51620   .StandardCreationdate = sOptions.StandardCreationdate
51630   .StandardDateformat = sOptions.StandardDateformat
51640   .StandardKeywords = sOptions.StandardKeywords
51650   .StandardMailDomain = sOptions.StandardMailDomain
51660   .StandardModifydate = sOptions.StandardModifydate
51670   .StandardSaveformat = sOptions.StandardSaveformat
51680   .StandardSubject = sOptions.StandardSubject
51690   .StandardTitle = sOptions.StandardTitle
51700   .StartStandardProgram = sOptions.StartStandardProgram
51710   .SVGResolution = sOptions.SVGResolution
51720   .TIFFColorscount = sOptions.TIFFColorscount
51730   .TIFFResolution = sOptions.TIFFResolution
51740   .Toolbars = sOptions.Toolbars
51750   .UpdateInterval = sOptions.UpdateInterval
51760   .UseAutosave = sOptions.UseAutosave
51770   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51780   .UseCreationDateNow = sOptions.UseCreationDateNow
51790   .UseCustomPaperSize = sOptions.UseCustomPaperSize
51800   .UseFixPapersize = sOptions.UseFixPapersize
51810   .UseStandardAuthor = sOptions.UseStandardAuthor
51820  End With
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

