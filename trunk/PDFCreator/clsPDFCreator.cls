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
 Set cOptionsNames = New Collection
 With cOptionsNames
  .Add "AdditionalGhostscriptParameters"
  .Add "AdditionalGhostscriptSearchpath"
  .Add "AddWindowsFontpath"
  .Add "AllowSpecialGSCharsInFilenames"
  .Add "AutosaveDirectory"
  .Add "AutosaveFilename"
  .Add "AutosaveFormat"
  .Add "AutosaveStartStandardProgram"
  .Add "BMPColorscount"
  .Add "BMPResolution"
  .Add "ClientComputerResolveIPAddress"
  .Add "Counter"
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
  .Add "JPEGResolution"
  .Add "Language"
  .Add "LastSaveDirectory"
  .Add "LastUpdateCheck"
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
  .Add "PCLColorsCount"
  .Add "PCLResolution"
  .Add "PCXColorscount"
  .Add "PCXResolution"
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
  .Add "PDFGeneralDefault"
  .Add "PDFGeneralOverprint"
  .Add "PDFGeneralResolution"
  .Add "PDFHighEncryption"
  .Add "PDFLowEncryption"
  .Add "PDFOptimize"
  .Add "PDFOwnerPass"
  .Add "PDFOwnerPasswordString"
  .Add "PDFSigningMultiSignature"
  .Add "PDFSigningPFXFile"
  .Add "PDFSigningPFXFilePassword"
  .Add "PDFSigningSignatureContact"
  .Add "PDFSigningSignatureLeftX"
  .Add "PDFSigningSignatureLeftY"
  .Add "PDFSigningSignatureLocation"
  .Add "PDFSigningSignatureOnPage"
  .Add "PDFSigningSignatureReason"
  .Add "PDFSigningSignatureRightX"
  .Add "PDFSigningSignatureRightY"
  .Add "PDFSigningSignatureVisible"
  .Add "PDFSigningSignPDF"
  .Add "PDFUpdateMetadata"
  .Add "PDFUserPass"
  .Add "PDFUserPasswordString"
  .Add "PDFUseSecurity"
  .Add "PNGColorscount"
  .Add "PNGResolution"
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
  .Add "PSDColorsCount"
  .Add "PSDResolution"
  .Add "PSLanguageLevel"
  .Add "RAWColorsCount"
  .Add "RAWResolution"
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
  .Add "SVGResolution"
  .Add "TIFFColorscount"
  .Add "TIFFResolution"
  .Add "Toolbars"
  .Add "UpdateInterval"
  .Add "UseAutosave"
  .Add "UseAutosaveDirectory"
  .Add "UseCreationDateNow"
  .Add "UseCustomPaperSize"
  .Add "UseFixPapersize"
  .Add "UseStandardAuthor"
 End With
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
 Select Case UCase$(PropertyName)
  Case "ADDITIONALGHOSTSCRIPTPARAMETERS": cGetOptionFromOptions = Options.AdditionalGhostscriptParameters
  Case "ADDITIONALGHOSTSCRIPTSEARCHPATH": cGetOptionFromOptions = Options.AdditionalGhostscriptSearchpath
  Case "ADDWINDOWSFONTPATH": cGetOptionFromOptions = Options.AddWindowsFontpath
  Case "ALLOWSPECIALGSCHARSINFILENAMES": cGetOptionFromOptions = Options.AllowSpecialGSCharsInFilenames
  Case "AUTOSAVEDIRECTORY": cGetOptionFromOptions = Options.AutosaveDirectory
  Case "AUTOSAVEFILENAME": cGetOptionFromOptions = Options.AutosaveFilename
  Case "AUTOSAVEFORMAT": cGetOptionFromOptions = Options.AutosaveFormat
  Case "AUTOSAVESTARTSTANDARDPROGRAM": cGetOptionFromOptions = Options.AutosaveStartStandardProgram
  Case "BMPCOLORSCOUNT": cGetOptionFromOptions = Options.BMPColorscount
  Case "BMPRESOLUTION": cGetOptionFromOptions = Options.BMPResolution
  Case "CLIENTCOMPUTERRESOLVEIPADDRESS": cGetOptionFromOptions = Options.ClientComputerResolveIPAddress
  Case "COUNTER": cGetOptionFromOptions = Options.Counter
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
  Case "JPEGRESOLUTION": cGetOptionFromOptions = Options.JPEGResolution
  Case "LANGUAGE": cGetOptionFromOptions = Options.Language
  Case "LASTSAVEDIRECTORY": cGetOptionFromOptions = Options.LastSaveDirectory
  Case "LASTUPDATECHECK": cGetOptionFromOptions = Options.LastUpdateCheck
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
  Case "PCLCOLORSCOUNT": cGetOptionFromOptions = Options.PCLColorsCount
  Case "PCLRESOLUTION": cGetOptionFromOptions = Options.PCLResolution
  Case "PCXCOLORSCOUNT": cGetOptionFromOptions = Options.PCXColorscount
  Case "PCXRESOLUTION": cGetOptionFromOptions = Options.PCXResolution
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
  Case "PDFGENERALDEFAULT": cGetOptionFromOptions = Options.PDFGeneralDefault
  Case "PDFGENERALOVERPRINT": cGetOptionFromOptions = Options.PDFGeneralOverprint
  Case "PDFGENERALRESOLUTION": cGetOptionFromOptions = Options.PDFGeneralResolution
  Case "PDFHIGHENCRYPTION": cGetOptionFromOptions = Options.PDFHighEncryption
  Case "PDFLOWENCRYPTION": cGetOptionFromOptions = Options.PDFLowEncryption
  Case "PDFOPTIMIZE": cGetOptionFromOptions = Options.PDFOptimize
  Case "PDFOWNERPASS": cGetOptionFromOptions = Options.PDFOwnerPass
  Case "PDFOWNERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFOwnerPasswordString
  Case "PDFSIGNINGMULTISIGNATURE": cGetOptionFromOptions = Options.PDFSigningMultiSignature
  Case "PDFSIGNINGPFXFILE": cGetOptionFromOptions = Options.PDFSigningPFXFile
  Case "PDFSIGNINGPFXFILEPASSWORD": cGetOptionFromOptions = Options.PDFSigningPFXFilePassword
  Case "PDFSIGNINGSIGNATURECONTACT": cGetOptionFromOptions = Options.PDFSigningSignatureContact
  Case "PDFSIGNINGSIGNATURELEFTX": cGetOptionFromOptions = Options.PDFSigningSignatureLeftX
  Case "PDFSIGNINGSIGNATURELEFTY": cGetOptionFromOptions = Options.PDFSigningSignatureLeftY
  Case "PDFSIGNINGSIGNATURELOCATION": cGetOptionFromOptions = Options.PDFSigningSignatureLocation
  Case "PDFSIGNINGSIGNATUREONPAGE": cGetOptionFromOptions = Options.PDFSigningSignatureOnPage
  Case "PDFSIGNINGSIGNATUREREASON": cGetOptionFromOptions = Options.PDFSigningSignatureReason
  Case "PDFSIGNINGSIGNATURERIGHTX": cGetOptionFromOptions = Options.PDFSigningSignatureRightX
  Case "PDFSIGNINGSIGNATURERIGHTY": cGetOptionFromOptions = Options.PDFSigningSignatureRightY
  Case "PDFSIGNINGSIGNATUREVISIBLE": cGetOptionFromOptions = Options.PDFSigningSignatureVisible
  Case "PDFSIGNINGSIGNPDF": cGetOptionFromOptions = Options.PDFSigningSignPDF
  Case "PDFUPDATEMETADATA": cGetOptionFromOptions = Options.PDFUpdateMetadata
  Case "PDFUSERPASS": cGetOptionFromOptions = Options.PDFUserPass
  Case "PDFUSERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFUserPasswordString
  Case "PDFUSESECURITY": cGetOptionFromOptions = Options.PDFUseSecurity
  Case "PNGCOLORSCOUNT": cGetOptionFromOptions = Options.PNGColorscount
  Case "PNGRESOLUTION": cGetOptionFromOptions = Options.PNGResolution
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
  Case "PSDCOLORSCOUNT": cGetOptionFromOptions = Options.PSDColorsCount
  Case "PSDRESOLUTION": cGetOptionFromOptions = Options.PSDResolution
  Case "PSLANGUAGELEVEL": cGetOptionFromOptions = Options.PSLanguageLevel
  Case "RAWCOLORSCOUNT": cGetOptionFromOptions = Options.RAWColorsCount
  Case "RAWRESOLUTION": cGetOptionFromOptions = Options.RAWResolution
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
  Case "SVGRESOLUTION": cGetOptionFromOptions = Options.SVGResolution
  Case "TIFFCOLORSCOUNT": cGetOptionFromOptions = Options.TIFFColorscount
  Case "TIFFRESOLUTION": cGetOptionFromOptions = Options.TIFFResolution
  Case "TOOLBARS": cGetOptionFromOptions = Options.Toolbars
  Case "UPDATEINTERVAL": cGetOptionFromOptions = Options.UpdateInterval
  Case "USEAUTOSAVE": cGetOptionFromOptions = Options.UseAutosave
  Case "USEAUTOSAVEDIRECTORY": cGetOptionFromOptions = Options.UseAutosaveDirectory
  Case "USECREATIONDATENOW": cGetOptionFromOptions = Options.UseCreationDateNow
  Case "USECUSTOMPAPERSIZE": cGetOptionFromOptions = Options.UseCustomPaperSize
  Case "USEFIXPAPERSIZE": cGetOptionFromOptions = Options.UseFixPapersize
  Case "USESTANDARDAUTHOR": cGetOptionFromOptions = Options.UseStandardAuthor
 End Select
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
 Set cGetOptions = New clsPDFCreatorOptions
 With cGetOptions
  .AdditionalGhostscriptParameters = sOptions.AdditionalGhostscriptParameters
  .AdditionalGhostscriptSearchpath = sOptions.AdditionalGhostscriptSearchpath
  .AddWindowsFontpath = sOptions.AddWindowsFontpath
  .AllowSpecialGSCharsInFilenames = sOptions.AllowSpecialGSCharsInFilenames
  .AutosaveDirectory = sOptions.AutosaveDirectory
  .AutosaveFilename = sOptions.AutosaveFilename
  .AutosaveFormat = sOptions.AutosaveFormat
  .AutosaveStartStandardProgram = sOptions.AutosaveStartStandardProgram
  .BMPColorscount = sOptions.BMPColorscount
  .BMPResolution = sOptions.BMPResolution
  .ClientComputerResolveIPAddress = sOptions.ClientComputerResolveIPAddress
  .Counter = sOptions.Counter
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
  .JPEGResolution = sOptions.JPEGResolution
  .Language = sOptions.Language
  .LastSaveDirectory = sOptions.LastSaveDirectory
  .LastUpdateCheck = sOptions.LastUpdateCheck
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
  .PCLColorsCount = sOptions.PCLColorsCount
  .PCLResolution = sOptions.PCLResolution
  .PCXColorscount = sOptions.PCXColorscount
  .PCXResolution = sOptions.PCXResolution
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
  .PDFGeneralDefault = sOptions.PDFGeneralDefault
  .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
  .PDFGeneralResolution = sOptions.PDFGeneralResolution
  .PDFHighEncryption = sOptions.PDFHighEncryption
  .PDFLowEncryption = sOptions.PDFLowEncryption
  .PDFOptimize = sOptions.PDFOptimize
  .PDFOwnerPass = sOptions.PDFOwnerPass
  .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
  .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
  .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
  .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
  .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
  .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
  .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
  .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
  .PDFSigningSignatureOnPage = sOptions.PDFSigningSignatureOnPage
  .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
  .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
  .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
  .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
  .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
  .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
  .PDFUserPass = sOptions.PDFUserPass
  .PDFUserPasswordString = sOptions.PDFUserPasswordString
  .PDFUseSecurity = sOptions.PDFUseSecurity
  .PNGColorscount = sOptions.PNGColorscount
  .PNGResolution = sOptions.PNGResolution
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
  .PSDColorsCount = sOptions.PSDColorsCount
  .PSDResolution = sOptions.PSDResolution
  .PSLanguageLevel = sOptions.PSLanguageLevel
  .RAWColorsCount = sOptions.RAWColorsCount
  .RAWResolution = sOptions.RAWResolution
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
  .SVGResolution = sOptions.SVGResolution
  .TIFFColorscount = sOptions.TIFFColorscount
  .TIFFResolution = sOptions.TIFFResolution
  .Toolbars = sOptions.Toolbars
  .UpdateInterval = sOptions.UpdateInterval
  .UseAutosave = sOptions.UseAutosave
  .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
  .UseCreationDateNow = sOptions.UseCreationDateNow
  .UseCustomPaperSize = sOptions.UseCustomPaperSize
  .UseFixPapersize = sOptions.UseFixPapersize
  .UseStandardAuthor = sOptions.UseStandardAuthor
 End With
End Function

Private Sub cLetOption(PropertyName As String, value As Variant, Options1 As tOptions, Profilename As String)
 Select Case UCase$(PropertyName)
  Case "ADDITIONALGHOSTSCRIPTPARAMETERS": Options1.AdditionalGhostscriptParameters = Value
  Case "ADDITIONALGHOSTSCRIPTSEARCHPATH": Options1.AdditionalGhostscriptSearchpath = Value
  Case "ADDWINDOWSFONTPATH": Options1.AddWindowsFontpath = Value
  Case "ALLOWSPECIALGSCHARSINFILENAMES": Options1.AllowSpecialGSCharsInFilenames = Value
  Case "AUTOSAVEDIRECTORY": Options1.AutosaveDirectory = Value
  Case "AUTOSAVEFILENAME": Options1.AutosaveFilename = Value
  Case "AUTOSAVEFORMAT": Options1.AutosaveFormat = Value
  Case "AUTOSAVESTARTSTANDARDPROGRAM": Options1.AutosaveStartStandardProgram = Value
  Case "BMPCOLORSCOUNT": Options1.BMPColorscount = Value
  Case "BMPRESOLUTION": Options1.BMPResolution = Value
  Case "CLIENTCOMPUTERRESOLVEIPADDRESS": Options1.ClientComputerResolveIPAddress = Value
  Case "COUNTER": Options1.Counter = Value
  Case "DEVICEHEIGHTPOINTS": Options1.DeviceHeightPoints = Value
  Case "DEVICEWIDTHPOINTS": Options1.DeviceWidthPoints = Value
  Case "DIRECTORYGHOSTSCRIPTBINARIES": Options1.DirectoryGhostscriptBinaries = Value
  Case "DIRECTORYGHOSTSCRIPTFONTS": Options1.DirectoryGhostscriptFonts = Value
  Case "DIRECTORYGHOSTSCRIPTLIBRARIES": Options1.DirectoryGhostscriptLibraries = Value
  Case "DIRECTORYGHOSTSCRIPTRESOURCE": Options1.DirectoryGhostscriptResource = Value
  Case "DISABLEEMAIL": Options1.DisableEmail = Value
  Case "DONTUSEDOCUMENTSETTINGS": Options1.DontUseDocumentSettings = Value
  Case "EPSLANGUAGELEVEL": Options1.EPSLanguageLevel = Value
  Case "FILENAMESUBSTITUTIONS": Options1.FilenameSubstitutions = Value
  Case "FILENAMESUBSTITUTIONSONLYINTITLE": Options1.FilenameSubstitutionsOnlyInTitle = Value
  Case "JPEGCOLORSCOUNT": Options1.JPEGColorscount = Value
  Case "JPEGQUALITY": Options1.JPEGQuality = Value
  Case "JPEGRESOLUTION": Options1.JPEGResolution = Value
  Case "LANGUAGE": Options1.Language = Value
  Case "LASTSAVEDIRECTORY": Options1.LastSaveDirectory = Value
  Case "LASTUPDATECHECK": Options1.LastUpdateCheck = Value
  Case "LOGGING": Options1.Logging = Value
  Case "LOGLINES": Options1.LogLines = Value
  Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": Options1.NoConfirmMessageSwitchingDefaultprinter = Value
  Case "NOPROCESSINGATSTARTUP": Options1.NoProcessingAtStartup = Value
  Case "NOPSCHECK": Options1.NoPSCheck = Value
  Case "ONEPAGEPERFILE": Options1.OnePagePerFile = Value
  Case "OPTIONSDESIGN": Options1.OptionsDesign = Value
  Case "OPTIONSENABLED": Options1.OptionsEnabled = Value
  Case "OPTIONSVISIBLE": Options1.OptionsVisible = Value
  Case "PAPERSIZE": Options1.Papersize = Value
  Case "PCLCOLORSCOUNT": Options1.PCLColorsCount = Value
  Case "PCLRESOLUTION": Options1.PCLResolution = Value
  Case "PCXCOLORSCOUNT": Options1.PCXColorscount = Value
  Case "PCXRESOLUTION": Options1.PCXResolution = Value
  Case "PDFALLOWASSEMBLY": Options1.PDFAllowAssembly = Value
  Case "PDFALLOWDEGRADEDPRINTING": Options1.PDFAllowDegradedPrinting = Value
  Case "PDFALLOWFILLIN": Options1.PDFAllowFillIn = Value
  Case "PDFALLOWSCREENREADERS": Options1.PDFAllowScreenReaders = Value
  Case "PDFCOLORSCMYKTORGB": Options1.PDFColorsCMYKToRGB = Value
  Case "PDFCOLORSCOLORMODEL": Options1.PDFColorsColorModel = Value
  Case "PDFCOLORSPRESERVEHALFTONE": Options1.PDFColorsPreserveHalftone = Value
  Case "PDFCOLORSPRESERVEOVERPRINT": Options1.PDFColorsPreserveOverprint = Value
  Case "PDFCOLORSPRESERVETRANSFER": Options1.PDFColorsPreserveTransfer = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSION": Options1.PDFCompressionColorCompression = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": Options1.PDFCompressionColorCompressionChoice = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": Options1.PDFCompressionColorCompressionJPEGHighFactor = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": Options1.PDFCompressionColorCompressionJPEGLowFactor = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": Options1.PDFCompressionColorCompressionJPEGMaximumFactor = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": Options1.PDFCompressionColorCompressionJPEGMediumFactor = Value
  Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": Options1.PDFCompressionColorCompressionJPEGMinimumFactor = Value
  Case "PDFCOMPRESSIONCOLORRESAMPLE": Options1.PDFCompressionColorResample = Value
  Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": Options1.PDFCompressionColorResampleChoice = Value
  Case "PDFCOMPRESSIONCOLORRESOLUTION": Options1.PDFCompressionColorResolution = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSION": Options1.PDFCompressionGreyCompression = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": Options1.PDFCompressionGreyCompressionChoice = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": Options1.PDFCompressionGreyCompressionJPEGHighFactor = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": Options1.PDFCompressionGreyCompressionJPEGLowFactor = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": Options1.PDFCompressionGreyCompressionJPEGMaximumFactor = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": Options1.PDFCompressionGreyCompressionJPEGMediumFactor = Value
  Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": Options1.PDFCompressionGreyCompressionJPEGMinimumFactor = Value
  Case "PDFCOMPRESSIONGREYRESAMPLE": Options1.PDFCompressionGreyResample = Value
  Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": Options1.PDFCompressionGreyResampleChoice = Value
  Case "PDFCOMPRESSIONGREYRESOLUTION": Options1.PDFCompressionGreyResolution = Value
  Case "PDFCOMPRESSIONMONOCOMPRESSION": Options1.PDFCompressionMonoCompression = Value
  Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": Options1.PDFCompressionMonoCompressionChoice = Value
  Case "PDFCOMPRESSIONMONORESAMPLE": Options1.PDFCompressionMonoResample = Value
  Case "PDFCOMPRESSIONMONORESAMPLECHOICE": Options1.PDFCompressionMonoResampleChoice = Value
  Case "PDFCOMPRESSIONMONORESOLUTION": Options1.PDFCompressionMonoResolution = Value
  Case "PDFCOMPRESSIONTEXTCOMPRESSION": Options1.PDFCompressionTextCompression = Value
  Case "PDFDISALLOWCOPY": Options1.PDFDisallowCopy = Value
  Case "PDFDISALLOWMODIFYANNOTATIONS": Options1.PDFDisallowModifyAnnotations = Value
  Case "PDFDISALLOWMODIFYCONTENTS": Options1.PDFDisallowModifyContents = Value
  Case "PDFDISALLOWPRINTING": Options1.PDFDisallowPrinting = Value
  Case "PDFENCRYPTOR": Options1.PDFEncryptor = Value
  Case "PDFFONTSEMBEDALL": Options1.PDFFontsEmbedAll = Value
  Case "PDFFONTSSUBSETFONTS": Options1.PDFFontsSubSetFonts = Value
  Case "PDFFONTSSUBSETFONTSPERCENT": Options1.PDFFontsSubSetFontsPercent = Value
  Case "PDFGENERALASCII85": Options1.PDFGeneralASCII85 = Value
  Case "PDFGENERALAUTOROTATE": Options1.PDFGeneralAutorotate = Value
  Case "PDFGENERALCOMPATIBILITY": Options1.PDFGeneralCompatibility = Value
  Case "PDFGENERALDEFAULT": Options1.PDFGeneralDefault = Value
  Case "PDFGENERALOVERPRINT": Options1.PDFGeneralOverprint = Value
  Case "PDFGENERALRESOLUTION": Options1.PDFGeneralResolution = Value
  Case "PDFHIGHENCRYPTION": Options1.PDFHighEncryption = Value
  Case "PDFLOWENCRYPTION": Options1.PDFLowEncryption = Value
  Case "PDFOPTIMIZE": Options1.PDFOptimize = Value
  Case "PDFOWNERPASS": Options1.PDFOwnerPass = Value
  Case "PDFOWNERPASSWORDSTRING": Options1.PDFOwnerPasswordString = Value
  Case "PDFSIGNINGMULTISIGNATURE": Options1.PDFSigningMultiSignature = Value
  Case "PDFSIGNINGPFXFILE": Options1.PDFSigningPFXFile = Value
  Case "PDFSIGNINGPFXFILEPASSWORD": Options1.PDFSigningPFXFilePassword = Value
  Case "PDFSIGNINGSIGNATURECONTACT": Options1.PDFSigningSignatureContact = Value
  Case "PDFSIGNINGSIGNATURELEFTX": Options1.PDFSigningSignatureLeftX = Value
  Case "PDFSIGNINGSIGNATURELEFTY": Options1.PDFSigningSignatureLeftY = Value
  Case "PDFSIGNINGSIGNATURELOCATION": Options1.PDFSigningSignatureLocation = Value
  Case "PDFSIGNINGSIGNATUREONPAGE": Options1.PDFSigningSignatureOnPage = Value
  Case "PDFSIGNINGSIGNATUREREASON": Options1.PDFSigningSignatureReason = Value
  Case "PDFSIGNINGSIGNATURERIGHTX": Options1.PDFSigningSignatureRightX = Value
  Case "PDFSIGNINGSIGNATURERIGHTY": Options1.PDFSigningSignatureRightY = Value
  Case "PDFSIGNINGSIGNATUREVISIBLE": Options1.PDFSigningSignatureVisible = Value
  Case "PDFSIGNINGSIGNPDF": Options1.PDFSigningSignPDF = Value
  Case "PDFUPDATEMETADATA": Options1.PDFUpdateMetadata = Value
  Case "PDFUSERPASS": Options1.PDFUserPass = Value
  Case "PDFUSERPASSWORDSTRING": Options1.PDFUserPasswordString = Value
  Case "PDFUSESECURITY": Options1.PDFUseSecurity = Value
  Case "PNGCOLORSCOUNT": Options1.PNGColorscount = Value
  Case "PNGRESOLUTION": Options1.PNGResolution = Value
  Case "PRINTAFTERSAVING": Options1.PrintAfterSaving = Value
  Case "PRINTAFTERSAVINGDUPLEX": Options1.PrintAfterSavingDuplex = Value
  Case "PRINTAFTERSAVINGNOCANCEL": Options1.PrintAfterSavingNoCancel = Value
  Case "PRINTAFTERSAVINGPRINTER": Options1.PrintAfterSavingPrinter = Value
  Case "PRINTAFTERSAVINGQUERYUSER": Options1.PrintAfterSavingQueryUser = Value
  Case "PRINTAFTERSAVINGTUMBLE": Options1.PrintAfterSavingTumble = Value
  Case "PRINTERSTOP": Options1.PrinterStop = Value
  Case "PRINTERTEMPPATH": Options1.PrinterTemppath = Value
  Case "PROCESSPRIORITY": Options1.ProcessPriority = Value
  Case "PROGRAMFONT": Options1.ProgramFont = Value
  Case "PROGRAMFONTCHARSET": Options1.ProgramFontCharset = Value
  Case "PROGRAMFONTSIZE": Options1.ProgramFontSize = Value
  Case "PSDCOLORSCOUNT": Options1.PSDColorsCount = Value
  Case "PSDRESOLUTION": Options1.PSDResolution = Value
  Case "PSLANGUAGELEVEL": Options1.PSLanguageLevel = Value
  Case "RAWCOLORSCOUNT": Options1.RAWColorsCount = Value
  Case "RAWRESOLUTION": Options1.RAWResolution = Value
  Case "REMOVEALLKNOWNFILEEXTENSIONS": Options1.RemoveAllKnownFileExtensions = Value
  Case "REMOVESPACES": Options1.RemoveSpaces = Value
  Case "RUNPROGRAMAFTERSAVING": Options1.RunProgramAfterSaving = Value
  Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": Options1.RunProgramAfterSavingProgramname = Value
  Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": Options1.RunProgramAfterSavingProgramParameters = Value
  Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": Options1.RunProgramAfterSavingWaitUntilReady = Value
  Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": Options1.RunProgramAfterSavingWindowstyle = Value
  Case "RUNPROGRAMBEFORESAVING": Options1.RunProgramBeforeSaving = Value
  Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": Options1.RunProgramBeforeSavingProgramname = Value
  Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": Options1.RunProgramBeforeSavingProgramParameters = Value
  Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": Options1.RunProgramBeforeSavingWindowstyle = Value
  Case "SAVEFILENAME": Options1.SaveFilename = Value
  Case "SENDEMAILAFTERAUTOSAVING": Options1.SendEmailAfterAutoSaving = Value
  Case "SENDMAILMETHOD": Options1.SendMailMethod = Value
  Case "SHOWANIMATION": Options1.ShowAnimation = Value
  Case "STAMPFONTCOLOR": Options1.StampFontColor = Value
  Case "STAMPFONTNAME": Options1.StampFontname = Value
  Case "STAMPFONTSIZE": Options1.StampFontsize = Value
  Case "STAMPOUTLINEFONTTHICKNESS": Options1.StampOutlineFontthickness = Value
  Case "STAMPSTRING": Options1.StampString = Value
  Case "STAMPUSEOUTLINEFONT": Options1.StampUseOutlineFont = Value
  Case "STANDARDAUTHOR": Options1.StandardAuthor = Value
  Case "STANDARDCREATIONDATE": Options1.StandardCreationdate = Value
  Case "STANDARDDATEFORMAT": Options1.StandardDateformat = Value
  Case "STANDARDKEYWORDS": Options1.StandardKeywords = Value
  Case "STANDARDMAILDOMAIN": Options1.StandardMailDomain = Value
  Case "STANDARDMODIFYDATE": Options1.StandardModifydate = Value
  Case "STANDARDSAVEFORMAT": Options1.StandardSaveformat = Value
  Case "STANDARDSUBJECT": Options1.StandardSubject = Value
  Case "STANDARDTITLE": Options1.StandardTitle = Value
  Case "STARTSTANDARDPROGRAM": Options1.StartStandardProgram = Value
  Case "SVGRESOLUTION": Options1.SVGResolution = Value
  Case "TIFFCOLORSCOUNT": Options1.TIFFColorscount = Value
  Case "TIFFRESOLUTION": Options1.TIFFResolution = Value
  Case "TOOLBARS": Options1.Toolbars = Value
  Case "UPDATEINTERVAL": Options1.UpdateInterval = Value
  Case "USEAUTOSAVE": Options1.UseAutosave = Value
  Case "USEAUTOSAVEDIRECTORY": Options1.UseAutosaveDirectory = Value
  Case "USECREATIONDATENOW": Options1.UseCreationDateNow = Value
  Case "USECUSTOMPAPERSIZE": Options1.UseCustomPaperSize = Value
  Case "USEFIXPAPERSIZE": Options1.UseFixPapersize = Value
  Case "USESTANDARDAUTHOR": Options1.UseStandardAuthor = Value
  Case Else:
   mError.Number = 3
   mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in cLetOption")
   RaiseEvent eError
 End Select
 If Profilename = "" Then
   Options = Options1
  Else
   SaveOptions Options1, Profilename
 End If
End Sub

Private Function cLetOptions(sOptions As Variant) As tOptions
 With cLetOptions
  .AdditionalGhostscriptParameters = sOptions.AdditionalGhostscriptParameters
  .AdditionalGhostscriptSearchpath = sOptions.AdditionalGhostscriptSearchpath
  .AddWindowsFontpath = sOptions.AddWindowsFontpath
  .AllowSpecialGSCharsInFilenames = sOptions.AllowSpecialGSCharsInFilenames
  .AutosaveDirectory = sOptions.AutosaveDirectory
  .AutosaveFilename = sOptions.AutosaveFilename
  .AutosaveFormat = sOptions.AutosaveFormat
  .AutosaveStartStandardProgram = sOptions.AutosaveStartStandardProgram
  .BMPColorscount = sOptions.BMPColorscount
  .BMPResolution = sOptions.BMPResolution
  .ClientComputerResolveIPAddress = sOptions.ClientComputerResolveIPAddress
  .Counter = sOptions.Counter
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
  .JPEGResolution = sOptions.JPEGResolution
  .Language = sOptions.Language
  .LastSaveDirectory = sOptions.LastSaveDirectory
  .LastUpdateCheck = sOptions.LastUpdateCheck
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
  .PCLColorsCount = sOptions.PCLColorsCount
  .PCLResolution = sOptions.PCLResolution
  .PCXColorscount = sOptions.PCXColorscount
  .PCXResolution = sOptions.PCXResolution
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
  .PDFGeneralDefault = sOptions.PDFGeneralDefault
  .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
  .PDFGeneralResolution = sOptions.PDFGeneralResolution
  .PDFHighEncryption = sOptions.PDFHighEncryption
  .PDFLowEncryption = sOptions.PDFLowEncryption
  .PDFOptimize = sOptions.PDFOptimize
  .PDFOwnerPass = sOptions.PDFOwnerPass
  .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
  .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
  .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
  .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
  .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
  .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
  .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
  .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
  .PDFSigningSignatureOnPage = sOptions.PDFSigningSignatureOnPage
  .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
  .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
  .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
  .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
  .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
  .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
  .PDFUserPass = sOptions.PDFUserPass
  .PDFUserPasswordString = sOptions.PDFUserPasswordString
  .PDFUseSecurity = sOptions.PDFUseSecurity
  .PNGColorscount = sOptions.PNGColorscount
  .PNGResolution = sOptions.PNGResolution
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
  .PSDColorsCount = sOptions.PSDColorsCount
  .PSDResolution = sOptions.PSDResolution
  .PSLanguageLevel = sOptions.PSLanguageLevel
  .RAWColorsCount = sOptions.RAWColorsCount
  .RAWResolution = sOptions.RAWResolution
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
  .SVGResolution = sOptions.SVGResolution
  .TIFFColorscount = sOptions.TIFFColorscount
  .TIFFResolution = sOptions.TIFFResolution
  .Toolbars = sOptions.Toolbars
  .UpdateInterval = sOptions.UpdateInterval
  .UseAutosave = sOptions.UseAutosave
  .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
  .UseCreationDateNow = sOptions.UseCreationDateNow
  .UseCustomPaperSize = sOptions.UseCustomPaperSize
  .UseFixPapersize = sOptions.UseFixPapersize
  .UseStandardAuthor = sOptions.UseStandardAuthor
 End With
End Function
' --- Private functions/Subroutines - end ---

