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
50040  If cProgramIsRunning Then
50050   With mError
50060    .Number = 2
50070    .Description = ErrDescr2
50080   End With
50090   mProgramIsRunning = True
50100   RaiseEvent eError
50110   Exit Sub
50120  End If
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
  .Add "UseStandardAuthor"
 End With
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
  Case "USESTANDARDAUTHOR": cGetOptionFromOptions = Options.UseStandardAuthor
 End Select
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
  .UseStandardAuthor = sOptions.UseStandardAuthor
 End With
End Function
' --- Private functions/Subroutines - end ---














