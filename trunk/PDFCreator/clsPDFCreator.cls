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
Private Const ErrDescr12 = "The profile '%1' already exists!"
Private Const ErrDescr13 = "The name for the profile could not be empty!"
Private Const ErrDescr14 = "You can't delete the profile '%1' because it is associate with the printer '%2'."
Private Const ErrDescr15 = "The old profile '%1' has the same name like the new profile '%2'."
Private Const ErrDescr16 = "The printer '%1' is not installed!"
Private Const ErrDescr17 = "The printer '%1' is already installed!"
Private Const ErrDescr18 = "You must be an administrator to install or delete a printer!"
Private Const ErrDescr19 = "The printer is not a PDFCreator printer!"

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
Public Property Get cPrinterProfile(ByVal PrinterName As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PrinterProfiles As Collection, i As Long
50020  If Not PrinterIsInstalled(PrinterName) Then
50030   mError.Number = 15
50040   mError.Description = Replace$(ErrDescr16, "%1", PrinterName)
50050   RaiseEvent eError
50060   Exit Property
50070  End If
50080
50090  Set PrinterProfiles = GetPrinterProfiles
50100  For i = 1 To PrinterProfiles.Count
50110   If StrComp(PrinterName, PrinterProfiles(i)(0)) = 0 Then
50120    cPrinterProfile = PrinterProfiles(i)(1)
50130   End If
50140  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPrinterProfile [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let cPrinterProfile(ByVal PrinterName As String, ByVal ProfileName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PrinterProfiles As Collection, PrinterProfiles2 As Collection, i As Long, sa(1) As String, _
  foundPrinter As Boolean
50030
50040  If Not PrinterIsInstalled(PrinterName) Then
50050   mError.Number = 15
50060   mError.Description = Replace$(ErrDescr16, "%1", PrinterName)
50070   RaiseEvent eError
50080   Exit Property
50090  End If
50100  ProfileName = Trim$(ProfileName)
50110  If ProfileExists(ProfileName) = False Then
50120   mError.Number = 11
50130   mError.Description = Replace$(ErrDescr11, "%1", ProfileName)
50140   RaiseEvent eError
50150   Exit Property
50160  End If
50170
50180  Set PrinterProfiles = GetPrinterProfiles
50190  Set PrinterProfiles2 = New Collection
50200
50210  foundPrinter = False
50220  For i = 1 To PrinterProfiles.Count
50230   sa(0) = PrinterProfiles(i)(0)
50240   If StrComp(sa(0), PrinterName) = 0 Then
50250     sa(1) = ProfileName
50260     foundPrinter = True
50270    Else
50280     sa(1) = PrinterProfiles(i)(1)
50290   End If
50300   PrinterProfiles2.Add sa
50310  Next i
50320  If foundPrinter = False Then
50330   sa(0) = PrinterName
50340   sa(1) = ProfileName
50350   PrinterProfiles2.Add sa
50360  End If
50370  SavePrinterProfiles PrinterProfiles2
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPrinterProfile [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

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
50010  cPDFCreatorApplicationPath = PDFCreatorApplicationPath
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
50010  cInstalledAsServer = CheckInstalledAsServer
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
50230   .Add "EditWithPDFArchitect"
50240   .Add "EPSLanguageLevel"
50250   .Add "FilenameSubstitutions"
50260   .Add "FilenameSubstitutionsOnlyInTitle"
50270   .Add "JPEGColorscount"
50280   .Add "JPEGQuality"
50290   .Add "JPEGResolution"
50300   .Add "Language"
50310   .Add "LastSaveDirectory"
50320   .Add "LastUpdateCheck"
50330   .Add "Logging"
50340   .Add "LogLines"
50350   .Add "MaximumCountOfPDFArchitectToolTip"
50360   .Add "NoConfirmMessageSwitchingDefaultprinter"
50370   .Add "NoProcessingAtStartup"
50380   .Add "NoPSCheck"
50390   .Add "OnePagePerFile"
50400   .Add "OpenOutputFile"
50410   .Add "OptionsDesign"
50420   .Add "OptionsEnabled"
50430   .Add "OptionsVisible"
50440   .Add "Papersize"
50450   .Add "PCLColorsCount"
50460   .Add "PCLResolution"
50470   .Add "PCXColorscount"
50480   .Add "PCXResolution"
50490   .Add "PDFAes128Encryption"
50500   .Add "PDFAllowAssembly"
50510   .Add "PDFAllowDegradedPrinting"
50520   .Add "PDFAllowFillIn"
50530   .Add "PDFAllowScreenReaders"
50540   .Add "PDFColorsCMYKToRGB"
50550   .Add "PDFColorsColorModel"
50560   .Add "PDFColorsPreserveHalftone"
50570   .Add "PDFColorsPreserveOverprint"
50580   .Add "PDFColorsPreserveTransfer"
50590   .Add "PDFCompressionColorCompression"
50600   .Add "PDFCompressionColorCompressionChoice"
50610   .Add "PDFCompressionColorCompressionJPEGHighFactor"
50620   .Add "PDFCompressionColorCompressionJPEGLowFactor"
50630   .Add "PDFCompressionColorCompressionJPEGManualFactor"
50640   .Add "PDFCompressionColorCompressionJPEGMaximumFactor"
50650   .Add "PDFCompressionColorCompressionJPEGMediumFactor"
50660   .Add "PDFCompressionColorCompressionJPEGMinimumFactor"
50670   .Add "PDFCompressionColorResample"
50680   .Add "PDFCompressionColorResampleChoice"
50690   .Add "PDFCompressionColorResolution"
50700   .Add "PDFCompressionGreyCompression"
50710   .Add "PDFCompressionGreyCompressionChoice"
50720   .Add "PDFCompressionGreyCompressionJPEGHighFactor"
50730   .Add "PDFCompressionGreyCompressionJPEGLowFactor"
50740   .Add "PDFCompressionGreyCompressionJPEGManualFactor"
50750   .Add "PDFCompressionGreyCompressionJPEGMaximumFactor"
50760   .Add "PDFCompressionGreyCompressionJPEGMediumFactor"
50770   .Add "PDFCompressionGreyCompressionJPEGMinimumFactor"
50780   .Add "PDFCompressionGreyResample"
50790   .Add "PDFCompressionGreyResampleChoice"
50800   .Add "PDFCompressionGreyResolution"
50810   .Add "PDFCompressionMonoCompression"
50820   .Add "PDFCompressionMonoCompressionChoice"
50830   .Add "PDFCompressionMonoResample"
50840   .Add "PDFCompressionMonoResampleChoice"
50850   .Add "PDFCompressionMonoResolution"
50860   .Add "PDFCompressionTextCompression"
50870   .Add "PDFDisallowCopy"
50880   .Add "PDFDisallowModifyAnnotations"
50890   .Add "PDFDisallowModifyContents"
50900   .Add "PDFDisallowPrinting"
50910   .Add "PDFEncryptor"
50920   .Add "PDFFontsEmbedAll"
50930   .Add "PDFFontsSubSetFonts"
50940   .Add "PDFFontsSubSetFontsPercent"
50950   .Add "PDFGeneralASCII85"
50960   .Add "PDFGeneralAutorotate"
50970   .Add "PDFGeneralCompatibility"
50980   .Add "PDFGeneralDefault"
50990   .Add "PDFGeneralOverprint"
51000   .Add "PDFGeneralResolution"
51010   .Add "PDFHighEncryption"
51020   .Add "PDFLowEncryption"
51030   .Add "PDFOptimize"
51040   .Add "PDFOwnerPass"
51050   .Add "PDFOwnerPasswordString"
51060   .Add "PDFSigningMultiSignature"
51070   .Add "PDFSigningPFXFile"
51080   .Add "PDFSigningPFXFilePassword"
51090   .Add "PDFSigningSignatureContact"
51100   .Add "PDFSigningSignatureLeftX"
51110   .Add "PDFSigningSignatureLeftY"
51120   .Add "PDFSigningSignatureLocation"
51130   .Add "PDFSigningSignatureOnPage"
51140   .Add "PDFSigningSignatureReason"
51150   .Add "PDFSigningSignatureRightX"
51160   .Add "PDFSigningSignatureRightY"
51170   .Add "PDFSigningSignatureVisible"
51180   .Add "PDFSigningSignPDF"
51190   .Add "PDFUpdateMetadata"
51200   .Add "PDFUserPass"
51210   .Add "PDFUserPasswordString"
51220   .Add "PDFUseSecurity"
51230   .Add "PNGColorscount"
51240   .Add "PNGResolution"
51250   .Add "PrintAfterSaving"
51260   .Add "PrintAfterSavingBitsPerPixel"
51270   .Add "PrintAfterSavingDuplex"
51280   .Add "PrintAfterSavingMaxResolution"
51290   .Add "PrintAfterSavingMaxResolutionEnabled"
51300   .Add "PrintAfterSavingNoCancel"
51310   .Add "PrintAfterSavingPrinter"
51320   .Add "PrintAfterSavingQueryUser"
51330   .Add "PrintAfterSavingTumble"
51340   .Add "PrinterStop"
51350   .Add "ProcessPriority"
51360   .Add "ProgramFont"
51370   .Add "ProgramFontCharset"
51380   .Add "ProgramFontSize"
51390   .Add "PSDColorsCount"
51400   .Add "PSDResolution"
51410   .Add "PSLanguageLevel"
51420   .Add "RAWColorsCount"
51430   .Add "RAWResolution"
51440   .Add "RemoveAllKnownFileExtensions"
51450   .Add "RemoveSpaces"
51460   .Add "RunProgramAfterSaving"
51470   .Add "RunProgramAfterSavingProgramname"
51480   .Add "RunProgramAfterSavingProgramParameters"
51490   .Add "RunProgramAfterSavingWaitUntilReady"
51500   .Add "RunProgramAfterSavingWindowstyle"
51510   .Add "RunProgramBeforeSaving"
51520   .Add "RunProgramBeforeSavingProgramname"
51530   .Add "RunProgramBeforeSavingProgramParameters"
51540   .Add "RunProgramBeforeSavingWindowstyle"
51550   .Add "SaveFilename"
51560   .Add "SendEmailAfterAutoSaving"
51570   .Add "SendMailMethod"
51580   .Add "ShowAnimation"
51590   .Add "StampFontColor"
51600   .Add "StampFontname"
51610   .Add "StampFontsize"
51620   .Add "StampOutlineFontthickness"
51630   .Add "StampString"
51640   .Add "StampUseOutlineFont"
51650   .Add "StandardAuthor"
51660   .Add "StandardCreationdate"
51670   .Add "StandardDateformat"
51680   .Add "StandardKeywords"
51690   .Add "StandardMailDomain"
51700   .Add "StandardModifydate"
51710   .Add "StandardSaveformat"
51720   .Add "StandardSubject"
51730   .Add "StandardTitle"
51740   .Add "SVGResolution"
51750   .Add "TIFFColorscount"
51760   .Add "TIFFResolution"
51770   .Add "Toolbars"
51780   .Add "UpdateInterval"
51790   .Add "UseAutosave"
51800   .Add "UseAutosaveDirectory"
51810   .Add "UseCreationDateNow"
51820   .Add "UseCustomPaperSize"
51830   .Add "UseFixPapersize"
51840   .Add "UseStandardAuthor"
51850  End With
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

Public Property Get cPostscriptInfo(ByVal PostScriptFilename As String, ByVal PropertyName As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PSHeader As tPSHeader
50020  If CheckForRunningProgramErrors Then
50030   RaiseEvent eError
50040   Exit Property
50050  End If
50060  If FileExists(PostScriptFilename) = False Then
50070   mError.Number = 5
50080   mError.Description = Replace$(ErrDescr5, "%1", PostScriptFilename)
50090   RaiseEvent eError
50100   Exit Property
50110  End If
50120  If FileInUse(PostScriptFilename) Then
50130   mError.Number = 7
50140   mError.Description = Replace$(ErrDescr7, "%1", PostScriptFilename)
50150   RaiseEvent eError
50160   Exit Property
50170  End If
50180  PSHeader = GetPSHeader(PostScriptFilename)
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

Public Property Let cDefaultPrinter(ByVal PrinterName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If CheckForRunningProgramErrors Then
50020   RaiseEvent eError
50030   Exit Property
50040  End If
50050  SetDefaultprinterInProg PrinterName
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
Public Function cIsAdministrator() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cIsAdministrator = IsAdmin
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cIsAdministrator")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function cPrinterIsInstalled(ByVal PrinterName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If PrinterIsInstalled(PrinterName) = True Then
50020    cPrinterIsInstalled = True
50030   Else
50040    cPrinterIsInstalled = False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cPrinterIsInstalled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function cAddPDFCreatorPrinter(ByVal PrinterName As String, Optional ByVal ProfileName As String = "") As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Not IsAdmin Then
50020   mError.Number = 17
50030   mError.Description = ErrDescr18
50040   RaiseEvent eError
50050   Exit Function
50060  End If
50070  If PrinterIsInstalled(PrinterName) = True Then
50080   mError.Number = 16
50090   mError.Description = Replace$(ErrDescr17, "%1", PrinterName)
50100   RaiseEvent eError
50110   Exit Function
50120  End If
50130  ProfileName = Trim$(ProfileName)
50140  If LenB(ProfileName) > 0 Then
50150   If ProfileExists(ProfileName) = False Then
50160    mError.Number = 11
50170    mError.Description = Replace$(ErrDescr11, "%1", ProfileName)
50180    RaiseEvent eError
50190    Exit Function
50200   End If
50210  End If
50220  cAddPDFCreatorPrinter = InstallPrinter(PrinterName, "PDFCreator", "PDFCreator:", "")
50230  If LenB(ProfileName) > 0 Then
50240   cPrinterProfile(PrinterName) = ProfileName
50250  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cAddPDFCreatorPrinter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function cProfileExists(ByVal ProfileName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cProfileExists = ProfileExists(ProfileName)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cProfileExists")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function cDeletePDFCreatorPrinter(ByVal PrinterName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PDFCreatorPrinters As Collection, i As Long
50020  If Not IsAdmin Then
50030   mError.Number = 17
50040   mError.Description = ErrDescr18
50050   RaiseEvent eError
50060   Exit Function
50070  End If
50080  If Not PrinterIsInstalled(PrinterName) Then
50090   mError.Number = 15
50100   mError.Description = Replace$(ErrDescr16, "%1", PrinterName)
50110   RaiseEvent eError
50120   Exit Function
50130  End If
50140  Set PDFCreatorPrinters = GetPDFCreatorPrinters()
50150
50160  For i = 1 To PDFCreatorPrinters.Count
50170   If StrComp(PDFCreatorPrinters(i), PrinterName, vbTextCompare) = 0 Then
50180    Exit For
50190   End If
50200  Next i
50210  If i > PDFCreatorPrinters.Count Then
50220   mError.Number = 18
50230   mError.Description = ErrDescr19
50240   RaiseEvent eError
50250   Exit Function
50260  End If
50270  cDeletePDFCreatorPrinter = UnInstallPrinter(PrinterName, "")
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cDeletePDFCreatorPrinter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function cGetProfileNames() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set cGetProfileNames = GetProfiles
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cGetProfileNames")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function cAddProfile(ByVal ProfileName As String, Optional ByVal Options1 As clsPDFCreatorOptions) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Profiles As Collection, i As Long
50020  ProfileName = Trim$(ProfileName)
50030  If LenB(ProfileName) = 0 Then
50040   mError.Number = 13
50050   mError.Description = ErrDescr13
50060   RaiseEvent eError
50070   Exit Function
50080  End If
50090  If ProfileExists(ProfileName) = True Then
50100   mError.Number = 12
50110   mError.Description = Replace$(ErrDescr12, "%1", ProfileName)
50120   RaiseEvent eError
50130   Exit Function
50140  End If
50150  If Options1 Is Nothing Then
50160   Set Options1 = cStandardOptions
50170  End If
50180  SaveOptions cLetOptions(Options1), ProfileName
50190  If ProfileExists(ProfileName) = True Then
50200   cAddProfile = True
50210  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cAddProfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function cRenameProfile(ByVal OldProfileName As String, ByVal NewProfileName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Options1 As clsPDFCreatorOptions, tmpOptions As tOptions
50020  OldProfileName = Trim$(OldProfileName)
50030  NewProfileName = Trim$(NewProfileName)
50040  If (LenB(OldProfileName) = 0) Or (LenB(NewProfileName) = 0) Then
50050   mError.Number = 13
50060   mError.Description = ErrDescr13
50070   RaiseEvent eError
50080   Exit Function
50090  End If
50100  If ProfileExists(OldProfileName) = False Then
50110   mError.Number = 11
50120   mError.Description = Replace$(ErrDescr11, "%1", OldProfileName)
50130   RaiseEvent eError
50140   Exit Function
50150  End If
50160  If StrComp(OldProfileName, NewProfileName, vbTextCompare) = 0 Then
50170   mError.Number = 15
50180   mError.Description = Replace$(Replace$(ErrDescr15, "%1", OldProfileName), "%2", NewProfileName)
50190   RaiseEvent eError
50200   Exit Function
50210  End If
50220  If ProfileExists(NewProfileName) = True Then
50230   mError.Number = 12
50240   mError.Description = Replace$(ErrDescr12, "%1", NewProfileName)
50250   RaiseEvent eError
50260   Exit Function
50270  End If
50280
50290  SaveOptions ReadOptions(, , OldProfileName), NewProfileName
50300  DeleteProfile OldProfileName
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cRenameProfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function cDeleteProfile(ByVal ProfileName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim errStr As String, sProfileAssociatedPrinters As String
50020  ProfileName = Trim$(ProfileName)
50030  If LenB(ProfileName) = 0 Then
50040   mError.Number = 13
50050   mError.Description = ErrDescr13
50060   RaiseEvent eError
50070   Exit Function
50080  End If
50090  If ProfileExists(ProfileName) = False Then
50100   mError.Number = 11
50110   mError.Description = Replace$(ErrDescr11, "%1", ProfileName)
50120   RaiseEvent eError
50130   Exit Function
50140  End If
50150
50160  errStr = ErrDescr14
50170  sProfileAssociatedPrinters = ProfileAssociatedPrinters(ProfileName)
50180  If LenB(sProfileAssociatedPrinters) > 0 Then
50190   mError.Number = 14
50200   If InStr(1, sProfileAssociatedPrinters, ",", vbTextCompare) > 0 Then
50210    errStr = Replace$(errStr, "printer", "printers", , vbTextCompare)
50220    mError.Description = errStr
50230   End If
50240   RaiseEvent eError
50250   Exit Function
50260  End If
50270
50280  DeleteProfile ProfileName
50290
50300  If ProfileExists(ProfileName) = False Then
50310   cDeleteProfile = True
50320  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cDeleteProfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub cAddPrintjob(ByVal filename As String) ' Supported file types: Postscript, JPEG, BMP
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ivgf As Boolean, spoolDirectory As String, psFileName As String, strGUID As String, File As String
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
50200    spoolDirectory = GetPDFCreatorSpoolDirectory
50210    strGUID = GetGUID
50220    File = spoolDirectory & strGUID
50230    psFileName = File & ".ps"
50240    If ivgf Then
50250      Image2PS filename, psFileName
50260     Else
50270      FileCopy filename, psFileName
50280    End If
50290    CreateInfoSpoolFile psFileName, File & ".inf"
50300   Else
50310    mError.Number = 9
50320    mError.Description = Replace$(ErrDescr9, "%1", filename)
50330    RaiseEvent eError
50340    Exit Sub
50350  End If
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
50140  Set eActionTimer = Nothing
50150
50160  For Each f In Forms
50170   If Not f Is Nothing Then
50180    If UCase$(f.Name) <> "FRMMAIN" Then
50190     Unload f
50200     cClose = True
50210    End If
50220   End If
50230  Next f
50240
50250  If Not mForm Is Nothing Then
50260   Unload mForm
50270   Set mForm = Nothing
50280  End If
50290  ProgramIsStarted = False
50300  cClose = True
50310
50320  Set mError = Nothing
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
50110  mProgramIsRunning = False
50120  StartProgram Params
50130  Set mForm = frmMain
50140  Set eActionTimer = mForm.Timer3
50150  eActionTimer.Interval = 10
50160  eActionTimer.Enabled = True
50170  cStart = True
50180  cVisible = ProgramIsVisible
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
50060   mForm.DocumentCombine True
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

Public Function cGetPrinterProfiles() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set cGetPrinterProfiles = GetPrinterProfiles
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreator", "cGetPrinterProfiles")
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
50050  If FileExists(filename) = False Then
50060   mError.Number = 5
50070   mError.Description = Replace$(ErrDescr5, "%1", filename)
50080   RaiseEvent eError
50090   Exit Sub
50100  End If
50110  PrintFile filename
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

Public Sub cPrintPrinterTestpage(Optional ByVal PrinterName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tPrintername As String
50020  If Not PrinterIsInstalled(PrinterName) Then
50030   mError.Number = 15
50040   mError.Description = Replace$(ErrDescr16, "%1", PrinterName)
50050   RaiseEvent eError
50060   Exit Sub
50070  End If
50080  If LenB(PrinterName) = 0 Then
50090    tPrintername = cDefaultPrinter
50100   Else
50110    tPrintername = PrinterName
50120  End If
50130  ShellAndWait 0, "open", "RUNDLL32.EXE", "PRINTUI.DLL,PrintUIEntry /n """ & tPrintername & """ /k", PDFCreatorApplicationPath, wNormal, WCNone
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
50180  Set cReadOptionsFromFile = cGetOptions(ReadOptionsINI(dummyOptions, INIFilename, True))
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
50220   Case "EDITWITHPDFARCHITECT": cGetOptionFromOptions = Options.EditWithPDFArchitect
50230   Case "EPSLANGUAGELEVEL": cGetOptionFromOptions = Options.EPSLanguageLevel
50240   Case "FILENAMESUBSTITUTIONS": cGetOptionFromOptions = Options.FilenameSubstitutions
50250   Case "FILENAMESUBSTITUTIONSONLYINTITLE": cGetOptionFromOptions = Options.FilenameSubstitutionsOnlyInTitle
50260   Case "JPEGCOLORSCOUNT": cGetOptionFromOptions = Options.JPEGColorscount
50270   Case "JPEGQUALITY": cGetOptionFromOptions = Options.JPEGQuality
50280   Case "JPEGRESOLUTION": cGetOptionFromOptions = Options.JPEGResolution
50290   Case "LANGUAGE": cGetOptionFromOptions = Options.Language
50300   Case "LASTSAVEDIRECTORY": cGetOptionFromOptions = Options.LastSaveDirectory
50310   Case "LASTUPDATECHECK": cGetOptionFromOptions = Options.LastUpdateCheck
50320   Case "LOGGING": cGetOptionFromOptions = Options.Logging
50330   Case "LOGLINES": cGetOptionFromOptions = Options.LogLines
50340   Case "MAXIMUMCOUNTOFPDFARCHITECTTOOLTIP": cGetOptionFromOptions = Options.MaximumCountOfPDFArchitectToolTip
50350   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": cGetOptionFromOptions = Options.NoConfirmMessageSwitchingDefaultprinter
50360   Case "NOPROCESSINGATSTARTUP": cGetOptionFromOptions = Options.NoProcessingAtStartup
50370   Case "NOPSCHECK": cGetOptionFromOptions = Options.NoPSCheck
50380   Case "ONEPAGEPERFILE": cGetOptionFromOptions = Options.OnePagePerFile
50390   Case "OPENOUTPUTFILE": cGetOptionFromOptions = Options.OpenOutputFile
50400   Case "OPTIONSDESIGN": cGetOptionFromOptions = Options.OptionsDesign
50410   Case "OPTIONSENABLED": cGetOptionFromOptions = Options.OptionsEnabled
50420   Case "OPTIONSVISIBLE": cGetOptionFromOptions = Options.OptionsVisible
50430   Case "PAPERSIZE": cGetOptionFromOptions = Options.Papersize
50440   Case "PCLCOLORSCOUNT": cGetOptionFromOptions = Options.PCLColorsCount
50450   Case "PCLRESOLUTION": cGetOptionFromOptions = Options.PCLResolution
50460   Case "PCXCOLORSCOUNT": cGetOptionFromOptions = Options.PCXColorscount
50470   Case "PCXRESOLUTION": cGetOptionFromOptions = Options.PCXResolution
50480   Case "PDFAES128ENCRYPTION": cGetOptionFromOptions = Options.PDFAes128Encryption
50490   Case "PDFALLOWASSEMBLY": cGetOptionFromOptions = Options.PDFAllowAssembly
50500   Case "PDFALLOWDEGRADEDPRINTING": cGetOptionFromOptions = Options.PDFAllowDegradedPrinting
50510   Case "PDFALLOWFILLIN": cGetOptionFromOptions = Options.PDFAllowFillIn
50520   Case "PDFALLOWSCREENREADERS": cGetOptionFromOptions = Options.PDFAllowScreenReaders
50530   Case "PDFCOLORSCMYKTORGB": cGetOptionFromOptions = Options.PDFColorsCMYKToRGB
50540   Case "PDFCOLORSCOLORMODEL": cGetOptionFromOptions = Options.PDFColorsColorModel
50550   Case "PDFCOLORSPRESERVEHALFTONE": cGetOptionFromOptions = Options.PDFColorsPreserveHalftone
50560   Case "PDFCOLORSPRESERVEOVERPRINT": cGetOptionFromOptions = Options.PDFColorsPreserveOverprint
50570   Case "PDFCOLORSPRESERVETRANSFER": cGetOptionFromOptions = Options.PDFColorsPreserveTransfer
50580   Case "PDFCOMPRESSIONCOLORCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionColorCompression
50590   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionColorCompressionChoice
50600   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGHighFactor
50610   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGLowFactor
50620   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMANUALFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGManualFactor
50630   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMaximumFactor
50640   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMediumFactor
50650   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMinimumFactor
50660   Case "PDFCOMPRESSIONCOLORRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionColorResample
50670   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionColorResampleChoice
50680   Case "PDFCOMPRESSIONCOLORRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionColorResolution
50690   Case "PDFCOMPRESSIONGREYCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionGreyCompression
50700   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionChoice
50710   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGHighFactor
50720   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGLowFactor
50730   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMANUALFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGManualFactor
50740   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMaximumFactor
50750   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMediumFactor
50760   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMinimumFactor
50770   Case "PDFCOMPRESSIONGREYRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionGreyResample
50780   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyResampleChoice
50790   Case "PDFCOMPRESSIONGREYRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionGreyResolution
50800   Case "PDFCOMPRESSIONMONOCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionMonoCompression
50810   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoCompressionChoice
50820   Case "PDFCOMPRESSIONMONORESAMPLE": cGetOptionFromOptions = Options.PDFCompressionMonoResample
50830   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoResampleChoice
50840   Case "PDFCOMPRESSIONMONORESOLUTION": cGetOptionFromOptions = Options.PDFCompressionMonoResolution
50850   Case "PDFCOMPRESSIONTEXTCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionTextCompression
50860   Case "PDFDISALLOWCOPY": cGetOptionFromOptions = Options.PDFDisallowCopy
50870   Case "PDFDISALLOWMODIFYANNOTATIONS": cGetOptionFromOptions = Options.PDFDisallowModifyAnnotations
50880   Case "PDFDISALLOWMODIFYCONTENTS": cGetOptionFromOptions = Options.PDFDisallowModifyContents
50890   Case "PDFDISALLOWPRINTING": cGetOptionFromOptions = Options.PDFDisallowPrinting
50900   Case "PDFENCRYPTOR": cGetOptionFromOptions = Options.PDFEncryptor
50910   Case "PDFFONTSEMBEDALL": cGetOptionFromOptions = Options.PDFFontsEmbedAll
50920   Case "PDFFONTSSUBSETFONTS": cGetOptionFromOptions = Options.PDFFontsSubSetFonts
50930   Case "PDFFONTSSUBSETFONTSPERCENT": cGetOptionFromOptions = Options.PDFFontsSubSetFontsPercent
50940   Case "PDFGENERALASCII85": cGetOptionFromOptions = Options.PDFGeneralASCII85
50950   Case "PDFGENERALAUTOROTATE": cGetOptionFromOptions = Options.PDFGeneralAutorotate
50960   Case "PDFGENERALCOMPATIBILITY": cGetOptionFromOptions = Options.PDFGeneralCompatibility
50970   Case "PDFGENERALDEFAULT": cGetOptionFromOptions = Options.PDFGeneralDefault
50980   Case "PDFGENERALOVERPRINT": cGetOptionFromOptions = Options.PDFGeneralOverprint
50990   Case "PDFGENERALRESOLUTION": cGetOptionFromOptions = Options.PDFGeneralResolution
51000   Case "PDFHIGHENCRYPTION": cGetOptionFromOptions = Options.PDFHighEncryption
51010   Case "PDFLOWENCRYPTION": cGetOptionFromOptions = Options.PDFLowEncryption
51020   Case "PDFOPTIMIZE": cGetOptionFromOptions = Options.PDFOptimize
51030   Case "PDFOWNERPASS": cGetOptionFromOptions = Options.PDFOwnerPass
51040   Case "PDFOWNERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFOwnerPasswordString
51050   Case "PDFSIGNINGMULTISIGNATURE": cGetOptionFromOptions = Options.PDFSigningMultiSignature
51060   Case "PDFSIGNINGPFXFILE": cGetOptionFromOptions = Options.PDFSigningPFXFile
51070   Case "PDFSIGNINGPFXFILEPASSWORD": cGetOptionFromOptions = Options.PDFSigningPFXFilePassword
51080   Case "PDFSIGNINGSIGNATURECONTACT": cGetOptionFromOptions = Options.PDFSigningSignatureContact
51090   Case "PDFSIGNINGSIGNATURELEFTX": cGetOptionFromOptions = Options.PDFSigningSignatureLeftX
51100   Case "PDFSIGNINGSIGNATURELEFTY": cGetOptionFromOptions = Options.PDFSigningSignatureLeftY
51110   Case "PDFSIGNINGSIGNATURELOCATION": cGetOptionFromOptions = Options.PDFSigningSignatureLocation
51120   Case "PDFSIGNINGSIGNATUREONPAGE": cGetOptionFromOptions = Options.PDFSigningSignatureOnPage
51130   Case "PDFSIGNINGSIGNATUREREASON": cGetOptionFromOptions = Options.PDFSigningSignatureReason
51140   Case "PDFSIGNINGSIGNATURERIGHTX": cGetOptionFromOptions = Options.PDFSigningSignatureRightX
51150   Case "PDFSIGNINGSIGNATURERIGHTY": cGetOptionFromOptions = Options.PDFSigningSignatureRightY
51160   Case "PDFSIGNINGSIGNATUREVISIBLE": cGetOptionFromOptions = Options.PDFSigningSignatureVisible
51170   Case "PDFSIGNINGSIGNPDF": cGetOptionFromOptions = Options.PDFSigningSignPDF
51180   Case "PDFUPDATEMETADATA": cGetOptionFromOptions = Options.PDFUpdateMetadata
51190   Case "PDFUSERPASS": cGetOptionFromOptions = Options.PDFUserPass
51200   Case "PDFUSERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFUserPasswordString
51210   Case "PDFUSESECURITY": cGetOptionFromOptions = Options.PDFUseSecurity
51220   Case "PNGCOLORSCOUNT": cGetOptionFromOptions = Options.PNGColorscount
51230   Case "PNGRESOLUTION": cGetOptionFromOptions = Options.PNGResolution
51240   Case "PRINTAFTERSAVING": cGetOptionFromOptions = Options.PrintAfterSaving
51250   Case "PRINTAFTERSAVINGBITSPERPIXEL": cGetOptionFromOptions = Options.PrintAfterSavingBitsPerPixel
51260   Case "PRINTAFTERSAVINGDUPLEX": cGetOptionFromOptions = Options.PrintAfterSavingDuplex
51270   Case "PRINTAFTERSAVINGMAXRESOLUTION": cGetOptionFromOptions = Options.PrintAfterSavingMaxResolution
51280   Case "PRINTAFTERSAVINGMAXRESOLUTIONENABLED": cGetOptionFromOptions = Options.PrintAfterSavingMaxResolutionEnabled
51290   Case "PRINTAFTERSAVINGNOCANCEL": cGetOptionFromOptions = Options.PrintAfterSavingNoCancel
51300   Case "PRINTAFTERSAVINGPRINTER": cGetOptionFromOptions = Options.PrintAfterSavingPrinter
51310   Case "PRINTAFTERSAVINGQUERYUSER": cGetOptionFromOptions = Options.PrintAfterSavingQueryUser
51320   Case "PRINTAFTERSAVINGTUMBLE": cGetOptionFromOptions = Options.PrintAfterSavingTumble
51330   Case "PRINTERSTOP": cGetOptionFromOptions = Options.PrinterStop
51340   Case "PROCESSPRIORITY": cGetOptionFromOptions = Options.ProcessPriority
51350   Case "PROGRAMFONT": cGetOptionFromOptions = Options.ProgramFont
51360   Case "PROGRAMFONTCHARSET": cGetOptionFromOptions = Options.ProgramFontCharset
51370   Case "PROGRAMFONTSIZE": cGetOptionFromOptions = Options.ProgramFontSize
51380   Case "PSDCOLORSCOUNT": cGetOptionFromOptions = Options.PSDColorsCount
51390   Case "PSDRESOLUTION": cGetOptionFromOptions = Options.PSDResolution
51400   Case "PSLANGUAGELEVEL": cGetOptionFromOptions = Options.PSLanguageLevel
51410   Case "RAWCOLORSCOUNT": cGetOptionFromOptions = Options.RAWColorsCount
51420   Case "RAWRESOLUTION": cGetOptionFromOptions = Options.RAWResolution
51430   Case "REMOVEALLKNOWNFILEEXTENSIONS": cGetOptionFromOptions = Options.RemoveAllKnownFileExtensions
51440   Case "REMOVESPACES": cGetOptionFromOptions = Options.RemoveSpaces
51450   Case "RUNPROGRAMAFTERSAVING": cGetOptionFromOptions = Options.RunProgramAfterSaving
51460   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramname
51470   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramParameters
51480   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": cGetOptionFromOptions = Options.RunProgramAfterSavingWaitUntilReady
51490   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramAfterSavingWindowstyle
51500   Case "RUNPROGRAMBEFORESAVING": cGetOptionFromOptions = Options.RunProgramBeforeSaving
51510   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramname
51520   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramParameters
51530   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramBeforeSavingWindowstyle
51540   Case "SAVEFILENAME": cGetOptionFromOptions = Options.SaveFilename
51550   Case "SENDEMAILAFTERAUTOSAVING": cGetOptionFromOptions = Options.SendEmailAfterAutoSaving
51560   Case "SENDMAILMETHOD": cGetOptionFromOptions = Options.SendMailMethod
51570   Case "SHOWANIMATION": cGetOptionFromOptions = Options.ShowAnimation
51580   Case "STAMPFONTCOLOR": cGetOptionFromOptions = Options.StampFontColor
51590   Case "STAMPFONTNAME": cGetOptionFromOptions = Options.StampFontname
51600   Case "STAMPFONTSIZE": cGetOptionFromOptions = Options.StampFontsize
51610   Case "STAMPOUTLINEFONTTHICKNESS": cGetOptionFromOptions = Options.StampOutlineFontthickness
51620   Case "STAMPSTRING": cGetOptionFromOptions = Options.StampString
51630   Case "STAMPUSEOUTLINEFONT": cGetOptionFromOptions = Options.StampUseOutlineFont
51640   Case "STANDARDAUTHOR": cGetOptionFromOptions = Options.StandardAuthor
51650   Case "STANDARDCREATIONDATE": cGetOptionFromOptions = Options.StandardCreationdate
51660   Case "STANDARDDATEFORMAT": cGetOptionFromOptions = Options.StandardDateformat
51670   Case "STANDARDKEYWORDS": cGetOptionFromOptions = Options.StandardKeywords
51680   Case "STANDARDMAILDOMAIN": cGetOptionFromOptions = Options.StandardMailDomain
51690   Case "STANDARDMODIFYDATE": cGetOptionFromOptions = Options.StandardModifydate
51700   Case "STANDARDSAVEFORMAT": cGetOptionFromOptions = Options.StandardSaveformat
51710   Case "STANDARDSUBJECT": cGetOptionFromOptions = Options.StandardSubject
51720   Case "STANDARDTITLE": cGetOptionFromOptions = Options.StandardTitle
51730   Case "SVGRESOLUTION": cGetOptionFromOptions = Options.SVGResolution
51740   Case "TIFFCOLORSCOUNT": cGetOptionFromOptions = Options.TIFFColorscount
51750   Case "TIFFRESOLUTION": cGetOptionFromOptions = Options.TIFFResolution
51760   Case "TOOLBARS": cGetOptionFromOptions = Options.Toolbars
51770   Case "UPDATEINTERVAL": cGetOptionFromOptions = Options.UpdateInterval
51780   Case "USEAUTOSAVE": cGetOptionFromOptions = Options.UseAutosave
51790   Case "USEAUTOSAVEDIRECTORY": cGetOptionFromOptions = Options.UseAutosaveDirectory
51800   Case "USECREATIONDATENOW": cGetOptionFromOptions = Options.UseCreationDateNow
51810   Case "USECUSTOMPAPERSIZE": cGetOptionFromOptions = Options.UseCustomPaperSize
51820   Case "USEFIXPAPERSIZE": cGetOptionFromOptions = Options.UseFixPapersize
51830   Case "USESTANDARDAUTHOR": cGetOptionFromOptions = Options.UseStandardAuthor
51840  End Select
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
50230   .EditWithPDFArchitect = sOptions.EditWithPDFArchitect
50240   .EPSLanguageLevel = sOptions.EPSLanguageLevel
50250   .FilenameSubstitutions = sOptions.FilenameSubstitutions
50260   .FilenameSubstitutionsOnlyInTitle = sOptions.FilenameSubstitutionsOnlyInTitle
50270   .JPEGColorscount = sOptions.JPEGColorscount
50280   .JPEGQuality = sOptions.JPEGQuality
50290   .JPEGResolution = sOptions.JPEGResolution
50300   .Language = sOptions.Language
50310   .LastSaveDirectory = sOptions.LastSaveDirectory
50320   .LastUpdateCheck = sOptions.LastUpdateCheck
50330   .Logging = sOptions.Logging
50340   .LogLines = sOptions.LogLines
50350   .MaximumCountOfPDFArchitectToolTip = sOptions.MaximumCountOfPDFArchitectToolTip
50360   .NoConfirmMessageSwitchingDefaultprinter = sOptions.NoConfirmMessageSwitchingDefaultprinter
50370   .NoProcessingAtStartup = sOptions.NoProcessingAtStartup
50380   .NoPSCheck = sOptions.NoPSCheck
50390   .OnePagePerFile = sOptions.OnePagePerFile
50400   .OpenOutputFile = sOptions.OpenOutputFile
50410   .OptionsDesign = sOptions.OptionsDesign
50420   .OptionsEnabled = sOptions.OptionsEnabled
50430   .OptionsVisible = sOptions.OptionsVisible
50440   .Papersize = sOptions.Papersize
50450   .PCLColorsCount = sOptions.PCLColorsCount
50460   .PCLResolution = sOptions.PCLResolution
50470   .PCXColorscount = sOptions.PCXColorscount
50480   .PCXResolution = sOptions.PCXResolution
50490   .PDFAes128Encryption = sOptions.PDFAes128Encryption
50500   .PDFAllowAssembly = sOptions.PDFAllowAssembly
50510   .PDFAllowDegradedPrinting = sOptions.PDFAllowDegradedPrinting
50520   .PDFAllowFillIn = sOptions.PDFAllowFillIn
50530   .PDFAllowScreenReaders = sOptions.PDFAllowScreenReaders
50540   .PDFColorsCMYKToRGB = sOptions.PDFColorsCMYKToRGB
50550   .PDFColorsColorModel = sOptions.PDFColorsColorModel
50560   .PDFColorsPreserveHalftone = sOptions.PDFColorsPreserveHalftone
50570   .PDFColorsPreserveOverprint = sOptions.PDFColorsPreserveOverprint
50580   .PDFColorsPreserveTransfer = sOptions.PDFColorsPreserveTransfer
50590   .PDFCompressionColorCompression = sOptions.PDFCompressionColorCompression
50600   .PDFCompressionColorCompressionChoice = sOptions.PDFCompressionColorCompressionChoice
50610   .PDFCompressionColorCompressionJPEGHighFactor = sOptions.PDFCompressionColorCompressionJPEGHighFactor
50620   .PDFCompressionColorCompressionJPEGLowFactor = sOptions.PDFCompressionColorCompressionJPEGLowFactor
50630   .PDFCompressionColorCompressionJPEGManualFactor = sOptions.PDFCompressionColorCompressionJPEGManualFactor
50640   .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
50650   .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
50660   .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
50670   .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
50680   .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
50690   .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
50700   .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
50710   .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
50720   .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
50730   .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
50740   .PDFCompressionGreyCompressionJPEGManualFactor = sOptions.PDFCompressionGreyCompressionJPEGManualFactor
50750   .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
50760   .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
50770   .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
50780   .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
50790   .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
50800   .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
50810   .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
50820   .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
50830   .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
50840   .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
50850   .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
50860   .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
50870   .PDFDisallowCopy = sOptions.PDFDisallowCopy
50880   .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
50890   .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
50900   .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
50910   .PDFEncryptor = sOptions.PDFEncryptor
50920   .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
50930   .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
50940   .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
50950   .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
50960   .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
50970   .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
50980   .PDFGeneralDefault = sOptions.PDFGeneralDefault
50990   .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
51000   .PDFGeneralResolution = sOptions.PDFGeneralResolution
51010   .PDFHighEncryption = sOptions.PDFHighEncryption
51020   .PDFLowEncryption = sOptions.PDFLowEncryption
51030   .PDFOptimize = sOptions.PDFOptimize
51040   .PDFOwnerPass = sOptions.PDFOwnerPass
51050   .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
51060   .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
51070   .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
51080   .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
51090   .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
51100   .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
51110   .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
51120   .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
51130   .PDFSigningSignatureOnPage = sOptions.PDFSigningSignatureOnPage
51140   .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
51150   .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
51160   .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
51170   .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
51180   .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
51190   .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
51200   .PDFUserPass = sOptions.PDFUserPass
51210   .PDFUserPasswordString = sOptions.PDFUserPasswordString
51220   .PDFUseSecurity = sOptions.PDFUseSecurity
51230   .PNGColorscount = sOptions.PNGColorscount
51240   .PNGResolution = sOptions.PNGResolution
51250   .PrintAfterSaving = sOptions.PrintAfterSaving
51260   .PrintAfterSavingBitsPerPixel = sOptions.PrintAfterSavingBitsPerPixel
51270   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
51280   .PrintAfterSavingMaxResolution = sOptions.PrintAfterSavingMaxResolution
51290   .PrintAfterSavingMaxResolutionEnabled = sOptions.PrintAfterSavingMaxResolutionEnabled
51300   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
51310   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
51320   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51330   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51340   .PrinterStop = sOptions.PrinterStop
51350   .ProcessPriority = sOptions.ProcessPriority
51360   .ProgramFont = sOptions.ProgramFont
51370   .ProgramFontCharset = sOptions.ProgramFontCharset
51380   .ProgramFontSize = sOptions.ProgramFontSize
51390   .PSDColorsCount = sOptions.PSDColorsCount
51400   .PSDResolution = sOptions.PSDResolution
51410   .PSLanguageLevel = sOptions.PSLanguageLevel
51420   .RAWColorsCount = sOptions.RAWColorsCount
51430   .RAWResolution = sOptions.RAWResolution
51440   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51450   .RemoveSpaces = sOptions.RemoveSpaces
51460   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51470   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51480   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51490   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51500   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51510   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51520   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51530   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51540   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51550   .SaveFilename = sOptions.SaveFilename
51560   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51570   .SendMailMethod = sOptions.SendMailMethod
51580   .ShowAnimation = sOptions.ShowAnimation
51590   .StampFontColor = sOptions.StampFontColor
51600   .StampFontname = sOptions.StampFontname
51610   .StampFontsize = sOptions.StampFontsize
51620   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51630   .StampString = sOptions.StampString
51640   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51650   .StandardAuthor = sOptions.StandardAuthor
51660   .StandardCreationdate = sOptions.StandardCreationdate
51670   .StandardDateformat = sOptions.StandardDateformat
51680   .StandardKeywords = sOptions.StandardKeywords
51690   .StandardMailDomain = sOptions.StandardMailDomain
51700   .StandardModifydate = sOptions.StandardModifydate
51710   .StandardSaveformat = sOptions.StandardSaveformat
51720   .StandardSubject = sOptions.StandardSubject
51730   .StandardTitle = sOptions.StandardTitle
51740   .SVGResolution = sOptions.SVGResolution
51750   .TIFFColorscount = sOptions.TIFFColorscount
51760   .TIFFResolution = sOptions.TIFFResolution
51770   .Toolbars = sOptions.Toolbars
51780   .UpdateInterval = sOptions.UpdateInterval
51790   .UseAutosave = sOptions.UseAutosave
51800   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51810   .UseCreationDateNow = sOptions.UseCreationDateNow
51820   .UseCustomPaperSize = sOptions.UseCustomPaperSize
51830   .UseFixPapersize = sOptions.UseFixPapersize
51840   .UseStandardAuthor = sOptions.UseStandardAuthor
51850  End With
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
50220   Case "EDITWITHPDFARCHITECT": Options1.EditWithPDFArchitect = value
50230   Case "EPSLANGUAGELEVEL": Options1.EPSLanguageLevel = value
50240   Case "FILENAMESUBSTITUTIONS": Options1.FilenameSubstitutions = value
50250   Case "FILENAMESUBSTITUTIONSONLYINTITLE": Options1.FilenameSubstitutionsOnlyInTitle = value
50260   Case "JPEGCOLORSCOUNT": Options1.JPEGColorscount = value
50270   Case "JPEGQUALITY": Options1.JPEGQuality = value
50280   Case "JPEGRESOLUTION": Options1.JPEGResolution = value
50290   Case "LANGUAGE": Options1.Language = value
50300   Case "LASTSAVEDIRECTORY": Options1.LastSaveDirectory = value
50310   Case "LASTUPDATECHECK": Options1.LastUpdateCheck = value
50320   Case "LOGGING": Options1.Logging = value
50330   Case "LOGLINES": Options1.LogLines = value
50340   Case "MAXIMUMCOUNTOFPDFARCHITECTTOOLTIP": Options1.MaximumCountOfPDFArchitectToolTip = value
50350   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": Options1.NoConfirmMessageSwitchingDefaultprinter = value
50360   Case "NOPROCESSINGATSTARTUP": Options1.NoProcessingAtStartup = value
50370   Case "NOPSCHECK": Options1.NoPSCheck = value
50380   Case "ONEPAGEPERFILE": Options1.OnePagePerFile = value
50390   Case "OPENOUTPUTFILE": Options1.OpenOutputFile = value
50400   Case "OPTIONSDESIGN": Options1.OptionsDesign = value
50410   Case "OPTIONSENABLED": Options1.OptionsEnabled = value
50420   Case "OPTIONSVISIBLE": Options1.OptionsVisible = value
50430   Case "PAPERSIZE": Options1.Papersize = value
50440   Case "PCLCOLORSCOUNT": Options1.PCLColorsCount = value
50450   Case "PCLRESOLUTION": Options1.PCLResolution = value
50460   Case "PCXCOLORSCOUNT": Options1.PCXColorscount = value
50470   Case "PCXRESOLUTION": Options1.PCXResolution = value
50480   Case "PDFAES128ENCRYPTION": Options1.PDFAes128Encryption = value
50490   Case "PDFALLOWASSEMBLY": Options1.PDFAllowAssembly = value
50500   Case "PDFALLOWDEGRADEDPRINTING": Options1.PDFAllowDegradedPrinting = value
50510   Case "PDFALLOWFILLIN": Options1.PDFAllowFillIn = value
50520   Case "PDFALLOWSCREENREADERS": Options1.PDFAllowScreenReaders = value
50530   Case "PDFCOLORSCMYKTORGB": Options1.PDFColorsCMYKToRGB = value
50540   Case "PDFCOLORSCOLORMODEL": Options1.PDFColorsColorModel = value
50550   Case "PDFCOLORSPRESERVEHALFTONE": Options1.PDFColorsPreserveHalftone = value
50560   Case "PDFCOLORSPRESERVEOVERPRINT": Options1.PDFColorsPreserveOverprint = value
50570   Case "PDFCOLORSPRESERVETRANSFER": Options1.PDFColorsPreserveTransfer = value
50580   Case "PDFCOMPRESSIONCOLORCOMPRESSION": Options1.PDFCompressionColorCompression = value
50590   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": Options1.PDFCompressionColorCompressionChoice = value
50600   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": Options1.PDFCompressionColorCompressionJPEGHighFactor = value
50610   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": Options1.PDFCompressionColorCompressionJPEGLowFactor = value
50620   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMANUALFACTOR": Options1.PDFCompressionColorCompressionJPEGManualFactor = value
50630   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": Options1.PDFCompressionColorCompressionJPEGMaximumFactor = value
50640   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": Options1.PDFCompressionColorCompressionJPEGMediumFactor = value
50650   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": Options1.PDFCompressionColorCompressionJPEGMinimumFactor = value
50660   Case "PDFCOMPRESSIONCOLORRESAMPLE": Options1.PDFCompressionColorResample = value
50670   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": Options1.PDFCompressionColorResampleChoice = value
50680   Case "PDFCOMPRESSIONCOLORRESOLUTION": Options1.PDFCompressionColorResolution = value
50690   Case "PDFCOMPRESSIONGREYCOMPRESSION": Options1.PDFCompressionGreyCompression = value
50700   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": Options1.PDFCompressionGreyCompressionChoice = value
50710   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": Options1.PDFCompressionGreyCompressionJPEGHighFactor = value
50720   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": Options1.PDFCompressionGreyCompressionJPEGLowFactor = value
50730   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMANUALFACTOR": Options1.PDFCompressionGreyCompressionJPEGManualFactor = value
50740   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": Options1.PDFCompressionGreyCompressionJPEGMaximumFactor = value
50750   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": Options1.PDFCompressionGreyCompressionJPEGMediumFactor = value
50760   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": Options1.PDFCompressionGreyCompressionJPEGMinimumFactor = value
50770   Case "PDFCOMPRESSIONGREYRESAMPLE": Options1.PDFCompressionGreyResample = value
50780   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": Options1.PDFCompressionGreyResampleChoice = value
50790   Case "PDFCOMPRESSIONGREYRESOLUTION": Options1.PDFCompressionGreyResolution = value
50800   Case "PDFCOMPRESSIONMONOCOMPRESSION": Options1.PDFCompressionMonoCompression = value
50810   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": Options1.PDFCompressionMonoCompressionChoice = value
50820   Case "PDFCOMPRESSIONMONORESAMPLE": Options1.PDFCompressionMonoResample = value
50830   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": Options1.PDFCompressionMonoResampleChoice = value
50840   Case "PDFCOMPRESSIONMONORESOLUTION": Options1.PDFCompressionMonoResolution = value
50850   Case "PDFCOMPRESSIONTEXTCOMPRESSION": Options1.PDFCompressionTextCompression = value
50860   Case "PDFDISALLOWCOPY": Options1.PDFDisallowCopy = value
50870   Case "PDFDISALLOWMODIFYANNOTATIONS": Options1.PDFDisallowModifyAnnotations = value
50880   Case "PDFDISALLOWMODIFYCONTENTS": Options1.PDFDisallowModifyContents = value
50890   Case "PDFDISALLOWPRINTING": Options1.PDFDisallowPrinting = value
50900   Case "PDFENCRYPTOR": Options1.PDFEncryptor = value
50910   Case "PDFFONTSEMBEDALL": Options1.PDFFontsEmbedAll = value
50920   Case "PDFFONTSSUBSETFONTS": Options1.PDFFontsSubSetFonts = value
50930   Case "PDFFONTSSUBSETFONTSPERCENT": Options1.PDFFontsSubSetFontsPercent = value
50940   Case "PDFGENERALASCII85": Options1.PDFGeneralASCII85 = value
50950   Case "PDFGENERALAUTOROTATE": Options1.PDFGeneralAutorotate = value
50960   Case "PDFGENERALCOMPATIBILITY": Options1.PDFGeneralCompatibility = value
50970   Case "PDFGENERALDEFAULT": Options1.PDFGeneralDefault = value
50980   Case "PDFGENERALOVERPRINT": Options1.PDFGeneralOverprint = value
50990   Case "PDFGENERALRESOLUTION": Options1.PDFGeneralResolution = value
51000   Case "PDFHIGHENCRYPTION": Options1.PDFHighEncryption = value
51010   Case "PDFLOWENCRYPTION": Options1.PDFLowEncryption = value
51020   Case "PDFOPTIMIZE": Options1.PDFOptimize = value
51030   Case "PDFOWNERPASS": Options1.PDFOwnerPass = value
51040   Case "PDFOWNERPASSWORDSTRING": Options1.PDFOwnerPasswordString = value
51050   Case "PDFSIGNINGMULTISIGNATURE": Options1.PDFSigningMultiSignature = value
51060   Case "PDFSIGNINGPFXFILE": Options1.PDFSigningPFXFile = value
51070   Case "PDFSIGNINGPFXFILEPASSWORD": Options1.PDFSigningPFXFilePassword = value
51080   Case "PDFSIGNINGSIGNATURECONTACT": Options1.PDFSigningSignatureContact = value
51090   Case "PDFSIGNINGSIGNATURELEFTX": Options1.PDFSigningSignatureLeftX = value
51100   Case "PDFSIGNINGSIGNATURELEFTY": Options1.PDFSigningSignatureLeftY = value
51110   Case "PDFSIGNINGSIGNATURELOCATION": Options1.PDFSigningSignatureLocation = value
51120   Case "PDFSIGNINGSIGNATUREONPAGE": Options1.PDFSigningSignatureOnPage = value
51130   Case "PDFSIGNINGSIGNATUREREASON": Options1.PDFSigningSignatureReason = value
51140   Case "PDFSIGNINGSIGNATURERIGHTX": Options1.PDFSigningSignatureRightX = value
51150   Case "PDFSIGNINGSIGNATURERIGHTY": Options1.PDFSigningSignatureRightY = value
51160   Case "PDFSIGNINGSIGNATUREVISIBLE": Options1.PDFSigningSignatureVisible = value
51170   Case "PDFSIGNINGSIGNPDF": Options1.PDFSigningSignPDF = value
51180   Case "PDFUPDATEMETADATA": Options1.PDFUpdateMetadata = value
51190   Case "PDFUSERPASS": Options1.PDFUserPass = value
51200   Case "PDFUSERPASSWORDSTRING": Options1.PDFUserPasswordString = value
51210   Case "PDFUSESECURITY": Options1.PDFUseSecurity = value
51220   Case "PNGCOLORSCOUNT": Options1.PNGColorscount = value
51230   Case "PNGRESOLUTION": Options1.PNGResolution = value
51240   Case "PRINTAFTERSAVING": Options1.PrintAfterSaving = value
51250   Case "PRINTAFTERSAVINGBITSPERPIXEL": Options1.PrintAfterSavingBitsPerPixel = value
51260   Case "PRINTAFTERSAVINGDUPLEX": Options1.PrintAfterSavingDuplex = value
51270   Case "PRINTAFTERSAVINGMAXRESOLUTION": Options1.PrintAfterSavingMaxResolution = value
51280   Case "PRINTAFTERSAVINGMAXRESOLUTIONENABLED": Options1.PrintAfterSavingMaxResolutionEnabled = value
51290   Case "PRINTAFTERSAVINGNOCANCEL": Options1.PrintAfterSavingNoCancel = value
51300   Case "PRINTAFTERSAVINGPRINTER": Options1.PrintAfterSavingPrinter = value
51310   Case "PRINTAFTERSAVINGQUERYUSER": Options1.PrintAfterSavingQueryUser = value
51320   Case "PRINTAFTERSAVINGTUMBLE": Options1.PrintAfterSavingTumble = value
51330   Case "PRINTERSTOP": Options1.PrinterStop = value
51340   Case "PROCESSPRIORITY": Options1.ProcessPriority = value
51350   Case "PROGRAMFONT": Options1.ProgramFont = value
51360   Case "PROGRAMFONTCHARSET": Options1.ProgramFontCharset = value
51370   Case "PROGRAMFONTSIZE": Options1.ProgramFontSize = value
51380   Case "PSDCOLORSCOUNT": Options1.PSDColorsCount = value
51390   Case "PSDRESOLUTION": Options1.PSDResolution = value
51400   Case "PSLANGUAGELEVEL": Options1.PSLanguageLevel = value
51410   Case "RAWCOLORSCOUNT": Options1.RAWColorsCount = value
51420   Case "RAWRESOLUTION": Options1.RAWResolution = value
51430   Case "REMOVEALLKNOWNFILEEXTENSIONS": Options1.RemoveAllKnownFileExtensions = value
51440   Case "REMOVESPACES": Options1.RemoveSpaces = value
51450   Case "RUNPROGRAMAFTERSAVING": Options1.RunProgramAfterSaving = value
51460   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": Options1.RunProgramAfterSavingProgramname = value
51470   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": Options1.RunProgramAfterSavingProgramParameters = value
51480   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": Options1.RunProgramAfterSavingWaitUntilReady = value
51490   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": Options1.RunProgramAfterSavingWindowstyle = value
51500   Case "RUNPROGRAMBEFORESAVING": Options1.RunProgramBeforeSaving = value
51510   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": Options1.RunProgramBeforeSavingProgramname = value
51520   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": Options1.RunProgramBeforeSavingProgramParameters = value
51530   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": Options1.RunProgramBeforeSavingWindowstyle = value
51540   Case "SAVEFILENAME": Options1.SaveFilename = value
51550   Case "SENDEMAILAFTERAUTOSAVING": Options1.SendEmailAfterAutoSaving = value
51560   Case "SENDMAILMETHOD": Options1.SendMailMethod = value
51570   Case "SHOWANIMATION": Options1.ShowAnimation = value
51580   Case "STAMPFONTCOLOR": Options1.StampFontColor = value
51590   Case "STAMPFONTNAME": Options1.StampFontname = value
51600   Case "STAMPFONTSIZE": Options1.StampFontsize = value
51610   Case "STAMPOUTLINEFONTTHICKNESS": Options1.StampOutlineFontthickness = value
51620   Case "STAMPSTRING": Options1.StampString = value
51630   Case "STAMPUSEOUTLINEFONT": Options1.StampUseOutlineFont = value
51640   Case "STANDARDAUTHOR": Options1.StandardAuthor = value
51650   Case "STANDARDCREATIONDATE": Options1.StandardCreationdate = value
51660   Case "STANDARDDATEFORMAT": Options1.StandardDateformat = value
51670   Case "STANDARDKEYWORDS": Options1.StandardKeywords = value
51680   Case "STANDARDMAILDOMAIN": Options1.StandardMailDomain = value
51690   Case "STANDARDMODIFYDATE": Options1.StandardModifydate = value
51700   Case "STANDARDSAVEFORMAT": Options1.StandardSaveformat = value
51710   Case "STANDARDSUBJECT": Options1.StandardSubject = value
51720   Case "STANDARDTITLE": Options1.StandardTitle = value
51730   Case "SVGRESOLUTION": Options1.SVGResolution = value
51740   Case "TIFFCOLORSCOUNT": Options1.TIFFColorscount = value
51750   Case "TIFFRESOLUTION": Options1.TIFFResolution = value
51760   Case "TOOLBARS": Options1.Toolbars = value
51770   Case "UPDATEINTERVAL": Options1.UpdateInterval = value
51780   Case "USEAUTOSAVE": Options1.UseAutosave = value
51790   Case "USEAUTOSAVEDIRECTORY": Options1.UseAutosaveDirectory = value
51800   Case "USECREATIONDATENOW": Options1.UseCreationDateNow = value
51810   Case "USECUSTOMPAPERSIZE": Options1.UseCustomPaperSize = value
51820   Case "USEFIXPAPERSIZE": Options1.UseFixPapersize = value
51830   Case "USESTANDARDAUTHOR": Options1.UseStandardAuthor = value
51840   Case Else:
51850    mError.Number = 3
51860    mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in cLetOption")
51870    RaiseEvent eError
51880  End Select
51890  If ProfileName = "" Then
51900    Options = Options1
51910   Else
51920    SaveOptions Options1, ProfileName
51930  End If
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
50220   .EditWithPDFArchitect = sOptions.EditWithPDFArchitect
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
50340   .MaximumCountOfPDFArchitectToolTip = sOptions.MaximumCountOfPDFArchitectToolTip
50350   .NoConfirmMessageSwitchingDefaultprinter = sOptions.NoConfirmMessageSwitchingDefaultprinter
50360   .NoProcessingAtStartup = sOptions.NoProcessingAtStartup
50370   .NoPSCheck = sOptions.NoPSCheck
50380   .OnePagePerFile = sOptions.OnePagePerFile
50390   .OpenOutputFile = sOptions.OpenOutputFile
50400   .OptionsDesign = sOptions.OptionsDesign
50410   .OptionsEnabled = sOptions.OptionsEnabled
50420   .OptionsVisible = sOptions.OptionsVisible
50430   .Papersize = sOptions.Papersize
50440   .PCLColorsCount = sOptions.PCLColorsCount
50450   .PCLResolution = sOptions.PCLResolution
50460   .PCXColorscount = sOptions.PCXColorscount
50470   .PCXResolution = sOptions.PCXResolution
50480   .PDFAes128Encryption = sOptions.PDFAes128Encryption
50490   .PDFAllowAssembly = sOptions.PDFAllowAssembly
50500   .PDFAllowDegradedPrinting = sOptions.PDFAllowDegradedPrinting
50510   .PDFAllowFillIn = sOptions.PDFAllowFillIn
50520   .PDFAllowScreenReaders = sOptions.PDFAllowScreenReaders
50530   .PDFColorsCMYKToRGB = sOptions.PDFColorsCMYKToRGB
50540   .PDFColorsColorModel = sOptions.PDFColorsColorModel
50550   .PDFColorsPreserveHalftone = sOptions.PDFColorsPreserveHalftone
50560   .PDFColorsPreserveOverprint = sOptions.PDFColorsPreserveOverprint
50570   .PDFColorsPreserveTransfer = sOptions.PDFColorsPreserveTransfer
50580   .PDFCompressionColorCompression = sOptions.PDFCompressionColorCompression
50590   .PDFCompressionColorCompressionChoice = sOptions.PDFCompressionColorCompressionChoice
50600   .PDFCompressionColorCompressionJPEGHighFactor = sOptions.PDFCompressionColorCompressionJPEGHighFactor
50610   .PDFCompressionColorCompressionJPEGLowFactor = sOptions.PDFCompressionColorCompressionJPEGLowFactor
50620   .PDFCompressionColorCompressionJPEGManualFactor = sOptions.PDFCompressionColorCompressionJPEGManualFactor
50630   .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
50640   .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
50650   .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
50660   .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
50670   .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
50680   .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
50690   .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
50700   .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
50710   .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
50720   .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
50730   .PDFCompressionGreyCompressionJPEGManualFactor = sOptions.PDFCompressionGreyCompressionJPEGManualFactor
50740   .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
50750   .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
50760   .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
50770   .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
50780   .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
50790   .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
50800   .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
50810   .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
50820   .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
50830   .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
50840   .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
50850   .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
50860   .PDFDisallowCopy = sOptions.PDFDisallowCopy
50870   .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
50880   .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
50890   .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
50900   .PDFEncryptor = sOptions.PDFEncryptor
50910   .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
50920   .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
50930   .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
50940   .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
50950   .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
50960   .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
50970   .PDFGeneralDefault = sOptions.PDFGeneralDefault
50980   .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
50990   .PDFGeneralResolution = sOptions.PDFGeneralResolution
51000   .PDFHighEncryption = sOptions.PDFHighEncryption
51010   .PDFLowEncryption = sOptions.PDFLowEncryption
51020   .PDFOptimize = sOptions.PDFOptimize
51030   .PDFOwnerPass = sOptions.PDFOwnerPass
51040   .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
51050   .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
51060   .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
51070   .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
51080   .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
51090   .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
51100   .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
51110   .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
51120   .PDFSigningSignatureOnPage = sOptions.PDFSigningSignatureOnPage
51130   .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
51140   .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
51150   .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
51160   .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
51170   .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
51180   .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
51190   .PDFUserPass = sOptions.PDFUserPass
51200   .PDFUserPasswordString = sOptions.PDFUserPasswordString
51210   .PDFUseSecurity = sOptions.PDFUseSecurity
51220   .PNGColorscount = sOptions.PNGColorscount
51230   .PNGResolution = sOptions.PNGResolution
51240   .PrintAfterSaving = sOptions.PrintAfterSaving
51250   .PrintAfterSavingBitsPerPixel = sOptions.PrintAfterSavingBitsPerPixel
51260   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
51270   .PrintAfterSavingMaxResolution = sOptions.PrintAfterSavingMaxResolution
51280   .PrintAfterSavingMaxResolutionEnabled = sOptions.PrintAfterSavingMaxResolutionEnabled
51290   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
51300   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
51310   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51320   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51330   .PrinterStop = sOptions.PrinterStop
51340   .ProcessPriority = sOptions.ProcessPriority
51350   .ProgramFont = sOptions.ProgramFont
51360   .ProgramFontCharset = sOptions.ProgramFontCharset
51370   .ProgramFontSize = sOptions.ProgramFontSize
51380   .PSDColorsCount = sOptions.PSDColorsCount
51390   .PSDResolution = sOptions.PSDResolution
51400   .PSLanguageLevel = sOptions.PSLanguageLevel
51410   .RAWColorsCount = sOptions.RAWColorsCount
51420   .RAWResolution = sOptions.RAWResolution
51430   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51440   .RemoveSpaces = sOptions.RemoveSpaces
51450   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51460   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51470   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51480   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51490   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51500   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51510   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51520   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51530   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51540   .SaveFilename = sOptions.SaveFilename
51550   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51560   .SendMailMethod = sOptions.SendMailMethod
51570   .ShowAnimation = sOptions.ShowAnimation
51580   .StampFontColor = sOptions.StampFontColor
51590   .StampFontname = sOptions.StampFontname
51600   .StampFontsize = sOptions.StampFontsize
51610   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51620   .StampString = sOptions.StampString
51630   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51640   .StandardAuthor = sOptions.StandardAuthor
51650   .StandardCreationdate = sOptions.StandardCreationdate
51660   .StandardDateformat = sOptions.StandardDateformat
51670   .StandardKeywords = sOptions.StandardKeywords
51680   .StandardMailDomain = sOptions.StandardMailDomain
51690   .StandardModifydate = sOptions.StandardModifydate
51700   .StandardSaveformat = sOptions.StandardSaveformat
51710   .StandardSubject = sOptions.StandardSubject
51720   .StandardTitle = sOptions.StandardTitle
51730   .SVGResolution = sOptions.SVGResolution
51740   .TIFFColorscount = sOptions.TIFFColorscount
51750   .TIFFResolution = sOptions.TIFFResolution
51760   .Toolbars = sOptions.Toolbars
51770   .UpdateInterval = sOptions.UpdateInterval
51780   .UseAutosave = sOptions.UseAutosave
51790   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51800   .UseCreationDateNow = sOptions.UseCreationDateNow
51810   .UseCustomPaperSize = sOptions.UseCustomPaperSize
51820   .UseFixPapersize = sOptions.UseFixPapersize
51830   .UseStandardAuthor = sOptions.UseStandardAuthor
51840  End With
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

