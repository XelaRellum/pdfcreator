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
50220   .Add "DisableUpdateCheck"
50230   .Add "DontUseDocumentSettings"
50240   .Add "EditWithPDFArchitect"
50250   .Add "EPSLanguageLevel"
50260   .Add "FilenameSubstitutions"
50270   .Add "FilenameSubstitutionsOnlyInTitle"
50280   .Add "JPEGColorscount"
50290   .Add "JPEGQuality"
50300   .Add "JPEGResolution"
50310   .Add "Language"
50320   .Add "LastSaveDirectory"
50330   .Add "LastUpdateCheck"
50340   .Add "Logging"
50350   .Add "LogLines"
50360   .Add "MaximumCountOfPDFArchitectToolTip"
50370   .Add "NoConfirmMessageSwitchingDefaultprinter"
50380   .Add "NoProcessingAtStartup"
50390   .Add "NoPSCheck"
50400   .Add "OneFilePerPage"
50410   .Add "OpenOutputFile"
50420   .Add "OptionsDesign"
50430   .Add "OptionsEnabled"
50440   .Add "OptionsVisible"
50450   .Add "Papersize"
50460   .Add "PCLColorsCount"
50470   .Add "PCLResolution"
50480   .Add "PCXColorscount"
50490   .Add "PCXResolution"
50500   .Add "PDFAes128Encryption"
50510   .Add "PDFAllowAssembly"
50520   .Add "PDFAllowDegradedPrinting"
50530   .Add "PDFAllowFillIn"
50540   .Add "PDFAllowScreenReaders"
50550   .Add "PDFColorsCMYKToRGB"
50560   .Add "PDFColorsColorModel"
50570   .Add "PDFColorsPreserveHalftone"
50580   .Add "PDFColorsPreserveOverprint"
50590   .Add "PDFColorsPreserveTransfer"
50600   .Add "PDFCompressionColorCompression"
50610   .Add "PDFCompressionColorCompressionChoice"
50620   .Add "PDFCompressionColorCompressionJPEGHighFactor"
50630   .Add "PDFCompressionColorCompressionJPEGLowFactor"
50640   .Add "PDFCompressionColorCompressionJPEGManualFactor"
50650   .Add "PDFCompressionColorCompressionJPEGMaximumFactor"
50660   .Add "PDFCompressionColorCompressionJPEGMediumFactor"
50670   .Add "PDFCompressionColorCompressionJPEGMinimumFactor"
50680   .Add "PDFCompressionColorResample"
50690   .Add "PDFCompressionColorResampleChoice"
50700   .Add "PDFCompressionColorResolution"
50710   .Add "PDFCompressionGreyCompression"
50720   .Add "PDFCompressionGreyCompressionChoice"
50730   .Add "PDFCompressionGreyCompressionJPEGHighFactor"
50740   .Add "PDFCompressionGreyCompressionJPEGLowFactor"
50750   .Add "PDFCompressionGreyCompressionJPEGManualFactor"
50760   .Add "PDFCompressionGreyCompressionJPEGMaximumFactor"
50770   .Add "PDFCompressionGreyCompressionJPEGMediumFactor"
50780   .Add "PDFCompressionGreyCompressionJPEGMinimumFactor"
50790   .Add "PDFCompressionGreyResample"
50800   .Add "PDFCompressionGreyResampleChoice"
50810   .Add "PDFCompressionGreyResolution"
50820   .Add "PDFCompressionMonoCompression"
50830   .Add "PDFCompressionMonoCompressionChoice"
50840   .Add "PDFCompressionMonoResample"
50850   .Add "PDFCompressionMonoResampleChoice"
50860   .Add "PDFCompressionMonoResolution"
50870   .Add "PDFCompressionTextCompression"
50880   .Add "PDFDisallowCopy"
50890   .Add "PDFDisallowModifyAnnotations"
50900   .Add "PDFDisallowModifyContents"
50910   .Add "PDFDisallowPrinting"
50920   .Add "PDFEncryptor"
50930   .Add "PDFFontsEmbedAll"
50940   .Add "PDFFontsSubSetFonts"
50950   .Add "PDFFontsSubSetFontsPercent"
50960   .Add "PDFGeneralASCII85"
50970   .Add "PDFGeneralAutorotate"
50980   .Add "PDFGeneralCompatibility"
50990   .Add "PDFGeneralDefault"
51000   .Add "PDFGeneralOverprint"
51010   .Add "PDFGeneralResolution"
51020   .Add "PDFHighEncryption"
51030   .Add "PDFLowEncryption"
51040   .Add "PDFOptimize"
51050   .Add "PDFOwnerPass"
51060   .Add "PDFOwnerPasswordString"
51070   .Add "PDFPageLayout"
51080   .Add "PDFPageMode"
51090   .Add "PDFSigningMultiSignature"
51100   .Add "PDFSigningPFXFile"
51110   .Add "PDFSigningPFXFilePassword"
51120   .Add "PDFSigningSignatureContact"
51130   .Add "PDFSigningSignatureLeftX"
51140   .Add "PDFSigningSignatureLeftY"
51150   .Add "PDFSigningSignatureLocation"
51160   .Add "PDFSigningSignatureOnPage"
51170   .Add "PDFSigningSignatureReason"
51180   .Add "PDFSigningSignatureRightX"
51190   .Add "PDFSigningSignatureRightY"
51200   .Add "PDFSigningSignatureVisible"
51210   .Add "PDFSigningSignPDF"
51220   .Add "PDFSigningTimeServerUrl"
51230   .Add "PDFStartPage"
51240   .Add "PDFUpdateMetadata"
51250   .Add "PDFUserPass"
51260   .Add "PDFUserPasswordString"
51270   .Add "PDFUseSecurity"
51280   .Add "PNGColorscount"
51290   .Add "PNGResolution"
51300   .Add "PrintAfterSaving"
51310   .Add "PrintAfterSavingBitsPerPixel"
51320   .Add "PrintAfterSavingDuplex"
51330   .Add "PrintAfterSavingMaxResolution"
51340   .Add "PrintAfterSavingMaxResolutionEnabled"
51350   .Add "PrintAfterSavingNoCancel"
51360   .Add "PrintAfterSavingPrinter"
51370   .Add "PrintAfterSavingQueryUser"
51380   .Add "PrintAfterSavingTumble"
51390   .Add "PrinterStop"
51400   .Add "ProcessPriority"
51410   .Add "ProgramFont"
51420   .Add "ProgramFontCharset"
51430   .Add "ProgramFontSize"
51440   .Add "PSDColorsCount"
51450   .Add "PSDResolution"
51460   .Add "PSLanguageLevel"
51470   .Add "RAWColorsCount"
51480   .Add "RAWResolution"
51490   .Add "RemoveAllKnownFileExtensions"
51500   .Add "RemoveSpaces"
51510   .Add "RunProgramAfterSaving"
51520   .Add "RunProgramAfterSavingProgramname"
51530   .Add "RunProgramAfterSavingProgramParameters"
51540   .Add "RunProgramAfterSavingWaitUntilReady"
51550   .Add "RunProgramAfterSavingWindowstyle"
51560   .Add "RunProgramBeforeSaving"
51570   .Add "RunProgramBeforeSavingProgramname"
51580   .Add "RunProgramBeforeSavingProgramParameters"
51590   .Add "RunProgramBeforeSavingWindowstyle"
51600   .Add "SaveFilename"
51610   .Add "SendEmailAfterAutoSaving"
51620   .Add "SendMailMethod"
51630   .Add "ShowAnimation"
51640   .Add "StampFontColor"
51650   .Add "StampFontname"
51660   .Add "StampFontsize"
51670   .Add "StampOutlineFontthickness"
51680   .Add "StampString"
51690   .Add "StampUseOutlineFont"
51700   .Add "StandardAuthor"
51710   .Add "StandardCreationdate"
51720   .Add "StandardDateformat"
51730   .Add "StandardKeywords"
51740   .Add "StandardMailDomain"
51750   .Add "StandardModifydate"
51760   .Add "StandardSaveformat"
51770   .Add "StandardSubject"
51780   .Add "StandardTitle"
51790   .Add "SVGResolution"
51800   .Add "TIFFColorscount"
51810   .Add "TIFFResolution"
51820   .Add "Toolbars"
51830   .Add "UpdateInterval"
51840   .Add "UseAutosave"
51850   .Add "UseAutosaveDirectory"
51860   .Add "UseCreationDateNow"
51870   .Add "UseCustomPaperSize"
51880   .Add "UseFixPapersize"
51890   .Add "UseStandardAuthor"
51900  End With
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
50220  cAddPDFCreatorPrinter = InstallPrinter(PrinterName, "PDFCreator", "pdfcmon", "")
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
50210   Case "DISABLEUPDATECHECK": cGetOptionFromOptions = Options.DisableUpdateCheck
50220   Case "DONTUSEDOCUMENTSETTINGS": cGetOptionFromOptions = Options.DontUseDocumentSettings
50230   Case "EDITWITHPDFARCHITECT": cGetOptionFromOptions = Options.EditWithPDFArchitect
50240   Case "EPSLANGUAGELEVEL": cGetOptionFromOptions = Options.EPSLanguageLevel
50250   Case "FILENAMESUBSTITUTIONS": cGetOptionFromOptions = Options.FilenameSubstitutions
50260   Case "FILENAMESUBSTITUTIONSONLYINTITLE": cGetOptionFromOptions = Options.FilenameSubstitutionsOnlyInTitle
50270   Case "JPEGCOLORSCOUNT": cGetOptionFromOptions = Options.JPEGColorscount
50280   Case "JPEGQUALITY": cGetOptionFromOptions = Options.JPEGQuality
50290   Case "JPEGRESOLUTION": cGetOptionFromOptions = Options.JPEGResolution
50300   Case "LANGUAGE": cGetOptionFromOptions = Options.Language
50310   Case "LASTSAVEDIRECTORY": cGetOptionFromOptions = Options.LastSaveDirectory
50320   Case "LASTUPDATECHECK": cGetOptionFromOptions = Options.LastUpdateCheck
50330   Case "LOGGING": cGetOptionFromOptions = Options.Logging
50340   Case "LOGLINES": cGetOptionFromOptions = Options.LogLines
50350   Case "MAXIMUMCOUNTOFPDFARCHITECTTOOLTIP": cGetOptionFromOptions = Options.MaximumCountOfPDFArchitectToolTip
50360   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": cGetOptionFromOptions = Options.NoConfirmMessageSwitchingDefaultprinter
50370   Case "NOPROCESSINGATSTARTUP": cGetOptionFromOptions = Options.NoProcessingAtStartup
50380   Case "NOPSCHECK": cGetOptionFromOptions = Options.NoPSCheck
50390   Case "ONEFILEPERPAGE": cGetOptionFromOptions = Options.OneFilePerPage
50400   Case "OPENOUTPUTFILE": cGetOptionFromOptions = Options.OpenOutputFile
50410   Case "OPTIONSDESIGN": cGetOptionFromOptions = Options.OptionsDesign
50420   Case "OPTIONSENABLED": cGetOptionFromOptions = Options.OptionsEnabled
50430   Case "OPTIONSVISIBLE": cGetOptionFromOptions = Options.OptionsVisible
50440   Case "PAPERSIZE": cGetOptionFromOptions = Options.Papersize
50450   Case "PCLCOLORSCOUNT": cGetOptionFromOptions = Options.PCLColorsCount
50460   Case "PCLRESOLUTION": cGetOptionFromOptions = Options.PCLResolution
50470   Case "PCXCOLORSCOUNT": cGetOptionFromOptions = Options.PCXColorscount
50480   Case "PCXRESOLUTION": cGetOptionFromOptions = Options.PCXResolution
50490   Case "PDFAES128ENCRYPTION": cGetOptionFromOptions = Options.PDFAes128Encryption
50500   Case "PDFALLOWASSEMBLY": cGetOptionFromOptions = Options.PDFAllowAssembly
50510   Case "PDFALLOWDEGRADEDPRINTING": cGetOptionFromOptions = Options.PDFAllowDegradedPrinting
50520   Case "PDFALLOWFILLIN": cGetOptionFromOptions = Options.PDFAllowFillIn
50530   Case "PDFALLOWSCREENREADERS": cGetOptionFromOptions = Options.PDFAllowScreenReaders
50540   Case "PDFCOLORSCMYKTORGB": cGetOptionFromOptions = Options.PDFColorsCMYKToRGB
50550   Case "PDFCOLORSCOLORMODEL": cGetOptionFromOptions = Options.PDFColorsColorModel
50560   Case "PDFCOLORSPRESERVEHALFTONE": cGetOptionFromOptions = Options.PDFColorsPreserveHalftone
50570   Case "PDFCOLORSPRESERVEOVERPRINT": cGetOptionFromOptions = Options.PDFColorsPreserveOverprint
50580   Case "PDFCOLORSPRESERVETRANSFER": cGetOptionFromOptions = Options.PDFColorsPreserveTransfer
50590   Case "PDFCOMPRESSIONCOLORCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionColorCompression
50600   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionColorCompressionChoice
50610   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGHighFactor
50620   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGLowFactor
50630   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMANUALFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGManualFactor
50640   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMaximumFactor
50650   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMediumFactor
50660   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionColorCompressionJPEGMinimumFactor
50670   Case "PDFCOMPRESSIONCOLORRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionColorResample
50680   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionColorResampleChoice
50690   Case "PDFCOMPRESSIONCOLORRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionColorResolution
50700   Case "PDFCOMPRESSIONGREYCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionGreyCompression
50710   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionChoice
50720   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGHighFactor
50730   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGLowFactor
50740   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMANUALFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGManualFactor
50750   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMaximumFactor
50760   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMediumFactor
50770   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": cGetOptionFromOptions = Options.PDFCompressionGreyCompressionJPEGMinimumFactor
50780   Case "PDFCOMPRESSIONGREYRESAMPLE": cGetOptionFromOptions = Options.PDFCompressionGreyResample
50790   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionGreyResampleChoice
50800   Case "PDFCOMPRESSIONGREYRESOLUTION": cGetOptionFromOptions = Options.PDFCompressionGreyResolution
50810   Case "PDFCOMPRESSIONMONOCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionMonoCompression
50820   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoCompressionChoice
50830   Case "PDFCOMPRESSIONMONORESAMPLE": cGetOptionFromOptions = Options.PDFCompressionMonoResample
50840   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": cGetOptionFromOptions = Options.PDFCompressionMonoResampleChoice
50850   Case "PDFCOMPRESSIONMONORESOLUTION": cGetOptionFromOptions = Options.PDFCompressionMonoResolution
50860   Case "PDFCOMPRESSIONTEXTCOMPRESSION": cGetOptionFromOptions = Options.PDFCompressionTextCompression
50870   Case "PDFDISALLOWCOPY": cGetOptionFromOptions = Options.PDFDisallowCopy
50880   Case "PDFDISALLOWMODIFYANNOTATIONS": cGetOptionFromOptions = Options.PDFDisallowModifyAnnotations
50890   Case "PDFDISALLOWMODIFYCONTENTS": cGetOptionFromOptions = Options.PDFDisallowModifyContents
50900   Case "PDFDISALLOWPRINTING": cGetOptionFromOptions = Options.PDFDisallowPrinting
50910   Case "PDFENCRYPTOR": cGetOptionFromOptions = Options.PDFEncryptor
50920   Case "PDFFONTSEMBEDALL": cGetOptionFromOptions = Options.PDFFontsEmbedAll
50930   Case "PDFFONTSSUBSETFONTS": cGetOptionFromOptions = Options.PDFFontsSubSetFonts
50940   Case "PDFFONTSSUBSETFONTSPERCENT": cGetOptionFromOptions = Options.PDFFontsSubSetFontsPercent
50950   Case "PDFGENERALASCII85": cGetOptionFromOptions = Options.PDFGeneralASCII85
50960   Case "PDFGENERALAUTOROTATE": cGetOptionFromOptions = Options.PDFGeneralAutorotate
50970   Case "PDFGENERALCOMPATIBILITY": cGetOptionFromOptions = Options.PDFGeneralCompatibility
50980   Case "PDFGENERALDEFAULT": cGetOptionFromOptions = Options.PDFGeneralDefault
50990   Case "PDFGENERALOVERPRINT": cGetOptionFromOptions = Options.PDFGeneralOverprint
51000   Case "PDFGENERALRESOLUTION": cGetOptionFromOptions = Options.PDFGeneralResolution
51010   Case "PDFHIGHENCRYPTION": cGetOptionFromOptions = Options.PDFHighEncryption
51020   Case "PDFLOWENCRYPTION": cGetOptionFromOptions = Options.PDFLowEncryption
51030   Case "PDFOPTIMIZE": cGetOptionFromOptions = Options.PDFOptimize
51040   Case "PDFOWNERPASS": cGetOptionFromOptions = Options.PDFOwnerPass
51050   Case "PDFOWNERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFOwnerPasswordString
51060   Case "PDFPAGELAYOUT": cGetOptionFromOptions = Options.PDFPageLayout
51070   Case "PDFPAGEMODE": cGetOptionFromOptions = Options.PDFPageMode
51080   Case "PDFSIGNINGMULTISIGNATURE": cGetOptionFromOptions = Options.PDFSigningMultiSignature
51090   Case "PDFSIGNINGPFXFILE": cGetOptionFromOptions = Options.PDFSigningPFXFile
51100   Case "PDFSIGNINGPFXFILEPASSWORD": cGetOptionFromOptions = Options.PDFSigningPFXFilePassword
51110   Case "PDFSIGNINGSIGNATURECONTACT": cGetOptionFromOptions = Options.PDFSigningSignatureContact
51120   Case "PDFSIGNINGSIGNATURELEFTX": cGetOptionFromOptions = Options.PDFSigningSignatureLeftX
51130   Case "PDFSIGNINGSIGNATURELEFTY": cGetOptionFromOptions = Options.PDFSigningSignatureLeftY
51140   Case "PDFSIGNINGSIGNATURELOCATION": cGetOptionFromOptions = Options.PDFSigningSignatureLocation
51150   Case "PDFSIGNINGSIGNATUREONPAGE": cGetOptionFromOptions = Options.PDFSigningSignatureOnPage
51160   Case "PDFSIGNINGSIGNATUREREASON": cGetOptionFromOptions = Options.PDFSigningSignatureReason
51170   Case "PDFSIGNINGSIGNATURERIGHTX": cGetOptionFromOptions = Options.PDFSigningSignatureRightX
51180   Case "PDFSIGNINGSIGNATURERIGHTY": cGetOptionFromOptions = Options.PDFSigningSignatureRightY
51190   Case "PDFSIGNINGSIGNATUREVISIBLE": cGetOptionFromOptions = Options.PDFSigningSignatureVisible
51200   Case "PDFSIGNINGSIGNPDF": cGetOptionFromOptions = Options.PDFSigningSignPDF
51210   Case "PDFSIGNINGTIMESERVERURL": cGetOptionFromOptions = Options.PDFSigningTimeServerUrl
51220   Case "PDFSTARTPAGE": cGetOptionFromOptions = Options.PDFStartPage
51230   Case "PDFUPDATEMETADATA": cGetOptionFromOptions = Options.PDFUpdateMetadata
51240   Case "PDFUSERPASS": cGetOptionFromOptions = Options.PDFUserPass
51250   Case "PDFUSERPASSWORDSTRING": cGetOptionFromOptions = Options.PDFUserPasswordString
51260   Case "PDFUSESECURITY": cGetOptionFromOptions = Options.PDFUseSecurity
51270   Case "PNGCOLORSCOUNT": cGetOptionFromOptions = Options.PNGColorscount
51280   Case "PNGRESOLUTION": cGetOptionFromOptions = Options.PNGResolution
51290   Case "PRINTAFTERSAVING": cGetOptionFromOptions = Options.PrintAfterSaving
51300   Case "PRINTAFTERSAVINGBITSPERPIXEL": cGetOptionFromOptions = Options.PrintAfterSavingBitsPerPixel
51310   Case "PRINTAFTERSAVINGDUPLEX": cGetOptionFromOptions = Options.PrintAfterSavingDuplex
51320   Case "PRINTAFTERSAVINGMAXRESOLUTION": cGetOptionFromOptions = Options.PrintAfterSavingMaxResolution
51330   Case "PRINTAFTERSAVINGMAXRESOLUTIONENABLED": cGetOptionFromOptions = Options.PrintAfterSavingMaxResolutionEnabled
51340   Case "PRINTAFTERSAVINGNOCANCEL": cGetOptionFromOptions = Options.PrintAfterSavingNoCancel
51350   Case "PRINTAFTERSAVINGPRINTER": cGetOptionFromOptions = Options.PrintAfterSavingPrinter
51360   Case "PRINTAFTERSAVINGQUERYUSER": cGetOptionFromOptions = Options.PrintAfterSavingQueryUser
51370   Case "PRINTAFTERSAVINGTUMBLE": cGetOptionFromOptions = Options.PrintAfterSavingTumble
51380   Case "PRINTERSTOP": cGetOptionFromOptions = Options.PrinterStop
51390   Case "PROCESSPRIORITY": cGetOptionFromOptions = Options.ProcessPriority
51400   Case "PROGRAMFONT": cGetOptionFromOptions = Options.ProgramFont
51410   Case "PROGRAMFONTCHARSET": cGetOptionFromOptions = Options.ProgramFontCharset
51420   Case "PROGRAMFONTSIZE": cGetOptionFromOptions = Options.ProgramFontSize
51430   Case "PSDCOLORSCOUNT": cGetOptionFromOptions = Options.PSDColorsCount
51440   Case "PSDRESOLUTION": cGetOptionFromOptions = Options.PSDResolution
51450   Case "PSLANGUAGELEVEL": cGetOptionFromOptions = Options.PSLanguageLevel
51460   Case "RAWCOLORSCOUNT": cGetOptionFromOptions = Options.RAWColorsCount
51470   Case "RAWRESOLUTION": cGetOptionFromOptions = Options.RAWResolution
51480   Case "REMOVEALLKNOWNFILEEXTENSIONS": cGetOptionFromOptions = Options.RemoveAllKnownFileExtensions
51490   Case "REMOVESPACES": cGetOptionFromOptions = Options.RemoveSpaces
51500   Case "RUNPROGRAMAFTERSAVING": cGetOptionFromOptions = Options.RunProgramAfterSaving
51510   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramname
51520   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramAfterSavingProgramParameters
51530   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": cGetOptionFromOptions = Options.RunProgramAfterSavingWaitUntilReady
51540   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramAfterSavingWindowstyle
51550   Case "RUNPROGRAMBEFORESAVING": cGetOptionFromOptions = Options.RunProgramBeforeSaving
51560   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramname
51570   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": cGetOptionFromOptions = Options.RunProgramBeforeSavingProgramParameters
51580   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": cGetOptionFromOptions = Options.RunProgramBeforeSavingWindowstyle
51590   Case "SAVEFILENAME": cGetOptionFromOptions = Options.SaveFilename
51600   Case "SENDEMAILAFTERAUTOSAVING": cGetOptionFromOptions = Options.SendEmailAfterAutoSaving
51610   Case "SENDMAILMETHOD": cGetOptionFromOptions = Options.SendMailMethod
51620   Case "SHOWANIMATION": cGetOptionFromOptions = Options.ShowAnimation
51630   Case "STAMPFONTCOLOR": cGetOptionFromOptions = Options.StampFontColor
51640   Case "STAMPFONTNAME": cGetOptionFromOptions = Options.StampFontname
51650   Case "STAMPFONTSIZE": cGetOptionFromOptions = Options.StampFontsize
51660   Case "STAMPOUTLINEFONTTHICKNESS": cGetOptionFromOptions = Options.StampOutlineFontthickness
51670   Case "STAMPSTRING": cGetOptionFromOptions = Options.StampString
51680   Case "STAMPUSEOUTLINEFONT": cGetOptionFromOptions = Options.StampUseOutlineFont
51690   Case "STANDARDAUTHOR": cGetOptionFromOptions = Options.StandardAuthor
51700   Case "STANDARDCREATIONDATE": cGetOptionFromOptions = Options.StandardCreationdate
51710   Case "STANDARDDATEFORMAT": cGetOptionFromOptions = Options.StandardDateformat
51720   Case "STANDARDKEYWORDS": cGetOptionFromOptions = Options.StandardKeywords
51730   Case "STANDARDMAILDOMAIN": cGetOptionFromOptions = Options.StandardMailDomain
51740   Case "STANDARDMODIFYDATE": cGetOptionFromOptions = Options.StandardModifydate
51750   Case "STANDARDSAVEFORMAT": cGetOptionFromOptions = Options.StandardSaveformat
51760   Case "STANDARDSUBJECT": cGetOptionFromOptions = Options.StandardSubject
51770   Case "STANDARDTITLE": cGetOptionFromOptions = Options.StandardTitle
51780   Case "SVGRESOLUTION": cGetOptionFromOptions = Options.SVGResolution
51790   Case "TIFFCOLORSCOUNT": cGetOptionFromOptions = Options.TIFFColorscount
51800   Case "TIFFRESOLUTION": cGetOptionFromOptions = Options.TIFFResolution
51810   Case "TOOLBARS": cGetOptionFromOptions = Options.Toolbars
51820   Case "UPDATEINTERVAL": cGetOptionFromOptions = Options.UpdateInterval
51830   Case "USEAUTOSAVE": cGetOptionFromOptions = Options.UseAutosave
51840   Case "USEAUTOSAVEDIRECTORY": cGetOptionFromOptions = Options.UseAutosaveDirectory
51850   Case "USECREATIONDATENOW": cGetOptionFromOptions = Options.UseCreationDateNow
51860   Case "USECUSTOMPAPERSIZE": cGetOptionFromOptions = Options.UseCustomPaperSize
51870   Case "USEFIXPAPERSIZE": cGetOptionFromOptions = Options.UseFixPapersize
51880   Case "USESTANDARDAUTHOR": cGetOptionFromOptions = Options.UseStandardAuthor
51890  End Select
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
50220   .DisableUpdateCheck = sOptions.DisableUpdateCheck
50230   .DontUseDocumentSettings = sOptions.DontUseDocumentSettings
50240   .EditWithPDFArchitect = sOptions.EditWithPDFArchitect
50250   .EPSLanguageLevel = sOptions.EPSLanguageLevel
50260   .FilenameSubstitutions = sOptions.FilenameSubstitutions
50270   .FilenameSubstitutionsOnlyInTitle = sOptions.FilenameSubstitutionsOnlyInTitle
50280   .JPEGColorscount = sOptions.JPEGColorscount
50290   .JPEGQuality = sOptions.JPEGQuality
50300   .JPEGResolution = sOptions.JPEGResolution
50310   .Language = sOptions.Language
50320   .LastSaveDirectory = sOptions.LastSaveDirectory
50330   .LastUpdateCheck = sOptions.LastUpdateCheck
50340   .Logging = sOptions.Logging
50350   .LogLines = sOptions.LogLines
50360   .MaximumCountOfPDFArchitectToolTip = sOptions.MaximumCountOfPDFArchitectToolTip
50370   .NoConfirmMessageSwitchingDefaultprinter = sOptions.NoConfirmMessageSwitchingDefaultprinter
50380   .NoProcessingAtStartup = sOptions.NoProcessingAtStartup
50390   .NoPSCheck = sOptions.NoPSCheck
50400   .OneFilePerPage = sOptions.OneFilePerPage
50410   .OpenOutputFile = sOptions.OpenOutputFile
50420   .OptionsDesign = sOptions.OptionsDesign
50430   .OptionsEnabled = sOptions.OptionsEnabled
50440   .OptionsVisible = sOptions.OptionsVisible
50450   .Papersize = sOptions.Papersize
50460   .PCLColorsCount = sOptions.PCLColorsCount
50470   .PCLResolution = sOptions.PCLResolution
50480   .PCXColorscount = sOptions.PCXColorscount
50490   .PCXResolution = sOptions.PCXResolution
50500   .PDFAes128Encryption = sOptions.PDFAes128Encryption
50510   .PDFAllowAssembly = sOptions.PDFAllowAssembly
50520   .PDFAllowDegradedPrinting = sOptions.PDFAllowDegradedPrinting
50530   .PDFAllowFillIn = sOptions.PDFAllowFillIn
50540   .PDFAllowScreenReaders = sOptions.PDFAllowScreenReaders
50550   .PDFColorsCMYKToRGB = sOptions.PDFColorsCMYKToRGB
50560   .PDFColorsColorModel = sOptions.PDFColorsColorModel
50570   .PDFColorsPreserveHalftone = sOptions.PDFColorsPreserveHalftone
50580   .PDFColorsPreserveOverprint = sOptions.PDFColorsPreserveOverprint
50590   .PDFColorsPreserveTransfer = sOptions.PDFColorsPreserveTransfer
50600   .PDFCompressionColorCompression = sOptions.PDFCompressionColorCompression
50610   .PDFCompressionColorCompressionChoice = sOptions.PDFCompressionColorCompressionChoice
50620   .PDFCompressionColorCompressionJPEGHighFactor = sOptions.PDFCompressionColorCompressionJPEGHighFactor
50630   .PDFCompressionColorCompressionJPEGLowFactor = sOptions.PDFCompressionColorCompressionJPEGLowFactor
50640   .PDFCompressionColorCompressionJPEGManualFactor = sOptions.PDFCompressionColorCompressionJPEGManualFactor
50650   .PDFCompressionColorCompressionJPEGMaximumFactor = sOptions.PDFCompressionColorCompressionJPEGMaximumFactor
50660   .PDFCompressionColorCompressionJPEGMediumFactor = sOptions.PDFCompressionColorCompressionJPEGMediumFactor
50670   .PDFCompressionColorCompressionJPEGMinimumFactor = sOptions.PDFCompressionColorCompressionJPEGMinimumFactor
50680   .PDFCompressionColorResample = sOptions.PDFCompressionColorResample
50690   .PDFCompressionColorResampleChoice = sOptions.PDFCompressionColorResampleChoice
50700   .PDFCompressionColorResolution = sOptions.PDFCompressionColorResolution
50710   .PDFCompressionGreyCompression = sOptions.PDFCompressionGreyCompression
50720   .PDFCompressionGreyCompressionChoice = sOptions.PDFCompressionGreyCompressionChoice
50730   .PDFCompressionGreyCompressionJPEGHighFactor = sOptions.PDFCompressionGreyCompressionJPEGHighFactor
50740   .PDFCompressionGreyCompressionJPEGLowFactor = sOptions.PDFCompressionGreyCompressionJPEGLowFactor
50750   .PDFCompressionGreyCompressionJPEGManualFactor = sOptions.PDFCompressionGreyCompressionJPEGManualFactor
50760   .PDFCompressionGreyCompressionJPEGMaximumFactor = sOptions.PDFCompressionGreyCompressionJPEGMaximumFactor
50770   .PDFCompressionGreyCompressionJPEGMediumFactor = sOptions.PDFCompressionGreyCompressionJPEGMediumFactor
50780   .PDFCompressionGreyCompressionJPEGMinimumFactor = sOptions.PDFCompressionGreyCompressionJPEGMinimumFactor
50790   .PDFCompressionGreyResample = sOptions.PDFCompressionGreyResample
50800   .PDFCompressionGreyResampleChoice = sOptions.PDFCompressionGreyResampleChoice
50810   .PDFCompressionGreyResolution = sOptions.PDFCompressionGreyResolution
50820   .PDFCompressionMonoCompression = sOptions.PDFCompressionMonoCompression
50830   .PDFCompressionMonoCompressionChoice = sOptions.PDFCompressionMonoCompressionChoice
50840   .PDFCompressionMonoResample = sOptions.PDFCompressionMonoResample
50850   .PDFCompressionMonoResampleChoice = sOptions.PDFCompressionMonoResampleChoice
50860   .PDFCompressionMonoResolution = sOptions.PDFCompressionMonoResolution
50870   .PDFCompressionTextCompression = sOptions.PDFCompressionTextCompression
50880   .PDFDisallowCopy = sOptions.PDFDisallowCopy
50890   .PDFDisallowModifyAnnotations = sOptions.PDFDisallowModifyAnnotations
50900   .PDFDisallowModifyContents = sOptions.PDFDisallowModifyContents
50910   .PDFDisallowPrinting = sOptions.PDFDisallowPrinting
50920   .PDFEncryptor = sOptions.PDFEncryptor
50930   .PDFFontsEmbedAll = sOptions.PDFFontsEmbedAll
50940   .PDFFontsSubSetFonts = sOptions.PDFFontsSubSetFonts
50950   .PDFFontsSubSetFontsPercent = sOptions.PDFFontsSubSetFontsPercent
50960   .PDFGeneralASCII85 = sOptions.PDFGeneralASCII85
50970   .PDFGeneralAutorotate = sOptions.PDFGeneralAutorotate
50980   .PDFGeneralCompatibility = sOptions.PDFGeneralCompatibility
50990   .PDFGeneralDefault = sOptions.PDFGeneralDefault
51000   .PDFGeneralOverprint = sOptions.PDFGeneralOverprint
51010   .PDFGeneralResolution = sOptions.PDFGeneralResolution
51020   .PDFHighEncryption = sOptions.PDFHighEncryption
51030   .PDFLowEncryption = sOptions.PDFLowEncryption
51040   .PDFOptimize = sOptions.PDFOptimize
51050   .PDFOwnerPass = sOptions.PDFOwnerPass
51060   .PDFOwnerPasswordString = sOptions.PDFOwnerPasswordString
51070   .PDFPageLayout = sOptions.PDFPageLayout
51080   .PDFPageMode = sOptions.PDFPageMode
51090   .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
51100   .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
51110   .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
51120   .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
51130   .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
51140   .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
51150   .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
51160   .PDFSigningSignatureOnPage = sOptions.PDFSigningSignatureOnPage
51170   .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
51180   .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
51190   .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
51200   .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
51210   .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
51220   .PDFSigningTimeServerUrl = sOptions.PDFSigningTimeServerUrl
51230   .PDFStartPage = sOptions.PDFStartPage
51240   .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
51250   .PDFUserPass = sOptions.PDFUserPass
51260   .PDFUserPasswordString = sOptions.PDFUserPasswordString
51270   .PDFUseSecurity = sOptions.PDFUseSecurity
51280   .PNGColorscount = sOptions.PNGColorscount
51290   .PNGResolution = sOptions.PNGResolution
51300   .PrintAfterSaving = sOptions.PrintAfterSaving
51310   .PrintAfterSavingBitsPerPixel = sOptions.PrintAfterSavingBitsPerPixel
51320   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
51330   .PrintAfterSavingMaxResolution = sOptions.PrintAfterSavingMaxResolution
51340   .PrintAfterSavingMaxResolutionEnabled = sOptions.PrintAfterSavingMaxResolutionEnabled
51350   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
51360   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
51370   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51380   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51390   .PrinterStop = sOptions.PrinterStop
51400   .ProcessPriority = sOptions.ProcessPriority
51410   .ProgramFont = sOptions.ProgramFont
51420   .ProgramFontCharset = sOptions.ProgramFontCharset
51430   .ProgramFontSize = sOptions.ProgramFontSize
51440   .PSDColorsCount = sOptions.PSDColorsCount
51450   .PSDResolution = sOptions.PSDResolution
51460   .PSLanguageLevel = sOptions.PSLanguageLevel
51470   .RAWColorsCount = sOptions.RAWColorsCount
51480   .RAWResolution = sOptions.RAWResolution
51490   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51500   .RemoveSpaces = sOptions.RemoveSpaces
51510   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51520   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51530   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51540   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51550   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51560   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51570   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51580   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51590   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51600   .SaveFilename = sOptions.SaveFilename
51610   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51620   .SendMailMethod = sOptions.SendMailMethod
51630   .ShowAnimation = sOptions.ShowAnimation
51640   .StampFontColor = sOptions.StampFontColor
51650   .StampFontname = sOptions.StampFontname
51660   .StampFontsize = sOptions.StampFontsize
51670   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51680   .StampString = sOptions.StampString
51690   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51700   .StandardAuthor = sOptions.StandardAuthor
51710   .StandardCreationdate = sOptions.StandardCreationdate
51720   .StandardDateformat = sOptions.StandardDateformat
51730   .StandardKeywords = sOptions.StandardKeywords
51740   .StandardMailDomain = sOptions.StandardMailDomain
51750   .StandardModifydate = sOptions.StandardModifydate
51760   .StandardSaveformat = sOptions.StandardSaveformat
51770   .StandardSubject = sOptions.StandardSubject
51780   .StandardTitle = sOptions.StandardTitle
51790   .SVGResolution = sOptions.SVGResolution
51800   .TIFFColorscount = sOptions.TIFFColorscount
51810   .TIFFResolution = sOptions.TIFFResolution
51820   .Toolbars = sOptions.Toolbars
51830   .UpdateInterval = sOptions.UpdateInterval
51840   .UseAutosave = sOptions.UseAutosave
51850   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51860   .UseCreationDateNow = sOptions.UseCreationDateNow
51870   .UseCustomPaperSize = sOptions.UseCustomPaperSize
51880   .UseFixPapersize = sOptions.UseFixPapersize
51890   .UseStandardAuthor = sOptions.UseStandardAuthor
51900  End With
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
50210   Case "DISABLEUPDATECHECK": Options1.DisableUpdateCheck = value
50220   Case "DONTUSEDOCUMENTSETTINGS": Options1.DontUseDocumentSettings = value
50230   Case "EDITWITHPDFARCHITECT": Options1.EditWithPDFArchitect = value
50240   Case "EPSLANGUAGELEVEL": Options1.EPSLanguageLevel = value
50250   Case "FILENAMESUBSTITUTIONS": Options1.FilenameSubstitutions = value
50260   Case "FILENAMESUBSTITUTIONSONLYINTITLE": Options1.FilenameSubstitutionsOnlyInTitle = value
50270   Case "JPEGCOLORSCOUNT": Options1.JPEGColorscount = value
50280   Case "JPEGQUALITY": Options1.JPEGQuality = value
50290   Case "JPEGRESOLUTION": Options1.JPEGResolution = value
50300   Case "LANGUAGE": Options1.Language = value
50310   Case "LASTSAVEDIRECTORY": Options1.LastSaveDirectory = value
50320   Case "LASTUPDATECHECK": Options1.LastUpdateCheck = value
50330   Case "LOGGING": Options1.Logging = value
50340   Case "LOGLINES": Options1.LogLines = value
50350   Case "MAXIMUMCOUNTOFPDFARCHITECTTOOLTIP": Options1.MaximumCountOfPDFArchitectToolTip = value
50360   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": Options1.NoConfirmMessageSwitchingDefaultprinter = value
50370   Case "NOPROCESSINGATSTARTUP": Options1.NoProcessingAtStartup = value
50380   Case "NOPSCHECK": Options1.NoPSCheck = value
50390   Case "ONEFILEPERPAGE": Options1.OneFilePerPage = value
50400   Case "OPENOUTPUTFILE": Options1.OpenOutputFile = value
50410   Case "OPTIONSDESIGN": Options1.OptionsDesign = value
50420   Case "OPTIONSENABLED": Options1.OptionsEnabled = value
50430   Case "OPTIONSVISIBLE": Options1.OptionsVisible = value
50440   Case "PAPERSIZE": Options1.Papersize = value
50450   Case "PCLCOLORSCOUNT": Options1.PCLColorsCount = value
50460   Case "PCLRESOLUTION": Options1.PCLResolution = value
50470   Case "PCXCOLORSCOUNT": Options1.PCXColorscount = value
50480   Case "PCXRESOLUTION": Options1.PCXResolution = value
50490   Case "PDFAES128ENCRYPTION": Options1.PDFAes128Encryption = value
50500   Case "PDFALLOWASSEMBLY": Options1.PDFAllowAssembly = value
50510   Case "PDFALLOWDEGRADEDPRINTING": Options1.PDFAllowDegradedPrinting = value
50520   Case "PDFALLOWFILLIN": Options1.PDFAllowFillIn = value
50530   Case "PDFALLOWSCREENREADERS": Options1.PDFAllowScreenReaders = value
50540   Case "PDFCOLORSCMYKTORGB": Options1.PDFColorsCMYKToRGB = value
50550   Case "PDFCOLORSCOLORMODEL": Options1.PDFColorsColorModel = value
50560   Case "PDFCOLORSPRESERVEHALFTONE": Options1.PDFColorsPreserveHalftone = value
50570   Case "PDFCOLORSPRESERVEOVERPRINT": Options1.PDFColorsPreserveOverprint = value
50580   Case "PDFCOLORSPRESERVETRANSFER": Options1.PDFColorsPreserveTransfer = value
50590   Case "PDFCOMPRESSIONCOLORCOMPRESSION": Options1.PDFCompressionColorCompression = value
50600   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": Options1.PDFCompressionColorCompressionChoice = value
50610   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": Options1.PDFCompressionColorCompressionJPEGHighFactor = value
50620   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": Options1.PDFCompressionColorCompressionJPEGLowFactor = value
50630   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMANUALFACTOR": Options1.PDFCompressionColorCompressionJPEGManualFactor = value
50640   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": Options1.PDFCompressionColorCompressionJPEGMaximumFactor = value
50650   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": Options1.PDFCompressionColorCompressionJPEGMediumFactor = value
50660   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": Options1.PDFCompressionColorCompressionJPEGMinimumFactor = value
50670   Case "PDFCOMPRESSIONCOLORRESAMPLE": Options1.PDFCompressionColorResample = value
50680   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": Options1.PDFCompressionColorResampleChoice = value
50690   Case "PDFCOMPRESSIONCOLORRESOLUTION": Options1.PDFCompressionColorResolution = value
50700   Case "PDFCOMPRESSIONGREYCOMPRESSION": Options1.PDFCompressionGreyCompression = value
50710   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": Options1.PDFCompressionGreyCompressionChoice = value
50720   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": Options1.PDFCompressionGreyCompressionJPEGHighFactor = value
50730   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": Options1.PDFCompressionGreyCompressionJPEGLowFactor = value
50740   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMANUALFACTOR": Options1.PDFCompressionGreyCompressionJPEGManualFactor = value
50750   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": Options1.PDFCompressionGreyCompressionJPEGMaximumFactor = value
50760   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": Options1.PDFCompressionGreyCompressionJPEGMediumFactor = value
50770   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": Options1.PDFCompressionGreyCompressionJPEGMinimumFactor = value
50780   Case "PDFCOMPRESSIONGREYRESAMPLE": Options1.PDFCompressionGreyResample = value
50790   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": Options1.PDFCompressionGreyResampleChoice = value
50800   Case "PDFCOMPRESSIONGREYRESOLUTION": Options1.PDFCompressionGreyResolution = value
50810   Case "PDFCOMPRESSIONMONOCOMPRESSION": Options1.PDFCompressionMonoCompression = value
50820   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": Options1.PDFCompressionMonoCompressionChoice = value
50830   Case "PDFCOMPRESSIONMONORESAMPLE": Options1.PDFCompressionMonoResample = value
50840   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": Options1.PDFCompressionMonoResampleChoice = value
50850   Case "PDFCOMPRESSIONMONORESOLUTION": Options1.PDFCompressionMonoResolution = value
50860   Case "PDFCOMPRESSIONTEXTCOMPRESSION": Options1.PDFCompressionTextCompression = value
50870   Case "PDFDISALLOWCOPY": Options1.PDFDisallowCopy = value
50880   Case "PDFDISALLOWMODIFYANNOTATIONS": Options1.PDFDisallowModifyAnnotations = value
50890   Case "PDFDISALLOWMODIFYCONTENTS": Options1.PDFDisallowModifyContents = value
50900   Case "PDFDISALLOWPRINTING": Options1.PDFDisallowPrinting = value
50910   Case "PDFENCRYPTOR": Options1.PDFEncryptor = value
50920   Case "PDFFONTSEMBEDALL": Options1.PDFFontsEmbedAll = value
50930   Case "PDFFONTSSUBSETFONTS": Options1.PDFFontsSubSetFonts = value
50940   Case "PDFFONTSSUBSETFONTSPERCENT": Options1.PDFFontsSubSetFontsPercent = value
50950   Case "PDFGENERALASCII85": Options1.PDFGeneralASCII85 = value
50960   Case "PDFGENERALAUTOROTATE": Options1.PDFGeneralAutorotate = value
50970   Case "PDFGENERALCOMPATIBILITY": Options1.PDFGeneralCompatibility = value
50980   Case "PDFGENERALDEFAULT": Options1.PDFGeneralDefault = value
50990   Case "PDFGENERALOVERPRINT": Options1.PDFGeneralOverprint = value
51000   Case "PDFGENERALRESOLUTION": Options1.PDFGeneralResolution = value
51010   Case "PDFHIGHENCRYPTION": Options1.PDFHighEncryption = value
51020   Case "PDFLOWENCRYPTION": Options1.PDFLowEncryption = value
51030   Case "PDFOPTIMIZE": Options1.PDFOptimize = value
51040   Case "PDFOWNERPASS": Options1.PDFOwnerPass = value
51050   Case "PDFOWNERPASSWORDSTRING": Options1.PDFOwnerPasswordString = value
51060   Case "PDFPAGELAYOUT": Options1.PDFPageLayout = value
51070   Case "PDFPAGEMODE": Options1.PDFPageMode = value
51080   Case "PDFSIGNINGMULTISIGNATURE": Options1.PDFSigningMultiSignature = value
51090   Case "PDFSIGNINGPFXFILE": Options1.PDFSigningPFXFile = value
51100   Case "PDFSIGNINGPFXFILEPASSWORD": Options1.PDFSigningPFXFilePassword = value
51110   Case "PDFSIGNINGSIGNATURECONTACT": Options1.PDFSigningSignatureContact = value
51120   Case "PDFSIGNINGSIGNATURELEFTX": Options1.PDFSigningSignatureLeftX = value
51130   Case "PDFSIGNINGSIGNATURELEFTY": Options1.PDFSigningSignatureLeftY = value
51140   Case "PDFSIGNINGSIGNATURELOCATION": Options1.PDFSigningSignatureLocation = value
51150   Case "PDFSIGNINGSIGNATUREONPAGE": Options1.PDFSigningSignatureOnPage = value
51160   Case "PDFSIGNINGSIGNATUREREASON": Options1.PDFSigningSignatureReason = value
51170   Case "PDFSIGNINGSIGNATURERIGHTX": Options1.PDFSigningSignatureRightX = value
51180   Case "PDFSIGNINGSIGNATURERIGHTY": Options1.PDFSigningSignatureRightY = value
51190   Case "PDFSIGNINGSIGNATUREVISIBLE": Options1.PDFSigningSignatureVisible = value
51200   Case "PDFSIGNINGSIGNPDF": Options1.PDFSigningSignPDF = value
51210   Case "PDFSIGNINGTIMESERVERURL": Options1.PDFSigningTimeServerUrl = value
51220   Case "PDFSTARTPAGE": Options1.PDFStartPage = value
51230   Case "PDFUPDATEMETADATA": Options1.PDFUpdateMetadata = value
51240   Case "PDFUSERPASS": Options1.PDFUserPass = value
51250   Case "PDFUSERPASSWORDSTRING": Options1.PDFUserPasswordString = value
51260   Case "PDFUSESECURITY": Options1.PDFUseSecurity = value
51270   Case "PNGCOLORSCOUNT": Options1.PNGColorscount = value
51280   Case "PNGRESOLUTION": Options1.PNGResolution = value
51290   Case "PRINTAFTERSAVING": Options1.PrintAfterSaving = value
51300   Case "PRINTAFTERSAVINGBITSPERPIXEL": Options1.PrintAfterSavingBitsPerPixel = value
51310   Case "PRINTAFTERSAVINGDUPLEX": Options1.PrintAfterSavingDuplex = value
51320   Case "PRINTAFTERSAVINGMAXRESOLUTION": Options1.PrintAfterSavingMaxResolution = value
51330   Case "PRINTAFTERSAVINGMAXRESOLUTIONENABLED": Options1.PrintAfterSavingMaxResolutionEnabled = value
51340   Case "PRINTAFTERSAVINGNOCANCEL": Options1.PrintAfterSavingNoCancel = value
51350   Case "PRINTAFTERSAVINGPRINTER": Options1.PrintAfterSavingPrinter = value
51360   Case "PRINTAFTERSAVINGQUERYUSER": Options1.PrintAfterSavingQueryUser = value
51370   Case "PRINTAFTERSAVINGTUMBLE": Options1.PrintAfterSavingTumble = value
51380   Case "PRINTERSTOP": Options1.PrinterStop = value
51390   Case "PROCESSPRIORITY": Options1.ProcessPriority = value
51400   Case "PROGRAMFONT": Options1.ProgramFont = value
51410   Case "PROGRAMFONTCHARSET": Options1.ProgramFontCharset = value
51420   Case "PROGRAMFONTSIZE": Options1.ProgramFontSize = value
51430   Case "PSDCOLORSCOUNT": Options1.PSDColorsCount = value
51440   Case "PSDRESOLUTION": Options1.PSDResolution = value
51450   Case "PSLANGUAGELEVEL": Options1.PSLanguageLevel = value
51460   Case "RAWCOLORSCOUNT": Options1.RAWColorsCount = value
51470   Case "RAWRESOLUTION": Options1.RAWResolution = value
51480   Case "REMOVEALLKNOWNFILEEXTENSIONS": Options1.RemoveAllKnownFileExtensions = value
51490   Case "REMOVESPACES": Options1.RemoveSpaces = value
51500   Case "RUNPROGRAMAFTERSAVING": Options1.RunProgramAfterSaving = value
51510   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": Options1.RunProgramAfterSavingProgramname = value
51520   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": Options1.RunProgramAfterSavingProgramParameters = value
51530   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": Options1.RunProgramAfterSavingWaitUntilReady = value
51540   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": Options1.RunProgramAfterSavingWindowstyle = value
51550   Case "RUNPROGRAMBEFORESAVING": Options1.RunProgramBeforeSaving = value
51560   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": Options1.RunProgramBeforeSavingProgramname = value
51570   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": Options1.RunProgramBeforeSavingProgramParameters = value
51580   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": Options1.RunProgramBeforeSavingWindowstyle = value
51590   Case "SAVEFILENAME": Options1.SaveFilename = value
51600   Case "SENDEMAILAFTERAUTOSAVING": Options1.SendEmailAfterAutoSaving = value
51610   Case "SENDMAILMETHOD": Options1.SendMailMethod = value
51620   Case "SHOWANIMATION": Options1.ShowAnimation = value
51630   Case "STAMPFONTCOLOR": Options1.StampFontColor = value
51640   Case "STAMPFONTNAME": Options1.StampFontname = value
51650   Case "STAMPFONTSIZE": Options1.StampFontsize = value
51660   Case "STAMPOUTLINEFONTTHICKNESS": Options1.StampOutlineFontthickness = value
51670   Case "STAMPSTRING": Options1.StampString = value
51680   Case "STAMPUSEOUTLINEFONT": Options1.StampUseOutlineFont = value
51690   Case "STANDARDAUTHOR": Options1.StandardAuthor = value
51700   Case "STANDARDCREATIONDATE": Options1.StandardCreationdate = value
51710   Case "STANDARDDATEFORMAT": Options1.StandardDateformat = value
51720   Case "STANDARDKEYWORDS": Options1.StandardKeywords = value
51730   Case "STANDARDMAILDOMAIN": Options1.StandardMailDomain = value
51740   Case "STANDARDMODIFYDATE": Options1.StandardModifydate = value
51750   Case "STANDARDSAVEFORMAT": Options1.StandardSaveformat = value
51760   Case "STANDARDSUBJECT": Options1.StandardSubject = value
51770   Case "STANDARDTITLE": Options1.StandardTitle = value
51780   Case "SVGRESOLUTION": Options1.SVGResolution = value
51790   Case "TIFFCOLORSCOUNT": Options1.TIFFColorscount = value
51800   Case "TIFFRESOLUTION": Options1.TIFFResolution = value
51810   Case "TOOLBARS": Options1.Toolbars = value
51820   Case "UPDATEINTERVAL": Options1.UpdateInterval = value
51830   Case "USEAUTOSAVE": Options1.UseAutosave = value
51840   Case "USEAUTOSAVEDIRECTORY": Options1.UseAutosaveDirectory = value
51850   Case "USECREATIONDATENOW": Options1.UseCreationDateNow = value
51860   Case "USECUSTOMPAPERSIZE": Options1.UseCustomPaperSize = value
51870   Case "USEFIXPAPERSIZE": Options1.UseFixPapersize = value
51880   Case "USESTANDARDAUTHOR": Options1.UseStandardAuthor = value
51890   Case Else:
51900    mError.Number = 3
51910    mError.Description = Replace$(Replace$(ErrDescr3, "%1", PropertyName), "%2", "in cLetOption")
51920    RaiseEvent eError
51930  End Select
51940  If ProfileName = "" Then
51950    Options = Options1
51960   Else
51970    SaveOptions Options1, ProfileName
51980  End If
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
50210   .DisableUpdateCheck = sOptions.DisableUpdateCheck
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
50390   .OneFilePerPage = sOptions.OneFilePerPage
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
51060   .PDFPageLayout = sOptions.PDFPageLayout
51070   .PDFPageMode = sOptions.PDFPageMode
51080   .PDFSigningMultiSignature = sOptions.PDFSigningMultiSignature
51090   .PDFSigningPFXFile = sOptions.PDFSigningPFXFile
51100   .PDFSigningPFXFilePassword = sOptions.PDFSigningPFXFilePassword
51110   .PDFSigningSignatureContact = sOptions.PDFSigningSignatureContact
51120   .PDFSigningSignatureLeftX = sOptions.PDFSigningSignatureLeftX
51130   .PDFSigningSignatureLeftY = sOptions.PDFSigningSignatureLeftY
51140   .PDFSigningSignatureLocation = sOptions.PDFSigningSignatureLocation
51150   .PDFSigningSignatureOnPage = sOptions.PDFSigningSignatureOnPage
51160   .PDFSigningSignatureReason = sOptions.PDFSigningSignatureReason
51170   .PDFSigningSignatureRightX = sOptions.PDFSigningSignatureRightX
51180   .PDFSigningSignatureRightY = sOptions.PDFSigningSignatureRightY
51190   .PDFSigningSignatureVisible = sOptions.PDFSigningSignatureVisible
51200   .PDFSigningSignPDF = sOptions.PDFSigningSignPDF
51210   .PDFSigningTimeServerUrl = sOptions.PDFSigningTimeServerUrl
51220   .PDFStartPage = sOptions.PDFStartPage
51230   .PDFUpdateMetadata = sOptions.PDFUpdateMetadata
51240   .PDFUserPass = sOptions.PDFUserPass
51250   .PDFUserPasswordString = sOptions.PDFUserPasswordString
51260   .PDFUseSecurity = sOptions.PDFUseSecurity
51270   .PNGColorscount = sOptions.PNGColorscount
51280   .PNGResolution = sOptions.PNGResolution
51290   .PrintAfterSaving = sOptions.PrintAfterSaving
51300   .PrintAfterSavingBitsPerPixel = sOptions.PrintAfterSavingBitsPerPixel
51310   .PrintAfterSavingDuplex = sOptions.PrintAfterSavingDuplex
51320   .PrintAfterSavingMaxResolution = sOptions.PrintAfterSavingMaxResolution
51330   .PrintAfterSavingMaxResolutionEnabled = sOptions.PrintAfterSavingMaxResolutionEnabled
51340   .PrintAfterSavingNoCancel = sOptions.PrintAfterSavingNoCancel
51350   .PrintAfterSavingPrinter = sOptions.PrintAfterSavingPrinter
51360   .PrintAfterSavingQueryUser = sOptions.PrintAfterSavingQueryUser
51370   .PrintAfterSavingTumble = sOptions.PrintAfterSavingTumble
51380   .PrinterStop = sOptions.PrinterStop
51390   .ProcessPriority = sOptions.ProcessPriority
51400   .ProgramFont = sOptions.ProgramFont
51410   .ProgramFontCharset = sOptions.ProgramFontCharset
51420   .ProgramFontSize = sOptions.ProgramFontSize
51430   .PSDColorsCount = sOptions.PSDColorsCount
51440   .PSDResolution = sOptions.PSDResolution
51450   .PSLanguageLevel = sOptions.PSLanguageLevel
51460   .RAWColorsCount = sOptions.RAWColorsCount
51470   .RAWResolution = sOptions.RAWResolution
51480   .RemoveAllKnownFileExtensions = sOptions.RemoveAllKnownFileExtensions
51490   .RemoveSpaces = sOptions.RemoveSpaces
51500   .RunProgramAfterSaving = sOptions.RunProgramAfterSaving
51510   .RunProgramAfterSavingProgramname = sOptions.RunProgramAfterSavingProgramname
51520   .RunProgramAfterSavingProgramParameters = sOptions.RunProgramAfterSavingProgramParameters
51530   .RunProgramAfterSavingWaitUntilReady = sOptions.RunProgramAfterSavingWaitUntilReady
51540   .RunProgramAfterSavingWindowstyle = sOptions.RunProgramAfterSavingWindowstyle
51550   .RunProgramBeforeSaving = sOptions.RunProgramBeforeSaving
51560   .RunProgramBeforeSavingProgramname = sOptions.RunProgramBeforeSavingProgramname
51570   .RunProgramBeforeSavingProgramParameters = sOptions.RunProgramBeforeSavingProgramParameters
51580   .RunProgramBeforeSavingWindowstyle = sOptions.RunProgramBeforeSavingWindowstyle
51590   .SaveFilename = sOptions.SaveFilename
51600   .SendEmailAfterAutoSaving = sOptions.SendEmailAfterAutoSaving
51610   .SendMailMethod = sOptions.SendMailMethod
51620   .ShowAnimation = sOptions.ShowAnimation
51630   .StampFontColor = sOptions.StampFontColor
51640   .StampFontname = sOptions.StampFontname
51650   .StampFontsize = sOptions.StampFontsize
51660   .StampOutlineFontthickness = sOptions.StampOutlineFontthickness
51670   .StampString = sOptions.StampString
51680   .StampUseOutlineFont = sOptions.StampUseOutlineFont
51690   .StandardAuthor = sOptions.StandardAuthor
51700   .StandardCreationdate = sOptions.StandardCreationdate
51710   .StandardDateformat = sOptions.StandardDateformat
51720   .StandardKeywords = sOptions.StandardKeywords
51730   .StandardMailDomain = sOptions.StandardMailDomain
51740   .StandardModifydate = sOptions.StandardModifydate
51750   .StandardSaveformat = sOptions.StandardSaveformat
51760   .StandardSubject = sOptions.StandardSubject
51770   .StandardTitle = sOptions.StandardTitle
51780   .SVGResolution = sOptions.SVGResolution
51790   .TIFFColorscount = sOptions.TIFFColorscount
51800   .TIFFResolution = sOptions.TIFFResolution
51810   .Toolbars = sOptions.Toolbars
51820   .UpdateInterval = sOptions.UpdateInterval
51830   .UseAutosave = sOptions.UseAutosave
51840   .UseAutosaveDirectory = sOptions.UseAutosaveDirectory
51850   .UseCreationDateNow = sOptions.UseCreationDateNow
51860   .UseCustomPaperSize = sOptions.UseCustomPaperSize
51870   .UseFixPapersize = sOptions.UseFixPapersize
51880   .UseStandardAuthor = sOptions.UseStandardAuthor
51890  End With
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

