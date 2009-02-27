Attribute VB_Name = "modOptions"
Option Explicit

' Automatically generated with DeveloperTool by Frank Heindörfer
' 2003 - 2007
' Email: thesmilyface@users.sourceforge.net

Public Type tOptions
 AdditionalGhostscriptParameters As String
 AdditionalGhostscriptSearchpath As String
 AddWindowsFontpath As Long
 AutosaveDirectory As String
 AutosaveFilename As String
 AutosaveFormat As Long
 AutosaveStartStandardProgram As Long
 BMPColorscount As Long
 BMPResolution As Long
 ClientComputerResolveIPAddress As Long
 Counter As Currency
 DeviceHeightPoints As Double
 DeviceWidthPoints As Double
 DirectoryGhostscriptBinaries As String
 DirectoryGhostscriptFonts As String
 DirectoryGhostscriptLibraries As String
 DirectoryGhostscriptResource As String
 DisableEmail As Long
 DontUseDocumentSettings As Long
 EPSLanguageLevel As Long
 FilenameSubstitutions As String
 FilenameSubstitutionsOnlyInTitle As Long
 JPEGColorscount As Long
 JPEGQuality As Long
 JPEGResolution As Long
 Language As String
 LastSaveDirectory As String
 LastUpdateCheck As String
 Logging As Long
 LogLines As Long
 NoConfirmMessageSwitchingDefaultprinter As Long
 NoProcessingAtStartup As Long
 NoPSCheck As Long
 OnePagePerFile As Long
 OptionsDesign As Long
 OptionsEnabled As Long
 OptionsVisible As Long
 Papersize As String
 PCLColorsCount As Long
 PCLResolution As Long
 PCXColorscount As Long
 PCXResolution As Long
 PDFAllowAssembly As Long
 PDFAllowDegradedPrinting As Long
 PDFAllowFillIn As Long
 PDFAllowScreenReaders As Long
 PDFColorsCMYKToRGB As Long
 PDFColorsColorModel As Long
 PDFColorsPreserveHalftone As Long
 PDFColorsPreserveOverprint As Long
 PDFColorsPreserveTransfer As Long
 PDFCompressionColorCompression As Long
 PDFCompressionColorCompressionChoice As Long
 PDFCompressionColorCompressionJPEGHighFactor As Double
 PDFCompressionColorCompressionJPEGLowFactor As Double
 PDFCompressionColorCompressionJPEGMaximumFactor As Double
 PDFCompressionColorCompressionJPEGMediumFactor As Double
 PDFCompressionColorCompressionJPEGMinimumFactor As Double
 PDFCompressionColorResample As Long
 PDFCompressionColorResampleChoice As Long
 PDFCompressionColorResolution As Long
 PDFCompressionGreyCompression As Long
 PDFCompressionGreyCompressionChoice As Long
 PDFCompressionGreyCompressionJPEGHighFactor As Double
 PDFCompressionGreyCompressionJPEGLowFactor As Double
 PDFCompressionGreyCompressionJPEGMaximumFactor As Double
 PDFCompressionGreyCompressionJPEGMediumFactor As Double
 PDFCompressionGreyCompressionJPEGMinimumFactor As Double
 PDFCompressionGreyResample As Long
 PDFCompressionGreyResampleChoice As Long
 PDFCompressionGreyResolution As Long
 PDFCompressionMonoCompression As Long
 PDFCompressionMonoCompressionChoice As Long
 PDFCompressionMonoResample As Long
 PDFCompressionMonoResampleChoice As Long
 PDFCompressionMonoResolution As Long
 PDFCompressionTextCompression As Long
 PDFDisallowCopy As Long
 PDFDisallowModifyAnnotations As Long
 PDFDisallowModifyContents As Long
 PDFDisallowPrinting As Long
 PDFEncryptor As Long
 PDFFontsEmbedAll As Long
 PDFFontsSubSetFonts As Long
 PDFFontsSubSetFontsPercent As Long
 PDFGeneralASCII85 As Long
 PDFGeneralAutorotate As Long
 PDFGeneralCompatibility As Long
 PDFGeneralDefault As Long
 PDFGeneralOverprint As Long
 PDFGeneralResolution As Long
 PDFHighEncryption As Long
 PDFLowEncryption As Long
 PDFOptimize As Long
 PDFOwnerPass As Long
 PDFOwnerPasswordString As String
 PDFSigningMultiSignature As Long
 PDFSigningPFXFile As String
 PDFSigningPFXFilePassword As String
 PDFSigningSignatureContact As String
 PDFSigningSignatureLeftX As Double
 PDFSigningSignatureLeftY As Double
 PDFSigningSignatureLocation As String
 PDFSigningSignatureReason As String
 PDFSigningSignatureRightX As Double
 PDFSigningSignatureRightY As Double
 PDFSigningSignatureVisible As Long
 PDFSigningSignPDF As Long
 PDFUpdateMetadata As Long
 PDFUserPass As Long
 PDFUserPasswordString As String
 PDFUseSecurity As Long
 PNGColorscount As Long
 PNGResolution As Long
 PrintAfterSaving As Long
 PrintAfterSavingDuplex As Long
 PrintAfterSavingNoCancel As Long
 PrintAfterSavingPrinter As String
 PrintAfterSavingQueryUser As Long
 PrintAfterSavingTumble As Long
 PrinterStop As Long
 PrinterTemppath As String
 ProcessPriority As Long
 ProgramFont As String
 ProgramFontCharset As Long
 ProgramFontSize As Long
 PSDColorsCount As Long
 PSDResolution As Long
 PSLanguageLevel As Long
 RAWColorsCount As Long
 RAWResolution As Long
 RemoveAllKnownFileExtensions As Long
 RemoveSpaces As Long
 RunProgramAfterSaving As Long
 RunProgramAfterSavingProgramname As String
 RunProgramAfterSavingProgramParameters As String
 RunProgramAfterSavingWaitUntilReady As Long
 RunProgramAfterSavingWindowstyle As Long
 RunProgramBeforeSaving As Long
 RunProgramBeforeSavingProgramname As String
 RunProgramBeforeSavingProgramParameters As String
 RunProgramBeforeSavingWindowstyle As Long
 SaveFilename As String
 SendEmailAfterAutoSaving As Long
 SendMailMethod As Long
 ShowAnimation As Long
 StampFontColor As String
 StampFontname As String
 StampFontsize As Long
 StampOutlineFontthickness As Long
 StampString As String
 StampUseOutlineFont As Long
 StandardAuthor As String
 StandardCreationdate As String
 StandardDateformat As String
 StandardKeywords As String
 StandardMailDomain As String
 StandardModifydate As String
 StandardSaveformat As Long
 StandardSubject As String
 StandardTitle As String
 StartStandardProgram As Long
 TIFFColorscount As Long
 TIFFResolution As Long
 Toolbars As Long
 UpdateInterval As Long
 UseAutosave As Long
 UseAutosaveDirectory As Long
 UseCreationDateNow As Long
 UseCustomPaperSize As String
 UseFixPapersize As Long
 UseStandardAuthor As Long
End Type

Public Options As tOptions

Public Function StandardOptions() As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Const Hash As String = "0001B4FD-9EA3-4D90-A79E-FD14BA3AB01D"
50020  Dim myOptions As tOptions, reg As clsRegistry
50030  With myOptions
50040   .AdditionalGhostscriptParameters = vbNullString
50050   .AdditionalGhostscriptSearchpath = vbNullString
50060   .AddWindowsFontpath = "1"
50070   If InstalledAsServer Then
50080     .AutosaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
50090    Else
50100     .AutosaveDirectory = "<MyFiles>"
50110   End If
50120   .AutosaveFilename = "<DateTime>"
50130   .AutosaveFormat = "0"
50140   .AutosaveStartStandardProgram = "0"
50150   .BMPColorscount = "1"
50160   .BMPResolution = "150"
50170   .ClientComputerResolveIPAddress = "0"
50180   .Counter = "0"
50190   .DeviceHeightPoints = Replace$("842", ".", GetDecimalChar)
50200   .DeviceWidthPoints = Replace$("595", ".", GetDecimalChar)
50210   Set reg = New clsRegistry
50220   reg.hkey = HKEY_LOCAL_MACHINE
50230   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50240   .DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50250   Set reg = Nothing
50260   Set reg = New clsRegistry
50270   reg.hkey = HKEY_LOCAL_MACHINE
50280   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50290   .DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50300   Set reg = Nothing
50310   Set reg = New clsRegistry
50320   reg.hkey = HKEY_LOCAL_MACHINE
50330   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50340   .DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50350   Set reg = Nothing
50360   Set reg = New clsRegistry
50370   reg.hkey = HKEY_LOCAL_MACHINE
50380   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50390   .DirectoryGhostscriptResource = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryResource"))
50400   Set reg = Nothing
50410   .DisableEmail = "0"
50420   .DontUseDocumentSettings = "0"
50430   .EPSLanguageLevel = "2"
50440   .FilenameSubstitutions = "Microsoft Word - \.docx\.doc\Microsoft Excel - \.xlsx\.xls\Microsoft PowerPoint - \.pptx\.ppt"
50450   .FilenameSubstitutionsOnlyInTitle = "1"
50460   .JPEGColorscount = "0"
50470   .JPEGQuality = "75"
50480   .JPEGResolution = "150"
50490   .Language = "english"
50500   If InstalledAsServer Then
50510     .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
50520    Else
50530     .LastSaveDirectory = "<MyFiles>"
50540   End If
50550   .LastUpdateCheck = vbNullString
50560   .Logging = "0"
50570   .LogLines = "100"
50580   .NoConfirmMessageSwitchingDefaultprinter = "0"
50590   .NoProcessingAtStartup = "0"
50600   .NoPSCheck = "0"
50610   .OnePagePerFile = "0"
50620   .OptionsDesign = "0"
50630   .OptionsEnabled = "1"
50640   .OptionsVisible = "1"
50650   .Papersize = "a4"
50660   .PCLColorsCount = "0"
50670   .PCLResolution = "150"
50680   .PCXColorscount = "0"
50690   .PCXResolution = "150"
50700   .PDFAllowAssembly = "0"
50710   .PDFAllowDegradedPrinting = "0"
50720   .PDFAllowFillIn = "0"
50730   .PDFAllowScreenReaders = "0"
50740   .PDFColorsCMYKToRGB = "0"
50750   .PDFColorsColorModel = "1"
50760   .PDFColorsPreserveHalftone = "0"
50770   .PDFColorsPreserveOverprint = "1"
50780   .PDFColorsPreserveTransfer = "1"
50790   .PDFCompressionColorCompression = "1"
50800   .PDFCompressionColorCompressionChoice = "0"
50810   .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
50820   .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
50830   .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
50840   .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
50850   .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
50860   .PDFCompressionColorResample = "0"
50870   .PDFCompressionColorResampleChoice = "0"
50880   .PDFCompressionColorResolution = "300"
50890   .PDFCompressionGreyCompression = "1"
50900   .PDFCompressionGreyCompressionChoice = "0"
50910   .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
50920   .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
50930   .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
50940   .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
50950   .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
50960   .PDFCompressionGreyResample = "0"
50970   .PDFCompressionGreyResampleChoice = "0"
50980   .PDFCompressionGreyResolution = "300"
50990   .PDFCompressionMonoCompression = "1"
51000   .PDFCompressionMonoCompressionChoice = "0"
51010   .PDFCompressionMonoResample = "0"
51020   .PDFCompressionMonoResampleChoice = "0"
51030   .PDFCompressionMonoResolution = "1200"
51040   .PDFCompressionTextCompression = "1"
51050   .PDFDisallowCopy = "1"
51060   .PDFDisallowModifyAnnotations = "0"
51070   .PDFDisallowModifyContents = "0"
51080   .PDFDisallowPrinting = "0"
51090   .PDFEncryptor = "0"
51100   .PDFFontsEmbedAll = "1"
51110   .PDFFontsSubSetFonts = "1"
51120   .PDFFontsSubSetFontsPercent = "100"
51130   .PDFGeneralASCII85 = "0"
51140   .PDFGeneralAutorotate = "2"
51150   .PDFGeneralCompatibility = "2"
51160   .PDFGeneralDefault = "0"
51170   .PDFGeneralOverprint = "0"
51180   .PDFGeneralResolution = "600"
51190   .PDFHighEncryption = "0"
51200   .PDFLowEncryption = "1"
51210   .PDFOptimize = "0"
51220   .PDFOwnerPass = "0"
51230   .PDFOwnerPasswordString = vbNullString
51240   .PDFSigningMultiSignature = "0"
51250   .PDFSigningPFXFile = vbNullString
51260   .PDFSigningPFXFilePassword = vbNullString
51270   .PDFSigningSignatureContact = vbNullString
51280   .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
51290   .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
51300   .PDFSigningSignatureLocation = vbNullString
51310   .PDFSigningSignatureReason = vbNullString
51320   .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
51330   .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
51340   .PDFSigningSignatureVisible = "0"
51350   .PDFSigningSignPDF = "0"
51360   .PDFUpdateMetadata = "1"
51370   .PDFUserPass = "0"
51380   .PDFUserPasswordString = vbNullString
51390   .PDFUseSecurity = "0"
51400   .PNGColorscount = "0"
51410   .PNGResolution = "150"
51420   .PrintAfterSaving = "0"
51430   .PrintAfterSavingDuplex = "0"
51440   .PrintAfterSavingNoCancel = "0"
51450   .PrintAfterSavingPrinter = vbNullString
51460   .PrintAfterSavingQueryUser = "0"
51470   .PrintAfterSavingTumble = "0"
51480   .PrinterStop = "0"
51490   If InstalledAsServer Then
51500     .PrinterTemppath = CompletePath(GetPDFCreatorApplicationPath) & "Temp\"
51510    Else
51520     .PrinterTemppath = "<Temp>PDFCreator\"
51530   End If
51540   .ProcessPriority = "1"
51550   .ProgramFont = "MS Sans Serif"
51560   .ProgramFontCharset = "0"
51570   .ProgramFontSize = "8"
51580   .PSDColorsCount = "0"
51590   .PSDResolution = "150"
51600   .PSLanguageLevel = "2"
51610   .RAWColorsCount = "0"
51620   .RAWResolution = "150"
51630   .RemoveAllKnownFileExtensions = "1"
51640   .RemoveSpaces = "1"
51650   .RunProgramAfterSaving = "0"
51660   .RunProgramAfterSavingProgramname = vbNullString
51670   .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
51680   .RunProgramAfterSavingWaitUntilReady = "1"
51690   .RunProgramAfterSavingWindowstyle = "1"
51700   .RunProgramBeforeSaving = "0"
51710   .RunProgramBeforeSavingProgramname = vbNullString
51720   .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
51730   .RunProgramBeforeSavingWindowstyle = "1"
51740   .SaveFilename = "<Title>"
51750   .SendEmailAfterAutoSaving = "0"
51760   .SendMailMethod = "0"
51770   .ShowAnimation = "1"
51780   .StampFontColor = "#FF0000"
51790   .StampFontname = "Arial"
51800   .StampFontsize = "48"
51810   .StampOutlineFontthickness = "0"
51820   .StampString = vbNullString
51830   .StampUseOutlineFont = "1"
51840   .StandardAuthor = vbNullString
51850   .StandardCreationdate = vbNullString
51860   .StandardDateformat = "YYYYMMDDHHNNSS"
51870   .StandardKeywords = vbNullString
51880   .StandardMailDomain = vbNullString
51890   .StandardModifydate = vbNullString
51900   .StandardSaveformat = "0"
51910   .StandardSubject = vbNullString
51920   .StandardTitle = vbNullString
51930   .StartStandardProgram = "1"
51940   .TIFFColorscount = "0"
51950   .TIFFResolution = "150"
51960   .Toolbars = "1"
51970   .UpdateInterval = "2"
51980   .UseAutosave = "0"
51990   .UseAutosaveDirectory = "1"
52000   .UseCreationDateNow = "0"
52010   .UseCustomPaperSize = "0"
52020   .UseFixPapersize = "0"
52030   .UseStandardAuthor = "0"
52040  End With
52050  If UseINI Then
52060    If Not IsWin9xMe Then
52070     myOptions = ReadOptionsINI(myOptions, CompletePath(GetDefaultAppData) & "PDFCreator.ini", False, False, False)
52080    End If
52090   Else
52100    If Not IsWin9xMe Then
52110     myOptions = ReadOptionsReg(myOptions, ".DEFAULT\Software\PDFCreator", HKEY_USERS, False, False)
52120    End If
52130  End If
52140  StandardOptions = myOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "StandardOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ReadOptions(Optional NoMsg As Boolean = False, Optional hProfile As hkey = HKEY_CURRENT_USER) As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim myOptions As tOptions, Str1 As String
50020  If InstalledAsServer Then
50030    If UseINI Then
50040      WriteToSpecialLogfile "INI-Read options: CommonAppData"
50050      myOptions = ReadOptionsINI(myOptions, CompletePath(GetCommonAppData) & "PDFCreator.ini", HKEY_LOCAL_MACHINE, NoMsg)
50060     Else
50070      WriteToSpecialLogfile "Reg-Read options: HKEY_LOCAL_MACHINE"
50080      myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", HKEY_LOCAL_MACHINE, NoMsg)
50090    End If
50100   Else
50110    If UseINI Then
50120      If Not IsWin9xMe Then
50130        WriteToSpecialLogfile "INI-Read options: DefaultAppData"
50140        myOptions = ReadOptionsINI(myOptions, CompletePath(GetDefaultAppData) & "PDFCreator.ini", HKEY_USERS, NoMsg)
50150        WriteToSpecialLogfile "INI-Read options: User"
50160        myOptions = ReadOptionsINI(myOptions, PDFCreatorINIFile, hProfile, NoMsg, False)
50170       Else
50180        WriteToSpecialLogfile "INI-Read options: User"
50190        myOptions = ReadOptionsINI(myOptions, PDFCreatorINIFile, hProfile, NoMsg)
50200      End If
50210      WriteToSpecialLogfile "INI-Read options: CommonAppData"
50220      myOptions = ReadOptionsINI(myOptions, CompletePath(GetCommonAppData) & "PDFCreator.ini", HKEY_LOCAL_MACHINE, NoMsg, False)
50230     Else
50240      If Not IsWin9xMe Then
50250        WriteToSpecialLogfile "Reg-Read options: HKEY_USERS"
50260        myOptions = ReadOptionsReg(myOptions, ".DEFAULT\Software\PDFCreator", HKEY_USERS, NoMsg)
50270        WriteToSpecialLogfile "Reg-Read options: HKEY_CURRENT_USER [" & hProfile & "]"
50280        myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", hProfile, NoMsg, False)
50290       Else
50300        WriteToSpecialLogfile "Reg-Read options: HKEY_CURRENT_USER [" & hProfile & "]"
50310        myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", hProfile, NoMsg)
50320      End If
50330      WriteToSpecialLogfile "Reg-Read options: HKEY_LOCAL_MACHINE"
50340      myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", HKEY_LOCAL_MACHINE, NoMsg, False)
50350    End If
50360  End If
50370  Str1 = "7777772E706466666F7267652E6F7267"
50380  myOptions = CorrectOptionsAfterLoading(myOptions)
50390  ReadOptions = myOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "ReadOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ReadOptionsINI(myOptions As tOptions, PDFCreatorINIFile As String, Optional hkey1 As hkey = HKEY_CURRENT_USER, Optional NoMsg As Boolean = False, Optional UseStandard As Boolean = True) As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI, tStr As String, hOpt As New clsHash
50020  ReadOptionsINI = myOptions
50030  Set ini = New clsINI
50040  ini.filename = PDFCreatorINIFile
50050  ini.Section = "Options"
50060  If ini.CheckIniFile = False Then
50070   If UseStandard Then
50080    ReadOptionsINI = StandardOptions
50090   End If
50100   Exit Function
50110  End If
50120  ReadINISection PDFCreatorINIFile, "Options", hOpt
50130  With myOptions
50140   tStr = hOpt.Retrieve("AdditionalGhostscriptParameters")
50150   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
50160     .AdditionalGhostscriptParameters = ""
50170    Else
50180     If LenB(tStr) > 0 Then
50190      .AdditionalGhostscriptParameters = tStr
50200     End If
50210   End If
50220   tStr = hOpt.Retrieve("AdditionalGhostscriptSearchpath")
50230   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
50240     .AdditionalGhostscriptSearchpath = ""
50250    Else
50260     If LenB(tStr) > 0 Then
50270      .AdditionalGhostscriptSearchpath = tStr
50280     End If
50290   End If
50300   tStr = hOpt.Retrieve("AddWindowsFontpath")
50310   If IsNumeric(tStr) Then
50320     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
50330       .AddWindowsFontpath = CLng(tStr)
50340      Else
50350       If UseStandard Then
50360        .AddWindowsFontpath = 1
50370       End If
50380     End If
50390    Else
50400     If UseStandard Then
50410      .AddWindowsFontpath = 1
50420     End If
50430   End If
50440   tStr = hOpt.Retrieve("AutosaveDirectory")
50450   If LenB(Trim$(tStr)) > 0 Then
50460     .AutosaveDirectory = CompletePath(tStr)
50470    Else
50480     If UseStandard Then
50490      If InstalledAsServer Then
50500        .AutosaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
50510       Else
50520        .AutosaveDirectory = "<MyFiles>"
50530      End If
50540     End If
50550   End If
50560   tStr = hOpt.Retrieve("AutosaveFilename")
50570   If LenB(tStr) = 0 And LenB("<DateTime>") > 0 And UseStandard Then
50580     .AutosaveFilename = "<DateTime>"
50590    Else
50600     If LenB(tStr) > 0 Then
50610      .AutosaveFilename = tStr
50620     End If
50630   End If
50640   tStr = hOpt.Retrieve("AutosaveFormat")
50650   If IsNumeric(tStr) Then
50660     If CLng(tStr) >= 0 And CLng(tStr) <= 13 Then
50670       .AutosaveFormat = CLng(tStr)
50680      Else
50690       If UseStandard Then
50700        .AutosaveFormat = 0
50710       End If
50720     End If
50730    Else
50740     If UseStandard Then
50750      .AutosaveFormat = 0
50760     End If
50770   End If
50780   tStr = hOpt.Retrieve("AutosaveStartStandardProgram")
50790   If IsNumeric(tStr) Then
50800     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
50810       .AutosaveStartStandardProgram = CLng(tStr)
50820      Else
50830       If UseStandard Then
50840        .AutosaveStartStandardProgram = 0
50850       End If
50860     End If
50870    Else
50880     If UseStandard Then
50890      .AutosaveStartStandardProgram = 0
50900     End If
50910   End If
50920   tStr = hOpt.Retrieve("BMPColorscount")
50930   If IsNumeric(tStr) Then
50940     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
50950       .BMPColorscount = CLng(tStr)
50960      Else
50970       If UseStandard Then
50980        .BMPColorscount = 1
50990       End If
51000     End If
51010    Else
51020     If UseStandard Then
51030      .BMPColorscount = 1
51040     End If
51050   End If
51060   tStr = hOpt.Retrieve("BMPResolution")
51070   If IsNumeric(tStr) Then
51080     If CLng(tStr) >= 1 Then
51090       .BMPResolution = CLng(tStr)
51100      Else
51110       If UseStandard Then
51120        .BMPResolution = 150
51130       End If
51140     End If
51150    Else
51160     If UseStandard Then
51170      .BMPResolution = 150
51180     End If
51190   End If
51200   tStr = hOpt.Retrieve("ClientComputerResolveIPAddress")
51210   If IsNumeric(tStr) Then
51220     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51230       .ClientComputerResolveIPAddress = CLng(tStr)
51240      Else
51250       If UseStandard Then
51260        .ClientComputerResolveIPAddress = 0
51270       End If
51280     End If
51290    Else
51300     If UseStandard Then
51310      .ClientComputerResolveIPAddress = 0
51320     End If
51330   End If
51340   tStr = hOpt.Retrieve("Counter")
51350   If IsNumeric(tStr) Then
51360     If CCur(tStr) >= 0 And CCur(tStr) <= 922337203685477# Then
51370       .Counter = CCur(tStr)
51380      Else
51390       If UseStandard Then
51400        .Counter = 0
51410       End If
51420     End If
51430    Else
51440     If UseStandard Then
51450      .Counter = 0
51460     End If
51470   End If
51480   tStr = hOpt.Retrieve("DeviceHeightPoints")
51490   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
51500     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 1 Then
51510       .DeviceHeightPoints = CDbl(Replace$(tStr, ".", GetDecimalChar))
51520      Else
51530       If UseStandard Then
51540        .DeviceHeightPoints = Replace$("842", ".", GetDecimalChar)
51550       End If
51560     End If
51570    Else
51580     If UseStandard Then
51590      .DeviceHeightPoints = Replace$("842", ".", GetDecimalChar)
51600     End If
51610   End If
51620   tStr = hOpt.Retrieve("DeviceWidthPoints")
51630   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
51640     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 1 Then
51650       .DeviceWidthPoints = CDbl(Replace$(tStr, ".", GetDecimalChar))
51660      Else
51670       If UseStandard Then
51680        .DeviceWidthPoints = Replace$("595", ".", GetDecimalChar)
51690       End If
51700     End If
51710    Else
51720     If UseStandard Then
51730      .DeviceWidthPoints = Replace$("595", ".", GetDecimalChar)
51740     End If
51750   End If
51760   tStr = hOpt.Retrieve("DirectoryGhostscriptBinaries")
51770   If LenB(Trim$(tStr)) > 0 Then
51780     .DirectoryGhostscriptBinaries = CompletePath(tStr)
51790    Else
51800     If UseStandard Then
51810      tStr = GetPDFCreatorApplicationPath
51820      .DirectoryGhostscriptBinaries = CompletePath(tStr)
51830     End If
51840   End If
51850   tStr = hOpt.Retrieve("DirectoryGhostscriptFonts")
51860   If LenB(Trim$(tStr)) > 0 Then
51870     .DirectoryGhostscriptFonts = CompletePath(tStr)
51880    Else
51890     If UseStandard Then
51900      tStr = GetPDFCreatorApplicationPath & "fonts"
51910      .DirectoryGhostscriptFonts = CompletePath(tStr)
51920     End If
51930   End If
51940   tStr = hOpt.Retrieve("DirectoryGhostscriptLibraries")
51950   If LenB(Trim$(tStr)) > 0 Then
51960     .DirectoryGhostscriptLibraries = CompletePath(tStr)
51970    Else
51980     If UseStandard Then
51990      tStr = GetPDFCreatorApplicationPath & "lib"
52000      .DirectoryGhostscriptLibraries = CompletePath(tStr)
52010     End If
52020   End If
52030   tStr = hOpt.Retrieve("DirectoryGhostscriptResource")
52040   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
52050     .DirectoryGhostscriptResource = ""
52060    Else
52070     If LenB(tStr) > 0 Then
52080      .DirectoryGhostscriptResource = tStr
52090     End If
52100   End If
52110   tStr = hOpt.Retrieve("DisableEmail")
52120   If IsNumeric(tStr) Then
52130     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52140       .DisableEmail = CLng(tStr)
52150      Else
52160       If UseStandard Then
52170        .DisableEmail = 0
52180       End If
52190     End If
52200    Else
52210     If UseStandard Then
52220      .DisableEmail = 0
52230     End If
52240   End If
52250   tStr = hOpt.Retrieve("DontUseDocumentSettings")
52260   If IsNumeric(tStr) Then
52270     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52280       .DontUseDocumentSettings = CLng(tStr)
52290      Else
52300       If UseStandard Then
52310        .DontUseDocumentSettings = 0
52320       End If
52330     End If
52340    Else
52350     If UseStandard Then
52360      .DontUseDocumentSettings = 0
52370     End If
52380   End If
52390   tStr = hOpt.Retrieve("EPSLanguageLevel")
52400   If IsNumeric(tStr) Then
52410     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
52420       .EPSLanguageLevel = CLng(tStr)
52430      Else
52440       If UseStandard Then
52450        .EPSLanguageLevel = 2
52460       End If
52470     End If
52480    Else
52490     If UseStandard Then
52500      .EPSLanguageLevel = 2
52510     End If
52520   End If
52530   tStr = hOpt.Retrieve("FilenameSubstitutions")
52540   If LenB(tStr) = 0 And LenB("Microsoft Word - \.docx\.doc\Microsoft Excel - \.xlsx\.xls\Microsoft PowerPoint - \.pptx\.ppt") > 0 And UseStandard Then
52550     .FilenameSubstitutions = "Microsoft Word - \.docx\.doc\Microsoft Excel - \.xlsx\.xls\Microsoft PowerPoint - \.pptx\.ppt"
52560    Else
52570     If LenB(tStr) > 0 Then
52580      .FilenameSubstitutions = tStr
52590     End If
52600   End If
52610   tStr = hOpt.Retrieve("FilenameSubstitutionsOnlyInTitle")
52620   If IsNumeric(tStr) Then
52630     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52640       .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
52650      Else
52660       If UseStandard Then
52670        .FilenameSubstitutionsOnlyInTitle = 1
52680       End If
52690     End If
52700    Else
52710     If UseStandard Then
52720      .FilenameSubstitutionsOnlyInTitle = 1
52730     End If
52740   End If
52750   tStr = hOpt.Retrieve("JPEGColorscount")
52760   If IsNumeric(tStr) Then
52770     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
52780       .JPEGColorscount = CLng(tStr)
52790      Else
52800       If UseStandard Then
52810        .JPEGColorscount = 0
52820       End If
52830     End If
52840    Else
52850     If UseStandard Then
52860      .JPEGColorscount = 0
52870     End If
52880   End If
52890   tStr = hOpt.Retrieve("JPEGQuality")
52900   If IsNumeric(tStr) Then
52910     If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
52920       .JPEGQuality = CLng(tStr)
52930      Else
52940       If UseStandard Then
52950        .JPEGQuality = 75
52960       End If
52970     End If
52980    Else
52990     If UseStandard Then
53000      .JPEGQuality = 75
53010     End If
53020   End If
53030   tStr = hOpt.Retrieve("JPEGResolution")
53040   If IsNumeric(tStr) Then
53050     If CLng(tStr) >= 1 Then
53060       .JPEGResolution = CLng(tStr)
53070      Else
53080       If UseStandard Then
53090        .JPEGResolution = 150
53100       End If
53110     End If
53120    Else
53130     If UseStandard Then
53140      .JPEGResolution = 150
53150     End If
53160   End If
53170   tStr = hOpt.Retrieve("Language")
53180   If LenB(tStr) = 0 And LenB("english") > 0 And UseStandard Then
53190     .Language = "english"
53200    Else
53210     If LenB(tStr) > 0 Then
53220      .Language = tStr
53230     End If
53240   End If
53250   tStr = hOpt.Retrieve("LastSaveDirectory")
53260   If LenB(Trim$(tStr)) > 0 Then
53270     .LastSaveDirectory = CompletePath(tStr)
53280    Else
53290     If UseStandard Then
53300      If InstalledAsServer Then
53310        .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
53320       Else
53330        .LastSaveDirectory = "<MyFiles>"
53340      End If
53350     End If
53360   End If
53370   tStr = hOpt.Retrieve("LastUpdateCheck")
53380   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
53390     .LastUpdateCheck = ""
53400    Else
53410     If LenB(tStr) > 0 Then
53420      .LastUpdateCheck = tStr
53430     End If
53440   End If
53450   tStr = hOpt.Retrieve("Logging")
53460   If IsNumeric(tStr) Then
53470     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53480       .Logging = CLng(tStr)
53490      Else
53500       If UseStandard Then
53510        .Logging = 0
53520       End If
53530     End If
53540    Else
53550     If UseStandard Then
53560      .Logging = 0
53570     End If
53580   End If
53590   tStr = hOpt.Retrieve("LogLines")
53600   If IsNumeric(tStr) Then
53610     If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
53620       .LogLines = CLng(tStr)
53630      Else
53640       If UseStandard Then
53650        .LogLines = 100
53660       End If
53670     End If
53680    Else
53690     If UseStandard Then
53700      .LogLines = 100
53710     End If
53720   End If
53730   tStr = hOpt.Retrieve("NoConfirmMessageSwitchingDefaultprinter")
53740   If IsNumeric(tStr) Then
53750     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53760       .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
53770      Else
53780       If UseStandard Then
53790        .NoConfirmMessageSwitchingDefaultprinter = 0
53800       End If
53810     End If
53820    Else
53830     If UseStandard Then
53840      .NoConfirmMessageSwitchingDefaultprinter = 0
53850     End If
53860   End If
53870   tStr = hOpt.Retrieve("NoProcessingAtStartup")
53880   If IsNumeric(tStr) Then
53890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53900       .NoProcessingAtStartup = CLng(tStr)
53910      Else
53920       If UseStandard Then
53930        .NoProcessingAtStartup = 0
53940       End If
53950     End If
53960    Else
53970     If UseStandard Then
53980      .NoProcessingAtStartup = 0
53990     End If
54000   End If
54010   tStr = hOpt.Retrieve("NoPSCheck")
54020   If IsNumeric(tStr) Then
54030     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54040       .NoPSCheck = CLng(tStr)
54050      Else
54060       If UseStandard Then
54070        .NoPSCheck = 0
54080       End If
54090     End If
54100    Else
54110     If UseStandard Then
54120      .NoPSCheck = 0
54130     End If
54140   End If
54150   tStr = hOpt.Retrieve("OnePagePerFile")
54160   If IsNumeric(tStr) Then
54170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54180       .OnePagePerFile = CLng(tStr)
54190      Else
54200       If UseStandard Then
54210        .OnePagePerFile = 0
54220       End If
54230     End If
54240    Else
54250     If UseStandard Then
54260      .OnePagePerFile = 0
54270     End If
54280   End If
54290   tStr = hOpt.Retrieve("OptionsDesign")
54300   If IsNumeric(tStr) Then
54310     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
54320       .OptionsDesign = CLng(tStr)
54330      Else
54340       If UseStandard Then
54350        .OptionsDesign = 0
54360       End If
54370     End If
54380    Else
54390     If UseStandard Then
54400      .OptionsDesign = 0
54410     End If
54420   End If
54430   tStr = hOpt.Retrieve("OptionsEnabled")
54440   If IsNumeric(tStr) Then
54450     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54460       .OptionsEnabled = CLng(tStr)
54470      Else
54480       If UseStandard Then
54490        .OptionsEnabled = 1
54500       End If
54510     End If
54520    Else
54530     If UseStandard Then
54540      .OptionsEnabled = 1
54550     End If
54560   End If
54570   tStr = hOpt.Retrieve("OptionsVisible")
54580   If IsNumeric(tStr) Then
54590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54600       .OptionsVisible = CLng(tStr)
54610      Else
54620       If UseStandard Then
54630        .OptionsVisible = 1
54640       End If
54650     End If
54660    Else
54670     If UseStandard Then
54680      .OptionsVisible = 1
54690     End If
54700   End If
54710   tStr = hOpt.Retrieve("Papersize")
54720   If LenB(tStr) = 0 And LenB("a4") > 0 And UseStandard Then
54730     .Papersize = "a4"
54740    Else
54750     If LenB(tStr) > 0 Then
54760      .Papersize = tStr
54770     End If
54780   End If
54790   tStr = hOpt.Retrieve("PCLColorsCount")
54800   If IsNumeric(tStr) Then
54810     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
54820       .PCLColorsCount = CLng(tStr)
54830      Else
54840       If UseStandard Then
54850        .PCLColorsCount = 0
54860       End If
54870     End If
54880    Else
54890     If UseStandard Then
54900      .PCLColorsCount = 0
54910     End If
54920   End If
54930   tStr = hOpt.Retrieve("PCLResolution")
54940   If IsNumeric(tStr) Then
54950     If CLng(tStr) >= 1 Then
54960       .PCLResolution = CLng(tStr)
54970      Else
54980       If UseStandard Then
54990        .PCLResolution = 150
55000       End If
55010     End If
55020    Else
55030     If UseStandard Then
55040      .PCLResolution = 150
55050     End If
55060   End If
55070   tStr = hOpt.Retrieve("PCXColorscount")
55080   If IsNumeric(tStr) Then
55090     If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
55100       .PCXColorscount = CLng(tStr)
55110      Else
55120       If UseStandard Then
55130        .PCXColorscount = 0
55140       End If
55150     End If
55160    Else
55170     If UseStandard Then
55180      .PCXColorscount = 0
55190     End If
55200   End If
55210   tStr = hOpt.Retrieve("PCXResolution")
55220   If IsNumeric(tStr) Then
55230     If CLng(tStr) >= 1 Then
55240       .PCXResolution = CLng(tStr)
55250      Else
55260       If UseStandard Then
55270        .PCXResolution = 150
55280       End If
55290     End If
55300    Else
55310     If UseStandard Then
55320      .PCXResolution = 150
55330     End If
55340   End If
55350   tStr = hOpt.Retrieve("PDFAllowAssembly")
55360   If IsNumeric(tStr) Then
55370     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55380       .PDFAllowAssembly = CLng(tStr)
55390      Else
55400       If UseStandard Then
55410        .PDFAllowAssembly = 0
55420       End If
55430     End If
55440    Else
55450     If UseStandard Then
55460      .PDFAllowAssembly = 0
55470     End If
55480   End If
55490   tStr = hOpt.Retrieve("PDFAllowDegradedPrinting")
55500   If IsNumeric(tStr) Then
55510     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55520       .PDFAllowDegradedPrinting = CLng(tStr)
55530      Else
55540       If UseStandard Then
55550        .PDFAllowDegradedPrinting = 0
55560       End If
55570     End If
55580    Else
55590     If UseStandard Then
55600      .PDFAllowDegradedPrinting = 0
55610     End If
55620   End If
55630   tStr = hOpt.Retrieve("PDFAllowFillIn")
55640   If IsNumeric(tStr) Then
55650     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55660       .PDFAllowFillIn = CLng(tStr)
55670      Else
55680       If UseStandard Then
55690        .PDFAllowFillIn = 0
55700       End If
55710     End If
55720    Else
55730     If UseStandard Then
55740      .PDFAllowFillIn = 0
55750     End If
55760   End If
55770   tStr = hOpt.Retrieve("PDFAllowScreenReaders")
55780   If IsNumeric(tStr) Then
55790     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55800       .PDFAllowScreenReaders = CLng(tStr)
55810      Else
55820       If UseStandard Then
55830        .PDFAllowScreenReaders = 0
55840       End If
55850     End If
55860    Else
55870     If UseStandard Then
55880      .PDFAllowScreenReaders = 0
55890     End If
55900   End If
55910   tStr = hOpt.Retrieve("PDFColorsCMYKToRGB")
55920   If IsNumeric(tStr) Then
55930     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55940       .PDFColorsCMYKToRGB = CLng(tStr)
55950      Else
55960       If UseStandard Then
55970        .PDFColorsCMYKToRGB = 0
55980       End If
55990     End If
56000    Else
56010     If UseStandard Then
56020      .PDFColorsCMYKToRGB = 0
56030     End If
56040   End If
56050   tStr = hOpt.Retrieve("PDFColorsColorModel")
56060   If IsNumeric(tStr) Then
56070     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
56080       .PDFColorsColorModel = CLng(tStr)
56090      Else
56100       If UseStandard Then
56110        .PDFColorsColorModel = 1
56120       End If
56130     End If
56140    Else
56150     If UseStandard Then
56160      .PDFColorsColorModel = 1
56170     End If
56180   End If
56190   tStr = hOpt.Retrieve("PDFColorsPreserveHalftone")
56200   If IsNumeric(tStr) Then
56210     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56220       .PDFColorsPreserveHalftone = CLng(tStr)
56230      Else
56240       If UseStandard Then
56250        .PDFColorsPreserveHalftone = 0
56260       End If
56270     End If
56280    Else
56290     If UseStandard Then
56300      .PDFColorsPreserveHalftone = 0
56310     End If
56320   End If
56330   tStr = hOpt.Retrieve("PDFColorsPreserveOverprint")
56340   If IsNumeric(tStr) Then
56350     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56360       .PDFColorsPreserveOverprint = CLng(tStr)
56370      Else
56380       If UseStandard Then
56390        .PDFColorsPreserveOverprint = 1
56400       End If
56410     End If
56420    Else
56430     If UseStandard Then
56440      .PDFColorsPreserveOverprint = 1
56450     End If
56460   End If
56470   tStr = hOpt.Retrieve("PDFColorsPreserveTransfer")
56480   If IsNumeric(tStr) Then
56490     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56500       .PDFColorsPreserveTransfer = CLng(tStr)
56510      Else
56520       If UseStandard Then
56530        .PDFColorsPreserveTransfer = 1
56540       End If
56550     End If
56560    Else
56570     If UseStandard Then
56580      .PDFColorsPreserveTransfer = 1
56590     End If
56600   End If
56610   tStr = hOpt.Retrieve("PDFCompressionColorCompression")
56620   If IsNumeric(tStr) Then
56630     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56640       .PDFCompressionColorCompression = CLng(tStr)
56650      Else
56660       If UseStandard Then
56670        .PDFCompressionColorCompression = 1
56680       End If
56690     End If
56700    Else
56710     If UseStandard Then
56720      .PDFCompressionColorCompression = 1
56730     End If
56740   End If
56750   tStr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
56760   If IsNumeric(tStr) Then
56770     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
56780       .PDFCompressionColorCompressionChoice = CLng(tStr)
56790      Else
56800       If UseStandard Then
56810        .PDFCompressionColorCompressionChoice = 0
56820       End If
56830     End If
56840    Else
56850     If UseStandard Then
56860      .PDFCompressionColorCompressionChoice = 0
56870     End If
56880   End If
56890   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGHighFactor")
56900   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56910     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56920       .PDFCompressionColorCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56930      Else
56940       If UseStandard Then
56950        .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56960       End If
56970     End If
56980    Else
56990     If UseStandard Then
57000      .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
57010     End If
57020   End If
57030   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGLowFactor")
57040   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57050     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57060       .PDFCompressionColorCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57070      Else
57080       If UseStandard Then
57090        .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
57100       End If
57110     End If
57120    Else
57130     If UseStandard Then
57140      .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
57150     End If
57160   End If
57170   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMaximumFactor")
57180   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57190     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57200       .PDFCompressionColorCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57210      Else
57220       If UseStandard Then
57230        .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57240       End If
57250     End If
57260    Else
57270     If UseStandard Then
57280      .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57290     End If
57300   End If
57310   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMediumFactor")
57320   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57330     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57340       .PDFCompressionColorCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57350      Else
57360       If UseStandard Then
57370        .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57380       End If
57390     End If
57400    Else
57410     If UseStandard Then
57420      .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57430     End If
57440   End If
57450   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMinimumFactor")
57460   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57470     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57480       .PDFCompressionColorCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57490      Else
57500       If UseStandard Then
57510        .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57520       End If
57530     End If
57540    Else
57550     If UseStandard Then
57560      .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57570     End If
57580   End If
57590   tStr = hOpt.Retrieve("PDFCompressionColorResample")
57600   If IsNumeric(tStr) Then
57610     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57620       .PDFCompressionColorResample = CLng(tStr)
57630      Else
57640       If UseStandard Then
57650        .PDFCompressionColorResample = 0
57660       End If
57670     End If
57680    Else
57690     If UseStandard Then
57700      .PDFCompressionColorResample = 0
57710     End If
57720   End If
57730   tStr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
57740   If IsNumeric(tStr) Then
57750     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
57760       .PDFCompressionColorResampleChoice = CLng(tStr)
57770      Else
57780       If UseStandard Then
57790        .PDFCompressionColorResampleChoice = 0
57800       End If
57810     End If
57820    Else
57830     If UseStandard Then
57840      .PDFCompressionColorResampleChoice = 0
57850     End If
57860   End If
57870   tStr = hOpt.Retrieve("PDFCompressionColorResolution")
57880   If IsNumeric(tStr) Then
57890     If CLng(tStr) >= 0 Then
57900       .PDFCompressionColorResolution = CLng(tStr)
57910      Else
57920       If UseStandard Then
57930        .PDFCompressionColorResolution = 300
57940       End If
57950     End If
57960    Else
57970     If UseStandard Then
57980      .PDFCompressionColorResolution = 300
57990     End If
58000   End If
58010   tStr = hOpt.Retrieve("PDFCompressionGreyCompression")
58020   If IsNumeric(tStr) Then
58030     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58040       .PDFCompressionGreyCompression = CLng(tStr)
58050      Else
58060       If UseStandard Then
58070        .PDFCompressionGreyCompression = 1
58080       End If
58090     End If
58100    Else
58110     If UseStandard Then
58120      .PDFCompressionGreyCompression = 1
58130     End If
58140   End If
58150   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
58160   If IsNumeric(tStr) Then
58170     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
58180       .PDFCompressionGreyCompressionChoice = CLng(tStr)
58190      Else
58200       If UseStandard Then
58210        .PDFCompressionGreyCompressionChoice = 0
58220       End If
58230     End If
58240    Else
58250     If UseStandard Then
58260      .PDFCompressionGreyCompressionChoice = 0
58270     End If
58280   End If
58290   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGHighFactor")
58300   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58310     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58320       .PDFCompressionGreyCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58330      Else
58340       If UseStandard Then
58350        .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
58360       End If
58370     End If
58380    Else
58390     If UseStandard Then
58400      .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
58410     End If
58420   End If
58430   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGLowFactor")
58440   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58450     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58460       .PDFCompressionGreyCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58470      Else
58480       If UseStandard Then
58490        .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58500       End If
58510     End If
58520    Else
58530     If UseStandard Then
58540      .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58550     End If
58560   End If
58570   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMaximumFactor")
58580   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58590     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58600       .PDFCompressionGreyCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58610      Else
58620       If UseStandard Then
58630        .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
58640       End If
58650     End If
58660    Else
58670     If UseStandard Then
58680      .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
58690     End If
58700   End If
58710   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMediumFactor")
58720   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58730     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58740       .PDFCompressionGreyCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58750      Else
58760       If UseStandard Then
58770        .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58780       End If
58790     End If
58800    Else
58810     If UseStandard Then
58820      .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58830     End If
58840   End If
58850   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMinimumFactor")
58860   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58870     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58880       .PDFCompressionGreyCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58890      Else
58900       If UseStandard Then
58910        .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58920       End If
58930     End If
58940    Else
58950     If UseStandard Then
58960      .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58970     End If
58980   End If
58990   tStr = hOpt.Retrieve("PDFCompressionGreyResample")
59000   If IsNumeric(tStr) Then
59010     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59020       .PDFCompressionGreyResample = CLng(tStr)
59030      Else
59040       If UseStandard Then
59050        .PDFCompressionGreyResample = 0
59060       End If
59070     End If
59080    Else
59090     If UseStandard Then
59100      .PDFCompressionGreyResample = 0
59110     End If
59120   End If
59130   tStr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
59140   If IsNumeric(tStr) Then
59150     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
59160       .PDFCompressionGreyResampleChoice = CLng(tStr)
59170      Else
59180       If UseStandard Then
59190        .PDFCompressionGreyResampleChoice = 0
59200       End If
59210     End If
59220    Else
59230     If UseStandard Then
59240      .PDFCompressionGreyResampleChoice = 0
59250     End If
59260   End If
59270   tStr = hOpt.Retrieve("PDFCompressionGreyResolution")
59280   If IsNumeric(tStr) Then
59290     If CLng(tStr) >= 0 Then
59300       .PDFCompressionGreyResolution = CLng(tStr)
59310      Else
59320       If UseStandard Then
59330        .PDFCompressionGreyResolution = 300
59340       End If
59350     End If
59360    Else
59370     If UseStandard Then
59380      .PDFCompressionGreyResolution = 300
59390     End If
59400   End If
59410   tStr = hOpt.Retrieve("PDFCompressionMonoCompression")
59420   If IsNumeric(tStr) Then
59430     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59440       .PDFCompressionMonoCompression = CLng(tStr)
59450      Else
59460       If UseStandard Then
59470        .PDFCompressionMonoCompression = 1
59480       End If
59490     End If
59500    Else
59510     If UseStandard Then
59520      .PDFCompressionMonoCompression = 1
59530     End If
59540   End If
59550   tStr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
59560   If IsNumeric(tStr) Then
59570     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
59580       .PDFCompressionMonoCompressionChoice = CLng(tStr)
59590      Else
59600       If UseStandard Then
59610        .PDFCompressionMonoCompressionChoice = 0
59620       End If
59630     End If
59640    Else
59650     If UseStandard Then
59660      .PDFCompressionMonoCompressionChoice = 0
59670     End If
59680   End If
59690   tStr = hOpt.Retrieve("PDFCompressionMonoResample")
59700   If IsNumeric(tStr) Then
59710     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59720       .PDFCompressionMonoResample = CLng(tStr)
59730      Else
59740       If UseStandard Then
59750        .PDFCompressionMonoResample = 0
59760       End If
59770     End If
59780    Else
59790     If UseStandard Then
59800      .PDFCompressionMonoResample = 0
59810     End If
59820   End If
59830   tStr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
59840   If IsNumeric(tStr) Then
59850     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
59860       .PDFCompressionMonoResampleChoice = CLng(tStr)
59870      Else
59880       If UseStandard Then
59890        .PDFCompressionMonoResampleChoice = 0
59900       End If
59910     End If
59920    Else
59930     If UseStandard Then
59940      .PDFCompressionMonoResampleChoice = 0
59950     End If
59960   End If
59970   tStr = hOpt.Retrieve("PDFCompressionMonoResolution")
59980   If IsNumeric(tStr) Then
59990     If CLng(tStr) >= 0 Then
60000       .PDFCompressionMonoResolution = CLng(tStr)
60010      Else
60020       If UseStandard Then
60030        .PDFCompressionMonoResolution = 1200
60040       End If
60050     End If
60060    Else
60070     If UseStandard Then
60080      .PDFCompressionMonoResolution = 1200
60090     End If
60100   End If
60110   tStr = hOpt.Retrieve("PDFCompressionTextCompression")
60120   If IsNumeric(tStr) Then
60130     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60140       .PDFCompressionTextCompression = CLng(tStr)
60150      Else
60160       If UseStandard Then
60170        .PDFCompressionTextCompression = 1
60180       End If
60190     End If
60200    Else
60210     If UseStandard Then
60220      .PDFCompressionTextCompression = 1
60230     End If
60240   End If
60250   tStr = hOpt.Retrieve("PDFDisallowCopy")
60260   If IsNumeric(tStr) Then
60270     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60280       .PDFDisallowCopy = CLng(tStr)
60290      Else
60300       If UseStandard Then
60310        .PDFDisallowCopy = 1
60320       End If
60330     End If
60340    Else
60350     If UseStandard Then
60360      .PDFDisallowCopy = 1
60370     End If
60380   End If
60390   tStr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
60400   If IsNumeric(tStr) Then
60410     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60420       .PDFDisallowModifyAnnotations = CLng(tStr)
60430      Else
60440       If UseStandard Then
60450        .PDFDisallowModifyAnnotations = 0
60460       End If
60470     End If
60480    Else
60490     If UseStandard Then
60500      .PDFDisallowModifyAnnotations = 0
60510     End If
60520   End If
60530   tStr = hOpt.Retrieve("PDFDisallowModifyContents")
60540   If IsNumeric(tStr) Then
60550     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60560       .PDFDisallowModifyContents = CLng(tStr)
60570      Else
60580       If UseStandard Then
60590        .PDFDisallowModifyContents = 0
60600       End If
60610     End If
60620    Else
60630     If UseStandard Then
60640      .PDFDisallowModifyContents = 0
60650     End If
60660   End If
60670   tStr = hOpt.Retrieve("PDFDisallowPrinting")
60680   If IsNumeric(tStr) Then
60690     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60700       .PDFDisallowPrinting = CLng(tStr)
60710      Else
60720       If UseStandard Then
60730        .PDFDisallowPrinting = 0
60740       End If
60750     End If
60760    Else
60770     If UseStandard Then
60780      .PDFDisallowPrinting = 0
60790     End If
60800   End If
60810   tStr = hOpt.Retrieve("PDFEncryptor")
60820   If IsNumeric(tStr) Then
60830     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
60840       .PDFEncryptor = CLng(tStr)
60850      Else
60860       If UseStandard Then
60870        .PDFEncryptor = 0
60880       End If
60890     End If
60900    Else
60910     If UseStandard Then
60920      .PDFEncryptor = 0
60930     End If
60940   End If
60950   tStr = hOpt.Retrieve("PDFFontsEmbedAll")
60960   If IsNumeric(tStr) Then
60970     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60980       .PDFFontsEmbedAll = CLng(tStr)
60990      Else
61000       If UseStandard Then
61010        .PDFFontsEmbedAll = 1
61020       End If
61030     End If
61040    Else
61050     If UseStandard Then
61060      .PDFFontsEmbedAll = 1
61070     End If
61080   End If
61090   tStr = hOpt.Retrieve("PDFFontsSubSetFonts")
61100   If IsNumeric(tStr) Then
61110     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61120       .PDFFontsSubSetFonts = CLng(tStr)
61130      Else
61140       If UseStandard Then
61150        .PDFFontsSubSetFonts = 1
61160       End If
61170     End If
61180    Else
61190     If UseStandard Then
61200      .PDFFontsSubSetFonts = 1
61210     End If
61220   End If
61230   tStr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
61240   If IsNumeric(tStr) Then
61250     If CLng(tStr) >= 0 Then
61260       .PDFFontsSubSetFontsPercent = CLng(tStr)
61270      Else
61280       If UseStandard Then
61290        .PDFFontsSubSetFontsPercent = 100
61300       End If
61310     End If
61320    Else
61330     If UseStandard Then
61340      .PDFFontsSubSetFontsPercent = 100
61350     End If
61360   End If
61370   tStr = hOpt.Retrieve("PDFGeneralASCII85")
61380   If IsNumeric(tStr) Then
61390     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61400       .PDFGeneralASCII85 = CLng(tStr)
61410      Else
61420       If UseStandard Then
61430        .PDFGeneralASCII85 = 0
61440       End If
61450     End If
61460    Else
61470     If UseStandard Then
61480      .PDFGeneralASCII85 = 0
61490     End If
61500   End If
61510   tStr = hOpt.Retrieve("PDFGeneralAutorotate")
61520   If IsNumeric(tStr) Then
61530     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
61540       .PDFGeneralAutorotate = CLng(tStr)
61550      Else
61560       If UseStandard Then
61570        .PDFGeneralAutorotate = 2
61580       End If
61590     End If
61600    Else
61610     If UseStandard Then
61620      .PDFGeneralAutorotate = 2
61630     End If
61640   End If
61650   tStr = hOpt.Retrieve("PDFGeneralCompatibility")
61660   If IsNumeric(tStr) Then
61670     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
61680       .PDFGeneralCompatibility = CLng(tStr)
61690      Else
61700       If UseStandard Then
61710        .PDFGeneralCompatibility = 2
61720       End If
61730     End If
61740    Else
61750     If UseStandard Then
61760      .PDFGeneralCompatibility = 2
61770     End If
61780   End If
61790   tStr = hOpt.Retrieve("PDFGeneralDefault")
61800   If IsNumeric(tStr) Then
61810     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
61820       .PDFGeneralDefault = CLng(tStr)
61830      Else
61840       If UseStandard Then
61850        .PDFGeneralDefault = 0
61860       End If
61870     End If
61880    Else
61890     If UseStandard Then
61900      .PDFGeneralDefault = 0
61910     End If
61920   End If
61930   tStr = hOpt.Retrieve("PDFGeneralOverprint")
61940   If IsNumeric(tStr) Then
61950     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
61960       .PDFGeneralOverprint = CLng(tStr)
61970      Else
61980       If UseStandard Then
61990        .PDFGeneralOverprint = 0
62000       End If
62010     End If
62020    Else
62030     If UseStandard Then
62040      .PDFGeneralOverprint = 0
62050     End If
62060   End If
62070   tStr = hOpt.Retrieve("PDFGeneralResolution")
62080   If IsNumeric(tStr) Then
62090     If CLng(tStr) >= 0 Then
62100       .PDFGeneralResolution = CLng(tStr)
62110      Else
62120       If UseStandard Then
62130        .PDFGeneralResolution = 600
62140       End If
62150     End If
62160    Else
62170     If UseStandard Then
62180      .PDFGeneralResolution = 600
62190     End If
62200   End If
62210   tStr = hOpt.Retrieve("PDFHighEncryption")
62220   If IsNumeric(tStr) Then
62230     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62240       .PDFHighEncryption = CLng(tStr)
62250      Else
62260       If UseStandard Then
62270        .PDFHighEncryption = 0
62280       End If
62290     End If
62300    Else
62310     If UseStandard Then
62320      .PDFHighEncryption = 0
62330     End If
62340   End If
62350   tStr = hOpt.Retrieve("PDFLowEncryption")
62360   If IsNumeric(tStr) Then
62370     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62380       .PDFLowEncryption = CLng(tStr)
62390      Else
62400       If UseStandard Then
62410        .PDFLowEncryption = 1
62420       End If
62430     End If
62440    Else
62450     If UseStandard Then
62460      .PDFLowEncryption = 1
62470     End If
62480   End If
62490   tStr = hOpt.Retrieve("PDFOptimize")
62500   If IsNumeric(tStr) Then
62510     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62520       .PDFOptimize = CLng(tStr)
62530      Else
62540       If UseStandard Then
62550        .PDFOptimize = 0
62560       End If
62570     End If
62580    Else
62590     If UseStandard Then
62600      .PDFOptimize = 0
62610     End If
62620   End If
62630   tStr = hOpt.Retrieve("PDFOwnerPass")
62640   If IsNumeric(tStr) Then
62650     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62660       .PDFOwnerPass = CLng(tStr)
62670      Else
62680       If UseStandard Then
62690        .PDFOwnerPass = 0
62700       End If
62710     End If
62720    Else
62730     If UseStandard Then
62740      .PDFOwnerPass = 0
62750     End If
62760   End If
62770   tStr = hOpt.Retrieve("PDFOwnerPasswordString")
62780   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62790     .PDFOwnerPasswordString = ""
62800    Else
62810     If LenB(tStr) > 0 Then
62820      .PDFOwnerPasswordString = tStr
62830     End If
62840   End If
62850   tStr = hOpt.Retrieve("PDFSigningMultiSignature")
62860   If IsNumeric(tStr) Then
62870     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62880       .PDFSigningMultiSignature = CLng(tStr)
62890      Else
62900       If UseStandard Then
62910        .PDFSigningMultiSignature = 0
62920       End If
62930     End If
62940    Else
62950     If UseStandard Then
62960      .PDFSigningMultiSignature = 0
62970     End If
62980   End If
62990   tStr = hOpt.Retrieve("PDFSigningPFXFile")
63000   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63010     .PDFSigningPFXFile = ""
63020    Else
63030     If LenB(tStr) > 0 Then
63040      .PDFSigningPFXFile = tStr
63050     End If
63060   End If
63070   tStr = hOpt.Retrieve("PDFSigningPFXFilePassword")
63080   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63090     .PDFSigningPFXFilePassword = ""
63100    Else
63110     If LenB(tStr) > 0 Then
63120      .PDFSigningPFXFilePassword = tStr
63130     End If
63140   End If
63150   tStr = hOpt.Retrieve("PDFSigningSignatureContact")
63160   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63170     .PDFSigningSignatureContact = ""
63180    Else
63190     If LenB(tStr) > 0 Then
63200      .PDFSigningSignatureContact = tStr
63210     End If
63220   End If
63230   tStr = hOpt.Retrieve("PDFSigningSignatureLeftX")
63240   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63250     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63260       .PDFSigningSignatureLeftX = CDbl(Replace$(tStr, ".", GetDecimalChar))
63270      Else
63280       If UseStandard Then
63290        .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
63300       End If
63310     End If
63320    Else
63330     If UseStandard Then
63340      .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
63350     End If
63360   End If
63370   tStr = hOpt.Retrieve("PDFSigningSignatureLeftY")
63380   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63390     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63400       .PDFSigningSignatureLeftY = CDbl(Replace$(tStr, ".", GetDecimalChar))
63410      Else
63420       If UseStandard Then
63430        .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
63440       End If
63450     End If
63460    Else
63470     If UseStandard Then
63480      .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
63490     End If
63500   End If
63510   tStr = hOpt.Retrieve("PDFSigningSignatureLocation")
63520   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63530     .PDFSigningSignatureLocation = ""
63540    Else
63550     If LenB(tStr) > 0 Then
63560      .PDFSigningSignatureLocation = tStr
63570     End If
63580   End If
63590   tStr = hOpt.Retrieve("PDFSigningSignatureReason")
63600   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63610     .PDFSigningSignatureReason = ""
63620    Else
63630     If LenB(tStr) > 0 Then
63640      .PDFSigningSignatureReason = tStr
63650     End If
63660   End If
63670   tStr = hOpt.Retrieve("PDFSigningSignatureRightX")
63680   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63690     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63700       .PDFSigningSignatureRightX = CDbl(Replace$(tStr, ".", GetDecimalChar))
63710      Else
63720       If UseStandard Then
63730        .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
63740       End If
63750     End If
63760    Else
63770     If UseStandard Then
63780      .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
63790     End If
63800   End If
63810   tStr = hOpt.Retrieve("PDFSigningSignatureRightY")
63820   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63830     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63840       .PDFSigningSignatureRightY = CDbl(Replace$(tStr, ".", GetDecimalChar))
63850      Else
63860       If UseStandard Then
63870        .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
63880       End If
63890     End If
63900    Else
63910     If UseStandard Then
63920      .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
63930     End If
63940   End If
63950   tStr = hOpt.Retrieve("PDFSigningSignatureVisible")
63960   If IsNumeric(tStr) Then
63970     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63980       .PDFSigningSignatureVisible = CLng(tStr)
63990      Else
64000       If UseStandard Then
64010        .PDFSigningSignatureVisible = 0
64020       End If
64030     End If
64040    Else
64050     If UseStandard Then
64060      .PDFSigningSignatureVisible = 0
64070     End If
64080   End If
64090   tStr = hOpt.Retrieve("PDFSigningSignPDF")
64100   If IsNumeric(tStr) Then
64110     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64120       .PDFSigningSignPDF = CLng(tStr)
64130      Else
64140       If UseStandard Then
64150        .PDFSigningSignPDF = 0
64160       End If
64170     End If
64180    Else
64190     If UseStandard Then
64200      .PDFSigningSignPDF = 0
64210     End If
64220   End If
64230   tStr = hOpt.Retrieve("PDFUpdateMetadata")
64240   If IsNumeric(tStr) Then
64250     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
64260       .PDFUpdateMetadata = CLng(tStr)
64270      Else
64280       If UseStandard Then
64290        .PDFUpdateMetadata = 1
64300       End If
64310     End If
64320    Else
64330     If UseStandard Then
64340      .PDFUpdateMetadata = 1
64350     End If
64360   End If
64370   tStr = hOpt.Retrieve("PDFUserPass")
64380   If IsNumeric(tStr) Then
64390     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64400       .PDFUserPass = CLng(tStr)
64410      Else
64420       If UseStandard Then
64430        .PDFUserPass = 0
64440       End If
64450     End If
64460    Else
64470     If UseStandard Then
64480      .PDFUserPass = 0
64490     End If
64500   End If
64510   tStr = hOpt.Retrieve("PDFUserPasswordString")
64520   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64530     .PDFUserPasswordString = ""
64540    Else
64550     If LenB(tStr) > 0 Then
64560      .PDFUserPasswordString = tStr
64570     End If
64580   End If
64590   tStr = hOpt.Retrieve("PDFUseSecurity")
64600   If IsNumeric(tStr) Then
64610     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64620       .PDFUseSecurity = CLng(tStr)
64630      Else
64640       If UseStandard Then
64650        .PDFUseSecurity = 0
64660       End If
64670     End If
64680    Else
64690     If UseStandard Then
64700      .PDFUseSecurity = 0
64710     End If
64720   End If
64730   tStr = hOpt.Retrieve("PNGColorscount")
64740   If IsNumeric(tStr) Then
64750     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
64760       .PNGColorscount = CLng(tStr)
64770      Else
64780       If UseStandard Then
64790        .PNGColorscount = 0
64800       End If
64810     End If
64820    Else
64830     If UseStandard Then
64840      .PNGColorscount = 0
64850     End If
64860   End If
64870   tStr = hOpt.Retrieve("PNGResolution")
64880   If IsNumeric(tStr) Then
64890     If CLng(tStr) >= 1 Then
64900       .PNGResolution = CLng(tStr)
64910      Else
64920       If UseStandard Then
64930        .PNGResolution = 150
64940       End If
64950     End If
64960    Else
64970     If UseStandard Then
64980      .PNGResolution = 150
64990     End If
65000   End If
65010   tStr = hOpt.Retrieve("PrintAfterSaving")
65020   If IsNumeric(tStr) Then
65030     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65040       .PrintAfterSaving = CLng(tStr)
65050      Else
65060       If UseStandard Then
65070        .PrintAfterSaving = 0
65080       End If
65090     End If
65100    Else
65110     If UseStandard Then
65120      .PrintAfterSaving = 0
65130     End If
65140   End If
65150   tStr = hOpt.Retrieve("PrintAfterSavingDuplex")
65160   If IsNumeric(tStr) Then
65170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65180       .PrintAfterSavingDuplex = CLng(tStr)
65190      Else
65200       If UseStandard Then
65210        .PrintAfterSavingDuplex = 0
65220       End If
65230     End If
65240    Else
65250     If UseStandard Then
65260      .PrintAfterSavingDuplex = 0
65270     End If
65280   End If
65290   tStr = hOpt.Retrieve("PrintAfterSavingNoCancel")
65300   If IsNumeric(tStr) Then
65310     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65320       .PrintAfterSavingNoCancel = CLng(tStr)
65330      Else
65340       If UseStandard Then
65350        .PrintAfterSavingNoCancel = 0
65360       End If
65370     End If
65380    Else
65390     If UseStandard Then
65400      .PrintAfterSavingNoCancel = 0
65410     End If
65420   End If
65430   tStr = hOpt.Retrieve("PrintAfterSavingPrinter")
65440   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
65450     .PrintAfterSavingPrinter = ""
65460    Else
65470     If LenB(tStr) > 0 Then
65480      .PrintAfterSavingPrinter = tStr
65490     End If
65500   End If
65510   tStr = hOpt.Retrieve("PrintAfterSavingQueryUser")
65520   If IsNumeric(tStr) Then
65530     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
65540       .PrintAfterSavingQueryUser = CLng(tStr)
65550      Else
65560       If UseStandard Then
65570        .PrintAfterSavingQueryUser = 0
65580       End If
65590     End If
65600    Else
65610     If UseStandard Then
65620      .PrintAfterSavingQueryUser = 0
65630     End If
65640   End If
65650   tStr = hOpt.Retrieve("PrintAfterSavingTumble")
65660   If IsNumeric(tStr) Then
65670     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
65680       .PrintAfterSavingTumble = CLng(tStr)
65690      Else
65700       If UseStandard Then
65710        .PrintAfterSavingTumble = 0
65720       End If
65730     End If
65740    Else
65750     If UseStandard Then
65760      .PrintAfterSavingTumble = 0
65770     End If
65780   End If
65790   tStr = hOpt.Retrieve("PrinterStop")
65800   If IsNumeric(tStr) Then
65810     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65820       .PrinterStop = CLng(tStr)
65830      Else
65840       If UseStandard Then
65850        .PrinterStop = 0
65860       End If
65870     End If
65880    Else
65890     If UseStandard Then
65900      .PrinterStop = 0
65910     End If
65920   End If
65930   tStr = hOpt.Retrieve("PrinterTemppath")
65940   WriteToSpecialLogfile "hOpt.Retrieve(""PrinterTemppath"")=" & tStr
65950   WriteToSpecialLogfile "Options.PrinterTemppath1=" & .PrinterTemppath
65960   If hkey1 = HKEY_USERS Then
65970     If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then
65980       .PrinterTemppath = tStr
65990      Else
66000       If UseStandard Then
66010         .PrinterTemppath = GetTempPath
66020        Else
66030         .PrinterTemppath = tStr
66040       End If
66050     End If
66060    Else
66070     If LenB(Trim$(tStr)) > 0 Then
66080      If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then
66090        .PrinterTemppath = tStr
66100       Else
66110        MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))
66120        If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then
66130          If UseStandard Then
66140            .PrinterTemppath = GetTempPath
66150           Else
66160            .PrinterTemppath = ""
66170            If NoMsg = False Then
66180             MsgBox "PrinterTemppath: '" & tStr & "' = '" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & "'" & _
             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
66200            End If
66210          End If
66220         Else
66230          .PrinterTemppath = tStr
66240        End If
66250      End If
66260     End If
66270   End If
66280   WriteToSpecialLogfile "Options.PrinterTemppath2=" & .PrinterTemppath
66290   tStr = hOpt.Retrieve("ProcessPriority")
66300   If IsNumeric(tStr) Then
66310     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
66320       .ProcessPriority = CLng(tStr)
66330      Else
66340       If UseStandard Then
66350        .ProcessPriority = 1
66360       End If
66370     End If
66380    Else
66390     If UseStandard Then
66400      .ProcessPriority = 1
66410     End If
66420   End If
66430   tStr = hOpt.Retrieve("ProgramFont")
66440   If LenB(tStr) = 0 And LenB("MS Sans Serif") > 0 And UseStandard Then
66450     .ProgramFont = "MS Sans Serif"
66460    Else
66470     If LenB(tStr) > 0 Then
66480      .ProgramFont = tStr
66490     End If
66500   End If
66510   tStr = hOpt.Retrieve("ProgramFontCharset")
66520   If IsNumeric(tStr) Then
66530     If CLng(tStr) >= 0 Then
66540       .ProgramFontCharset = CLng(tStr)
66550      Else
66560       If UseStandard Then
66570        .ProgramFontCharset = 0
66580       End If
66590     End If
66600    Else
66610     If UseStandard Then
66620      .ProgramFontCharset = 0
66630     End If
66640   End If
66650   tStr = hOpt.Retrieve("ProgramFontSize")
66660   If IsNumeric(tStr) Then
66670     If CLng(tStr) >= 6 And CLng(tStr) <= 72 Then
66680       .ProgramFontSize = CLng(tStr)
66690      Else
66700       If UseStandard Then
66710        .ProgramFontSize = 8
66720       End If
66730     End If
66740    Else
66750     If UseStandard Then
66760      .ProgramFontSize = 8
66770     End If
66780   End If
66790   tStr = hOpt.Retrieve("PSDColorsCount")
66800   If IsNumeric(tStr) Then
66810     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
66820       .PSDColorsCount = CLng(tStr)
66830      Else
66840       If UseStandard Then
66850        .PSDColorsCount = 0
66860       End If
66870     End If
66880    Else
66890     If UseStandard Then
66900      .PSDColorsCount = 0
66910     End If
66920   End If
66930   tStr = hOpt.Retrieve("PSDResolution")
66940   If IsNumeric(tStr) Then
66950     If CLng(tStr) >= 1 Then
66960       .PSDResolution = CLng(tStr)
66970      Else
66980       If UseStandard Then
66990        .PSDResolution = 150
67000       End If
67010     End If
67020    Else
67030     If UseStandard Then
67040      .PSDResolution = 150
67050     End If
67060   End If
67070   tStr = hOpt.Retrieve("PSLanguageLevel")
67080   If IsNumeric(tStr) Then
67090     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
67100       .PSLanguageLevel = CLng(tStr)
67110      Else
67120       If UseStandard Then
67130        .PSLanguageLevel = 2
67140       End If
67150     End If
67160    Else
67170     If UseStandard Then
67180      .PSLanguageLevel = 2
67190     End If
67200   End If
67210   tStr = hOpt.Retrieve("RAWColorsCount")
67220   If IsNumeric(tStr) Then
67230     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
67240       .RAWColorsCount = CLng(tStr)
67250      Else
67260       If UseStandard Then
67270        .RAWColorsCount = 0
67280       End If
67290     End If
67300    Else
67310     If UseStandard Then
67320      .RAWColorsCount = 0
67330     End If
67340   End If
67350   tStr = hOpt.Retrieve("RAWResolution")
67360   If IsNumeric(tStr) Then
67370     If CLng(tStr) >= 1 Then
67380       .RAWResolution = CLng(tStr)
67390      Else
67400       If UseStandard Then
67410        .RAWResolution = 150
67420       End If
67430     End If
67440    Else
67450     If UseStandard Then
67460      .RAWResolution = 150
67470     End If
67480   End If
67490   tStr = hOpt.Retrieve("RemoveAllKnownFileExtensions")
67500   If IsNumeric(tStr) Then
67510     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67520       .RemoveAllKnownFileExtensions = CLng(tStr)
67530      Else
67540       If UseStandard Then
67550        .RemoveAllKnownFileExtensions = 1
67560       End If
67570     End If
67580    Else
67590     If UseStandard Then
67600      .RemoveAllKnownFileExtensions = 1
67610     End If
67620   End If
67630   tStr = hOpt.Retrieve("RemoveSpaces")
67640   If IsNumeric(tStr) Then
67650     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67660       .RemoveSpaces = CLng(tStr)
67670      Else
67680       If UseStandard Then
67690        .RemoveSpaces = 1
67700       End If
67710     End If
67720    Else
67730     If UseStandard Then
67740      .RemoveSpaces = 1
67750     End If
67760   End If
67770   tStr = hOpt.Retrieve("RunProgramAfterSaving")
67780   If IsNumeric(tStr) Then
67790     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67800       .RunProgramAfterSaving = CLng(tStr)
67810      Else
67820       If UseStandard Then
67830        .RunProgramAfterSaving = 0
67840       End If
67850     End If
67860    Else
67870     If UseStandard Then
67880      .RunProgramAfterSaving = 0
67890     End If
67900   End If
67910   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramname")
67920   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67930     .RunProgramAfterSavingProgramname = ""
67940    Else
67950     If LenB(tStr) > 0 Then
67960      .RunProgramAfterSavingProgramname = tStr
67970     End If
67980   End If
67990   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramParameters")
68000   If LenB(tStr) = 0 And LenB("""<OutputFilename>""") > 0 And UseStandard Then
68010     .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
68020    Else
68030     If LenB(tStr) > 0 Then
68040      .RunProgramAfterSavingProgramParameters = tStr
68050     End If
68060   End If
68070   tStr = hOpt.Retrieve("RunProgramAfterSavingWaitUntilReady")
68080   If IsNumeric(tStr) Then
68090     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68100       .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
68110      Else
68120       If UseStandard Then
68130        .RunProgramAfterSavingWaitUntilReady = 1
68140       End If
68150     End If
68160    Else
68170     If UseStandard Then
68180      .RunProgramAfterSavingWaitUntilReady = 1
68190     End If
68200   End If
68210   tStr = hOpt.Retrieve("RunProgramAfterSavingWindowstyle")
68220   If IsNumeric(tStr) Then
68230     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
68240       .RunProgramAfterSavingWindowstyle = CLng(tStr)
68250      Else
68260       If UseStandard Then
68270        .RunProgramAfterSavingWindowstyle = 1
68280       End If
68290     End If
68300    Else
68310     If UseStandard Then
68320      .RunProgramAfterSavingWindowstyle = 1
68330     End If
68340   End If
68350   tStr = hOpt.Retrieve("RunProgramBeforeSaving")
68360   If IsNumeric(tStr) Then
68370     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68380       .RunProgramBeforeSaving = CLng(tStr)
68390      Else
68400       If UseStandard Then
68410        .RunProgramBeforeSaving = 0
68420       End If
68430     End If
68440    Else
68450     If UseStandard Then
68460      .RunProgramBeforeSaving = 0
68470     End If
68480   End If
68490   tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramname")
68500   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
68510     .RunProgramBeforeSavingProgramname = ""
68520    Else
68530     If LenB(tStr) > 0 Then
68540      .RunProgramBeforeSavingProgramname = tStr
68550     End If
68560   End If
68570   tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramParameters")
68580   If LenB(tStr) = 0 And LenB("""<TempFilename>""") > 0 And UseStandard Then
68590     .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
68600    Else
68610     If LenB(tStr) > 0 Then
68620      .RunProgramBeforeSavingProgramParameters = tStr
68630     End If
68640   End If
68650   tStr = hOpt.Retrieve("RunProgramBeforeSavingWindowstyle")
68660   If IsNumeric(tStr) Then
68670     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
68680       .RunProgramBeforeSavingWindowstyle = CLng(tStr)
68690      Else
68700       If UseStandard Then
68710        .RunProgramBeforeSavingWindowstyle = 1
68720       End If
68730     End If
68740    Else
68750     If UseStandard Then
68760      .RunProgramBeforeSavingWindowstyle = 1
68770     End If
68780   End If
68790   tStr = hOpt.Retrieve("SaveFilename")
68800   If LenB(tStr) = 0 And LenB("<Title>") > 0 And UseStandard Then
68810     .SaveFilename = "<Title>"
68820    Else
68830     If LenB(tStr) > 0 Then
68840      .SaveFilename = tStr
68850     End If
68860   End If
68870   tStr = hOpt.Retrieve("SendEmailAfterAutoSaving")
68880   If IsNumeric(tStr) Then
68890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68900       .SendEmailAfterAutoSaving = CLng(tStr)
68910      Else
68920       If UseStandard Then
68930        .SendEmailAfterAutoSaving = 0
68940       End If
68950     End If
68960    Else
68970     If UseStandard Then
68980      .SendEmailAfterAutoSaving = 0
68990     End If
69000   End If
69010   tStr = hOpt.Retrieve("SendMailMethod")
69020   If IsNumeric(tStr) Then
69030     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
69040       .SendMailMethod = CLng(tStr)
69050      Else
69060       If UseStandard Then
69070        .SendMailMethod = 0
69080       End If
69090     End If
69100    Else
69110     If UseStandard Then
69120      .SendMailMethod = 0
69130     End If
69140   End If
69150   tStr = hOpt.Retrieve("ShowAnimation")
69160   If IsNumeric(tStr) Then
69170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69180       .ShowAnimation = CLng(tStr)
69190      Else
69200       If UseStandard Then
69210        .ShowAnimation = 1
69220       End If
69230     End If
69240    Else
69250     If UseStandard Then
69260      .ShowAnimation = 1
69270     End If
69280   End If
69290   tStr = hOpt.Retrieve("StampFontColor")
69300   If LenB(tStr) = 0 And LenB("#FF0000") > 0 And UseStandard Then
69310     .StampFontColor = "#FF0000"
69320    Else
69330     If LenB(tStr) > 0 Then
69340      .StampFontColor = tStr
69350     End If
69360   End If
69370   tStr = hOpt.Retrieve("StampFontname")
69380   If LenB(tStr) = 0 And LenB("Arial") > 0 And UseStandard Then
69390     .StampFontname = "Arial"
69400    Else
69410     If LenB(tStr) > 0 Then
69420      .StampFontname = tStr
69430     End If
69440   End If
69450   tStr = hOpt.Retrieve("StampFontsize")
69460   If IsNumeric(tStr) Then
69470     If CLng(tStr) >= 1 Then
69480       .StampFontsize = CLng(tStr)
69490      Else
69500       If UseStandard Then
69510        .StampFontsize = 48
69520       End If
69530     End If
69540    Else
69550     If UseStandard Then
69560      .StampFontsize = 48
69570     End If
69580   End If
69590   tStr = hOpt.Retrieve("StampOutlineFontthickness")
69600   If IsNumeric(tStr) Then
69610     If CLng(tStr) >= 0 Then
69620       .StampOutlineFontthickness = CLng(tStr)
69630      Else
69640       If UseStandard Then
69650        .StampOutlineFontthickness = 0
69660       End If
69670     End If
69680    Else
69690     If UseStandard Then
69700      .StampOutlineFontthickness = 0
69710     End If
69720   End If
69730   tStr = hOpt.Retrieve("StampString")
69740   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69750     .StampString = ""
69760    Else
69770     If LenB(tStr) > 0 Then
69780      .StampString = tStr
69790     End If
69800   End If
69810   tStr = hOpt.Retrieve("StampUseOutlineFont")
69820   If IsNumeric(tStr) Then
69830     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69840       .StampUseOutlineFont = CLng(tStr)
69850      Else
69860       If UseStandard Then
69870        .StampUseOutlineFont = 1
69880       End If
69890     End If
69900    Else
69910     If UseStandard Then
69920      .StampUseOutlineFont = 1
69930     End If
69940   End If
69950   tStr = hOpt.Retrieve("StandardAuthor")
69960   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69970     .StandardAuthor = ""
69980    Else
69990     If LenB(tStr) > 0 Then
70000      .StandardAuthor = tStr
70010     End If
70020   End If
70030   tStr = hOpt.Retrieve("StandardCreationdate")
70040   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70050     .StandardCreationdate = ""
70060    Else
70070     If LenB(tStr) > 0 Then
70080      .StandardCreationdate = tStr
70090     End If
70100   End If
70110   tStr = hOpt.Retrieve("StandardDateformat")
70120   If LenB(tStr) = 0 And LenB("YYYYMMDDHHNNSS") > 0 And UseStandard Then
70130     .StandardDateformat = "YYYYMMDDHHNNSS"
70140    Else
70150     If LenB(tStr) > 0 Then
70160      .StandardDateformat = tStr
70170     End If
70180   End If
70190   tStr = hOpt.Retrieve("StandardKeywords")
70200   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70210     .StandardKeywords = ""
70220    Else
70230     If LenB(tStr) > 0 Then
70240      .StandardKeywords = tStr
70250     End If
70260   End If
70270   tStr = hOpt.Retrieve("StandardMailDomain")
70280   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70290     .StandardMailDomain = ""
70300    Else
70310     If LenB(tStr) > 0 Then
70320      .StandardMailDomain = tStr
70330     End If
70340   End If
70350   tStr = hOpt.Retrieve("StandardModifydate")
70360   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70370     .StandardModifydate = ""
70380    Else
70390     If LenB(tStr) > 0 Then
70400      .StandardModifydate = tStr
70410     End If
70420   End If
70430   tStr = hOpt.Retrieve("StandardSaveformat")
70440   If IsNumeric(tStr) Then
70450     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
70460       .StandardSaveformat = CLng(tStr)
70470      Else
70480       If UseStandard Then
70490        .StandardSaveformat = 0
70500       End If
70510     End If
70520    Else
70530     If UseStandard Then
70540      .StandardSaveformat = 0
70550     End If
70560   End If
70570   tStr = hOpt.Retrieve("StandardSubject")
70580   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70590     .StandardSubject = ""
70600    Else
70610     If LenB(tStr) > 0 Then
70620      .StandardSubject = tStr
70630     End If
70640   End If
70650   tStr = hOpt.Retrieve("StandardTitle")
70660   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70670     .StandardTitle = ""
70680    Else
70690     If LenB(tStr) > 0 Then
70700      .StandardTitle = tStr
70710     End If
70720   End If
70730   tStr = hOpt.Retrieve("StartStandardProgram")
70740   If IsNumeric(tStr) Then
70750     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70760       .StartStandardProgram = CLng(tStr)
70770      Else
70780       If UseStandard Then
70790        .StartStandardProgram = 1
70800       End If
70810     End If
70820    Else
70830     If UseStandard Then
70840      .StartStandardProgram = 1
70850     End If
70860   End If
70870   tStr = hOpt.Retrieve("TIFFColorscount")
70880   If IsNumeric(tStr) Then
70890     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
70900       .TIFFColorscount = CLng(tStr)
70910      Else
70920       If UseStandard Then
70930        .TIFFColorscount = 0
70940       End If
70950     End If
70960    Else
70970     If UseStandard Then
70980      .TIFFColorscount = 0
70990     End If
71000   End If
71010   tStr = hOpt.Retrieve("TIFFResolution")
71020   If IsNumeric(tStr) Then
71030     If CLng(tStr) >= 1 Then
71040       .TIFFResolution = CLng(tStr)
71050      Else
71060       If UseStandard Then
71070        .TIFFResolution = 150
71080       End If
71090     End If
71100    Else
71110     If UseStandard Then
71120      .TIFFResolution = 150
71130     End If
71140   End If
71150   tStr = hOpt.Retrieve("Toolbars")
71160   If IsNumeric(tStr) Then
71170     If CLng(tStr) >= 0 Then
71180       .Toolbars = CLng(tStr)
71190      Else
71200       If UseStandard Then
71210        .Toolbars = 1
71220       End If
71230     End If
71240    Else
71250     If UseStandard Then
71260      .Toolbars = 1
71270     End If
71280   End If
71290   tStr = hOpt.Retrieve("UpdateInterval")
71300   If IsNumeric(tStr) Then
71310     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
71320       .UpdateInterval = CLng(tStr)
71330      Else
71340       If UseStandard Then
71350        .UpdateInterval = 2
71360       End If
71370     End If
71380    Else
71390     If UseStandard Then
71400      .UpdateInterval = 2
71410     End If
71420   End If
71430   tStr = hOpt.Retrieve("UseAutosave")
71440   If IsNumeric(tStr) Then
71450     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71460       .UseAutosave = CLng(tStr)
71470      Else
71480       If UseStandard Then
71490        .UseAutosave = 0
71500       End If
71510     End If
71520    Else
71530     If UseStandard Then
71540      .UseAutosave = 0
71550     End If
71560   End If
71570   tStr = hOpt.Retrieve("UseAutosaveDirectory")
71580   If IsNumeric(tStr) Then
71590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71600       .UseAutosaveDirectory = CLng(tStr)
71610      Else
71620       If UseStandard Then
71630        .UseAutosaveDirectory = 1
71640       End If
71650     End If
71660    Else
71670     If UseStandard Then
71680      .UseAutosaveDirectory = 1
71690     End If
71700   End If
71710   tStr = hOpt.Retrieve("UseCreationDateNow")
71720   If IsNumeric(tStr) Then
71730     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71740       .UseCreationDateNow = CLng(tStr)
71750      Else
71760       If UseStandard Then
71770        .UseCreationDateNow = 0
71780       End If
71790     End If
71800    Else
71810     If UseStandard Then
71820      .UseCreationDateNow = 0
71830     End If
71840   End If
71850   tStr = hOpt.Retrieve("UseCustomPaperSize")
71860   If LenB(tStr) = 0 And LenB("0") > 0 And UseStandard Then
71870     .UseCustomPaperSize = "0"
71880    Else
71890     If LenB(tStr) > 0 Then
71900      .UseCustomPaperSize = tStr
71910     End If
71920   End If
71930   tStr = hOpt.Retrieve("UseFixPapersize")
71940   If IsNumeric(tStr) Then
71950     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71960       .UseFixPapersize = CLng(tStr)
71970      Else
71980       If UseStandard Then
71990        .UseFixPapersize = 0
72000       End If
72010     End If
72020    Else
72030     If UseStandard Then
72040      .UseFixPapersize = 0
72050     End If
72060   End If
72070   tStr = hOpt.Retrieve("UseStandardAuthor")
72080   If IsNumeric(tStr) Then
72090     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
72100       .UseStandardAuthor = CLng(tStr)
72110      Else
72120       If UseStandard Then
72130        .UseStandardAuthor = 0
72140       End If
72150     End If
72160    Else
72170     If UseStandard Then
72180      .UseStandardAuthor = 0
72190     End If
72200   End If
72210  End With
72220  Set ini = Nothing
72230  ReadOptionsINI = myOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "ReadOptionsINI")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CorrectOptionsAfterLoading(Options As tOptions) As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fns As String, fnsf() As String, nfnsf() As String, i As Long, j As Long
50020  fns = Options.FilenameSubstitutions
50030  If LenB(fns) = 0 Then
50040   CorrectOptionsAfterLoading = Options
50050   Exit Function
50060  End If
50070  If InStr(1, fns, "\") = 0 Then
50080   CorrectOptionsAfterLoading = Options
50090   Exit Function
50100  End If
50110  fnsf = Split(fns, "\")
50120  ReDim nfnsf(0)
50130  nfnsf(0) = fnsf(0)
50140  For i = 1 To UBound(fnsf)
50150   For j = LBound(nfnsf) To UBound(nfnsf)
50160    If nfnsf(j) = fnsf(i) Then
50170     Exit For
50180    End If
50190    DoEvents
50200   Next j
50210   If j > UBound(nfnsf) Then
50220    ReDim Preserve nfnsf(UBound(nfnsf) + 1)
50230    nfnsf(UBound(nfnsf)) = fnsf(i)
50240   End If
50250   DoEvents
50260  Next i
50270  fns = nfnsf(0)
50280  For i = 1 To UBound(nfnsf)
50290   fns = fns & "\" & nfnsf(i)
50300   DoEvents
50310  Next i
50320  Options.FilenameSubstitutions = fns
50330  CorrectOptionsAfterLoading = Options
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "CorrectOptionsAfterLoading")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub CorrectOptionsBeforeSaving()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Options.AutosaveDirectory = Trim$(Options.AutosaveDirectory)
50020  Options.PrinterTemppath = Trim$(Options.PrinterTemppath)
50030  If LenB(Options.AutosaveDirectory) = 0 Then
50040   Options.AutosaveDirectory = "<MyFiles>\"
50050  End If
50060  If LenB(Options.PrinterTemppath) = 0 Then
50070   Options.PrinterTemppath = "<Temp>PDFCreator\"
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "CorrectOptionsBeforeSaving")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SaveOptions(sOptions As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  CorrectOptionsBeforeSaving
50020  If InstalledAsServer Then
50030    If UseINI Then
50040      SaveOptionsINI sOptions, CompletePath(GetCommonAppData) & "PDFCreator.ini"
50050     Else
50060      SaveOptionsREG sOptions, HKEY_LOCAL_MACHINE
50070    End If
50080   Else
50090    If UseINI Then
50100      SaveOptionsINI sOptions, PDFCreatorINIFile
50110     Else
50120      SaveOptionsREG sOptions
50130    End If
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "SaveOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SaveOption(sOptions As tOptions, OptionName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If InstalledAsServer Then
50020    If UseINI Then
50030      SaveOptionINI sOptions, OptionName, CompletePath(GetCommonAppData) & "PDFCreator.ini"
50040     Else
50050      SaveOptionREG sOptions, OptionName, HKEY_LOCAL_MACHINE
50060    End If
50070   Else
50080    If UseINI Then
50090      SaveOptionINI sOptions, OptionName, PDFCreatorINIFile
50100     Else
50110      SaveOptionREG sOptions, OptionName
50120    End If
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "SaveOption")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SaveOptionINI(sOptions As tOptions, OptionName As String, PDFCreatorINIFile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI
50020  Set ini = New clsINI
50030  ini.filename = PDFCreatorINIFile
50040  ini.Section = "Options"
50050  If ini.CheckIniFile = False Then
50060   ini.CreateIniFile
50070  End If
50080  With sOptions
50091   Select Case UCase$(OptionName)
        Case "ADDITIONALGHOSTSCRIPTPARAMETERS": ini.SaveKey CStr(.AdditionalGhostscriptParameters), "AdditionalGhostscriptParameters"
50110   Case "ADDITIONALGHOSTSCRIPTSEARCHPATH": ini.SaveKey CStr(.AdditionalGhostscriptSearchpath), "AdditionalGhostscriptSearchpath"
50120   Case "ADDWINDOWSFONTPATH": ini.SaveKey CStr(Abs(.AddWindowsFontpath)), "AddWindowsFontpath"
50130   Case "AUTOSAVEDIRECTORY": ini.SaveKey CStr(.AutosaveDirectory), "AutosaveDirectory"
50140   Case "AUTOSAVEFILENAME": ini.SaveKey CStr(.AutosaveFilename), "AutosaveFilename"
50150   Case "AUTOSAVEFORMAT": ini.SaveKey CStr(.AutosaveFormat), "AutosaveFormat"
50160   Case "AUTOSAVESTARTSTANDARDPROGRAM": ini.SaveKey CStr(Abs(.AutosaveStartStandardProgram)), "AutosaveStartStandardProgram"
50170   Case "BMPCOLORSCOUNT": ini.SaveKey CStr(.BMPColorscount), "BMPColorscount"
50180   Case "BMPRESOLUTION": ini.SaveKey CStr(.BMPResolution), "BMPResolution"
50190   Case "CLIENTCOMPUTERRESOLVEIPADDRESS": ini.SaveKey CStr(Abs(.ClientComputerResolveIPAddress)), "ClientComputerResolveIPAddress"
50200   Case "COUNTER": ini.SaveKey CStr(.Counter), "Counter"
50210   Case "DEVICEHEIGHTPOINTS": ini.SaveKey Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), "DeviceHeightPoints"
50220   Case "DEVICEWIDTHPOINTS": ini.SaveKey Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), "DeviceWidthPoints"
50230   Case "DIRECTORYGHOSTSCRIPTBINARIES": ini.SaveKey CStr(.DirectoryGhostscriptBinaries), "DirectoryGhostscriptBinaries"
50240   Case "DIRECTORYGHOSTSCRIPTFONTS": ini.SaveKey CStr(.DirectoryGhostscriptFonts), "DirectoryGhostscriptFonts"
50250   Case "DIRECTORYGHOSTSCRIPTLIBRARIES": ini.SaveKey CStr(.DirectoryGhostscriptLibraries), "DirectoryGhostscriptLibraries"
50260   Case "DIRECTORYGHOSTSCRIPTRESOURCE": ini.SaveKey CStr(.DirectoryGhostscriptResource), "DirectoryGhostscriptResource"
50270   Case "DISABLEEMAIL": ini.SaveKey CStr(Abs(.DisableEmail)), "DisableEmail"
50280   Case "DONTUSEDOCUMENTSETTINGS": ini.SaveKey CStr(Abs(.DontUseDocumentSettings)), "DontUseDocumentSettings"
50290   Case "EPSLANGUAGELEVEL": ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
50300   Case "FILENAMESUBSTITUTIONS": ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
50310   Case "FILENAMESUBSTITUTIONSONLYINTITLE": ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
50320   Case "JPEGCOLORSCOUNT": ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50330   Case "JPEGQUALITY": ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50340   Case "JPEGRESOLUTION": ini.SaveKey CStr(.JPEGResolution), "JPEGResolution"
50350   Case "LANGUAGE": ini.SaveKey CStr(.Language), "Language"
50360   Case "LASTSAVEDIRECTORY": ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50370   Case "LASTUPDATECHECK": ini.SaveKey CStr(.LastUpdateCheck), "LastUpdateCheck"
50380   Case "LOGGING": ini.SaveKey CStr(Abs(.Logging)), "Logging"
50390   Case "LOGLINES": ini.SaveKey CStr(.LogLines), "LogLines"
50400   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50410   Case "NOPROCESSINGATSTARTUP": ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
50420   Case "NOPSCHECK": ini.SaveKey CStr(Abs(.NoPSCheck)), "NoPSCheck"
50430   Case "ONEPAGEPERFILE": ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50440   Case "OPTIONSDESIGN": ini.SaveKey CStr(.OptionsDesign), "OptionsDesign"
50450   Case "OPTIONSENABLED": ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50460   Case "OPTIONSVISIBLE": ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50470   Case "PAPERSIZE": ini.SaveKey CStr(.Papersize), "Papersize"
50480   Case "PCLCOLORSCOUNT": ini.SaveKey CStr(.PCLColorsCount), "PCLColorsCount"
50490   Case "PCLRESOLUTION": ini.SaveKey CStr(.PCLResolution), "PCLResolution"
50500   Case "PCXCOLORSCOUNT": ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50510   Case "PCXRESOLUTION": ini.SaveKey CStr(.PCXResolution), "PCXResolution"
50520   Case "PDFALLOWASSEMBLY": ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50530   Case "PDFALLOWDEGRADEDPRINTING": ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50540   Case "PDFALLOWFILLIN": ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50550   Case "PDFALLOWSCREENREADERS": ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50560   Case "PDFCOLORSCMYKTORGB": ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50570   Case "PDFCOLORSCOLORMODEL": ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50580   Case "PDFCOLORSPRESERVEHALFTONE": ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50590   Case "PDFCOLORSPRESERVEOVERPRINT": ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50600   Case "PDFCOLORSPRESERVETRANSFER": ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50610   Case "PDFCOMPRESSIONCOLORCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50620   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50630   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGHighFactor"
50640   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGLowFactor"
50650   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMaximumFactor"
50660   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMediumFactor"
50670   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMinimumFactor"
50680   Case "PDFCOMPRESSIONCOLORRESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50690   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50700   Case "PDFCOMPRESSIONCOLORRESOLUTION": ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50710   Case "PDFCOMPRESSIONGREYCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50720   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50730   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGHighFactor"
50740   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGLowFactor"
50750   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMaximumFactor"
50760   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMediumFactor"
50770   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMinimumFactor"
50780   Case "PDFCOMPRESSIONGREYRESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50790   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50800   Case "PDFCOMPRESSIONGREYRESOLUTION": ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50810   Case "PDFCOMPRESSIONMONOCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50820   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50830   Case "PDFCOMPRESSIONMONORESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50840   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50850   Case "PDFCOMPRESSIONMONORESOLUTION": ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50860   Case "PDFCOMPRESSIONTEXTCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50870   Case "PDFDISALLOWCOPY": ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50880   Case "PDFDISALLOWMODIFYANNOTATIONS": ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50890   Case "PDFDISALLOWMODIFYCONTENTS": ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50900   Case "PDFDISALLOWPRINTING": ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50910   Case "PDFENCRYPTOR": ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50920   Case "PDFFONTSEMBEDALL": ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50930   Case "PDFFONTSSUBSETFONTS": ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50940   Case "PDFFONTSSUBSETFONTSPERCENT": ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50950   Case "PDFGENERALASCII85": ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50960   Case "PDFGENERALAUTOROTATE": ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50970   Case "PDFGENERALCOMPATIBILITY": ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50980   Case "PDFGENERALDEFAULT": ini.SaveKey CStr(.PDFGeneralDefault), "PDFGeneralDefault"
50990   Case "PDFGENERALOVERPRINT": ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
51000   Case "PDFGENERALRESOLUTION": ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
51010   Case "PDFHIGHENCRYPTION": ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
51020   Case "PDFLOWENCRYPTION": ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
51030   Case "PDFOPTIMIZE": ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
51040   Case "PDFOWNERPASS": ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
51050   Case "PDFOWNERPASSWORDSTRING": ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
51060   Case "PDFSIGNINGMULTISIGNATURE": ini.SaveKey CStr(Abs(.PDFSigningMultiSignature)), "PDFSigningMultiSignature"
51070   Case "PDFSIGNINGPFXFILE": ini.SaveKey CStr(.PDFSigningPFXFile), "PDFSigningPFXFile"
51080   Case "PDFSIGNINGPFXFILEPASSWORD": ini.SaveKey CStr(.PDFSigningPFXFilePassword), "PDFSigningPFXFilePassword"
51090   Case "PDFSIGNINGSIGNATURECONTACT": ini.SaveKey CStr(.PDFSigningSignatureContact), "PDFSigningSignatureContact"
51100   Case "PDFSIGNINGSIGNATURELEFTX": ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), "PDFSigningSignatureLeftX"
51110   Case "PDFSIGNINGSIGNATURELEFTY": ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), "PDFSigningSignatureLeftY"
51120   Case "PDFSIGNINGSIGNATURELOCATION": ini.SaveKey CStr(.PDFSigningSignatureLocation), "PDFSigningSignatureLocation"
51130   Case "PDFSIGNINGSIGNATUREREASON": ini.SaveKey CStr(.PDFSigningSignatureReason), "PDFSigningSignatureReason"
51140   Case "PDFSIGNINGSIGNATURERIGHTX": ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), "PDFSigningSignatureRightX"
51150   Case "PDFSIGNINGSIGNATURERIGHTY": ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), "PDFSigningSignatureRightY"
51160   Case "PDFSIGNINGSIGNATUREVISIBLE": ini.SaveKey CStr(Abs(.PDFSigningSignatureVisible)), "PDFSigningSignatureVisible"
51170   Case "PDFSIGNINGSIGNPDF": ini.SaveKey CStr(Abs(.PDFSigningSignPDF)), "PDFSigningSignPDF"
51180   Case "PDFUPDATEMETADATA": ini.SaveKey CStr(.PDFUpdateMetadata), "PDFUpdateMetadata"
51190   Case "PDFUSERPASS": ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
51200   Case "PDFUSERPASSWORDSTRING": ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
51210   Case "PDFUSESECURITY": ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
51220   Case "PNGCOLORSCOUNT": ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
51230   Case "PNGRESOLUTION": ini.SaveKey CStr(.PNGResolution), "PNGResolution"
51240   Case "PRINTAFTERSAVING": ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
51250   Case "PRINTAFTERSAVINGDUPLEX": ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
51260   Case "PRINTAFTERSAVINGNOCANCEL": ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
51270   Case "PRINTAFTERSAVINGPRINTER": ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
51280   Case "PRINTAFTERSAVINGQUERYUSER": ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
51290   Case "PRINTAFTERSAVINGTUMBLE": ini.SaveKey CStr(.PrintAfterSavingTumble), "PrintAfterSavingTumble"
51300   Case "PRINTERSTOP": ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
51310   Case "PRINTERTEMPPATH": ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
51320   Case "PROCESSPRIORITY": ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
51330   Case "PROGRAMFONT": ini.SaveKey CStr(.ProgramFont), "ProgramFont"
51340   Case "PROGRAMFONTCHARSET": ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
51350   Case "PROGRAMFONTSIZE": ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
51360   Case "PSDCOLORSCOUNT": ini.SaveKey CStr(.PSDColorsCount), "PSDColorsCount"
51370   Case "PSDRESOLUTION": ini.SaveKey CStr(.PSDResolution), "PSDResolution"
51380   Case "PSLANGUAGELEVEL": ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
51390   Case "RAWCOLORSCOUNT": ini.SaveKey CStr(.RAWColorsCount), "RAWColorsCount"
51400   Case "RAWRESOLUTION": ini.SaveKey CStr(.RAWResolution), "RAWResolution"
51410   Case "REMOVEALLKNOWNFILEEXTENSIONS": ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
51420   Case "REMOVESPACES": ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
51430   Case "RUNPROGRAMAFTERSAVING": ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
51440   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
51450   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
51460   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51470   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51480   Case "RUNPROGRAMBEFORESAVING": ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51490   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51500   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51510   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51520   Case "SAVEFILENAME": ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51530   Case "SENDEMAILAFTERAUTOSAVING": ini.SaveKey CStr(Abs(.SendEmailAfterAutoSaving)), "SendEmailAfterAutoSaving"
51540   Case "SENDMAILMETHOD": ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51550   Case "SHOWANIMATION": ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51560   Case "STAMPFONTCOLOR": ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51570   Case "STAMPFONTNAME": ini.SaveKey CStr(.StampFontname), "StampFontname"
51580   Case "STAMPFONTSIZE": ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51590   Case "STAMPOUTLINEFONTTHICKNESS": ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51600   Case "STAMPSTRING": ini.SaveKey CStr(.StampString), "StampString"
51610   Case "STAMPUSEOUTLINEFONT": ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51620   Case "STANDARDAUTHOR": ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51630   Case "STANDARDCREATIONDATE": ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51640   Case "STANDARDDATEFORMAT": ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51650   Case "STANDARDKEYWORDS": ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51660   Case "STANDARDMAILDOMAIN": ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51670   Case "STANDARDMODIFYDATE": ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51680   Case "STANDARDSAVEFORMAT": ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51690   Case "STANDARDSUBJECT": ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51700   Case "STANDARDTITLE": ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51710   Case "STARTSTANDARDPROGRAM": ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51720   Case "TIFFCOLORSCOUNT": ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51730   Case "TIFFRESOLUTION": ini.SaveKey CStr(.TIFFResolution), "TIFFResolution"
51740   Case "TOOLBARS": ini.SaveKey CStr(.Toolbars), "Toolbars"
51750   Case "UPDATEINTERVAL": ini.SaveKey CStr(.UpdateInterval), "UpdateInterval"
51760   Case "USEAUTOSAVE": ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51770   Case "USEAUTOSAVEDIRECTORY": ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51780   Case "USECREATIONDATENOW": ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51790   Case "USECUSTOMPAPERSIZE": ini.SaveKey CStr(.UseCustomPaperSize), "UseCustomPaperSize"
51800   Case "USEFIXPAPERSIZE": ini.SaveKey CStr(Abs(.UseFixPapersize)), "UseFixPapersize"
51810   Case "USESTANDARDAUTHOR": ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51820   End Select
51830  End With
51840  Set ini = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "SaveOptionINI")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SaveOptionsINI(sOptions As tOptions, PDFCreatorINIFile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI
50020  Set ini = New clsINI
50030  ini.filename = PDFCreatorINIFile
50040  ini.Section = "Options"
50050  If ini.CheckIniFile = False Then
50060   ini.CreateIniFile
50070  End If
50080  With sOptions
50090   ini.SaveKey CStr(.AdditionalGhostscriptParameters), "AdditionalGhostscriptParameters"
50100   ini.SaveKey CStr(.AdditionalGhostscriptSearchpath), "AdditionalGhostscriptSearchpath"
50110   ini.SaveKey CStr(Abs(.AddWindowsFontpath)), "AddWindowsFontpath"
50120   ini.SaveKey CStr(.AutosaveDirectory), "AutosaveDirectory"
50130   ini.SaveKey CStr(.AutosaveFilename), "AutosaveFilename"
50140   ini.SaveKey CStr(.AutosaveFormat), "AutosaveFormat"
50150   ini.SaveKey CStr(Abs(.AutosaveStartStandardProgram)), "AutosaveStartStandardProgram"
50160   ini.SaveKey CStr(.BMPColorscount), "BMPColorscount"
50170   ini.SaveKey CStr(.BMPResolution), "BMPResolution"
50180   ini.SaveKey CStr(Abs(.ClientComputerResolveIPAddress)), "ClientComputerResolveIPAddress"
50190   ini.SaveKey CStr(.Counter), "Counter"
50200   ini.SaveKey Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), "DeviceHeightPoints"
50210   ini.SaveKey Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), "DeviceWidthPoints"
50220   ini.SaveKey CStr(.DirectoryGhostscriptBinaries), "DirectoryGhostscriptBinaries"
50230   ini.SaveKey CStr(.DirectoryGhostscriptFonts), "DirectoryGhostscriptFonts"
50240   ini.SaveKey CStr(.DirectoryGhostscriptLibraries), "DirectoryGhostscriptLibraries"
50250   ini.SaveKey CStr(.DirectoryGhostscriptResource), "DirectoryGhostscriptResource"
50260   ini.SaveKey CStr(Abs(.DisableEmail)), "DisableEmail"
50270   ini.SaveKey CStr(Abs(.DontUseDocumentSettings)), "DontUseDocumentSettings"
50280   ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
50290   ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
50300   ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
50310   ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50320   ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50330   ini.SaveKey CStr(.JPEGResolution), "JPEGResolution"
50340   ini.SaveKey CStr(.Language), "Language"
50350   ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50360   ini.SaveKey CStr(.LastUpdateCheck), "LastUpdateCheck"
50370   ini.SaveKey CStr(Abs(.Logging)), "Logging"
50380   ini.SaveKey CStr(.LogLines), "LogLines"
50390   ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50400   ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
50410   ini.SaveKey CStr(Abs(.NoPSCheck)), "NoPSCheck"
50420   ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50430   ini.SaveKey CStr(.OptionsDesign), "OptionsDesign"
50440   ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50450   ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50460   ini.SaveKey CStr(.Papersize), "Papersize"
50470   ini.SaveKey CStr(.PCLColorsCount), "PCLColorsCount"
50480   ini.SaveKey CStr(.PCLResolution), "PCLResolution"
50490   ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50500   ini.SaveKey CStr(.PCXResolution), "PCXResolution"
50510   ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50520   ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50530   ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50540   ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50550   ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50560   ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50570   ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50580   ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50590   ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50600   ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50610   ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50620   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGHighFactor"
50630   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGLowFactor"
50640   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMaximumFactor"
50650   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMediumFactor"
50660   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMinimumFactor"
50670   ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50680   ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50690   ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50700   ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50710   ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50720   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGHighFactor"
50730   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGLowFactor"
50740   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMaximumFactor"
50750   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMediumFactor"
50760   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMinimumFactor"
50770   ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50780   ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50790   ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50800   ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50810   ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50820   ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50830   ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50840   ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50850   ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50860   ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50870   ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50880   ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50890   ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50900   ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50910   ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50920   ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50930   ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50940   ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50950   ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50960   ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50970   ini.SaveKey CStr(.PDFGeneralDefault), "PDFGeneralDefault"
50980   ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50990   ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
51000   ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
51010   ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
51020   ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
51030   ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
51040   ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
51050   ini.SaveKey CStr(Abs(.PDFSigningMultiSignature)), "PDFSigningMultiSignature"
51060   ini.SaveKey CStr(.PDFSigningPFXFile), "PDFSigningPFXFile"
51070   ini.SaveKey CStr(.PDFSigningPFXFilePassword), "PDFSigningPFXFilePassword"
51080   ini.SaveKey CStr(.PDFSigningSignatureContact), "PDFSigningSignatureContact"
51090   ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), "PDFSigningSignatureLeftX"
51100   ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), "PDFSigningSignatureLeftY"
51110   ini.SaveKey CStr(.PDFSigningSignatureLocation), "PDFSigningSignatureLocation"
51120   ini.SaveKey CStr(.PDFSigningSignatureReason), "PDFSigningSignatureReason"
51130   ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), "PDFSigningSignatureRightX"
51140   ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), "PDFSigningSignatureRightY"
51150   ini.SaveKey CStr(Abs(.PDFSigningSignatureVisible)), "PDFSigningSignatureVisible"
51160   ini.SaveKey CStr(Abs(.PDFSigningSignPDF)), "PDFSigningSignPDF"
51170   ini.SaveKey CStr(.PDFUpdateMetadata), "PDFUpdateMetadata"
51180   ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
51190   ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
51200   ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
51210   ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
51220   ini.SaveKey CStr(.PNGResolution), "PNGResolution"
51230   ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
51240   ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
51250   ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
51260   ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
51270   ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
51280   ini.SaveKey CStr(.PrintAfterSavingTumble), "PrintAfterSavingTumble"
51290   ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
51300   ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
51310   ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
51320   ini.SaveKey CStr(.ProgramFont), "ProgramFont"
51330   ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
51340   ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
51350   ini.SaveKey CStr(.PSDColorsCount), "PSDColorsCount"
51360   ini.SaveKey CStr(.PSDResolution), "PSDResolution"
51370   ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
51380   ini.SaveKey CStr(.RAWColorsCount), "RAWColorsCount"
51390   ini.SaveKey CStr(.RAWResolution), "RAWResolution"
51400   ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
51410   ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
51420   ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
51430   ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
51440   ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
51450   ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51460   ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51470   ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51480   ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51490   ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51500   ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51510   ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51520   ini.SaveKey CStr(Abs(.SendEmailAfterAutoSaving)), "SendEmailAfterAutoSaving"
51530   ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51540   ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51550   ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51560   ini.SaveKey CStr(.StampFontname), "StampFontname"
51570   ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51580   ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51590   ini.SaveKey CStr(.StampString), "StampString"
51600   ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51610   ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51620   ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51630   ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51640   ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51650   ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51660   ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51670   ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51680   ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51690   ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51700   ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51710   ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51720   ini.SaveKey CStr(.TIFFResolution), "TIFFResolution"
51730   ini.SaveKey CStr(.Toolbars), "Toolbars"
51740   ini.SaveKey CStr(.UpdateInterval), "UpdateInterval"
51750   ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51760   ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51770   ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51780   ini.SaveKey CStr(.UseCustomPaperSize), "UseCustomPaperSize"
51790   ini.SaveKey CStr(Abs(.UseFixPapersize)), "UseFixPapersize"
51800   ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51810  End With
51820  Set ini = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "SaveOptionsINI")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function ReadOptionsReg(myOptions As tOptions, KeyRoot As String, Optional hkey1 As hkey = HKEY_CURRENT_USER, Optional NoMsg As Boolean = False, Optional UseStandard As Boolean = True) As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, tStr As String
50020  Set reg = New clsRegistry
50030  reg.hkey = hkey1
50040  reg.KeyRoot = KeyRoot
50050  With myOptions
50060   reg.Subkey = "Ghostscript"
50070   tStr = reg.GetRegistryValue("DirectoryGhostscriptBinaries")
50080   If LenB(Trim$(tStr)) > 0 Then
50090     .DirectoryGhostscriptBinaries = CompletePath(tStr)
50100    Else
50110     If UseStandard Then
50120      tStr = GetPDFCreatorApplicationPath
50130      .DirectoryGhostscriptBinaries = CompletePath(tStr)
50140     End If
50150   End If
50160   tStr = reg.GetRegistryValue("DirectoryGhostscriptFonts")
50170   If LenB(Trim$(tStr)) > 0 Then
50180     .DirectoryGhostscriptFonts = CompletePath(tStr)
50190    Else
50200     If UseStandard Then
50210      tStr = GetPDFCreatorApplicationPath & "fonts"
50220      .DirectoryGhostscriptFonts = CompletePath(tStr)
50230     End If
50240   End If
50250   tStr = reg.GetRegistryValue("DirectoryGhostscriptLibraries")
50260   If LenB(Trim$(tStr)) > 0 Then
50270     .DirectoryGhostscriptLibraries = CompletePath(tStr)
50280    Else
50290     If UseStandard Then
50300      tStr = GetPDFCreatorApplicationPath & "lib"
50310      .DirectoryGhostscriptLibraries = CompletePath(tStr)
50320     End If
50330   End If
50340   tStr = reg.GetRegistryValue("DirectoryGhostscriptResource")
50350   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
50360     .DirectoryGhostscriptResource = ""
50370    Else
50380     If LenB(tStr) > 0 Then
50390      .DirectoryGhostscriptResource = tStr
50400     End If
50410   End If
50420   reg.Subkey = "Printing"
50430   tStr = reg.GetRegistryValue("Counter")
50440   If IsNumeric(tStr) Then
50450     If CCur(tStr) >= 0 And CCur(tStr) <= 922337203685477# Then
50460       .Counter = CCur(tStr)
50470      Else
50480       If UseStandard Then
50490        .Counter = 0
50500       End If
50510     End If
50520    Else
50530     If UseStandard Then
50540      .Counter = 0
50550     End If
50560   End If
50570   tStr = reg.GetRegistryValue("DeviceHeightPoints")
50580   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
50590     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 1 Then
50600       .DeviceHeightPoints = CDbl(Replace$(tStr, ".", GetDecimalChar))
50610      Else
50620       If UseStandard Then
50630        .DeviceHeightPoints = Replace$("842", ".", GetDecimalChar)
50640       End If
50650     End If
50660    Else
50670     If UseStandard Then
50680      .DeviceHeightPoints = Replace$("842", ".", GetDecimalChar)
50690     End If
50700   End If
50710   tStr = reg.GetRegistryValue("DeviceWidthPoints")
50720   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
50730     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 1 Then
50740       .DeviceWidthPoints = CDbl(Replace$(tStr, ".", GetDecimalChar))
50750      Else
50760       If UseStandard Then
50770        .DeviceWidthPoints = Replace$("595", ".", GetDecimalChar)
50780       End If
50790     End If
50800    Else
50810     If UseStandard Then
50820      .DeviceWidthPoints = Replace$("595", ".", GetDecimalChar)
50830     End If
50840   End If
50850   tStr = reg.GetRegistryValue("OnePagePerFile")
50860   If IsNumeric(tStr) Then
50870     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
50880       .OnePagePerFile = CLng(tStr)
50890      Else
50900       If UseStandard Then
50910        .OnePagePerFile = 0
50920       End If
50930     End If
50940    Else
50950     If UseStandard Then
50960      .OnePagePerFile = 0
50970     End If
50980   End If
50990   tStr = reg.GetRegistryValue("Papersize")
51000   If LenB(tStr) = 0 And LenB("a4") > 0 And UseStandard Then
51010     .Papersize = "a4"
51020    Else
51030     If LenB(tStr) > 0 Then
51040      .Papersize = tStr
51050     End If
51060   End If
51070   tStr = reg.GetRegistryValue("StampFontColor")
51080   If LenB(tStr) = 0 And LenB("#FF0000") > 0 And UseStandard Then
51090     .StampFontColor = "#FF0000"
51100    Else
51110     If LenB(tStr) > 0 Then
51120      .StampFontColor = tStr
51130     End If
51140   End If
51150   tStr = reg.GetRegistryValue("StampFontname")
51160   If LenB(tStr) = 0 And LenB("Arial") > 0 And UseStandard Then
51170     .StampFontname = "Arial"
51180    Else
51190     If LenB(tStr) > 0 Then
51200      .StampFontname = tStr
51210     End If
51220   End If
51230   tStr = reg.GetRegistryValue("StampFontsize")
51240   If IsNumeric(tStr) Then
51250     If CLng(tStr) >= 1 Then
51260       .StampFontsize = CLng(tStr)
51270      Else
51280       If UseStandard Then
51290        .StampFontsize = 48
51300       End If
51310     End If
51320    Else
51330     If UseStandard Then
51340      .StampFontsize = 48
51350     End If
51360   End If
51370   tStr = reg.GetRegistryValue("StampOutlineFontthickness")
51380   If IsNumeric(tStr) Then
51390     If CLng(tStr) >= 0 Then
51400       .StampOutlineFontthickness = CLng(tStr)
51410      Else
51420       If UseStandard Then
51430        .StampOutlineFontthickness = 0
51440       End If
51450     End If
51460    Else
51470     If UseStandard Then
51480      .StampOutlineFontthickness = 0
51490     End If
51500   End If
51510   tStr = reg.GetRegistryValue("StampString")
51520   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
51530     .StampString = ""
51540    Else
51550     If LenB(tStr) > 0 Then
51560      .StampString = tStr
51570     End If
51580   End If
51590   tStr = reg.GetRegistryValue("StampUseOutlineFont")
51600   If IsNumeric(tStr) Then
51610     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51620       .StampUseOutlineFont = CLng(tStr)
51630      Else
51640       If UseStandard Then
51650        .StampUseOutlineFont = 1
51660       End If
51670     End If
51680    Else
51690     If UseStandard Then
51700      .StampUseOutlineFont = 1
51710     End If
51720   End If
51730   tStr = reg.GetRegistryValue("StandardAuthor")
51740   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
51750     .StandardAuthor = ""
51760    Else
51770     If LenB(tStr) > 0 Then
51780      .StandardAuthor = tStr
51790     End If
51800   End If
51810   tStr = reg.GetRegistryValue("StandardCreationdate")
51820   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
51830     .StandardCreationdate = ""
51840    Else
51850     If LenB(tStr) > 0 Then
51860      .StandardCreationdate = tStr
51870     End If
51880   End If
51890   tStr = reg.GetRegistryValue("StandardDateformat")
51900   If LenB(tStr) = 0 And LenB("YYYYMMDDHHNNSS") > 0 And UseStandard Then
51910     .StandardDateformat = "YYYYMMDDHHNNSS"
51920    Else
51930     If LenB(tStr) > 0 Then
51940      .StandardDateformat = tStr
51950     End If
51960   End If
51970   tStr = reg.GetRegistryValue("StandardKeywords")
51980   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
51990     .StandardKeywords = ""
52000    Else
52010     If LenB(tStr) > 0 Then
52020      .StandardKeywords = tStr
52030     End If
52040   End If
52050   tStr = reg.GetRegistryValue("StandardMailDomain")
52060   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
52070     .StandardMailDomain = ""
52080    Else
52090     If LenB(tStr) > 0 Then
52100      .StandardMailDomain = tStr
52110     End If
52120   End If
52130   tStr = reg.GetRegistryValue("StandardModifydate")
52140   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
52150     .StandardModifydate = ""
52160    Else
52170     If LenB(tStr) > 0 Then
52180      .StandardModifydate = tStr
52190     End If
52200   End If
52210   tStr = reg.GetRegistryValue("StandardSaveformat")
52220   If IsNumeric(tStr) Then
52230     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
52240       .StandardSaveformat = CLng(tStr)
52250      Else
52260       If UseStandard Then
52270        .StandardSaveformat = 0
52280       End If
52290     End If
52300    Else
52310     If UseStandard Then
52320      .StandardSaveformat = 0
52330     End If
52340   End If
52350   tStr = reg.GetRegistryValue("StandardSubject")
52360   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
52370     .StandardSubject = ""
52380    Else
52390     If LenB(tStr) > 0 Then
52400      .StandardSubject = tStr
52410     End If
52420   End If
52430   tStr = reg.GetRegistryValue("StandardTitle")
52440   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
52450     .StandardTitle = ""
52460    Else
52470     If LenB(tStr) > 0 Then
52480      .StandardTitle = tStr
52490     End If
52500   End If
52510   tStr = reg.GetRegistryValue("UseCreationDateNow")
52520   If IsNumeric(tStr) Then
52530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52540       .UseCreationDateNow = CLng(tStr)
52550      Else
52560       If UseStandard Then
52570        .UseCreationDateNow = 0
52580       End If
52590     End If
52600    Else
52610     If UseStandard Then
52620      .UseCreationDateNow = 0
52630     End If
52640   End If
52650   tStr = reg.GetRegistryValue("UseCustomPaperSize")
52660   If LenB(tStr) = 0 And LenB("0") > 0 And UseStandard Then
52670     .UseCustomPaperSize = "0"
52680    Else
52690     If LenB(tStr) > 0 Then
52700      .UseCustomPaperSize = tStr
52710     End If
52720   End If
52730   tStr = reg.GetRegistryValue("UseFixPapersize")
52740   If IsNumeric(tStr) Then
52750     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52760       .UseFixPapersize = CLng(tStr)
52770      Else
52780       If UseStandard Then
52790        .UseFixPapersize = 0
52800       End If
52810     End If
52820    Else
52830     If UseStandard Then
52840      .UseFixPapersize = 0
52850     End If
52860   End If
52870   tStr = reg.GetRegistryValue("UseStandardAuthor")
52880   If IsNumeric(tStr) Then
52890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52900       .UseStandardAuthor = CLng(tStr)
52910      Else
52920       If UseStandard Then
52930        .UseStandardAuthor = 0
52940       End If
52950     End If
52960    Else
52970     If UseStandard Then
52980      .UseStandardAuthor = 0
52990     End If
53000   End If
53010   reg.Subkey = "Printing\Formats\Bitmap\Colors"
53020   tStr = reg.GetRegistryValue("BMPColorscount")
53030   If IsNumeric(tStr) Then
53040     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
53050       .BMPColorscount = CLng(tStr)
53060      Else
53070       If UseStandard Then
53080        .BMPColorscount = 1
53090       End If
53100     End If
53110    Else
53120     If UseStandard Then
53130      .BMPColorscount = 1
53140     End If
53150   End If
53160   tStr = reg.GetRegistryValue("BMPResolution")
53170   If IsNumeric(tStr) Then
53180     If CLng(tStr) >= 1 Then
53190       .BMPResolution = CLng(tStr)
53200      Else
53210       If UseStandard Then
53220        .BMPResolution = 150
53230       End If
53240     End If
53250    Else
53260     If UseStandard Then
53270      .BMPResolution = 150
53280     End If
53290   End If
53300   tStr = reg.GetRegistryValue("JPEGColorscount")
53310   If IsNumeric(tStr) Then
53320     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
53330       .JPEGColorscount = CLng(tStr)
53340      Else
53350       If UseStandard Then
53360        .JPEGColorscount = 0
53370       End If
53380     End If
53390    Else
53400     If UseStandard Then
53410      .JPEGColorscount = 0
53420     End If
53430   End If
53440   tStr = reg.GetRegistryValue("JPEGQuality")
53450   If IsNumeric(tStr) Then
53460     If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
53470       .JPEGQuality = CLng(tStr)
53480      Else
53490       If UseStandard Then
53500        .JPEGQuality = 75
53510       End If
53520     End If
53530    Else
53540     If UseStandard Then
53550      .JPEGQuality = 75
53560     End If
53570   End If
53580   tStr = reg.GetRegistryValue("JPEGResolution")
53590   If IsNumeric(tStr) Then
53600     If CLng(tStr) >= 1 Then
53610       .JPEGResolution = CLng(tStr)
53620      Else
53630       If UseStandard Then
53640        .JPEGResolution = 150
53650       End If
53660     End If
53670    Else
53680     If UseStandard Then
53690      .JPEGResolution = 150
53700     End If
53710   End If
53720   tStr = reg.GetRegistryValue("PCLColorsCount")
53730   If IsNumeric(tStr) Then
53740     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
53750       .PCLColorsCount = CLng(tStr)
53760      Else
53770       If UseStandard Then
53780        .PCLColorsCount = 0
53790       End If
53800     End If
53810    Else
53820     If UseStandard Then
53830      .PCLColorsCount = 0
53840     End If
53850   End If
53860   tStr = reg.GetRegistryValue("PCLResolution")
53870   If IsNumeric(tStr) Then
53880     If CLng(tStr) >= 1 Then
53890       .PCLResolution = CLng(tStr)
53900      Else
53910       If UseStandard Then
53920        .PCLResolution = 150
53930       End If
53940     End If
53950    Else
53960     If UseStandard Then
53970      .PCLResolution = 150
53980     End If
53990   End If
54000   tStr = reg.GetRegistryValue("PCXColorscount")
54010   If IsNumeric(tStr) Then
54020     If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
54030       .PCXColorscount = CLng(tStr)
54040      Else
54050       If UseStandard Then
54060        .PCXColorscount = 0
54070       End If
54080     End If
54090    Else
54100     If UseStandard Then
54110      .PCXColorscount = 0
54120     End If
54130   End If
54140   tStr = reg.GetRegistryValue("PCXResolution")
54150   If IsNumeric(tStr) Then
54160     If CLng(tStr) >= 1 Then
54170       .PCXResolution = CLng(tStr)
54180      Else
54190       If UseStandard Then
54200        .PCXResolution = 150
54210       End If
54220     End If
54230    Else
54240     If UseStandard Then
54250      .PCXResolution = 150
54260     End If
54270   End If
54280   tStr = reg.GetRegistryValue("PNGColorscount")
54290   If IsNumeric(tStr) Then
54300     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
54310       .PNGColorscount = CLng(tStr)
54320      Else
54330       If UseStandard Then
54340        .PNGColorscount = 0
54350       End If
54360     End If
54370    Else
54380     If UseStandard Then
54390      .PNGColorscount = 0
54400     End If
54410   End If
54420   tStr = reg.GetRegistryValue("PNGResolution")
54430   If IsNumeric(tStr) Then
54440     If CLng(tStr) >= 1 Then
54450       .PNGResolution = CLng(tStr)
54460      Else
54470       If UseStandard Then
54480        .PNGResolution = 150
54490       End If
54500     End If
54510    Else
54520     If UseStandard Then
54530      .PNGResolution = 150
54540     End If
54550   End If
54560   tStr = reg.GetRegistryValue("PSDColorsCount")
54570   If IsNumeric(tStr) Then
54580     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
54590       .PSDColorsCount = CLng(tStr)
54600      Else
54610       If UseStandard Then
54620        .PSDColorsCount = 0
54630       End If
54640     End If
54650    Else
54660     If UseStandard Then
54670      .PSDColorsCount = 0
54680     End If
54690   End If
54700   tStr = reg.GetRegistryValue("PSDResolution")
54710   If IsNumeric(tStr) Then
54720     If CLng(tStr) >= 1 Then
54730       .PSDResolution = CLng(tStr)
54740      Else
54750       If UseStandard Then
54760        .PSDResolution = 150
54770       End If
54780     End If
54790    Else
54800     If UseStandard Then
54810      .PSDResolution = 150
54820     End If
54830   End If
54840   tStr = reg.GetRegistryValue("RAWColorsCount")
54850   If IsNumeric(tStr) Then
54860     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
54870       .RAWColorsCount = CLng(tStr)
54880      Else
54890       If UseStandard Then
54900        .RAWColorsCount = 0
54910       End If
54920     End If
54930    Else
54940     If UseStandard Then
54950      .RAWColorsCount = 0
54960     End If
54970   End If
54980   tStr = reg.GetRegistryValue("RAWResolution")
54990   If IsNumeric(tStr) Then
55000     If CLng(tStr) >= 1 Then
55010       .RAWResolution = CLng(tStr)
55020      Else
55030       If UseStandard Then
55040        .RAWResolution = 150
55050       End If
55060     End If
55070    Else
55080     If UseStandard Then
55090      .RAWResolution = 150
55100     End If
55110   End If
55120   tStr = reg.GetRegistryValue("TIFFColorscount")
55130   If IsNumeric(tStr) Then
55140     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
55150       .TIFFColorscount = CLng(tStr)
55160      Else
55170       If UseStandard Then
55180        .TIFFColorscount = 0
55190       End If
55200     End If
55210    Else
55220     If UseStandard Then
55230      .TIFFColorscount = 0
55240     End If
55250   End If
55260   tStr = reg.GetRegistryValue("TIFFResolution")
55270   If IsNumeric(tStr) Then
55280     If CLng(tStr) >= 1 Then
55290       .TIFFResolution = CLng(tStr)
55300      Else
55310       If UseStandard Then
55320        .TIFFResolution = 150
55330       End If
55340     End If
55350    Else
55360     If UseStandard Then
55370      .TIFFResolution = 150
55380     End If
55390   End If
55400   reg.Subkey = "Printing\Formats\PDF\Colors"
55410   tStr = reg.GetRegistryValue("PDFColorsCMYKToRGB")
55420   If IsNumeric(tStr) Then
55430     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55440       .PDFColorsCMYKToRGB = CLng(tStr)
55450      Else
55460       If UseStandard Then
55470        .PDFColorsCMYKToRGB = 0
55480       End If
55490     End If
55500    Else
55510     If UseStandard Then
55520      .PDFColorsCMYKToRGB = 0
55530     End If
55540   End If
55550   tStr = reg.GetRegistryValue("PDFColorsColorModel")
55560   If IsNumeric(tStr) Then
55570     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
55580       .PDFColorsColorModel = CLng(tStr)
55590      Else
55600       If UseStandard Then
55610        .PDFColorsColorModel = 1
55620       End If
55630     End If
55640    Else
55650     If UseStandard Then
55660      .PDFColorsColorModel = 1
55670     End If
55680   End If
55690   tStr = reg.GetRegistryValue("PDFColorsPreserveHalftone")
55700   If IsNumeric(tStr) Then
55710     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55720       .PDFColorsPreserveHalftone = CLng(tStr)
55730      Else
55740       If UseStandard Then
55750        .PDFColorsPreserveHalftone = 0
55760       End If
55770     End If
55780    Else
55790     If UseStandard Then
55800      .PDFColorsPreserveHalftone = 0
55810     End If
55820   End If
55830   tStr = reg.GetRegistryValue("PDFColorsPreserveOverprint")
55840   If IsNumeric(tStr) Then
55850     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55860       .PDFColorsPreserveOverprint = CLng(tStr)
55870      Else
55880       If UseStandard Then
55890        .PDFColorsPreserveOverprint = 1
55900       End If
55910     End If
55920    Else
55930     If UseStandard Then
55940      .PDFColorsPreserveOverprint = 1
55950     End If
55960   End If
55970   tStr = reg.GetRegistryValue("PDFColorsPreserveTransfer")
55980   If IsNumeric(tStr) Then
55990     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56000       .PDFColorsPreserveTransfer = CLng(tStr)
56010      Else
56020       If UseStandard Then
56030        .PDFColorsPreserveTransfer = 1
56040       End If
56050     End If
56060    Else
56070     If UseStandard Then
56080      .PDFColorsPreserveTransfer = 1
56090     End If
56100   End If
56110   reg.Subkey = "Printing\Formats\PDF\Compression"
56120   tStr = reg.GetRegistryValue("PDFCompressionColorCompression")
56130   If IsNumeric(tStr) Then
56140     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56150       .PDFCompressionColorCompression = CLng(tStr)
56160      Else
56170       If UseStandard Then
56180        .PDFCompressionColorCompression = 1
56190       End If
56200     End If
56210    Else
56220     If UseStandard Then
56230      .PDFCompressionColorCompression = 1
56240     End If
56250   End If
56260   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionChoice")
56270   If IsNumeric(tStr) Then
56280     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
56290       .PDFCompressionColorCompressionChoice = CLng(tStr)
56300      Else
56310       If UseStandard Then
56320        .PDFCompressionColorCompressionChoice = 0
56330       End If
56340     End If
56350    Else
56360     If UseStandard Then
56370      .PDFCompressionColorCompressionChoice = 0
56380     End If
56390   End If
56400   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGHighFactor")
56410   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56420     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56430       .PDFCompressionColorCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56440      Else
56450       If UseStandard Then
56460        .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56470       End If
56480     End If
56490    Else
56500     If UseStandard Then
56510      .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56520     End If
56530   End If
56540   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGLowFactor")
56550   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56560     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56570       .PDFCompressionColorCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56580      Else
56590       If UseStandard Then
56600        .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56610       End If
56620     End If
56630    Else
56640     If UseStandard Then
56650      .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56660     End If
56670   End If
56680   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMaximumFactor")
56690   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56700     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56710       .PDFCompressionColorCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56720      Else
56730       If UseStandard Then
56740        .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56750       End If
56760     End If
56770    Else
56780     If UseStandard Then
56790      .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56800     End If
56810   End If
56820   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMediumFactor")
56830   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56840     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56850       .PDFCompressionColorCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56860      Else
56870       If UseStandard Then
56880        .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56890       End If
56900     End If
56910    Else
56920     If UseStandard Then
56930      .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56940     End If
56950   End If
56960   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMinimumFactor")
56970   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56980     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56990       .PDFCompressionColorCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57000      Else
57010       If UseStandard Then
57020        .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57030       End If
57040     End If
57050    Else
57060     If UseStandard Then
57070      .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57080     End If
57090   End If
57100   tStr = reg.GetRegistryValue("PDFCompressionColorResample")
57110   If IsNumeric(tStr) Then
57120     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57130       .PDFCompressionColorResample = CLng(tStr)
57140      Else
57150       If UseStandard Then
57160        .PDFCompressionColorResample = 0
57170       End If
57180     End If
57190    Else
57200     If UseStandard Then
57210      .PDFCompressionColorResample = 0
57220     End If
57230   End If
57240   tStr = reg.GetRegistryValue("PDFCompressionColorResampleChoice")
57250   If IsNumeric(tStr) Then
57260     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
57270       .PDFCompressionColorResampleChoice = CLng(tStr)
57280      Else
57290       If UseStandard Then
57300        .PDFCompressionColorResampleChoice = 0
57310       End If
57320     End If
57330    Else
57340     If UseStandard Then
57350      .PDFCompressionColorResampleChoice = 0
57360     End If
57370   End If
57380   tStr = reg.GetRegistryValue("PDFCompressionColorResolution")
57390   If IsNumeric(tStr) Then
57400     If CLng(tStr) >= 0 Then
57410       .PDFCompressionColorResolution = CLng(tStr)
57420      Else
57430       If UseStandard Then
57440        .PDFCompressionColorResolution = 300
57450       End If
57460     End If
57470    Else
57480     If UseStandard Then
57490      .PDFCompressionColorResolution = 300
57500     End If
57510   End If
57520   tStr = reg.GetRegistryValue("PDFCompressionGreyCompression")
57530   If IsNumeric(tStr) Then
57540     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57550       .PDFCompressionGreyCompression = CLng(tStr)
57560      Else
57570       If UseStandard Then
57580        .PDFCompressionGreyCompression = 1
57590       End If
57600     End If
57610    Else
57620     If UseStandard Then
57630      .PDFCompressionGreyCompression = 1
57640     End If
57650   End If
57660   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionChoice")
57670   If IsNumeric(tStr) Then
57680     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
57690       .PDFCompressionGreyCompressionChoice = CLng(tStr)
57700      Else
57710       If UseStandard Then
57720        .PDFCompressionGreyCompressionChoice = 0
57730       End If
57740     End If
57750    Else
57760     If UseStandard Then
57770      .PDFCompressionGreyCompressionChoice = 0
57780     End If
57790   End If
57800   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGHighFactor")
57810   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57820     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57830       .PDFCompressionGreyCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57840      Else
57850       If UseStandard Then
57860        .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
57870       End If
57880     End If
57890    Else
57900     If UseStandard Then
57910      .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
57920     End If
57930   End If
57940   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGLowFactor")
57950   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57960     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57970       .PDFCompressionGreyCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57980      Else
57990       If UseStandard Then
58000        .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58010       End If
58020     End If
58030    Else
58040     If UseStandard Then
58050      .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58060     End If
58070   End If
58080   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMaximumFactor")
58090   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58100     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58110       .PDFCompressionGreyCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58120      Else
58130       If UseStandard Then
58140        .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
58150       End If
58160     End If
58170    Else
58180     If UseStandard Then
58190      .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
58200     End If
58210   End If
58220   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMediumFactor")
58230   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58240     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58250       .PDFCompressionGreyCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58260      Else
58270       If UseStandard Then
58280        .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58290       End If
58300     End If
58310    Else
58320     If UseStandard Then
58330      .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58340     End If
58350   End If
58360   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMinimumFactor")
58370   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58380     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58390       .PDFCompressionGreyCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58400      Else
58410       If UseStandard Then
58420        .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58430       End If
58440     End If
58450    Else
58460     If UseStandard Then
58470      .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58480     End If
58490   End If
58500   tStr = reg.GetRegistryValue("PDFCompressionGreyResample")
58510   If IsNumeric(tStr) Then
58520     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58530       .PDFCompressionGreyResample = CLng(tStr)
58540      Else
58550       If UseStandard Then
58560        .PDFCompressionGreyResample = 0
58570       End If
58580     End If
58590    Else
58600     If UseStandard Then
58610      .PDFCompressionGreyResample = 0
58620     End If
58630   End If
58640   tStr = reg.GetRegistryValue("PDFCompressionGreyResampleChoice")
58650   If IsNumeric(tStr) Then
58660     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
58670       .PDFCompressionGreyResampleChoice = CLng(tStr)
58680      Else
58690       If UseStandard Then
58700        .PDFCompressionGreyResampleChoice = 0
58710       End If
58720     End If
58730    Else
58740     If UseStandard Then
58750      .PDFCompressionGreyResampleChoice = 0
58760     End If
58770   End If
58780   tStr = reg.GetRegistryValue("PDFCompressionGreyResolution")
58790   If IsNumeric(tStr) Then
58800     If CLng(tStr) >= 0 Then
58810       .PDFCompressionGreyResolution = CLng(tStr)
58820      Else
58830       If UseStandard Then
58840        .PDFCompressionGreyResolution = 300
58850       End If
58860     End If
58870    Else
58880     If UseStandard Then
58890      .PDFCompressionGreyResolution = 300
58900     End If
58910   End If
58920   tStr = reg.GetRegistryValue("PDFCompressionMonoCompression")
58930   If IsNumeric(tStr) Then
58940     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58950       .PDFCompressionMonoCompression = CLng(tStr)
58960      Else
58970       If UseStandard Then
58980        .PDFCompressionMonoCompression = 1
58990       End If
59000     End If
59010    Else
59020     If UseStandard Then
59030      .PDFCompressionMonoCompression = 1
59040     End If
59050   End If
59060   tStr = reg.GetRegistryValue("PDFCompressionMonoCompressionChoice")
59070   If IsNumeric(tStr) Then
59080     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
59090       .PDFCompressionMonoCompressionChoice = CLng(tStr)
59100      Else
59110       If UseStandard Then
59120        .PDFCompressionMonoCompressionChoice = 0
59130       End If
59140     End If
59150    Else
59160     If UseStandard Then
59170      .PDFCompressionMonoCompressionChoice = 0
59180     End If
59190   End If
59200   tStr = reg.GetRegistryValue("PDFCompressionMonoResample")
59210   If IsNumeric(tStr) Then
59220     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59230       .PDFCompressionMonoResample = CLng(tStr)
59240      Else
59250       If UseStandard Then
59260        .PDFCompressionMonoResample = 0
59270       End If
59280     End If
59290    Else
59300     If UseStandard Then
59310      .PDFCompressionMonoResample = 0
59320     End If
59330   End If
59340   tStr = reg.GetRegistryValue("PDFCompressionMonoResampleChoice")
59350   If IsNumeric(tStr) Then
59360     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
59370       .PDFCompressionMonoResampleChoice = CLng(tStr)
59380      Else
59390       If UseStandard Then
59400        .PDFCompressionMonoResampleChoice = 0
59410       End If
59420     End If
59430    Else
59440     If UseStandard Then
59450      .PDFCompressionMonoResampleChoice = 0
59460     End If
59470   End If
59480   tStr = reg.GetRegistryValue("PDFCompressionMonoResolution")
59490   If IsNumeric(tStr) Then
59500     If CLng(tStr) >= 0 Then
59510       .PDFCompressionMonoResolution = CLng(tStr)
59520      Else
59530       If UseStandard Then
59540        .PDFCompressionMonoResolution = 1200
59550       End If
59560     End If
59570    Else
59580     If UseStandard Then
59590      .PDFCompressionMonoResolution = 1200
59600     End If
59610   End If
59620   tStr = reg.GetRegistryValue("PDFCompressionTextCompression")
59630   If IsNumeric(tStr) Then
59640     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59650       .PDFCompressionTextCompression = CLng(tStr)
59660      Else
59670       If UseStandard Then
59680        .PDFCompressionTextCompression = 1
59690       End If
59700     End If
59710    Else
59720     If UseStandard Then
59730      .PDFCompressionTextCompression = 1
59740     End If
59750   End If
59760   reg.Subkey = "Printing\Formats\PDF\Fonts"
59770   tStr = reg.GetRegistryValue("PDFFontsEmbedAll")
59780   If IsNumeric(tStr) Then
59790     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59800       .PDFFontsEmbedAll = CLng(tStr)
59810      Else
59820       If UseStandard Then
59830        .PDFFontsEmbedAll = 1
59840       End If
59850     End If
59860    Else
59870     If UseStandard Then
59880      .PDFFontsEmbedAll = 1
59890     End If
59900   End If
59910   tStr = reg.GetRegistryValue("PDFFontsSubSetFonts")
59920   If IsNumeric(tStr) Then
59930     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59940       .PDFFontsSubSetFonts = CLng(tStr)
59950      Else
59960       If UseStandard Then
59970        .PDFFontsSubSetFonts = 1
59980       End If
59990     End If
60000    Else
60010     If UseStandard Then
60020      .PDFFontsSubSetFonts = 1
60030     End If
60040   End If
60050   tStr = reg.GetRegistryValue("PDFFontsSubSetFontsPercent")
60060   If IsNumeric(tStr) Then
60070     If CLng(tStr) >= 0 Then
60080       .PDFFontsSubSetFontsPercent = CLng(tStr)
60090      Else
60100       If UseStandard Then
60110        .PDFFontsSubSetFontsPercent = 100
60120       End If
60130     End If
60140    Else
60150     If UseStandard Then
60160      .PDFFontsSubSetFontsPercent = 100
60170     End If
60180   End If
60190   reg.Subkey = "Printing\Formats\PDF\General"
60200   tStr = reg.GetRegistryValue("PDFGeneralASCII85")
60210   If IsNumeric(tStr) Then
60220     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60230       .PDFGeneralASCII85 = CLng(tStr)
60240      Else
60250       If UseStandard Then
60260        .PDFGeneralASCII85 = 0
60270       End If
60280     End If
60290    Else
60300     If UseStandard Then
60310      .PDFGeneralASCII85 = 0
60320     End If
60330   End If
60340   tStr = reg.GetRegistryValue("PDFGeneralAutorotate")
60350   If IsNumeric(tStr) Then
60360     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
60370       .PDFGeneralAutorotate = CLng(tStr)
60380      Else
60390       If UseStandard Then
60400        .PDFGeneralAutorotate = 2
60410       End If
60420     End If
60430    Else
60440     If UseStandard Then
60450      .PDFGeneralAutorotate = 2
60460     End If
60470   End If
60480   tStr = reg.GetRegistryValue("PDFGeneralCompatibility")
60490   If IsNumeric(tStr) Then
60500     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
60510       .PDFGeneralCompatibility = CLng(tStr)
60520      Else
60530       If UseStandard Then
60540        .PDFGeneralCompatibility = 2
60550       End If
60560     End If
60570    Else
60580     If UseStandard Then
60590      .PDFGeneralCompatibility = 2
60600     End If
60610   End If
60620   tStr = reg.GetRegistryValue("PDFGeneralDefault")
60630   If IsNumeric(tStr) Then
60640     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
60650       .PDFGeneralDefault = CLng(tStr)
60660      Else
60670       If UseStandard Then
60680        .PDFGeneralDefault = 0
60690       End If
60700     End If
60710    Else
60720     If UseStandard Then
60730      .PDFGeneralDefault = 0
60740     End If
60750   End If
60760   tStr = reg.GetRegistryValue("PDFGeneralOverprint")
60770   If IsNumeric(tStr) Then
60780     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
60790       .PDFGeneralOverprint = CLng(tStr)
60800      Else
60810       If UseStandard Then
60820        .PDFGeneralOverprint = 0
60830       End If
60840     End If
60850    Else
60860     If UseStandard Then
60870      .PDFGeneralOverprint = 0
60880     End If
60890   End If
60900   tStr = reg.GetRegistryValue("PDFGeneralResolution")
60910   If IsNumeric(tStr) Then
60920     If CLng(tStr) >= 0 Then
60930       .PDFGeneralResolution = CLng(tStr)
60940      Else
60950       If UseStandard Then
60960        .PDFGeneralResolution = 600
60970       End If
60980     End If
60990    Else
61000     If UseStandard Then
61010      .PDFGeneralResolution = 600
61020     End If
61030   End If
61040   tStr = reg.GetRegistryValue("PDFOptimize")
61050   If IsNumeric(tStr) Then
61060     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61070       .PDFOptimize = CLng(tStr)
61080      Else
61090       If UseStandard Then
61100        .PDFOptimize = 0
61110       End If
61120     End If
61130    Else
61140     If UseStandard Then
61150      .PDFOptimize = 0
61160     End If
61170   End If
61180   tStr = reg.GetRegistryValue("PDFUpdateMetadata")
61190   If IsNumeric(tStr) Then
61200     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
61210       .PDFUpdateMetadata = CLng(tStr)
61220      Else
61230       If UseStandard Then
61240        .PDFUpdateMetadata = 1
61250       End If
61260     End If
61270    Else
61280     If UseStandard Then
61290      .PDFUpdateMetadata = 1
61300     End If
61310   End If
61320   reg.Subkey = "Printing\Formats\PDF\Security"
61330   tStr = reg.GetRegistryValue("PDFAllowAssembly")
61340   If IsNumeric(tStr) Then
61350     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61360       .PDFAllowAssembly = CLng(tStr)
61370      Else
61380       If UseStandard Then
61390        .PDFAllowAssembly = 0
61400       End If
61410     End If
61420    Else
61430     If UseStandard Then
61440      .PDFAllowAssembly = 0
61450     End If
61460   End If
61470   tStr = reg.GetRegistryValue("PDFAllowDegradedPrinting")
61480   If IsNumeric(tStr) Then
61490     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61500       .PDFAllowDegradedPrinting = CLng(tStr)
61510      Else
61520       If UseStandard Then
61530        .PDFAllowDegradedPrinting = 0
61540       End If
61550     End If
61560    Else
61570     If UseStandard Then
61580      .PDFAllowDegradedPrinting = 0
61590     End If
61600   End If
61610   tStr = reg.GetRegistryValue("PDFAllowFillIn")
61620   If IsNumeric(tStr) Then
61630     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61640       .PDFAllowFillIn = CLng(tStr)
61650      Else
61660       If UseStandard Then
61670        .PDFAllowFillIn = 0
61680       End If
61690     End If
61700    Else
61710     If UseStandard Then
61720      .PDFAllowFillIn = 0
61730     End If
61740   End If
61750   tStr = reg.GetRegistryValue("PDFAllowScreenReaders")
61760   If IsNumeric(tStr) Then
61770     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61780       .PDFAllowScreenReaders = CLng(tStr)
61790      Else
61800       If UseStandard Then
61810        .PDFAllowScreenReaders = 0
61820       End If
61830     End If
61840    Else
61850     If UseStandard Then
61860      .PDFAllowScreenReaders = 0
61870     End If
61880   End If
61890   tStr = reg.GetRegistryValue("PDFDisallowCopy")
61900   If IsNumeric(tStr) Then
61910     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61920       .PDFDisallowCopy = CLng(tStr)
61930      Else
61940       If UseStandard Then
61950        .PDFDisallowCopy = 1
61960       End If
61970     End If
61980    Else
61990     If UseStandard Then
62000      .PDFDisallowCopy = 1
62010     End If
62020   End If
62030   tStr = reg.GetRegistryValue("PDFDisallowModifyAnnotations")
62040   If IsNumeric(tStr) Then
62050     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62060       .PDFDisallowModifyAnnotations = CLng(tStr)
62070      Else
62080       If UseStandard Then
62090        .PDFDisallowModifyAnnotations = 0
62100       End If
62110     End If
62120    Else
62130     If UseStandard Then
62140      .PDFDisallowModifyAnnotations = 0
62150     End If
62160   End If
62170   tStr = reg.GetRegistryValue("PDFDisallowModifyContents")
62180   If IsNumeric(tStr) Then
62190     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62200       .PDFDisallowModifyContents = CLng(tStr)
62210      Else
62220       If UseStandard Then
62230        .PDFDisallowModifyContents = 0
62240       End If
62250     End If
62260    Else
62270     If UseStandard Then
62280      .PDFDisallowModifyContents = 0
62290     End If
62300   End If
62310   tStr = reg.GetRegistryValue("PDFDisallowPrinting")
62320   If IsNumeric(tStr) Then
62330     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62340       .PDFDisallowPrinting = CLng(tStr)
62350      Else
62360       If UseStandard Then
62370        .PDFDisallowPrinting = 0
62380       End If
62390     End If
62400    Else
62410     If UseStandard Then
62420      .PDFDisallowPrinting = 0
62430     End If
62440   End If
62450   tStr = reg.GetRegistryValue("PDFEncryptor")
62460   If IsNumeric(tStr) Then
62470     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
62480       .PDFEncryptor = CLng(tStr)
62490      Else
62500       If UseStandard Then
62510        .PDFEncryptor = 0
62520       End If
62530     End If
62540    Else
62550     If UseStandard Then
62560      .PDFEncryptor = 0
62570     End If
62580   End If
62590   tStr = reg.GetRegistryValue("PDFHighEncryption")
62600   If IsNumeric(tStr) Then
62610     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62620       .PDFHighEncryption = CLng(tStr)
62630      Else
62640       If UseStandard Then
62650        .PDFHighEncryption = 0
62660       End If
62670     End If
62680    Else
62690     If UseStandard Then
62700      .PDFHighEncryption = 0
62710     End If
62720   End If
62730   tStr = reg.GetRegistryValue("PDFLowEncryption")
62740   If IsNumeric(tStr) Then
62750     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62760       .PDFLowEncryption = CLng(tStr)
62770      Else
62780       If UseStandard Then
62790        .PDFLowEncryption = 1
62800       End If
62810     End If
62820    Else
62830     If UseStandard Then
62840      .PDFLowEncryption = 1
62850     End If
62860   End If
62870   tStr = reg.GetRegistryValue("PDFOwnerPass")
62880   If IsNumeric(tStr) Then
62890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62900       .PDFOwnerPass = CLng(tStr)
62910      Else
62920       If UseStandard Then
62930        .PDFOwnerPass = 0
62940       End If
62950     End If
62960    Else
62970     If UseStandard Then
62980      .PDFOwnerPass = 0
62990     End If
63000   End If
63010   tStr = reg.GetRegistryValue("PDFOwnerPasswordString")
63020   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63030     .PDFOwnerPasswordString = ""
63040    Else
63050     If LenB(tStr) > 0 Then
63060      .PDFOwnerPasswordString = tStr
63070     End If
63080   End If
63090   tStr = reg.GetRegistryValue("PDFUserPass")
63100   If IsNumeric(tStr) Then
63110     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63120       .PDFUserPass = CLng(tStr)
63130      Else
63140       If UseStandard Then
63150        .PDFUserPass = 0
63160       End If
63170     End If
63180    Else
63190     If UseStandard Then
63200      .PDFUserPass = 0
63210     End If
63220   End If
63230   tStr = reg.GetRegistryValue("PDFUserPasswordString")
63240   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63250     .PDFUserPasswordString = ""
63260    Else
63270     If LenB(tStr) > 0 Then
63280      .PDFUserPasswordString = tStr
63290     End If
63300   End If
63310   tStr = reg.GetRegistryValue("PDFUseSecurity")
63320   If IsNumeric(tStr) Then
63330     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63340       .PDFUseSecurity = CLng(tStr)
63350      Else
63360       If UseStandard Then
63370        .PDFUseSecurity = 0
63380       End If
63390     End If
63400    Else
63410     If UseStandard Then
63420      .PDFUseSecurity = 0
63430     End If
63440   End If
63450   reg.Subkey = "Printing\Formats\PDF\Signing"
63460   tStr = reg.GetRegistryValue("PDFSigningMultiSignature")
63470   If IsNumeric(tStr) Then
63480     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63490       .PDFSigningMultiSignature = CLng(tStr)
63500      Else
63510       If UseStandard Then
63520        .PDFSigningMultiSignature = 0
63530       End If
63540     End If
63550    Else
63560     If UseStandard Then
63570      .PDFSigningMultiSignature = 0
63580     End If
63590   End If
63600   tStr = reg.GetRegistryValue("PDFSigningPFXFile")
63610   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63620     .PDFSigningPFXFile = ""
63630    Else
63640     If LenB(tStr) > 0 Then
63650      .PDFSigningPFXFile = tStr
63660     End If
63670   End If
63680   tStr = reg.GetRegistryValue("PDFSigningPFXFilePassword")
63690   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63700     .PDFSigningPFXFilePassword = ""
63710    Else
63720     If LenB(tStr) > 0 Then
63730      .PDFSigningPFXFilePassword = tStr
63740     End If
63750   End If
63760   tStr = reg.GetRegistryValue("PDFSigningSignatureContact")
63770   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63780     .PDFSigningSignatureContact = ""
63790    Else
63800     If LenB(tStr) > 0 Then
63810      .PDFSigningSignatureContact = tStr
63820     End If
63830   End If
63840   tStr = reg.GetRegistryValue("PDFSigningSignatureLeftX")
63850   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63860     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63870       .PDFSigningSignatureLeftX = CDbl(Replace$(tStr, ".", GetDecimalChar))
63880      Else
63890       If UseStandard Then
63900        .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
63910       End If
63920     End If
63930    Else
63940     If UseStandard Then
63950      .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
63960     End If
63970   End If
63980   tStr = reg.GetRegistryValue("PDFSigningSignatureLeftY")
63990   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
64000     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
64010       .PDFSigningSignatureLeftY = CDbl(Replace$(tStr, ".", GetDecimalChar))
64020      Else
64030       If UseStandard Then
64040        .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
64050       End If
64060     End If
64070    Else
64080     If UseStandard Then
64090      .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
64100     End If
64110   End If
64120   tStr = reg.GetRegistryValue("PDFSigningSignatureLocation")
64130   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64140     .PDFSigningSignatureLocation = ""
64150    Else
64160     If LenB(tStr) > 0 Then
64170      .PDFSigningSignatureLocation = tStr
64180     End If
64190   End If
64200   tStr = reg.GetRegistryValue("PDFSigningSignatureReason")
64210   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64220     .PDFSigningSignatureReason = ""
64230    Else
64240     If LenB(tStr) > 0 Then
64250      .PDFSigningSignatureReason = tStr
64260     End If
64270   End If
64280   tStr = reg.GetRegistryValue("PDFSigningSignatureRightX")
64290   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
64300     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
64310       .PDFSigningSignatureRightX = CDbl(Replace$(tStr, ".", GetDecimalChar))
64320      Else
64330       If UseStandard Then
64340        .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
64350       End If
64360     End If
64370    Else
64380     If UseStandard Then
64390      .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
64400     End If
64410   End If
64420   tStr = reg.GetRegistryValue("PDFSigningSignatureRightY")
64430   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
64440     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
64450       .PDFSigningSignatureRightY = CDbl(Replace$(tStr, ".", GetDecimalChar))
64460      Else
64470       If UseStandard Then
64480        .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
64490       End If
64500     End If
64510    Else
64520     If UseStandard Then
64530      .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
64540     End If
64550   End If
64560   tStr = reg.GetRegistryValue("PDFSigningSignatureVisible")
64570   If IsNumeric(tStr) Then
64580     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64590       .PDFSigningSignatureVisible = CLng(tStr)
64600      Else
64610       If UseStandard Then
64620        .PDFSigningSignatureVisible = 0
64630       End If
64640     End If
64650    Else
64660     If UseStandard Then
64670      .PDFSigningSignatureVisible = 0
64680     End If
64690   End If
64700   tStr = reg.GetRegistryValue("PDFSigningSignPDF")
64710   If IsNumeric(tStr) Then
64720     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64730       .PDFSigningSignPDF = CLng(tStr)
64740      Else
64750       If UseStandard Then
64760        .PDFSigningSignPDF = 0
64770       End If
64780     End If
64790    Else
64800     If UseStandard Then
64810      .PDFSigningSignPDF = 0
64820     End If
64830   End If
64840   reg.Subkey = "Printing\Formats\PS\LanguageLevel"
64850   tStr = reg.GetRegistryValue("EPSLanguageLevel")
64860   If IsNumeric(tStr) Then
64870     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
64880       .EPSLanguageLevel = CLng(tStr)
64890      Else
64900       If UseStandard Then
64910        .EPSLanguageLevel = 2
64920       End If
64930     End If
64940    Else
64950     If UseStandard Then
64960      .EPSLanguageLevel = 2
64970     End If
64980   End If
64990   tStr = reg.GetRegistryValue("PSLanguageLevel")
65000   If IsNumeric(tStr) Then
65010     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
65020       .PSLanguageLevel = CLng(tStr)
65030      Else
65040       If UseStandard Then
65050        .PSLanguageLevel = 2
65060       End If
65070     End If
65080    Else
65090     If UseStandard Then
65100      .PSLanguageLevel = 2
65110     End If
65120   End If
65130   reg.Subkey = "Program"
65140   tStr = reg.GetRegistryValue("AdditionalGhostscriptParameters")
65150   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
65160     .AdditionalGhostscriptParameters = ""
65170    Else
65180     If LenB(tStr) > 0 Then
65190      .AdditionalGhostscriptParameters = tStr
65200     End If
65210   End If
65220   tStr = reg.GetRegistryValue("AdditionalGhostscriptSearchpath")
65230   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
65240     .AdditionalGhostscriptSearchpath = ""
65250    Else
65260     If LenB(tStr) > 0 Then
65270      .AdditionalGhostscriptSearchpath = tStr
65280     End If
65290   End If
65300   tStr = reg.GetRegistryValue("AddWindowsFontpath")
65310   If IsNumeric(tStr) Then
65320     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65330       .AddWindowsFontpath = CLng(tStr)
65340      Else
65350       If UseStandard Then
65360        .AddWindowsFontpath = 1
65370       End If
65380     End If
65390    Else
65400     If UseStandard Then
65410      .AddWindowsFontpath = 1
65420     End If
65430   End If
65440   tStr = reg.GetRegistryValue("AutosaveDirectory")
65450   If LenB(Trim$(tStr)) > 0 Then
65460     .AutosaveDirectory = CompletePath(tStr)
65470    Else
65480     If UseStandard Then
65490      If InstalledAsServer Then
65500        .AutosaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
65510       Else
65520        .AutosaveDirectory = "<MyFiles>"
65530      End If
65540     End If
65550   End If
65560   tStr = reg.GetRegistryValue("AutosaveFilename")
65570   If LenB(tStr) = 0 And LenB("<DateTime>") > 0 And UseStandard Then
65580     .AutosaveFilename = "<DateTime>"
65590    Else
65600     If LenB(tStr) > 0 Then
65610      .AutosaveFilename = tStr
65620     End If
65630   End If
65640   tStr = reg.GetRegistryValue("AutosaveFormat")
65650   If IsNumeric(tStr) Then
65660     If CLng(tStr) >= 0 And CLng(tStr) <= 13 Then
65670       .AutosaveFormat = CLng(tStr)
65680      Else
65690       If UseStandard Then
65700        .AutosaveFormat = 0
65710       End If
65720     End If
65730    Else
65740     If UseStandard Then
65750      .AutosaveFormat = 0
65760     End If
65770   End If
65780   tStr = reg.GetRegistryValue("AutosaveStartStandardProgram")
65790   If IsNumeric(tStr) Then
65800     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65810       .AutosaveStartStandardProgram = CLng(tStr)
65820      Else
65830       If UseStandard Then
65840        .AutosaveStartStandardProgram = 0
65850       End If
65860     End If
65870    Else
65880     If UseStandard Then
65890      .AutosaveStartStandardProgram = 0
65900     End If
65910   End If
65920   tStr = reg.GetRegistryValue("ClientComputerResolveIPAddress")
65930   If IsNumeric(tStr) Then
65940     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65950       .ClientComputerResolveIPAddress = CLng(tStr)
65960      Else
65970       If UseStandard Then
65980        .ClientComputerResolveIPAddress = 0
65990       End If
66000     End If
66010    Else
66020     If UseStandard Then
66030      .ClientComputerResolveIPAddress = 0
66040     End If
66050   End If
66060   tStr = reg.GetRegistryValue("DisableEmail")
66070   If IsNumeric(tStr) Then
66080     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66090       .DisableEmail = CLng(tStr)
66100      Else
66110       If UseStandard Then
66120        .DisableEmail = 0
66130       End If
66140     End If
66150    Else
66160     If UseStandard Then
66170      .DisableEmail = 0
66180     End If
66190   End If
66200   tStr = reg.GetRegistryValue("DontUseDocumentSettings")
66210   If IsNumeric(tStr) Then
66220     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66230       .DontUseDocumentSettings = CLng(tStr)
66240      Else
66250       If UseStandard Then
66260        .DontUseDocumentSettings = 0
66270       End If
66280     End If
66290    Else
66300     If UseStandard Then
66310      .DontUseDocumentSettings = 0
66320     End If
66330   End If
66340   tStr = reg.GetRegistryValue("FilenameSubstitutions")
66350   If LenB(tStr) = 0 And LenB("Microsoft Word - \.docx\.doc\Microsoft Excel - \.xlsx\.xls\Microsoft PowerPoint - \.pptx\.ppt") > 0 And UseStandard Then
66360     .FilenameSubstitutions = "Microsoft Word - \.docx\.doc\Microsoft Excel - \.xlsx\.xls\Microsoft PowerPoint - \.pptx\.ppt"
66370    Else
66380     If LenB(tStr) > 0 Then
66390      .FilenameSubstitutions = tStr
66400     End If
66410   End If
66420   tStr = reg.GetRegistryValue("FilenameSubstitutionsOnlyInTitle")
66430   If IsNumeric(tStr) Then
66440     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66450       .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
66460      Else
66470       If UseStandard Then
66480        .FilenameSubstitutionsOnlyInTitle = 1
66490       End If
66500     End If
66510    Else
66520     If UseStandard Then
66530      .FilenameSubstitutionsOnlyInTitle = 1
66540     End If
66550   End If
66560   tStr = reg.GetRegistryValue("Language")
66570   If LenB(tStr) = 0 And LenB("english") > 0 And UseStandard Then
66580     .Language = "english"
66590    Else
66600     If LenB(tStr) > 0 Then
66610      .Language = tStr
66620     End If
66630   End If
66640   tStr = reg.GetRegistryValue("LastSaveDirectory")
66650   If LenB(Trim$(tStr)) > 0 Then
66660     .LastSaveDirectory = CompletePath(tStr)
66670    Else
66680     If UseStandard Then
66690      If InstalledAsServer Then
66700        .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
66710       Else
66720        .LastSaveDirectory = "<MyFiles>"
66730      End If
66740     End If
66750   End If
66760   tStr = reg.GetRegistryValue("LastUpdateCheck")
66770   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
66780     .LastUpdateCheck = ""
66790    Else
66800     If LenB(tStr) > 0 Then
66810      .LastUpdateCheck = tStr
66820     End If
66830   End If
66840   tStr = reg.GetRegistryValue("Logging")
66850   If IsNumeric(tStr) Then
66860     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66870       .Logging = CLng(tStr)
66880      Else
66890       If UseStandard Then
66900        .Logging = 0
66910       End If
66920     End If
66930    Else
66940     If UseStandard Then
66950      .Logging = 0
66960     End If
66970   End If
66980   tStr = reg.GetRegistryValue("LogLines")
66990   If IsNumeric(tStr) Then
67000     If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
67010       .LogLines = CLng(tStr)
67020      Else
67030       If UseStandard Then
67040        .LogLines = 100
67050       End If
67060     End If
67070    Else
67080     If UseStandard Then
67090      .LogLines = 100
67100     End If
67110   End If
67120   tStr = reg.GetRegistryValue("NoConfirmMessageSwitchingDefaultprinter")
67130   If IsNumeric(tStr) Then
67140     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67150       .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
67160      Else
67170       If UseStandard Then
67180        .NoConfirmMessageSwitchingDefaultprinter = 0
67190       End If
67200     End If
67210    Else
67220     If UseStandard Then
67230      .NoConfirmMessageSwitchingDefaultprinter = 0
67240     End If
67250   End If
67260   tStr = reg.GetRegistryValue("NoProcessingAtStartup")
67270   If IsNumeric(tStr) Then
67280     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67290       .NoProcessingAtStartup = CLng(tStr)
67300      Else
67310       If UseStandard Then
67320        .NoProcessingAtStartup = 0
67330       End If
67340     End If
67350    Else
67360     If UseStandard Then
67370      .NoProcessingAtStartup = 0
67380     End If
67390   End If
67400   tStr = reg.GetRegistryValue("NoPSCheck")
67410   If IsNumeric(tStr) Then
67420     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67430       .NoPSCheck = CLng(tStr)
67440      Else
67450       If UseStandard Then
67460        .NoPSCheck = 0
67470       End If
67480     End If
67490    Else
67500     If UseStandard Then
67510      .NoPSCheck = 0
67520     End If
67530   End If
67540   tStr = reg.GetRegistryValue("OptionsDesign")
67550   If IsNumeric(tStr) Then
67560     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
67570       .OptionsDesign = CLng(tStr)
67580      Else
67590       If UseStandard Then
67600        .OptionsDesign = 0
67610       End If
67620     End If
67630    Else
67640     If UseStandard Then
67650      .OptionsDesign = 0
67660     End If
67670   End If
67680   tStr = reg.GetRegistryValue("OptionsEnabled")
67690   If IsNumeric(tStr) Then
67700     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67710       .OptionsEnabled = CLng(tStr)
67720      Else
67730       If UseStandard Then
67740        .OptionsEnabled = 1
67750       End If
67760     End If
67770    Else
67780     If UseStandard Then
67790      .OptionsEnabled = 1
67800     End If
67810   End If
67820   tStr = reg.GetRegistryValue("OptionsVisible")
67830   If IsNumeric(tStr) Then
67840     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67850       .OptionsVisible = CLng(tStr)
67860      Else
67870       If UseStandard Then
67880        .OptionsVisible = 1
67890       End If
67900     End If
67910    Else
67920     If UseStandard Then
67930      .OptionsVisible = 1
67940     End If
67950   End If
67960   tStr = reg.GetRegistryValue("PrintAfterSaving")
67970   If IsNumeric(tStr) Then
67980     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67990       .PrintAfterSaving = CLng(tStr)
68000      Else
68010       If UseStandard Then
68020        .PrintAfterSaving = 0
68030       End If
68040     End If
68050    Else
68060     If UseStandard Then
68070      .PrintAfterSaving = 0
68080     End If
68090   End If
68100   tStr = reg.GetRegistryValue("PrintAfterSavingDuplex")
68110   If IsNumeric(tStr) Then
68120     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68130       .PrintAfterSavingDuplex = CLng(tStr)
68140      Else
68150       If UseStandard Then
68160        .PrintAfterSavingDuplex = 0
68170       End If
68180     End If
68190    Else
68200     If UseStandard Then
68210      .PrintAfterSavingDuplex = 0
68220     End If
68230   End If
68240   tStr = reg.GetRegistryValue("PrintAfterSavingNoCancel")
68250   If IsNumeric(tStr) Then
68260     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68270       .PrintAfterSavingNoCancel = CLng(tStr)
68280      Else
68290       If UseStandard Then
68300        .PrintAfterSavingNoCancel = 0
68310       End If
68320     End If
68330    Else
68340     If UseStandard Then
68350      .PrintAfterSavingNoCancel = 0
68360     End If
68370   End If
68380   tStr = reg.GetRegistryValue("PrintAfterSavingPrinter")
68390   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
68400     .PrintAfterSavingPrinter = ""
68410    Else
68420     If LenB(tStr) > 0 Then
68430      .PrintAfterSavingPrinter = tStr
68440     End If
68450   End If
68460   tStr = reg.GetRegistryValue("PrintAfterSavingQueryUser")
68470   If IsNumeric(tStr) Then
68480     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
68490       .PrintAfterSavingQueryUser = CLng(tStr)
68500      Else
68510       If UseStandard Then
68520        .PrintAfterSavingQueryUser = 0
68530       End If
68540     End If
68550    Else
68560     If UseStandard Then
68570      .PrintAfterSavingQueryUser = 0
68580     End If
68590   End If
68600   tStr = reg.GetRegistryValue("PrintAfterSavingTumble")
68610   If IsNumeric(tStr) Then
68620     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
68630       .PrintAfterSavingTumble = CLng(tStr)
68640      Else
68650       If UseStandard Then
68660        .PrintAfterSavingTumble = 0
68670       End If
68680     End If
68690    Else
68700     If UseStandard Then
68710      .PrintAfterSavingTumble = 0
68720     End If
68730   End If
68740   tStr = reg.GetRegistryValue("PrinterStop")
68750   If IsNumeric(tStr) Then
68760     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68770       .PrinterStop = CLng(tStr)
68780      Else
68790       If UseStandard Then
68800        .PrinterStop = 0
68810       End If
68820     End If
68830    Else
68840     If UseStandard Then
68850      .PrinterStop = 0
68860     End If
68870   End If
68880   tStr = reg.GetRegistryValue("PrinterTemppath")
68890   WriteToSpecialLogfile "reg.GetRegistryValue(""PrinterTemppath"")=" & tStr
68900   WriteToSpecialLogfile "Options.PrinterTemppath1=" & .PrinterTemppath
68910   If hkey1 = HKEY_USERS Then
68920     If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then
68930       .PrinterTemppath = tStr
68940      Else
68950       If UseStandard Then
68960         .PrinterTemppath = GetTempPath
68970        Else
68980         .PrinterTemppath = tStr
68990       End If
69000     End If
69010    Else
69020     If LenB(Trim$(tStr)) > 0 Then
69030      If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then
69040        .PrinterTemppath = tStr
69050       Else
69060        MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))
69070        If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then
69080          If UseStandard Then
69090            .PrinterTemppath = GetTempPath
69100           Else
69110            .PrinterTemppath = ""
69120            If NoMsg = False Then
69130             MsgBox "PrinterTemppath: '" & tStr & "' = '" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & "'" & _
             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
69150            End If
69160          End If
69170         Else
69180          .PrinterTemppath = tStr
69190        End If
69200      End If
69210     End If
69220   End If
69230   WriteToSpecialLogfile "Options.PrinterTemppath2=" & .PrinterTemppath
69240   tStr = reg.GetRegistryValue("ProcessPriority")
69250   If IsNumeric(tStr) Then
69260     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
69270       .ProcessPriority = CLng(tStr)
69280      Else
69290       If UseStandard Then
69300        .ProcessPriority = 1
69310       End If
69320     End If
69330    Else
69340     If UseStandard Then
69350      .ProcessPriority = 1
69360     End If
69370   End If
69380   tStr = reg.GetRegistryValue("ProgramFont")
69390   If LenB(tStr) = 0 And LenB("MS Sans Serif") > 0 And UseStandard Then
69400     .ProgramFont = "MS Sans Serif"
69410    Else
69420     If LenB(tStr) > 0 Then
69430      .ProgramFont = tStr
69440     End If
69450   End If
69460   tStr = reg.GetRegistryValue("ProgramFontCharset")
69470   If IsNumeric(tStr) Then
69480     If CLng(tStr) >= 0 Then
69490       .ProgramFontCharset = CLng(tStr)
69500      Else
69510       If UseStandard Then
69520        .ProgramFontCharset = 0
69530       End If
69540     End If
69550    Else
69560     If UseStandard Then
69570      .ProgramFontCharset = 0
69580     End If
69590   End If
69600   tStr = reg.GetRegistryValue("ProgramFontSize")
69610   If IsNumeric(tStr) Then
69620     If CLng(tStr) >= 6 And CLng(tStr) <= 72 Then
69630       .ProgramFontSize = CLng(tStr)
69640      Else
69650       If UseStandard Then
69660        .ProgramFontSize = 8
69670       End If
69680     End If
69690    Else
69700     If UseStandard Then
69710      .ProgramFontSize = 8
69720     End If
69730   End If
69740   tStr = reg.GetRegistryValue("RemoveAllKnownFileExtensions")
69750   If IsNumeric(tStr) Then
69760     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69770       .RemoveAllKnownFileExtensions = CLng(tStr)
69780      Else
69790       If UseStandard Then
69800        .RemoveAllKnownFileExtensions = 1
69810       End If
69820     End If
69830    Else
69840     If UseStandard Then
69850      .RemoveAllKnownFileExtensions = 1
69860     End If
69870   End If
69880   tStr = reg.GetRegistryValue("RemoveSpaces")
69890   If IsNumeric(tStr) Then
69900     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69910       .RemoveSpaces = CLng(tStr)
69920      Else
69930       If UseStandard Then
69940        .RemoveSpaces = 1
69950       End If
69960     End If
69970    Else
69980     If UseStandard Then
69990      .RemoveSpaces = 1
70000     End If
70010   End If
70020   tStr = reg.GetRegistryValue("RunProgramAfterSaving")
70030   If IsNumeric(tStr) Then
70040     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70050       .RunProgramAfterSaving = CLng(tStr)
70060      Else
70070       If UseStandard Then
70080        .RunProgramAfterSaving = 0
70090       End If
70100     End If
70110    Else
70120     If UseStandard Then
70130      .RunProgramAfterSaving = 0
70140     End If
70150   End If
70160   tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramname")
70170   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70180     .RunProgramAfterSavingProgramname = ""
70190    Else
70200     If LenB(tStr) > 0 Then
70210      .RunProgramAfterSavingProgramname = tStr
70220     End If
70230   End If
70240   tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramParameters")
70250   If LenB(tStr) = 0 And LenB("""<OutputFilename>""") > 0 And UseStandard Then
70260     .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
70270    Else
70280     If LenB(tStr) > 0 Then
70290      .RunProgramAfterSavingProgramParameters = tStr
70300     End If
70310   End If
70320   tStr = reg.GetRegistryValue("RunProgramAfterSavingWaitUntilReady")
70330   If IsNumeric(tStr) Then
70340     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70350       .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
70360      Else
70370       If UseStandard Then
70380        .RunProgramAfterSavingWaitUntilReady = 1
70390       End If
70400     End If
70410    Else
70420     If UseStandard Then
70430      .RunProgramAfterSavingWaitUntilReady = 1
70440     End If
70450   End If
70460   tStr = reg.GetRegistryValue("RunProgramAfterSavingWindowstyle")
70470   If IsNumeric(tStr) Then
70480     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
70490       .RunProgramAfterSavingWindowstyle = CLng(tStr)
70500      Else
70510       If UseStandard Then
70520        .RunProgramAfterSavingWindowstyle = 1
70530       End If
70540     End If
70550    Else
70560     If UseStandard Then
70570      .RunProgramAfterSavingWindowstyle = 1
70580     End If
70590   End If
70600   tStr = reg.GetRegistryValue("RunProgramBeforeSaving")
70610   If IsNumeric(tStr) Then
70620     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70630       .RunProgramBeforeSaving = CLng(tStr)
70640      Else
70650       If UseStandard Then
70660        .RunProgramBeforeSaving = 0
70670       End If
70680     End If
70690    Else
70700     If UseStandard Then
70710      .RunProgramBeforeSaving = 0
70720     End If
70730   End If
70740   tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramname")
70750   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70760     .RunProgramBeforeSavingProgramname = ""
70770    Else
70780     If LenB(tStr) > 0 Then
70790      .RunProgramBeforeSavingProgramname = tStr
70800     End If
70810   End If
70820   tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramParameters")
70830   If LenB(tStr) = 0 And LenB("""<TempFilename>""") > 0 And UseStandard Then
70840     .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
70850    Else
70860     If LenB(tStr) > 0 Then
70870      .RunProgramBeforeSavingProgramParameters = tStr
70880     End If
70890   End If
70900   tStr = reg.GetRegistryValue("RunProgramBeforeSavingWindowstyle")
70910   If IsNumeric(tStr) Then
70920     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
70930       .RunProgramBeforeSavingWindowstyle = CLng(tStr)
70940      Else
70950       If UseStandard Then
70960        .RunProgramBeforeSavingWindowstyle = 1
70970       End If
70980     End If
70990    Else
71000     If UseStandard Then
71010      .RunProgramBeforeSavingWindowstyle = 1
71020     End If
71030   End If
71040   tStr = reg.GetRegistryValue("SaveFilename")
71050   If LenB(tStr) = 0 And LenB("<Title>") > 0 And UseStandard Then
71060     .SaveFilename = "<Title>"
71070    Else
71080     If LenB(tStr) > 0 Then
71090      .SaveFilename = tStr
71100     End If
71110   End If
71120   tStr = reg.GetRegistryValue("SendEmailAfterAutoSaving")
71130   If IsNumeric(tStr) Then
71140     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71150       .SendEmailAfterAutoSaving = CLng(tStr)
71160      Else
71170       If UseStandard Then
71180        .SendEmailAfterAutoSaving = 0
71190       End If
71200     End If
71210    Else
71220     If UseStandard Then
71230      .SendEmailAfterAutoSaving = 0
71240     End If
71250   End If
71260   tStr = reg.GetRegistryValue("SendMailMethod")
71270   If IsNumeric(tStr) Then
71280     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
71290       .SendMailMethod = CLng(tStr)
71300      Else
71310       If UseStandard Then
71320        .SendMailMethod = 0
71330       End If
71340     End If
71350    Else
71360     If UseStandard Then
71370      .SendMailMethod = 0
71380     End If
71390   End If
71400   tStr = reg.GetRegistryValue("ShowAnimation")
71410   If IsNumeric(tStr) Then
71420     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71430       .ShowAnimation = CLng(tStr)
71440      Else
71450       If UseStandard Then
71460        .ShowAnimation = 1
71470       End If
71480     End If
71490    Else
71500     If UseStandard Then
71510      .ShowAnimation = 1
71520     End If
71530   End If
71540   tStr = reg.GetRegistryValue("StartStandardProgram")
71550   If IsNumeric(tStr) Then
71560     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71570       .StartStandardProgram = CLng(tStr)
71580      Else
71590       If UseStandard Then
71600        .StartStandardProgram = 1
71610       End If
71620     End If
71630    Else
71640     If UseStandard Then
71650      .StartStandardProgram = 1
71660     End If
71670   End If
71680   tStr = reg.GetRegistryValue("Toolbars")
71690   If IsNumeric(tStr) Then
71700     If CLng(tStr) >= 0 Then
71710       .Toolbars = CLng(tStr)
71720      Else
71730       If UseStandard Then
71740        .Toolbars = 1
71750       End If
71760     End If
71770    Else
71780     If UseStandard Then
71790      .Toolbars = 1
71800     End If
71810   End If
71820   tStr = reg.GetRegistryValue("UpdateInterval")
71830   If IsNumeric(tStr) Then
71840     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
71850       .UpdateInterval = CLng(tStr)
71860      Else
71870       If UseStandard Then
71880        .UpdateInterval = 2
71890       End If
71900     End If
71910    Else
71920     If UseStandard Then
71930      .UpdateInterval = 2
71940     End If
71950   End If
71960   tStr = reg.GetRegistryValue("UseAutosave")
71970   If IsNumeric(tStr) Then
71980     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71990       .UseAutosave = CLng(tStr)
72000      Else
72010       If UseStandard Then
72020        .UseAutosave = 0
72030       End If
72040     End If
72050    Else
72060     If UseStandard Then
72070      .UseAutosave = 0
72080     End If
72090   End If
72100   tStr = reg.GetRegistryValue("UseAutosaveDirectory")
72110   If IsNumeric(tStr) Then
72120     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
72130       .UseAutosaveDirectory = CLng(tStr)
72140      Else
72150       If UseStandard Then
72160        .UseAutosaveDirectory = 1
72170       End If
72180     End If
72190    Else
72200     If UseStandard Then
72210      .UseAutosaveDirectory = 1
72220     End If
72230   End If
72240  End With
72250  Set reg = Nothing
72260  ReadOptionsReg = myOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "ReadOptionsReg")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub SaveOptionREG(sOptions As tOptions, OptionName As String, Optional hkey1 As hkey = HKEY_CURRENT_USER)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = hkey1
50040  reg.KeyRoot = "Software\PDFCreator"
50050  With sOptions
50060   reg.Subkey = "Ghostscript"
50070   If UCase$(OptionName) = "DIRECTORYGHOSTSCRIPTBINARIES" Then
50080    If Not reg.KeyExists Then
50090     reg.CreateKey
50100    End If
50110    reg.SetRegistryValue "DirectoryGhostscriptBinaries", CStr(.DirectoryGhostscriptBinaries), REG_SZ
50120    Set reg = Nothing
50130    Exit Sub
50140   End If
50150   If UCase$(OptionName) = "DIRECTORYGHOSTSCRIPTFONTS" Then
50160    If Not reg.KeyExists Then
50170     reg.CreateKey
50180    End If
50190    reg.SetRegistryValue "DirectoryGhostscriptFonts", CStr(.DirectoryGhostscriptFonts), REG_SZ
50200    Set reg = Nothing
50210    Exit Sub
50220   End If
50230   If UCase$(OptionName) = "DIRECTORYGHOSTSCRIPTLIBRARIES" Then
50240    If Not reg.KeyExists Then
50250     reg.CreateKey
50260    End If
50270    reg.SetRegistryValue "DirectoryGhostscriptLibraries", CStr(.DirectoryGhostscriptLibraries), REG_SZ
50280    Set reg = Nothing
50290    Exit Sub
50300   End If
50310   If UCase$(OptionName) = "DIRECTORYGHOSTSCRIPTRESOURCE" Then
50320    If Not reg.KeyExists Then
50330     reg.CreateKey
50340    End If
50350    reg.SetRegistryValue "DirectoryGhostscriptResource", CStr(.DirectoryGhostscriptResource), REG_SZ
50360    Set reg = Nothing
50370    Exit Sub
50380   End If
50390   reg.Subkey = "Printing"
50400   If UCase$(OptionName) = "COUNTER" Then
50410    If Not reg.KeyExists Then
50420     reg.CreateKey
50430    End If
50440    reg.SetRegistryValue "Counter", CStr(.Counter), REG_SZ
50450    Set reg = Nothing
50460    Exit Sub
50470   End If
50480   If UCase$(OptionName) = "DEVICEHEIGHTPOINTS" Then
50490    If Not reg.KeyExists Then
50500     reg.CreateKey
50510    End If
50520   reg.SetRegistryValue "DeviceHeightPoints", Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), REG_SZ
50530    Set reg = Nothing
50540    Exit Sub
50550   End If
50560   If UCase$(OptionName) = "DEVICEWIDTHPOINTS" Then
50570    If Not reg.KeyExists Then
50580     reg.CreateKey
50590    End If
50600   reg.SetRegistryValue "DeviceWidthPoints", Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), REG_SZ
50610    Set reg = Nothing
50620    Exit Sub
50630   End If
50640   If UCase$(OptionName) = "ONEPAGEPERFILE" Then
50650    If Not reg.KeyExists Then
50660     reg.CreateKey
50670    End If
50680    reg.SetRegistryValue "OnePagePerFile", CStr(Abs(.OnePagePerFile)), REG_SZ
50690    Set reg = Nothing
50700    Exit Sub
50710   End If
50720   If UCase$(OptionName) = "PAPERSIZE" Then
50730    If Not reg.KeyExists Then
50740     reg.CreateKey
50750    End If
50760    reg.SetRegistryValue "Papersize", CStr(.Papersize), REG_SZ
50770    Set reg = Nothing
50780    Exit Sub
50790   End If
50800   If UCase$(OptionName) = "STAMPFONTCOLOR" Then
50810    If Not reg.KeyExists Then
50820     reg.CreateKey
50830    End If
50840    reg.SetRegistryValue "StampFontColor", CStr(.StampFontColor), REG_SZ
50850    Set reg = Nothing
50860    Exit Sub
50870   End If
50880   If UCase$(OptionName) = "STAMPFONTNAME" Then
50890    If Not reg.KeyExists Then
50900     reg.CreateKey
50910    End If
50920    reg.SetRegistryValue "StampFontname", CStr(.StampFontname), REG_SZ
50930    Set reg = Nothing
50940    Exit Sub
50950   End If
50960   If UCase$(OptionName) = "STAMPFONTSIZE" Then
50970    If Not reg.KeyExists Then
50980     reg.CreateKey
50990    End If
51000    reg.SetRegistryValue "StampFontsize", CStr(.StampFontsize), REG_SZ
51010    Set reg = Nothing
51020    Exit Sub
51030   End If
51040   If UCase$(OptionName) = "STAMPOUTLINEFONTTHICKNESS" Then
51050    If Not reg.KeyExists Then
51060     reg.CreateKey
51070    End If
51080    reg.SetRegistryValue "StampOutlineFontthickness", CStr(.StampOutlineFontthickness), REG_SZ
51090    Set reg = Nothing
51100    Exit Sub
51110   End If
51120   If UCase$(OptionName) = "STAMPSTRING" Then
51130    If Not reg.KeyExists Then
51140     reg.CreateKey
51150    End If
51160    reg.SetRegistryValue "StampString", CStr(.StampString), REG_SZ
51170    Set reg = Nothing
51180    Exit Sub
51190   End If
51200   If UCase$(OptionName) = "STAMPUSEOUTLINEFONT" Then
51210    If Not reg.KeyExists Then
51220     reg.CreateKey
51230    End If
51240    reg.SetRegistryValue "StampUseOutlineFont", CStr(Abs(.StampUseOutlineFont)), REG_SZ
51250    Set reg = Nothing
51260    Exit Sub
51270   End If
51280   If UCase$(OptionName) = "STANDARDAUTHOR" Then
51290    If Not reg.KeyExists Then
51300     reg.CreateKey
51310    End If
51320    reg.SetRegistryValue "StandardAuthor", CStr(.StandardAuthor), REG_SZ
51330    Set reg = Nothing
51340    Exit Sub
51350   End If
51360   If UCase$(OptionName) = "STANDARDCREATIONDATE" Then
51370    If Not reg.KeyExists Then
51380     reg.CreateKey
51390    End If
51400    reg.SetRegistryValue "StandardCreationdate", CStr(.StandardCreationdate), REG_SZ
51410    Set reg = Nothing
51420    Exit Sub
51430   End If
51440   If UCase$(OptionName) = "STANDARDDATEFORMAT" Then
51450    If Not reg.KeyExists Then
51460     reg.CreateKey
51470    End If
51480    reg.SetRegistryValue "StandardDateformat", CStr(.StandardDateformat), REG_SZ
51490    Set reg = Nothing
51500    Exit Sub
51510   End If
51520   If UCase$(OptionName) = "STANDARDKEYWORDS" Then
51530    If Not reg.KeyExists Then
51540     reg.CreateKey
51550    End If
51560    reg.SetRegistryValue "StandardKeywords", CStr(.StandardKeywords), REG_SZ
51570    Set reg = Nothing
51580    Exit Sub
51590   End If
51600   If UCase$(OptionName) = "STANDARDMAILDOMAIN" Then
51610    If Not reg.KeyExists Then
51620     reg.CreateKey
51630    End If
51640    reg.SetRegistryValue "StandardMailDomain", CStr(.StandardMailDomain), REG_SZ
51650    Set reg = Nothing
51660    Exit Sub
51670   End If
51680   If UCase$(OptionName) = "STANDARDMODIFYDATE" Then
51690    If Not reg.KeyExists Then
51700     reg.CreateKey
51710    End If
51720    reg.SetRegistryValue "StandardModifydate", CStr(.StandardModifydate), REG_SZ
51730    Set reg = Nothing
51740    Exit Sub
51750   End If
51760   If UCase$(OptionName) = "STANDARDSAVEFORMAT" Then
51770    If Not reg.KeyExists Then
51780     reg.CreateKey
51790    End If
51800    reg.SetRegistryValue "StandardSaveformat", CStr(.StandardSaveformat), REG_SZ
51810    Set reg = Nothing
51820    Exit Sub
51830   End If
51840   If UCase$(OptionName) = "STANDARDSUBJECT" Then
51850    If Not reg.KeyExists Then
51860     reg.CreateKey
51870    End If
51880    reg.SetRegistryValue "StandardSubject", CStr(.StandardSubject), REG_SZ
51890    Set reg = Nothing
51900    Exit Sub
51910   End If
51920   If UCase$(OptionName) = "STANDARDTITLE" Then
51930    If Not reg.KeyExists Then
51940     reg.CreateKey
51950    End If
51960    reg.SetRegistryValue "StandardTitle", CStr(.StandardTitle), REG_SZ
51970    Set reg = Nothing
51980    Exit Sub
51990   End If
52000   If UCase$(OptionName) = "USECREATIONDATENOW" Then
52010    If Not reg.KeyExists Then
52020     reg.CreateKey
52030    End If
52040    reg.SetRegistryValue "UseCreationDateNow", CStr(Abs(.UseCreationDateNow)), REG_SZ
52050    Set reg = Nothing
52060    Exit Sub
52070   End If
52080   If UCase$(OptionName) = "USECUSTOMPAPERSIZE" Then
52090    If Not reg.KeyExists Then
52100     reg.CreateKey
52110    End If
52120    reg.SetRegistryValue "UseCustomPaperSize", CStr(.UseCustomPaperSize), REG_SZ
52130    Set reg = Nothing
52140    Exit Sub
52150   End If
52160   If UCase$(OptionName) = "USEFIXPAPERSIZE" Then
52170    If Not reg.KeyExists Then
52180     reg.CreateKey
52190    End If
52200    reg.SetRegistryValue "UseFixPapersize", CStr(Abs(.UseFixPapersize)), REG_SZ
52210    Set reg = Nothing
52220    Exit Sub
52230   End If
52240   If UCase$(OptionName) = "USESTANDARDAUTHOR" Then
52250    If Not reg.KeyExists Then
52260     reg.CreateKey
52270    End If
52280    reg.SetRegistryValue "UseStandardAuthor", CStr(Abs(.UseStandardAuthor)), REG_SZ
52290    Set reg = Nothing
52300    Exit Sub
52310   End If
52320   reg.Subkey = "Printing\Formats\Bitmap\Colors"
52330   If UCase$(OptionName) = "BMPCOLORSCOUNT" Then
52340    If Not reg.KeyExists Then
52350     reg.CreateKey
52360    End If
52370    reg.SetRegistryValue "BMPColorscount", CStr(.BMPColorscount), REG_SZ
52380    Set reg = Nothing
52390    Exit Sub
52400   End If
52410   If UCase$(OptionName) = "BMPRESOLUTION" Then
52420    If Not reg.KeyExists Then
52430     reg.CreateKey
52440    End If
52450    reg.SetRegistryValue "BMPResolution", CStr(.BMPResolution), REG_SZ
52460    Set reg = Nothing
52470    Exit Sub
52480   End If
52490   If UCase$(OptionName) = "JPEGCOLORSCOUNT" Then
52500    If Not reg.KeyExists Then
52510     reg.CreateKey
52520    End If
52530    reg.SetRegistryValue "JPEGColorscount", CStr(.JPEGColorscount), REG_SZ
52540    Set reg = Nothing
52550    Exit Sub
52560   End If
52570   If UCase$(OptionName) = "JPEGQUALITY" Then
52580    If Not reg.KeyExists Then
52590     reg.CreateKey
52600    End If
52610    reg.SetRegistryValue "JPEGQuality", CStr(.JPEGQuality), REG_SZ
52620    Set reg = Nothing
52630    Exit Sub
52640   End If
52650   If UCase$(OptionName) = "JPEGRESOLUTION" Then
52660    If Not reg.KeyExists Then
52670     reg.CreateKey
52680    End If
52690    reg.SetRegistryValue "JPEGResolution", CStr(.JPEGResolution), REG_SZ
52700    Set reg = Nothing
52710    Exit Sub
52720   End If
52730   If UCase$(OptionName) = "PCLCOLORSCOUNT" Then
52740    If Not reg.KeyExists Then
52750     reg.CreateKey
52760    End If
52770    reg.SetRegistryValue "PCLColorsCount", CStr(.PCLColorsCount), REG_SZ
52780    Set reg = Nothing
52790    Exit Sub
52800   End If
52810   If UCase$(OptionName) = "PCLRESOLUTION" Then
52820    If Not reg.KeyExists Then
52830     reg.CreateKey
52840    End If
52850    reg.SetRegistryValue "PCLResolution", CStr(.PCLResolution), REG_SZ
52860    Set reg = Nothing
52870    Exit Sub
52880   End If
52890   If UCase$(OptionName) = "PCXCOLORSCOUNT" Then
52900    If Not reg.KeyExists Then
52910     reg.CreateKey
52920    End If
52930    reg.SetRegistryValue "PCXColorscount", CStr(.PCXColorscount), REG_SZ
52940    Set reg = Nothing
52950    Exit Sub
52960   End If
52970   If UCase$(OptionName) = "PCXRESOLUTION" Then
52980    If Not reg.KeyExists Then
52990     reg.CreateKey
53000    End If
53010    reg.SetRegistryValue "PCXResolution", CStr(.PCXResolution), REG_SZ
53020    Set reg = Nothing
53030    Exit Sub
53040   End If
53050   If UCase$(OptionName) = "PNGCOLORSCOUNT" Then
53060    If Not reg.KeyExists Then
53070     reg.CreateKey
53080    End If
53090    reg.SetRegistryValue "PNGColorscount", CStr(.PNGColorscount), REG_SZ
53100    Set reg = Nothing
53110    Exit Sub
53120   End If
53130   If UCase$(OptionName) = "PNGRESOLUTION" Then
53140    If Not reg.KeyExists Then
53150     reg.CreateKey
53160    End If
53170    reg.SetRegistryValue "PNGResolution", CStr(.PNGResolution), REG_SZ
53180    Set reg = Nothing
53190    Exit Sub
53200   End If
53210   If UCase$(OptionName) = "PSDCOLORSCOUNT" Then
53220    If Not reg.KeyExists Then
53230     reg.CreateKey
53240    End If
53250    reg.SetRegistryValue "PSDColorsCount", CStr(.PSDColorsCount), REG_SZ
53260    Set reg = Nothing
53270    Exit Sub
53280   End If
53290   If UCase$(OptionName) = "PSDRESOLUTION" Then
53300    If Not reg.KeyExists Then
53310     reg.CreateKey
53320    End If
53330    reg.SetRegistryValue "PSDResolution", CStr(.PSDResolution), REG_SZ
53340    Set reg = Nothing
53350    Exit Sub
53360   End If
53370   If UCase$(OptionName) = "RAWCOLORSCOUNT" Then
53380    If Not reg.KeyExists Then
53390     reg.CreateKey
53400    End If
53410    reg.SetRegistryValue "RAWColorsCount", CStr(.RAWColorsCount), REG_SZ
53420    Set reg = Nothing
53430    Exit Sub
53440   End If
53450   If UCase$(OptionName) = "RAWRESOLUTION" Then
53460    If Not reg.KeyExists Then
53470     reg.CreateKey
53480    End If
53490    reg.SetRegistryValue "RAWResolution", CStr(.RAWResolution), REG_SZ
53500    Set reg = Nothing
53510    Exit Sub
53520   End If
53530   If UCase$(OptionName) = "TIFFCOLORSCOUNT" Then
53540    If Not reg.KeyExists Then
53550     reg.CreateKey
53560    End If
53570    reg.SetRegistryValue "TIFFColorscount", CStr(.TIFFColorscount), REG_SZ
53580    Set reg = Nothing
53590    Exit Sub
53600   End If
53610   If UCase$(OptionName) = "TIFFRESOLUTION" Then
53620    If Not reg.KeyExists Then
53630     reg.CreateKey
53640    End If
53650    reg.SetRegistryValue "TIFFResolution", CStr(.TIFFResolution), REG_SZ
53660    Set reg = Nothing
53670    Exit Sub
53680   End If
53690   reg.Subkey = "Printing\Formats\PDF\Colors"
53700   If UCase$(OptionName) = "PDFCOLORSCMYKTORGB" Then
53710    If Not reg.KeyExists Then
53720     reg.CreateKey
53730    End If
53740    reg.SetRegistryValue "PDFColorsCMYKToRGB", CStr(Abs(.PDFColorsCMYKToRGB)), REG_SZ
53750    Set reg = Nothing
53760    Exit Sub
53770   End If
53780   If UCase$(OptionName) = "PDFCOLORSCOLORMODEL" Then
53790    If Not reg.KeyExists Then
53800     reg.CreateKey
53810    End If
53820    reg.SetRegistryValue "PDFColorsColorModel", CStr(.PDFColorsColorModel), REG_SZ
53830    Set reg = Nothing
53840    Exit Sub
53850   End If
53860   If UCase$(OptionName) = "PDFCOLORSPRESERVEHALFTONE" Then
53870    If Not reg.KeyExists Then
53880     reg.CreateKey
53890    End If
53900    reg.SetRegistryValue "PDFColorsPreserveHalftone", CStr(Abs(.PDFColorsPreserveHalftone)), REG_SZ
53910    Set reg = Nothing
53920    Exit Sub
53930   End If
53940   If UCase$(OptionName) = "PDFCOLORSPRESERVEOVERPRINT" Then
53950    If Not reg.KeyExists Then
53960     reg.CreateKey
53970    End If
53980    reg.SetRegistryValue "PDFColorsPreserveOverprint", CStr(Abs(.PDFColorsPreserveOverprint)), REG_SZ
53990    Set reg = Nothing
54000    Exit Sub
54010   End If
54020   If UCase$(OptionName) = "PDFCOLORSPRESERVETRANSFER" Then
54030    If Not reg.KeyExists Then
54040     reg.CreateKey
54050    End If
54060    reg.SetRegistryValue "PDFColorsPreserveTransfer", CStr(Abs(.PDFColorsPreserveTransfer)), REG_SZ
54070    Set reg = Nothing
54080    Exit Sub
54090   End If
54100   reg.Subkey = "Printing\Formats\PDF\Compression"
54110   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSION" Then
54120    If Not reg.KeyExists Then
54130     reg.CreateKey
54140    End If
54150    reg.SetRegistryValue "PDFCompressionColorCompression", CStr(Abs(.PDFCompressionColorCompression)), REG_SZ
54160    Set reg = Nothing
54170    Exit Sub
54180   End If
54190   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE" Then
54200    If Not reg.KeyExists Then
54210     reg.CreateKey
54220    End If
54230    reg.SetRegistryValue "PDFCompressionColorCompressionChoice", CStr(.PDFCompressionColorCompressionChoice), REG_SZ
54240    Set reg = Nothing
54250    Exit Sub
54260   End If
54270   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR" Then
54280    If Not reg.KeyExists Then
54290     reg.CreateKey
54300    End If
54310   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
54320    Set reg = Nothing
54330    Exit Sub
54340   End If
54350   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR" Then
54360    If Not reg.KeyExists Then
54370     reg.CreateKey
54380    End If
54390   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
54400    Set reg = Nothing
54410    Exit Sub
54420   End If
54430   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR" Then
54440    If Not reg.KeyExists Then
54450     reg.CreateKey
54460    End If
54470   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
54480    Set reg = Nothing
54490    Exit Sub
54500   End If
54510   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR" Then
54520    If Not reg.KeyExists Then
54530     reg.CreateKey
54540    End If
54550   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
54560    Set reg = Nothing
54570    Exit Sub
54580   End If
54590   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR" Then
54600    If Not reg.KeyExists Then
54610     reg.CreateKey
54620    End If
54630   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
54640    Set reg = Nothing
54650    Exit Sub
54660   End If
54670   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESAMPLE" Then
54680    If Not reg.KeyExists Then
54690     reg.CreateKey
54700    End If
54710    reg.SetRegistryValue "PDFCompressionColorResample", CStr(Abs(.PDFCompressionColorResample)), REG_SZ
54720    Set reg = Nothing
54730    Exit Sub
54740   End If
54750   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESAMPLECHOICE" Then
54760    If Not reg.KeyExists Then
54770     reg.CreateKey
54780    End If
54790    reg.SetRegistryValue "PDFCompressionColorResampleChoice", CStr(.PDFCompressionColorResampleChoice), REG_SZ
54800    Set reg = Nothing
54810    Exit Sub
54820   End If
54830   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESOLUTION" Then
54840    If Not reg.KeyExists Then
54850     reg.CreateKey
54860    End If
54870    reg.SetRegistryValue "PDFCompressionColorResolution", CStr(.PDFCompressionColorResolution), REG_SZ
54880    Set reg = Nothing
54890    Exit Sub
54900   End If
54910   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSION" Then
54920    If Not reg.KeyExists Then
54930     reg.CreateKey
54940    End If
54950    reg.SetRegistryValue "PDFCompressionGreyCompression", CStr(Abs(.PDFCompressionGreyCompression)), REG_SZ
54960    Set reg = Nothing
54970    Exit Sub
54980   End If
54990   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE" Then
55000    If Not reg.KeyExists Then
55010     reg.CreateKey
55020    End If
55030    reg.SetRegistryValue "PDFCompressionGreyCompressionChoice", CStr(.PDFCompressionGreyCompressionChoice), REG_SZ
55040    Set reg = Nothing
55050    Exit Sub
55060   End If
55070   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR" Then
55080    If Not reg.KeyExists Then
55090     reg.CreateKey
55100    End If
55110   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
55120    Set reg = Nothing
55130    Exit Sub
55140   End If
55150   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR" Then
55160    If Not reg.KeyExists Then
55170     reg.CreateKey
55180    End If
55190   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
55200    Set reg = Nothing
55210    Exit Sub
55220   End If
55230   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR" Then
55240    If Not reg.KeyExists Then
55250     reg.CreateKey
55260    End If
55270   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
55280    Set reg = Nothing
55290    Exit Sub
55300   End If
55310   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR" Then
55320    If Not reg.KeyExists Then
55330     reg.CreateKey
55340    End If
55350   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
55360    Set reg = Nothing
55370    Exit Sub
55380   End If
55390   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR" Then
55400    If Not reg.KeyExists Then
55410     reg.CreateKey
55420    End If
55430   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
55440    Set reg = Nothing
55450    Exit Sub
55460   End If
55470   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESAMPLE" Then
55480    If Not reg.KeyExists Then
55490     reg.CreateKey
55500    End If
55510    reg.SetRegistryValue "PDFCompressionGreyResample", CStr(Abs(.PDFCompressionGreyResample)), REG_SZ
55520    Set reg = Nothing
55530    Exit Sub
55540   End If
55550   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESAMPLECHOICE" Then
55560    If Not reg.KeyExists Then
55570     reg.CreateKey
55580    End If
55590    reg.SetRegistryValue "PDFCompressionGreyResampleChoice", CStr(.PDFCompressionGreyResampleChoice), REG_SZ
55600    Set reg = Nothing
55610    Exit Sub
55620   End If
55630   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESOLUTION" Then
55640    If Not reg.KeyExists Then
55650     reg.CreateKey
55660    End If
55670    reg.SetRegistryValue "PDFCompressionGreyResolution", CStr(.PDFCompressionGreyResolution), REG_SZ
55680    Set reg = Nothing
55690    Exit Sub
55700   End If
55710   If UCase$(OptionName) = "PDFCOMPRESSIONMONOCOMPRESSION" Then
55720    If Not reg.KeyExists Then
55730     reg.CreateKey
55740    End If
55750    reg.SetRegistryValue "PDFCompressionMonoCompression", CStr(Abs(.PDFCompressionMonoCompression)), REG_SZ
55760    Set reg = Nothing
55770    Exit Sub
55780   End If
55790   If UCase$(OptionName) = "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE" Then
55800    If Not reg.KeyExists Then
55810     reg.CreateKey
55820    End If
55830    reg.SetRegistryValue "PDFCompressionMonoCompressionChoice", CStr(.PDFCompressionMonoCompressionChoice), REG_SZ
55840    Set reg = Nothing
55850    Exit Sub
55860   End If
55870   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESAMPLE" Then
55880    If Not reg.KeyExists Then
55890     reg.CreateKey
55900    End If
55910    reg.SetRegistryValue "PDFCompressionMonoResample", CStr(Abs(.PDFCompressionMonoResample)), REG_SZ
55920    Set reg = Nothing
55930    Exit Sub
55940   End If
55950   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESAMPLECHOICE" Then
55960    If Not reg.KeyExists Then
55970     reg.CreateKey
55980    End If
55990    reg.SetRegistryValue "PDFCompressionMonoResampleChoice", CStr(.PDFCompressionMonoResampleChoice), REG_SZ
56000    Set reg = Nothing
56010    Exit Sub
56020   End If
56030   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESOLUTION" Then
56040    If Not reg.KeyExists Then
56050     reg.CreateKey
56060    End If
56070    reg.SetRegistryValue "PDFCompressionMonoResolution", CStr(.PDFCompressionMonoResolution), REG_SZ
56080    Set reg = Nothing
56090    Exit Sub
56100   End If
56110   If UCase$(OptionName) = "PDFCOMPRESSIONTEXTCOMPRESSION" Then
56120    If Not reg.KeyExists Then
56130     reg.CreateKey
56140    End If
56150    reg.SetRegistryValue "PDFCompressionTextCompression", CStr(Abs(.PDFCompressionTextCompression)), REG_SZ
56160    Set reg = Nothing
56170    Exit Sub
56180   End If
56190   reg.Subkey = "Printing\Formats\PDF\Fonts"
56200   If UCase$(OptionName) = "PDFFONTSEMBEDALL" Then
56210    If Not reg.KeyExists Then
56220     reg.CreateKey
56230    End If
56240    reg.SetRegistryValue "PDFFontsEmbedAll", CStr(Abs(.PDFFontsEmbedAll)), REG_SZ
56250    Set reg = Nothing
56260    Exit Sub
56270   End If
56280   If UCase$(OptionName) = "PDFFONTSSUBSETFONTS" Then
56290    If Not reg.KeyExists Then
56300     reg.CreateKey
56310    End If
56320    reg.SetRegistryValue "PDFFontsSubSetFonts", CStr(Abs(.PDFFontsSubSetFonts)), REG_SZ
56330    Set reg = Nothing
56340    Exit Sub
56350   End If
56360   If UCase$(OptionName) = "PDFFONTSSUBSETFONTSPERCENT" Then
56370    If Not reg.KeyExists Then
56380     reg.CreateKey
56390    End If
56400    reg.SetRegistryValue "PDFFontsSubSetFontsPercent", CStr(.PDFFontsSubSetFontsPercent), REG_SZ
56410    Set reg = Nothing
56420    Exit Sub
56430   End If
56440   reg.Subkey = "Printing\Formats\PDF\General"
56450   If UCase$(OptionName) = "PDFGENERALASCII85" Then
56460    If Not reg.KeyExists Then
56470     reg.CreateKey
56480    End If
56490    reg.SetRegistryValue "PDFGeneralASCII85", CStr(Abs(.PDFGeneralASCII85)), REG_SZ
56500    Set reg = Nothing
56510    Exit Sub
56520   End If
56530   If UCase$(OptionName) = "PDFGENERALAUTOROTATE" Then
56540    If Not reg.KeyExists Then
56550     reg.CreateKey
56560    End If
56570    reg.SetRegistryValue "PDFGeneralAutorotate", CStr(.PDFGeneralAutorotate), REG_SZ
56580    Set reg = Nothing
56590    Exit Sub
56600   End If
56610   If UCase$(OptionName) = "PDFGENERALCOMPATIBILITY" Then
56620    If Not reg.KeyExists Then
56630     reg.CreateKey
56640    End If
56650    reg.SetRegistryValue "PDFGeneralCompatibility", CStr(.PDFGeneralCompatibility), REG_SZ
56660    Set reg = Nothing
56670    Exit Sub
56680   End If
56690   If UCase$(OptionName) = "PDFGENERALDEFAULT" Then
56700    If Not reg.KeyExists Then
56710     reg.CreateKey
56720    End If
56730    reg.SetRegistryValue "PDFGeneralDefault", CStr(.PDFGeneralDefault), REG_SZ
56740    Set reg = Nothing
56750    Exit Sub
56760   End If
56770   If UCase$(OptionName) = "PDFGENERALOVERPRINT" Then
56780    If Not reg.KeyExists Then
56790     reg.CreateKey
56800    End If
56810    reg.SetRegistryValue "PDFGeneralOverprint", CStr(.PDFGeneralOverprint), REG_SZ
56820    Set reg = Nothing
56830    Exit Sub
56840   End If
56850   If UCase$(OptionName) = "PDFGENERALRESOLUTION" Then
56860    If Not reg.KeyExists Then
56870     reg.CreateKey
56880    End If
56890    reg.SetRegistryValue "PDFGeneralResolution", CStr(.PDFGeneralResolution), REG_SZ
56900    Set reg = Nothing
56910    Exit Sub
56920   End If
56930   If UCase$(OptionName) = "PDFOPTIMIZE" Then
56940    If Not reg.KeyExists Then
56950     reg.CreateKey
56960    End If
56970    reg.SetRegistryValue "PDFOptimize", CStr(Abs(.PDFOptimize)), REG_SZ
56980    Set reg = Nothing
56990    Exit Sub
57000   End If
57010   If UCase$(OptionName) = "PDFUPDATEMETADATA" Then
57020    If Not reg.KeyExists Then
57030     reg.CreateKey
57040    End If
57050    reg.SetRegistryValue "PDFUpdateMetadata", CStr(.PDFUpdateMetadata), REG_SZ
57060    Set reg = Nothing
57070    Exit Sub
57080   End If
57090   reg.Subkey = "Printing\Formats\PDF\Security"
57100   If UCase$(OptionName) = "PDFALLOWASSEMBLY" Then
57110    If Not reg.KeyExists Then
57120     reg.CreateKey
57130    End If
57140    reg.SetRegistryValue "PDFAllowAssembly", CStr(Abs(.PDFAllowAssembly)), REG_SZ
57150    Set reg = Nothing
57160    Exit Sub
57170   End If
57180   If UCase$(OptionName) = "PDFALLOWDEGRADEDPRINTING" Then
57190    If Not reg.KeyExists Then
57200     reg.CreateKey
57210    End If
57220    reg.SetRegistryValue "PDFAllowDegradedPrinting", CStr(Abs(.PDFAllowDegradedPrinting)), REG_SZ
57230    Set reg = Nothing
57240    Exit Sub
57250   End If
57260   If UCase$(OptionName) = "PDFALLOWFILLIN" Then
57270    If Not reg.KeyExists Then
57280     reg.CreateKey
57290    End If
57300    reg.SetRegistryValue "PDFAllowFillIn", CStr(Abs(.PDFAllowFillIn)), REG_SZ
57310    Set reg = Nothing
57320    Exit Sub
57330   End If
57340   If UCase$(OptionName) = "PDFALLOWSCREENREADERS" Then
57350    If Not reg.KeyExists Then
57360     reg.CreateKey
57370    End If
57380    reg.SetRegistryValue "PDFAllowScreenReaders", CStr(Abs(.PDFAllowScreenReaders)), REG_SZ
57390    Set reg = Nothing
57400    Exit Sub
57410   End If
57420   If UCase$(OptionName) = "PDFDISALLOWCOPY" Then
57430    If Not reg.KeyExists Then
57440     reg.CreateKey
57450    End If
57460    reg.SetRegistryValue "PDFDisallowCopy", CStr(Abs(.PDFDisallowCopy)), REG_SZ
57470    Set reg = Nothing
57480    Exit Sub
57490   End If
57500   If UCase$(OptionName) = "PDFDISALLOWMODIFYANNOTATIONS" Then
57510    If Not reg.KeyExists Then
57520     reg.CreateKey
57530    End If
57540    reg.SetRegistryValue "PDFDisallowModifyAnnotations", CStr(Abs(.PDFDisallowModifyAnnotations)), REG_SZ
57550    Set reg = Nothing
57560    Exit Sub
57570   End If
57580   If UCase$(OptionName) = "PDFDISALLOWMODIFYCONTENTS" Then
57590    If Not reg.KeyExists Then
57600     reg.CreateKey
57610    End If
57620    reg.SetRegistryValue "PDFDisallowModifyContents", CStr(Abs(.PDFDisallowModifyContents)), REG_SZ
57630    Set reg = Nothing
57640    Exit Sub
57650   End If
57660   If UCase$(OptionName) = "PDFDISALLOWPRINTING" Then
57670    If Not reg.KeyExists Then
57680     reg.CreateKey
57690    End If
57700    reg.SetRegistryValue "PDFDisallowPrinting", CStr(Abs(.PDFDisallowPrinting)), REG_SZ
57710    Set reg = Nothing
57720    Exit Sub
57730   End If
57740   If UCase$(OptionName) = "PDFENCRYPTOR" Then
57750    If Not reg.KeyExists Then
57760     reg.CreateKey
57770    End If
57780    reg.SetRegistryValue "PDFEncryptor", CStr(.PDFEncryptor), REG_SZ
57790    Set reg = Nothing
57800    Exit Sub
57810   End If
57820   If UCase$(OptionName) = "PDFHIGHENCRYPTION" Then
57830    If Not reg.KeyExists Then
57840     reg.CreateKey
57850    End If
57860    reg.SetRegistryValue "PDFHighEncryption", CStr(Abs(.PDFHighEncryption)), REG_SZ
57870    Set reg = Nothing
57880    Exit Sub
57890   End If
57900   If UCase$(OptionName) = "PDFLOWENCRYPTION" Then
57910    If Not reg.KeyExists Then
57920     reg.CreateKey
57930    End If
57940    reg.SetRegistryValue "PDFLowEncryption", CStr(Abs(.PDFLowEncryption)), REG_SZ
57950    Set reg = Nothing
57960    Exit Sub
57970   End If
57980   If UCase$(OptionName) = "PDFOWNERPASS" Then
57990    If Not reg.KeyExists Then
58000     reg.CreateKey
58010    End If
58020    reg.SetRegistryValue "PDFOwnerPass", CStr(Abs(.PDFOwnerPass)), REG_SZ
58030    Set reg = Nothing
58040    Exit Sub
58050   End If
58060   If UCase$(OptionName) = "PDFOWNERPASSWORDSTRING" Then
58070    If Not reg.KeyExists Then
58080     reg.CreateKey
58090    End If
58100    reg.SetRegistryValue "PDFOwnerPasswordString", CStr(.PDFOwnerPasswordString), REG_SZ
58110    Set reg = Nothing
58120    Exit Sub
58130   End If
58140   If UCase$(OptionName) = "PDFUSERPASS" Then
58150    If Not reg.KeyExists Then
58160     reg.CreateKey
58170    End If
58180    reg.SetRegistryValue "PDFUserPass", CStr(Abs(.PDFUserPass)), REG_SZ
58190    Set reg = Nothing
58200    Exit Sub
58210   End If
58220   If UCase$(OptionName) = "PDFUSERPASSWORDSTRING" Then
58230    If Not reg.KeyExists Then
58240     reg.CreateKey
58250    End If
58260    reg.SetRegistryValue "PDFUserPasswordString", CStr(.PDFUserPasswordString), REG_SZ
58270    Set reg = Nothing
58280    Exit Sub
58290   End If
58300   If UCase$(OptionName) = "PDFUSESECURITY" Then
58310    If Not reg.KeyExists Then
58320     reg.CreateKey
58330    End If
58340    reg.SetRegistryValue "PDFUseSecurity", CStr(Abs(.PDFUseSecurity)), REG_SZ
58350    Set reg = Nothing
58360    Exit Sub
58370   End If
58380   reg.Subkey = "Printing\Formats\PDF\Signing"
58390   If UCase$(OptionName) = "PDFSIGNINGMULTISIGNATURE" Then
58400    If Not reg.KeyExists Then
58410     reg.CreateKey
58420    End If
58430    reg.SetRegistryValue "PDFSigningMultiSignature", CStr(Abs(.PDFSigningMultiSignature)), REG_SZ
58440    Set reg = Nothing
58450    Exit Sub
58460   End If
58470   If UCase$(OptionName) = "PDFSIGNINGPFXFILE" Then
58480    If Not reg.KeyExists Then
58490     reg.CreateKey
58500    End If
58510    reg.SetRegistryValue "PDFSigningPFXFile", CStr(.PDFSigningPFXFile), REG_SZ
58520    Set reg = Nothing
58530    Exit Sub
58540   End If
58550   If UCase$(OptionName) = "PDFSIGNINGPFXFILEPASSWORD" Then
58560    If Not reg.KeyExists Then
58570     reg.CreateKey
58580    End If
58590    reg.SetRegistryValue "PDFSigningPFXFilePassword", CStr(.PDFSigningPFXFilePassword), REG_SZ
58600    Set reg = Nothing
58610    Exit Sub
58620   End If
58630   If UCase$(OptionName) = "PDFSIGNINGSIGNATURECONTACT" Then
58640    If Not reg.KeyExists Then
58650     reg.CreateKey
58660    End If
58670    reg.SetRegistryValue "PDFSigningSignatureContact", CStr(.PDFSigningSignatureContact), REG_SZ
58680    Set reg = Nothing
58690    Exit Sub
58700   End If
58710   If UCase$(OptionName) = "PDFSIGNINGSIGNATURELEFTX" Then
58720    If Not reg.KeyExists Then
58730     reg.CreateKey
58740    End If
58750   reg.SetRegistryValue "PDFSigningSignatureLeftX", Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), REG_SZ
58760    Set reg = Nothing
58770    Exit Sub
58780   End If
58790   If UCase$(OptionName) = "PDFSIGNINGSIGNATURELEFTY" Then
58800    If Not reg.KeyExists Then
58810     reg.CreateKey
58820    End If
58830   reg.SetRegistryValue "PDFSigningSignatureLeftY", Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), REG_SZ
58840    Set reg = Nothing
58850    Exit Sub
58860   End If
58870   If UCase$(OptionName) = "PDFSIGNINGSIGNATURELOCATION" Then
58880    If Not reg.KeyExists Then
58890     reg.CreateKey
58900    End If
58910    reg.SetRegistryValue "PDFSigningSignatureLocation", CStr(.PDFSigningSignatureLocation), REG_SZ
58920    Set reg = Nothing
58930    Exit Sub
58940   End If
58950   If UCase$(OptionName) = "PDFSIGNINGSIGNATUREREASON" Then
58960    If Not reg.KeyExists Then
58970     reg.CreateKey
58980    End If
58990    reg.SetRegistryValue "PDFSigningSignatureReason", CStr(.PDFSigningSignatureReason), REG_SZ
59000    Set reg = Nothing
59010    Exit Sub
59020   End If
59030   If UCase$(OptionName) = "PDFSIGNINGSIGNATURERIGHTX" Then
59040    If Not reg.KeyExists Then
59050     reg.CreateKey
59060    End If
59070   reg.SetRegistryValue "PDFSigningSignatureRightX", Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), REG_SZ
59080    Set reg = Nothing
59090    Exit Sub
59100   End If
59110   If UCase$(OptionName) = "PDFSIGNINGSIGNATURERIGHTY" Then
59120    If Not reg.KeyExists Then
59130     reg.CreateKey
59140    End If
59150   reg.SetRegistryValue "PDFSigningSignatureRightY", Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), REG_SZ
59160    Set reg = Nothing
59170    Exit Sub
59180   End If
59190   If UCase$(OptionName) = "PDFSIGNINGSIGNATUREVISIBLE" Then
59200    If Not reg.KeyExists Then
59210     reg.CreateKey
59220    End If
59230    reg.SetRegistryValue "PDFSigningSignatureVisible", CStr(Abs(.PDFSigningSignatureVisible)), REG_SZ
59240    Set reg = Nothing
59250    Exit Sub
59260   End If
59270   If UCase$(OptionName) = "PDFSIGNINGSIGNPDF" Then
59280    If Not reg.KeyExists Then
59290     reg.CreateKey
59300    End If
59310    reg.SetRegistryValue "PDFSigningSignPDF", CStr(Abs(.PDFSigningSignPDF)), REG_SZ
59320    Set reg = Nothing
59330    Exit Sub
59340   End If
59350   reg.Subkey = "Printing\Formats\PS\LanguageLevel"
59360   If UCase$(OptionName) = "EPSLANGUAGELEVEL" Then
59370    If Not reg.KeyExists Then
59380     reg.CreateKey
59390    End If
59400    reg.SetRegistryValue "EPSLanguageLevel", CStr(.EPSLanguageLevel), REG_SZ
59410    Set reg = Nothing
59420    Exit Sub
59430   End If
59440   If UCase$(OptionName) = "PSLANGUAGELEVEL" Then
59450    If Not reg.KeyExists Then
59460     reg.CreateKey
59470    End If
59480    reg.SetRegistryValue "PSLanguageLevel", CStr(.PSLanguageLevel), REG_SZ
59490    Set reg = Nothing
59500    Exit Sub
59510   End If
59520   reg.Subkey = "Program"
59530   If UCase$(OptionName) = "ADDITIONALGHOSTSCRIPTPARAMETERS" Then
59540    If Not reg.KeyExists Then
59550     reg.CreateKey
59560    End If
59570    reg.SetRegistryValue "AdditionalGhostscriptParameters", CStr(.AdditionalGhostscriptParameters), REG_SZ
59580    Set reg = Nothing
59590    Exit Sub
59600   End If
59610   If UCase$(OptionName) = "ADDITIONALGHOSTSCRIPTSEARCHPATH" Then
59620    If Not reg.KeyExists Then
59630     reg.CreateKey
59640    End If
59650    reg.SetRegistryValue "AdditionalGhostscriptSearchpath", CStr(.AdditionalGhostscriptSearchpath), REG_SZ
59660    Set reg = Nothing
59670    Exit Sub
59680   End If
59690   If UCase$(OptionName) = "ADDWINDOWSFONTPATH" Then
59700    If Not reg.KeyExists Then
59710     reg.CreateKey
59720    End If
59730    reg.SetRegistryValue "AddWindowsFontpath", CStr(Abs(.AddWindowsFontpath)), REG_SZ
59740    Set reg = Nothing
59750    Exit Sub
59760   End If
59770   If UCase$(OptionName) = "AUTOSAVEDIRECTORY" Then
59780    If Not reg.KeyExists Then
59790     reg.CreateKey
59800    End If
59810    reg.SetRegistryValue "AutosaveDirectory", CStr(.AutosaveDirectory), REG_SZ
59820    Set reg = Nothing
59830    Exit Sub
59840   End If
59850   If UCase$(OptionName) = "AUTOSAVEFILENAME" Then
59860    If Not reg.KeyExists Then
59870     reg.CreateKey
59880    End If
59890    reg.SetRegistryValue "AutosaveFilename", CStr(.AutosaveFilename), REG_SZ
59900    Set reg = Nothing
59910    Exit Sub
59920   End If
59930   If UCase$(OptionName) = "AUTOSAVEFORMAT" Then
59940    If Not reg.KeyExists Then
59950     reg.CreateKey
59960    End If
59970    reg.SetRegistryValue "AutosaveFormat", CStr(.AutosaveFormat), REG_SZ
59980    Set reg = Nothing
59990    Exit Sub
60000   End If
60010   If UCase$(OptionName) = "AUTOSAVESTARTSTANDARDPROGRAM" Then
60020    If Not reg.KeyExists Then
60030     reg.CreateKey
60040    End If
60050    reg.SetRegistryValue "AutosaveStartStandardProgram", CStr(Abs(.AutosaveStartStandardProgram)), REG_SZ
60060    Set reg = Nothing
60070    Exit Sub
60080   End If
60090   If UCase$(OptionName) = "CLIENTCOMPUTERRESOLVEIPADDRESS" Then
60100    If Not reg.KeyExists Then
60110     reg.CreateKey
60120    End If
60130    reg.SetRegistryValue "ClientComputerResolveIPAddress", CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
60140    Set reg = Nothing
60150    Exit Sub
60160   End If
60170   If UCase$(OptionName) = "DISABLEEMAIL" Then
60180    If Not reg.KeyExists Then
60190     reg.CreateKey
60200    End If
60210    reg.SetRegistryValue "DisableEmail", CStr(Abs(.DisableEmail)), REG_SZ
60220    Set reg = Nothing
60230    Exit Sub
60240   End If
60250   If UCase$(OptionName) = "DONTUSEDOCUMENTSETTINGS" Then
60260    If Not reg.KeyExists Then
60270     reg.CreateKey
60280    End If
60290    reg.SetRegistryValue "DontUseDocumentSettings", CStr(Abs(.DontUseDocumentSettings)), REG_SZ
60300    Set reg = Nothing
60310    Exit Sub
60320   End If
60330   If UCase$(OptionName) = "FILENAMESUBSTITUTIONS" Then
60340    If Not reg.KeyExists Then
60350     reg.CreateKey
60360    End If
60370    reg.SetRegistryValue "FilenameSubstitutions", CStr(.FilenameSubstitutions), REG_SZ
60380    Set reg = Nothing
60390    Exit Sub
60400   End If
60410   If UCase$(OptionName) = "FILENAMESUBSTITUTIONSONLYINTITLE" Then
60420    If Not reg.KeyExists Then
60430     reg.CreateKey
60440    End If
60450    reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle", CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
60460    Set reg = Nothing
60470    Exit Sub
60480   End If
60490   If UCase$(OptionName) = "LANGUAGE" Then
60500    If Not reg.KeyExists Then
60510     reg.CreateKey
60520    End If
60530    reg.SetRegistryValue "Language", CStr(.Language), REG_SZ
60540    Set reg = Nothing
60550    Exit Sub
60560   End If
60570   If UCase$(OptionName) = "LASTSAVEDIRECTORY" Then
60580    If Not reg.KeyExists Then
60590     reg.CreateKey
60600    End If
60610    reg.SetRegistryValue "LastSaveDirectory", CStr(.LastSaveDirectory), REG_SZ
60620    Set reg = Nothing
60630    Exit Sub
60640   End If
60650   If UCase$(OptionName) = "LASTUPDATECHECK" Then
60660    If Not reg.KeyExists Then
60670     reg.CreateKey
60680    End If
60690    reg.SetRegistryValue "LastUpdateCheck", CStr(.LastUpdateCheck), REG_SZ
60700    Set reg = Nothing
60710    Exit Sub
60720   End If
60730   If UCase$(OptionName) = "LOGGING" Then
60740    If Not reg.KeyExists Then
60750     reg.CreateKey
60760    End If
60770    reg.SetRegistryValue "Logging", CStr(Abs(.Logging)), REG_SZ
60780    Set reg = Nothing
60790    Exit Sub
60800   End If
60810   If UCase$(OptionName) = "LOGLINES" Then
60820    If Not reg.KeyExists Then
60830     reg.CreateKey
60840    End If
60850    reg.SetRegistryValue "LogLines", CStr(.LogLines), REG_SZ
60860    Set reg = Nothing
60870    Exit Sub
60880   End If
60890   If UCase$(OptionName) = "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER" Then
60900    If Not reg.KeyExists Then
60910     reg.CreateKey
60920    End If
60930    reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter", CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
60940    Set reg = Nothing
60950    Exit Sub
60960   End If
60970   If UCase$(OptionName) = "NOPROCESSINGATSTARTUP" Then
60980    If Not reg.KeyExists Then
60990     reg.CreateKey
61000    End If
61010    reg.SetRegistryValue "NoProcessingAtStartup", CStr(Abs(.NoProcessingAtStartup)), REG_SZ
61020    Set reg = Nothing
61030    Exit Sub
61040   End If
61050   If UCase$(OptionName) = "NOPSCHECK" Then
61060    If Not reg.KeyExists Then
61070     reg.CreateKey
61080    End If
61090    reg.SetRegistryValue "NoPSCheck", CStr(Abs(.NoPSCheck)), REG_SZ
61100    Set reg = Nothing
61110    Exit Sub
61120   End If
61130   If UCase$(OptionName) = "OPTIONSDESIGN" Then
61140    If Not reg.KeyExists Then
61150     reg.CreateKey
61160    End If
61170    reg.SetRegistryValue "OptionsDesign", CStr(.OptionsDesign), REG_SZ
61180    Set reg = Nothing
61190    Exit Sub
61200   End If
61210   If UCase$(OptionName) = "OPTIONSENABLED" Then
61220    If Not reg.KeyExists Then
61230     reg.CreateKey
61240    End If
61250    reg.SetRegistryValue "OptionsEnabled", CStr(Abs(.OptionsEnabled)), REG_SZ
61260    Set reg = Nothing
61270    Exit Sub
61280   End If
61290   If UCase$(OptionName) = "OPTIONSVISIBLE" Then
61300    If Not reg.KeyExists Then
61310     reg.CreateKey
61320    End If
61330    reg.SetRegistryValue "OptionsVisible", CStr(Abs(.OptionsVisible)), REG_SZ
61340    Set reg = Nothing
61350    Exit Sub
61360   End If
61370   If UCase$(OptionName) = "PRINTAFTERSAVING" Then
61380    If Not reg.KeyExists Then
61390     reg.CreateKey
61400    End If
61410    reg.SetRegistryValue "PrintAfterSaving", CStr(Abs(.PrintAfterSaving)), REG_SZ
61420    Set reg = Nothing
61430    Exit Sub
61440   End If
61450   If UCase$(OptionName) = "PRINTAFTERSAVINGDUPLEX" Then
61460    If Not reg.KeyExists Then
61470     reg.CreateKey
61480    End If
61490    reg.SetRegistryValue "PrintAfterSavingDuplex", CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
61500    Set reg = Nothing
61510    Exit Sub
61520   End If
61530   If UCase$(OptionName) = "PRINTAFTERSAVINGNOCANCEL" Then
61540    If Not reg.KeyExists Then
61550     reg.CreateKey
61560    End If
61570    reg.SetRegistryValue "PrintAfterSavingNoCancel", CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
61580    Set reg = Nothing
61590    Exit Sub
61600   End If
61610   If UCase$(OptionName) = "PRINTAFTERSAVINGPRINTER" Then
61620    If Not reg.KeyExists Then
61630     reg.CreateKey
61640    End If
61650    reg.SetRegistryValue "PrintAfterSavingPrinter", CStr(.PrintAfterSavingPrinter), REG_SZ
61660    Set reg = Nothing
61670    Exit Sub
61680   End If
61690   If UCase$(OptionName) = "PRINTAFTERSAVINGQUERYUSER" Then
61700    If Not reg.KeyExists Then
61710     reg.CreateKey
61720    End If
61730    reg.SetRegistryValue "PrintAfterSavingQueryUser", CStr(.PrintAfterSavingQueryUser), REG_SZ
61740    Set reg = Nothing
61750    Exit Sub
61760   End If
61770   If UCase$(OptionName) = "PRINTAFTERSAVINGTUMBLE" Then
61780    If Not reg.KeyExists Then
61790     reg.CreateKey
61800    End If
61810    reg.SetRegistryValue "PrintAfterSavingTumble", CStr(.PrintAfterSavingTumble), REG_SZ
61820    Set reg = Nothing
61830    Exit Sub
61840   End If
61850   If UCase$(OptionName) = "PRINTERSTOP" Then
61860    If Not reg.KeyExists Then
61870     reg.CreateKey
61880    End If
61890    reg.SetRegistryValue "PrinterStop", CStr(Abs(.PrinterStop)), REG_SZ
61900    Set reg = Nothing
61910    Exit Sub
61920   End If
61930   If UCase$(OptionName) = "PRINTERTEMPPATH" Then
61940    If Not reg.KeyExists Then
61950     reg.CreateKey
61960    End If
61970    reg.SetRegistryValue "PrinterTemppath", CStr(.PrinterTemppath), REG_SZ
61980    Set reg = Nothing
61990    Exit Sub
62000   End If
62010   If UCase$(OptionName) = "PROCESSPRIORITY" Then
62020    If Not reg.KeyExists Then
62030     reg.CreateKey
62040    End If
62050    reg.SetRegistryValue "ProcessPriority", CStr(.ProcessPriority), REG_SZ
62060    Set reg = Nothing
62070    Exit Sub
62080   End If
62090   If UCase$(OptionName) = "PROGRAMFONT" Then
62100    If Not reg.KeyExists Then
62110     reg.CreateKey
62120    End If
62130    reg.SetRegistryValue "ProgramFont", CStr(.ProgramFont), REG_SZ
62140    Set reg = Nothing
62150    Exit Sub
62160   End If
62170   If UCase$(OptionName) = "PROGRAMFONTCHARSET" Then
62180    If Not reg.KeyExists Then
62190     reg.CreateKey
62200    End If
62210    reg.SetRegistryValue "ProgramFontCharset", CStr(.ProgramFontCharset), REG_SZ
62220    Set reg = Nothing
62230    Exit Sub
62240   End If
62250   If UCase$(OptionName) = "PROGRAMFONTSIZE" Then
62260    If Not reg.KeyExists Then
62270     reg.CreateKey
62280    End If
62290    reg.SetRegistryValue "ProgramFontSize", CStr(.ProgramFontSize), REG_SZ
62300    Set reg = Nothing
62310    Exit Sub
62320   End If
62330   If UCase$(OptionName) = "REMOVEALLKNOWNFILEEXTENSIONS" Then
62340    If Not reg.KeyExists Then
62350     reg.CreateKey
62360    End If
62370    reg.SetRegistryValue "RemoveAllKnownFileExtensions", CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
62380    Set reg = Nothing
62390    Exit Sub
62400   End If
62410   If UCase$(OptionName) = "REMOVESPACES" Then
62420    If Not reg.KeyExists Then
62430     reg.CreateKey
62440    End If
62450    reg.SetRegistryValue "RemoveSpaces", CStr(Abs(.RemoveSpaces)), REG_SZ
62460    Set reg = Nothing
62470    Exit Sub
62480   End If
62490   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVING" Then
62500    If Not reg.KeyExists Then
62510     reg.CreateKey
62520    End If
62530    reg.SetRegistryValue "RunProgramAfterSaving", CStr(Abs(.RunProgramAfterSaving)), REG_SZ
62540    Set reg = Nothing
62550    Exit Sub
62560   End If
62570   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGPROGRAMNAME" Then
62580    If Not reg.KeyExists Then
62590     reg.CreateKey
62600    End If
62610    reg.SetRegistryValue "RunProgramAfterSavingProgramname", CStr(.RunProgramAfterSavingProgramname), REG_SZ
62620    Set reg = Nothing
62630    Exit Sub
62640   End If
62650   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS" Then
62660    If Not reg.KeyExists Then
62670     reg.CreateKey
62680    End If
62690    reg.SetRegistryValue "RunProgramAfterSavingProgramParameters", CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
62700    Set reg = Nothing
62710    Exit Sub
62720   End If
62730   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGWAITUNTILREADY" Then
62740    If Not reg.KeyExists Then
62750     reg.CreateKey
62760    End If
62770    reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady", CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
62780    Set reg = Nothing
62790    Exit Sub
62800   End If
62810   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGWINDOWSTYLE" Then
62820    If Not reg.KeyExists Then
62830     reg.CreateKey
62840    End If
62850    reg.SetRegistryValue "RunProgramAfterSavingWindowstyle", CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
62860    Set reg = Nothing
62870    Exit Sub
62880   End If
62890   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVING" Then
62900    If Not reg.KeyExists Then
62910     reg.CreateKey
62920    End If
62930    reg.SetRegistryValue "RunProgramBeforeSaving", CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
62940    Set reg = Nothing
62950    Exit Sub
62960   End If
62970   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGPROGRAMNAME" Then
62980    If Not reg.KeyExists Then
62990     reg.CreateKey
63000    End If
63010    reg.SetRegistryValue "RunProgramBeforeSavingProgramname", CStr(.RunProgramBeforeSavingProgramname), REG_SZ
63020    Set reg = Nothing
63030    Exit Sub
63040   End If
63050   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS" Then
63060    If Not reg.KeyExists Then
63070     reg.CreateKey
63080    End If
63090    reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters", CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
63100    Set reg = Nothing
63110    Exit Sub
63120   End If
63130   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGWINDOWSTYLE" Then
63140    If Not reg.KeyExists Then
63150     reg.CreateKey
63160    End If
63170    reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle", CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
63180    Set reg = Nothing
63190    Exit Sub
63200   End If
63210   If UCase$(OptionName) = "SAVEFILENAME" Then
63220    If Not reg.KeyExists Then
63230     reg.CreateKey
63240    End If
63250    reg.SetRegistryValue "SaveFilename", CStr(.SaveFilename), REG_SZ
63260    Set reg = Nothing
63270    Exit Sub
63280   End If
63290   If UCase$(OptionName) = "SENDEMAILAFTERAUTOSAVING" Then
63300    If Not reg.KeyExists Then
63310     reg.CreateKey
63320    End If
63330    reg.SetRegistryValue "SendEmailAfterAutoSaving", CStr(Abs(.SendEmailAfterAutoSaving)), REG_SZ
63340    Set reg = Nothing
63350    Exit Sub
63360   End If
63370   If UCase$(OptionName) = "SENDMAILMETHOD" Then
63380    If Not reg.KeyExists Then
63390     reg.CreateKey
63400    End If
63410    reg.SetRegistryValue "SendMailMethod", CStr(.SendMailMethod), REG_SZ
63420    Set reg = Nothing
63430    Exit Sub
63440   End If
63450   If UCase$(OptionName) = "SHOWANIMATION" Then
63460    If Not reg.KeyExists Then
63470     reg.CreateKey
63480    End If
63490    reg.SetRegistryValue "ShowAnimation", CStr(Abs(.ShowAnimation)), REG_SZ
63500    Set reg = Nothing
63510    Exit Sub
63520   End If
63530   If UCase$(OptionName) = "STARTSTANDARDPROGRAM" Then
63540    If Not reg.KeyExists Then
63550     reg.CreateKey
63560    End If
63570    reg.SetRegistryValue "StartStandardProgram", CStr(Abs(.StartStandardProgram)), REG_SZ
63580    Set reg = Nothing
63590    Exit Sub
63600   End If
63610   If UCase$(OptionName) = "TOOLBARS" Then
63620    If Not reg.KeyExists Then
63630     reg.CreateKey
63640    End If
63650    reg.SetRegistryValue "Toolbars", CStr(.Toolbars), REG_SZ
63660    Set reg = Nothing
63670    Exit Sub
63680   End If
63690   If UCase$(OptionName) = "UPDATEINTERVAL" Then
63700    If Not reg.KeyExists Then
63710     reg.CreateKey
63720    End If
63730    reg.SetRegistryValue "UpdateInterval", CStr(.UpdateInterval), REG_SZ
63740    Set reg = Nothing
63750    Exit Sub
63760   End If
63770   If UCase$(OptionName) = "USEAUTOSAVE" Then
63780    If Not reg.KeyExists Then
63790     reg.CreateKey
63800    End If
63810    reg.SetRegistryValue "UseAutosave", CStr(Abs(.UseAutosave)), REG_SZ
63820    Set reg = Nothing
63830    Exit Sub
63840   End If
63850   If UCase$(OptionName) = "USEAUTOSAVEDIRECTORY" Then
63860    If Not reg.KeyExists Then
63870     reg.CreateKey
63880    End If
63890    reg.SetRegistryValue "UseAutosaveDirectory", CStr(Abs(.UseAutosaveDirectory)), REG_SZ
63900    Set reg = Nothing
63910    Exit Sub
63920   End If
63930  End With
63940  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "SaveOptionREG")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SaveOptionsREG(sOptions As tOptions, Optional hkey1 As hkey = HKEY_CURRENT_USER)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = hkey1
50040  reg.KeyRoot = "Software\PDFCreator"
50050  If Not reg.KeyExists Then
50060   reg.CreateKey
50070  End If
50080  With sOptions
50090   reg.Subkey = "Ghostscript"
50100   If Not reg.KeyExists Then
50110    reg.CreateKey
50120   End If
50130   reg.SetRegistryValue "DirectoryGhostscriptBinaries", CStr(.DirectoryGhostscriptBinaries), REG_SZ
50140   reg.SetRegistryValue "DirectoryGhostscriptFonts", CStr(.DirectoryGhostscriptFonts), REG_SZ
50150   reg.SetRegistryValue "DirectoryGhostscriptLibraries", CStr(.DirectoryGhostscriptLibraries), REG_SZ
50160   reg.SetRegistryValue "DirectoryGhostscriptResource", CStr(.DirectoryGhostscriptResource), REG_SZ
50170   reg.Subkey = "Printing"
50180   If Not reg.KeyExists Then
50190    reg.CreateKey
50200   End If
50210   reg.SetRegistryValue "Counter", CStr(.Counter), REG_SZ
50220   reg.SetRegistryValue "DeviceHeightPoints", Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), REG_SZ
50230   reg.SetRegistryValue "DeviceWidthPoints", Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), REG_SZ
50240   reg.SetRegistryValue "OnePagePerFile", CStr(Abs(.OnePagePerFile)), REG_SZ
50250   reg.SetRegistryValue "Papersize", CStr(.Papersize), REG_SZ
50260   reg.SetRegistryValue "StampFontColor", CStr(.StampFontColor), REG_SZ
50270   reg.SetRegistryValue "StampFontname", CStr(.StampFontname), REG_SZ
50280   reg.SetRegistryValue "StampFontsize", CStr(.StampFontsize), REG_SZ
50290   reg.SetRegistryValue "StampOutlineFontthickness", CStr(.StampOutlineFontthickness), REG_SZ
50300   reg.SetRegistryValue "StampString", CStr(.StampString), REG_SZ
50310   reg.SetRegistryValue "StampUseOutlineFont", CStr(Abs(.StampUseOutlineFont)), REG_SZ
50320   reg.SetRegistryValue "StandardAuthor", CStr(.StandardAuthor), REG_SZ
50330   reg.SetRegistryValue "StandardCreationdate", CStr(.StandardCreationdate), REG_SZ
50340   reg.SetRegistryValue "StandardDateformat", CStr(.StandardDateformat), REG_SZ
50350   reg.SetRegistryValue "StandardKeywords", CStr(.StandardKeywords), REG_SZ
50360   reg.SetRegistryValue "StandardMailDomain", CStr(.StandardMailDomain), REG_SZ
50370   reg.SetRegistryValue "StandardModifydate", CStr(.StandardModifydate), REG_SZ
50380   reg.SetRegistryValue "StandardSaveformat", CStr(.StandardSaveformat), REG_SZ
50390   reg.SetRegistryValue "StandardSubject", CStr(.StandardSubject), REG_SZ
50400   reg.SetRegistryValue "StandardTitle", CStr(.StandardTitle), REG_SZ
50410   reg.SetRegistryValue "UseCreationDateNow", CStr(Abs(.UseCreationDateNow)), REG_SZ
50420   reg.SetRegistryValue "UseCustomPaperSize", CStr(.UseCustomPaperSize), REG_SZ
50430   reg.SetRegistryValue "UseFixPapersize", CStr(Abs(.UseFixPapersize)), REG_SZ
50440   reg.SetRegistryValue "UseStandardAuthor", CStr(Abs(.UseStandardAuthor)), REG_SZ
50450   reg.Subkey = "Printing\Formats\Bitmap\Colors"
50460   If Not reg.KeyExists Then
50470    reg.CreateKey
50480   End If
50490   reg.SetRegistryValue "BMPColorscount", CStr(.BMPColorscount), REG_SZ
50500   reg.SetRegistryValue "BMPResolution", CStr(.BMPResolution), REG_SZ
50510   reg.SetRegistryValue "JPEGColorscount", CStr(.JPEGColorscount), REG_SZ
50520   reg.SetRegistryValue "JPEGQuality", CStr(.JPEGQuality), REG_SZ
50530   reg.SetRegistryValue "JPEGResolution", CStr(.JPEGResolution), REG_SZ
50540   reg.SetRegistryValue "PCLColorsCount", CStr(.PCLColorsCount), REG_SZ
50550   reg.SetRegistryValue "PCLResolution", CStr(.PCLResolution), REG_SZ
50560   reg.SetRegistryValue "PCXColorscount", CStr(.PCXColorscount), REG_SZ
50570   reg.SetRegistryValue "PCXResolution", CStr(.PCXResolution), REG_SZ
50580   reg.SetRegistryValue "PNGColorscount", CStr(.PNGColorscount), REG_SZ
50590   reg.SetRegistryValue "PNGResolution", CStr(.PNGResolution), REG_SZ
50600   reg.SetRegistryValue "PSDColorsCount", CStr(.PSDColorsCount), REG_SZ
50610   reg.SetRegistryValue "PSDResolution", CStr(.PSDResolution), REG_SZ
50620   reg.SetRegistryValue "RAWColorsCount", CStr(.RAWColorsCount), REG_SZ
50630   reg.SetRegistryValue "RAWResolution", CStr(.RAWResolution), REG_SZ
50640   reg.SetRegistryValue "TIFFColorscount", CStr(.TIFFColorscount), REG_SZ
50650   reg.SetRegistryValue "TIFFResolution", CStr(.TIFFResolution), REG_SZ
50660   reg.Subkey = "Printing\Formats\PDF\Colors"
50670   If Not reg.KeyExists Then
50680    reg.CreateKey
50690   End If
50700   reg.SetRegistryValue "PDFColorsCMYKToRGB", CStr(Abs(.PDFColorsCMYKToRGB)), REG_SZ
50710   reg.SetRegistryValue "PDFColorsColorModel", CStr(.PDFColorsColorModel), REG_SZ
50720   reg.SetRegistryValue "PDFColorsPreserveHalftone", CStr(Abs(.PDFColorsPreserveHalftone)), REG_SZ
50730   reg.SetRegistryValue "PDFColorsPreserveOverprint", CStr(Abs(.PDFColorsPreserveOverprint)), REG_SZ
50740   reg.SetRegistryValue "PDFColorsPreserveTransfer", CStr(Abs(.PDFColorsPreserveTransfer)), REG_SZ
50750   reg.Subkey = "Printing\Formats\PDF\Compression"
50760   If Not reg.KeyExists Then
50770    reg.CreateKey
50780   End If
50790   reg.SetRegistryValue "PDFCompressionColorCompression", CStr(Abs(.PDFCompressionColorCompression)), REG_SZ
50800   reg.SetRegistryValue "PDFCompressionColorCompressionChoice", CStr(.PDFCompressionColorCompressionChoice), REG_SZ
50810   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
50820   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
50830   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
50840   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
50850   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
50860   reg.SetRegistryValue "PDFCompressionColorResample", CStr(Abs(.PDFCompressionColorResample)), REG_SZ
50870   reg.SetRegistryValue "PDFCompressionColorResampleChoice", CStr(.PDFCompressionColorResampleChoice), REG_SZ
50880   reg.SetRegistryValue "PDFCompressionColorResolution", CStr(.PDFCompressionColorResolution), REG_SZ
50890   reg.SetRegistryValue "PDFCompressionGreyCompression", CStr(Abs(.PDFCompressionGreyCompression)), REG_SZ
50900   reg.SetRegistryValue "PDFCompressionGreyCompressionChoice", CStr(.PDFCompressionGreyCompressionChoice), REG_SZ
50910   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
50920   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
50930   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
50940   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
50950   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
50960   reg.SetRegistryValue "PDFCompressionGreyResample", CStr(Abs(.PDFCompressionGreyResample)), REG_SZ
50970   reg.SetRegistryValue "PDFCompressionGreyResampleChoice", CStr(.PDFCompressionGreyResampleChoice), REG_SZ
50980   reg.SetRegistryValue "PDFCompressionGreyResolution", CStr(.PDFCompressionGreyResolution), REG_SZ
50990   reg.SetRegistryValue "PDFCompressionMonoCompression", CStr(Abs(.PDFCompressionMonoCompression)), REG_SZ
51000   reg.SetRegistryValue "PDFCompressionMonoCompressionChoice", CStr(.PDFCompressionMonoCompressionChoice), REG_SZ
51010   reg.SetRegistryValue "PDFCompressionMonoResample", CStr(Abs(.PDFCompressionMonoResample)), REG_SZ
51020   reg.SetRegistryValue "PDFCompressionMonoResampleChoice", CStr(.PDFCompressionMonoResampleChoice), REG_SZ
51030   reg.SetRegistryValue "PDFCompressionMonoResolution", CStr(.PDFCompressionMonoResolution), REG_SZ
51040   reg.SetRegistryValue "PDFCompressionTextCompression", CStr(Abs(.PDFCompressionTextCompression)), REG_SZ
51050   reg.Subkey = "Printing\Formats\PDF\Fonts"
51060   If Not reg.KeyExists Then
51070    reg.CreateKey
51080   End If
51090   reg.SetRegistryValue "PDFFontsEmbedAll", CStr(Abs(.PDFFontsEmbedAll)), REG_SZ
51100   reg.SetRegistryValue "PDFFontsSubSetFonts", CStr(Abs(.PDFFontsSubSetFonts)), REG_SZ
51110   reg.SetRegistryValue "PDFFontsSubSetFontsPercent", CStr(.PDFFontsSubSetFontsPercent), REG_SZ
51120   reg.Subkey = "Printing\Formats\PDF\General"
51130   If Not reg.KeyExists Then
51140    reg.CreateKey
51150   End If
51160   reg.SetRegistryValue "PDFGeneralASCII85", CStr(Abs(.PDFGeneralASCII85)), REG_SZ
51170   reg.SetRegistryValue "PDFGeneralAutorotate", CStr(.PDFGeneralAutorotate), REG_SZ
51180   reg.SetRegistryValue "PDFGeneralCompatibility", CStr(.PDFGeneralCompatibility), REG_SZ
51190   reg.SetRegistryValue "PDFGeneralDefault", CStr(.PDFGeneralDefault), REG_SZ
51200   reg.SetRegistryValue "PDFGeneralOverprint", CStr(.PDFGeneralOverprint), REG_SZ
51210   reg.SetRegistryValue "PDFGeneralResolution", CStr(.PDFGeneralResolution), REG_SZ
51220   reg.SetRegistryValue "PDFOptimize", CStr(Abs(.PDFOptimize)), REG_SZ
51230   reg.SetRegistryValue "PDFUpdateMetadata", CStr(.PDFUpdateMetadata), REG_SZ
51240   reg.Subkey = "Printing\Formats\PDF\Security"
51250   If Not reg.KeyExists Then
51260    reg.CreateKey
51270   End If
51280   reg.SetRegistryValue "PDFAllowAssembly", CStr(Abs(.PDFAllowAssembly)), REG_SZ
51290   reg.SetRegistryValue "PDFAllowDegradedPrinting", CStr(Abs(.PDFAllowDegradedPrinting)), REG_SZ
51300   reg.SetRegistryValue "PDFAllowFillIn", CStr(Abs(.PDFAllowFillIn)), REG_SZ
51310   reg.SetRegistryValue "PDFAllowScreenReaders", CStr(Abs(.PDFAllowScreenReaders)), REG_SZ
51320   reg.SetRegistryValue "PDFDisallowCopy", CStr(Abs(.PDFDisallowCopy)), REG_SZ
51330   reg.SetRegistryValue "PDFDisallowModifyAnnotations", CStr(Abs(.PDFDisallowModifyAnnotations)), REG_SZ
51340   reg.SetRegistryValue "PDFDisallowModifyContents", CStr(Abs(.PDFDisallowModifyContents)), REG_SZ
51350   reg.SetRegistryValue "PDFDisallowPrinting", CStr(Abs(.PDFDisallowPrinting)), REG_SZ
51360   reg.SetRegistryValue "PDFEncryptor", CStr(.PDFEncryptor), REG_SZ
51370   reg.SetRegistryValue "PDFHighEncryption", CStr(Abs(.PDFHighEncryption)), REG_SZ
51380   reg.SetRegistryValue "PDFLowEncryption", CStr(Abs(.PDFLowEncryption)), REG_SZ
51390   reg.SetRegistryValue "PDFOwnerPass", CStr(Abs(.PDFOwnerPass)), REG_SZ
51400   reg.SetRegistryValue "PDFOwnerPasswordString", CStr(.PDFOwnerPasswordString), REG_SZ
51410   reg.SetRegistryValue "PDFUserPass", CStr(Abs(.PDFUserPass)), REG_SZ
51420   reg.SetRegistryValue "PDFUserPasswordString", CStr(.PDFUserPasswordString), REG_SZ
51430   reg.SetRegistryValue "PDFUseSecurity", CStr(Abs(.PDFUseSecurity)), REG_SZ
51440   reg.Subkey = "Printing\Formats\PDF\Signing"
51450   If Not reg.KeyExists Then
51460    reg.CreateKey
51470   End If
51480   reg.SetRegistryValue "PDFSigningMultiSignature", CStr(Abs(.PDFSigningMultiSignature)), REG_SZ
51490   reg.SetRegistryValue "PDFSigningPFXFile", CStr(.PDFSigningPFXFile), REG_SZ
51500   reg.SetRegistryValue "PDFSigningPFXFilePassword", CStr(.PDFSigningPFXFilePassword), REG_SZ
51510   reg.SetRegistryValue "PDFSigningSignatureContact", CStr(.PDFSigningSignatureContact), REG_SZ
51520   reg.SetRegistryValue "PDFSigningSignatureLeftX", Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), REG_SZ
51530   reg.SetRegistryValue "PDFSigningSignatureLeftY", Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), REG_SZ
51540   reg.SetRegistryValue "PDFSigningSignatureLocation", CStr(.PDFSigningSignatureLocation), REG_SZ
51550   reg.SetRegistryValue "PDFSigningSignatureReason", CStr(.PDFSigningSignatureReason), REG_SZ
51560   reg.SetRegistryValue "PDFSigningSignatureRightX", Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), REG_SZ
51570   reg.SetRegistryValue "PDFSigningSignatureRightY", Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), REG_SZ
51580   reg.SetRegistryValue "PDFSigningSignatureVisible", CStr(Abs(.PDFSigningSignatureVisible)), REG_SZ
51590   reg.SetRegistryValue "PDFSigningSignPDF", CStr(Abs(.PDFSigningSignPDF)), REG_SZ
51600   reg.Subkey = "Printing\Formats\PS\LanguageLevel"
51610   If Not reg.KeyExists Then
51620    reg.CreateKey
51630   End If
51640   reg.SetRegistryValue "EPSLanguageLevel", CStr(.EPSLanguageLevel), REG_SZ
51650   reg.SetRegistryValue "PSLanguageLevel", CStr(.PSLanguageLevel), REG_SZ
51660   reg.Subkey = "Program"
51670   If Not reg.KeyExists Then
51680    reg.CreateKey
51690   End If
51700   reg.SetRegistryValue "AdditionalGhostscriptParameters", CStr(.AdditionalGhostscriptParameters), REG_SZ
51710   reg.SetRegistryValue "AdditionalGhostscriptSearchpath", CStr(.AdditionalGhostscriptSearchpath), REG_SZ
51720   reg.SetRegistryValue "AddWindowsFontpath", CStr(Abs(.AddWindowsFontpath)), REG_SZ
51730   reg.SetRegistryValue "AutosaveDirectory", CStr(.AutosaveDirectory), REG_SZ
51740   reg.SetRegistryValue "AutosaveFilename", CStr(.AutosaveFilename), REG_SZ
51750   reg.SetRegistryValue "AutosaveFormat", CStr(.AutosaveFormat), REG_SZ
51760   reg.SetRegistryValue "AutosaveStartStandardProgram", CStr(Abs(.AutosaveStartStandardProgram)), REG_SZ
51770   reg.SetRegistryValue "ClientComputerResolveIPAddress", CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
51780   reg.SetRegistryValue "DisableEmail", CStr(Abs(.DisableEmail)), REG_SZ
51790   reg.SetRegistryValue "DontUseDocumentSettings", CStr(Abs(.DontUseDocumentSettings)), REG_SZ
51800   reg.SetRegistryValue "FilenameSubstitutions", CStr(.FilenameSubstitutions), REG_SZ
51810   reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle", CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
51820   reg.SetRegistryValue "Language", CStr(.Language), REG_SZ
51830   reg.SetRegistryValue "LastSaveDirectory", CStr(.LastSaveDirectory), REG_SZ
51840   reg.SetRegistryValue "LastUpdateCheck", CStr(.LastUpdateCheck), REG_SZ
51850   reg.SetRegistryValue "Logging", CStr(Abs(.Logging)), REG_SZ
51860   reg.SetRegistryValue "LogLines", CStr(.LogLines), REG_SZ
51870   reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter", CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
51880   reg.SetRegistryValue "NoProcessingAtStartup", CStr(Abs(.NoProcessingAtStartup)), REG_SZ
51890   reg.SetRegistryValue "NoPSCheck", CStr(Abs(.NoPSCheck)), REG_SZ
51900   reg.SetRegistryValue "OptionsDesign", CStr(.OptionsDesign), REG_SZ
51910   reg.SetRegistryValue "OptionsEnabled", CStr(Abs(.OptionsEnabled)), REG_SZ
51920   reg.SetRegistryValue "OptionsVisible", CStr(Abs(.OptionsVisible)), REG_SZ
51930   reg.SetRegistryValue "PrintAfterSaving", CStr(Abs(.PrintAfterSaving)), REG_SZ
51940   reg.SetRegistryValue "PrintAfterSavingDuplex", CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
51950   reg.SetRegistryValue "PrintAfterSavingNoCancel", CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
51960   reg.SetRegistryValue "PrintAfterSavingPrinter", CStr(.PrintAfterSavingPrinter), REG_SZ
51970   reg.SetRegistryValue "PrintAfterSavingQueryUser", CStr(.PrintAfterSavingQueryUser), REG_SZ
51980   reg.SetRegistryValue "PrintAfterSavingTumble", CStr(.PrintAfterSavingTumble), REG_SZ
51990   reg.SetRegistryValue "PrinterStop", CStr(Abs(.PrinterStop)), REG_SZ
52000   reg.SetRegistryValue "PrinterTemppath", CStr(.PrinterTemppath), REG_SZ
52010   reg.SetRegistryValue "ProcessPriority", CStr(.ProcessPriority), REG_SZ
52020   reg.SetRegistryValue "ProgramFont", CStr(.ProgramFont), REG_SZ
52030   reg.SetRegistryValue "ProgramFontCharset", CStr(.ProgramFontCharset), REG_SZ
52040   reg.SetRegistryValue "ProgramFontSize", CStr(.ProgramFontSize), REG_SZ
52050   reg.SetRegistryValue "RemoveAllKnownFileExtensions", CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
52060   reg.SetRegistryValue "RemoveSpaces", CStr(Abs(.RemoveSpaces)), REG_SZ
52070   reg.SetRegistryValue "RunProgramAfterSaving", CStr(Abs(.RunProgramAfterSaving)), REG_SZ
52080   reg.SetRegistryValue "RunProgramAfterSavingProgramname", CStr(.RunProgramAfterSavingProgramname), REG_SZ
52090   reg.SetRegistryValue "RunProgramAfterSavingProgramParameters", CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
52100   reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady", CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
52110   reg.SetRegistryValue "RunProgramAfterSavingWindowstyle", CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
52120   reg.SetRegistryValue "RunProgramBeforeSaving", CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
52130   reg.SetRegistryValue "RunProgramBeforeSavingProgramname", CStr(.RunProgramBeforeSavingProgramname), REG_SZ
52140   reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters", CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
52150   reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle", CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
52160   reg.SetRegistryValue "SaveFilename", CStr(.SaveFilename), REG_SZ
52170   reg.SetRegistryValue "SendEmailAfterAutoSaving", CStr(Abs(.SendEmailAfterAutoSaving)), REG_SZ
52180   reg.SetRegistryValue "SendMailMethod", CStr(.SendMailMethod), REG_SZ
52190   reg.SetRegistryValue "ShowAnimation", CStr(Abs(.ShowAnimation)), REG_SZ
52200   reg.SetRegistryValue "StartStandardProgram", CStr(Abs(.StartStandardProgram)), REG_SZ
52210   reg.SetRegistryValue "Toolbars", CStr(.Toolbars), REG_SZ
52220   reg.SetRegistryValue "UpdateInterval", CStr(.UpdateInterval), REG_SZ
52230   reg.SetRegistryValue "UseAutosave", CStr(Abs(.UseAutosave)), REG_SZ
52240   reg.SetRegistryValue "UseAutosaveDirectory", CStr(Abs(.UseAutosaveDirectory)), REG_SZ
52250  End With
52260  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "SaveOptionsREG")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetPrinterStop(StopPrinter As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If StopPrinter = True Then
50020    Options.PrinterStop = 1
50030    PrinterStop = True
50040    PrintSelectedJobs = False
50050   Else
50060    Options.PrinterStop = 0
50070    PrinterStop = False
50080  End If
50090  SaveOption Options, "Printerstop"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "SetPrinterStop")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetLogging(Logging As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Logging = True Then
50020    Options.Logging = 1
50030   Else
50040    Options.Logging = 0
50050  End If
50060  SaveOption Options, "Logging"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "SetLogging")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetLanguage(Language As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Options.Language = Language
50020  SaveOptions Options
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "SetLanguage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub ReadLanguageFromOptions(Optional hProfile As hkey = HKEY_CURRENT_USER)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim sLanguage As String
50020  If InstalledAsServer Then
50030    If UseINI Then
50040      sLanguage = ReadLanguageFromOptionsINI(sLanguage, CompletePath(GetCommonAppData) & "PDFCreator.ini")
50050     Else
50060      sLanguage = ReadLanguageFromOptionsReg(sLanguage, "Software\PDFCreator", HKEY_LOCAL_MACHINE)
50070    End If
50080   Else
50090    If UseINI Then
50100      If Not IsWin9xMe Then
50110        sLanguage = ReadLanguageFromOptionsINI(sLanguage, CompletePath(GetDefaultAppData) & "PDFCreator.ini")
50120        sLanguage = ReadLanguageFromOptionsINI(sLanguage, PDFCreatorINIFile, False)
50130       Else
50140        sLanguage = ReadLanguageFromOptionsINI(sLanguage, PDFCreatorINIFile)
50150      End If
50160      sLanguage = ReadLanguageFromOptionsINI(sLanguage, CompletePath(GetCommonAppData) & "PDFCreator.ini", False)
50170     Else
50180      If Not IsWin9xMe Then
50190        sLanguage = ReadLanguageFromOptionsReg(sLanguage, ".DEFAULT\Software\PDFCreator", HKEY_USERS)
50200        sLanguage = ReadLanguageFromOptionsReg(sLanguage, "Software\PDFCreator", hProfile, False)
50210       Else
50220        sLanguage = ReadLanguageFromOptionsReg(sLanguage, "Software\PDFCreator", hProfile)
50230      End If
50240      sLanguage = ReadLanguageFromOptionsReg(sLanguage, "Software\PDFCreator", HKEY_LOCAL_MACHINE, False)
50250    End If
50260  End If
50270  Options.Language = sLanguage
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "ReadLanguageFromOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function ReadLanguageFromOptionsINI(Language As String, PDFCreatorINIFile As String, Optional UseStandard As Boolean = True) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hOpt As clsHash, tStr As String, opt As tOptions
50020  ReadLanguageFromOptionsINI = Language
50030  If FileExists(PDFCreatorINIFile) = False Then
50040   If UseStandard Then
50050    opt = StandardOptions
50060    ReadLanguageFromOptionsINI = opt.Language
50070   End If
50080   Exit Function
50090  End If
50100  Set hOpt = New clsHash
50110  ReadINISection PDFCreatorINIFile, "Options", hOpt
50120  tStr = Trim$(hOpt.Retrieve("Language"))
50130  If LenB(tStr) > 0 Then
50140    ReadLanguageFromOptionsINI = tStr
50150   Else
50160    If UseStandard Then
50170      ReadLanguageFromOptionsINI = "english"
50180     Else
50190      ReadLanguageFromOptionsINI = Language
50200    End If
50210  End If
50220  Set hOpt = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "ReadLanguageFromOptionsINI")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ReadLanguageFromOptionsReg(Language As String, KeyRoot As String, Optional hProfile As hkey = HKEY_CURRENT_USER, Optional UseStandard As Boolean = True) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, tStr As String
50020  Set reg = New clsRegistry
50030  With reg
50040   .KeyRoot = KeyRoot
50050   .Subkey = "Program"
50060   .hkey = hProfile
50070   tStr = Trim$(reg.GetRegistryValue("Language"))
50080  End With
50090  If LenB(tStr) > 0 Then
50100    ReadLanguageFromOptionsReg = tStr
50110   Else
50120    If UseStandard Then
50130      ReadLanguageFromOptionsReg = "english"
50140     Else
50150      ReadLanguageFromOptionsReg = Language
50160    End If
50170  End If
50180  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "ReadLanguageFromOptionsReg")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function UseINI() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, tStr As String
50020  Set reg = New clsRegistry
50030  UseINI = False
50040  With reg
50050   .hkey = HKEY_LOCAL_MACHINE
50060   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50070   tStr = Trim$(.GetRegistryValue("UseINI"))
50080   If tStr = "1" Then
50090    UseINI = True
50100   End If
50110  End With
50120  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "UseINI")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

