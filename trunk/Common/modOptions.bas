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
 BitmapResolution As Long
 BMPColorscount As Long
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
 Language As String
 LastSaveDirectory As String
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
 PCXColorscount As Long
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
 PSLanguageLevel As Long
 RAWColorsCount As Long
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
 Toolbars As Long
 UseAutosave As Long
 UseAutosaveDirectory As Long
 UseCreationDateNow As Long
 UseCustomPaperSize As String
 UseFixPapersize As Long
 UseStandardAuthor As Long
 XCFColorsCount As Long
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
50150   .BitmapResolution = "150"
50160   .BMPColorscount = "1"
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
50440   .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
50450   .FilenameSubstitutionsOnlyInTitle = "1"
50460   .JPEGColorscount = "0"
50470   .JPEGQuality = "75"
50480   .Language = "english"
50490   If InstalledAsServer Then
50500     .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
50510    Else
50520     .LastSaveDirectory = "<MyFiles>"
50530   End If
50540   .Logging = "0"
50550   .LogLines = "100"
50560   .NoConfirmMessageSwitchingDefaultprinter = "0"
50570   .NoProcessingAtStartup = "0"
50580   .NoPSCheck = "0"
50590   .OnePagePerFile = "0"
50600   .OptionsDesign = "0"
50610   .OptionsEnabled = "1"
50620   .OptionsVisible = "1"
50630   .Papersize = "a4"
50640   .PCLColorsCount = "0"
50650   .PCXColorscount = "0"
50660   .PDFAllowAssembly = "0"
50670   .PDFAllowDegradedPrinting = "0"
50680   .PDFAllowFillIn = "0"
50690   .PDFAllowScreenReaders = "0"
50700   .PDFColorsCMYKToRGB = "0"
50710   .PDFColorsColorModel = "1"
50720   .PDFColorsPreserveHalftone = "0"
50730   .PDFColorsPreserveOverprint = "1"
50740   .PDFColorsPreserveTransfer = "1"
50750   .PDFCompressionColorCompression = "1"
50760   .PDFCompressionColorCompressionChoice = "0"
50770   .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
50780   .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
50790   .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
50800   .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
50810   .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
50820   .PDFCompressionColorResample = "0"
50830   .PDFCompressionColorResampleChoice = "0"
50840   .PDFCompressionColorResolution = "300"
50850   .PDFCompressionGreyCompression = "1"
50860   .PDFCompressionGreyCompressionChoice = "0"
50870   .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
50880   .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
50890   .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
50900   .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
50910   .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
50920   .PDFCompressionGreyResample = "0"
50930   .PDFCompressionGreyResampleChoice = "0"
50940   .PDFCompressionGreyResolution = "300"
50950   .PDFCompressionMonoCompression = "1"
50960   .PDFCompressionMonoCompressionChoice = "0"
50970   .PDFCompressionMonoResample = "0"
50980   .PDFCompressionMonoResampleChoice = "0"
50990   .PDFCompressionMonoResolution = "1200"
51000   .PDFCompressionTextCompression = "1"
51010   .PDFDisallowCopy = "1"
51020   .PDFDisallowModifyAnnotations = "0"
51030   .PDFDisallowModifyContents = "0"
51040   .PDFDisallowPrinting = "0"
51050   .PDFEncryptor = "0"
51060   .PDFFontsEmbedAll = "1"
51070   .PDFFontsSubSetFonts = "1"
51080   .PDFFontsSubSetFontsPercent = "100"
51090   .PDFGeneralASCII85 = "0"
51100   .PDFGeneralAutorotate = "2"
51110   .PDFGeneralCompatibility = "2"
51120   .PDFGeneralDefault = "0"
51130   .PDFGeneralOverprint = "0"
51140   .PDFGeneralResolution = "600"
51150   .PDFHighEncryption = "0"
51160   .PDFLowEncryption = "1"
51170   .PDFOptimize = "0"
51180   .PDFOwnerPass = "0"
51190   .PDFOwnerPasswordString = vbNullString
51200   .PDFSigningMultiSignature = "0"
51210   .PDFSigningPFXFile = vbNullString
51220   .PDFSigningPFXFilePassword = vbNullString
51230   .PDFSigningSignatureContact = vbNullString
51240   .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
51250   .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
51260   .PDFSigningSignatureLocation = vbNullString
51270   .PDFSigningSignatureReason = vbNullString
51280   .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
51290   .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
51300   .PDFSigningSignatureVisible = "0"
51310   .PDFSigningSignPDF = "0"
51320   .PDFUpdateMetadata = "1"
51330   .PDFUserPass = "0"
51340   .PDFUserPasswordString = vbNullString
51350   .PDFUseSecurity = "0"
51360   .PNGColorscount = "0"
51370   .PrintAfterSaving = "0"
51380   .PrintAfterSavingDuplex = "0"
51390   .PrintAfterSavingNoCancel = "0"
51400   .PrintAfterSavingPrinter = vbNullString
51410   .PrintAfterSavingQueryUser = "0"
51420   .PrintAfterSavingTumble = "0"
51430   .PrinterStop = "0"
51440   If InstalledAsServer Then
51450     .PrinterTemppath = CompletePath(GetPDFCreatorApplicationPath) & "Temp\"
51460    Else
51470     .PrinterTemppath = "<Temp>PDFCreator\"
51480   End If
51490   .ProcessPriority = "1"
51500   .ProgramFont = "MS Sans Serif"
51510   .ProgramFontCharset = "0"
51520   .ProgramFontSize = "8"
51530   .PSDColorsCount = "0"
51540   .PSLanguageLevel = "2"
51550   .RAWColorsCount = "0"
51560   .RemoveAllKnownFileExtensions = "1"
51570   .RemoveSpaces = "1"
51580   .RunProgramAfterSaving = "0"
51590   .RunProgramAfterSavingProgramname = vbNullString
51600   .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
51610   .RunProgramAfterSavingWaitUntilReady = "1"
51620   .RunProgramAfterSavingWindowstyle = "1"
51630   .RunProgramBeforeSaving = "0"
51640   .RunProgramBeforeSavingProgramname = vbNullString
51650   .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
51660   .RunProgramBeforeSavingWindowstyle = "1"
51670   .SaveFilename = "<Title>"
51680   .SendEmailAfterAutoSaving = "0"
51690   .SendMailMethod = "0"
51700   .ShowAnimation = "1"
51710   .StampFontColor = "#FF0000"
51720   .StampFontname = "Arial"
51730   .StampFontsize = "48"
51740   .StampOutlineFontthickness = "0"
51750   .StampString = vbNullString
51760   .StampUseOutlineFont = "1"
51770   .StandardAuthor = vbNullString
51780   .StandardCreationdate = vbNullString
51790   .StandardDateformat = "YYYYMMDDHHNNSS"
51800   .StandardKeywords = vbNullString
51810   .StandardMailDomain = vbNullString
51820   .StandardModifydate = vbNullString
51830   .StandardSaveformat = "0"
51840   .StandardSubject = vbNullString
51850   .StandardTitle = vbNullString
51860   .StartStandardProgram = "1"
51870   .TIFFColorscount = "0"
51880   .Toolbars = "1"
51890   .UseAutosave = "0"
51900   .UseAutosaveDirectory = "1"
51910   .UseCreationDateNow = "0"
51920   .UseCustomPaperSize = "0"
51930   .UseFixPapersize = "0"
51940   .UseStandardAuthor = "0"
51950   .XCFColorsCount = "0"
51960  End With
51970  If UseINI Then
51980    If Not IsWin9xMe Then
51990     myOptions = ReadOptionsINI(myOptions, CompletePath(GetDefaultAppData) & "PDFCreator.ini", False, False, False)
52000    End If
52010   Else
52020    If Not IsWin9xMe Then
52030     myOptions = ReadOptionsReg(myOptions, ".DEFAULT\Software\PDFCreator", HKEY_USERS, False, False)
52040    End If
52050  End If
52060  StandardOptions = myOptions
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
50920   tStr = hOpt.Retrieve("BitmapResolution")
50930   If IsNumeric(tStr) Then
50940     If CLng(tStr) >= 1 Then
50950       .BitmapResolution = CLng(tStr)
50960      Else
50970       If UseStandard Then
50980        .BitmapResolution = 150
50990       End If
51000     End If
51010    Else
51020     If UseStandard Then
51030      .BitmapResolution = 150
51040     End If
51050   End If
51060   tStr = hOpt.Retrieve("BMPColorscount")
51070   If IsNumeric(tStr) Then
51080     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
51090       .BMPColorscount = CLng(tStr)
51100      Else
51110       If UseStandard Then
51120        .BMPColorscount = 1
51130       End If
51140     End If
51150    Else
51160     If UseStandard Then
51170      .BMPColorscount = 1
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
52540   If LenB(tStr) = 0 And LenB("Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt") > 0 And UseStandard Then
52550     .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
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
53030   tStr = hOpt.Retrieve("Language")
53040   If LenB(tStr) = 0 And LenB("english") > 0 And UseStandard Then
53050     .Language = "english"
53060    Else
53070     If LenB(tStr) > 0 Then
53080      .Language = tStr
53090     End If
53100   End If
53110   tStr = hOpt.Retrieve("LastSaveDirectory")
53120   If LenB(Trim$(tStr)) > 0 Then
53130     .LastSaveDirectory = CompletePath(tStr)
53140    Else
53150     If UseStandard Then
53160      If InstalledAsServer Then
53170        .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
53180       Else
53190        .LastSaveDirectory = "<MyFiles>"
53200      End If
53210     End If
53220   End If
53230   tStr = hOpt.Retrieve("Logging")
53240   If IsNumeric(tStr) Then
53250     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53260       .Logging = CLng(tStr)
53270      Else
53280       If UseStandard Then
53290        .Logging = 0
53300       End If
53310     End If
53320    Else
53330     If UseStandard Then
53340      .Logging = 0
53350     End If
53360   End If
53370   tStr = hOpt.Retrieve("LogLines")
53380   If IsNumeric(tStr) Then
53390     If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
53400       .LogLines = CLng(tStr)
53410      Else
53420       If UseStandard Then
53430        .LogLines = 100
53440       End If
53450     End If
53460    Else
53470     If UseStandard Then
53480      .LogLines = 100
53490     End If
53500   End If
53510   tStr = hOpt.Retrieve("NoConfirmMessageSwitchingDefaultprinter")
53520   If IsNumeric(tStr) Then
53530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53540       .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
53550      Else
53560       If UseStandard Then
53570        .NoConfirmMessageSwitchingDefaultprinter = 0
53580       End If
53590     End If
53600    Else
53610     If UseStandard Then
53620      .NoConfirmMessageSwitchingDefaultprinter = 0
53630     End If
53640   End If
53650   tStr = hOpt.Retrieve("NoProcessingAtStartup")
53660   If IsNumeric(tStr) Then
53670     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53680       .NoProcessingAtStartup = CLng(tStr)
53690      Else
53700       If UseStandard Then
53710        .NoProcessingAtStartup = 0
53720       End If
53730     End If
53740    Else
53750     If UseStandard Then
53760      .NoProcessingAtStartup = 0
53770     End If
53780   End If
53790   tStr = hOpt.Retrieve("NoPSCheck")
53800   If IsNumeric(tStr) Then
53810     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53820       .NoPSCheck = CLng(tStr)
53830      Else
53840       If UseStandard Then
53850        .NoPSCheck = 0
53860       End If
53870     End If
53880    Else
53890     If UseStandard Then
53900      .NoPSCheck = 0
53910     End If
53920   End If
53930   tStr = hOpt.Retrieve("OnePagePerFile")
53940   If IsNumeric(tStr) Then
53950     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53960       .OnePagePerFile = CLng(tStr)
53970      Else
53980       If UseStandard Then
53990        .OnePagePerFile = 0
54000       End If
54010     End If
54020    Else
54030     If UseStandard Then
54040      .OnePagePerFile = 0
54050     End If
54060   End If
54070   tStr = hOpt.Retrieve("OptionsDesign")
54080   If IsNumeric(tStr) Then
54090     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
54100       .OptionsDesign = CLng(tStr)
54110      Else
54120       If UseStandard Then
54130        .OptionsDesign = 0
54140       End If
54150     End If
54160    Else
54170     If UseStandard Then
54180      .OptionsDesign = 0
54190     End If
54200   End If
54210   tStr = hOpt.Retrieve("OptionsEnabled")
54220   If IsNumeric(tStr) Then
54230     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54240       .OptionsEnabled = CLng(tStr)
54250      Else
54260       If UseStandard Then
54270        .OptionsEnabled = 1
54280       End If
54290     End If
54300    Else
54310     If UseStandard Then
54320      .OptionsEnabled = 1
54330     End If
54340   End If
54350   tStr = hOpt.Retrieve("OptionsVisible")
54360   If IsNumeric(tStr) Then
54370     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54380       .OptionsVisible = CLng(tStr)
54390      Else
54400       If UseStandard Then
54410        .OptionsVisible = 1
54420       End If
54430     End If
54440    Else
54450     If UseStandard Then
54460      .OptionsVisible = 1
54470     End If
54480   End If
54490   tStr = hOpt.Retrieve("Papersize")
54500   If LenB(tStr) = 0 And LenB("a4") > 0 And UseStandard Then
54510     .Papersize = "a4"
54520    Else
54530     If LenB(tStr) > 0 Then
54540      .Papersize = tStr
54550     End If
54560   End If
54570   tStr = hOpt.Retrieve("PCLColorsCount")
54580   If IsNumeric(tStr) Then
54590     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
54600       .PCLColorsCount = CLng(tStr)
54610      Else
54620       If UseStandard Then
54630        .PCLColorsCount = 0
54640       End If
54650     End If
54660    Else
54670     If UseStandard Then
54680      .PCLColorsCount = 0
54690     End If
54700   End If
54710   tStr = hOpt.Retrieve("PCXColorscount")
54720   If IsNumeric(tStr) Then
54730     If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
54740       .PCXColorscount = CLng(tStr)
54750      Else
54760       If UseStandard Then
54770        .PCXColorscount = 0
54780       End If
54790     End If
54800    Else
54810     If UseStandard Then
54820      .PCXColorscount = 0
54830     End If
54840   End If
54850   tStr = hOpt.Retrieve("PDFAllowAssembly")
54860   If IsNumeric(tStr) Then
54870     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54880       .PDFAllowAssembly = CLng(tStr)
54890      Else
54900       If UseStandard Then
54910        .PDFAllowAssembly = 0
54920       End If
54930     End If
54940    Else
54950     If UseStandard Then
54960      .PDFAllowAssembly = 0
54970     End If
54980   End If
54990   tStr = hOpt.Retrieve("PDFAllowDegradedPrinting")
55000   If IsNumeric(tStr) Then
55010     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55020       .PDFAllowDegradedPrinting = CLng(tStr)
55030      Else
55040       If UseStandard Then
55050        .PDFAllowDegradedPrinting = 0
55060       End If
55070     End If
55080    Else
55090     If UseStandard Then
55100      .PDFAllowDegradedPrinting = 0
55110     End If
55120   End If
55130   tStr = hOpt.Retrieve("PDFAllowFillIn")
55140   If IsNumeric(tStr) Then
55150     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55160       .PDFAllowFillIn = CLng(tStr)
55170      Else
55180       If UseStandard Then
55190        .PDFAllowFillIn = 0
55200       End If
55210     End If
55220    Else
55230     If UseStandard Then
55240      .PDFAllowFillIn = 0
55250     End If
55260   End If
55270   tStr = hOpt.Retrieve("PDFAllowScreenReaders")
55280   If IsNumeric(tStr) Then
55290     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55300       .PDFAllowScreenReaders = CLng(tStr)
55310      Else
55320       If UseStandard Then
55330        .PDFAllowScreenReaders = 0
55340       End If
55350     End If
55360    Else
55370     If UseStandard Then
55380      .PDFAllowScreenReaders = 0
55390     End If
55400   End If
55410   tStr = hOpt.Retrieve("PDFColorsCMYKToRGB")
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
55550   tStr = hOpt.Retrieve("PDFColorsColorModel")
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
55690   tStr = hOpt.Retrieve("PDFColorsPreserveHalftone")
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
55830   tStr = hOpt.Retrieve("PDFColorsPreserveOverprint")
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
55970   tStr = hOpt.Retrieve("PDFColorsPreserveTransfer")
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
56110   tStr = hOpt.Retrieve("PDFCompressionColorCompression")
56120   If IsNumeric(tStr) Then
56130     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56140       .PDFCompressionColorCompression = CLng(tStr)
56150      Else
56160       If UseStandard Then
56170        .PDFCompressionColorCompression = 1
56180       End If
56190     End If
56200    Else
56210     If UseStandard Then
56220      .PDFCompressionColorCompression = 1
56230     End If
56240   End If
56250   tStr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
56260   If IsNumeric(tStr) Then
56270     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
56280       .PDFCompressionColorCompressionChoice = CLng(tStr)
56290      Else
56300       If UseStandard Then
56310        .PDFCompressionColorCompressionChoice = 0
56320       End If
56330     End If
56340    Else
56350     If UseStandard Then
56360      .PDFCompressionColorCompressionChoice = 0
56370     End If
56380   End If
56390   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGHighFactor")
56400   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56410     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56420       .PDFCompressionColorCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56430      Else
56440       If UseStandard Then
56450        .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56460       End If
56470     End If
56480    Else
56490     If UseStandard Then
56500      .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56510     End If
56520   End If
56530   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGLowFactor")
56540   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56550     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56560       .PDFCompressionColorCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56570      Else
56580       If UseStandard Then
56590        .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56600       End If
56610     End If
56620    Else
56630     If UseStandard Then
56640      .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56650     End If
56660   End If
56670   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMaximumFactor")
56680   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56690     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56700       .PDFCompressionColorCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56710      Else
56720       If UseStandard Then
56730        .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56740       End If
56750     End If
56760    Else
56770     If UseStandard Then
56780      .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56790     End If
56800   End If
56810   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMediumFactor")
56820   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56830     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56840       .PDFCompressionColorCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56850      Else
56860       If UseStandard Then
56870        .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56880       End If
56890     End If
56900    Else
56910     If UseStandard Then
56920      .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56930     End If
56940   End If
56950   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMinimumFactor")
56960   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56970     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56980       .PDFCompressionColorCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56990      Else
57000       If UseStandard Then
57010        .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57020       End If
57030     End If
57040    Else
57050     If UseStandard Then
57060      .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57070     End If
57080   End If
57090   tStr = hOpt.Retrieve("PDFCompressionColorResample")
57100   If IsNumeric(tStr) Then
57110     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57120       .PDFCompressionColorResample = CLng(tStr)
57130      Else
57140       If UseStandard Then
57150        .PDFCompressionColorResample = 0
57160       End If
57170     End If
57180    Else
57190     If UseStandard Then
57200      .PDFCompressionColorResample = 0
57210     End If
57220   End If
57230   tStr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
57240   If IsNumeric(tStr) Then
57250     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
57260       .PDFCompressionColorResampleChoice = CLng(tStr)
57270      Else
57280       If UseStandard Then
57290        .PDFCompressionColorResampleChoice = 0
57300       End If
57310     End If
57320    Else
57330     If UseStandard Then
57340      .PDFCompressionColorResampleChoice = 0
57350     End If
57360   End If
57370   tStr = hOpt.Retrieve("PDFCompressionColorResolution")
57380   If IsNumeric(tStr) Then
57390     If CLng(tStr) >= 0 Then
57400       .PDFCompressionColorResolution = CLng(tStr)
57410      Else
57420       If UseStandard Then
57430        .PDFCompressionColorResolution = 300
57440       End If
57450     End If
57460    Else
57470     If UseStandard Then
57480      .PDFCompressionColorResolution = 300
57490     End If
57500   End If
57510   tStr = hOpt.Retrieve("PDFCompressionGreyCompression")
57520   If IsNumeric(tStr) Then
57530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57540       .PDFCompressionGreyCompression = CLng(tStr)
57550      Else
57560       If UseStandard Then
57570        .PDFCompressionGreyCompression = 1
57580       End If
57590     End If
57600    Else
57610     If UseStandard Then
57620      .PDFCompressionGreyCompression = 1
57630     End If
57640   End If
57650   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
57660   If IsNumeric(tStr) Then
57670     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
57680       .PDFCompressionGreyCompressionChoice = CLng(tStr)
57690      Else
57700       If UseStandard Then
57710        .PDFCompressionGreyCompressionChoice = 0
57720       End If
57730     End If
57740    Else
57750     If UseStandard Then
57760      .PDFCompressionGreyCompressionChoice = 0
57770     End If
57780   End If
57790   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGHighFactor")
57800   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57810     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57820       .PDFCompressionGreyCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57830      Else
57840       If UseStandard Then
57850        .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
57860       End If
57870     End If
57880    Else
57890     If UseStandard Then
57900      .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
57910     End If
57920   End If
57930   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGLowFactor")
57940   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57950     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57960       .PDFCompressionGreyCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57970      Else
57980       If UseStandard Then
57990        .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58000       End If
58010     End If
58020    Else
58030     If UseStandard Then
58040      .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58050     End If
58060   End If
58070   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMaximumFactor")
58080   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58090     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58100       .PDFCompressionGreyCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58110      Else
58120       If UseStandard Then
58130        .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
58140       End If
58150     End If
58160    Else
58170     If UseStandard Then
58180      .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
58190     End If
58200   End If
58210   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMediumFactor")
58220   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58230     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58240       .PDFCompressionGreyCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58250      Else
58260       If UseStandard Then
58270        .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58280       End If
58290     End If
58300    Else
58310     If UseStandard Then
58320      .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58330     End If
58340   End If
58350   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMinimumFactor")
58360   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58370     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58380       .PDFCompressionGreyCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58390      Else
58400       If UseStandard Then
58410        .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58420       End If
58430     End If
58440    Else
58450     If UseStandard Then
58460      .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58470     End If
58480   End If
58490   tStr = hOpt.Retrieve("PDFCompressionGreyResample")
58500   If IsNumeric(tStr) Then
58510     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58520       .PDFCompressionGreyResample = CLng(tStr)
58530      Else
58540       If UseStandard Then
58550        .PDFCompressionGreyResample = 0
58560       End If
58570     End If
58580    Else
58590     If UseStandard Then
58600      .PDFCompressionGreyResample = 0
58610     End If
58620   End If
58630   tStr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
58640   If IsNumeric(tStr) Then
58650     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
58660       .PDFCompressionGreyResampleChoice = CLng(tStr)
58670      Else
58680       If UseStandard Then
58690        .PDFCompressionGreyResampleChoice = 0
58700       End If
58710     End If
58720    Else
58730     If UseStandard Then
58740      .PDFCompressionGreyResampleChoice = 0
58750     End If
58760   End If
58770   tStr = hOpt.Retrieve("PDFCompressionGreyResolution")
58780   If IsNumeric(tStr) Then
58790     If CLng(tStr) >= 0 Then
58800       .PDFCompressionGreyResolution = CLng(tStr)
58810      Else
58820       If UseStandard Then
58830        .PDFCompressionGreyResolution = 300
58840       End If
58850     End If
58860    Else
58870     If UseStandard Then
58880      .PDFCompressionGreyResolution = 300
58890     End If
58900   End If
58910   tStr = hOpt.Retrieve("PDFCompressionMonoCompression")
58920   If IsNumeric(tStr) Then
58930     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58940       .PDFCompressionMonoCompression = CLng(tStr)
58950      Else
58960       If UseStandard Then
58970        .PDFCompressionMonoCompression = 1
58980       End If
58990     End If
59000    Else
59010     If UseStandard Then
59020      .PDFCompressionMonoCompression = 1
59030     End If
59040   End If
59050   tStr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
59060   If IsNumeric(tStr) Then
59070     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
59080       .PDFCompressionMonoCompressionChoice = CLng(tStr)
59090      Else
59100       If UseStandard Then
59110        .PDFCompressionMonoCompressionChoice = 0
59120       End If
59130     End If
59140    Else
59150     If UseStandard Then
59160      .PDFCompressionMonoCompressionChoice = 0
59170     End If
59180   End If
59190   tStr = hOpt.Retrieve("PDFCompressionMonoResample")
59200   If IsNumeric(tStr) Then
59210     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59220       .PDFCompressionMonoResample = CLng(tStr)
59230      Else
59240       If UseStandard Then
59250        .PDFCompressionMonoResample = 0
59260       End If
59270     End If
59280    Else
59290     If UseStandard Then
59300      .PDFCompressionMonoResample = 0
59310     End If
59320   End If
59330   tStr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
59340   If IsNumeric(tStr) Then
59350     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
59360       .PDFCompressionMonoResampleChoice = CLng(tStr)
59370      Else
59380       If UseStandard Then
59390        .PDFCompressionMonoResampleChoice = 0
59400       End If
59410     End If
59420    Else
59430     If UseStandard Then
59440      .PDFCompressionMonoResampleChoice = 0
59450     End If
59460   End If
59470   tStr = hOpt.Retrieve("PDFCompressionMonoResolution")
59480   If IsNumeric(tStr) Then
59490     If CLng(tStr) >= 0 Then
59500       .PDFCompressionMonoResolution = CLng(tStr)
59510      Else
59520       If UseStandard Then
59530        .PDFCompressionMonoResolution = 1200
59540       End If
59550     End If
59560    Else
59570     If UseStandard Then
59580      .PDFCompressionMonoResolution = 1200
59590     End If
59600   End If
59610   tStr = hOpt.Retrieve("PDFCompressionTextCompression")
59620   If IsNumeric(tStr) Then
59630     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59640       .PDFCompressionTextCompression = CLng(tStr)
59650      Else
59660       If UseStandard Then
59670        .PDFCompressionTextCompression = 1
59680       End If
59690     End If
59700    Else
59710     If UseStandard Then
59720      .PDFCompressionTextCompression = 1
59730     End If
59740   End If
59750   tStr = hOpt.Retrieve("PDFDisallowCopy")
59760   If IsNumeric(tStr) Then
59770     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59780       .PDFDisallowCopy = CLng(tStr)
59790      Else
59800       If UseStandard Then
59810        .PDFDisallowCopy = 1
59820       End If
59830     End If
59840    Else
59850     If UseStandard Then
59860      .PDFDisallowCopy = 1
59870     End If
59880   End If
59890   tStr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
59900   If IsNumeric(tStr) Then
59910     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59920       .PDFDisallowModifyAnnotations = CLng(tStr)
59930      Else
59940       If UseStandard Then
59950        .PDFDisallowModifyAnnotations = 0
59960       End If
59970     End If
59980    Else
59990     If UseStandard Then
60000      .PDFDisallowModifyAnnotations = 0
60010     End If
60020   End If
60030   tStr = hOpt.Retrieve("PDFDisallowModifyContents")
60040   If IsNumeric(tStr) Then
60050     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60060       .PDFDisallowModifyContents = CLng(tStr)
60070      Else
60080       If UseStandard Then
60090        .PDFDisallowModifyContents = 0
60100       End If
60110     End If
60120    Else
60130     If UseStandard Then
60140      .PDFDisallowModifyContents = 0
60150     End If
60160   End If
60170   tStr = hOpt.Retrieve("PDFDisallowPrinting")
60180   If IsNumeric(tStr) Then
60190     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60200       .PDFDisallowPrinting = CLng(tStr)
60210      Else
60220       If UseStandard Then
60230        .PDFDisallowPrinting = 0
60240       End If
60250     End If
60260    Else
60270     If UseStandard Then
60280      .PDFDisallowPrinting = 0
60290     End If
60300   End If
60310   tStr = hOpt.Retrieve("PDFEncryptor")
60320   If IsNumeric(tStr) Then
60330     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
60340       .PDFEncryptor = CLng(tStr)
60350      Else
60360       If UseStandard Then
60370        .PDFEncryptor = 0
60380       End If
60390     End If
60400    Else
60410     If UseStandard Then
60420      .PDFEncryptor = 0
60430     End If
60440   End If
60450   tStr = hOpt.Retrieve("PDFFontsEmbedAll")
60460   If IsNumeric(tStr) Then
60470     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60480       .PDFFontsEmbedAll = CLng(tStr)
60490      Else
60500       If UseStandard Then
60510        .PDFFontsEmbedAll = 1
60520       End If
60530     End If
60540    Else
60550     If UseStandard Then
60560      .PDFFontsEmbedAll = 1
60570     End If
60580   End If
60590   tStr = hOpt.Retrieve("PDFFontsSubSetFonts")
60600   If IsNumeric(tStr) Then
60610     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60620       .PDFFontsSubSetFonts = CLng(tStr)
60630      Else
60640       If UseStandard Then
60650        .PDFFontsSubSetFonts = 1
60660       End If
60670     End If
60680    Else
60690     If UseStandard Then
60700      .PDFFontsSubSetFonts = 1
60710     End If
60720   End If
60730   tStr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
60740   If IsNumeric(tStr) Then
60750     If CLng(tStr) >= 0 Then
60760       .PDFFontsSubSetFontsPercent = CLng(tStr)
60770      Else
60780       If UseStandard Then
60790        .PDFFontsSubSetFontsPercent = 100
60800       End If
60810     End If
60820    Else
60830     If UseStandard Then
60840      .PDFFontsSubSetFontsPercent = 100
60850     End If
60860   End If
60870   tStr = hOpt.Retrieve("PDFGeneralASCII85")
60880   If IsNumeric(tStr) Then
60890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60900       .PDFGeneralASCII85 = CLng(tStr)
60910      Else
60920       If UseStandard Then
60930        .PDFGeneralASCII85 = 0
60940       End If
60950     End If
60960    Else
60970     If UseStandard Then
60980      .PDFGeneralASCII85 = 0
60990     End If
61000   End If
61010   tStr = hOpt.Retrieve("PDFGeneralAutorotate")
61020   If IsNumeric(tStr) Then
61030     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
61040       .PDFGeneralAutorotate = CLng(tStr)
61050      Else
61060       If UseStandard Then
61070        .PDFGeneralAutorotate = 2
61080       End If
61090     End If
61100    Else
61110     If UseStandard Then
61120      .PDFGeneralAutorotate = 2
61130     End If
61140   End If
61150   tStr = hOpt.Retrieve("PDFGeneralCompatibility")
61160   If IsNumeric(tStr) Then
61170     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
61180       .PDFGeneralCompatibility = CLng(tStr)
61190      Else
61200       If UseStandard Then
61210        .PDFGeneralCompatibility = 2
61220       End If
61230     End If
61240    Else
61250     If UseStandard Then
61260      .PDFGeneralCompatibility = 2
61270     End If
61280   End If
61290   tStr = hOpt.Retrieve("PDFGeneralDefault")
61300   If IsNumeric(tStr) Then
61310     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
61320       .PDFGeneralDefault = CLng(tStr)
61330      Else
61340       If UseStandard Then
61350        .PDFGeneralDefault = 0
61360       End If
61370     End If
61380    Else
61390     If UseStandard Then
61400      .PDFGeneralDefault = 0
61410     End If
61420   End If
61430   tStr = hOpt.Retrieve("PDFGeneralOverprint")
61440   If IsNumeric(tStr) Then
61450     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
61460       .PDFGeneralOverprint = CLng(tStr)
61470      Else
61480       If UseStandard Then
61490        .PDFGeneralOverprint = 0
61500       End If
61510     End If
61520    Else
61530     If UseStandard Then
61540      .PDFGeneralOverprint = 0
61550     End If
61560   End If
61570   tStr = hOpt.Retrieve("PDFGeneralResolution")
61580   If IsNumeric(tStr) Then
61590     If CLng(tStr) >= 0 Then
61600       .PDFGeneralResolution = CLng(tStr)
61610      Else
61620       If UseStandard Then
61630        .PDFGeneralResolution = 600
61640       End If
61650     End If
61660    Else
61670     If UseStandard Then
61680      .PDFGeneralResolution = 600
61690     End If
61700   End If
61710   tStr = hOpt.Retrieve("PDFHighEncryption")
61720   If IsNumeric(tStr) Then
61730     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61740       .PDFHighEncryption = CLng(tStr)
61750      Else
61760       If UseStandard Then
61770        .PDFHighEncryption = 0
61780       End If
61790     End If
61800    Else
61810     If UseStandard Then
61820      .PDFHighEncryption = 0
61830     End If
61840   End If
61850   tStr = hOpt.Retrieve("PDFLowEncryption")
61860   If IsNumeric(tStr) Then
61870     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61880       .PDFLowEncryption = CLng(tStr)
61890      Else
61900       If UseStandard Then
61910        .PDFLowEncryption = 1
61920       End If
61930     End If
61940    Else
61950     If UseStandard Then
61960      .PDFLowEncryption = 1
61970     End If
61980   End If
61990   tStr = hOpt.Retrieve("PDFOptimize")
62000   If IsNumeric(tStr) Then
62010     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62020       .PDFOptimize = CLng(tStr)
62030      Else
62040       If UseStandard Then
62050        .PDFOptimize = 0
62060       End If
62070     End If
62080    Else
62090     If UseStandard Then
62100      .PDFOptimize = 0
62110     End If
62120   End If
62130   tStr = hOpt.Retrieve("PDFOwnerPass")
62140   If IsNumeric(tStr) Then
62150     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62160       .PDFOwnerPass = CLng(tStr)
62170      Else
62180       If UseStandard Then
62190        .PDFOwnerPass = 0
62200       End If
62210     End If
62220    Else
62230     If UseStandard Then
62240      .PDFOwnerPass = 0
62250     End If
62260   End If
62270   tStr = hOpt.Retrieve("PDFOwnerPasswordString")
62280   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62290     .PDFOwnerPasswordString = ""
62300    Else
62310     If LenB(tStr) > 0 Then
62320      .PDFOwnerPasswordString = tStr
62330     End If
62340   End If
62350   tStr = hOpt.Retrieve("PDFSigningMultiSignature")
62360   If IsNumeric(tStr) Then
62370     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62380       .PDFSigningMultiSignature = CLng(tStr)
62390      Else
62400       If UseStandard Then
62410        .PDFSigningMultiSignature = 0
62420       End If
62430     End If
62440    Else
62450     If UseStandard Then
62460      .PDFSigningMultiSignature = 0
62470     End If
62480   End If
62490   tStr = hOpt.Retrieve("PDFSigningPFXFile")
62500   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62510     .PDFSigningPFXFile = ""
62520    Else
62530     If LenB(tStr) > 0 Then
62540      .PDFSigningPFXFile = tStr
62550     End If
62560   End If
62570   tStr = hOpt.Retrieve("PDFSigningPFXFilePassword")
62580   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62590     .PDFSigningPFXFilePassword = ""
62600    Else
62610     If LenB(tStr) > 0 Then
62620      .PDFSigningPFXFilePassword = tStr
62630     End If
62640   End If
62650   tStr = hOpt.Retrieve("PDFSigningSignatureContact")
62660   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62670     .PDFSigningSignatureContact = ""
62680    Else
62690     If LenB(tStr) > 0 Then
62700      .PDFSigningSignatureContact = tStr
62710     End If
62720   End If
62730   tStr = hOpt.Retrieve("PDFSigningSignatureLeftX")
62740   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
62750     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
62760       .PDFSigningSignatureLeftX = CDbl(Replace$(tStr, ".", GetDecimalChar))
62770      Else
62780       If UseStandard Then
62790        .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
62800       End If
62810     End If
62820    Else
62830     If UseStandard Then
62840      .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
62850     End If
62860   End If
62870   tStr = hOpt.Retrieve("PDFSigningSignatureLeftY")
62880   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
62890     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
62900       .PDFSigningSignatureLeftY = CDbl(Replace$(tStr, ".", GetDecimalChar))
62910      Else
62920       If UseStandard Then
62930        .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
62940       End If
62950     End If
62960    Else
62970     If UseStandard Then
62980      .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
62990     End If
63000   End If
63010   tStr = hOpt.Retrieve("PDFSigningSignatureLocation")
63020   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63030     .PDFSigningSignatureLocation = ""
63040    Else
63050     If LenB(tStr) > 0 Then
63060      .PDFSigningSignatureLocation = tStr
63070     End If
63080   End If
63090   tStr = hOpt.Retrieve("PDFSigningSignatureReason")
63100   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63110     .PDFSigningSignatureReason = ""
63120    Else
63130     If LenB(tStr) > 0 Then
63140      .PDFSigningSignatureReason = tStr
63150     End If
63160   End If
63170   tStr = hOpt.Retrieve("PDFSigningSignatureRightX")
63180   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63190     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63200       .PDFSigningSignatureRightX = CDbl(Replace$(tStr, ".", GetDecimalChar))
63210      Else
63220       If UseStandard Then
63230        .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
63240       End If
63250     End If
63260    Else
63270     If UseStandard Then
63280      .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
63290     End If
63300   End If
63310   tStr = hOpt.Retrieve("PDFSigningSignatureRightY")
63320   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63330     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63340       .PDFSigningSignatureRightY = CDbl(Replace$(tStr, ".", GetDecimalChar))
63350      Else
63360       If UseStandard Then
63370        .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
63380       End If
63390     End If
63400    Else
63410     If UseStandard Then
63420      .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
63430     End If
63440   End If
63450   tStr = hOpt.Retrieve("PDFSigningSignatureVisible")
63460   If IsNumeric(tStr) Then
63470     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63480       .PDFSigningSignatureVisible = CLng(tStr)
63490      Else
63500       If UseStandard Then
63510        .PDFSigningSignatureVisible = 0
63520       End If
63530     End If
63540    Else
63550     If UseStandard Then
63560      .PDFSigningSignatureVisible = 0
63570     End If
63580   End If
63590   tStr = hOpt.Retrieve("PDFSigningSignPDF")
63600   If IsNumeric(tStr) Then
63610     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63620       .PDFSigningSignPDF = CLng(tStr)
63630      Else
63640       If UseStandard Then
63650        .PDFSigningSignPDF = 0
63660       End If
63670     End If
63680    Else
63690     If UseStandard Then
63700      .PDFSigningSignPDF = 0
63710     End If
63720   End If
63730   tStr = hOpt.Retrieve("PDFUpdateMetadata")
63740   If IsNumeric(tStr) Then
63750     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
63760       .PDFUpdateMetadata = CLng(tStr)
63770      Else
63780       If UseStandard Then
63790        .PDFUpdateMetadata = 1
63800       End If
63810     End If
63820    Else
63830     If UseStandard Then
63840      .PDFUpdateMetadata = 1
63850     End If
63860   End If
63870   tStr = hOpt.Retrieve("PDFUserPass")
63880   If IsNumeric(tStr) Then
63890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63900       .PDFUserPass = CLng(tStr)
63910      Else
63920       If UseStandard Then
63930        .PDFUserPass = 0
63940       End If
63950     End If
63960    Else
63970     If UseStandard Then
63980      .PDFUserPass = 0
63990     End If
64000   End If
64010   tStr = hOpt.Retrieve("PDFUserPasswordString")
64020   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64030     .PDFUserPasswordString = ""
64040    Else
64050     If LenB(tStr) > 0 Then
64060      .PDFUserPasswordString = tStr
64070     End If
64080   End If
64090   tStr = hOpt.Retrieve("PDFUseSecurity")
64100   If IsNumeric(tStr) Then
64110     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64120       .PDFUseSecurity = CLng(tStr)
64130      Else
64140       If UseStandard Then
64150        .PDFUseSecurity = 0
64160       End If
64170     End If
64180    Else
64190     If UseStandard Then
64200      .PDFUseSecurity = 0
64210     End If
64220   End If
64230   tStr = hOpt.Retrieve("PNGColorscount")
64240   If IsNumeric(tStr) Then
64250     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
64260       .PNGColorscount = CLng(tStr)
64270      Else
64280       If UseStandard Then
64290        .PNGColorscount = 0
64300       End If
64310     End If
64320    Else
64330     If UseStandard Then
64340      .PNGColorscount = 0
64350     End If
64360   End If
64370   tStr = hOpt.Retrieve("PrintAfterSaving")
64380   If IsNumeric(tStr) Then
64390     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64400       .PrintAfterSaving = CLng(tStr)
64410      Else
64420       If UseStandard Then
64430        .PrintAfterSaving = 0
64440       End If
64450     End If
64460    Else
64470     If UseStandard Then
64480      .PrintAfterSaving = 0
64490     End If
64500   End If
64510   tStr = hOpt.Retrieve("PrintAfterSavingDuplex")
64520   If IsNumeric(tStr) Then
64530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64540       .PrintAfterSavingDuplex = CLng(tStr)
64550      Else
64560       If UseStandard Then
64570        .PrintAfterSavingDuplex = 0
64580       End If
64590     End If
64600    Else
64610     If UseStandard Then
64620      .PrintAfterSavingDuplex = 0
64630     End If
64640   End If
64650   tStr = hOpt.Retrieve("PrintAfterSavingNoCancel")
64660   If IsNumeric(tStr) Then
64670     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64680       .PrintAfterSavingNoCancel = CLng(tStr)
64690      Else
64700       If UseStandard Then
64710        .PrintAfterSavingNoCancel = 0
64720       End If
64730     End If
64740    Else
64750     If UseStandard Then
64760      .PrintAfterSavingNoCancel = 0
64770     End If
64780   End If
64790   tStr = hOpt.Retrieve("PrintAfterSavingPrinter")
64800   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64810     .PrintAfterSavingPrinter = ""
64820    Else
64830     If LenB(tStr) > 0 Then
64840      .PrintAfterSavingPrinter = tStr
64850     End If
64860   End If
64870   tStr = hOpt.Retrieve("PrintAfterSavingQueryUser")
64880   If IsNumeric(tStr) Then
64890     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
64900       .PrintAfterSavingQueryUser = CLng(tStr)
64910      Else
64920       If UseStandard Then
64930        .PrintAfterSavingQueryUser = 0
64940       End If
64950     End If
64960    Else
64970     If UseStandard Then
64980      .PrintAfterSavingQueryUser = 0
64990     End If
65000   End If
65010   tStr = hOpt.Retrieve("PrintAfterSavingTumble")
65020   If IsNumeric(tStr) Then
65030     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
65040       .PrintAfterSavingTumble = CLng(tStr)
65050      Else
65060       If UseStandard Then
65070        .PrintAfterSavingTumble = 0
65080       End If
65090     End If
65100    Else
65110     If UseStandard Then
65120      .PrintAfterSavingTumble = 0
65130     End If
65140   End If
65150   tStr = hOpt.Retrieve("PrinterStop")
65160   If IsNumeric(tStr) Then
65170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65180       .PrinterStop = CLng(tStr)
65190      Else
65200       If UseStandard Then
65210        .PrinterStop = 0
65220       End If
65230     End If
65240    Else
65250     If UseStandard Then
65260      .PrinterStop = 0
65270     End If
65280   End If
65290   tStr = hOpt.Retrieve("PrinterTemppath")
65300   WriteToSpecialLogfile "hOpt.Retrieve(""PrinterTemppath"")=" & tStr
65310   WriteToSpecialLogfile "Options.PrinterTemppath1=" & .PrinterTemppath
65320   If hkey1 = HKEY_USERS Then
65330     If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then
65340       .PrinterTemppath = tStr
65350      Else
65360       If UseStandard Then
65370         .PrinterTemppath = GetTempPath
65380        Else
65390         .PrinterTemppath = tStr
65400       End If
65410     End If
65420    Else
65430     If LenB(Trim$(tStr)) > 0 Then
65440      If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then
65450        .PrinterTemppath = tStr
65460       Else
65470        MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))
65480        If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then
65490          If UseStandard Then
65500            .PrinterTemppath = GetTempPath
65510           Else
65520            .PrinterTemppath = ""
65530            If NoMsg = False Then
65540             MsgBox "PrinterTemppath: '" & tStr & "' = '" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & "'" & _
             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
65560            End If
65570          End If
65580         Else
65590          .PrinterTemppath = tStr
65600        End If
65610      End If
65620     End If
65630   End If
65640   WriteToSpecialLogfile "Options.PrinterTemppath2=" & .PrinterTemppath
65650   tStr = hOpt.Retrieve("ProcessPriority")
65660   If IsNumeric(tStr) Then
65670     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
65680       .ProcessPriority = CLng(tStr)
65690      Else
65700       If UseStandard Then
65710        .ProcessPriority = 1
65720       End If
65730     End If
65740    Else
65750     If UseStandard Then
65760      .ProcessPriority = 1
65770     End If
65780   End If
65790   tStr = hOpt.Retrieve("ProgramFont")
65800   If LenB(tStr) = 0 And LenB("MS Sans Serif") > 0 And UseStandard Then
65810     .ProgramFont = "MS Sans Serif"
65820    Else
65830     If LenB(tStr) > 0 Then
65840      .ProgramFont = tStr
65850     End If
65860   End If
65870   tStr = hOpt.Retrieve("ProgramFontCharset")
65880   If IsNumeric(tStr) Then
65890     If CLng(tStr) >= 0 Then
65900       .ProgramFontCharset = CLng(tStr)
65910      Else
65920       If UseStandard Then
65930        .ProgramFontCharset = 0
65940       End If
65950     End If
65960    Else
65970     If UseStandard Then
65980      .ProgramFontCharset = 0
65990     End If
66000   End If
66010   tStr = hOpt.Retrieve("ProgramFontSize")
66020   If IsNumeric(tStr) Then
66030     If CLng(tStr) >= 6 And CLng(tStr) <= 72 Then
66040       .ProgramFontSize = CLng(tStr)
66050      Else
66060       If UseStandard Then
66070        .ProgramFontSize = 8
66080       End If
66090     End If
66100    Else
66110     If UseStandard Then
66120      .ProgramFontSize = 8
66130     End If
66140   End If
66150   tStr = hOpt.Retrieve("PSDColorsCount")
66160   If IsNumeric(tStr) Then
66170     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
66180       .PSDColorsCount = CLng(tStr)
66190      Else
66200       If UseStandard Then
66210        .PSDColorsCount = 0
66220       End If
66230     End If
66240    Else
66250     If UseStandard Then
66260      .PSDColorsCount = 0
66270     End If
66280   End If
66290   tStr = hOpt.Retrieve("PSLanguageLevel")
66300   If IsNumeric(tStr) Then
66310     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
66320       .PSLanguageLevel = CLng(tStr)
66330      Else
66340       If UseStandard Then
66350        .PSLanguageLevel = 2
66360       End If
66370     End If
66380    Else
66390     If UseStandard Then
66400      .PSLanguageLevel = 2
66410     End If
66420   End If
66430   tStr = hOpt.Retrieve("RAWColorsCount")
66440   If IsNumeric(tStr) Then
66450     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
66460       .RAWColorsCount = CLng(tStr)
66470      Else
66480       If UseStandard Then
66490        .RAWColorsCount = 0
66500       End If
66510     End If
66520    Else
66530     If UseStandard Then
66540      .RAWColorsCount = 0
66550     End If
66560   End If
66570   tStr = hOpt.Retrieve("RemoveAllKnownFileExtensions")
66580   If IsNumeric(tStr) Then
66590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66600       .RemoveAllKnownFileExtensions = CLng(tStr)
66610      Else
66620       If UseStandard Then
66630        .RemoveAllKnownFileExtensions = 1
66640       End If
66650     End If
66660    Else
66670     If UseStandard Then
66680      .RemoveAllKnownFileExtensions = 1
66690     End If
66700   End If
66710   tStr = hOpt.Retrieve("RemoveSpaces")
66720   If IsNumeric(tStr) Then
66730     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66740       .RemoveSpaces = CLng(tStr)
66750      Else
66760       If UseStandard Then
66770        .RemoveSpaces = 1
66780       End If
66790     End If
66800    Else
66810     If UseStandard Then
66820      .RemoveSpaces = 1
66830     End If
66840   End If
66850   tStr = hOpt.Retrieve("RunProgramAfterSaving")
66860   If IsNumeric(tStr) Then
66870     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66880       .RunProgramAfterSaving = CLng(tStr)
66890      Else
66900       If UseStandard Then
66910        .RunProgramAfterSaving = 0
66920       End If
66930     End If
66940    Else
66950     If UseStandard Then
66960      .RunProgramAfterSaving = 0
66970     End If
66980   End If
66990   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramname")
67000   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67010     .RunProgramAfterSavingProgramname = ""
67020    Else
67030     If LenB(tStr) > 0 Then
67040      .RunProgramAfterSavingProgramname = tStr
67050     End If
67060   End If
67070   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramParameters")
67080   If LenB(tStr) = 0 And LenB("""<OutputFilename>""") > 0 And UseStandard Then
67090     .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
67100    Else
67110     If LenB(tStr) > 0 Then
67120      .RunProgramAfterSavingProgramParameters = tStr
67130     End If
67140   End If
67150   tStr = hOpt.Retrieve("RunProgramAfterSavingWaitUntilReady")
67160   If IsNumeric(tStr) Then
67170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67180       .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
67190      Else
67200       If UseStandard Then
67210        .RunProgramAfterSavingWaitUntilReady = 1
67220       End If
67230     End If
67240    Else
67250     If UseStandard Then
67260      .RunProgramAfterSavingWaitUntilReady = 1
67270     End If
67280   End If
67290   tStr = hOpt.Retrieve("RunProgramAfterSavingWindowstyle")
67300   If IsNumeric(tStr) Then
67310     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
67320       .RunProgramAfterSavingWindowstyle = CLng(tStr)
67330      Else
67340       If UseStandard Then
67350        .RunProgramAfterSavingWindowstyle = 1
67360       End If
67370     End If
67380    Else
67390     If UseStandard Then
67400      .RunProgramAfterSavingWindowstyle = 1
67410     End If
67420   End If
67430   tStr = hOpt.Retrieve("RunProgramBeforeSaving")
67440   If IsNumeric(tStr) Then
67450     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67460       .RunProgramBeforeSaving = CLng(tStr)
67470      Else
67480       If UseStandard Then
67490        .RunProgramBeforeSaving = 0
67500       End If
67510     End If
67520    Else
67530     If UseStandard Then
67540      .RunProgramBeforeSaving = 0
67550     End If
67560   End If
67570   tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramname")
67580   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67590     .RunProgramBeforeSavingProgramname = ""
67600    Else
67610     If LenB(tStr) > 0 Then
67620      .RunProgramBeforeSavingProgramname = tStr
67630     End If
67640   End If
67650   tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramParameters")
67660   If LenB(tStr) = 0 And LenB("""<TempFilename>""") > 0 And UseStandard Then
67670     .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
67680    Else
67690     If LenB(tStr) > 0 Then
67700      .RunProgramBeforeSavingProgramParameters = tStr
67710     End If
67720   End If
67730   tStr = hOpt.Retrieve("RunProgramBeforeSavingWindowstyle")
67740   If IsNumeric(tStr) Then
67750     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
67760       .RunProgramBeforeSavingWindowstyle = CLng(tStr)
67770      Else
67780       If UseStandard Then
67790        .RunProgramBeforeSavingWindowstyle = 1
67800       End If
67810     End If
67820    Else
67830     If UseStandard Then
67840      .RunProgramBeforeSavingWindowstyle = 1
67850     End If
67860   End If
67870   tStr = hOpt.Retrieve("SaveFilename")
67880   If LenB(tStr) = 0 And LenB("<Title>") > 0 And UseStandard Then
67890     .SaveFilename = "<Title>"
67900    Else
67910     If LenB(tStr) > 0 Then
67920      .SaveFilename = tStr
67930     End If
67940   End If
67950   tStr = hOpt.Retrieve("SendEmailAfterAutoSaving")
67960   If IsNumeric(tStr) Then
67970     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67980       .SendEmailAfterAutoSaving = CLng(tStr)
67990      Else
68000       If UseStandard Then
68010        .SendEmailAfterAutoSaving = 0
68020       End If
68030     End If
68040    Else
68050     If UseStandard Then
68060      .SendEmailAfterAutoSaving = 0
68070     End If
68080   End If
68090   tStr = hOpt.Retrieve("SendMailMethod")
68100   If IsNumeric(tStr) Then
68110     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
68120       .SendMailMethod = CLng(tStr)
68130      Else
68140       If UseStandard Then
68150        .SendMailMethod = 0
68160       End If
68170     End If
68180    Else
68190     If UseStandard Then
68200      .SendMailMethod = 0
68210     End If
68220   End If
68230   tStr = hOpt.Retrieve("ShowAnimation")
68240   If IsNumeric(tStr) Then
68250     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68260       .ShowAnimation = CLng(tStr)
68270      Else
68280       If UseStandard Then
68290        .ShowAnimation = 1
68300       End If
68310     End If
68320    Else
68330     If UseStandard Then
68340      .ShowAnimation = 1
68350     End If
68360   End If
68370   tStr = hOpt.Retrieve("StampFontColor")
68380   If LenB(tStr) = 0 And LenB("#FF0000") > 0 And UseStandard Then
68390     .StampFontColor = "#FF0000"
68400    Else
68410     If LenB(tStr) > 0 Then
68420      .StampFontColor = tStr
68430     End If
68440   End If
68450   tStr = hOpt.Retrieve("StampFontname")
68460   If LenB(tStr) = 0 And LenB("Arial") > 0 And UseStandard Then
68470     .StampFontname = "Arial"
68480    Else
68490     If LenB(tStr) > 0 Then
68500      .StampFontname = tStr
68510     End If
68520   End If
68530   tStr = hOpt.Retrieve("StampFontsize")
68540   If IsNumeric(tStr) Then
68550     If CLng(tStr) >= 1 Then
68560       .StampFontsize = CLng(tStr)
68570      Else
68580       If UseStandard Then
68590        .StampFontsize = 48
68600       End If
68610     End If
68620    Else
68630     If UseStandard Then
68640      .StampFontsize = 48
68650     End If
68660   End If
68670   tStr = hOpt.Retrieve("StampOutlineFontthickness")
68680   If IsNumeric(tStr) Then
68690     If CLng(tStr) >= 0 Then
68700       .StampOutlineFontthickness = CLng(tStr)
68710      Else
68720       If UseStandard Then
68730        .StampOutlineFontthickness = 0
68740       End If
68750     End If
68760    Else
68770     If UseStandard Then
68780      .StampOutlineFontthickness = 0
68790     End If
68800   End If
68810   tStr = hOpt.Retrieve("StampString")
68820   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
68830     .StampString = ""
68840    Else
68850     If LenB(tStr) > 0 Then
68860      .StampString = tStr
68870     End If
68880   End If
68890   tStr = hOpt.Retrieve("StampUseOutlineFont")
68900   If IsNumeric(tStr) Then
68910     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68920       .StampUseOutlineFont = CLng(tStr)
68930      Else
68940       If UseStandard Then
68950        .StampUseOutlineFont = 1
68960       End If
68970     End If
68980    Else
68990     If UseStandard Then
69000      .StampUseOutlineFont = 1
69010     End If
69020   End If
69030   tStr = hOpt.Retrieve("StandardAuthor")
69040   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69050     .StandardAuthor = ""
69060    Else
69070     If LenB(tStr) > 0 Then
69080      .StandardAuthor = tStr
69090     End If
69100   End If
69110   tStr = hOpt.Retrieve("StandardCreationdate")
69120   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69130     .StandardCreationdate = ""
69140    Else
69150     If LenB(tStr) > 0 Then
69160      .StandardCreationdate = tStr
69170     End If
69180   End If
69190   tStr = hOpt.Retrieve("StandardDateformat")
69200   If LenB(tStr) = 0 And LenB("YYYYMMDDHHNNSS") > 0 And UseStandard Then
69210     .StandardDateformat = "YYYYMMDDHHNNSS"
69220    Else
69230     If LenB(tStr) > 0 Then
69240      .StandardDateformat = tStr
69250     End If
69260   End If
69270   tStr = hOpt.Retrieve("StandardKeywords")
69280   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69290     .StandardKeywords = ""
69300    Else
69310     If LenB(tStr) > 0 Then
69320      .StandardKeywords = tStr
69330     End If
69340   End If
69350   tStr = hOpt.Retrieve("StandardMailDomain")
69360   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69370     .StandardMailDomain = ""
69380    Else
69390     If LenB(tStr) > 0 Then
69400      .StandardMailDomain = tStr
69410     End If
69420   End If
69430   tStr = hOpt.Retrieve("StandardModifydate")
69440   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69450     .StandardModifydate = ""
69460    Else
69470     If LenB(tStr) > 0 Then
69480      .StandardModifydate = tStr
69490     End If
69500   End If
69510   tStr = hOpt.Retrieve("StandardSaveformat")
69520   If IsNumeric(tStr) Then
69530     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
69540       .StandardSaveformat = CLng(tStr)
69550      Else
69560       If UseStandard Then
69570        .StandardSaveformat = 0
69580       End If
69590     End If
69600    Else
69610     If UseStandard Then
69620      .StandardSaveformat = 0
69630     End If
69640   End If
69650   tStr = hOpt.Retrieve("StandardSubject")
69660   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69670     .StandardSubject = ""
69680    Else
69690     If LenB(tStr) > 0 Then
69700      .StandardSubject = tStr
69710     End If
69720   End If
69730   tStr = hOpt.Retrieve("StandardTitle")
69740   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69750     .StandardTitle = ""
69760    Else
69770     If LenB(tStr) > 0 Then
69780      .StandardTitle = tStr
69790     End If
69800   End If
69810   tStr = hOpt.Retrieve("StartStandardProgram")
69820   If IsNumeric(tStr) Then
69830     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69840       .StartStandardProgram = CLng(tStr)
69850      Else
69860       If UseStandard Then
69870        .StartStandardProgram = 1
69880       End If
69890     End If
69900    Else
69910     If UseStandard Then
69920      .StartStandardProgram = 1
69930     End If
69940   End If
69950   tStr = hOpt.Retrieve("TIFFColorscount")
69960   If IsNumeric(tStr) Then
69970     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
69980       .TIFFColorscount = CLng(tStr)
69990      Else
70000       If UseStandard Then
70010        .TIFFColorscount = 0
70020       End If
70030     End If
70040    Else
70050     If UseStandard Then
70060      .TIFFColorscount = 0
70070     End If
70080   End If
70090   tStr = hOpt.Retrieve("Toolbars")
70100   If IsNumeric(tStr) Then
70110     If CLng(tStr) >= 0 Then
70120       .Toolbars = CLng(tStr)
70130      Else
70140       If UseStandard Then
70150        .Toolbars = 1
70160       End If
70170     End If
70180    Else
70190     If UseStandard Then
70200      .Toolbars = 1
70210     End If
70220   End If
70230   tStr = hOpt.Retrieve("UseAutosave")
70240   If IsNumeric(tStr) Then
70250     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70260       .UseAutosave = CLng(tStr)
70270      Else
70280       If UseStandard Then
70290        .UseAutosave = 0
70300       End If
70310     End If
70320    Else
70330     If UseStandard Then
70340      .UseAutosave = 0
70350     End If
70360   End If
70370   tStr = hOpt.Retrieve("UseAutosaveDirectory")
70380   If IsNumeric(tStr) Then
70390     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70400       .UseAutosaveDirectory = CLng(tStr)
70410      Else
70420       If UseStandard Then
70430        .UseAutosaveDirectory = 1
70440       End If
70450     End If
70460    Else
70470     If UseStandard Then
70480      .UseAutosaveDirectory = 1
70490     End If
70500   End If
70510   tStr = hOpt.Retrieve("UseCreationDateNow")
70520   If IsNumeric(tStr) Then
70530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70540       .UseCreationDateNow = CLng(tStr)
70550      Else
70560       If UseStandard Then
70570        .UseCreationDateNow = 0
70580       End If
70590     End If
70600    Else
70610     If UseStandard Then
70620      .UseCreationDateNow = 0
70630     End If
70640   End If
70650   tStr = hOpt.Retrieve("UseCustomPaperSize")
70660   If LenB(tStr) = 0 And LenB("0") > 0 And UseStandard Then
70670     .UseCustomPaperSize = "0"
70680    Else
70690     If LenB(tStr) > 0 Then
70700      .UseCustomPaperSize = tStr
70710     End If
70720   End If
70730   tStr = hOpt.Retrieve("UseFixPapersize")
70740   If IsNumeric(tStr) Then
70750     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70760       .UseFixPapersize = CLng(tStr)
70770      Else
70780       If UseStandard Then
70790        .UseFixPapersize = 0
70800       End If
70810     End If
70820    Else
70830     If UseStandard Then
70840      .UseFixPapersize = 0
70850     End If
70860   End If
70870   tStr = hOpt.Retrieve("UseStandardAuthor")
70880   If IsNumeric(tStr) Then
70890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70900       .UseStandardAuthor = CLng(tStr)
70910      Else
70920       If UseStandard Then
70930        .UseStandardAuthor = 0
70940       End If
70950     End If
70960    Else
70970     If UseStandard Then
70980      .UseStandardAuthor = 0
70990     End If
71000   End If
71010   tStr = hOpt.Retrieve("XCFColorsCount")
71020   If IsNumeric(tStr) Then
71030     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
71040       .XCFColorsCount = CLng(tStr)
71050      Else
71060       If UseStandard Then
71070        .XCFColorsCount = 0
71080       End If
71090     End If
71100    Else
71110     If UseStandard Then
71120      .XCFColorsCount = 0
71130     End If
71140   End If
71150  End With
71160  Set ini = Nothing
71170  ReadOptionsINI = myOptions
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
50170   Case "BITMAPRESOLUTION": ini.SaveKey CStr(.BitmapResolution), "BitmapResolution"
50180   Case "BMPCOLORSCOUNT": ini.SaveKey CStr(.BMPColorscount), "BMPColorscount"
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
50340   Case "LANGUAGE": ini.SaveKey CStr(.Language), "Language"
50350   Case "LASTSAVEDIRECTORY": ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50360   Case "LOGGING": ini.SaveKey CStr(Abs(.Logging)), "Logging"
50370   Case "LOGLINES": ini.SaveKey CStr(.LogLines), "LogLines"
50380   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50390   Case "NOPROCESSINGATSTARTUP": ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
50400   Case "NOPSCHECK": ini.SaveKey CStr(Abs(.NoPSCheck)), "NoPSCheck"
50410   Case "ONEPAGEPERFILE": ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50420   Case "OPTIONSDESIGN": ini.SaveKey CStr(.OptionsDesign), "OptionsDesign"
50430   Case "OPTIONSENABLED": ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50440   Case "OPTIONSVISIBLE": ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50450   Case "PAPERSIZE": ini.SaveKey CStr(.Papersize), "Papersize"
50460   Case "PCLCOLORSCOUNT": ini.SaveKey CStr(.PCLColorsCount), "PCLColorsCount"
50470   Case "PCXCOLORSCOUNT": ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50480   Case "PDFALLOWASSEMBLY": ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50490   Case "PDFALLOWDEGRADEDPRINTING": ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50500   Case "PDFALLOWFILLIN": ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50510   Case "PDFALLOWSCREENREADERS": ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50520   Case "PDFCOLORSCMYKTORGB": ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50530   Case "PDFCOLORSCOLORMODEL": ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50540   Case "PDFCOLORSPRESERVEHALFTONE": ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50550   Case "PDFCOLORSPRESERVEOVERPRINT": ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50560   Case "PDFCOLORSPRESERVETRANSFER": ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50570   Case "PDFCOMPRESSIONCOLORCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50580   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50590   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGHighFactor"
50600   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGLowFactor"
50610   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMaximumFactor"
50620   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMediumFactor"
50630   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMinimumFactor"
50640   Case "PDFCOMPRESSIONCOLORRESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50650   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50660   Case "PDFCOMPRESSIONCOLORRESOLUTION": ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50670   Case "PDFCOMPRESSIONGREYCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50680   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50690   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGHighFactor"
50700   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGLowFactor"
50710   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMaximumFactor"
50720   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMediumFactor"
50730   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMinimumFactor"
50740   Case "PDFCOMPRESSIONGREYRESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50750   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50760   Case "PDFCOMPRESSIONGREYRESOLUTION": ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50770   Case "PDFCOMPRESSIONMONOCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50780   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50790   Case "PDFCOMPRESSIONMONORESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50800   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50810   Case "PDFCOMPRESSIONMONORESOLUTION": ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50820   Case "PDFCOMPRESSIONTEXTCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50830   Case "PDFDISALLOWCOPY": ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50840   Case "PDFDISALLOWMODIFYANNOTATIONS": ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50850   Case "PDFDISALLOWMODIFYCONTENTS": ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50860   Case "PDFDISALLOWPRINTING": ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50870   Case "PDFENCRYPTOR": ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50880   Case "PDFFONTSEMBEDALL": ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50890   Case "PDFFONTSSUBSETFONTS": ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50900   Case "PDFFONTSSUBSETFONTSPERCENT": ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50910   Case "PDFGENERALASCII85": ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50920   Case "PDFGENERALAUTOROTATE": ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50930   Case "PDFGENERALCOMPATIBILITY": ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50940   Case "PDFGENERALDEFAULT": ini.SaveKey CStr(.PDFGeneralDefault), "PDFGeneralDefault"
50950   Case "PDFGENERALOVERPRINT": ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50960   Case "PDFGENERALRESOLUTION": ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
50970   Case "PDFHIGHENCRYPTION": ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
50980   Case "PDFLOWENCRYPTION": ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
50990   Case "PDFOPTIMIZE": ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
51000   Case "PDFOWNERPASS": ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
51010   Case "PDFOWNERPASSWORDSTRING": ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
51020   Case "PDFSIGNINGMULTISIGNATURE": ini.SaveKey CStr(Abs(.PDFSigningMultiSignature)), "PDFSigningMultiSignature"
51030   Case "PDFSIGNINGPFXFILE": ini.SaveKey CStr(.PDFSigningPFXFile), "PDFSigningPFXFile"
51040   Case "PDFSIGNINGPFXFILEPASSWORD": ini.SaveKey CStr(.PDFSigningPFXFilePassword), "PDFSigningPFXFilePassword"
51050   Case "PDFSIGNINGSIGNATURECONTACT": ini.SaveKey CStr(.PDFSigningSignatureContact), "PDFSigningSignatureContact"
51060   Case "PDFSIGNINGSIGNATURELEFTX": ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), "PDFSigningSignatureLeftX"
51070   Case "PDFSIGNINGSIGNATURELEFTY": ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), "PDFSigningSignatureLeftY"
51080   Case "PDFSIGNINGSIGNATURELOCATION": ini.SaveKey CStr(.PDFSigningSignatureLocation), "PDFSigningSignatureLocation"
51090   Case "PDFSIGNINGSIGNATUREREASON": ini.SaveKey CStr(.PDFSigningSignatureReason), "PDFSigningSignatureReason"
51100   Case "PDFSIGNINGSIGNATURERIGHTX": ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), "PDFSigningSignatureRightX"
51110   Case "PDFSIGNINGSIGNATURERIGHTY": ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), "PDFSigningSignatureRightY"
51120   Case "PDFSIGNINGSIGNATUREVISIBLE": ini.SaveKey CStr(Abs(.PDFSigningSignatureVisible)), "PDFSigningSignatureVisible"
51130   Case "PDFSIGNINGSIGNPDF": ini.SaveKey CStr(Abs(.PDFSigningSignPDF)), "PDFSigningSignPDF"
51140   Case "PDFUPDATEMETADATA": ini.SaveKey CStr(.PDFUpdateMetadata), "PDFUpdateMetadata"
51150   Case "PDFUSERPASS": ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
51160   Case "PDFUSERPASSWORDSTRING": ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
51170   Case "PDFUSESECURITY": ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
51180   Case "PNGCOLORSCOUNT": ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
51190   Case "PRINTAFTERSAVING": ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
51200   Case "PRINTAFTERSAVINGDUPLEX": ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
51210   Case "PRINTAFTERSAVINGNOCANCEL": ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
51220   Case "PRINTAFTERSAVINGPRINTER": ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
51230   Case "PRINTAFTERSAVINGQUERYUSER": ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
51240   Case "PRINTAFTERSAVINGTUMBLE": ini.SaveKey CStr(.PrintAfterSavingTumble), "PrintAfterSavingTumble"
51250   Case "PRINTERSTOP": ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
51260   Case "PRINTERTEMPPATH": ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
51270   Case "PROCESSPRIORITY": ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
51280   Case "PROGRAMFONT": ini.SaveKey CStr(.ProgramFont), "ProgramFont"
51290   Case "PROGRAMFONTCHARSET": ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
51300   Case "PROGRAMFONTSIZE": ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
51310   Case "PSDCOLORSCOUNT": ini.SaveKey CStr(.PSDColorsCount), "PSDColorsCount"
51320   Case "PSLANGUAGELEVEL": ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
51330   Case "RAWCOLORSCOUNT": ini.SaveKey CStr(.RAWColorsCount), "RAWColorsCount"
51340   Case "REMOVEALLKNOWNFILEEXTENSIONS": ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
51350   Case "REMOVESPACES": ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
51360   Case "RUNPROGRAMAFTERSAVING": ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
51370   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
51380   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
51390   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51400   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51410   Case "RUNPROGRAMBEFORESAVING": ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51420   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51430   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51440   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51450   Case "SAVEFILENAME": ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51460   Case "SENDEMAILAFTERAUTOSAVING": ini.SaveKey CStr(Abs(.SendEmailAfterAutoSaving)), "SendEmailAfterAutoSaving"
51470   Case "SENDMAILMETHOD": ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51480   Case "SHOWANIMATION": ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51490   Case "STAMPFONTCOLOR": ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51500   Case "STAMPFONTNAME": ini.SaveKey CStr(.StampFontname), "StampFontname"
51510   Case "STAMPFONTSIZE": ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51520   Case "STAMPOUTLINEFONTTHICKNESS": ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51530   Case "STAMPSTRING": ini.SaveKey CStr(.StampString), "StampString"
51540   Case "STAMPUSEOUTLINEFONT": ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51550   Case "STANDARDAUTHOR": ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51560   Case "STANDARDCREATIONDATE": ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51570   Case "STANDARDDATEFORMAT": ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51580   Case "STANDARDKEYWORDS": ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51590   Case "STANDARDMAILDOMAIN": ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51600   Case "STANDARDMODIFYDATE": ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51610   Case "STANDARDSAVEFORMAT": ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51620   Case "STANDARDSUBJECT": ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51630   Case "STANDARDTITLE": ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51640   Case "STARTSTANDARDPROGRAM": ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51650   Case "TIFFCOLORSCOUNT": ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51660   Case "TOOLBARS": ini.SaveKey CStr(.Toolbars), "Toolbars"
51670   Case "USEAUTOSAVE": ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51680   Case "USEAUTOSAVEDIRECTORY": ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51690   Case "USECREATIONDATENOW": ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51700   Case "USECUSTOMPAPERSIZE": ini.SaveKey CStr(.UseCustomPaperSize), "UseCustomPaperSize"
51710   Case "USEFIXPAPERSIZE": ini.SaveKey CStr(Abs(.UseFixPapersize)), "UseFixPapersize"
51720   Case "USESTANDARDAUTHOR": ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51730   Case "XCFCOLORSCOUNT": ini.SaveKey CStr(.XCFColorsCount), "XCFColorsCount"
51740   End Select
51750  End With
51760  Set ini = Nothing
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
50160   ini.SaveKey CStr(.BitmapResolution), "BitmapResolution"
50170   ini.SaveKey CStr(.BMPColorscount), "BMPColorscount"
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
50330   ini.SaveKey CStr(.Language), "Language"
50340   ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50350   ini.SaveKey CStr(Abs(.Logging)), "Logging"
50360   ini.SaveKey CStr(.LogLines), "LogLines"
50370   ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50380   ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
50390   ini.SaveKey CStr(Abs(.NoPSCheck)), "NoPSCheck"
50400   ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50410   ini.SaveKey CStr(.OptionsDesign), "OptionsDesign"
50420   ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50430   ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50440   ini.SaveKey CStr(.Papersize), "Papersize"
50450   ini.SaveKey CStr(.PCLColorsCount), "PCLColorsCount"
50460   ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50470   ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50480   ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50490   ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50500   ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50510   ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50520   ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50530   ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50540   ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50550   ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50560   ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50570   ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50580   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGHighFactor"
50590   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGLowFactor"
50600   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMaximumFactor"
50610   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMediumFactor"
50620   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMinimumFactor"
50630   ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50640   ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50650   ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50660   ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50670   ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50680   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGHighFactor"
50690   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGLowFactor"
50700   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMaximumFactor"
50710   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMediumFactor"
50720   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMinimumFactor"
50730   ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50740   ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50750   ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50760   ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50770   ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50780   ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50790   ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50800   ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50810   ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50820   ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50830   ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50840   ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50850   ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50860   ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50870   ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50880   ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50890   ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50900   ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50910   ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50920   ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50930   ini.SaveKey CStr(.PDFGeneralDefault), "PDFGeneralDefault"
50940   ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50950   ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
50960   ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
50970   ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
50980   ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
50990   ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
51000   ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
51010   ini.SaveKey CStr(Abs(.PDFSigningMultiSignature)), "PDFSigningMultiSignature"
51020   ini.SaveKey CStr(.PDFSigningPFXFile), "PDFSigningPFXFile"
51030   ini.SaveKey CStr(.PDFSigningPFXFilePassword), "PDFSigningPFXFilePassword"
51040   ini.SaveKey CStr(.PDFSigningSignatureContact), "PDFSigningSignatureContact"
51050   ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), "PDFSigningSignatureLeftX"
51060   ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), "PDFSigningSignatureLeftY"
51070   ini.SaveKey CStr(.PDFSigningSignatureLocation), "PDFSigningSignatureLocation"
51080   ini.SaveKey CStr(.PDFSigningSignatureReason), "PDFSigningSignatureReason"
51090   ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), "PDFSigningSignatureRightX"
51100   ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), "PDFSigningSignatureRightY"
51110   ini.SaveKey CStr(Abs(.PDFSigningSignatureVisible)), "PDFSigningSignatureVisible"
51120   ini.SaveKey CStr(Abs(.PDFSigningSignPDF)), "PDFSigningSignPDF"
51130   ini.SaveKey CStr(.PDFUpdateMetadata), "PDFUpdateMetadata"
51140   ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
51150   ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
51160   ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
51170   ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
51180   ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
51190   ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
51200   ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
51210   ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
51220   ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
51230   ini.SaveKey CStr(.PrintAfterSavingTumble), "PrintAfterSavingTumble"
51240   ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
51250   ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
51260   ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
51270   ini.SaveKey CStr(.ProgramFont), "ProgramFont"
51280   ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
51290   ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
51300   ini.SaveKey CStr(.PSDColorsCount), "PSDColorsCount"
51310   ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
51320   ini.SaveKey CStr(.RAWColorsCount), "RAWColorsCount"
51330   ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
51340   ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
51350   ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
51360   ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
51370   ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
51380   ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51390   ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51400   ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51410   ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51420   ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51430   ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51440   ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51450   ini.SaveKey CStr(Abs(.SendEmailAfterAutoSaving)), "SendEmailAfterAutoSaving"
51460   ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51470   ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51480   ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51490   ini.SaveKey CStr(.StampFontname), "StampFontname"
51500   ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51510   ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51520   ini.SaveKey CStr(.StampString), "StampString"
51530   ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51540   ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51550   ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51560   ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51570   ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51580   ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51590   ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51600   ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51610   ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51620   ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51630   ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51640   ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51650   ini.SaveKey CStr(.Toolbars), "Toolbars"
51660   ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51670   ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51680   ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51690   ini.SaveKey CStr(.UseCustomPaperSize), "UseCustomPaperSize"
51700   ini.SaveKey CStr(Abs(.UseFixPapersize)), "UseFixPapersize"
51710   ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51720   ini.SaveKey CStr(.XCFColorsCount), "XCFColorsCount"
51730  End With
51740  Set ini = Nothing
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
53020   tStr = reg.GetRegistryValue("BitmapResolution")
53030   If IsNumeric(tStr) Then
53040     If CLng(tStr) >= 1 Then
53050       .BitmapResolution = CLng(tStr)
53060      Else
53070       If UseStandard Then
53080        .BitmapResolution = 150
53090       End If
53100     End If
53110    Else
53120     If UseStandard Then
53130      .BitmapResolution = 150
53140     End If
53150   End If
53160   tStr = reg.GetRegistryValue("BMPColorscount")
53170   If IsNumeric(tStr) Then
53180     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
53190       .BMPColorscount = CLng(tStr)
53200      Else
53210       If UseStandard Then
53220        .BMPColorscount = 1
53230       End If
53240     End If
53250    Else
53260     If UseStandard Then
53270      .BMPColorscount = 1
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
53580   tStr = reg.GetRegistryValue("PCLColorsCount")
53590   If IsNumeric(tStr) Then
53600     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
53610       .PCLColorsCount = CLng(tStr)
53620      Else
53630       If UseStandard Then
53640        .PCLColorsCount = 0
53650       End If
53660     End If
53670    Else
53680     If UseStandard Then
53690      .PCLColorsCount = 0
53700     End If
53710   End If
53720   tStr = reg.GetRegistryValue("PCXColorscount")
53730   If IsNumeric(tStr) Then
53740     If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
53750       .PCXColorscount = CLng(tStr)
53760      Else
53770       If UseStandard Then
53780        .PCXColorscount = 0
53790       End If
53800     End If
53810    Else
53820     If UseStandard Then
53830      .PCXColorscount = 0
53840     End If
53850   End If
53860   tStr = reg.GetRegistryValue("PNGColorscount")
53870   If IsNumeric(tStr) Then
53880     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
53890       .PNGColorscount = CLng(tStr)
53900      Else
53910       If UseStandard Then
53920        .PNGColorscount = 0
53930       End If
53940     End If
53950    Else
53960     If UseStandard Then
53970      .PNGColorscount = 0
53980     End If
53990   End If
54000   tStr = reg.GetRegistryValue("PSDColorsCount")
54010   If IsNumeric(tStr) Then
54020     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
54030       .PSDColorsCount = CLng(tStr)
54040      Else
54050       If UseStandard Then
54060        .PSDColorsCount = 0
54070       End If
54080     End If
54090    Else
54100     If UseStandard Then
54110      .PSDColorsCount = 0
54120     End If
54130   End If
54140   tStr = reg.GetRegistryValue("RAWColorsCount")
54150   If IsNumeric(tStr) Then
54160     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
54170       .RAWColorsCount = CLng(tStr)
54180      Else
54190       If UseStandard Then
54200        .RAWColorsCount = 0
54210       End If
54220     End If
54230    Else
54240     If UseStandard Then
54250      .RAWColorsCount = 0
54260     End If
54270   End If
54280   tStr = reg.GetRegistryValue("TIFFColorscount")
54290   If IsNumeric(tStr) Then
54300     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
54310       .TIFFColorscount = CLng(tStr)
54320      Else
54330       If UseStandard Then
54340        .TIFFColorscount = 0
54350       End If
54360     End If
54370    Else
54380     If UseStandard Then
54390      .TIFFColorscount = 0
54400     End If
54410   End If
54420   tStr = reg.GetRegistryValue("XCFColorsCount")
54430   If IsNumeric(tStr) Then
54440     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
54450       .XCFColorsCount = CLng(tStr)
54460      Else
54470       If UseStandard Then
54480        .XCFColorsCount = 0
54490       End If
54500     End If
54510    Else
54520     If UseStandard Then
54530      .XCFColorsCount = 0
54540     End If
54550   End If
54560   reg.Subkey = "Printing\Formats\PDF\Colors"
54570   tStr = reg.GetRegistryValue("PDFColorsCMYKToRGB")
54580   If IsNumeric(tStr) Then
54590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54600       .PDFColorsCMYKToRGB = CLng(tStr)
54610      Else
54620       If UseStandard Then
54630        .PDFColorsCMYKToRGB = 0
54640       End If
54650     End If
54660    Else
54670     If UseStandard Then
54680      .PDFColorsCMYKToRGB = 0
54690     End If
54700   End If
54710   tStr = reg.GetRegistryValue("PDFColorsColorModel")
54720   If IsNumeric(tStr) Then
54730     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
54740       .PDFColorsColorModel = CLng(tStr)
54750      Else
54760       If UseStandard Then
54770        .PDFColorsColorModel = 1
54780       End If
54790     End If
54800    Else
54810     If UseStandard Then
54820      .PDFColorsColorModel = 1
54830     End If
54840   End If
54850   tStr = reg.GetRegistryValue("PDFColorsPreserveHalftone")
54860   If IsNumeric(tStr) Then
54870     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54880       .PDFColorsPreserveHalftone = CLng(tStr)
54890      Else
54900       If UseStandard Then
54910        .PDFColorsPreserveHalftone = 0
54920       End If
54930     End If
54940    Else
54950     If UseStandard Then
54960      .PDFColorsPreserveHalftone = 0
54970     End If
54980   End If
54990   tStr = reg.GetRegistryValue("PDFColorsPreserveOverprint")
55000   If IsNumeric(tStr) Then
55010     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55020       .PDFColorsPreserveOverprint = CLng(tStr)
55030      Else
55040       If UseStandard Then
55050        .PDFColorsPreserveOverprint = 1
55060       End If
55070     End If
55080    Else
55090     If UseStandard Then
55100      .PDFColorsPreserveOverprint = 1
55110     End If
55120   End If
55130   tStr = reg.GetRegistryValue("PDFColorsPreserveTransfer")
55140   If IsNumeric(tStr) Then
55150     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55160       .PDFColorsPreserveTransfer = CLng(tStr)
55170      Else
55180       If UseStandard Then
55190        .PDFColorsPreserveTransfer = 1
55200       End If
55210     End If
55220    Else
55230     If UseStandard Then
55240      .PDFColorsPreserveTransfer = 1
55250     End If
55260   End If
55270   reg.Subkey = "Printing\Formats\PDF\Compression"
55280   tStr = reg.GetRegistryValue("PDFCompressionColorCompression")
55290   If IsNumeric(tStr) Then
55300     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55310       .PDFCompressionColorCompression = CLng(tStr)
55320      Else
55330       If UseStandard Then
55340        .PDFCompressionColorCompression = 1
55350       End If
55360     End If
55370    Else
55380     If UseStandard Then
55390      .PDFCompressionColorCompression = 1
55400     End If
55410   End If
55420   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionChoice")
55430   If IsNumeric(tStr) Then
55440     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
55450       .PDFCompressionColorCompressionChoice = CLng(tStr)
55460      Else
55470       If UseStandard Then
55480        .PDFCompressionColorCompressionChoice = 0
55490       End If
55500     End If
55510    Else
55520     If UseStandard Then
55530      .PDFCompressionColorCompressionChoice = 0
55540     End If
55550   End If
55560   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGHighFactor")
55570   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
55580     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
55590       .PDFCompressionColorCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
55600      Else
55610       If UseStandard Then
55620        .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
55630       End If
55640     End If
55650    Else
55660     If UseStandard Then
55670      .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
55680     End If
55690   End If
55700   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGLowFactor")
55710   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
55720     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
55730       .PDFCompressionColorCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
55740      Else
55750       If UseStandard Then
55760        .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
55770       End If
55780     End If
55790    Else
55800     If UseStandard Then
55810      .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
55820     End If
55830   End If
55840   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMaximumFactor")
55850   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
55860     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
55870       .PDFCompressionColorCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
55880      Else
55890       If UseStandard Then
55900        .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
55910       End If
55920     End If
55930    Else
55940     If UseStandard Then
55950      .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
55960     End If
55970   End If
55980   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMediumFactor")
55990   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56000     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56010       .PDFCompressionColorCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56020      Else
56030       If UseStandard Then
56040        .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56050       End If
56060     End If
56070    Else
56080     If UseStandard Then
56090      .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56100     End If
56110   End If
56120   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMinimumFactor")
56130   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56140     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56150       .PDFCompressionColorCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56160      Else
56170       If UseStandard Then
56180        .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
56190       End If
56200     End If
56210    Else
56220     If UseStandard Then
56230      .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
56240     End If
56250   End If
56260   tStr = reg.GetRegistryValue("PDFCompressionColorResample")
56270   If IsNumeric(tStr) Then
56280     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56290       .PDFCompressionColorResample = CLng(tStr)
56300      Else
56310       If UseStandard Then
56320        .PDFCompressionColorResample = 0
56330       End If
56340     End If
56350    Else
56360     If UseStandard Then
56370      .PDFCompressionColorResample = 0
56380     End If
56390   End If
56400   tStr = reg.GetRegistryValue("PDFCompressionColorResampleChoice")
56410   If IsNumeric(tStr) Then
56420     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
56430       .PDFCompressionColorResampleChoice = CLng(tStr)
56440      Else
56450       If UseStandard Then
56460        .PDFCompressionColorResampleChoice = 0
56470       End If
56480     End If
56490    Else
56500     If UseStandard Then
56510      .PDFCompressionColorResampleChoice = 0
56520     End If
56530   End If
56540   tStr = reg.GetRegistryValue("PDFCompressionColorResolution")
56550   If IsNumeric(tStr) Then
56560     If CLng(tStr) >= 0 Then
56570       .PDFCompressionColorResolution = CLng(tStr)
56580      Else
56590       If UseStandard Then
56600        .PDFCompressionColorResolution = 300
56610       End If
56620     End If
56630    Else
56640     If UseStandard Then
56650      .PDFCompressionColorResolution = 300
56660     End If
56670   End If
56680   tStr = reg.GetRegistryValue("PDFCompressionGreyCompression")
56690   If IsNumeric(tStr) Then
56700     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56710       .PDFCompressionGreyCompression = CLng(tStr)
56720      Else
56730       If UseStandard Then
56740        .PDFCompressionGreyCompression = 1
56750       End If
56760     End If
56770    Else
56780     If UseStandard Then
56790      .PDFCompressionGreyCompression = 1
56800     End If
56810   End If
56820   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionChoice")
56830   If IsNumeric(tStr) Then
56840     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
56850       .PDFCompressionGreyCompressionChoice = CLng(tStr)
56860      Else
56870       If UseStandard Then
56880        .PDFCompressionGreyCompressionChoice = 0
56890       End If
56900     End If
56910    Else
56920     If UseStandard Then
56930      .PDFCompressionGreyCompressionChoice = 0
56940     End If
56950   End If
56960   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGHighFactor")
56970   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56980     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56990       .PDFCompressionGreyCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57000      Else
57010       If UseStandard Then
57020        .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
57030       End If
57040     End If
57050    Else
57060     If UseStandard Then
57070      .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
57080     End If
57090   End If
57100   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGLowFactor")
57110   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57120     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57130       .PDFCompressionGreyCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57140      Else
57150       If UseStandard Then
57160        .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
57170       End If
57180     End If
57190    Else
57200     If UseStandard Then
57210      .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
57220     End If
57230   End If
57240   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMaximumFactor")
57250   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57260     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57270       .PDFCompressionGreyCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57280      Else
57290       If UseStandard Then
57300        .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57310       End If
57320     End If
57330    Else
57340     If UseStandard Then
57350      .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57360     End If
57370   End If
57380   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMediumFactor")
57390   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57400     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57410       .PDFCompressionGreyCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57420      Else
57430       If UseStandard Then
57440        .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57450       End If
57460     End If
57470    Else
57480     If UseStandard Then
57490      .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57500     End If
57510   End If
57520   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMinimumFactor")
57530   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57540     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57550       .PDFCompressionGreyCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57560      Else
57570       If UseStandard Then
57580        .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57590       End If
57600     End If
57610    Else
57620     If UseStandard Then
57630      .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57640     End If
57650   End If
57660   tStr = reg.GetRegistryValue("PDFCompressionGreyResample")
57670   If IsNumeric(tStr) Then
57680     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57690       .PDFCompressionGreyResample = CLng(tStr)
57700      Else
57710       If UseStandard Then
57720        .PDFCompressionGreyResample = 0
57730       End If
57740     End If
57750    Else
57760     If UseStandard Then
57770      .PDFCompressionGreyResample = 0
57780     End If
57790   End If
57800   tStr = reg.GetRegistryValue("PDFCompressionGreyResampleChoice")
57810   If IsNumeric(tStr) Then
57820     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
57830       .PDFCompressionGreyResampleChoice = CLng(tStr)
57840      Else
57850       If UseStandard Then
57860        .PDFCompressionGreyResampleChoice = 0
57870       End If
57880     End If
57890    Else
57900     If UseStandard Then
57910      .PDFCompressionGreyResampleChoice = 0
57920     End If
57930   End If
57940   tStr = reg.GetRegistryValue("PDFCompressionGreyResolution")
57950   If IsNumeric(tStr) Then
57960     If CLng(tStr) >= 0 Then
57970       .PDFCompressionGreyResolution = CLng(tStr)
57980      Else
57990       If UseStandard Then
58000        .PDFCompressionGreyResolution = 300
58010       End If
58020     End If
58030    Else
58040     If UseStandard Then
58050      .PDFCompressionGreyResolution = 300
58060     End If
58070   End If
58080   tStr = reg.GetRegistryValue("PDFCompressionMonoCompression")
58090   If IsNumeric(tStr) Then
58100     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58110       .PDFCompressionMonoCompression = CLng(tStr)
58120      Else
58130       If UseStandard Then
58140        .PDFCompressionMonoCompression = 1
58150       End If
58160     End If
58170    Else
58180     If UseStandard Then
58190      .PDFCompressionMonoCompression = 1
58200     End If
58210   End If
58220   tStr = reg.GetRegistryValue("PDFCompressionMonoCompressionChoice")
58230   If IsNumeric(tStr) Then
58240     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
58250       .PDFCompressionMonoCompressionChoice = CLng(tStr)
58260      Else
58270       If UseStandard Then
58280        .PDFCompressionMonoCompressionChoice = 0
58290       End If
58300     End If
58310    Else
58320     If UseStandard Then
58330      .PDFCompressionMonoCompressionChoice = 0
58340     End If
58350   End If
58360   tStr = reg.GetRegistryValue("PDFCompressionMonoResample")
58370   If IsNumeric(tStr) Then
58380     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58390       .PDFCompressionMonoResample = CLng(tStr)
58400      Else
58410       If UseStandard Then
58420        .PDFCompressionMonoResample = 0
58430       End If
58440     End If
58450    Else
58460     If UseStandard Then
58470      .PDFCompressionMonoResample = 0
58480     End If
58490   End If
58500   tStr = reg.GetRegistryValue("PDFCompressionMonoResampleChoice")
58510   If IsNumeric(tStr) Then
58520     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
58530       .PDFCompressionMonoResampleChoice = CLng(tStr)
58540      Else
58550       If UseStandard Then
58560        .PDFCompressionMonoResampleChoice = 0
58570       End If
58580     End If
58590    Else
58600     If UseStandard Then
58610      .PDFCompressionMonoResampleChoice = 0
58620     End If
58630   End If
58640   tStr = reg.GetRegistryValue("PDFCompressionMonoResolution")
58650   If IsNumeric(tStr) Then
58660     If CLng(tStr) >= 0 Then
58670       .PDFCompressionMonoResolution = CLng(tStr)
58680      Else
58690       If UseStandard Then
58700        .PDFCompressionMonoResolution = 1200
58710       End If
58720     End If
58730    Else
58740     If UseStandard Then
58750      .PDFCompressionMonoResolution = 1200
58760     End If
58770   End If
58780   tStr = reg.GetRegistryValue("PDFCompressionTextCompression")
58790   If IsNumeric(tStr) Then
58800     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58810       .PDFCompressionTextCompression = CLng(tStr)
58820      Else
58830       If UseStandard Then
58840        .PDFCompressionTextCompression = 1
58850       End If
58860     End If
58870    Else
58880     If UseStandard Then
58890      .PDFCompressionTextCompression = 1
58900     End If
58910   End If
58920   reg.Subkey = "Printing\Formats\PDF\Fonts"
58930   tStr = reg.GetRegistryValue("PDFFontsEmbedAll")
58940   If IsNumeric(tStr) Then
58950     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58960       .PDFFontsEmbedAll = CLng(tStr)
58970      Else
58980       If UseStandard Then
58990        .PDFFontsEmbedAll = 1
59000       End If
59010     End If
59020    Else
59030     If UseStandard Then
59040      .PDFFontsEmbedAll = 1
59050     End If
59060   End If
59070   tStr = reg.GetRegistryValue("PDFFontsSubSetFonts")
59080   If IsNumeric(tStr) Then
59090     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59100       .PDFFontsSubSetFonts = CLng(tStr)
59110      Else
59120       If UseStandard Then
59130        .PDFFontsSubSetFonts = 1
59140       End If
59150     End If
59160    Else
59170     If UseStandard Then
59180      .PDFFontsSubSetFonts = 1
59190     End If
59200   End If
59210   tStr = reg.GetRegistryValue("PDFFontsSubSetFontsPercent")
59220   If IsNumeric(tStr) Then
59230     If CLng(tStr) >= 0 Then
59240       .PDFFontsSubSetFontsPercent = CLng(tStr)
59250      Else
59260       If UseStandard Then
59270        .PDFFontsSubSetFontsPercent = 100
59280       End If
59290     End If
59300    Else
59310     If UseStandard Then
59320      .PDFFontsSubSetFontsPercent = 100
59330     End If
59340   End If
59350   reg.Subkey = "Printing\Formats\PDF\General"
59360   tStr = reg.GetRegistryValue("PDFGeneralASCII85")
59370   If IsNumeric(tStr) Then
59380     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59390       .PDFGeneralASCII85 = CLng(tStr)
59400      Else
59410       If UseStandard Then
59420        .PDFGeneralASCII85 = 0
59430       End If
59440     End If
59450    Else
59460     If UseStandard Then
59470      .PDFGeneralASCII85 = 0
59480     End If
59490   End If
59500   tStr = reg.GetRegistryValue("PDFGeneralAutorotate")
59510   If IsNumeric(tStr) Then
59520     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
59530       .PDFGeneralAutorotate = CLng(tStr)
59540      Else
59550       If UseStandard Then
59560        .PDFGeneralAutorotate = 2
59570       End If
59580     End If
59590    Else
59600     If UseStandard Then
59610      .PDFGeneralAutorotate = 2
59620     End If
59630   End If
59640   tStr = reg.GetRegistryValue("PDFGeneralCompatibility")
59650   If IsNumeric(tStr) Then
59660     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
59670       .PDFGeneralCompatibility = CLng(tStr)
59680      Else
59690       If UseStandard Then
59700        .PDFGeneralCompatibility = 2
59710       End If
59720     End If
59730    Else
59740     If UseStandard Then
59750      .PDFGeneralCompatibility = 2
59760     End If
59770   End If
59780   tStr = reg.GetRegistryValue("PDFGeneralDefault")
59790   If IsNumeric(tStr) Then
59800     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
59810       .PDFGeneralDefault = CLng(tStr)
59820      Else
59830       If UseStandard Then
59840        .PDFGeneralDefault = 0
59850       End If
59860     End If
59870    Else
59880     If UseStandard Then
59890      .PDFGeneralDefault = 0
59900     End If
59910   End If
59920   tStr = reg.GetRegistryValue("PDFGeneralOverprint")
59930   If IsNumeric(tStr) Then
59940     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
59950       .PDFGeneralOverprint = CLng(tStr)
59960      Else
59970       If UseStandard Then
59980        .PDFGeneralOverprint = 0
59990       End If
60000     End If
60010    Else
60020     If UseStandard Then
60030      .PDFGeneralOverprint = 0
60040     End If
60050   End If
60060   tStr = reg.GetRegistryValue("PDFGeneralResolution")
60070   If IsNumeric(tStr) Then
60080     If CLng(tStr) >= 0 Then
60090       .PDFGeneralResolution = CLng(tStr)
60100      Else
60110       If UseStandard Then
60120        .PDFGeneralResolution = 600
60130       End If
60140     End If
60150    Else
60160     If UseStandard Then
60170      .PDFGeneralResolution = 600
60180     End If
60190   End If
60200   tStr = reg.GetRegistryValue("PDFOptimize")
60210   If IsNumeric(tStr) Then
60220     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60230       .PDFOptimize = CLng(tStr)
60240      Else
60250       If UseStandard Then
60260        .PDFOptimize = 0
60270       End If
60280     End If
60290    Else
60300     If UseStandard Then
60310      .PDFOptimize = 0
60320     End If
60330   End If
60340   tStr = reg.GetRegistryValue("PDFUpdateMetadata")
60350   If IsNumeric(tStr) Then
60360     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
60370       .PDFUpdateMetadata = CLng(tStr)
60380      Else
60390       If UseStandard Then
60400        .PDFUpdateMetadata = 1
60410       End If
60420     End If
60430    Else
60440     If UseStandard Then
60450      .PDFUpdateMetadata = 1
60460     End If
60470   End If
60480   reg.Subkey = "Printing\Formats\PDF\Security"
60490   tStr = reg.GetRegistryValue("PDFAllowAssembly")
60500   If IsNumeric(tStr) Then
60510     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60520       .PDFAllowAssembly = CLng(tStr)
60530      Else
60540       If UseStandard Then
60550        .PDFAllowAssembly = 0
60560       End If
60570     End If
60580    Else
60590     If UseStandard Then
60600      .PDFAllowAssembly = 0
60610     End If
60620   End If
60630   tStr = reg.GetRegistryValue("PDFAllowDegradedPrinting")
60640   If IsNumeric(tStr) Then
60650     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60660       .PDFAllowDegradedPrinting = CLng(tStr)
60670      Else
60680       If UseStandard Then
60690        .PDFAllowDegradedPrinting = 0
60700       End If
60710     End If
60720    Else
60730     If UseStandard Then
60740      .PDFAllowDegradedPrinting = 0
60750     End If
60760   End If
60770   tStr = reg.GetRegistryValue("PDFAllowFillIn")
60780   If IsNumeric(tStr) Then
60790     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60800       .PDFAllowFillIn = CLng(tStr)
60810      Else
60820       If UseStandard Then
60830        .PDFAllowFillIn = 0
60840       End If
60850     End If
60860    Else
60870     If UseStandard Then
60880      .PDFAllowFillIn = 0
60890     End If
60900   End If
60910   tStr = reg.GetRegistryValue("PDFAllowScreenReaders")
60920   If IsNumeric(tStr) Then
60930     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60940       .PDFAllowScreenReaders = CLng(tStr)
60950      Else
60960       If UseStandard Then
60970        .PDFAllowScreenReaders = 0
60980       End If
60990     End If
61000    Else
61010     If UseStandard Then
61020      .PDFAllowScreenReaders = 0
61030     End If
61040   End If
61050   tStr = reg.GetRegistryValue("PDFDisallowCopy")
61060   If IsNumeric(tStr) Then
61070     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61080       .PDFDisallowCopy = CLng(tStr)
61090      Else
61100       If UseStandard Then
61110        .PDFDisallowCopy = 1
61120       End If
61130     End If
61140    Else
61150     If UseStandard Then
61160      .PDFDisallowCopy = 1
61170     End If
61180   End If
61190   tStr = reg.GetRegistryValue("PDFDisallowModifyAnnotations")
61200   If IsNumeric(tStr) Then
61210     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61220       .PDFDisallowModifyAnnotations = CLng(tStr)
61230      Else
61240       If UseStandard Then
61250        .PDFDisallowModifyAnnotations = 0
61260       End If
61270     End If
61280    Else
61290     If UseStandard Then
61300      .PDFDisallowModifyAnnotations = 0
61310     End If
61320   End If
61330   tStr = reg.GetRegistryValue("PDFDisallowModifyContents")
61340   If IsNumeric(tStr) Then
61350     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61360       .PDFDisallowModifyContents = CLng(tStr)
61370      Else
61380       If UseStandard Then
61390        .PDFDisallowModifyContents = 0
61400       End If
61410     End If
61420    Else
61430     If UseStandard Then
61440      .PDFDisallowModifyContents = 0
61450     End If
61460   End If
61470   tStr = reg.GetRegistryValue("PDFDisallowPrinting")
61480   If IsNumeric(tStr) Then
61490     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61500       .PDFDisallowPrinting = CLng(tStr)
61510      Else
61520       If UseStandard Then
61530        .PDFDisallowPrinting = 0
61540       End If
61550     End If
61560    Else
61570     If UseStandard Then
61580      .PDFDisallowPrinting = 0
61590     End If
61600   End If
61610   tStr = reg.GetRegistryValue("PDFEncryptor")
61620   If IsNumeric(tStr) Then
61630     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
61640       .PDFEncryptor = CLng(tStr)
61650      Else
61660       If UseStandard Then
61670        .PDFEncryptor = 0
61680       End If
61690     End If
61700    Else
61710     If UseStandard Then
61720      .PDFEncryptor = 0
61730     End If
61740   End If
61750   tStr = reg.GetRegistryValue("PDFHighEncryption")
61760   If IsNumeric(tStr) Then
61770     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61780       .PDFHighEncryption = CLng(tStr)
61790      Else
61800       If UseStandard Then
61810        .PDFHighEncryption = 0
61820       End If
61830     End If
61840    Else
61850     If UseStandard Then
61860      .PDFHighEncryption = 0
61870     End If
61880   End If
61890   tStr = reg.GetRegistryValue("PDFLowEncryption")
61900   If IsNumeric(tStr) Then
61910     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61920       .PDFLowEncryption = CLng(tStr)
61930      Else
61940       If UseStandard Then
61950        .PDFLowEncryption = 1
61960       End If
61970     End If
61980    Else
61990     If UseStandard Then
62000      .PDFLowEncryption = 1
62010     End If
62020   End If
62030   tStr = reg.GetRegistryValue("PDFOwnerPass")
62040   If IsNumeric(tStr) Then
62050     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62060       .PDFOwnerPass = CLng(tStr)
62070      Else
62080       If UseStandard Then
62090        .PDFOwnerPass = 0
62100       End If
62110     End If
62120    Else
62130     If UseStandard Then
62140      .PDFOwnerPass = 0
62150     End If
62160   End If
62170   tStr = reg.GetRegistryValue("PDFOwnerPasswordString")
62180   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62190     .PDFOwnerPasswordString = ""
62200    Else
62210     If LenB(tStr) > 0 Then
62220      .PDFOwnerPasswordString = tStr
62230     End If
62240   End If
62250   tStr = reg.GetRegistryValue("PDFUserPass")
62260   If IsNumeric(tStr) Then
62270     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62280       .PDFUserPass = CLng(tStr)
62290      Else
62300       If UseStandard Then
62310        .PDFUserPass = 0
62320       End If
62330     End If
62340    Else
62350     If UseStandard Then
62360      .PDFUserPass = 0
62370     End If
62380   End If
62390   tStr = reg.GetRegistryValue("PDFUserPasswordString")
62400   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62410     .PDFUserPasswordString = ""
62420    Else
62430     If LenB(tStr) > 0 Then
62440      .PDFUserPasswordString = tStr
62450     End If
62460   End If
62470   tStr = reg.GetRegistryValue("PDFUseSecurity")
62480   If IsNumeric(tStr) Then
62490     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62500       .PDFUseSecurity = CLng(tStr)
62510      Else
62520       If UseStandard Then
62530        .PDFUseSecurity = 0
62540       End If
62550     End If
62560    Else
62570     If UseStandard Then
62580      .PDFUseSecurity = 0
62590     End If
62600   End If
62610   reg.Subkey = "Printing\Formats\PDF\Signing"
62620   tStr = reg.GetRegistryValue("PDFSigningMultiSignature")
62630   If IsNumeric(tStr) Then
62640     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62650       .PDFSigningMultiSignature = CLng(tStr)
62660      Else
62670       If UseStandard Then
62680        .PDFSigningMultiSignature = 0
62690       End If
62700     End If
62710    Else
62720     If UseStandard Then
62730      .PDFSigningMultiSignature = 0
62740     End If
62750   End If
62760   tStr = reg.GetRegistryValue("PDFSigningPFXFile")
62770   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62780     .PDFSigningPFXFile = ""
62790    Else
62800     If LenB(tStr) > 0 Then
62810      .PDFSigningPFXFile = tStr
62820     End If
62830   End If
62840   tStr = reg.GetRegistryValue("PDFSigningPFXFilePassword")
62850   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62860     .PDFSigningPFXFilePassword = ""
62870    Else
62880     If LenB(tStr) > 0 Then
62890      .PDFSigningPFXFilePassword = tStr
62900     End If
62910   End If
62920   tStr = reg.GetRegistryValue("PDFSigningSignatureContact")
62930   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62940     .PDFSigningSignatureContact = ""
62950    Else
62960     If LenB(tStr) > 0 Then
62970      .PDFSigningSignatureContact = tStr
62980     End If
62990   End If
63000   tStr = reg.GetRegistryValue("PDFSigningSignatureLeftX")
63010   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63020     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63030       .PDFSigningSignatureLeftX = CDbl(Replace$(tStr, ".", GetDecimalChar))
63040      Else
63050       If UseStandard Then
63060        .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
63070       End If
63080     End If
63090    Else
63100     If UseStandard Then
63110      .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
63120     End If
63130   End If
63140   tStr = reg.GetRegistryValue("PDFSigningSignatureLeftY")
63150   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63160     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63170       .PDFSigningSignatureLeftY = CDbl(Replace$(tStr, ".", GetDecimalChar))
63180      Else
63190       If UseStandard Then
63200        .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
63210       End If
63220     End If
63230    Else
63240     If UseStandard Then
63250      .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
63260     End If
63270   End If
63280   tStr = reg.GetRegistryValue("PDFSigningSignatureLocation")
63290   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63300     .PDFSigningSignatureLocation = ""
63310    Else
63320     If LenB(tStr) > 0 Then
63330      .PDFSigningSignatureLocation = tStr
63340     End If
63350   End If
63360   tStr = reg.GetRegistryValue("PDFSigningSignatureReason")
63370   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63380     .PDFSigningSignatureReason = ""
63390    Else
63400     If LenB(tStr) > 0 Then
63410      .PDFSigningSignatureReason = tStr
63420     End If
63430   End If
63440   tStr = reg.GetRegistryValue("PDFSigningSignatureRightX")
63450   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63460     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63470       .PDFSigningSignatureRightX = CDbl(Replace$(tStr, ".", GetDecimalChar))
63480      Else
63490       If UseStandard Then
63500        .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
63510       End If
63520     End If
63530    Else
63540     If UseStandard Then
63550      .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
63560     End If
63570   End If
63580   tStr = reg.GetRegistryValue("PDFSigningSignatureRightY")
63590   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63600     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63610       .PDFSigningSignatureRightY = CDbl(Replace$(tStr, ".", GetDecimalChar))
63620      Else
63630       If UseStandard Then
63640        .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
63650       End If
63660     End If
63670    Else
63680     If UseStandard Then
63690      .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
63700     End If
63710   End If
63720   tStr = reg.GetRegistryValue("PDFSigningSignatureVisible")
63730   If IsNumeric(tStr) Then
63740     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63750       .PDFSigningSignatureVisible = CLng(tStr)
63760      Else
63770       If UseStandard Then
63780        .PDFSigningSignatureVisible = 0
63790       End If
63800     End If
63810    Else
63820     If UseStandard Then
63830      .PDFSigningSignatureVisible = 0
63840     End If
63850   End If
63860   tStr = reg.GetRegistryValue("PDFSigningSignPDF")
63870   If IsNumeric(tStr) Then
63880     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63890       .PDFSigningSignPDF = CLng(tStr)
63900      Else
63910       If UseStandard Then
63920        .PDFSigningSignPDF = 0
63930       End If
63940     End If
63950    Else
63960     If UseStandard Then
63970      .PDFSigningSignPDF = 0
63980     End If
63990   End If
64000   reg.Subkey = "Printing\Formats\PS\LanguageLevel"
64010   tStr = reg.GetRegistryValue("EPSLanguageLevel")
64020   If IsNumeric(tStr) Then
64030     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
64040       .EPSLanguageLevel = CLng(tStr)
64050      Else
64060       If UseStandard Then
64070        .EPSLanguageLevel = 2
64080       End If
64090     End If
64100    Else
64110     If UseStandard Then
64120      .EPSLanguageLevel = 2
64130     End If
64140   End If
64150   tStr = reg.GetRegistryValue("PSLanguageLevel")
64160   If IsNumeric(tStr) Then
64170     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
64180       .PSLanguageLevel = CLng(tStr)
64190      Else
64200       If UseStandard Then
64210        .PSLanguageLevel = 2
64220       End If
64230     End If
64240    Else
64250     If UseStandard Then
64260      .PSLanguageLevel = 2
64270     End If
64280   End If
64290   reg.Subkey = "Program"
64300   tStr = reg.GetRegistryValue("AdditionalGhostscriptParameters")
64310   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64320     .AdditionalGhostscriptParameters = ""
64330    Else
64340     If LenB(tStr) > 0 Then
64350      .AdditionalGhostscriptParameters = tStr
64360     End If
64370   End If
64380   tStr = reg.GetRegistryValue("AdditionalGhostscriptSearchpath")
64390   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64400     .AdditionalGhostscriptSearchpath = ""
64410    Else
64420     If LenB(tStr) > 0 Then
64430      .AdditionalGhostscriptSearchpath = tStr
64440     End If
64450   End If
64460   tStr = reg.GetRegistryValue("AddWindowsFontpath")
64470   If IsNumeric(tStr) Then
64480     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64490       .AddWindowsFontpath = CLng(tStr)
64500      Else
64510       If UseStandard Then
64520        .AddWindowsFontpath = 1
64530       End If
64540     End If
64550    Else
64560     If UseStandard Then
64570      .AddWindowsFontpath = 1
64580     End If
64590   End If
64600   tStr = reg.GetRegistryValue("AutosaveDirectory")
64610   If LenB(Trim$(tStr)) > 0 Then
64620     .AutosaveDirectory = CompletePath(tStr)
64630    Else
64640     If UseStandard Then
64650      If InstalledAsServer Then
64660        .AutosaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
64670       Else
64680        .AutosaveDirectory = "<MyFiles>"
64690      End If
64700     End If
64710   End If
64720   tStr = reg.GetRegistryValue("AutosaveFilename")
64730   If LenB(tStr) = 0 And LenB("<DateTime>") > 0 And UseStandard Then
64740     .AutosaveFilename = "<DateTime>"
64750    Else
64760     If LenB(tStr) > 0 Then
64770      .AutosaveFilename = tStr
64780     End If
64790   End If
64800   tStr = reg.GetRegistryValue("AutosaveFormat")
64810   If IsNumeric(tStr) Then
64820     If CLng(tStr) >= 0 And CLng(tStr) <= 13 Then
64830       .AutosaveFormat = CLng(tStr)
64840      Else
64850       If UseStandard Then
64860        .AutosaveFormat = 0
64870       End If
64880     End If
64890    Else
64900     If UseStandard Then
64910      .AutosaveFormat = 0
64920     End If
64930   End If
64940   tStr = reg.GetRegistryValue("AutosaveStartStandardProgram")
64950   If IsNumeric(tStr) Then
64960     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64970       .AutosaveStartStandardProgram = CLng(tStr)
64980      Else
64990       If UseStandard Then
65000        .AutosaveStartStandardProgram = 0
65010       End If
65020     End If
65030    Else
65040     If UseStandard Then
65050      .AutosaveStartStandardProgram = 0
65060     End If
65070   End If
65080   tStr = reg.GetRegistryValue("ClientComputerResolveIPAddress")
65090   If IsNumeric(tStr) Then
65100     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65110       .ClientComputerResolveIPAddress = CLng(tStr)
65120      Else
65130       If UseStandard Then
65140        .ClientComputerResolveIPAddress = 0
65150       End If
65160     End If
65170    Else
65180     If UseStandard Then
65190      .ClientComputerResolveIPAddress = 0
65200     End If
65210   End If
65220   tStr = reg.GetRegistryValue("DisableEmail")
65230   If IsNumeric(tStr) Then
65240     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65250       .DisableEmail = CLng(tStr)
65260      Else
65270       If UseStandard Then
65280        .DisableEmail = 0
65290       End If
65300     End If
65310    Else
65320     If UseStandard Then
65330      .DisableEmail = 0
65340     End If
65350   End If
65360   tStr = reg.GetRegistryValue("DontUseDocumentSettings")
65370   If IsNumeric(tStr) Then
65380     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65390       .DontUseDocumentSettings = CLng(tStr)
65400      Else
65410       If UseStandard Then
65420        .DontUseDocumentSettings = 0
65430       End If
65440     End If
65450    Else
65460     If UseStandard Then
65470      .DontUseDocumentSettings = 0
65480     End If
65490   End If
65500   tStr = reg.GetRegistryValue("FilenameSubstitutions")
65510   If LenB(tStr) = 0 And LenB("Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt") > 0 And UseStandard Then
65520     .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
65530    Else
65540     If LenB(tStr) > 0 Then
65550      .FilenameSubstitutions = tStr
65560     End If
65570   End If
65580   tStr = reg.GetRegistryValue("FilenameSubstitutionsOnlyInTitle")
65590   If IsNumeric(tStr) Then
65600     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65610       .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
65620      Else
65630       If UseStandard Then
65640        .FilenameSubstitutionsOnlyInTitle = 1
65650       End If
65660     End If
65670    Else
65680     If UseStandard Then
65690      .FilenameSubstitutionsOnlyInTitle = 1
65700     End If
65710   End If
65720   tStr = reg.GetRegistryValue("Language")
65730   If LenB(tStr) = 0 And LenB("english") > 0 And UseStandard Then
65740     .Language = "english"
65750    Else
65760     If LenB(tStr) > 0 Then
65770      .Language = tStr
65780     End If
65790   End If
65800   tStr = reg.GetRegistryValue("LastSaveDirectory")
65810   If LenB(Trim$(tStr)) > 0 Then
65820     .LastSaveDirectory = CompletePath(tStr)
65830    Else
65840     If UseStandard Then
65850      If InstalledAsServer Then
65860        .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
65870       Else
65880        .LastSaveDirectory = "<MyFiles>"
65890      End If
65900     End If
65910   End If
65920   tStr = reg.GetRegistryValue("Logging")
65930   If IsNumeric(tStr) Then
65940     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65950       .Logging = CLng(tStr)
65960      Else
65970       If UseStandard Then
65980        .Logging = 0
65990       End If
66000     End If
66010    Else
66020     If UseStandard Then
66030      .Logging = 0
66040     End If
66050   End If
66060   tStr = reg.GetRegistryValue("LogLines")
66070   If IsNumeric(tStr) Then
66080     If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
66090       .LogLines = CLng(tStr)
66100      Else
66110       If UseStandard Then
66120        .LogLines = 100
66130       End If
66140     End If
66150    Else
66160     If UseStandard Then
66170      .LogLines = 100
66180     End If
66190   End If
66200   tStr = reg.GetRegistryValue("NoConfirmMessageSwitchingDefaultprinter")
66210   If IsNumeric(tStr) Then
66220     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66230       .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
66240      Else
66250       If UseStandard Then
66260        .NoConfirmMessageSwitchingDefaultprinter = 0
66270       End If
66280     End If
66290    Else
66300     If UseStandard Then
66310      .NoConfirmMessageSwitchingDefaultprinter = 0
66320     End If
66330   End If
66340   tStr = reg.GetRegistryValue("NoProcessingAtStartup")
66350   If IsNumeric(tStr) Then
66360     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66370       .NoProcessingAtStartup = CLng(tStr)
66380      Else
66390       If UseStandard Then
66400        .NoProcessingAtStartup = 0
66410       End If
66420     End If
66430    Else
66440     If UseStandard Then
66450      .NoProcessingAtStartup = 0
66460     End If
66470   End If
66480   tStr = reg.GetRegistryValue("NoPSCheck")
66490   If IsNumeric(tStr) Then
66500     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66510       .NoPSCheck = CLng(tStr)
66520      Else
66530       If UseStandard Then
66540        .NoPSCheck = 0
66550       End If
66560     End If
66570    Else
66580     If UseStandard Then
66590      .NoPSCheck = 0
66600     End If
66610   End If
66620   tStr = reg.GetRegistryValue("OptionsDesign")
66630   If IsNumeric(tStr) Then
66640     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
66650       .OptionsDesign = CLng(tStr)
66660      Else
66670       If UseStandard Then
66680        .OptionsDesign = 0
66690       End If
66700     End If
66710    Else
66720     If UseStandard Then
66730      .OptionsDesign = 0
66740     End If
66750   End If
66760   tStr = reg.GetRegistryValue("OptionsEnabled")
66770   If IsNumeric(tStr) Then
66780     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66790       .OptionsEnabled = CLng(tStr)
66800      Else
66810       If UseStandard Then
66820        .OptionsEnabled = 1
66830       End If
66840     End If
66850    Else
66860     If UseStandard Then
66870      .OptionsEnabled = 1
66880     End If
66890   End If
66900   tStr = reg.GetRegistryValue("OptionsVisible")
66910   If IsNumeric(tStr) Then
66920     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66930       .OptionsVisible = CLng(tStr)
66940      Else
66950       If UseStandard Then
66960        .OptionsVisible = 1
66970       End If
66980     End If
66990    Else
67000     If UseStandard Then
67010      .OptionsVisible = 1
67020     End If
67030   End If
67040   tStr = reg.GetRegistryValue("PrintAfterSaving")
67050   If IsNumeric(tStr) Then
67060     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67070       .PrintAfterSaving = CLng(tStr)
67080      Else
67090       If UseStandard Then
67100        .PrintAfterSaving = 0
67110       End If
67120     End If
67130    Else
67140     If UseStandard Then
67150      .PrintAfterSaving = 0
67160     End If
67170   End If
67180   tStr = reg.GetRegistryValue("PrintAfterSavingDuplex")
67190   If IsNumeric(tStr) Then
67200     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67210       .PrintAfterSavingDuplex = CLng(tStr)
67220      Else
67230       If UseStandard Then
67240        .PrintAfterSavingDuplex = 0
67250       End If
67260     End If
67270    Else
67280     If UseStandard Then
67290      .PrintAfterSavingDuplex = 0
67300     End If
67310   End If
67320   tStr = reg.GetRegistryValue("PrintAfterSavingNoCancel")
67330   If IsNumeric(tStr) Then
67340     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67350       .PrintAfterSavingNoCancel = CLng(tStr)
67360      Else
67370       If UseStandard Then
67380        .PrintAfterSavingNoCancel = 0
67390       End If
67400     End If
67410    Else
67420     If UseStandard Then
67430      .PrintAfterSavingNoCancel = 0
67440     End If
67450   End If
67460   tStr = reg.GetRegistryValue("PrintAfterSavingPrinter")
67470   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67480     .PrintAfterSavingPrinter = ""
67490    Else
67500     If LenB(tStr) > 0 Then
67510      .PrintAfterSavingPrinter = tStr
67520     End If
67530   End If
67540   tStr = reg.GetRegistryValue("PrintAfterSavingQueryUser")
67550   If IsNumeric(tStr) Then
67560     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
67570       .PrintAfterSavingQueryUser = CLng(tStr)
67580      Else
67590       If UseStandard Then
67600        .PrintAfterSavingQueryUser = 0
67610       End If
67620     End If
67630    Else
67640     If UseStandard Then
67650      .PrintAfterSavingQueryUser = 0
67660     End If
67670   End If
67680   tStr = reg.GetRegistryValue("PrintAfterSavingTumble")
67690   If IsNumeric(tStr) Then
67700     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
67710       .PrintAfterSavingTumble = CLng(tStr)
67720      Else
67730       If UseStandard Then
67740        .PrintAfterSavingTumble = 0
67750       End If
67760     End If
67770    Else
67780     If UseStandard Then
67790      .PrintAfterSavingTumble = 0
67800     End If
67810   End If
67820   tStr = reg.GetRegistryValue("PrinterStop")
67830   If IsNumeric(tStr) Then
67840     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67850       .PrinterStop = CLng(tStr)
67860      Else
67870       If UseStandard Then
67880        .PrinterStop = 0
67890       End If
67900     End If
67910    Else
67920     If UseStandard Then
67930      .PrinterStop = 0
67940     End If
67950   End If
67960   tStr = reg.GetRegistryValue("PrinterTemppath")
67970   WriteToSpecialLogfile "reg.GetRegistryValue(""PrinterTemppath"")=" & tStr
67980   WriteToSpecialLogfile "Options.PrinterTemppath1=" & .PrinterTemppath
67990   If hkey1 = HKEY_USERS Then
68000     If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then
68010       .PrinterTemppath = tStr
68020      Else
68030       If UseStandard Then
68040         .PrinterTemppath = GetTempPath
68050        Else
68060         .PrinterTemppath = tStr
68070       End If
68080     End If
68090    Else
68100     If LenB(Trim$(tStr)) > 0 Then
68110      If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then
68120        .PrinterTemppath = tStr
68130       Else
68140        MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))
68150        If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then
68160          If UseStandard Then
68170            .PrinterTemppath = GetTempPath
68180           Else
68190            .PrinterTemppath = ""
68200            If NoMsg = False Then
68210             MsgBox "PrinterTemppath: '" & tStr & "' = '" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & "'" & _
             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
68230            End If
68240          End If
68250         Else
68260          .PrinterTemppath = tStr
68270        End If
68280      End If
68290     End If
68300   End If
68310   WriteToSpecialLogfile "Options.PrinterTemppath2=" & .PrinterTemppath
68320   tStr = reg.GetRegistryValue("ProcessPriority")
68330   If IsNumeric(tStr) Then
68340     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
68350       .ProcessPriority = CLng(tStr)
68360      Else
68370       If UseStandard Then
68380        .ProcessPriority = 1
68390       End If
68400     End If
68410    Else
68420     If UseStandard Then
68430      .ProcessPriority = 1
68440     End If
68450   End If
68460   tStr = reg.GetRegistryValue("ProgramFont")
68470   If LenB(tStr) = 0 And LenB("MS Sans Serif") > 0 And UseStandard Then
68480     .ProgramFont = "MS Sans Serif"
68490    Else
68500     If LenB(tStr) > 0 Then
68510      .ProgramFont = tStr
68520     End If
68530   End If
68540   tStr = reg.GetRegistryValue("ProgramFontCharset")
68550   If IsNumeric(tStr) Then
68560     If CLng(tStr) >= 0 Then
68570       .ProgramFontCharset = CLng(tStr)
68580      Else
68590       If UseStandard Then
68600        .ProgramFontCharset = 0
68610       End If
68620     End If
68630    Else
68640     If UseStandard Then
68650      .ProgramFontCharset = 0
68660     End If
68670   End If
68680   tStr = reg.GetRegistryValue("ProgramFontSize")
68690   If IsNumeric(tStr) Then
68700     If CLng(tStr) >= 6 And CLng(tStr) <= 72 Then
68710       .ProgramFontSize = CLng(tStr)
68720      Else
68730       If UseStandard Then
68740        .ProgramFontSize = 8
68750       End If
68760     End If
68770    Else
68780     If UseStandard Then
68790      .ProgramFontSize = 8
68800     End If
68810   End If
68820   tStr = reg.GetRegistryValue("RemoveAllKnownFileExtensions")
68830   If IsNumeric(tStr) Then
68840     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68850       .RemoveAllKnownFileExtensions = CLng(tStr)
68860      Else
68870       If UseStandard Then
68880        .RemoveAllKnownFileExtensions = 1
68890       End If
68900     End If
68910    Else
68920     If UseStandard Then
68930      .RemoveAllKnownFileExtensions = 1
68940     End If
68950   End If
68960   tStr = reg.GetRegistryValue("RemoveSpaces")
68970   If IsNumeric(tStr) Then
68980     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68990       .RemoveSpaces = CLng(tStr)
69000      Else
69010       If UseStandard Then
69020        .RemoveSpaces = 1
69030       End If
69040     End If
69050    Else
69060     If UseStandard Then
69070      .RemoveSpaces = 1
69080     End If
69090   End If
69100   tStr = reg.GetRegistryValue("RunProgramAfterSaving")
69110   If IsNumeric(tStr) Then
69120     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69130       .RunProgramAfterSaving = CLng(tStr)
69140      Else
69150       If UseStandard Then
69160        .RunProgramAfterSaving = 0
69170       End If
69180     End If
69190    Else
69200     If UseStandard Then
69210      .RunProgramAfterSaving = 0
69220     End If
69230   End If
69240   tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramname")
69250   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69260     .RunProgramAfterSavingProgramname = ""
69270    Else
69280     If LenB(tStr) > 0 Then
69290      .RunProgramAfterSavingProgramname = tStr
69300     End If
69310   End If
69320   tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramParameters")
69330   If LenB(tStr) = 0 And LenB("""<OutputFilename>""") > 0 And UseStandard Then
69340     .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
69350    Else
69360     If LenB(tStr) > 0 Then
69370      .RunProgramAfterSavingProgramParameters = tStr
69380     End If
69390   End If
69400   tStr = reg.GetRegistryValue("RunProgramAfterSavingWaitUntilReady")
69410   If IsNumeric(tStr) Then
69420     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69430       .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
69440      Else
69450       If UseStandard Then
69460        .RunProgramAfterSavingWaitUntilReady = 1
69470       End If
69480     End If
69490    Else
69500     If UseStandard Then
69510      .RunProgramAfterSavingWaitUntilReady = 1
69520     End If
69530   End If
69540   tStr = reg.GetRegistryValue("RunProgramAfterSavingWindowstyle")
69550   If IsNumeric(tStr) Then
69560     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
69570       .RunProgramAfterSavingWindowstyle = CLng(tStr)
69580      Else
69590       If UseStandard Then
69600        .RunProgramAfterSavingWindowstyle = 1
69610       End If
69620     End If
69630    Else
69640     If UseStandard Then
69650      .RunProgramAfterSavingWindowstyle = 1
69660     End If
69670   End If
69680   tStr = reg.GetRegistryValue("RunProgramBeforeSaving")
69690   If IsNumeric(tStr) Then
69700     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69710       .RunProgramBeforeSaving = CLng(tStr)
69720      Else
69730       If UseStandard Then
69740        .RunProgramBeforeSaving = 0
69750       End If
69760     End If
69770    Else
69780     If UseStandard Then
69790      .RunProgramBeforeSaving = 0
69800     End If
69810   End If
69820   tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramname")
69830   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69840     .RunProgramBeforeSavingProgramname = ""
69850    Else
69860     If LenB(tStr) > 0 Then
69870      .RunProgramBeforeSavingProgramname = tStr
69880     End If
69890   End If
69900   tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramParameters")
69910   If LenB(tStr) = 0 And LenB("""<TempFilename>""") > 0 And UseStandard Then
69920     .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
69930    Else
69940     If LenB(tStr) > 0 Then
69950      .RunProgramBeforeSavingProgramParameters = tStr
69960     End If
69970   End If
69980   tStr = reg.GetRegistryValue("RunProgramBeforeSavingWindowstyle")
69990   If IsNumeric(tStr) Then
70000     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
70010       .RunProgramBeforeSavingWindowstyle = CLng(tStr)
70020      Else
70030       If UseStandard Then
70040        .RunProgramBeforeSavingWindowstyle = 1
70050       End If
70060     End If
70070    Else
70080     If UseStandard Then
70090      .RunProgramBeforeSavingWindowstyle = 1
70100     End If
70110   End If
70120   tStr = reg.GetRegistryValue("SaveFilename")
70130   If LenB(tStr) = 0 And LenB("<Title>") > 0 And UseStandard Then
70140     .SaveFilename = "<Title>"
70150    Else
70160     If LenB(tStr) > 0 Then
70170      .SaveFilename = tStr
70180     End If
70190   End If
70200   tStr = reg.GetRegistryValue("SendEmailAfterAutoSaving")
70210   If IsNumeric(tStr) Then
70220     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70230       .SendEmailAfterAutoSaving = CLng(tStr)
70240      Else
70250       If UseStandard Then
70260        .SendEmailAfterAutoSaving = 0
70270       End If
70280     End If
70290    Else
70300     If UseStandard Then
70310      .SendEmailAfterAutoSaving = 0
70320     End If
70330   End If
70340   tStr = reg.GetRegistryValue("SendMailMethod")
70350   If IsNumeric(tStr) Then
70360     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
70370       .SendMailMethod = CLng(tStr)
70380      Else
70390       If UseStandard Then
70400        .SendMailMethod = 0
70410       End If
70420     End If
70430    Else
70440     If UseStandard Then
70450      .SendMailMethod = 0
70460     End If
70470   End If
70480   tStr = reg.GetRegistryValue("ShowAnimation")
70490   If IsNumeric(tStr) Then
70500     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70510       .ShowAnimation = CLng(tStr)
70520      Else
70530       If UseStandard Then
70540        .ShowAnimation = 1
70550       End If
70560     End If
70570    Else
70580     If UseStandard Then
70590      .ShowAnimation = 1
70600     End If
70610   End If
70620   tStr = reg.GetRegistryValue("StartStandardProgram")
70630   If IsNumeric(tStr) Then
70640     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70650       .StartStandardProgram = CLng(tStr)
70660      Else
70670       If UseStandard Then
70680        .StartStandardProgram = 1
70690       End If
70700     End If
70710    Else
70720     If UseStandard Then
70730      .StartStandardProgram = 1
70740     End If
70750   End If
70760   tStr = reg.GetRegistryValue("Toolbars")
70770   If IsNumeric(tStr) Then
70780     If CLng(tStr) >= 0 Then
70790       .Toolbars = CLng(tStr)
70800      Else
70810       If UseStandard Then
70820        .Toolbars = 1
70830       End If
70840     End If
70850    Else
70860     If UseStandard Then
70870      .Toolbars = 1
70880     End If
70890   End If
70900   tStr = reg.GetRegistryValue("UseAutosave")
70910   If IsNumeric(tStr) Then
70920     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70930       .UseAutosave = CLng(tStr)
70940      Else
70950       If UseStandard Then
70960        .UseAutosave = 0
70970       End If
70980     End If
70990    Else
71000     If UseStandard Then
71010      .UseAutosave = 0
71020     End If
71030   End If
71040   tStr = reg.GetRegistryValue("UseAutosaveDirectory")
71050   If IsNumeric(tStr) Then
71060     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71070       .UseAutosaveDirectory = CLng(tStr)
71080      Else
71090       If UseStandard Then
71100        .UseAutosaveDirectory = 1
71110       End If
71120     End If
71130    Else
71140     If UseStandard Then
71150      .UseAutosaveDirectory = 1
71160     End If
71170   End If
71180  End With
71190  Set reg = Nothing
71200  ReadOptionsReg = myOptions
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
52330   If UCase$(OptionName) = "BITMAPRESOLUTION" Then
52340    If Not reg.KeyExists Then
52350     reg.CreateKey
52360    End If
52370    reg.SetRegistryValue "BitmapResolution", CStr(.BitmapResolution), REG_SZ
52380    Set reg = Nothing
52390    Exit Sub
52400   End If
52410   If UCase$(OptionName) = "BMPCOLORSCOUNT" Then
52420    If Not reg.KeyExists Then
52430     reg.CreateKey
52440    End If
52450    reg.SetRegistryValue "BMPColorscount", CStr(.BMPColorscount), REG_SZ
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
52650   If UCase$(OptionName) = "PCLCOLORSCOUNT" Then
52660    If Not reg.KeyExists Then
52670     reg.CreateKey
52680    End If
52690    reg.SetRegistryValue "PCLColorsCount", CStr(.PCLColorsCount), REG_SZ
52700    Set reg = Nothing
52710    Exit Sub
52720   End If
52730   If UCase$(OptionName) = "PCXCOLORSCOUNT" Then
52740    If Not reg.KeyExists Then
52750     reg.CreateKey
52760    End If
52770    reg.SetRegistryValue "PCXColorscount", CStr(.PCXColorscount), REG_SZ
52780    Set reg = Nothing
52790    Exit Sub
52800   End If
52810   If UCase$(OptionName) = "PNGCOLORSCOUNT" Then
52820    If Not reg.KeyExists Then
52830     reg.CreateKey
52840    End If
52850    reg.SetRegistryValue "PNGColorscount", CStr(.PNGColorscount), REG_SZ
52860    Set reg = Nothing
52870    Exit Sub
52880   End If
52890   If UCase$(OptionName) = "PSDCOLORSCOUNT" Then
52900    If Not reg.KeyExists Then
52910     reg.CreateKey
52920    End If
52930    reg.SetRegistryValue "PSDColorsCount", CStr(.PSDColorsCount), REG_SZ
52940    Set reg = Nothing
52950    Exit Sub
52960   End If
52970   If UCase$(OptionName) = "RAWCOLORSCOUNT" Then
52980    If Not reg.KeyExists Then
52990     reg.CreateKey
53000    End If
53010    reg.SetRegistryValue "RAWColorsCount", CStr(.RAWColorsCount), REG_SZ
53020    Set reg = Nothing
53030    Exit Sub
53040   End If
53050   If UCase$(OptionName) = "TIFFCOLORSCOUNT" Then
53060    If Not reg.KeyExists Then
53070     reg.CreateKey
53080    End If
53090    reg.SetRegistryValue "TIFFColorscount", CStr(.TIFFColorscount), REG_SZ
53100    Set reg = Nothing
53110    Exit Sub
53120   End If
53130   If UCase$(OptionName) = "XCFCOLORSCOUNT" Then
53140    If Not reg.KeyExists Then
53150     reg.CreateKey
53160    End If
53170    reg.SetRegistryValue "XCFColorsCount", CStr(.XCFColorsCount), REG_SZ
53180    Set reg = Nothing
53190    Exit Sub
53200   End If
53210   reg.Subkey = "Printing\Formats\PDF\Colors"
53220   If UCase$(OptionName) = "PDFCOLORSCMYKTORGB" Then
53230    If Not reg.KeyExists Then
53240     reg.CreateKey
53250    End If
53260    reg.SetRegistryValue "PDFColorsCMYKToRGB", CStr(Abs(.PDFColorsCMYKToRGB)), REG_SZ
53270    Set reg = Nothing
53280    Exit Sub
53290   End If
53300   If UCase$(OptionName) = "PDFCOLORSCOLORMODEL" Then
53310    If Not reg.KeyExists Then
53320     reg.CreateKey
53330    End If
53340    reg.SetRegistryValue "PDFColorsColorModel", CStr(.PDFColorsColorModel), REG_SZ
53350    Set reg = Nothing
53360    Exit Sub
53370   End If
53380   If UCase$(OptionName) = "PDFCOLORSPRESERVEHALFTONE" Then
53390    If Not reg.KeyExists Then
53400     reg.CreateKey
53410    End If
53420    reg.SetRegistryValue "PDFColorsPreserveHalftone", CStr(Abs(.PDFColorsPreserveHalftone)), REG_SZ
53430    Set reg = Nothing
53440    Exit Sub
53450   End If
53460   If UCase$(OptionName) = "PDFCOLORSPRESERVEOVERPRINT" Then
53470    If Not reg.KeyExists Then
53480     reg.CreateKey
53490    End If
53500    reg.SetRegistryValue "PDFColorsPreserveOverprint", CStr(Abs(.PDFColorsPreserveOverprint)), REG_SZ
53510    Set reg = Nothing
53520    Exit Sub
53530   End If
53540   If UCase$(OptionName) = "PDFCOLORSPRESERVETRANSFER" Then
53550    If Not reg.KeyExists Then
53560     reg.CreateKey
53570    End If
53580    reg.SetRegistryValue "PDFColorsPreserveTransfer", CStr(Abs(.PDFColorsPreserveTransfer)), REG_SZ
53590    Set reg = Nothing
53600    Exit Sub
53610   End If
53620   reg.Subkey = "Printing\Formats\PDF\Compression"
53630   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSION" Then
53640    If Not reg.KeyExists Then
53650     reg.CreateKey
53660    End If
53670    reg.SetRegistryValue "PDFCompressionColorCompression", CStr(Abs(.PDFCompressionColorCompression)), REG_SZ
53680    Set reg = Nothing
53690    Exit Sub
53700   End If
53710   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE" Then
53720    If Not reg.KeyExists Then
53730     reg.CreateKey
53740    End If
53750    reg.SetRegistryValue "PDFCompressionColorCompressionChoice", CStr(.PDFCompressionColorCompressionChoice), REG_SZ
53760    Set reg = Nothing
53770    Exit Sub
53780   End If
53790   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR" Then
53800    If Not reg.KeyExists Then
53810     reg.CreateKey
53820    End If
53830   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
53840    Set reg = Nothing
53850    Exit Sub
53860   End If
53870   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR" Then
53880    If Not reg.KeyExists Then
53890     reg.CreateKey
53900    End If
53910   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
53920    Set reg = Nothing
53930    Exit Sub
53940   End If
53950   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR" Then
53960    If Not reg.KeyExists Then
53970     reg.CreateKey
53980    End If
53990   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
54000    Set reg = Nothing
54010    Exit Sub
54020   End If
54030   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR" Then
54040    If Not reg.KeyExists Then
54050     reg.CreateKey
54060    End If
54070   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
54080    Set reg = Nothing
54090    Exit Sub
54100   End If
54110   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR" Then
54120    If Not reg.KeyExists Then
54130     reg.CreateKey
54140    End If
54150   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
54160    Set reg = Nothing
54170    Exit Sub
54180   End If
54190   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESAMPLE" Then
54200    If Not reg.KeyExists Then
54210     reg.CreateKey
54220    End If
54230    reg.SetRegistryValue "PDFCompressionColorResample", CStr(Abs(.PDFCompressionColorResample)), REG_SZ
54240    Set reg = Nothing
54250    Exit Sub
54260   End If
54270   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESAMPLECHOICE" Then
54280    If Not reg.KeyExists Then
54290     reg.CreateKey
54300    End If
54310    reg.SetRegistryValue "PDFCompressionColorResampleChoice", CStr(.PDFCompressionColorResampleChoice), REG_SZ
54320    Set reg = Nothing
54330    Exit Sub
54340   End If
54350   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESOLUTION" Then
54360    If Not reg.KeyExists Then
54370     reg.CreateKey
54380    End If
54390    reg.SetRegistryValue "PDFCompressionColorResolution", CStr(.PDFCompressionColorResolution), REG_SZ
54400    Set reg = Nothing
54410    Exit Sub
54420   End If
54430   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSION" Then
54440    If Not reg.KeyExists Then
54450     reg.CreateKey
54460    End If
54470    reg.SetRegistryValue "PDFCompressionGreyCompression", CStr(Abs(.PDFCompressionGreyCompression)), REG_SZ
54480    Set reg = Nothing
54490    Exit Sub
54500   End If
54510   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE" Then
54520    If Not reg.KeyExists Then
54530     reg.CreateKey
54540    End If
54550    reg.SetRegistryValue "PDFCompressionGreyCompressionChoice", CStr(.PDFCompressionGreyCompressionChoice), REG_SZ
54560    Set reg = Nothing
54570    Exit Sub
54580   End If
54590   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR" Then
54600    If Not reg.KeyExists Then
54610     reg.CreateKey
54620    End If
54630   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
54640    Set reg = Nothing
54650    Exit Sub
54660   End If
54670   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR" Then
54680    If Not reg.KeyExists Then
54690     reg.CreateKey
54700    End If
54710   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
54720    Set reg = Nothing
54730    Exit Sub
54740   End If
54750   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR" Then
54760    If Not reg.KeyExists Then
54770     reg.CreateKey
54780    End If
54790   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
54800    Set reg = Nothing
54810    Exit Sub
54820   End If
54830   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR" Then
54840    If Not reg.KeyExists Then
54850     reg.CreateKey
54860    End If
54870   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
54880    Set reg = Nothing
54890    Exit Sub
54900   End If
54910   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR" Then
54920    If Not reg.KeyExists Then
54930     reg.CreateKey
54940    End If
54950   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
54960    Set reg = Nothing
54970    Exit Sub
54980   End If
54990   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESAMPLE" Then
55000    If Not reg.KeyExists Then
55010     reg.CreateKey
55020    End If
55030    reg.SetRegistryValue "PDFCompressionGreyResample", CStr(Abs(.PDFCompressionGreyResample)), REG_SZ
55040    Set reg = Nothing
55050    Exit Sub
55060   End If
55070   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESAMPLECHOICE" Then
55080    If Not reg.KeyExists Then
55090     reg.CreateKey
55100    End If
55110    reg.SetRegistryValue "PDFCompressionGreyResampleChoice", CStr(.PDFCompressionGreyResampleChoice), REG_SZ
55120    Set reg = Nothing
55130    Exit Sub
55140   End If
55150   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESOLUTION" Then
55160    If Not reg.KeyExists Then
55170     reg.CreateKey
55180    End If
55190    reg.SetRegistryValue "PDFCompressionGreyResolution", CStr(.PDFCompressionGreyResolution), REG_SZ
55200    Set reg = Nothing
55210    Exit Sub
55220   End If
55230   If UCase$(OptionName) = "PDFCOMPRESSIONMONOCOMPRESSION" Then
55240    If Not reg.KeyExists Then
55250     reg.CreateKey
55260    End If
55270    reg.SetRegistryValue "PDFCompressionMonoCompression", CStr(Abs(.PDFCompressionMonoCompression)), REG_SZ
55280    Set reg = Nothing
55290    Exit Sub
55300   End If
55310   If UCase$(OptionName) = "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE" Then
55320    If Not reg.KeyExists Then
55330     reg.CreateKey
55340    End If
55350    reg.SetRegistryValue "PDFCompressionMonoCompressionChoice", CStr(.PDFCompressionMonoCompressionChoice), REG_SZ
55360    Set reg = Nothing
55370    Exit Sub
55380   End If
55390   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESAMPLE" Then
55400    If Not reg.KeyExists Then
55410     reg.CreateKey
55420    End If
55430    reg.SetRegistryValue "PDFCompressionMonoResample", CStr(Abs(.PDFCompressionMonoResample)), REG_SZ
55440    Set reg = Nothing
55450    Exit Sub
55460   End If
55470   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESAMPLECHOICE" Then
55480    If Not reg.KeyExists Then
55490     reg.CreateKey
55500    End If
55510    reg.SetRegistryValue "PDFCompressionMonoResampleChoice", CStr(.PDFCompressionMonoResampleChoice), REG_SZ
55520    Set reg = Nothing
55530    Exit Sub
55540   End If
55550   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESOLUTION" Then
55560    If Not reg.KeyExists Then
55570     reg.CreateKey
55580    End If
55590    reg.SetRegistryValue "PDFCompressionMonoResolution", CStr(.PDFCompressionMonoResolution), REG_SZ
55600    Set reg = Nothing
55610    Exit Sub
55620   End If
55630   If UCase$(OptionName) = "PDFCOMPRESSIONTEXTCOMPRESSION" Then
55640    If Not reg.KeyExists Then
55650     reg.CreateKey
55660    End If
55670    reg.SetRegistryValue "PDFCompressionTextCompression", CStr(Abs(.PDFCompressionTextCompression)), REG_SZ
55680    Set reg = Nothing
55690    Exit Sub
55700   End If
55710   reg.Subkey = "Printing\Formats\PDF\Fonts"
55720   If UCase$(OptionName) = "PDFFONTSEMBEDALL" Then
55730    If Not reg.KeyExists Then
55740     reg.CreateKey
55750    End If
55760    reg.SetRegistryValue "PDFFontsEmbedAll", CStr(Abs(.PDFFontsEmbedAll)), REG_SZ
55770    Set reg = Nothing
55780    Exit Sub
55790   End If
55800   If UCase$(OptionName) = "PDFFONTSSUBSETFONTS" Then
55810    If Not reg.KeyExists Then
55820     reg.CreateKey
55830    End If
55840    reg.SetRegistryValue "PDFFontsSubSetFonts", CStr(Abs(.PDFFontsSubSetFonts)), REG_SZ
55850    Set reg = Nothing
55860    Exit Sub
55870   End If
55880   If UCase$(OptionName) = "PDFFONTSSUBSETFONTSPERCENT" Then
55890    If Not reg.KeyExists Then
55900     reg.CreateKey
55910    End If
55920    reg.SetRegistryValue "PDFFontsSubSetFontsPercent", CStr(.PDFFontsSubSetFontsPercent), REG_SZ
55930    Set reg = Nothing
55940    Exit Sub
55950   End If
55960   reg.Subkey = "Printing\Formats\PDF\General"
55970   If UCase$(OptionName) = "PDFGENERALASCII85" Then
55980    If Not reg.KeyExists Then
55990     reg.CreateKey
56000    End If
56010    reg.SetRegistryValue "PDFGeneralASCII85", CStr(Abs(.PDFGeneralASCII85)), REG_SZ
56020    Set reg = Nothing
56030    Exit Sub
56040   End If
56050   If UCase$(OptionName) = "PDFGENERALAUTOROTATE" Then
56060    If Not reg.KeyExists Then
56070     reg.CreateKey
56080    End If
56090    reg.SetRegistryValue "PDFGeneralAutorotate", CStr(.PDFGeneralAutorotate), REG_SZ
56100    Set reg = Nothing
56110    Exit Sub
56120   End If
56130   If UCase$(OptionName) = "PDFGENERALCOMPATIBILITY" Then
56140    If Not reg.KeyExists Then
56150     reg.CreateKey
56160    End If
56170    reg.SetRegistryValue "PDFGeneralCompatibility", CStr(.PDFGeneralCompatibility), REG_SZ
56180    Set reg = Nothing
56190    Exit Sub
56200   End If
56210   If UCase$(OptionName) = "PDFGENERALDEFAULT" Then
56220    If Not reg.KeyExists Then
56230     reg.CreateKey
56240    End If
56250    reg.SetRegistryValue "PDFGeneralDefault", CStr(.PDFGeneralDefault), REG_SZ
56260    Set reg = Nothing
56270    Exit Sub
56280   End If
56290   If UCase$(OptionName) = "PDFGENERALOVERPRINT" Then
56300    If Not reg.KeyExists Then
56310     reg.CreateKey
56320    End If
56330    reg.SetRegistryValue "PDFGeneralOverprint", CStr(.PDFGeneralOverprint), REG_SZ
56340    Set reg = Nothing
56350    Exit Sub
56360   End If
56370   If UCase$(OptionName) = "PDFGENERALRESOLUTION" Then
56380    If Not reg.KeyExists Then
56390     reg.CreateKey
56400    End If
56410    reg.SetRegistryValue "PDFGeneralResolution", CStr(.PDFGeneralResolution), REG_SZ
56420    Set reg = Nothing
56430    Exit Sub
56440   End If
56450   If UCase$(OptionName) = "PDFOPTIMIZE" Then
56460    If Not reg.KeyExists Then
56470     reg.CreateKey
56480    End If
56490    reg.SetRegistryValue "PDFOptimize", CStr(Abs(.PDFOptimize)), REG_SZ
56500    Set reg = Nothing
56510    Exit Sub
56520   End If
56530   If UCase$(OptionName) = "PDFUPDATEMETADATA" Then
56540    If Not reg.KeyExists Then
56550     reg.CreateKey
56560    End If
56570    reg.SetRegistryValue "PDFUpdateMetadata", CStr(.PDFUpdateMetadata), REG_SZ
56580    Set reg = Nothing
56590    Exit Sub
56600   End If
56610   reg.Subkey = "Printing\Formats\PDF\Security"
56620   If UCase$(OptionName) = "PDFALLOWASSEMBLY" Then
56630    If Not reg.KeyExists Then
56640     reg.CreateKey
56650    End If
56660    reg.SetRegistryValue "PDFAllowAssembly", CStr(Abs(.PDFAllowAssembly)), REG_SZ
56670    Set reg = Nothing
56680    Exit Sub
56690   End If
56700   If UCase$(OptionName) = "PDFALLOWDEGRADEDPRINTING" Then
56710    If Not reg.KeyExists Then
56720     reg.CreateKey
56730    End If
56740    reg.SetRegistryValue "PDFAllowDegradedPrinting", CStr(Abs(.PDFAllowDegradedPrinting)), REG_SZ
56750    Set reg = Nothing
56760    Exit Sub
56770   End If
56780   If UCase$(OptionName) = "PDFALLOWFILLIN" Then
56790    If Not reg.KeyExists Then
56800     reg.CreateKey
56810    End If
56820    reg.SetRegistryValue "PDFAllowFillIn", CStr(Abs(.PDFAllowFillIn)), REG_SZ
56830    Set reg = Nothing
56840    Exit Sub
56850   End If
56860   If UCase$(OptionName) = "PDFALLOWSCREENREADERS" Then
56870    If Not reg.KeyExists Then
56880     reg.CreateKey
56890    End If
56900    reg.SetRegistryValue "PDFAllowScreenReaders", CStr(Abs(.PDFAllowScreenReaders)), REG_SZ
56910    Set reg = Nothing
56920    Exit Sub
56930   End If
56940   If UCase$(OptionName) = "PDFDISALLOWCOPY" Then
56950    If Not reg.KeyExists Then
56960     reg.CreateKey
56970    End If
56980    reg.SetRegistryValue "PDFDisallowCopy", CStr(Abs(.PDFDisallowCopy)), REG_SZ
56990    Set reg = Nothing
57000    Exit Sub
57010   End If
57020   If UCase$(OptionName) = "PDFDISALLOWMODIFYANNOTATIONS" Then
57030    If Not reg.KeyExists Then
57040     reg.CreateKey
57050    End If
57060    reg.SetRegistryValue "PDFDisallowModifyAnnotations", CStr(Abs(.PDFDisallowModifyAnnotations)), REG_SZ
57070    Set reg = Nothing
57080    Exit Sub
57090   End If
57100   If UCase$(OptionName) = "PDFDISALLOWMODIFYCONTENTS" Then
57110    If Not reg.KeyExists Then
57120     reg.CreateKey
57130    End If
57140    reg.SetRegistryValue "PDFDisallowModifyContents", CStr(Abs(.PDFDisallowModifyContents)), REG_SZ
57150    Set reg = Nothing
57160    Exit Sub
57170   End If
57180   If UCase$(OptionName) = "PDFDISALLOWPRINTING" Then
57190    If Not reg.KeyExists Then
57200     reg.CreateKey
57210    End If
57220    reg.SetRegistryValue "PDFDisallowPrinting", CStr(Abs(.PDFDisallowPrinting)), REG_SZ
57230    Set reg = Nothing
57240    Exit Sub
57250   End If
57260   If UCase$(OptionName) = "PDFENCRYPTOR" Then
57270    If Not reg.KeyExists Then
57280     reg.CreateKey
57290    End If
57300    reg.SetRegistryValue "PDFEncryptor", CStr(.PDFEncryptor), REG_SZ
57310    Set reg = Nothing
57320    Exit Sub
57330   End If
57340   If UCase$(OptionName) = "PDFHIGHENCRYPTION" Then
57350    If Not reg.KeyExists Then
57360     reg.CreateKey
57370    End If
57380    reg.SetRegistryValue "PDFHighEncryption", CStr(Abs(.PDFHighEncryption)), REG_SZ
57390    Set reg = Nothing
57400    Exit Sub
57410   End If
57420   If UCase$(OptionName) = "PDFLOWENCRYPTION" Then
57430    If Not reg.KeyExists Then
57440     reg.CreateKey
57450    End If
57460    reg.SetRegistryValue "PDFLowEncryption", CStr(Abs(.PDFLowEncryption)), REG_SZ
57470    Set reg = Nothing
57480    Exit Sub
57490   End If
57500   If UCase$(OptionName) = "PDFOWNERPASS" Then
57510    If Not reg.KeyExists Then
57520     reg.CreateKey
57530    End If
57540    reg.SetRegistryValue "PDFOwnerPass", CStr(Abs(.PDFOwnerPass)), REG_SZ
57550    Set reg = Nothing
57560    Exit Sub
57570   End If
57580   If UCase$(OptionName) = "PDFOWNERPASSWORDSTRING" Then
57590    If Not reg.KeyExists Then
57600     reg.CreateKey
57610    End If
57620    reg.SetRegistryValue "PDFOwnerPasswordString", CStr(.PDFOwnerPasswordString), REG_SZ
57630    Set reg = Nothing
57640    Exit Sub
57650   End If
57660   If UCase$(OptionName) = "PDFUSERPASS" Then
57670    If Not reg.KeyExists Then
57680     reg.CreateKey
57690    End If
57700    reg.SetRegistryValue "PDFUserPass", CStr(Abs(.PDFUserPass)), REG_SZ
57710    Set reg = Nothing
57720    Exit Sub
57730   End If
57740   If UCase$(OptionName) = "PDFUSERPASSWORDSTRING" Then
57750    If Not reg.KeyExists Then
57760     reg.CreateKey
57770    End If
57780    reg.SetRegistryValue "PDFUserPasswordString", CStr(.PDFUserPasswordString), REG_SZ
57790    Set reg = Nothing
57800    Exit Sub
57810   End If
57820   If UCase$(OptionName) = "PDFUSESECURITY" Then
57830    If Not reg.KeyExists Then
57840     reg.CreateKey
57850    End If
57860    reg.SetRegistryValue "PDFUseSecurity", CStr(Abs(.PDFUseSecurity)), REG_SZ
57870    Set reg = Nothing
57880    Exit Sub
57890   End If
57900   reg.Subkey = "Printing\Formats\PDF\Signing"
57910   If UCase$(OptionName) = "PDFSIGNINGMULTISIGNATURE" Then
57920    If Not reg.KeyExists Then
57930     reg.CreateKey
57940    End If
57950    reg.SetRegistryValue "PDFSigningMultiSignature", CStr(Abs(.PDFSigningMultiSignature)), REG_SZ
57960    Set reg = Nothing
57970    Exit Sub
57980   End If
57990   If UCase$(OptionName) = "PDFSIGNINGPFXFILE" Then
58000    If Not reg.KeyExists Then
58010     reg.CreateKey
58020    End If
58030    reg.SetRegistryValue "PDFSigningPFXFile", CStr(.PDFSigningPFXFile), REG_SZ
58040    Set reg = Nothing
58050    Exit Sub
58060   End If
58070   If UCase$(OptionName) = "PDFSIGNINGPFXFILEPASSWORD" Then
58080    If Not reg.KeyExists Then
58090     reg.CreateKey
58100    End If
58110    reg.SetRegistryValue "PDFSigningPFXFilePassword", CStr(.PDFSigningPFXFilePassword), REG_SZ
58120    Set reg = Nothing
58130    Exit Sub
58140   End If
58150   If UCase$(OptionName) = "PDFSIGNINGSIGNATURECONTACT" Then
58160    If Not reg.KeyExists Then
58170     reg.CreateKey
58180    End If
58190    reg.SetRegistryValue "PDFSigningSignatureContact", CStr(.PDFSigningSignatureContact), REG_SZ
58200    Set reg = Nothing
58210    Exit Sub
58220   End If
58230   If UCase$(OptionName) = "PDFSIGNINGSIGNATURELEFTX" Then
58240    If Not reg.KeyExists Then
58250     reg.CreateKey
58260    End If
58270   reg.SetRegistryValue "PDFSigningSignatureLeftX", Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), REG_SZ
58280    Set reg = Nothing
58290    Exit Sub
58300   End If
58310   If UCase$(OptionName) = "PDFSIGNINGSIGNATURELEFTY" Then
58320    If Not reg.KeyExists Then
58330     reg.CreateKey
58340    End If
58350   reg.SetRegistryValue "PDFSigningSignatureLeftY", Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), REG_SZ
58360    Set reg = Nothing
58370    Exit Sub
58380   End If
58390   If UCase$(OptionName) = "PDFSIGNINGSIGNATURELOCATION" Then
58400    If Not reg.KeyExists Then
58410     reg.CreateKey
58420    End If
58430    reg.SetRegistryValue "PDFSigningSignatureLocation", CStr(.PDFSigningSignatureLocation), REG_SZ
58440    Set reg = Nothing
58450    Exit Sub
58460   End If
58470   If UCase$(OptionName) = "PDFSIGNINGSIGNATUREREASON" Then
58480    If Not reg.KeyExists Then
58490     reg.CreateKey
58500    End If
58510    reg.SetRegistryValue "PDFSigningSignatureReason", CStr(.PDFSigningSignatureReason), REG_SZ
58520    Set reg = Nothing
58530    Exit Sub
58540   End If
58550   If UCase$(OptionName) = "PDFSIGNINGSIGNATURERIGHTX" Then
58560    If Not reg.KeyExists Then
58570     reg.CreateKey
58580    End If
58590   reg.SetRegistryValue "PDFSigningSignatureRightX", Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), REG_SZ
58600    Set reg = Nothing
58610    Exit Sub
58620   End If
58630   If UCase$(OptionName) = "PDFSIGNINGSIGNATURERIGHTY" Then
58640    If Not reg.KeyExists Then
58650     reg.CreateKey
58660    End If
58670   reg.SetRegistryValue "PDFSigningSignatureRightY", Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), REG_SZ
58680    Set reg = Nothing
58690    Exit Sub
58700   End If
58710   If UCase$(OptionName) = "PDFSIGNINGSIGNATUREVISIBLE" Then
58720    If Not reg.KeyExists Then
58730     reg.CreateKey
58740    End If
58750    reg.SetRegistryValue "PDFSigningSignatureVisible", CStr(Abs(.PDFSigningSignatureVisible)), REG_SZ
58760    Set reg = Nothing
58770    Exit Sub
58780   End If
58790   If UCase$(OptionName) = "PDFSIGNINGSIGNPDF" Then
58800    If Not reg.KeyExists Then
58810     reg.CreateKey
58820    End If
58830    reg.SetRegistryValue "PDFSigningSignPDF", CStr(Abs(.PDFSigningSignPDF)), REG_SZ
58840    Set reg = Nothing
58850    Exit Sub
58860   End If
58870   reg.Subkey = "Printing\Formats\PS\LanguageLevel"
58880   If UCase$(OptionName) = "EPSLANGUAGELEVEL" Then
58890    If Not reg.KeyExists Then
58900     reg.CreateKey
58910    End If
58920    reg.SetRegistryValue "EPSLanguageLevel", CStr(.EPSLanguageLevel), REG_SZ
58930    Set reg = Nothing
58940    Exit Sub
58950   End If
58960   If UCase$(OptionName) = "PSLANGUAGELEVEL" Then
58970    If Not reg.KeyExists Then
58980     reg.CreateKey
58990    End If
59000    reg.SetRegistryValue "PSLanguageLevel", CStr(.PSLanguageLevel), REG_SZ
59010    Set reg = Nothing
59020    Exit Sub
59030   End If
59040   reg.Subkey = "Program"
59050   If UCase$(OptionName) = "ADDITIONALGHOSTSCRIPTPARAMETERS" Then
59060    If Not reg.KeyExists Then
59070     reg.CreateKey
59080    End If
59090    reg.SetRegistryValue "AdditionalGhostscriptParameters", CStr(.AdditionalGhostscriptParameters), REG_SZ
59100    Set reg = Nothing
59110    Exit Sub
59120   End If
59130   If UCase$(OptionName) = "ADDITIONALGHOSTSCRIPTSEARCHPATH" Then
59140    If Not reg.KeyExists Then
59150     reg.CreateKey
59160    End If
59170    reg.SetRegistryValue "AdditionalGhostscriptSearchpath", CStr(.AdditionalGhostscriptSearchpath), REG_SZ
59180    Set reg = Nothing
59190    Exit Sub
59200   End If
59210   If UCase$(OptionName) = "ADDWINDOWSFONTPATH" Then
59220    If Not reg.KeyExists Then
59230     reg.CreateKey
59240    End If
59250    reg.SetRegistryValue "AddWindowsFontpath", CStr(Abs(.AddWindowsFontpath)), REG_SZ
59260    Set reg = Nothing
59270    Exit Sub
59280   End If
59290   If UCase$(OptionName) = "AUTOSAVEDIRECTORY" Then
59300    If Not reg.KeyExists Then
59310     reg.CreateKey
59320    End If
59330    reg.SetRegistryValue "AutosaveDirectory", CStr(.AutosaveDirectory), REG_SZ
59340    Set reg = Nothing
59350    Exit Sub
59360   End If
59370   If UCase$(OptionName) = "AUTOSAVEFILENAME" Then
59380    If Not reg.KeyExists Then
59390     reg.CreateKey
59400    End If
59410    reg.SetRegistryValue "AutosaveFilename", CStr(.AutosaveFilename), REG_SZ
59420    Set reg = Nothing
59430    Exit Sub
59440   End If
59450   If UCase$(OptionName) = "AUTOSAVEFORMAT" Then
59460    If Not reg.KeyExists Then
59470     reg.CreateKey
59480    End If
59490    reg.SetRegistryValue "AutosaveFormat", CStr(.AutosaveFormat), REG_SZ
59500    Set reg = Nothing
59510    Exit Sub
59520   End If
59530   If UCase$(OptionName) = "AUTOSAVESTARTSTANDARDPROGRAM" Then
59540    If Not reg.KeyExists Then
59550     reg.CreateKey
59560    End If
59570    reg.SetRegistryValue "AutosaveStartStandardProgram", CStr(Abs(.AutosaveStartStandardProgram)), REG_SZ
59580    Set reg = Nothing
59590    Exit Sub
59600   End If
59610   If UCase$(OptionName) = "CLIENTCOMPUTERRESOLVEIPADDRESS" Then
59620    If Not reg.KeyExists Then
59630     reg.CreateKey
59640    End If
59650    reg.SetRegistryValue "ClientComputerResolveIPAddress", CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
59660    Set reg = Nothing
59670    Exit Sub
59680   End If
59690   If UCase$(OptionName) = "DISABLEEMAIL" Then
59700    If Not reg.KeyExists Then
59710     reg.CreateKey
59720    End If
59730    reg.SetRegistryValue "DisableEmail", CStr(Abs(.DisableEmail)), REG_SZ
59740    Set reg = Nothing
59750    Exit Sub
59760   End If
59770   If UCase$(OptionName) = "DONTUSEDOCUMENTSETTINGS" Then
59780    If Not reg.KeyExists Then
59790     reg.CreateKey
59800    End If
59810    reg.SetRegistryValue "DontUseDocumentSettings", CStr(Abs(.DontUseDocumentSettings)), REG_SZ
59820    Set reg = Nothing
59830    Exit Sub
59840   End If
59850   If UCase$(OptionName) = "FILENAMESUBSTITUTIONS" Then
59860    If Not reg.KeyExists Then
59870     reg.CreateKey
59880    End If
59890    reg.SetRegistryValue "FilenameSubstitutions", CStr(.FilenameSubstitutions), REG_SZ
59900    Set reg = Nothing
59910    Exit Sub
59920   End If
59930   If UCase$(OptionName) = "FILENAMESUBSTITUTIONSONLYINTITLE" Then
59940    If Not reg.KeyExists Then
59950     reg.CreateKey
59960    End If
59970    reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle", CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
59980    Set reg = Nothing
59990    Exit Sub
60000   End If
60010   If UCase$(OptionName) = "LANGUAGE" Then
60020    If Not reg.KeyExists Then
60030     reg.CreateKey
60040    End If
60050    reg.SetRegistryValue "Language", CStr(.Language), REG_SZ
60060    Set reg = Nothing
60070    Exit Sub
60080   End If
60090   If UCase$(OptionName) = "LASTSAVEDIRECTORY" Then
60100    If Not reg.KeyExists Then
60110     reg.CreateKey
60120    End If
60130    reg.SetRegistryValue "LastSaveDirectory", CStr(.LastSaveDirectory), REG_SZ
60140    Set reg = Nothing
60150    Exit Sub
60160   End If
60170   If UCase$(OptionName) = "LOGGING" Then
60180    If Not reg.KeyExists Then
60190     reg.CreateKey
60200    End If
60210    reg.SetRegistryValue "Logging", CStr(Abs(.Logging)), REG_SZ
60220    Set reg = Nothing
60230    Exit Sub
60240   End If
60250   If UCase$(OptionName) = "LOGLINES" Then
60260    If Not reg.KeyExists Then
60270     reg.CreateKey
60280    End If
60290    reg.SetRegistryValue "LogLines", CStr(.LogLines), REG_SZ
60300    Set reg = Nothing
60310    Exit Sub
60320   End If
60330   If UCase$(OptionName) = "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER" Then
60340    If Not reg.KeyExists Then
60350     reg.CreateKey
60360    End If
60370    reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter", CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
60380    Set reg = Nothing
60390    Exit Sub
60400   End If
60410   If UCase$(OptionName) = "NOPROCESSINGATSTARTUP" Then
60420    If Not reg.KeyExists Then
60430     reg.CreateKey
60440    End If
60450    reg.SetRegistryValue "NoProcessingAtStartup", CStr(Abs(.NoProcessingAtStartup)), REG_SZ
60460    Set reg = Nothing
60470    Exit Sub
60480   End If
60490   If UCase$(OptionName) = "NOPSCHECK" Then
60500    If Not reg.KeyExists Then
60510     reg.CreateKey
60520    End If
60530    reg.SetRegistryValue "NoPSCheck", CStr(Abs(.NoPSCheck)), REG_SZ
60540    Set reg = Nothing
60550    Exit Sub
60560   End If
60570   If UCase$(OptionName) = "OPTIONSDESIGN" Then
60580    If Not reg.KeyExists Then
60590     reg.CreateKey
60600    End If
60610    reg.SetRegistryValue "OptionsDesign", CStr(.OptionsDesign), REG_SZ
60620    Set reg = Nothing
60630    Exit Sub
60640   End If
60650   If UCase$(OptionName) = "OPTIONSENABLED" Then
60660    If Not reg.KeyExists Then
60670     reg.CreateKey
60680    End If
60690    reg.SetRegistryValue "OptionsEnabled", CStr(Abs(.OptionsEnabled)), REG_SZ
60700    Set reg = Nothing
60710    Exit Sub
60720   End If
60730   If UCase$(OptionName) = "OPTIONSVISIBLE" Then
60740    If Not reg.KeyExists Then
60750     reg.CreateKey
60760    End If
60770    reg.SetRegistryValue "OptionsVisible", CStr(Abs(.OptionsVisible)), REG_SZ
60780    Set reg = Nothing
60790    Exit Sub
60800   End If
60810   If UCase$(OptionName) = "PRINTAFTERSAVING" Then
60820    If Not reg.KeyExists Then
60830     reg.CreateKey
60840    End If
60850    reg.SetRegistryValue "PrintAfterSaving", CStr(Abs(.PrintAfterSaving)), REG_SZ
60860    Set reg = Nothing
60870    Exit Sub
60880   End If
60890   If UCase$(OptionName) = "PRINTAFTERSAVINGDUPLEX" Then
60900    If Not reg.KeyExists Then
60910     reg.CreateKey
60920    End If
60930    reg.SetRegistryValue "PrintAfterSavingDuplex", CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
60940    Set reg = Nothing
60950    Exit Sub
60960   End If
60970   If UCase$(OptionName) = "PRINTAFTERSAVINGNOCANCEL" Then
60980    If Not reg.KeyExists Then
60990     reg.CreateKey
61000    End If
61010    reg.SetRegistryValue "PrintAfterSavingNoCancel", CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
61020    Set reg = Nothing
61030    Exit Sub
61040   End If
61050   If UCase$(OptionName) = "PRINTAFTERSAVINGPRINTER" Then
61060    If Not reg.KeyExists Then
61070     reg.CreateKey
61080    End If
61090    reg.SetRegistryValue "PrintAfterSavingPrinter", CStr(.PrintAfterSavingPrinter), REG_SZ
61100    Set reg = Nothing
61110    Exit Sub
61120   End If
61130   If UCase$(OptionName) = "PRINTAFTERSAVINGQUERYUSER" Then
61140    If Not reg.KeyExists Then
61150     reg.CreateKey
61160    End If
61170    reg.SetRegistryValue "PrintAfterSavingQueryUser", CStr(.PrintAfterSavingQueryUser), REG_SZ
61180    Set reg = Nothing
61190    Exit Sub
61200   End If
61210   If UCase$(OptionName) = "PRINTAFTERSAVINGTUMBLE" Then
61220    If Not reg.KeyExists Then
61230     reg.CreateKey
61240    End If
61250    reg.SetRegistryValue "PrintAfterSavingTumble", CStr(.PrintAfterSavingTumble), REG_SZ
61260    Set reg = Nothing
61270    Exit Sub
61280   End If
61290   If UCase$(OptionName) = "PRINTERSTOP" Then
61300    If Not reg.KeyExists Then
61310     reg.CreateKey
61320    End If
61330    reg.SetRegistryValue "PrinterStop", CStr(Abs(.PrinterStop)), REG_SZ
61340    Set reg = Nothing
61350    Exit Sub
61360   End If
61370   If UCase$(OptionName) = "PRINTERTEMPPATH" Then
61380    If Not reg.KeyExists Then
61390     reg.CreateKey
61400    End If
61410    reg.SetRegistryValue "PrinterTemppath", CStr(.PrinterTemppath), REG_SZ
61420    Set reg = Nothing
61430    Exit Sub
61440   End If
61450   If UCase$(OptionName) = "PROCESSPRIORITY" Then
61460    If Not reg.KeyExists Then
61470     reg.CreateKey
61480    End If
61490    reg.SetRegistryValue "ProcessPriority", CStr(.ProcessPriority), REG_SZ
61500    Set reg = Nothing
61510    Exit Sub
61520   End If
61530   If UCase$(OptionName) = "PROGRAMFONT" Then
61540    If Not reg.KeyExists Then
61550     reg.CreateKey
61560    End If
61570    reg.SetRegistryValue "ProgramFont", CStr(.ProgramFont), REG_SZ
61580    Set reg = Nothing
61590    Exit Sub
61600   End If
61610   If UCase$(OptionName) = "PROGRAMFONTCHARSET" Then
61620    If Not reg.KeyExists Then
61630     reg.CreateKey
61640    End If
61650    reg.SetRegistryValue "ProgramFontCharset", CStr(.ProgramFontCharset), REG_SZ
61660    Set reg = Nothing
61670    Exit Sub
61680   End If
61690   If UCase$(OptionName) = "PROGRAMFONTSIZE" Then
61700    If Not reg.KeyExists Then
61710     reg.CreateKey
61720    End If
61730    reg.SetRegistryValue "ProgramFontSize", CStr(.ProgramFontSize), REG_SZ
61740    Set reg = Nothing
61750    Exit Sub
61760   End If
61770   If UCase$(OptionName) = "REMOVEALLKNOWNFILEEXTENSIONS" Then
61780    If Not reg.KeyExists Then
61790     reg.CreateKey
61800    End If
61810    reg.SetRegistryValue "RemoveAllKnownFileExtensions", CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
61820    Set reg = Nothing
61830    Exit Sub
61840   End If
61850   If UCase$(OptionName) = "REMOVESPACES" Then
61860    If Not reg.KeyExists Then
61870     reg.CreateKey
61880    End If
61890    reg.SetRegistryValue "RemoveSpaces", CStr(Abs(.RemoveSpaces)), REG_SZ
61900    Set reg = Nothing
61910    Exit Sub
61920   End If
61930   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVING" Then
61940    If Not reg.KeyExists Then
61950     reg.CreateKey
61960    End If
61970    reg.SetRegistryValue "RunProgramAfterSaving", CStr(Abs(.RunProgramAfterSaving)), REG_SZ
61980    Set reg = Nothing
61990    Exit Sub
62000   End If
62010   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGPROGRAMNAME" Then
62020    If Not reg.KeyExists Then
62030     reg.CreateKey
62040    End If
62050    reg.SetRegistryValue "RunProgramAfterSavingProgramname", CStr(.RunProgramAfterSavingProgramname), REG_SZ
62060    Set reg = Nothing
62070    Exit Sub
62080   End If
62090   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS" Then
62100    If Not reg.KeyExists Then
62110     reg.CreateKey
62120    End If
62130    reg.SetRegistryValue "RunProgramAfterSavingProgramParameters", CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
62140    Set reg = Nothing
62150    Exit Sub
62160   End If
62170   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGWAITUNTILREADY" Then
62180    If Not reg.KeyExists Then
62190     reg.CreateKey
62200    End If
62210    reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady", CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
62220    Set reg = Nothing
62230    Exit Sub
62240   End If
62250   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGWINDOWSTYLE" Then
62260    If Not reg.KeyExists Then
62270     reg.CreateKey
62280    End If
62290    reg.SetRegistryValue "RunProgramAfterSavingWindowstyle", CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
62300    Set reg = Nothing
62310    Exit Sub
62320   End If
62330   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVING" Then
62340    If Not reg.KeyExists Then
62350     reg.CreateKey
62360    End If
62370    reg.SetRegistryValue "RunProgramBeforeSaving", CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
62380    Set reg = Nothing
62390    Exit Sub
62400   End If
62410   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGPROGRAMNAME" Then
62420    If Not reg.KeyExists Then
62430     reg.CreateKey
62440    End If
62450    reg.SetRegistryValue "RunProgramBeforeSavingProgramname", CStr(.RunProgramBeforeSavingProgramname), REG_SZ
62460    Set reg = Nothing
62470    Exit Sub
62480   End If
62490   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS" Then
62500    If Not reg.KeyExists Then
62510     reg.CreateKey
62520    End If
62530    reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters", CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
62540    Set reg = Nothing
62550    Exit Sub
62560   End If
62570   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGWINDOWSTYLE" Then
62580    If Not reg.KeyExists Then
62590     reg.CreateKey
62600    End If
62610    reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle", CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
62620    Set reg = Nothing
62630    Exit Sub
62640   End If
62650   If UCase$(OptionName) = "SAVEFILENAME" Then
62660    If Not reg.KeyExists Then
62670     reg.CreateKey
62680    End If
62690    reg.SetRegistryValue "SaveFilename", CStr(.SaveFilename), REG_SZ
62700    Set reg = Nothing
62710    Exit Sub
62720   End If
62730   If UCase$(OptionName) = "SENDEMAILAFTERAUTOSAVING" Then
62740    If Not reg.KeyExists Then
62750     reg.CreateKey
62760    End If
62770    reg.SetRegistryValue "SendEmailAfterAutoSaving", CStr(Abs(.SendEmailAfterAutoSaving)), REG_SZ
62780    Set reg = Nothing
62790    Exit Sub
62800   End If
62810   If UCase$(OptionName) = "SENDMAILMETHOD" Then
62820    If Not reg.KeyExists Then
62830     reg.CreateKey
62840    End If
62850    reg.SetRegistryValue "SendMailMethod", CStr(.SendMailMethod), REG_SZ
62860    Set reg = Nothing
62870    Exit Sub
62880   End If
62890   If UCase$(OptionName) = "SHOWANIMATION" Then
62900    If Not reg.KeyExists Then
62910     reg.CreateKey
62920    End If
62930    reg.SetRegistryValue "ShowAnimation", CStr(Abs(.ShowAnimation)), REG_SZ
62940    Set reg = Nothing
62950    Exit Sub
62960   End If
62970   If UCase$(OptionName) = "STARTSTANDARDPROGRAM" Then
62980    If Not reg.KeyExists Then
62990     reg.CreateKey
63000    End If
63010    reg.SetRegistryValue "StartStandardProgram", CStr(Abs(.StartStandardProgram)), REG_SZ
63020    Set reg = Nothing
63030    Exit Sub
63040   End If
63050   If UCase$(OptionName) = "TOOLBARS" Then
63060    If Not reg.KeyExists Then
63070     reg.CreateKey
63080    End If
63090    reg.SetRegistryValue "Toolbars", CStr(.Toolbars), REG_SZ
63100    Set reg = Nothing
63110    Exit Sub
63120   End If
63130   If UCase$(OptionName) = "USEAUTOSAVE" Then
63140    If Not reg.KeyExists Then
63150     reg.CreateKey
63160    End If
63170    reg.SetRegistryValue "UseAutosave", CStr(Abs(.UseAutosave)), REG_SZ
63180    Set reg = Nothing
63190    Exit Sub
63200   End If
63210   If UCase$(OptionName) = "USEAUTOSAVEDIRECTORY" Then
63220    If Not reg.KeyExists Then
63230     reg.CreateKey
63240    End If
63250    reg.SetRegistryValue "UseAutosaveDirectory", CStr(Abs(.UseAutosaveDirectory)), REG_SZ
63260    Set reg = Nothing
63270    Exit Sub
63280   End If
63290  End With
63300  Set reg = Nothing
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
50490   reg.SetRegistryValue "BitmapResolution", CStr(.BitmapResolution), REG_SZ
50500   reg.SetRegistryValue "BMPColorscount", CStr(.BMPColorscount), REG_SZ
50510   reg.SetRegistryValue "JPEGColorscount", CStr(.JPEGColorscount), REG_SZ
50520   reg.SetRegistryValue "JPEGQuality", CStr(.JPEGQuality), REG_SZ
50530   reg.SetRegistryValue "PCLColorsCount", CStr(.PCLColorsCount), REG_SZ
50540   reg.SetRegistryValue "PCXColorscount", CStr(.PCXColorscount), REG_SZ
50550   reg.SetRegistryValue "PNGColorscount", CStr(.PNGColorscount), REG_SZ
50560   reg.SetRegistryValue "PSDColorsCount", CStr(.PSDColorsCount), REG_SZ
50570   reg.SetRegistryValue "RAWColorsCount", CStr(.RAWColorsCount), REG_SZ
50580   reg.SetRegistryValue "TIFFColorscount", CStr(.TIFFColorscount), REG_SZ
50590   reg.SetRegistryValue "XCFColorsCount", CStr(.XCFColorsCount), REG_SZ
50600   reg.Subkey = "Printing\Formats\PDF\Colors"
50610   If Not reg.KeyExists Then
50620    reg.CreateKey
50630   End If
50640   reg.SetRegistryValue "PDFColorsCMYKToRGB", CStr(Abs(.PDFColorsCMYKToRGB)), REG_SZ
50650   reg.SetRegistryValue "PDFColorsColorModel", CStr(.PDFColorsColorModel), REG_SZ
50660   reg.SetRegistryValue "PDFColorsPreserveHalftone", CStr(Abs(.PDFColorsPreserveHalftone)), REG_SZ
50670   reg.SetRegistryValue "PDFColorsPreserveOverprint", CStr(Abs(.PDFColorsPreserveOverprint)), REG_SZ
50680   reg.SetRegistryValue "PDFColorsPreserveTransfer", CStr(Abs(.PDFColorsPreserveTransfer)), REG_SZ
50690   reg.Subkey = "Printing\Formats\PDF\Compression"
50700   If Not reg.KeyExists Then
50710    reg.CreateKey
50720   End If
50730   reg.SetRegistryValue "PDFCompressionColorCompression", CStr(Abs(.PDFCompressionColorCompression)), REG_SZ
50740   reg.SetRegistryValue "PDFCompressionColorCompressionChoice", CStr(.PDFCompressionColorCompressionChoice), REG_SZ
50750   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
50760   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
50770   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
50780   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
50790   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
50800   reg.SetRegistryValue "PDFCompressionColorResample", CStr(Abs(.PDFCompressionColorResample)), REG_SZ
50810   reg.SetRegistryValue "PDFCompressionColorResampleChoice", CStr(.PDFCompressionColorResampleChoice), REG_SZ
50820   reg.SetRegistryValue "PDFCompressionColorResolution", CStr(.PDFCompressionColorResolution), REG_SZ
50830   reg.SetRegistryValue "PDFCompressionGreyCompression", CStr(Abs(.PDFCompressionGreyCompression)), REG_SZ
50840   reg.SetRegistryValue "PDFCompressionGreyCompressionChoice", CStr(.PDFCompressionGreyCompressionChoice), REG_SZ
50850   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
50860   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
50870   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
50880   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
50890   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
50900   reg.SetRegistryValue "PDFCompressionGreyResample", CStr(Abs(.PDFCompressionGreyResample)), REG_SZ
50910   reg.SetRegistryValue "PDFCompressionGreyResampleChoice", CStr(.PDFCompressionGreyResampleChoice), REG_SZ
50920   reg.SetRegistryValue "PDFCompressionGreyResolution", CStr(.PDFCompressionGreyResolution), REG_SZ
50930   reg.SetRegistryValue "PDFCompressionMonoCompression", CStr(Abs(.PDFCompressionMonoCompression)), REG_SZ
50940   reg.SetRegistryValue "PDFCompressionMonoCompressionChoice", CStr(.PDFCompressionMonoCompressionChoice), REG_SZ
50950   reg.SetRegistryValue "PDFCompressionMonoResample", CStr(Abs(.PDFCompressionMonoResample)), REG_SZ
50960   reg.SetRegistryValue "PDFCompressionMonoResampleChoice", CStr(.PDFCompressionMonoResampleChoice), REG_SZ
50970   reg.SetRegistryValue "PDFCompressionMonoResolution", CStr(.PDFCompressionMonoResolution), REG_SZ
50980   reg.SetRegistryValue "PDFCompressionTextCompression", CStr(Abs(.PDFCompressionTextCompression)), REG_SZ
50990   reg.Subkey = "Printing\Formats\PDF\Fonts"
51000   If Not reg.KeyExists Then
51010    reg.CreateKey
51020   End If
51030   reg.SetRegistryValue "PDFFontsEmbedAll", CStr(Abs(.PDFFontsEmbedAll)), REG_SZ
51040   reg.SetRegistryValue "PDFFontsSubSetFonts", CStr(Abs(.PDFFontsSubSetFonts)), REG_SZ
51050   reg.SetRegistryValue "PDFFontsSubSetFontsPercent", CStr(.PDFFontsSubSetFontsPercent), REG_SZ
51060   reg.Subkey = "Printing\Formats\PDF\General"
51070   If Not reg.KeyExists Then
51080    reg.CreateKey
51090   End If
51100   reg.SetRegistryValue "PDFGeneralASCII85", CStr(Abs(.PDFGeneralASCII85)), REG_SZ
51110   reg.SetRegistryValue "PDFGeneralAutorotate", CStr(.PDFGeneralAutorotate), REG_SZ
51120   reg.SetRegistryValue "PDFGeneralCompatibility", CStr(.PDFGeneralCompatibility), REG_SZ
51130   reg.SetRegistryValue "PDFGeneralDefault", CStr(.PDFGeneralDefault), REG_SZ
51140   reg.SetRegistryValue "PDFGeneralOverprint", CStr(.PDFGeneralOverprint), REG_SZ
51150   reg.SetRegistryValue "PDFGeneralResolution", CStr(.PDFGeneralResolution), REG_SZ
51160   reg.SetRegistryValue "PDFOptimize", CStr(Abs(.PDFOptimize)), REG_SZ
51170   reg.SetRegistryValue "PDFUpdateMetadata", CStr(.PDFUpdateMetadata), REG_SZ
51180   reg.Subkey = "Printing\Formats\PDF\Security"
51190   If Not reg.KeyExists Then
51200    reg.CreateKey
51210   End If
51220   reg.SetRegistryValue "PDFAllowAssembly", CStr(Abs(.PDFAllowAssembly)), REG_SZ
51230   reg.SetRegistryValue "PDFAllowDegradedPrinting", CStr(Abs(.PDFAllowDegradedPrinting)), REG_SZ
51240   reg.SetRegistryValue "PDFAllowFillIn", CStr(Abs(.PDFAllowFillIn)), REG_SZ
51250   reg.SetRegistryValue "PDFAllowScreenReaders", CStr(Abs(.PDFAllowScreenReaders)), REG_SZ
51260   reg.SetRegistryValue "PDFDisallowCopy", CStr(Abs(.PDFDisallowCopy)), REG_SZ
51270   reg.SetRegistryValue "PDFDisallowModifyAnnotations", CStr(Abs(.PDFDisallowModifyAnnotations)), REG_SZ
51280   reg.SetRegistryValue "PDFDisallowModifyContents", CStr(Abs(.PDFDisallowModifyContents)), REG_SZ
51290   reg.SetRegistryValue "PDFDisallowPrinting", CStr(Abs(.PDFDisallowPrinting)), REG_SZ
51300   reg.SetRegistryValue "PDFEncryptor", CStr(.PDFEncryptor), REG_SZ
51310   reg.SetRegistryValue "PDFHighEncryption", CStr(Abs(.PDFHighEncryption)), REG_SZ
51320   reg.SetRegistryValue "PDFLowEncryption", CStr(Abs(.PDFLowEncryption)), REG_SZ
51330   reg.SetRegistryValue "PDFOwnerPass", CStr(Abs(.PDFOwnerPass)), REG_SZ
51340   reg.SetRegistryValue "PDFOwnerPasswordString", CStr(.PDFOwnerPasswordString), REG_SZ
51350   reg.SetRegistryValue "PDFUserPass", CStr(Abs(.PDFUserPass)), REG_SZ
51360   reg.SetRegistryValue "PDFUserPasswordString", CStr(.PDFUserPasswordString), REG_SZ
51370   reg.SetRegistryValue "PDFUseSecurity", CStr(Abs(.PDFUseSecurity)), REG_SZ
51380   reg.Subkey = "Printing\Formats\PDF\Signing"
51390   If Not reg.KeyExists Then
51400    reg.CreateKey
51410   End If
51420   reg.SetRegistryValue "PDFSigningMultiSignature", CStr(Abs(.PDFSigningMultiSignature)), REG_SZ
51430   reg.SetRegistryValue "PDFSigningPFXFile", CStr(.PDFSigningPFXFile), REG_SZ
51440   reg.SetRegistryValue "PDFSigningPFXFilePassword", CStr(.PDFSigningPFXFilePassword), REG_SZ
51450   reg.SetRegistryValue "PDFSigningSignatureContact", CStr(.PDFSigningSignatureContact), REG_SZ
51460   reg.SetRegistryValue "PDFSigningSignatureLeftX", Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), REG_SZ
51470   reg.SetRegistryValue "PDFSigningSignatureLeftY", Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), REG_SZ
51480   reg.SetRegistryValue "PDFSigningSignatureLocation", CStr(.PDFSigningSignatureLocation), REG_SZ
51490   reg.SetRegistryValue "PDFSigningSignatureReason", CStr(.PDFSigningSignatureReason), REG_SZ
51500   reg.SetRegistryValue "PDFSigningSignatureRightX", Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), REG_SZ
51510   reg.SetRegistryValue "PDFSigningSignatureRightY", Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), REG_SZ
51520   reg.SetRegistryValue "PDFSigningSignatureVisible", CStr(Abs(.PDFSigningSignatureVisible)), REG_SZ
51530   reg.SetRegistryValue "PDFSigningSignPDF", CStr(Abs(.PDFSigningSignPDF)), REG_SZ
51540   reg.Subkey = "Printing\Formats\PS\LanguageLevel"
51550   If Not reg.KeyExists Then
51560    reg.CreateKey
51570   End If
51580   reg.SetRegistryValue "EPSLanguageLevel", CStr(.EPSLanguageLevel), REG_SZ
51590   reg.SetRegistryValue "PSLanguageLevel", CStr(.PSLanguageLevel), REG_SZ
51600   reg.Subkey = "Program"
51610   If Not reg.KeyExists Then
51620    reg.CreateKey
51630   End If
51640   reg.SetRegistryValue "AdditionalGhostscriptParameters", CStr(.AdditionalGhostscriptParameters), REG_SZ
51650   reg.SetRegistryValue "AdditionalGhostscriptSearchpath", CStr(.AdditionalGhostscriptSearchpath), REG_SZ
51660   reg.SetRegistryValue "AddWindowsFontpath", CStr(Abs(.AddWindowsFontpath)), REG_SZ
51670   reg.SetRegistryValue "AutosaveDirectory", CStr(.AutosaveDirectory), REG_SZ
51680   reg.SetRegistryValue "AutosaveFilename", CStr(.AutosaveFilename), REG_SZ
51690   reg.SetRegistryValue "AutosaveFormat", CStr(.AutosaveFormat), REG_SZ
51700   reg.SetRegistryValue "AutosaveStartStandardProgram", CStr(Abs(.AutosaveStartStandardProgram)), REG_SZ
51710   reg.SetRegistryValue "ClientComputerResolveIPAddress", CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
51720   reg.SetRegistryValue "DisableEmail", CStr(Abs(.DisableEmail)), REG_SZ
51730   reg.SetRegistryValue "DontUseDocumentSettings", CStr(Abs(.DontUseDocumentSettings)), REG_SZ
51740   reg.SetRegistryValue "FilenameSubstitutions", CStr(.FilenameSubstitutions), REG_SZ
51750   reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle", CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
51760   reg.SetRegistryValue "Language", CStr(.Language), REG_SZ
51770   reg.SetRegistryValue "LastSaveDirectory", CStr(.LastSaveDirectory), REG_SZ
51780   reg.SetRegistryValue "Logging", CStr(Abs(.Logging)), REG_SZ
51790   reg.SetRegistryValue "LogLines", CStr(.LogLines), REG_SZ
51800   reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter", CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
51810   reg.SetRegistryValue "NoProcessingAtStartup", CStr(Abs(.NoProcessingAtStartup)), REG_SZ
51820   reg.SetRegistryValue "NoPSCheck", CStr(Abs(.NoPSCheck)), REG_SZ
51830   reg.SetRegistryValue "OptionsDesign", CStr(.OptionsDesign), REG_SZ
51840   reg.SetRegistryValue "OptionsEnabled", CStr(Abs(.OptionsEnabled)), REG_SZ
51850   reg.SetRegistryValue "OptionsVisible", CStr(Abs(.OptionsVisible)), REG_SZ
51860   reg.SetRegistryValue "PrintAfterSaving", CStr(Abs(.PrintAfterSaving)), REG_SZ
51870   reg.SetRegistryValue "PrintAfterSavingDuplex", CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
51880   reg.SetRegistryValue "PrintAfterSavingNoCancel", CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
51890   reg.SetRegistryValue "PrintAfterSavingPrinter", CStr(.PrintAfterSavingPrinter), REG_SZ
51900   reg.SetRegistryValue "PrintAfterSavingQueryUser", CStr(.PrintAfterSavingQueryUser), REG_SZ
51910   reg.SetRegistryValue "PrintAfterSavingTumble", CStr(.PrintAfterSavingTumble), REG_SZ
51920   reg.SetRegistryValue "PrinterStop", CStr(Abs(.PrinterStop)), REG_SZ
51930   reg.SetRegistryValue "PrinterTemppath", CStr(.PrinterTemppath), REG_SZ
51940   reg.SetRegistryValue "ProcessPriority", CStr(.ProcessPriority), REG_SZ
51950   reg.SetRegistryValue "ProgramFont", CStr(.ProgramFont), REG_SZ
51960   reg.SetRegistryValue "ProgramFontCharset", CStr(.ProgramFontCharset), REG_SZ
51970   reg.SetRegistryValue "ProgramFontSize", CStr(.ProgramFontSize), REG_SZ
51980   reg.SetRegistryValue "RemoveAllKnownFileExtensions", CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
51990   reg.SetRegistryValue "RemoveSpaces", CStr(Abs(.RemoveSpaces)), REG_SZ
52000   reg.SetRegistryValue "RunProgramAfterSaving", CStr(Abs(.RunProgramAfterSaving)), REG_SZ
52010   reg.SetRegistryValue "RunProgramAfterSavingProgramname", CStr(.RunProgramAfterSavingProgramname), REG_SZ
52020   reg.SetRegistryValue "RunProgramAfterSavingProgramParameters", CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
52030   reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady", CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
52040   reg.SetRegistryValue "RunProgramAfterSavingWindowstyle", CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
52050   reg.SetRegistryValue "RunProgramBeforeSaving", CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
52060   reg.SetRegistryValue "RunProgramBeforeSavingProgramname", CStr(.RunProgramBeforeSavingProgramname), REG_SZ
52070   reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters", CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
52080   reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle", CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
52090   reg.SetRegistryValue "SaveFilename", CStr(.SaveFilename), REG_SZ
52100   reg.SetRegistryValue "SendEmailAfterAutoSaving", CStr(Abs(.SendEmailAfterAutoSaving)), REG_SZ
52110   reg.SetRegistryValue "SendMailMethod", CStr(.SendMailMethod), REG_SZ
52120   reg.SetRegistryValue "ShowAnimation", CStr(Abs(.ShowAnimation)), REG_SZ
52130   reg.SetRegistryValue "StartStandardProgram", CStr(Abs(.StartStandardProgram)), REG_SZ
52140   reg.SetRegistryValue "Toolbars", CStr(.Toolbars), REG_SZ
52150   reg.SetRegistryValue "UseAutosave", CStr(Abs(.UseAutosave)), REG_SZ
52160   reg.SetRegistryValue "UseAutosaveDirectory", CStr(Abs(.UseAutosaveDirectory)), REG_SZ
52170  End With
52180  Set reg = Nothing
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

