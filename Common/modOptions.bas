Attribute VB_Name = "modOptions"
Option Explicit

' Module automatically generated with LanguagesTool from Frank Heindörfer
' 2003
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
 PDFGeneralOverprint As Long
 PDFGeneralResolution As Long
 PDFHighEncryption As Long
 PDFLowEncryption As Long
 PDFOptimize As Long
 PDFOwnerPass As Long
 PDFOwnerPasswordString As String
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
 PSLanguageLevel As Long
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
 UseStandardAuthor As Long
End Type

Public Options As tOptions

Public Function StandardOptions() As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim myOptions As tOptions, reg As clsRegistry
50020  With myOptions
50030   .AdditionalGhostscriptParameters = vbNullString
50040   .AdditionalGhostscriptSearchpath = vbNullString
50050   .AddWindowsFontpath = "1"
50060   If InstalledAsServer Then
50070     .AutosaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
50080    Else
50090     .AutosaveDirectory = "<MyFiles>"
50100   End If
50110   .AutosaveFilename = "<DateTime>"
50120   .AutosaveFormat = "0"
50130   .AutosaveStartStandardProgram = "0"
50140   .BitmapResolution = "150"
50150   .BMPColorscount = "1"
50160   .ClientComputerResolveIPAddress = "0"
50170   .DeviceHeightPoints = Replace$("-1", ".", GetDecimalChar)
50180   .DeviceWidthPoints = Replace$("-1", ".", GetDecimalChar)
50190   Set reg = New clsRegistry
50200   reg.hkey = HKEY_LOCAL_MACHINE
50210   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50220   .DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50230   Set reg = Nothing
50240   Set reg = New clsRegistry
50250   reg.hkey = HKEY_LOCAL_MACHINE
50260   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50270   .DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50280   Set reg = Nothing
50290   Set reg = New clsRegistry
50300   reg.hkey = HKEY_LOCAL_MACHINE
50310   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50320   .DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50330   Set reg = Nothing
50340   Set reg = New clsRegistry
50350   reg.hkey = HKEY_LOCAL_MACHINE
50360   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50370   .DirectoryGhostscriptResource = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryResource"))
50380   Set reg = Nothing
50390   .DisableEmail = "0"
50400   .DontUseDocumentSettings = "0"
50410   .EPSLanguageLevel = "2"
50420   .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
50430   .FilenameSubstitutionsOnlyInTitle = "1"
50440   .JPEGColorscount = "0"
50450   .JPEGQuality = "75"
50460   .Language = "english"
50470   If InstalledAsServer Then
50480     .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
50490    Else
50500     .LastSaveDirectory = "<MyFiles>"
50510   End If
50520   .Logging = "0"
50530   .LogLines = "100"
50540   .NoConfirmMessageSwitchingDefaultprinter = "0"
50550   .NoProcessingAtStartup = "0"
50560   .NoPSCheck = "0"
50570   .OnePagePerFile = "0"
50580   .OptionsDesign = "0"
50590   .OptionsEnabled = "1"
50600   .OptionsVisible = "1"
50610   .Papersize = vbNullString
50620   .PCXColorscount = "0"
50630   .PDFAllowAssembly = "0"
50640   .PDFAllowDegradedPrinting = "0"
50650   .PDFAllowFillIn = "0"
50660   .PDFAllowScreenReaders = "0"
50670   .PDFColorsCMYKToRGB = "0"
50680   .PDFColorsColorModel = "1"
50690   .PDFColorsPreserveHalftone = "0"
50700   .PDFColorsPreserveOverprint = "1"
50710   .PDFColorsPreserveTransfer = "1"
50720   .PDFCompressionColorCompression = "1"
50730   .PDFCompressionColorCompressionChoice = "0"
50740   .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
50750   .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
50760   .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
50770   .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
50780   .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
50790   .PDFCompressionColorResample = "0"
50800   .PDFCompressionColorResampleChoice = "0"
50810   .PDFCompressionColorResolution = "300"
50820   .PDFCompressionGreyCompression = "1"
50830   .PDFCompressionGreyCompressionChoice = "0"
50840   .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
50850   .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
50860   .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
50870   .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
50880   .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
50890   .PDFCompressionGreyResample = "0"
50900   .PDFCompressionGreyResampleChoice = "0"
50910   .PDFCompressionGreyResolution = "300"
50920   .PDFCompressionMonoCompression = "1"
50930   .PDFCompressionMonoCompressionChoice = "0"
50940   .PDFCompressionMonoResample = "0"
50950   .PDFCompressionMonoResampleChoice = "0"
50960   .PDFCompressionMonoResolution = "1200"
50970   .PDFCompressionTextCompression = "1"
50980   .PDFDisallowCopy = "1"
50990   .PDFDisallowModifyAnnotations = "0"
51000   .PDFDisallowModifyContents = "0"
51010   .PDFDisallowPrinting = "0"
51020   .PDFEncryptor = "0"
51030   .PDFFontsEmbedAll = "1"
51040   .PDFFontsSubSetFonts = "1"
51050   .PDFFontsSubSetFontsPercent = "100"
51060   .PDFGeneralASCII85 = "0"
51070   .PDFGeneralAutorotate = "2"
51080   .PDFGeneralCompatibility = "1"
51090   .PDFGeneralOverprint = "0"
51100   .PDFGeneralResolution = "600"
51110   .PDFHighEncryption = "0"
51120   .PDFLowEncryption = "1"
51130   .PDFOptimize = "0"
51140   .PDFOwnerPass = "0"
51150   .PDFOwnerPasswordString = vbNullString
51160   .PDFUserPass = "0"
51170   .PDFUserPasswordString = vbNullString
51180   .PDFUseSecurity = "0"
51190   .PNGColorscount = "0"
51200   .PrintAfterSaving = "0"
51210   .PrintAfterSavingDuplex = "0"
51220   .PrintAfterSavingNoCancel = "0"
51230   .PrintAfterSavingPrinter = vbNullString
51240   .PrintAfterSavingQueryUser = "0"
51250   .PrintAfterSavingTumble = "0"
51260   .PrinterStop = "0"
51270   If InstalledAsServer Then
51280     .PrinterTemppath = CompletePath(GetPDFCreatorApplicationPath) & "Temp\"
51290    Else
51300     .PrinterTemppath = "<Temp>PDFCreator\"
51310   End If
51320   .ProcessPriority = "1"
51330   .ProgramFont = "MS Sans Serif"
51340   .ProgramFontCharset = "0"
51350   .ProgramFontSize = "8"
51360   .PSLanguageLevel = "2"
51370   .RemoveAllKnownFileExtensions = "1"
51380   .RemoveSpaces = "1"
51390   .RunProgramAfterSaving = "0"
51400   .RunProgramAfterSavingProgramname = vbNullString
51410   .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
51420   .RunProgramAfterSavingWaitUntilReady = "1"
51430   .RunProgramAfterSavingWindowstyle = "1"
51440   .RunProgramBeforeSaving = "0"
51450   .RunProgramBeforeSavingProgramname = vbNullString
51460   .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
51470   .RunProgramBeforeSavingWindowstyle = "1"
51480   .SaveFilename = "<Title>"
51490   .SendEmailAfterAutoSaving = "0"
51500   .SendMailMethod = "0"
51510   .ShowAnimation = "1"
51520   .StampFontColor = "#FF0000"
51530   .StampFontname = "Arial"
51540   .StampFontsize = "48"
51550   .StampOutlineFontthickness = "0"
51560   .StampString = vbNullString
51570   .StampUseOutlineFont = "1"
51580   .StandardAuthor = vbNullString
51590   .StandardCreationdate = vbNullString
51600   .StandardDateformat = "YYYYMMDDHHNNSS"
51610   .StandardKeywords = vbNullString
51620   .StandardMailDomain = vbNullString
51630   .StandardModifydate = vbNullString
51640   .StandardSaveformat = "0"
51650   .StandardSubject = vbNullString
51660   .StandardTitle = vbNullString
51670   .StartStandardProgram = "1"
51680   .TIFFColorscount = "0"
51690   .Toolbars = "1"
51700   .UseAutosave = "0"
51710   .UseAutosaveDirectory = "1"
51720   .UseCreationDateNow = "0"
51730   .UseStandardAuthor = "0"
51740  End With
51750  If UseINI Then
51760    If Not IsWin9xMe Then
51770     myOptions = ReadOptionsINI(myOptions, CompletePath(GetDefaultAppData) & "PDFCreator.ini", False, False)
51780    End If
51790   Else
51800    If Not IsWin9xMe Then
51810     myOptions = ReadOptionsReg(myOptions, ".DEFAULT\Software\PDFCreator", HKEY_USERS, False, False)
51820    End If
51830  End If
51840  StandardOptions = myOptions
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
50010  Dim myOptions As tOptions
50020  If InstalledAsServer Then
50030    If UseINI Then
50040      WriteToSpecialLogfile "INI-Read options: CommonAppData"
50050      myOptions = ReadOptionsINI(myOptions, CompletePath(GetCommonAppData) & "PDFCreator.ini", HKEY_LOCAL_MACHINE, NoMsg)
50060     Else
50070      WriteToSpecialLogfile "Reg-Read options: HKEY_LOCAL_MACHINE"
50080      myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", HKEY_LOCAL_MACHINE, HKEY_LOCAL_MACHINE, NoMsg)
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
50370  ReadOptions = myOptions
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
50040  ini.Filename = PDFCreatorINIFile
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
50660     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
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
51340   tStr = hOpt.Retrieve("DeviceHeightPoints")
51350   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
51360     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= -1 Then
51370       .DeviceHeightPoints = CDbl(Replace$(tStr, ".", GetDecimalChar))
51380      Else
51390       If UseStandard Then
51400        .DeviceHeightPoints = Replace$("-1", ".", GetDecimalChar)
51410       End If
51420     End If
51430    Else
51440     If UseStandard Then
51450      .DeviceHeightPoints = Replace$("-1", ".", GetDecimalChar)
51460     End If
51470   End If
51480   tStr = hOpt.Retrieve("DeviceWidthPoints")
51490   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
51500     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= -1 Then
51510       .DeviceWidthPoints = CDbl(Replace$(tStr, ".", GetDecimalChar))
51520      Else
51530       If UseStandard Then
51540        .DeviceWidthPoints = Replace$("-1", ".", GetDecimalChar)
51550       End If
51560     End If
51570    Else
51580     If UseStandard Then
51590      .DeviceWidthPoints = Replace$("-1", ".", GetDecimalChar)
51600     End If
51610   End If
51620   tStr = hOpt.Retrieve("DirectoryGhostscriptBinaries")
51630   If LenB(Trim$(tStr)) > 0 Then
51640     .DirectoryGhostscriptBinaries = CompletePath(tStr)
51650    Else
51660     If UseStandard Then
51670      tStr = GetPDFCreatorApplicationPath
51680      .DirectoryGhostscriptBinaries = CompletePath(tStr)
51690     End If
51700   End If
51710   tStr = hOpt.Retrieve("DirectoryGhostscriptFonts")
51720   If LenB(Trim$(tStr)) > 0 Then
51730     .DirectoryGhostscriptFonts = CompletePath(tStr)
51740    Else
51750     If UseStandard Then
51760      tStr = GetPDFCreatorApplicationPath & "fonts"
51770      .DirectoryGhostscriptFonts = CompletePath(tStr)
51780     End If
51790   End If
51800   tStr = hOpt.Retrieve("DirectoryGhostscriptLibraries")
51810   If LenB(Trim$(tStr)) > 0 Then
51820     .DirectoryGhostscriptLibraries = CompletePath(tStr)
51830    Else
51840     If UseStandard Then
51850      tStr = GetPDFCreatorApplicationPath & "lib"
51860      .DirectoryGhostscriptLibraries = CompletePath(tStr)
51870     End If
51880   End If
51890   tStr = hOpt.Retrieve("DirectoryGhostscriptResource")
51900   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
51910     .DirectoryGhostscriptResource = ""
51920    Else
51930     If LenB(tStr) > 0 Then
51940      .DirectoryGhostscriptResource = tStr
51950     End If
51960   End If
51970   tStr = hOpt.Retrieve("DisableEmail")
51980   If IsNumeric(tStr) Then
51990     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52000       .DisableEmail = CLng(tStr)
52010      Else
52020       If UseStandard Then
52030        .DisableEmail = 0
52040       End If
52050     End If
52060    Else
52070     If UseStandard Then
52080      .DisableEmail = 0
52090     End If
52100   End If
52110   tStr = hOpt.Retrieve("DontUseDocumentSettings")
52120   If IsNumeric(tStr) Then
52130     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52140       .DontUseDocumentSettings = CLng(tStr)
52150      Else
52160       If UseStandard Then
52170        .DontUseDocumentSettings = 0
52180       End If
52190     End If
52200    Else
52210     If UseStandard Then
52220      .DontUseDocumentSettings = 0
52230     End If
52240   End If
52250   tStr = hOpt.Retrieve("EPSLanguageLevel")
52260   If IsNumeric(tStr) Then
52270     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
52280       .EPSLanguageLevel = CLng(tStr)
52290      Else
52300       If UseStandard Then
52310        .EPSLanguageLevel = 2
52320       End If
52330     End If
52340    Else
52350     If UseStandard Then
52360      .EPSLanguageLevel = 2
52370     End If
52380   End If
52390   tStr = hOpt.Retrieve("FilenameSubstitutions")
52400   If LenB(tStr) = 0 And LenB("Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt") > 0 And UseStandard Then
52410     .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
52420    Else
52430     If LenB(tStr) > 0 Then
52440      .FilenameSubstitutions = tStr
52450     End If
52460   End If
52470   tStr = hOpt.Retrieve("FilenameSubstitutionsOnlyInTitle")
52480   If IsNumeric(tStr) Then
52490     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52500       .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
52510      Else
52520       If UseStandard Then
52530        .FilenameSubstitutionsOnlyInTitle = 1
52540       End If
52550     End If
52560    Else
52570     If UseStandard Then
52580      .FilenameSubstitutionsOnlyInTitle = 1
52590     End If
52600   End If
52610   tStr = hOpt.Retrieve("JPEGColorscount")
52620   If IsNumeric(tStr) Then
52630     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
52640       .JPEGColorscount = CLng(tStr)
52650      Else
52660       If UseStandard Then
52670        .JPEGColorscount = 0
52680       End If
52690     End If
52700    Else
52710     If UseStandard Then
52720      .JPEGColorscount = 0
52730     End If
52740   End If
52750   tStr = hOpt.Retrieve("JPEGQuality")
52760   If IsNumeric(tStr) Then
52770     If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
52780       .JPEGQuality = CLng(tStr)
52790      Else
52800       If UseStandard Then
52810        .JPEGQuality = 75
52820       End If
52830     End If
52840    Else
52850     If UseStandard Then
52860      .JPEGQuality = 75
52870     End If
52880   End If
52890   tStr = hOpt.Retrieve("Language")
52900   If LenB(tStr) = 0 And LenB("english") > 0 And UseStandard Then
52910     .Language = "english"
52920    Else
52930     If LenB(tStr) > 0 Then
52940      .Language = tStr
52950     End If
52960   End If
52970   tStr = hOpt.Retrieve("LastSaveDirectory")
52980   If LenB(Trim$(tStr)) > 0 Then
52990     .LastSaveDirectory = CompletePath(tStr)
53000    Else
53010     If UseStandard Then
53020      If InstalledAsServer Then
53030        .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
53040       Else
53050        .LastSaveDirectory = "<MyFiles>"
53060      End If
53070     End If
53080   End If
53090   tStr = hOpt.Retrieve("Logging")
53100   If IsNumeric(tStr) Then
53110     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53120       .Logging = CLng(tStr)
53130      Else
53140       If UseStandard Then
53150        .Logging = 0
53160       End If
53170     End If
53180    Else
53190     If UseStandard Then
53200      .Logging = 0
53210     End If
53220   End If
53230   tStr = hOpt.Retrieve("LogLines")
53240   If IsNumeric(tStr) Then
53250     If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
53260       .LogLines = CLng(tStr)
53270      Else
53280       If UseStandard Then
53290        .LogLines = 100
53300       End If
53310     End If
53320    Else
53330     If UseStandard Then
53340      .LogLines = 100
53350     End If
53360   End If
53370   tStr = hOpt.Retrieve("NoConfirmMessageSwitchingDefaultprinter")
53380   If IsNumeric(tStr) Then
53390     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53400       .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
53410      Else
53420       If UseStandard Then
53430        .NoConfirmMessageSwitchingDefaultprinter = 0
53440       End If
53450     End If
53460    Else
53470     If UseStandard Then
53480      .NoConfirmMessageSwitchingDefaultprinter = 0
53490     End If
53500   End If
53510   tStr = hOpt.Retrieve("NoProcessingAtStartup")
53520   If IsNumeric(tStr) Then
53530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53540       .NoProcessingAtStartup = CLng(tStr)
53550      Else
53560       If UseStandard Then
53570        .NoProcessingAtStartup = 0
53580       End If
53590     End If
53600    Else
53610     If UseStandard Then
53620      .NoProcessingAtStartup = 0
53630     End If
53640   End If
53650   tStr = hOpt.Retrieve("NoPSCheck")
53660   If IsNumeric(tStr) Then
53670     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53680       .NoPSCheck = CLng(tStr)
53690      Else
53700       If UseStandard Then
53710        .NoPSCheck = 0
53720       End If
53730     End If
53740    Else
53750     If UseStandard Then
53760      .NoPSCheck = 0
53770     End If
53780   End If
53790   tStr = hOpt.Retrieve("OnePagePerFile")
53800   If IsNumeric(tStr) Then
53810     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53820       .OnePagePerFile = CLng(tStr)
53830      Else
53840       If UseStandard Then
53850        .OnePagePerFile = 0
53860       End If
53870     End If
53880    Else
53890     If UseStandard Then
53900      .OnePagePerFile = 0
53910     End If
53920   End If
53930   tStr = hOpt.Retrieve("OptionsDesign")
53940   If IsNumeric(tStr) Then
53950     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
53960       .OptionsDesign = CLng(tStr)
53970      Else
53980       If UseStandard Then
53990        .OptionsDesign = 0
54000       End If
54010     End If
54020    Else
54030     If UseStandard Then
54040      .OptionsDesign = 0
54050     End If
54060   End If
54070   tStr = hOpt.Retrieve("OptionsEnabled")
54080   If IsNumeric(tStr) Then
54090     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54100       .OptionsEnabled = CLng(tStr)
54110      Else
54120       If UseStandard Then
54130        .OptionsEnabled = 1
54140       End If
54150     End If
54160    Else
54170     If UseStandard Then
54180      .OptionsEnabled = 1
54190     End If
54200   End If
54210   tStr = hOpt.Retrieve("OptionsVisible")
54220   If IsNumeric(tStr) Then
54230     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54240       .OptionsVisible = CLng(tStr)
54250      Else
54260       If UseStandard Then
54270        .OptionsVisible = 1
54280       End If
54290     End If
54300    Else
54310     If UseStandard Then
54320      .OptionsVisible = 1
54330     End If
54340   End If
54350   tStr = hOpt.Retrieve("Papersize")
54360   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
54370     .Papersize = ""
54380    Else
54390     If LenB(tStr) > 0 Then
54400      .Papersize = tStr
54410     End If
54420   End If
54430   tStr = hOpt.Retrieve("PCXColorscount")
54440   If IsNumeric(tStr) Then
54450     If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
54460       .PCXColorscount = CLng(tStr)
54470      Else
54480       If UseStandard Then
54490        .PCXColorscount = 0
54500       End If
54510     End If
54520    Else
54530     If UseStandard Then
54540      .PCXColorscount = 0
54550     End If
54560   End If
54570   tStr = hOpt.Retrieve("PDFAllowAssembly")
54580   If IsNumeric(tStr) Then
54590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54600       .PDFAllowAssembly = CLng(tStr)
54610      Else
54620       If UseStandard Then
54630        .PDFAllowAssembly = 0
54640       End If
54650     End If
54660    Else
54670     If UseStandard Then
54680      .PDFAllowAssembly = 0
54690     End If
54700   End If
54710   tStr = hOpt.Retrieve("PDFAllowDegradedPrinting")
54720   If IsNumeric(tStr) Then
54730     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54740       .PDFAllowDegradedPrinting = CLng(tStr)
54750      Else
54760       If UseStandard Then
54770        .PDFAllowDegradedPrinting = 0
54780       End If
54790     End If
54800    Else
54810     If UseStandard Then
54820      .PDFAllowDegradedPrinting = 0
54830     End If
54840   End If
54850   tStr = hOpt.Retrieve("PDFAllowFillIn")
54860   If IsNumeric(tStr) Then
54870     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54880       .PDFAllowFillIn = CLng(tStr)
54890      Else
54900       If UseStandard Then
54910        .PDFAllowFillIn = 0
54920       End If
54930     End If
54940    Else
54950     If UseStandard Then
54960      .PDFAllowFillIn = 0
54970     End If
54980   End If
54990   tStr = hOpt.Retrieve("PDFAllowScreenReaders")
55000   If IsNumeric(tStr) Then
55010     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55020       .PDFAllowScreenReaders = CLng(tStr)
55030      Else
55040       If UseStandard Then
55050        .PDFAllowScreenReaders = 0
55060       End If
55070     End If
55080    Else
55090     If UseStandard Then
55100      .PDFAllowScreenReaders = 0
55110     End If
55120   End If
55130   tStr = hOpt.Retrieve("PDFColorsCMYKToRGB")
55140   If IsNumeric(tStr) Then
55150     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55160       .PDFColorsCMYKToRGB = CLng(tStr)
55170      Else
55180       If UseStandard Then
55190        .PDFColorsCMYKToRGB = 0
55200       End If
55210     End If
55220    Else
55230     If UseStandard Then
55240      .PDFColorsCMYKToRGB = 0
55250     End If
55260   End If
55270   tStr = hOpt.Retrieve("PDFColorsColorModel")
55280   If IsNumeric(tStr) Then
55290     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
55300       .PDFColorsColorModel = CLng(tStr)
55310      Else
55320       If UseStandard Then
55330        .PDFColorsColorModel = 1
55340       End If
55350     End If
55360    Else
55370     If UseStandard Then
55380      .PDFColorsColorModel = 1
55390     End If
55400   End If
55410   tStr = hOpt.Retrieve("PDFColorsPreserveHalftone")
55420   If IsNumeric(tStr) Then
55430     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55440       .PDFColorsPreserveHalftone = CLng(tStr)
55450      Else
55460       If UseStandard Then
55470        .PDFColorsPreserveHalftone = 0
55480       End If
55490     End If
55500    Else
55510     If UseStandard Then
55520      .PDFColorsPreserveHalftone = 0
55530     End If
55540   End If
55550   tStr = hOpt.Retrieve("PDFColorsPreserveOverprint")
55560   If IsNumeric(tStr) Then
55570     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55580       .PDFColorsPreserveOverprint = CLng(tStr)
55590      Else
55600       If UseStandard Then
55610        .PDFColorsPreserveOverprint = 1
55620       End If
55630     End If
55640    Else
55650     If UseStandard Then
55660      .PDFColorsPreserveOverprint = 1
55670     End If
55680   End If
55690   tStr = hOpt.Retrieve("PDFColorsPreserveTransfer")
55700   If IsNumeric(tStr) Then
55710     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55720       .PDFColorsPreserveTransfer = CLng(tStr)
55730      Else
55740       If UseStandard Then
55750        .PDFColorsPreserveTransfer = 1
55760       End If
55770     End If
55780    Else
55790     If UseStandard Then
55800      .PDFColorsPreserveTransfer = 1
55810     End If
55820   End If
55830   tStr = hOpt.Retrieve("PDFCompressionColorCompression")
55840   If IsNumeric(tStr) Then
55850     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55860       .PDFCompressionColorCompression = CLng(tStr)
55870      Else
55880       If UseStandard Then
55890        .PDFCompressionColorCompression = 1
55900       End If
55910     End If
55920    Else
55930     If UseStandard Then
55940      .PDFCompressionColorCompression = 1
55950     End If
55960   End If
55970   tStr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
55980   If IsNumeric(tStr) Then
55990     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
56000       .PDFCompressionColorCompressionChoice = CLng(tStr)
56010      Else
56020       If UseStandard Then
56030        .PDFCompressionColorCompressionChoice = 0
56040       End If
56050     End If
56060    Else
56070     If UseStandard Then
56080      .PDFCompressionColorCompressionChoice = 0
56090     End If
56100   End If
56110   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGHighFactor")
56120   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56130     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56140       .PDFCompressionColorCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56150      Else
56160       If UseStandard Then
56170        .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56180       End If
56190     End If
56200    Else
56210     If UseStandard Then
56220      .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56230     End If
56240   End If
56250   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGLowFactor")
56260   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56270     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56280       .PDFCompressionColorCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56290      Else
56300       If UseStandard Then
56310        .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56320       End If
56330     End If
56340    Else
56350     If UseStandard Then
56360      .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56370     End If
56380   End If
56390   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMaximumFactor")
56400   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56410     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56420       .PDFCompressionColorCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56430      Else
56440       If UseStandard Then
56450        .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56460       End If
56470     End If
56480    Else
56490     If UseStandard Then
56500      .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56510     End If
56520   End If
56530   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMediumFactor")
56540   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56550     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56560       .PDFCompressionColorCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56570      Else
56580       If UseStandard Then
56590        .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56600       End If
56610     End If
56620    Else
56630     If UseStandard Then
56640      .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56650     End If
56660   End If
56670   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMinimumFactor")
56680   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56690     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56700       .PDFCompressionColorCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56710      Else
56720       If UseStandard Then
56730        .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
56740       End If
56750     End If
56760    Else
56770     If UseStandard Then
56780      .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
56790     End If
56800   End If
56810   tStr = hOpt.Retrieve("PDFCompressionColorResample")
56820   If IsNumeric(tStr) Then
56830     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56840       .PDFCompressionColorResample = CLng(tStr)
56850      Else
56860       If UseStandard Then
56870        .PDFCompressionColorResample = 0
56880       End If
56890     End If
56900    Else
56910     If UseStandard Then
56920      .PDFCompressionColorResample = 0
56930     End If
56940   End If
56950   tStr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
56960   If IsNumeric(tStr) Then
56970     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
56980       .PDFCompressionColorResampleChoice = CLng(tStr)
56990      Else
57000       If UseStandard Then
57010        .PDFCompressionColorResampleChoice = 0
57020       End If
57030     End If
57040    Else
57050     If UseStandard Then
57060      .PDFCompressionColorResampleChoice = 0
57070     End If
57080   End If
57090   tStr = hOpt.Retrieve("PDFCompressionColorResolution")
57100   If IsNumeric(tStr) Then
57110     If CLng(tStr) >= 0 Then
57120       .PDFCompressionColorResolution = CLng(tStr)
57130      Else
57140       If UseStandard Then
57150        .PDFCompressionColorResolution = 300
57160       End If
57170     End If
57180    Else
57190     If UseStandard Then
57200      .PDFCompressionColorResolution = 300
57210     End If
57220   End If
57230   tStr = hOpt.Retrieve("PDFCompressionGreyCompression")
57240   If IsNumeric(tStr) Then
57250     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57260       .PDFCompressionGreyCompression = CLng(tStr)
57270      Else
57280       If UseStandard Then
57290        .PDFCompressionGreyCompression = 1
57300       End If
57310     End If
57320    Else
57330     If UseStandard Then
57340      .PDFCompressionGreyCompression = 1
57350     End If
57360   End If
57370   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
57380   If IsNumeric(tStr) Then
57390     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
57400       .PDFCompressionGreyCompressionChoice = CLng(tStr)
57410      Else
57420       If UseStandard Then
57430        .PDFCompressionGreyCompressionChoice = 0
57440       End If
57450     End If
57460    Else
57470     If UseStandard Then
57480      .PDFCompressionGreyCompressionChoice = 0
57490     End If
57500   End If
57510   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGHighFactor")
57520   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57530     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57540       .PDFCompressionGreyCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57550      Else
57560       If UseStandard Then
57570        .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
57580       End If
57590     End If
57600    Else
57610     If UseStandard Then
57620      .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
57630     End If
57640   End If
57650   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGLowFactor")
57660   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57670     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57680       .PDFCompressionGreyCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57690      Else
57700       If UseStandard Then
57710        .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
57720       End If
57730     End If
57740    Else
57750     If UseStandard Then
57760      .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
57770     End If
57780   End If
57790   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMaximumFactor")
57800   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57810     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57820       .PDFCompressionGreyCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57830      Else
57840       If UseStandard Then
57850        .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57860       End If
57870     End If
57880    Else
57890     If UseStandard Then
57900      .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57910     End If
57920   End If
57930   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMediumFactor")
57940   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57950     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57960       .PDFCompressionGreyCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57970      Else
57980       If UseStandard Then
57990        .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58000       End If
58010     End If
58020    Else
58030     If UseStandard Then
58040      .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58050     End If
58060   End If
58070   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMinimumFactor")
58080   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58090     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58100       .PDFCompressionGreyCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58110      Else
58120       If UseStandard Then
58130        .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58140       End If
58150     End If
58160    Else
58170     If UseStandard Then
58180      .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58190     End If
58200   End If
58210   tStr = hOpt.Retrieve("PDFCompressionGreyResample")
58220   If IsNumeric(tStr) Then
58230     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58240       .PDFCompressionGreyResample = CLng(tStr)
58250      Else
58260       If UseStandard Then
58270        .PDFCompressionGreyResample = 0
58280       End If
58290     End If
58300    Else
58310     If UseStandard Then
58320      .PDFCompressionGreyResample = 0
58330     End If
58340   End If
58350   tStr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
58360   If IsNumeric(tStr) Then
58370     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
58380       .PDFCompressionGreyResampleChoice = CLng(tStr)
58390      Else
58400       If UseStandard Then
58410        .PDFCompressionGreyResampleChoice = 0
58420       End If
58430     End If
58440    Else
58450     If UseStandard Then
58460      .PDFCompressionGreyResampleChoice = 0
58470     End If
58480   End If
58490   tStr = hOpt.Retrieve("PDFCompressionGreyResolution")
58500   If IsNumeric(tStr) Then
58510     If CLng(tStr) >= 0 Then
58520       .PDFCompressionGreyResolution = CLng(tStr)
58530      Else
58540       If UseStandard Then
58550        .PDFCompressionGreyResolution = 300
58560       End If
58570     End If
58580    Else
58590     If UseStandard Then
58600      .PDFCompressionGreyResolution = 300
58610     End If
58620   End If
58630   tStr = hOpt.Retrieve("PDFCompressionMonoCompression")
58640   If IsNumeric(tStr) Then
58650     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58660       .PDFCompressionMonoCompression = CLng(tStr)
58670      Else
58680       If UseStandard Then
58690        .PDFCompressionMonoCompression = 1
58700       End If
58710     End If
58720    Else
58730     If UseStandard Then
58740      .PDFCompressionMonoCompression = 1
58750     End If
58760   End If
58770   tStr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
58780   If IsNumeric(tStr) Then
58790     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
58800       .PDFCompressionMonoCompressionChoice = CLng(tStr)
58810      Else
58820       If UseStandard Then
58830        .PDFCompressionMonoCompressionChoice = 0
58840       End If
58850     End If
58860    Else
58870     If UseStandard Then
58880      .PDFCompressionMonoCompressionChoice = 0
58890     End If
58900   End If
58910   tStr = hOpt.Retrieve("PDFCompressionMonoResample")
58920   If IsNumeric(tStr) Then
58930     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58940       .PDFCompressionMonoResample = CLng(tStr)
58950      Else
58960       If UseStandard Then
58970        .PDFCompressionMonoResample = 0
58980       End If
58990     End If
59000    Else
59010     If UseStandard Then
59020      .PDFCompressionMonoResample = 0
59030     End If
59040   End If
59050   tStr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
59060   If IsNumeric(tStr) Then
59070     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
59080       .PDFCompressionMonoResampleChoice = CLng(tStr)
59090      Else
59100       If UseStandard Then
59110        .PDFCompressionMonoResampleChoice = 0
59120       End If
59130     End If
59140    Else
59150     If UseStandard Then
59160      .PDFCompressionMonoResampleChoice = 0
59170     End If
59180   End If
59190   tStr = hOpt.Retrieve("PDFCompressionMonoResolution")
59200   If IsNumeric(tStr) Then
59210     If CLng(tStr) >= 0 Then
59220       .PDFCompressionMonoResolution = CLng(tStr)
59230      Else
59240       If UseStandard Then
59250        .PDFCompressionMonoResolution = 1200
59260       End If
59270     End If
59280    Else
59290     If UseStandard Then
59300      .PDFCompressionMonoResolution = 1200
59310     End If
59320   End If
59330   tStr = hOpt.Retrieve("PDFCompressionTextCompression")
59340   If IsNumeric(tStr) Then
59350     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59360       .PDFCompressionTextCompression = CLng(tStr)
59370      Else
59380       If UseStandard Then
59390        .PDFCompressionTextCompression = 1
59400       End If
59410     End If
59420    Else
59430     If UseStandard Then
59440      .PDFCompressionTextCompression = 1
59450     End If
59460   End If
59470   tStr = hOpt.Retrieve("PDFDisallowCopy")
59480   If IsNumeric(tStr) Then
59490     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59500       .PDFDisallowCopy = CLng(tStr)
59510      Else
59520       If UseStandard Then
59530        .PDFDisallowCopy = 1
59540       End If
59550     End If
59560    Else
59570     If UseStandard Then
59580      .PDFDisallowCopy = 1
59590     End If
59600   End If
59610   tStr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
59620   If IsNumeric(tStr) Then
59630     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59640       .PDFDisallowModifyAnnotations = CLng(tStr)
59650      Else
59660       If UseStandard Then
59670        .PDFDisallowModifyAnnotations = 0
59680       End If
59690     End If
59700    Else
59710     If UseStandard Then
59720      .PDFDisallowModifyAnnotations = 0
59730     End If
59740   End If
59750   tStr = hOpt.Retrieve("PDFDisallowModifyContents")
59760   If IsNumeric(tStr) Then
59770     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59780       .PDFDisallowModifyContents = CLng(tStr)
59790      Else
59800       If UseStandard Then
59810        .PDFDisallowModifyContents = 0
59820       End If
59830     End If
59840    Else
59850     If UseStandard Then
59860      .PDFDisallowModifyContents = 0
59870     End If
59880   End If
59890   tStr = hOpt.Retrieve("PDFDisallowPrinting")
59900   If IsNumeric(tStr) Then
59910     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59920       .PDFDisallowPrinting = CLng(tStr)
59930      Else
59940       If UseStandard Then
59950        .PDFDisallowPrinting = 0
59960       End If
59970     End If
59980    Else
59990     If UseStandard Then
60000      .PDFDisallowPrinting = 0
60010     End If
60020   End If
60030   tStr = hOpt.Retrieve("PDFEncryptor")
60040   If IsNumeric(tStr) Then
60050     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
60060       .PDFEncryptor = CLng(tStr)
60070      Else
60080       If UseStandard Then
60090        .PDFEncryptor = 0
60100       End If
60110     End If
60120    Else
60130     If UseStandard Then
60140      .PDFEncryptor = 0
60150     End If
60160   End If
60170   tStr = hOpt.Retrieve("PDFFontsEmbedAll")
60180   If IsNumeric(tStr) Then
60190     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60200       .PDFFontsEmbedAll = CLng(tStr)
60210      Else
60220       If UseStandard Then
60230        .PDFFontsEmbedAll = 1
60240       End If
60250     End If
60260    Else
60270     If UseStandard Then
60280      .PDFFontsEmbedAll = 1
60290     End If
60300   End If
60310   tStr = hOpt.Retrieve("PDFFontsSubSetFonts")
60320   If IsNumeric(tStr) Then
60330     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60340       .PDFFontsSubSetFonts = CLng(tStr)
60350      Else
60360       If UseStandard Then
60370        .PDFFontsSubSetFonts = 1
60380       End If
60390     End If
60400    Else
60410     If UseStandard Then
60420      .PDFFontsSubSetFonts = 1
60430     End If
60440   End If
60450   tStr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
60460   If IsNumeric(tStr) Then
60470     If CLng(tStr) >= 0 Then
60480       .PDFFontsSubSetFontsPercent = CLng(tStr)
60490      Else
60500       If UseStandard Then
60510        .PDFFontsSubSetFontsPercent = 100
60520       End If
60530     End If
60540    Else
60550     If UseStandard Then
60560      .PDFFontsSubSetFontsPercent = 100
60570     End If
60580   End If
60590   tStr = hOpt.Retrieve("PDFGeneralASCII85")
60600   If IsNumeric(tStr) Then
60610     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60620       .PDFGeneralASCII85 = CLng(tStr)
60630      Else
60640       If UseStandard Then
60650        .PDFGeneralASCII85 = 0
60660       End If
60670     End If
60680    Else
60690     If UseStandard Then
60700      .PDFGeneralASCII85 = 0
60710     End If
60720   End If
60730   tStr = hOpt.Retrieve("PDFGeneralAutorotate")
60740   If IsNumeric(tStr) Then
60750     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
60760       .PDFGeneralAutorotate = CLng(tStr)
60770      Else
60780       If UseStandard Then
60790        .PDFGeneralAutorotate = 2
60800       End If
60810     End If
60820    Else
60830     If UseStandard Then
60840      .PDFGeneralAutorotate = 2
60850     End If
60860   End If
60870   tStr = hOpt.Retrieve("PDFGeneralCompatibility")
60880   If IsNumeric(tStr) Then
60890     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
60900       .PDFGeneralCompatibility = CLng(tStr)
60910      Else
60920       If UseStandard Then
60930        .PDFGeneralCompatibility = 1
60940       End If
60950     End If
60960    Else
60970     If UseStandard Then
60980      .PDFGeneralCompatibility = 1
60990     End If
61000   End If
61010   tStr = hOpt.Retrieve("PDFGeneralOverprint")
61020   If IsNumeric(tStr) Then
61030     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
61040       .PDFGeneralOverprint = CLng(tStr)
61050      Else
61060       If UseStandard Then
61070        .PDFGeneralOverprint = 0
61080       End If
61090     End If
61100    Else
61110     If UseStandard Then
61120      .PDFGeneralOverprint = 0
61130     End If
61140   End If
61150   tStr = hOpt.Retrieve("PDFGeneralResolution")
61160   If IsNumeric(tStr) Then
61170     If CLng(tStr) >= 0 Then
61180       .PDFGeneralResolution = CLng(tStr)
61190      Else
61200       If UseStandard Then
61210        .PDFGeneralResolution = 600
61220       End If
61230     End If
61240    Else
61250     If UseStandard Then
61260      .PDFGeneralResolution = 600
61270     End If
61280   End If
61290   tStr = hOpt.Retrieve("PDFHighEncryption")
61300   If IsNumeric(tStr) Then
61310     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61320       .PDFHighEncryption = CLng(tStr)
61330      Else
61340       If UseStandard Then
61350        .PDFHighEncryption = 0
61360       End If
61370     End If
61380    Else
61390     If UseStandard Then
61400      .PDFHighEncryption = 0
61410     End If
61420   End If
61430   tStr = hOpt.Retrieve("PDFLowEncryption")
61440   If IsNumeric(tStr) Then
61450     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61460       .PDFLowEncryption = CLng(tStr)
61470      Else
61480       If UseStandard Then
61490        .PDFLowEncryption = 1
61500       End If
61510     End If
61520    Else
61530     If UseStandard Then
61540      .PDFLowEncryption = 1
61550     End If
61560   End If
61570   tStr = hOpt.Retrieve("PDFOptimize")
61580   If IsNumeric(tStr) Then
61590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61600       .PDFOptimize = CLng(tStr)
61610      Else
61620       If UseStandard Then
61630        .PDFOptimize = 0
61640       End If
61650     End If
61660    Else
61670     If UseStandard Then
61680      .PDFOptimize = 0
61690     End If
61700   End If
61710   tStr = hOpt.Retrieve("PDFOwnerPass")
61720   If IsNumeric(tStr) Then
61730     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61740       .PDFOwnerPass = CLng(tStr)
61750      Else
61760       If UseStandard Then
61770        .PDFOwnerPass = 0
61780       End If
61790     End If
61800    Else
61810     If UseStandard Then
61820      .PDFOwnerPass = 0
61830     End If
61840   End If
61850   tStr = hOpt.Retrieve("PDFOwnerPasswordString")
61860   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
61870     .PDFOwnerPasswordString = ""
61880    Else
61890     If LenB(tStr) > 0 Then
61900      .PDFOwnerPasswordString = tStr
61910     End If
61920   End If
61930   tStr = hOpt.Retrieve("PDFUserPass")
61940   If IsNumeric(tStr) Then
61950     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61960       .PDFUserPass = CLng(tStr)
61970      Else
61980       If UseStandard Then
61990        .PDFUserPass = 0
62000       End If
62010     End If
62020    Else
62030     If UseStandard Then
62040      .PDFUserPass = 0
62050     End If
62060   End If
62070   tStr = hOpt.Retrieve("PDFUserPasswordString")
62080   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62090     .PDFUserPasswordString = ""
62100    Else
62110     If LenB(tStr) > 0 Then
62120      .PDFUserPasswordString = tStr
62130     End If
62140   End If
62150   tStr = hOpt.Retrieve("PDFUseSecurity")
62160   If IsNumeric(tStr) Then
62170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62180       .PDFUseSecurity = CLng(tStr)
62190      Else
62200       If UseStandard Then
62210        .PDFUseSecurity = 0
62220       End If
62230     End If
62240    Else
62250     If UseStandard Then
62260      .PDFUseSecurity = 0
62270     End If
62280   End If
62290   tStr = hOpt.Retrieve("PNGColorscount")
62300   If IsNumeric(tStr) Then
62310     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
62320       .PNGColorscount = CLng(tStr)
62330      Else
62340       If UseStandard Then
62350        .PNGColorscount = 0
62360       End If
62370     End If
62380    Else
62390     If UseStandard Then
62400      .PNGColorscount = 0
62410     End If
62420   End If
62430   tStr = hOpt.Retrieve("PrintAfterSaving")
62440   If IsNumeric(tStr) Then
62450     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62460       .PrintAfterSaving = CLng(tStr)
62470      Else
62480       If UseStandard Then
62490        .PrintAfterSaving = 0
62500       End If
62510     End If
62520    Else
62530     If UseStandard Then
62540      .PrintAfterSaving = 0
62550     End If
62560   End If
62570   tStr = hOpt.Retrieve("PrintAfterSavingDuplex")
62580   If IsNumeric(tStr) Then
62590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62600       .PrintAfterSavingDuplex = CLng(tStr)
62610      Else
62620       If UseStandard Then
62630        .PrintAfterSavingDuplex = 0
62640       End If
62650     End If
62660    Else
62670     If UseStandard Then
62680      .PrintAfterSavingDuplex = 0
62690     End If
62700   End If
62710   tStr = hOpt.Retrieve("PrintAfterSavingNoCancel")
62720   If IsNumeric(tStr) Then
62730     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62740       .PrintAfterSavingNoCancel = CLng(tStr)
62750      Else
62760       If UseStandard Then
62770        .PrintAfterSavingNoCancel = 0
62780       End If
62790     End If
62800    Else
62810     If UseStandard Then
62820      .PrintAfterSavingNoCancel = 0
62830     End If
62840   End If
62850   tStr = hOpt.Retrieve("PrintAfterSavingPrinter")
62860   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62870     .PrintAfterSavingPrinter = ""
62880    Else
62890     If LenB(tStr) > 0 Then
62900      .PrintAfterSavingPrinter = tStr
62910     End If
62920   End If
62930   tStr = hOpt.Retrieve("PrintAfterSavingQueryUser")
62940   If IsNumeric(tStr) Then
62950     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
62960       .PrintAfterSavingQueryUser = CLng(tStr)
62970      Else
62980       If UseStandard Then
62990        .PrintAfterSavingQueryUser = 0
63000       End If
63010     End If
63020    Else
63030     If UseStandard Then
63040      .PrintAfterSavingQueryUser = 0
63050     End If
63060   End If
63070   tStr = hOpt.Retrieve("PrintAfterSavingTumble")
63080   If IsNumeric(tStr) Then
63090     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
63100       .PrintAfterSavingTumble = CLng(tStr)
63110      Else
63120       If UseStandard Then
63130        .PrintAfterSavingTumble = 0
63140       End If
63150     End If
63160    Else
63170     If UseStandard Then
63180      .PrintAfterSavingTumble = 0
63190     End If
63200   End If
63210   tStr = hOpt.Retrieve("PrinterStop")
63220   If IsNumeric(tStr) Then
63230     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63240       .PrinterStop = CLng(tStr)
63250      Else
63260       If UseStandard Then
63270        .PrinterStop = 0
63280       End If
63290     End If
63300    Else
63310     If UseStandard Then
63320      .PrinterStop = 0
63330     End If
63340   End If
63350   tStr = hOpt.Retrieve("PrinterTemppath")
63360   WriteToSpecialLogfile "hOpt.Retrieve(""PrinterTemppath"")=" & tStr
63370   WriteToSpecialLogfile "Options.PrinterTemppath1=" & .PrinterTemppath
63380   If hkey1 = HKEY_USERS Then
63390     If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then
63400       .PrinterTemppath = tStr
63410      Else
63420       If UseStandard Then
63430         .PrinterTemppath = GetTempPath
63440        Else
63450         .PrinterTemppath = tStr
63460       End If
63470     End If
63480    Else
63490     If LenB(Trim$(tStr)) > 0 Then
63500      If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then
63510        .PrinterTemppath = tStr
63520       Else
63530        MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))
63540        If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then
63550          If UseStandard Then
63560            .PrinterTemppath = GetTempPath
63570           Else
63580            .PrinterTemppath = ""
63590            If NoMsg = False Then
63600             MsgBox "PrinterTemppath: '" & tStr & "' = '" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & "'" & _
             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
63620            End If
63630          End If
63640         Else
63650          .PrinterTemppath = tStr
63660        End If
63670      End If
63680     End If
63690   End If
63700   WriteToSpecialLogfile "Options.PrinterTemppath2=" & .PrinterTemppath
63710   tStr = hOpt.Retrieve("ProcessPriority")
63720   If IsNumeric(tStr) Then
63730     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
63740       .ProcessPriority = CLng(tStr)
63750      Else
63760       If UseStandard Then
63770        .ProcessPriority = 1
63780       End If
63790     End If
63800    Else
63810     If UseStandard Then
63820      .ProcessPriority = 1
63830     End If
63840   End If
63850   tStr = hOpt.Retrieve("ProgramFont")
63860   If LenB(tStr) = 0 And LenB("MS Sans Serif") > 0 And UseStandard Then
63870     .ProgramFont = "MS Sans Serif"
63880    Else
63890     If LenB(tStr) > 0 Then
63900      .ProgramFont = tStr
63910     End If
63920   End If
63930   tStr = hOpt.Retrieve("ProgramFontCharset")
63940   If IsNumeric(tStr) Then
63950     If CLng(tStr) >= 0 Then
63960       .ProgramFontCharset = CLng(tStr)
63970      Else
63980       If UseStandard Then
63990        .ProgramFontCharset = 0
64000       End If
64010     End If
64020    Else
64030     If UseStandard Then
64040      .ProgramFontCharset = 0
64050     End If
64060   End If
64070   tStr = hOpt.Retrieve("ProgramFontSize")
64080   If IsNumeric(tStr) Then
64090     If CLng(tStr) >= 1 And CLng(tStr) <= 72 Then
64100       .ProgramFontSize = CLng(tStr)
64110      Else
64120       If UseStandard Then
64130        .ProgramFontSize = 8
64140       End If
64150     End If
64160    Else
64170     If UseStandard Then
64180      .ProgramFontSize = 8
64190     End If
64200   End If
64210   tStr = hOpt.Retrieve("PSLanguageLevel")
64220   If IsNumeric(tStr) Then
64230     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
64240       .PSLanguageLevel = CLng(tStr)
64250      Else
64260       If UseStandard Then
64270        .PSLanguageLevel = 2
64280       End If
64290     End If
64300    Else
64310     If UseStandard Then
64320      .PSLanguageLevel = 2
64330     End If
64340   End If
64350   tStr = hOpt.Retrieve("RemoveAllKnownFileExtensions")
64360   If IsNumeric(tStr) Then
64370     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64380       .RemoveAllKnownFileExtensions = CLng(tStr)
64390      Else
64400       If UseStandard Then
64410        .RemoveAllKnownFileExtensions = 1
64420       End If
64430     End If
64440    Else
64450     If UseStandard Then
64460      .RemoveAllKnownFileExtensions = 1
64470     End If
64480   End If
64490   tStr = hOpt.Retrieve("RemoveSpaces")
64500   If IsNumeric(tStr) Then
64510     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64520       .RemoveSpaces = CLng(tStr)
64530      Else
64540       If UseStandard Then
64550        .RemoveSpaces = 1
64560       End If
64570     End If
64580    Else
64590     If UseStandard Then
64600      .RemoveSpaces = 1
64610     End If
64620   End If
64630   tStr = hOpt.Retrieve("RunProgramAfterSaving")
64640   If IsNumeric(tStr) Then
64650     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64660       .RunProgramAfterSaving = CLng(tStr)
64670      Else
64680       If UseStandard Then
64690        .RunProgramAfterSaving = 0
64700       End If
64710     End If
64720    Else
64730     If UseStandard Then
64740      .RunProgramAfterSaving = 0
64750     End If
64760   End If
64770   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramname")
64780   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64790     .RunProgramAfterSavingProgramname = ""
64800    Else
64810     If LenB(tStr) > 0 Then
64820      .RunProgramAfterSavingProgramname = tStr
64830     End If
64840   End If
64850   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramParameters")
64860   If LenB(tStr) = 0 And LenB("""<OutputFilename>""") > 0 And UseStandard Then
64870     .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
64880    Else
64890     If LenB(tStr) > 0 Then
64900      .RunProgramAfterSavingProgramParameters = tStr
64910     End If
64920   End If
64930   tStr = hOpt.Retrieve("RunProgramAfterSavingWaitUntilReady")
64940   If IsNumeric(tStr) Then
64950     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64960       .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
64970      Else
64980       If UseStandard Then
64990        .RunProgramAfterSavingWaitUntilReady = 1
65000       End If
65010     End If
65020    Else
65030     If UseStandard Then
65040      .RunProgramAfterSavingWaitUntilReady = 1
65050     End If
65060   End If
65070   tStr = hOpt.Retrieve("RunProgramAfterSavingWindowstyle")
65080   If IsNumeric(tStr) Then
65090     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
65100       .RunProgramAfterSavingWindowstyle = CLng(tStr)
65110      Else
65120       If UseStandard Then
65130        .RunProgramAfterSavingWindowstyle = 1
65140       End If
65150     End If
65160    Else
65170     If UseStandard Then
65180      .RunProgramAfterSavingWindowstyle = 1
65190     End If
65200   End If
65210   tStr = hOpt.Retrieve("RunProgramBeforeSaving")
65220   If IsNumeric(tStr) Then
65230     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65240       .RunProgramBeforeSaving = CLng(tStr)
65250      Else
65260       If UseStandard Then
65270        .RunProgramBeforeSaving = 0
65280       End If
65290     End If
65300    Else
65310     If UseStandard Then
65320      .RunProgramBeforeSaving = 0
65330     End If
65340   End If
65350   tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramname")
65360   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
65370     .RunProgramBeforeSavingProgramname = ""
65380    Else
65390     If LenB(tStr) > 0 Then
65400      .RunProgramBeforeSavingProgramname = tStr
65410     End If
65420   End If
65430   tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramParameters")
65440   If LenB(tStr) = 0 And LenB("""<TempFilename>""") > 0 And UseStandard Then
65450     .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
65460    Else
65470     If LenB(tStr) > 0 Then
65480      .RunProgramBeforeSavingProgramParameters = tStr
65490     End If
65500   End If
65510   tStr = hOpt.Retrieve("RunProgramBeforeSavingWindowstyle")
65520   If IsNumeric(tStr) Then
65530     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
65540       .RunProgramBeforeSavingWindowstyle = CLng(tStr)
65550      Else
65560       If UseStandard Then
65570        .RunProgramBeforeSavingWindowstyle = 1
65580       End If
65590     End If
65600    Else
65610     If UseStandard Then
65620      .RunProgramBeforeSavingWindowstyle = 1
65630     End If
65640   End If
65650   tStr = hOpt.Retrieve("SaveFilename")
65660   If LenB(tStr) = 0 And LenB("<Title>") > 0 And UseStandard Then
65670     .SaveFilename = "<Title>"
65680    Else
65690     If LenB(tStr) > 0 Then
65700      .SaveFilename = tStr
65710     End If
65720   End If
65730   tStr = hOpt.Retrieve("SendEmailAfterAutoSaving")
65740   If IsNumeric(tStr) Then
65750     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65760       .SendEmailAfterAutoSaving = CLng(tStr)
65770      Else
65780       If UseStandard Then
65790        .SendEmailAfterAutoSaving = 0
65800       End If
65810     End If
65820    Else
65830     If UseStandard Then
65840      .SendEmailAfterAutoSaving = 0
65850     End If
65860   End If
65870   tStr = hOpt.Retrieve("SendMailMethod")
65880   If IsNumeric(tStr) Then
65890     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
65900       .SendMailMethod = CLng(tStr)
65910      Else
65920       If UseStandard Then
65930        .SendMailMethod = 0
65940       End If
65950     End If
65960    Else
65970     If UseStandard Then
65980      .SendMailMethod = 0
65990     End If
66000   End If
66010   tStr = hOpt.Retrieve("ShowAnimation")
66020   If IsNumeric(tStr) Then
66030     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66040       .ShowAnimation = CLng(tStr)
66050      Else
66060       If UseStandard Then
66070        .ShowAnimation = 1
66080       End If
66090     End If
66100    Else
66110     If UseStandard Then
66120      .ShowAnimation = 1
66130     End If
66140   End If
66150   tStr = hOpt.Retrieve("StampFontColor")
66160   If LenB(tStr) = 0 And LenB("#FF0000") > 0 And UseStandard Then
66170     .StampFontColor = "#FF0000"
66180    Else
66190     If LenB(tStr) > 0 Then
66200      .StampFontColor = tStr
66210     End If
66220   End If
66230   tStr = hOpt.Retrieve("StampFontname")
66240   If LenB(tStr) = 0 And LenB("Arial") > 0 And UseStandard Then
66250     .StampFontname = "Arial"
66260    Else
66270     If LenB(tStr) > 0 Then
66280      .StampFontname = tStr
66290     End If
66300   End If
66310   tStr = hOpt.Retrieve("StampFontsize")
66320   If IsNumeric(tStr) Then
66330     If CLng(tStr) >= 1 Then
66340       .StampFontsize = CLng(tStr)
66350      Else
66360       If UseStandard Then
66370        .StampFontsize = 48
66380       End If
66390     End If
66400    Else
66410     If UseStandard Then
66420      .StampFontsize = 48
66430     End If
66440   End If
66450   tStr = hOpt.Retrieve("StampOutlineFontthickness")
66460   If IsNumeric(tStr) Then
66470     If CLng(tStr) >= 0 Then
66480       .StampOutlineFontthickness = CLng(tStr)
66490      Else
66500       If UseStandard Then
66510        .StampOutlineFontthickness = 0
66520       End If
66530     End If
66540    Else
66550     If UseStandard Then
66560      .StampOutlineFontthickness = 0
66570     End If
66580   End If
66590   tStr = hOpt.Retrieve("StampString")
66600   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
66610     .StampString = ""
66620    Else
66630     If LenB(tStr) > 0 Then
66640      .StampString = tStr
66650     End If
66660   End If
66670   tStr = hOpt.Retrieve("StampUseOutlineFont")
66680   If IsNumeric(tStr) Then
66690     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66700       .StampUseOutlineFont = CLng(tStr)
66710      Else
66720       If UseStandard Then
66730        .StampUseOutlineFont = 1
66740       End If
66750     End If
66760    Else
66770     If UseStandard Then
66780      .StampUseOutlineFont = 1
66790     End If
66800   End If
66810   tStr = hOpt.Retrieve("StandardAuthor")
66820   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
66830     .StandardAuthor = ""
66840    Else
66850     If LenB(tStr) > 0 Then
66860      .StandardAuthor = tStr
66870     End If
66880   End If
66890   tStr = hOpt.Retrieve("StandardCreationdate")
66900   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
66910     .StandardCreationdate = ""
66920    Else
66930     If LenB(tStr) > 0 Then
66940      .StandardCreationdate = tStr
66950     End If
66960   End If
66970   tStr = hOpt.Retrieve("StandardDateformat")
66980   If LenB(tStr) = 0 And LenB("YYYYMMDDHHNNSS") > 0 And UseStandard Then
66990     .StandardDateformat = "YYYYMMDDHHNNSS"
67000    Else
67010     If LenB(tStr) > 0 Then
67020      .StandardDateformat = tStr
67030     End If
67040   End If
67050   tStr = hOpt.Retrieve("StandardKeywords")
67060   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67070     .StandardKeywords = ""
67080    Else
67090     If LenB(tStr) > 0 Then
67100      .StandardKeywords = tStr
67110     End If
67120   End If
67130   tStr = hOpt.Retrieve("StandardMailDomain")
67140   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67150     .StandardMailDomain = ""
67160    Else
67170     If LenB(tStr) > 0 Then
67180      .StandardMailDomain = tStr
67190     End If
67200   End If
67210   tStr = hOpt.Retrieve("StandardModifydate")
67220   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67230     .StandardModifydate = ""
67240    Else
67250     If LenB(tStr) > 0 Then
67260      .StandardModifydate = tStr
67270     End If
67280   End If
67290   tStr = hOpt.Retrieve("StandardSaveformat")
67300   If IsNumeric(tStr) Then
67310     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
67320       .StandardSaveformat = CLng(tStr)
67330      Else
67340       If UseStandard Then
67350        .StandardSaveformat = 0
67360       End If
67370     End If
67380    Else
67390     If UseStandard Then
67400      .StandardSaveformat = 0
67410     End If
67420   End If
67430   tStr = hOpt.Retrieve("StandardSubject")
67440   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67450     .StandardSubject = ""
67460    Else
67470     If LenB(tStr) > 0 Then
67480      .StandardSubject = tStr
67490     End If
67500   End If
67510   tStr = hOpt.Retrieve("StandardTitle")
67520   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67530     .StandardTitle = ""
67540    Else
67550     If LenB(tStr) > 0 Then
67560      .StandardTitle = tStr
67570     End If
67580   End If
67590   tStr = hOpt.Retrieve("StartStandardProgram")
67600   If IsNumeric(tStr) Then
67610     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67620       .StartStandardProgram = CLng(tStr)
67630      Else
67640       If UseStandard Then
67650        .StartStandardProgram = 1
67660       End If
67670     End If
67680    Else
67690     If UseStandard Then
67700      .StartStandardProgram = 1
67710     End If
67720   End If
67730   tStr = hOpt.Retrieve("TIFFColorscount")
67740   If IsNumeric(tStr) Then
67750     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
67760       .TIFFColorscount = CLng(tStr)
67770      Else
67780       If UseStandard Then
67790        .TIFFColorscount = 0
67800       End If
67810     End If
67820    Else
67830     If UseStandard Then
67840      .TIFFColorscount = 0
67850     End If
67860   End If
67870   tStr = hOpt.Retrieve("Toolbars")
67880   If IsNumeric(tStr) Then
67890     If CLng(tStr) >= 0 Then
67900       .Toolbars = CLng(tStr)
67910      Else
67920       If UseStandard Then
67930        .Toolbars = 1
67940       End If
67950     End If
67960    Else
67970     If UseStandard Then
67980      .Toolbars = 1
67990     End If
68000   End If
68010   tStr = hOpt.Retrieve("UseAutosave")
68020   If IsNumeric(tStr) Then
68030     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68040       .UseAutosave = CLng(tStr)
68050      Else
68060       If UseStandard Then
68070        .UseAutosave = 0
68080       End If
68090     End If
68100    Else
68110     If UseStandard Then
68120      .UseAutosave = 0
68130     End If
68140   End If
68150   tStr = hOpt.Retrieve("UseAutosaveDirectory")
68160   If IsNumeric(tStr) Then
68170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68180       .UseAutosaveDirectory = CLng(tStr)
68190      Else
68200       If UseStandard Then
68210        .UseAutosaveDirectory = 1
68220       End If
68230     End If
68240    Else
68250     If UseStandard Then
68260      .UseAutosaveDirectory = 1
68270     End If
68280   End If
68290   tStr = hOpt.Retrieve("UseCreationDateNow")
68300   If IsNumeric(tStr) Then
68310     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68320       .UseCreationDateNow = CLng(tStr)
68330      Else
68340       If UseStandard Then
68350        .UseCreationDateNow = 0
68360       End If
68370     End If
68380    Else
68390     If UseStandard Then
68400      .UseCreationDateNow = 0
68410     End If
68420   End If
68430   tStr = hOpt.Retrieve("UseStandardAuthor")
68440   If IsNumeric(tStr) Then
68450     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68460       .UseStandardAuthor = CLng(tStr)
68470      Else
68480       If UseStandard Then
68490        .UseStandardAuthor = 0
68500       End If
68510     End If
68520    Else
68530     If UseStandard Then
68540      .UseStandardAuthor = 0
68550     End If
68560   End If
68570  End With
68580  Set ini = Nothing
68590  ReadOptionsINI = myOptions
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

Public Sub CorrectOptions()
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
Select Case ErrPtnr.OnError("modOptions", "CorrectOptions")
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
50010  CorrectOptions
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
50030  ini.Filename = PDFCreatorINIFile
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
50200   Case "DEVICEHEIGHTPOINTS": ini.SaveKey Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), "DeviceHeightPoints"
50210   Case "DEVICEWIDTHPOINTS": ini.SaveKey Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), "DeviceWidthPoints"
50220   Case "DIRECTORYGHOSTSCRIPTBINARIES": ini.SaveKey CStr(.DirectoryGhostscriptBinaries), "DirectoryGhostscriptBinaries"
50230   Case "DIRECTORYGHOSTSCRIPTFONTS": ini.SaveKey CStr(.DirectoryGhostscriptFonts), "DirectoryGhostscriptFonts"
50240   Case "DIRECTORYGHOSTSCRIPTLIBRARIES": ini.SaveKey CStr(.DirectoryGhostscriptLibraries), "DirectoryGhostscriptLibraries"
50250   Case "DIRECTORYGHOSTSCRIPTRESOURCE": ini.SaveKey CStr(.DirectoryGhostscriptResource), "DirectoryGhostscriptResource"
50260   Case "DISABLEEMAIL": ini.SaveKey CStr(Abs(.DisableEmail)), "DisableEmail"
50270   Case "DONTUSEDOCUMENTSETTINGS": ini.SaveKey CStr(Abs(.DontUseDocumentSettings)), "DontUseDocumentSettings"
50280   Case "EPSLANGUAGELEVEL": ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
50290   Case "FILENAMESUBSTITUTIONS": ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
50300   Case "FILENAMESUBSTITUTIONSONLYINTITLE": ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
50310   Case "JPEGCOLORSCOUNT": ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50320   Case "JPEGQUALITY": ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50330   Case "LANGUAGE": ini.SaveKey CStr(.Language), "Language"
50340   Case "LASTSAVEDIRECTORY": ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50350   Case "LOGGING": ini.SaveKey CStr(Abs(.Logging)), "Logging"
50360   Case "LOGLINES": ini.SaveKey CStr(.LogLines), "LogLines"
50370   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50380   Case "NOPROCESSINGATSTARTUP": ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
50390   Case "NOPSCHECK": ini.SaveKey CStr(Abs(.NoPSCheck)), "NoPSCheck"
50400   Case "ONEPAGEPERFILE": ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50410   Case "OPTIONSDESIGN": ini.SaveKey CStr(.OptionsDesign), "OptionsDesign"
50420   Case "OPTIONSENABLED": ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50430   Case "OPTIONSVISIBLE": ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50440   Case "PAPERSIZE": ini.SaveKey CStr(.Papersize), "Papersize"
50450   Case "PCXCOLORSCOUNT": ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50460   Case "PDFALLOWASSEMBLY": ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50470   Case "PDFALLOWDEGRADEDPRINTING": ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50480   Case "PDFALLOWFILLIN": ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50490   Case "PDFALLOWSCREENREADERS": ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50500   Case "PDFCOLORSCMYKTORGB": ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50510   Case "PDFCOLORSCOLORMODEL": ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50520   Case "PDFCOLORSPRESERVEHALFTONE": ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50530   Case "PDFCOLORSPRESERVEOVERPRINT": ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50540   Case "PDFCOLORSPRESERVETRANSFER": ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50550   Case "PDFCOMPRESSIONCOLORCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50560   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50570   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGHighFactor"
50580   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGLowFactor"
50590   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMaximumFactor"
50600   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMediumFactor"
50610   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMinimumFactor"
50620   Case "PDFCOMPRESSIONCOLORRESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50630   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50640   Case "PDFCOMPRESSIONCOLORRESOLUTION": ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50650   Case "PDFCOMPRESSIONGREYCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50660   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50670   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGHighFactor"
50680   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGLowFactor"
50690   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMaximumFactor"
50700   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMediumFactor"
50710   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMinimumFactor"
50720   Case "PDFCOMPRESSIONGREYRESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50730   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50740   Case "PDFCOMPRESSIONGREYRESOLUTION": ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50750   Case "PDFCOMPRESSIONMONOCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50760   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50770   Case "PDFCOMPRESSIONMONORESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50780   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50790   Case "PDFCOMPRESSIONMONORESOLUTION": ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50800   Case "PDFCOMPRESSIONTEXTCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50810   Case "PDFDISALLOWCOPY": ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50820   Case "PDFDISALLOWMODIFYANNOTATIONS": ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50830   Case "PDFDISALLOWMODIFYCONTENTS": ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50840   Case "PDFDISALLOWPRINTING": ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50850   Case "PDFENCRYPTOR": ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50860   Case "PDFFONTSEMBEDALL": ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50870   Case "PDFFONTSSUBSETFONTS": ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50880   Case "PDFFONTSSUBSETFONTSPERCENT": ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50890   Case "PDFGENERALASCII85": ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50900   Case "PDFGENERALAUTOROTATE": ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50910   Case "PDFGENERALCOMPATIBILITY": ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50920   Case "PDFGENERALOVERPRINT": ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50930   Case "PDFGENERALRESOLUTION": ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
50940   Case "PDFHIGHENCRYPTION": ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
50950   Case "PDFLOWENCRYPTION": ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
50960   Case "PDFOPTIMIZE": ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
50970   Case "PDFOWNERPASS": ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
50980   Case "PDFOWNERPASSWORDSTRING": ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
50990   Case "PDFUSERPASS": ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
51000   Case "PDFUSERPASSWORDSTRING": ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
51010   Case "PDFUSESECURITY": ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
51020   Case "PNGCOLORSCOUNT": ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
51030   Case "PRINTAFTERSAVING": ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
51040   Case "PRINTAFTERSAVINGDUPLEX": ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
51050   Case "PRINTAFTERSAVINGNOCANCEL": ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
51060   Case "PRINTAFTERSAVINGPRINTER": ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
51070   Case "PRINTAFTERSAVINGQUERYUSER": ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
51080   Case "PRINTAFTERSAVINGTUMBLE": ini.SaveKey CStr(.PrintAfterSavingTumble), "PrintAfterSavingTumble"
51090   Case "PRINTERSTOP": ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
51100   Case "PRINTERTEMPPATH": ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
51110   Case "PROCESSPRIORITY": ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
51120   Case "PROGRAMFONT": ini.SaveKey CStr(.ProgramFont), "ProgramFont"
51130   Case "PROGRAMFONTCHARSET": ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
51140   Case "PROGRAMFONTSIZE": ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
51150   Case "PSLANGUAGELEVEL": ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
51160   Case "REMOVEALLKNOWNFILEEXTENSIONS": ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
51170   Case "REMOVESPACES": ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
51180   Case "RUNPROGRAMAFTERSAVING": ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
51190   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
51200   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
51210   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51220   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51230   Case "RUNPROGRAMBEFORESAVING": ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51240   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51250   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51260   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51270   Case "SAVEFILENAME": ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51280   Case "SENDEMAILAFTERAUTOSAVING": ini.SaveKey CStr(Abs(.SendEmailAfterAutoSaving)), "SendEmailAfterAutoSaving"
51290   Case "SENDMAILMETHOD": ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51300   Case "SHOWANIMATION": ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51310   Case "STAMPFONTCOLOR": ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51320   Case "STAMPFONTNAME": ini.SaveKey CStr(.StampFontname), "StampFontname"
51330   Case "STAMPFONTSIZE": ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51340   Case "STAMPOUTLINEFONTTHICKNESS": ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51350   Case "STAMPSTRING": ini.SaveKey CStr(.StampString), "StampString"
51360   Case "STAMPUSEOUTLINEFONT": ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51370   Case "STANDARDAUTHOR": ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51380   Case "STANDARDCREATIONDATE": ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51390   Case "STANDARDDATEFORMAT": ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51400   Case "STANDARDKEYWORDS": ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51410   Case "STANDARDMAILDOMAIN": ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51420   Case "STANDARDMODIFYDATE": ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51430   Case "STANDARDSAVEFORMAT": ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51440   Case "STANDARDSUBJECT": ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51450   Case "STANDARDTITLE": ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51460   Case "STARTSTANDARDPROGRAM": ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51470   Case "TIFFCOLORSCOUNT": ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51480   Case "TOOLBARS": ini.SaveKey CStr(.Toolbars), "Toolbars"
51490   Case "USEAUTOSAVE": ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51500   Case "USEAUTOSAVEDIRECTORY": ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51510   Case "USECREATIONDATENOW": ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51520   Case "USESTANDARDAUTHOR": ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51530   End Select
51540  End With
51550  Set ini = Nothing
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
50030  ini.Filename = PDFCreatorINIFile
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
50190   ini.SaveKey Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), "DeviceHeightPoints"
50200   ini.SaveKey Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), "DeviceWidthPoints"
50210   ini.SaveKey CStr(.DirectoryGhostscriptBinaries), "DirectoryGhostscriptBinaries"
50220   ini.SaveKey CStr(.DirectoryGhostscriptFonts), "DirectoryGhostscriptFonts"
50230   ini.SaveKey CStr(.DirectoryGhostscriptLibraries), "DirectoryGhostscriptLibraries"
50240   ini.SaveKey CStr(.DirectoryGhostscriptResource), "DirectoryGhostscriptResource"
50250   ini.SaveKey CStr(Abs(.DisableEmail)), "DisableEmail"
50260   ini.SaveKey CStr(Abs(.DontUseDocumentSettings)), "DontUseDocumentSettings"
50270   ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
50280   ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
50290   ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
50300   ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50310   ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50320   ini.SaveKey CStr(.Language), "Language"
50330   ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50340   ini.SaveKey CStr(Abs(.Logging)), "Logging"
50350   ini.SaveKey CStr(.LogLines), "LogLines"
50360   ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50370   ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
50380   ini.SaveKey CStr(Abs(.NoPSCheck)), "NoPSCheck"
50390   ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50400   ini.SaveKey CStr(.OptionsDesign), "OptionsDesign"
50410   ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50420   ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50430   ini.SaveKey CStr(.Papersize), "Papersize"
50440   ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50450   ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50460   ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50470   ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50480   ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50490   ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50500   ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50510   ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50520   ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50530   ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50540   ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50550   ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50560   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGHighFactor"
50570   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGLowFactor"
50580   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMaximumFactor"
50590   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMediumFactor"
50600   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMinimumFactor"
50610   ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50620   ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50630   ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50640   ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50650   ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50660   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGHighFactor"
50670   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGLowFactor"
50680   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMaximumFactor"
50690   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMediumFactor"
50700   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMinimumFactor"
50710   ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50720   ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50730   ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50740   ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50750   ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50760   ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50770   ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50780   ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50790   ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50800   ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50810   ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50820   ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50830   ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50840   ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50850   ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50860   ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50870   ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50880   ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50890   ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50900   ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50910   ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50920   ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
50930   ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
50940   ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
50950   ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
50960   ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
50970   ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
50980   ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
50990   ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
51000   ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
51010   ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
51020   ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
51030   ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
51040   ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
51050   ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
51060   ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
51070   ini.SaveKey CStr(.PrintAfterSavingTumble), "PrintAfterSavingTumble"
51080   ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
51090   ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
51100   ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
51110   ini.SaveKey CStr(.ProgramFont), "ProgramFont"
51120   ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
51130   ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
51140   ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
51150   ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
51160   ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
51170   ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
51180   ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
51190   ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
51200   ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51210   ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51220   ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51230   ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51240   ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51250   ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51260   ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51270   ini.SaveKey CStr(Abs(.SendEmailAfterAutoSaving)), "SendEmailAfterAutoSaving"
51280   ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51290   ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51300   ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51310   ini.SaveKey CStr(.StampFontname), "StampFontname"
51320   ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51330   ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51340   ini.SaveKey CStr(.StampString), "StampString"
51350   ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51360   ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51370   ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51380   ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51390   ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51400   ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51410   ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51420   ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51430   ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51440   ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51450   ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51460   ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51470   ini.SaveKey CStr(.Toolbars), "Toolbars"
51480   ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51490   ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51500   ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51510   ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51520  End With
51530  Set ini = Nothing
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
50430   tStr = reg.GetRegistryValue("DeviceHeightPoints")
50440   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
50450     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= -1 Then
50460       .DeviceHeightPoints = CDbl(Replace$(tStr, ".", GetDecimalChar))
50470      Else
50480       If UseStandard Then
50490        .DeviceHeightPoints = Replace$("-1", ".", GetDecimalChar)
50500       End If
50510     End If
50520    Else
50530     If UseStandard Then
50540      .DeviceHeightPoints = Replace$("-1", ".", GetDecimalChar)
50550     End If
50560   End If
50570   tStr = reg.GetRegistryValue("DeviceWidthPoints")
50580   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
50590     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= -1 Then
50600       .DeviceWidthPoints = CDbl(Replace$(tStr, ".", GetDecimalChar))
50610      Else
50620       If UseStandard Then
50630        .DeviceWidthPoints = Replace$("-1", ".", GetDecimalChar)
50640       End If
50650     End If
50660    Else
50670     If UseStandard Then
50680      .DeviceWidthPoints = Replace$("-1", ".", GetDecimalChar)
50690     End If
50700   End If
50710   tStr = reg.GetRegistryValue("OnePagePerFile")
50720   If IsNumeric(tStr) Then
50730     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
50740       .OnePagePerFile = CLng(tStr)
50750      Else
50760       If UseStandard Then
50770        .OnePagePerFile = 0
50780       End If
50790     End If
50800    Else
50810     If UseStandard Then
50820      .OnePagePerFile = 0
50830     End If
50840   End If
50850   tStr = reg.GetRegistryValue("Papersize")
50860   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
50870     .Papersize = ""
50880    Else
50890     If LenB(tStr) > 0 Then
50900      .Papersize = tStr
50910     End If
50920   End If
50930   tStr = reg.GetRegistryValue("StampFontColor")
50940   If LenB(tStr) = 0 And LenB("#FF0000") > 0 And UseStandard Then
50950     .StampFontColor = "#FF0000"
50960    Else
50970     If LenB(tStr) > 0 Then
50980      .StampFontColor = tStr
50990     End If
51000   End If
51010   tStr = reg.GetRegistryValue("StampFontname")
51020   If LenB(tStr) = 0 And LenB("Arial") > 0 And UseStandard Then
51030     .StampFontname = "Arial"
51040    Else
51050     If LenB(tStr) > 0 Then
51060      .StampFontname = tStr
51070     End If
51080   End If
51090   tStr = reg.GetRegistryValue("StampFontsize")
51100   If IsNumeric(tStr) Then
51110     If CLng(tStr) >= 1 Then
51120       .StampFontsize = CLng(tStr)
51130      Else
51140       If UseStandard Then
51150        .StampFontsize = 48
51160       End If
51170     End If
51180    Else
51190     If UseStandard Then
51200      .StampFontsize = 48
51210     End If
51220   End If
51230   tStr = reg.GetRegistryValue("StampOutlineFontthickness")
51240   If IsNumeric(tStr) Then
51250     If CLng(tStr) >= 0 Then
51260       .StampOutlineFontthickness = CLng(tStr)
51270      Else
51280       If UseStandard Then
51290        .StampOutlineFontthickness = 0
51300       End If
51310     End If
51320    Else
51330     If UseStandard Then
51340      .StampOutlineFontthickness = 0
51350     End If
51360   End If
51370   tStr = reg.GetRegistryValue("StampString")
51380   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
51390     .StampString = ""
51400    Else
51410     If LenB(tStr) > 0 Then
51420      .StampString = tStr
51430     End If
51440   End If
51450   tStr = reg.GetRegistryValue("StampUseOutlineFont")
51460   If IsNumeric(tStr) Then
51470     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51480       .StampUseOutlineFont = CLng(tStr)
51490      Else
51500       If UseStandard Then
51510        .StampUseOutlineFont = 1
51520       End If
51530     End If
51540    Else
51550     If UseStandard Then
51560      .StampUseOutlineFont = 1
51570     End If
51580   End If
51590   tStr = reg.GetRegistryValue("StandardAuthor")
51600   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
51610     .StandardAuthor = ""
51620    Else
51630     If LenB(tStr) > 0 Then
51640      .StandardAuthor = tStr
51650     End If
51660   End If
51670   tStr = reg.GetRegistryValue("StandardCreationdate")
51680   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
51690     .StandardCreationdate = ""
51700    Else
51710     If LenB(tStr) > 0 Then
51720      .StandardCreationdate = tStr
51730     End If
51740   End If
51750   tStr = reg.GetRegistryValue("StandardDateformat")
51760   If LenB(tStr) = 0 And LenB("YYYYMMDDHHNNSS") > 0 And UseStandard Then
51770     .StandardDateformat = "YYYYMMDDHHNNSS"
51780    Else
51790     If LenB(tStr) > 0 Then
51800      .StandardDateformat = tStr
51810     End If
51820   End If
51830   tStr = reg.GetRegistryValue("StandardKeywords")
51840   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
51850     .StandardKeywords = ""
51860    Else
51870     If LenB(tStr) > 0 Then
51880      .StandardKeywords = tStr
51890     End If
51900   End If
51910   tStr = reg.GetRegistryValue("StandardMailDomain")
51920   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
51930     .StandardMailDomain = ""
51940    Else
51950     If LenB(tStr) > 0 Then
51960      .StandardMailDomain = tStr
51970     End If
51980   End If
51990   tStr = reg.GetRegistryValue("StandardModifydate")
52000   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
52010     .StandardModifydate = ""
52020    Else
52030     If LenB(tStr) > 0 Then
52040      .StandardModifydate = tStr
52050     End If
52060   End If
52070   tStr = reg.GetRegistryValue("StandardSaveformat")
52080   If IsNumeric(tStr) Then
52090     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
52100       .StandardSaveformat = CLng(tStr)
52110      Else
52120       If UseStandard Then
52130        .StandardSaveformat = 0
52140       End If
52150     End If
52160    Else
52170     If UseStandard Then
52180      .StandardSaveformat = 0
52190     End If
52200   End If
52210   tStr = reg.GetRegistryValue("StandardSubject")
52220   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
52230     .StandardSubject = ""
52240    Else
52250     If LenB(tStr) > 0 Then
52260      .StandardSubject = tStr
52270     End If
52280   End If
52290   tStr = reg.GetRegistryValue("StandardTitle")
52300   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
52310     .StandardTitle = ""
52320    Else
52330     If LenB(tStr) > 0 Then
52340      .StandardTitle = tStr
52350     End If
52360   End If
52370   tStr = reg.GetRegistryValue("UseCreationDateNow")
52380   If IsNumeric(tStr) Then
52390     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52400       .UseCreationDateNow = CLng(tStr)
52410      Else
52420       If UseStandard Then
52430        .UseCreationDateNow = 0
52440       End If
52450     End If
52460    Else
52470     If UseStandard Then
52480      .UseCreationDateNow = 0
52490     End If
52500   End If
52510   tStr = reg.GetRegistryValue("UseStandardAuthor")
52520   If IsNumeric(tStr) Then
52530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52540       .UseStandardAuthor = CLng(tStr)
52550      Else
52560       If UseStandard Then
52570        .UseStandardAuthor = 0
52580       End If
52590     End If
52600    Else
52610     If UseStandard Then
52620      .UseStandardAuthor = 0
52630     End If
52640   End If
52650   reg.Subkey = "Printing\Formats\Bitmap\Colors"
52660   tStr = reg.GetRegistryValue("BitmapResolution")
52670   If IsNumeric(tStr) Then
52680     If CLng(tStr) >= 1 Then
52690       .BitmapResolution = CLng(tStr)
52700      Else
52710       If UseStandard Then
52720        .BitmapResolution = 150
52730       End If
52740     End If
52750    Else
52760     If UseStandard Then
52770      .BitmapResolution = 150
52780     End If
52790   End If
52800   tStr = reg.GetRegistryValue("BMPColorscount")
52810   If IsNumeric(tStr) Then
52820     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
52830       .BMPColorscount = CLng(tStr)
52840      Else
52850       If UseStandard Then
52860        .BMPColorscount = 1
52870       End If
52880     End If
52890    Else
52900     If UseStandard Then
52910      .BMPColorscount = 1
52920     End If
52930   End If
52940   tStr = reg.GetRegistryValue("JPEGColorscount")
52950   If IsNumeric(tStr) Then
52960     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
52970       .JPEGColorscount = CLng(tStr)
52980      Else
52990       If UseStandard Then
53000        .JPEGColorscount = 0
53010       End If
53020     End If
53030    Else
53040     If UseStandard Then
53050      .JPEGColorscount = 0
53060     End If
53070   End If
53080   tStr = reg.GetRegistryValue("JPEGQuality")
53090   If IsNumeric(tStr) Then
53100     If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
53110       .JPEGQuality = CLng(tStr)
53120      Else
53130       If UseStandard Then
53140        .JPEGQuality = 75
53150       End If
53160     End If
53170    Else
53180     If UseStandard Then
53190      .JPEGQuality = 75
53200     End If
53210   End If
53220   tStr = reg.GetRegistryValue("PCXColorscount")
53230   If IsNumeric(tStr) Then
53240     If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
53250       .PCXColorscount = CLng(tStr)
53260      Else
53270       If UseStandard Then
53280        .PCXColorscount = 0
53290       End If
53300     End If
53310    Else
53320     If UseStandard Then
53330      .PCXColorscount = 0
53340     End If
53350   End If
53360   tStr = reg.GetRegistryValue("PNGColorscount")
53370   If IsNumeric(tStr) Then
53380     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
53390       .PNGColorscount = CLng(tStr)
53400      Else
53410       If UseStandard Then
53420        .PNGColorscount = 0
53430       End If
53440     End If
53450    Else
53460     If UseStandard Then
53470      .PNGColorscount = 0
53480     End If
53490   End If
53500   tStr = reg.GetRegistryValue("TIFFColorscount")
53510   If IsNumeric(tStr) Then
53520     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
53530       .TIFFColorscount = CLng(tStr)
53540      Else
53550       If UseStandard Then
53560        .TIFFColorscount = 0
53570       End If
53580     End If
53590    Else
53600     If UseStandard Then
53610      .TIFFColorscount = 0
53620     End If
53630   End If
53640   reg.Subkey = "Printing\Formats\PDF\Colors"
53650   tStr = reg.GetRegistryValue("PDFColorsCMYKToRGB")
53660   If IsNumeric(tStr) Then
53670     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53680       .PDFColorsCMYKToRGB = CLng(tStr)
53690      Else
53700       If UseStandard Then
53710        .PDFColorsCMYKToRGB = 0
53720       End If
53730     End If
53740    Else
53750     If UseStandard Then
53760      .PDFColorsCMYKToRGB = 0
53770     End If
53780   End If
53790   tStr = reg.GetRegistryValue("PDFColorsColorModel")
53800   If IsNumeric(tStr) Then
53810     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
53820       .PDFColorsColorModel = CLng(tStr)
53830      Else
53840       If UseStandard Then
53850        .PDFColorsColorModel = 1
53860       End If
53870     End If
53880    Else
53890     If UseStandard Then
53900      .PDFColorsColorModel = 1
53910     End If
53920   End If
53930   tStr = reg.GetRegistryValue("PDFColorsPreserveHalftone")
53940   If IsNumeric(tStr) Then
53950     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53960       .PDFColorsPreserveHalftone = CLng(tStr)
53970      Else
53980       If UseStandard Then
53990        .PDFColorsPreserveHalftone = 0
54000       End If
54010     End If
54020    Else
54030     If UseStandard Then
54040      .PDFColorsPreserveHalftone = 0
54050     End If
54060   End If
54070   tStr = reg.GetRegistryValue("PDFColorsPreserveOverprint")
54080   If IsNumeric(tStr) Then
54090     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54100       .PDFColorsPreserveOverprint = CLng(tStr)
54110      Else
54120       If UseStandard Then
54130        .PDFColorsPreserveOverprint = 1
54140       End If
54150     End If
54160    Else
54170     If UseStandard Then
54180      .PDFColorsPreserveOverprint = 1
54190     End If
54200   End If
54210   tStr = reg.GetRegistryValue("PDFColorsPreserveTransfer")
54220   If IsNumeric(tStr) Then
54230     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54240       .PDFColorsPreserveTransfer = CLng(tStr)
54250      Else
54260       If UseStandard Then
54270        .PDFColorsPreserveTransfer = 1
54280       End If
54290     End If
54300    Else
54310     If UseStandard Then
54320      .PDFColorsPreserveTransfer = 1
54330     End If
54340   End If
54350   reg.Subkey = "Printing\Formats\PDF\Compression"
54360   tStr = reg.GetRegistryValue("PDFCompressionColorCompression")
54370   If IsNumeric(tStr) Then
54380     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54390       .PDFCompressionColorCompression = CLng(tStr)
54400      Else
54410       If UseStandard Then
54420        .PDFCompressionColorCompression = 1
54430       End If
54440     End If
54450    Else
54460     If UseStandard Then
54470      .PDFCompressionColorCompression = 1
54480     End If
54490   End If
54500   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionChoice")
54510   If IsNumeric(tStr) Then
54520     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
54530       .PDFCompressionColorCompressionChoice = CLng(tStr)
54540      Else
54550       If UseStandard Then
54560        .PDFCompressionColorCompressionChoice = 0
54570       End If
54580     End If
54590    Else
54600     If UseStandard Then
54610      .PDFCompressionColorCompressionChoice = 0
54620     End If
54630   End If
54640   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGHighFactor")
54650   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
54660     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
54670       .PDFCompressionColorCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
54680      Else
54690       If UseStandard Then
54700        .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
54710       End If
54720     End If
54730    Else
54740     If UseStandard Then
54750      .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
54760     End If
54770   End If
54780   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGLowFactor")
54790   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
54800     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
54810       .PDFCompressionColorCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
54820      Else
54830       If UseStandard Then
54840        .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
54850       End If
54860     End If
54870    Else
54880     If UseStandard Then
54890      .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
54900     End If
54910   End If
54920   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMaximumFactor")
54930   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
54940     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
54950       .PDFCompressionColorCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
54960      Else
54970       If UseStandard Then
54980        .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
54990       End If
55000     End If
55010    Else
55020     If UseStandard Then
55030      .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
55040     End If
55050   End If
55060   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMediumFactor")
55070   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
55080     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
55090       .PDFCompressionColorCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
55100      Else
55110       If UseStandard Then
55120        .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
55130       End If
55140     End If
55150    Else
55160     If UseStandard Then
55170      .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
55180     End If
55190   End If
55200   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMinimumFactor")
55210   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
55220     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
55230       .PDFCompressionColorCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
55240      Else
55250       If UseStandard Then
55260        .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
55270       End If
55280     End If
55290    Else
55300     If UseStandard Then
55310      .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
55320     End If
55330   End If
55340   tStr = reg.GetRegistryValue("PDFCompressionColorResample")
55350   If IsNumeric(tStr) Then
55360     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55370       .PDFCompressionColorResample = CLng(tStr)
55380      Else
55390       If UseStandard Then
55400        .PDFCompressionColorResample = 0
55410       End If
55420     End If
55430    Else
55440     If UseStandard Then
55450      .PDFCompressionColorResample = 0
55460     End If
55470   End If
55480   tStr = reg.GetRegistryValue("PDFCompressionColorResampleChoice")
55490   If IsNumeric(tStr) Then
55500     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
55510       .PDFCompressionColorResampleChoice = CLng(tStr)
55520      Else
55530       If UseStandard Then
55540        .PDFCompressionColorResampleChoice = 0
55550       End If
55560     End If
55570    Else
55580     If UseStandard Then
55590      .PDFCompressionColorResampleChoice = 0
55600     End If
55610   End If
55620   tStr = reg.GetRegistryValue("PDFCompressionColorResolution")
55630   If IsNumeric(tStr) Then
55640     If CLng(tStr) >= 0 Then
55650       .PDFCompressionColorResolution = CLng(tStr)
55660      Else
55670       If UseStandard Then
55680        .PDFCompressionColorResolution = 300
55690       End If
55700     End If
55710    Else
55720     If UseStandard Then
55730      .PDFCompressionColorResolution = 300
55740     End If
55750   End If
55760   tStr = reg.GetRegistryValue("PDFCompressionGreyCompression")
55770   If IsNumeric(tStr) Then
55780     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55790       .PDFCompressionGreyCompression = CLng(tStr)
55800      Else
55810       If UseStandard Then
55820        .PDFCompressionGreyCompression = 1
55830       End If
55840     End If
55850    Else
55860     If UseStandard Then
55870      .PDFCompressionGreyCompression = 1
55880     End If
55890   End If
55900   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionChoice")
55910   If IsNumeric(tStr) Then
55920     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
55930       .PDFCompressionGreyCompressionChoice = CLng(tStr)
55940      Else
55950       If UseStandard Then
55960        .PDFCompressionGreyCompressionChoice = 0
55970       End If
55980     End If
55990    Else
56000     If UseStandard Then
56010      .PDFCompressionGreyCompressionChoice = 0
56020     End If
56030   End If
56040   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGHighFactor")
56050   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56060     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56070       .PDFCompressionGreyCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56080      Else
56090       If UseStandard Then
56100        .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56110       End If
56120     End If
56130    Else
56140     If UseStandard Then
56150      .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56160     End If
56170   End If
56180   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGLowFactor")
56190   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56200     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56210       .PDFCompressionGreyCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56220      Else
56230       If UseStandard Then
56240        .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56250       End If
56260     End If
56270    Else
56280     If UseStandard Then
56290      .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56300     End If
56310   End If
56320   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMaximumFactor")
56330   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56340     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56350       .PDFCompressionGreyCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56360      Else
56370       If UseStandard Then
56380        .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56390       End If
56400     End If
56410    Else
56420     If UseStandard Then
56430      .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56440     End If
56450   End If
56460   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMediumFactor")
56470   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56480     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56490       .PDFCompressionGreyCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56500      Else
56510       If UseStandard Then
56520        .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56530       End If
56540     End If
56550    Else
56560     If UseStandard Then
56570      .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56580     End If
56590   End If
56600   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMinimumFactor")
56610   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56620     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56630       .PDFCompressionGreyCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56640      Else
56650       If UseStandard Then
56660        .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
56670       End If
56680     End If
56690    Else
56700     If UseStandard Then
56710      .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
56720     End If
56730   End If
56740   tStr = reg.GetRegistryValue("PDFCompressionGreyResample")
56750   If IsNumeric(tStr) Then
56760     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56770       .PDFCompressionGreyResample = CLng(tStr)
56780      Else
56790       If UseStandard Then
56800        .PDFCompressionGreyResample = 0
56810       End If
56820     End If
56830    Else
56840     If UseStandard Then
56850      .PDFCompressionGreyResample = 0
56860     End If
56870   End If
56880   tStr = reg.GetRegistryValue("PDFCompressionGreyResampleChoice")
56890   If IsNumeric(tStr) Then
56900     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
56910       .PDFCompressionGreyResampleChoice = CLng(tStr)
56920      Else
56930       If UseStandard Then
56940        .PDFCompressionGreyResampleChoice = 0
56950       End If
56960     End If
56970    Else
56980     If UseStandard Then
56990      .PDFCompressionGreyResampleChoice = 0
57000     End If
57010   End If
57020   tStr = reg.GetRegistryValue("PDFCompressionGreyResolution")
57030   If IsNumeric(tStr) Then
57040     If CLng(tStr) >= 0 Then
57050       .PDFCompressionGreyResolution = CLng(tStr)
57060      Else
57070       If UseStandard Then
57080        .PDFCompressionGreyResolution = 300
57090       End If
57100     End If
57110    Else
57120     If UseStandard Then
57130      .PDFCompressionGreyResolution = 300
57140     End If
57150   End If
57160   tStr = reg.GetRegistryValue("PDFCompressionMonoCompression")
57170   If IsNumeric(tStr) Then
57180     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57190       .PDFCompressionMonoCompression = CLng(tStr)
57200      Else
57210       If UseStandard Then
57220        .PDFCompressionMonoCompression = 1
57230       End If
57240     End If
57250    Else
57260     If UseStandard Then
57270      .PDFCompressionMonoCompression = 1
57280     End If
57290   End If
57300   tStr = reg.GetRegistryValue("PDFCompressionMonoCompressionChoice")
57310   If IsNumeric(tStr) Then
57320     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
57330       .PDFCompressionMonoCompressionChoice = CLng(tStr)
57340      Else
57350       If UseStandard Then
57360        .PDFCompressionMonoCompressionChoice = 0
57370       End If
57380     End If
57390    Else
57400     If UseStandard Then
57410      .PDFCompressionMonoCompressionChoice = 0
57420     End If
57430   End If
57440   tStr = reg.GetRegistryValue("PDFCompressionMonoResample")
57450   If IsNumeric(tStr) Then
57460     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57470       .PDFCompressionMonoResample = CLng(tStr)
57480      Else
57490       If UseStandard Then
57500        .PDFCompressionMonoResample = 0
57510       End If
57520     End If
57530    Else
57540     If UseStandard Then
57550      .PDFCompressionMonoResample = 0
57560     End If
57570   End If
57580   tStr = reg.GetRegistryValue("PDFCompressionMonoResampleChoice")
57590   If IsNumeric(tStr) Then
57600     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
57610       .PDFCompressionMonoResampleChoice = CLng(tStr)
57620      Else
57630       If UseStandard Then
57640        .PDFCompressionMonoResampleChoice = 0
57650       End If
57660     End If
57670    Else
57680     If UseStandard Then
57690      .PDFCompressionMonoResampleChoice = 0
57700     End If
57710   End If
57720   tStr = reg.GetRegistryValue("PDFCompressionMonoResolution")
57730   If IsNumeric(tStr) Then
57740     If CLng(tStr) >= 0 Then
57750       .PDFCompressionMonoResolution = CLng(tStr)
57760      Else
57770       If UseStandard Then
57780        .PDFCompressionMonoResolution = 1200
57790       End If
57800     End If
57810    Else
57820     If UseStandard Then
57830      .PDFCompressionMonoResolution = 1200
57840     End If
57850   End If
57860   tStr = reg.GetRegistryValue("PDFCompressionTextCompression")
57870   If IsNumeric(tStr) Then
57880     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57890       .PDFCompressionTextCompression = CLng(tStr)
57900      Else
57910       If UseStandard Then
57920        .PDFCompressionTextCompression = 1
57930       End If
57940     End If
57950    Else
57960     If UseStandard Then
57970      .PDFCompressionTextCompression = 1
57980     End If
57990   End If
58000   reg.Subkey = "Printing\Formats\PDF\Fonts"
58010   tStr = reg.GetRegistryValue("PDFFontsEmbedAll")
58020   If IsNumeric(tStr) Then
58030     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58040       .PDFFontsEmbedAll = CLng(tStr)
58050      Else
58060       If UseStandard Then
58070        .PDFFontsEmbedAll = 1
58080       End If
58090     End If
58100    Else
58110     If UseStandard Then
58120      .PDFFontsEmbedAll = 1
58130     End If
58140   End If
58150   tStr = reg.GetRegistryValue("PDFFontsSubSetFonts")
58160   If IsNumeric(tStr) Then
58170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58180       .PDFFontsSubSetFonts = CLng(tStr)
58190      Else
58200       If UseStandard Then
58210        .PDFFontsSubSetFonts = 1
58220       End If
58230     End If
58240    Else
58250     If UseStandard Then
58260      .PDFFontsSubSetFonts = 1
58270     End If
58280   End If
58290   tStr = reg.GetRegistryValue("PDFFontsSubSetFontsPercent")
58300   If IsNumeric(tStr) Then
58310     If CLng(tStr) >= 0 Then
58320       .PDFFontsSubSetFontsPercent = CLng(tStr)
58330      Else
58340       If UseStandard Then
58350        .PDFFontsSubSetFontsPercent = 100
58360       End If
58370     End If
58380    Else
58390     If UseStandard Then
58400      .PDFFontsSubSetFontsPercent = 100
58410     End If
58420   End If
58430   reg.Subkey = "Printing\Formats\PDF\General"
58440   tStr = reg.GetRegistryValue("PDFGeneralASCII85")
58450   If IsNumeric(tStr) Then
58460     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58470       .PDFGeneralASCII85 = CLng(tStr)
58480      Else
58490       If UseStandard Then
58500        .PDFGeneralASCII85 = 0
58510       End If
58520     End If
58530    Else
58540     If UseStandard Then
58550      .PDFGeneralASCII85 = 0
58560     End If
58570   End If
58580   tStr = reg.GetRegistryValue("PDFGeneralAutorotate")
58590   If IsNumeric(tStr) Then
58600     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
58610       .PDFGeneralAutorotate = CLng(tStr)
58620      Else
58630       If UseStandard Then
58640        .PDFGeneralAutorotate = 2
58650       End If
58660     End If
58670    Else
58680     If UseStandard Then
58690      .PDFGeneralAutorotate = 2
58700     End If
58710   End If
58720   tStr = reg.GetRegistryValue("PDFGeneralCompatibility")
58730   If IsNumeric(tStr) Then
58740     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
58750       .PDFGeneralCompatibility = CLng(tStr)
58760      Else
58770       If UseStandard Then
58780        .PDFGeneralCompatibility = 1
58790       End If
58800     End If
58810    Else
58820     If UseStandard Then
58830      .PDFGeneralCompatibility = 1
58840     End If
58850   End If
58860   tStr = reg.GetRegistryValue("PDFGeneralOverprint")
58870   If IsNumeric(tStr) Then
58880     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
58890       .PDFGeneralOverprint = CLng(tStr)
58900      Else
58910       If UseStandard Then
58920        .PDFGeneralOverprint = 0
58930       End If
58940     End If
58950    Else
58960     If UseStandard Then
58970      .PDFGeneralOverprint = 0
58980     End If
58990   End If
59000   tStr = reg.GetRegistryValue("PDFGeneralResolution")
59010   If IsNumeric(tStr) Then
59020     If CLng(tStr) >= 0 Then
59030       .PDFGeneralResolution = CLng(tStr)
59040      Else
59050       If UseStandard Then
59060        .PDFGeneralResolution = 600
59070       End If
59080     End If
59090    Else
59100     If UseStandard Then
59110      .PDFGeneralResolution = 600
59120     End If
59130   End If
59140   tStr = reg.GetRegistryValue("PDFOptimize")
59150   If IsNumeric(tStr) Then
59160     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59170       .PDFOptimize = CLng(tStr)
59180      Else
59190       If UseStandard Then
59200        .PDFOptimize = 0
59210       End If
59220     End If
59230    Else
59240     If UseStandard Then
59250      .PDFOptimize = 0
59260     End If
59270   End If
59280   reg.Subkey = "Printing\Formats\PDF\Security"
59290   tStr = reg.GetRegistryValue("PDFAllowAssembly")
59300   If IsNumeric(tStr) Then
59310     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59320       .PDFAllowAssembly = CLng(tStr)
59330      Else
59340       If UseStandard Then
59350        .PDFAllowAssembly = 0
59360       End If
59370     End If
59380    Else
59390     If UseStandard Then
59400      .PDFAllowAssembly = 0
59410     End If
59420   End If
59430   tStr = reg.GetRegistryValue("PDFAllowDegradedPrinting")
59440   If IsNumeric(tStr) Then
59450     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59460       .PDFAllowDegradedPrinting = CLng(tStr)
59470      Else
59480       If UseStandard Then
59490        .PDFAllowDegradedPrinting = 0
59500       End If
59510     End If
59520    Else
59530     If UseStandard Then
59540      .PDFAllowDegradedPrinting = 0
59550     End If
59560   End If
59570   tStr = reg.GetRegistryValue("PDFAllowFillIn")
59580   If IsNumeric(tStr) Then
59590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59600       .PDFAllowFillIn = CLng(tStr)
59610      Else
59620       If UseStandard Then
59630        .PDFAllowFillIn = 0
59640       End If
59650     End If
59660    Else
59670     If UseStandard Then
59680      .PDFAllowFillIn = 0
59690     End If
59700   End If
59710   tStr = reg.GetRegistryValue("PDFAllowScreenReaders")
59720   If IsNumeric(tStr) Then
59730     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59740       .PDFAllowScreenReaders = CLng(tStr)
59750      Else
59760       If UseStandard Then
59770        .PDFAllowScreenReaders = 0
59780       End If
59790     End If
59800    Else
59810     If UseStandard Then
59820      .PDFAllowScreenReaders = 0
59830     End If
59840   End If
59850   tStr = reg.GetRegistryValue("PDFDisallowCopy")
59860   If IsNumeric(tStr) Then
59870     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59880       .PDFDisallowCopy = CLng(tStr)
59890      Else
59900       If UseStandard Then
59910        .PDFDisallowCopy = 1
59920       End If
59930     End If
59940    Else
59950     If UseStandard Then
59960      .PDFDisallowCopy = 1
59970     End If
59980   End If
59990   tStr = reg.GetRegistryValue("PDFDisallowModifyAnnotations")
60000   If IsNumeric(tStr) Then
60010     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60020       .PDFDisallowModifyAnnotations = CLng(tStr)
60030      Else
60040       If UseStandard Then
60050        .PDFDisallowModifyAnnotations = 0
60060       End If
60070     End If
60080    Else
60090     If UseStandard Then
60100      .PDFDisallowModifyAnnotations = 0
60110     End If
60120   End If
60130   tStr = reg.GetRegistryValue("PDFDisallowModifyContents")
60140   If IsNumeric(tStr) Then
60150     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60160       .PDFDisallowModifyContents = CLng(tStr)
60170      Else
60180       If UseStandard Then
60190        .PDFDisallowModifyContents = 0
60200       End If
60210     End If
60220    Else
60230     If UseStandard Then
60240      .PDFDisallowModifyContents = 0
60250     End If
60260   End If
60270   tStr = reg.GetRegistryValue("PDFDisallowPrinting")
60280   If IsNumeric(tStr) Then
60290     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60300       .PDFDisallowPrinting = CLng(tStr)
60310      Else
60320       If UseStandard Then
60330        .PDFDisallowPrinting = 0
60340       End If
60350     End If
60360    Else
60370     If UseStandard Then
60380      .PDFDisallowPrinting = 0
60390     End If
60400   End If
60410   tStr = reg.GetRegistryValue("PDFEncryptor")
60420   If IsNumeric(tStr) Then
60430     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
60440       .PDFEncryptor = CLng(tStr)
60450      Else
60460       If UseStandard Then
60470        .PDFEncryptor = 0
60480       End If
60490     End If
60500    Else
60510     If UseStandard Then
60520      .PDFEncryptor = 0
60530     End If
60540   End If
60550   tStr = reg.GetRegistryValue("PDFHighEncryption")
60560   If IsNumeric(tStr) Then
60570     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60580       .PDFHighEncryption = CLng(tStr)
60590      Else
60600       If UseStandard Then
60610        .PDFHighEncryption = 0
60620       End If
60630     End If
60640    Else
60650     If UseStandard Then
60660      .PDFHighEncryption = 0
60670     End If
60680   End If
60690   tStr = reg.GetRegistryValue("PDFLowEncryption")
60700   If IsNumeric(tStr) Then
60710     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60720       .PDFLowEncryption = CLng(tStr)
60730      Else
60740       If UseStandard Then
60750        .PDFLowEncryption = 1
60760       End If
60770     End If
60780    Else
60790     If UseStandard Then
60800      .PDFLowEncryption = 1
60810     End If
60820   End If
60830   tStr = reg.GetRegistryValue("PDFOwnerPass")
60840   If IsNumeric(tStr) Then
60850     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60860       .PDFOwnerPass = CLng(tStr)
60870      Else
60880       If UseStandard Then
60890        .PDFOwnerPass = 0
60900       End If
60910     End If
60920    Else
60930     If UseStandard Then
60940      .PDFOwnerPass = 0
60950     End If
60960   End If
60970   tStr = reg.GetRegistryValue("PDFOwnerPasswordString")
60980   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
60990     .PDFOwnerPasswordString = ""
61000    Else
61010     If LenB(tStr) > 0 Then
61020      .PDFOwnerPasswordString = tStr
61030     End If
61040   End If
61050   tStr = reg.GetRegistryValue("PDFUserPass")
61060   If IsNumeric(tStr) Then
61070     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61080       .PDFUserPass = CLng(tStr)
61090      Else
61100       If UseStandard Then
61110        .PDFUserPass = 0
61120       End If
61130     End If
61140    Else
61150     If UseStandard Then
61160      .PDFUserPass = 0
61170     End If
61180   End If
61190   tStr = reg.GetRegistryValue("PDFUserPasswordString")
61200   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
61210     .PDFUserPasswordString = ""
61220    Else
61230     If LenB(tStr) > 0 Then
61240      .PDFUserPasswordString = tStr
61250     End If
61260   End If
61270   tStr = reg.GetRegistryValue("PDFUseSecurity")
61280   If IsNumeric(tStr) Then
61290     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61300       .PDFUseSecurity = CLng(tStr)
61310      Else
61320       If UseStandard Then
61330        .PDFUseSecurity = 0
61340       End If
61350     End If
61360    Else
61370     If UseStandard Then
61380      .PDFUseSecurity = 0
61390     End If
61400   End If
61410   reg.Subkey = "Printing\Formats\PS\LanguageLevel"
61420   tStr = reg.GetRegistryValue("EPSLanguageLevel")
61430   If IsNumeric(tStr) Then
61440     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
61450       .EPSLanguageLevel = CLng(tStr)
61460      Else
61470       If UseStandard Then
61480        .EPSLanguageLevel = 2
61490       End If
61500     End If
61510    Else
61520     If UseStandard Then
61530      .EPSLanguageLevel = 2
61540     End If
61550   End If
61560   tStr = reg.GetRegistryValue("PSLanguageLevel")
61570   If IsNumeric(tStr) Then
61580     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
61590       .PSLanguageLevel = CLng(tStr)
61600      Else
61610       If UseStandard Then
61620        .PSLanguageLevel = 2
61630       End If
61640     End If
61650    Else
61660     If UseStandard Then
61670      .PSLanguageLevel = 2
61680     End If
61690   End If
61700   reg.Subkey = "Program"
61710   tStr = reg.GetRegistryValue("AdditionalGhostscriptParameters")
61720   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
61730     .AdditionalGhostscriptParameters = ""
61740    Else
61750     If LenB(tStr) > 0 Then
61760      .AdditionalGhostscriptParameters = tStr
61770     End If
61780   End If
61790   tStr = reg.GetRegistryValue("AdditionalGhostscriptSearchpath")
61800   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
61810     .AdditionalGhostscriptSearchpath = ""
61820    Else
61830     If LenB(tStr) > 0 Then
61840      .AdditionalGhostscriptSearchpath = tStr
61850     End If
61860   End If
61870   tStr = reg.GetRegistryValue("AddWindowsFontpath")
61880   If IsNumeric(tStr) Then
61890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61900       .AddWindowsFontpath = CLng(tStr)
61910      Else
61920       If UseStandard Then
61930        .AddWindowsFontpath = 1
61940       End If
61950     End If
61960    Else
61970     If UseStandard Then
61980      .AddWindowsFontpath = 1
61990     End If
62000   End If
62010   tStr = reg.GetRegistryValue("AutosaveDirectory")
62020   If LenB(Trim$(tStr)) > 0 Then
62030     .AutosaveDirectory = CompletePath(tStr)
62040    Else
62050     If UseStandard Then
62060      If InstalledAsServer Then
62070        .AutosaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
62080       Else
62090        .AutosaveDirectory = "<MyFiles>"
62100      End If
62110     End If
62120   End If
62130   tStr = reg.GetRegistryValue("AutosaveFilename")
62140   If LenB(tStr) = 0 And LenB("<DateTime>") > 0 And UseStandard Then
62150     .AutosaveFilename = "<DateTime>"
62160    Else
62170     If LenB(tStr) > 0 Then
62180      .AutosaveFilename = tStr
62190     End If
62200   End If
62210   tStr = reg.GetRegistryValue("AutosaveFormat")
62220   If IsNumeric(tStr) Then
62230     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
62240       .AutosaveFormat = CLng(tStr)
62250      Else
62260       If UseStandard Then
62270        .AutosaveFormat = 0
62280       End If
62290     End If
62300    Else
62310     If UseStandard Then
62320      .AutosaveFormat = 0
62330     End If
62340   End If
62350   tStr = reg.GetRegistryValue("AutosaveStartStandardProgram")
62360   If IsNumeric(tStr) Then
62370     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62380       .AutosaveStartStandardProgram = CLng(tStr)
62390      Else
62400       If UseStandard Then
62410        .AutosaveStartStandardProgram = 0
62420       End If
62430     End If
62440    Else
62450     If UseStandard Then
62460      .AutosaveStartStandardProgram = 0
62470     End If
62480   End If
62490   tStr = reg.GetRegistryValue("ClientComputerResolveIPAddress")
62500   If IsNumeric(tStr) Then
62510     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62520       .ClientComputerResolveIPAddress = CLng(tStr)
62530      Else
62540       If UseStandard Then
62550        .ClientComputerResolveIPAddress = 0
62560       End If
62570     End If
62580    Else
62590     If UseStandard Then
62600      .ClientComputerResolveIPAddress = 0
62610     End If
62620   End If
62630   tStr = reg.GetRegistryValue("DisableEmail")
62640   If IsNumeric(tStr) Then
62650     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62660       .DisableEmail = CLng(tStr)
62670      Else
62680       If UseStandard Then
62690        .DisableEmail = 0
62700       End If
62710     End If
62720    Else
62730     If UseStandard Then
62740      .DisableEmail = 0
62750     End If
62760   End If
62770   tStr = reg.GetRegistryValue("DontUseDocumentSettings")
62780   If IsNumeric(tStr) Then
62790     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62800       .DontUseDocumentSettings = CLng(tStr)
62810      Else
62820       If UseStandard Then
62830        .DontUseDocumentSettings = 0
62840       End If
62850     End If
62860    Else
62870     If UseStandard Then
62880      .DontUseDocumentSettings = 0
62890     End If
62900   End If
62910   tStr = reg.GetRegistryValue("FilenameSubstitutions")
62920   If LenB(tStr) = 0 And LenB("Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt") > 0 And UseStandard Then
62930     .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
62940    Else
62950     If LenB(tStr) > 0 Then
62960      .FilenameSubstitutions = tStr
62970     End If
62980   End If
62990   tStr = reg.GetRegistryValue("FilenameSubstitutionsOnlyInTitle")
63000   If IsNumeric(tStr) Then
63010     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63020       .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
63030      Else
63040       If UseStandard Then
63050        .FilenameSubstitutionsOnlyInTitle = 1
63060       End If
63070     End If
63080    Else
63090     If UseStandard Then
63100      .FilenameSubstitutionsOnlyInTitle = 1
63110     End If
63120   End If
63130   tStr = reg.GetRegistryValue("Language")
63140   If LenB(tStr) = 0 And LenB("english") > 0 And UseStandard Then
63150     .Language = "english"
63160    Else
63170     If LenB(tStr) > 0 Then
63180      .Language = tStr
63190     End If
63200   End If
63210   tStr = reg.GetRegistryValue("LastSaveDirectory")
63220   If LenB(Trim$(tStr)) > 0 Then
63230     .LastSaveDirectory = CompletePath(tStr)
63240    Else
63250     If UseStandard Then
63260      If InstalledAsServer Then
63270        .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
63280       Else
63290        .LastSaveDirectory = "<MyFiles>"
63300      End If
63310     End If
63320   End If
63330   tStr = reg.GetRegistryValue("Logging")
63340   If IsNumeric(tStr) Then
63350     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63360       .Logging = CLng(tStr)
63370      Else
63380       If UseStandard Then
63390        .Logging = 0
63400       End If
63410     End If
63420    Else
63430     If UseStandard Then
63440      .Logging = 0
63450     End If
63460   End If
63470   tStr = reg.GetRegistryValue("LogLines")
63480   If IsNumeric(tStr) Then
63490     If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
63500       .LogLines = CLng(tStr)
63510      Else
63520       If UseStandard Then
63530        .LogLines = 100
63540       End If
63550     End If
63560    Else
63570     If UseStandard Then
63580      .LogLines = 100
63590     End If
63600   End If
63610   tStr = reg.GetRegistryValue("NoConfirmMessageSwitchingDefaultprinter")
63620   If IsNumeric(tStr) Then
63630     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63640       .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
63650      Else
63660       If UseStandard Then
63670        .NoConfirmMessageSwitchingDefaultprinter = 0
63680       End If
63690     End If
63700    Else
63710     If UseStandard Then
63720      .NoConfirmMessageSwitchingDefaultprinter = 0
63730     End If
63740   End If
63750   tStr = reg.GetRegistryValue("NoProcessingAtStartup")
63760   If IsNumeric(tStr) Then
63770     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63780       .NoProcessingAtStartup = CLng(tStr)
63790      Else
63800       If UseStandard Then
63810        .NoProcessingAtStartup = 0
63820       End If
63830     End If
63840    Else
63850     If UseStandard Then
63860      .NoProcessingAtStartup = 0
63870     End If
63880   End If
63890   tStr = reg.GetRegistryValue("NoPSCheck")
63900   If IsNumeric(tStr) Then
63910     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63920       .NoPSCheck = CLng(tStr)
63930      Else
63940       If UseStandard Then
63950        .NoPSCheck = 0
63960       End If
63970     End If
63980    Else
63990     If UseStandard Then
64000      .NoPSCheck = 0
64010     End If
64020   End If
64030   tStr = reg.GetRegistryValue("OptionsDesign")
64040   If IsNumeric(tStr) Then
64050     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
64060       .OptionsDesign = CLng(tStr)
64070      Else
64080       If UseStandard Then
64090        .OptionsDesign = 0
64100       End If
64110     End If
64120    Else
64130     If UseStandard Then
64140      .OptionsDesign = 0
64150     End If
64160   End If
64170   tStr = reg.GetRegistryValue("OptionsEnabled")
64180   If IsNumeric(tStr) Then
64190     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64200       .OptionsEnabled = CLng(tStr)
64210      Else
64220       If UseStandard Then
64230        .OptionsEnabled = 1
64240       End If
64250     End If
64260    Else
64270     If UseStandard Then
64280      .OptionsEnabled = 1
64290     End If
64300   End If
64310   tStr = reg.GetRegistryValue("OptionsVisible")
64320   If IsNumeric(tStr) Then
64330     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64340       .OptionsVisible = CLng(tStr)
64350      Else
64360       If UseStandard Then
64370        .OptionsVisible = 1
64380       End If
64390     End If
64400    Else
64410     If UseStandard Then
64420      .OptionsVisible = 1
64430     End If
64440   End If
64450   tStr = reg.GetRegistryValue("PrintAfterSaving")
64460   If IsNumeric(tStr) Then
64470     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64480       .PrintAfterSaving = CLng(tStr)
64490      Else
64500       If UseStandard Then
64510        .PrintAfterSaving = 0
64520       End If
64530     End If
64540    Else
64550     If UseStandard Then
64560      .PrintAfterSaving = 0
64570     End If
64580   End If
64590   tStr = reg.GetRegistryValue("PrintAfterSavingDuplex")
64600   If IsNumeric(tStr) Then
64610     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64620       .PrintAfterSavingDuplex = CLng(tStr)
64630      Else
64640       If UseStandard Then
64650        .PrintAfterSavingDuplex = 0
64660       End If
64670     End If
64680    Else
64690     If UseStandard Then
64700      .PrintAfterSavingDuplex = 0
64710     End If
64720   End If
64730   tStr = reg.GetRegistryValue("PrintAfterSavingNoCancel")
64740   If IsNumeric(tStr) Then
64750     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64760       .PrintAfterSavingNoCancel = CLng(tStr)
64770      Else
64780       If UseStandard Then
64790        .PrintAfterSavingNoCancel = 0
64800       End If
64810     End If
64820    Else
64830     If UseStandard Then
64840      .PrintAfterSavingNoCancel = 0
64850     End If
64860   End If
64870   tStr = reg.GetRegistryValue("PrintAfterSavingPrinter")
64880   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64890     .PrintAfterSavingPrinter = ""
64900    Else
64910     If LenB(tStr) > 0 Then
64920      .PrintAfterSavingPrinter = tStr
64930     End If
64940   End If
64950   tStr = reg.GetRegistryValue("PrintAfterSavingQueryUser")
64960   If IsNumeric(tStr) Then
64970     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
64980       .PrintAfterSavingQueryUser = CLng(tStr)
64990      Else
65000       If UseStandard Then
65010        .PrintAfterSavingQueryUser = 0
65020       End If
65030     End If
65040    Else
65050     If UseStandard Then
65060      .PrintAfterSavingQueryUser = 0
65070     End If
65080   End If
65090   tStr = reg.GetRegistryValue("PrintAfterSavingTumble")
65100   If IsNumeric(tStr) Then
65110     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
65120       .PrintAfterSavingTumble = CLng(tStr)
65130      Else
65140       If UseStandard Then
65150        .PrintAfterSavingTumble = 0
65160       End If
65170     End If
65180    Else
65190     If UseStandard Then
65200      .PrintAfterSavingTumble = 0
65210     End If
65220   End If
65230   tStr = reg.GetRegistryValue("PrinterStop")
65240   If IsNumeric(tStr) Then
65250     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65260       .PrinterStop = CLng(tStr)
65270      Else
65280       If UseStandard Then
65290        .PrinterStop = 0
65300       End If
65310     End If
65320    Else
65330     If UseStandard Then
65340      .PrinterStop = 0
65350     End If
65360   End If
65370   tStr = reg.GetRegistryValue("PrinterTemppath")
65380   WriteToSpecialLogfile "reg.GetRegistryValue(""PrinterTemppath"")=" & tStr
65390   WriteToSpecialLogfile "Options.PrinterTemppath1=" & .PrinterTemppath
65400   If hkey1 = HKEY_USERS Then
65410     If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then
65420       .PrinterTemppath = tStr
65430      Else
65440       If UseStandard Then
65450         .PrinterTemppath = GetTempPath
65460        Else
65470         .PrinterTemppath = tStr
65480       End If
65490     End If
65500    Else
65510     If LenB(Trim$(tStr)) > 0 Then
65520      If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then
65530        .PrinterTemppath = tStr
65540       Else
65550        MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))
65560        If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then
65570          If UseStandard Then
65580            .PrinterTemppath = GetTempPath
65590           Else
65600            .PrinterTemppath = ""
65610            If NoMsg = False Then
65620             MsgBox "PrinterTemppath: '" & tStr & "' = '" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & "'" & _
             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
65640            End If
65650          End If
65660         Else
65670          .PrinterTemppath = tStr
65680        End If
65690      End If
65700     End If
65710   End If
65720   WriteToSpecialLogfile "Options.PrinterTemppath2=" & .PrinterTemppath
65730   tStr = reg.GetRegistryValue("ProcessPriority")
65740   If IsNumeric(tStr) Then
65750     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
65760       .ProcessPriority = CLng(tStr)
65770      Else
65780       If UseStandard Then
65790        .ProcessPriority = 1
65800       End If
65810     End If
65820    Else
65830     If UseStandard Then
65840      .ProcessPriority = 1
65850     End If
65860   End If
65870   tStr = reg.GetRegistryValue("ProgramFont")
65880   If LenB(tStr) = 0 And LenB("MS Sans Serif") > 0 And UseStandard Then
65890     .ProgramFont = "MS Sans Serif"
65900    Else
65910     If LenB(tStr) > 0 Then
65920      .ProgramFont = tStr
65930     End If
65940   End If
65950   tStr = reg.GetRegistryValue("ProgramFontCharset")
65960   If IsNumeric(tStr) Then
65970     If CLng(tStr) >= 0 Then
65980       .ProgramFontCharset = CLng(tStr)
65990      Else
66000       If UseStandard Then
66010        .ProgramFontCharset = 0
66020       End If
66030     End If
66040    Else
66050     If UseStandard Then
66060      .ProgramFontCharset = 0
66070     End If
66080   End If
66090   tStr = reg.GetRegistryValue("ProgramFontSize")
66100   If IsNumeric(tStr) Then
66110     If CLng(tStr) >= 1 And CLng(tStr) <= 72 Then
66120       .ProgramFontSize = CLng(tStr)
66130      Else
66140       If UseStandard Then
66150        .ProgramFontSize = 8
66160       End If
66170     End If
66180    Else
66190     If UseStandard Then
66200      .ProgramFontSize = 8
66210     End If
66220   End If
66230   tStr = reg.GetRegistryValue("RemoveAllKnownFileExtensions")
66240   If IsNumeric(tStr) Then
66250     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66260       .RemoveAllKnownFileExtensions = CLng(tStr)
66270      Else
66280       If UseStandard Then
66290        .RemoveAllKnownFileExtensions = 1
66300       End If
66310     End If
66320    Else
66330     If UseStandard Then
66340      .RemoveAllKnownFileExtensions = 1
66350     End If
66360   End If
66370   tStr = reg.GetRegistryValue("RemoveSpaces")
66380   If IsNumeric(tStr) Then
66390     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66400       .RemoveSpaces = CLng(tStr)
66410      Else
66420       If UseStandard Then
66430        .RemoveSpaces = 1
66440       End If
66450     End If
66460    Else
66470     If UseStandard Then
66480      .RemoveSpaces = 1
66490     End If
66500   End If
66510   tStr = reg.GetRegistryValue("RunProgramAfterSaving")
66520   If IsNumeric(tStr) Then
66530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66540       .RunProgramAfterSaving = CLng(tStr)
66550      Else
66560       If UseStandard Then
66570        .RunProgramAfterSaving = 0
66580       End If
66590     End If
66600    Else
66610     If UseStandard Then
66620      .RunProgramAfterSaving = 0
66630     End If
66640   End If
66650   tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramname")
66660   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
66670     .RunProgramAfterSavingProgramname = ""
66680    Else
66690     If LenB(tStr) > 0 Then
66700      .RunProgramAfterSavingProgramname = tStr
66710     End If
66720   End If
66730   tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramParameters")
66740   If LenB(tStr) = 0 And LenB("""<OutputFilename>""") > 0 And UseStandard Then
66750     .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
66760    Else
66770     If LenB(tStr) > 0 Then
66780      .RunProgramAfterSavingProgramParameters = tStr
66790     End If
66800   End If
66810   tStr = reg.GetRegistryValue("RunProgramAfterSavingWaitUntilReady")
66820   If IsNumeric(tStr) Then
66830     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66840       .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
66850      Else
66860       If UseStandard Then
66870        .RunProgramAfterSavingWaitUntilReady = 1
66880       End If
66890     End If
66900    Else
66910     If UseStandard Then
66920      .RunProgramAfterSavingWaitUntilReady = 1
66930     End If
66940   End If
66950   tStr = reg.GetRegistryValue("RunProgramAfterSavingWindowstyle")
66960   If IsNumeric(tStr) Then
66970     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
66980       .RunProgramAfterSavingWindowstyle = CLng(tStr)
66990      Else
67000       If UseStandard Then
67010        .RunProgramAfterSavingWindowstyle = 1
67020       End If
67030     End If
67040    Else
67050     If UseStandard Then
67060      .RunProgramAfterSavingWindowstyle = 1
67070     End If
67080   End If
67090   tStr = reg.GetRegistryValue("RunProgramBeforeSaving")
67100   If IsNumeric(tStr) Then
67110     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67120       .RunProgramBeforeSaving = CLng(tStr)
67130      Else
67140       If UseStandard Then
67150        .RunProgramBeforeSaving = 0
67160       End If
67170     End If
67180    Else
67190     If UseStandard Then
67200      .RunProgramBeforeSaving = 0
67210     End If
67220   End If
67230   tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramname")
67240   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67250     .RunProgramBeforeSavingProgramname = ""
67260    Else
67270     If LenB(tStr) > 0 Then
67280      .RunProgramBeforeSavingProgramname = tStr
67290     End If
67300   End If
67310   tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramParameters")
67320   If LenB(tStr) = 0 And LenB("""<TempFilename>""") > 0 And UseStandard Then
67330     .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
67340    Else
67350     If LenB(tStr) > 0 Then
67360      .RunProgramBeforeSavingProgramParameters = tStr
67370     End If
67380   End If
67390   tStr = reg.GetRegistryValue("RunProgramBeforeSavingWindowstyle")
67400   If IsNumeric(tStr) Then
67410     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
67420       .RunProgramBeforeSavingWindowstyle = CLng(tStr)
67430      Else
67440       If UseStandard Then
67450        .RunProgramBeforeSavingWindowstyle = 1
67460       End If
67470     End If
67480    Else
67490     If UseStandard Then
67500      .RunProgramBeforeSavingWindowstyle = 1
67510     End If
67520   End If
67530   tStr = reg.GetRegistryValue("SaveFilename")
67540   If LenB(tStr) = 0 And LenB("<Title>") > 0 And UseStandard Then
67550     .SaveFilename = "<Title>"
67560    Else
67570     If LenB(tStr) > 0 Then
67580      .SaveFilename = tStr
67590     End If
67600   End If
67610   tStr = reg.GetRegistryValue("SendEmailAfterAutoSaving")
67620   If IsNumeric(tStr) Then
67630     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67640       .SendEmailAfterAutoSaving = CLng(tStr)
67650      Else
67660       If UseStandard Then
67670        .SendEmailAfterAutoSaving = 0
67680       End If
67690     End If
67700    Else
67710     If UseStandard Then
67720      .SendEmailAfterAutoSaving = 0
67730     End If
67740   End If
67750   tStr = reg.GetRegistryValue("SendMailMethod")
67760   If IsNumeric(tStr) Then
67770     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
67780       .SendMailMethod = CLng(tStr)
67790      Else
67800       If UseStandard Then
67810        .SendMailMethod = 0
67820       End If
67830     End If
67840    Else
67850     If UseStandard Then
67860      .SendMailMethod = 0
67870     End If
67880   End If
67890   tStr = reg.GetRegistryValue("ShowAnimation")
67900   If IsNumeric(tStr) Then
67910     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67920       .ShowAnimation = CLng(tStr)
67930      Else
67940       If UseStandard Then
67950        .ShowAnimation = 1
67960       End If
67970     End If
67980    Else
67990     If UseStandard Then
68000      .ShowAnimation = 1
68010     End If
68020   End If
68030   tStr = reg.GetRegistryValue("StartStandardProgram")
68040   If IsNumeric(tStr) Then
68050     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68060       .StartStandardProgram = CLng(tStr)
68070      Else
68080       If UseStandard Then
68090        .StartStandardProgram = 1
68100       End If
68110     End If
68120    Else
68130     If UseStandard Then
68140      .StartStandardProgram = 1
68150     End If
68160   End If
68170   tStr = reg.GetRegistryValue("Toolbars")
68180   If IsNumeric(tStr) Then
68190     If CLng(tStr) >= 0 Then
68200       .Toolbars = CLng(tStr)
68210      Else
68220       If UseStandard Then
68230        .Toolbars = 1
68240       End If
68250     End If
68260    Else
68270     If UseStandard Then
68280      .Toolbars = 1
68290     End If
68300   End If
68310   tStr = reg.GetRegistryValue("UseAutosave")
68320   If IsNumeric(tStr) Then
68330     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68340       .UseAutosave = CLng(tStr)
68350      Else
68360       If UseStandard Then
68370        .UseAutosave = 0
68380       End If
68390     End If
68400    Else
68410     If UseStandard Then
68420      .UseAutosave = 0
68430     End If
68440   End If
68450   tStr = reg.GetRegistryValue("UseAutosaveDirectory")
68460   If IsNumeric(tStr) Then
68470     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68480       .UseAutosaveDirectory = CLng(tStr)
68490      Else
68500       If UseStandard Then
68510        .UseAutosaveDirectory = 1
68520       End If
68530     End If
68540    Else
68550     If UseStandard Then
68560      .UseAutosaveDirectory = 1
68570     End If
68580   End If
68590  End With
68600  Set reg = Nothing
68610  ReadOptionsReg = myOptions
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
50400   If UCase$(OptionName) = "DEVICEHEIGHTPOINTS" Then
50410    If Not reg.KeyExists Then
50420     reg.CreateKey
50430    End If
50440   reg.SetRegistryValue "DeviceHeightPoints", Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), REG_SZ
50450    Set reg = Nothing
50460    Exit Sub
50470   End If
50480   If UCase$(OptionName) = "DEVICEWIDTHPOINTS" Then
50490    If Not reg.KeyExists Then
50500     reg.CreateKey
50510    End If
50520   reg.SetRegistryValue "DeviceWidthPoints", Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), REG_SZ
50530    Set reg = Nothing
50540    Exit Sub
50550   End If
50560   If UCase$(OptionName) = "ONEPAGEPERFILE" Then
50570    If Not reg.KeyExists Then
50580     reg.CreateKey
50590    End If
50600    reg.SetRegistryValue "OnePagePerFile", CStr(Abs(.OnePagePerFile)), REG_SZ
50610    Set reg = Nothing
50620    Exit Sub
50630   End If
50640   If UCase$(OptionName) = "PAPERSIZE" Then
50650    If Not reg.KeyExists Then
50660     reg.CreateKey
50670    End If
50680    reg.SetRegistryValue "Papersize", CStr(.Papersize), REG_SZ
50690    Set reg = Nothing
50700    Exit Sub
50710   End If
50720   If UCase$(OptionName) = "STAMPFONTCOLOR" Then
50730    If Not reg.KeyExists Then
50740     reg.CreateKey
50750    End If
50760    reg.SetRegistryValue "StampFontColor", CStr(.StampFontColor), REG_SZ
50770    Set reg = Nothing
50780    Exit Sub
50790   End If
50800   If UCase$(OptionName) = "STAMPFONTNAME" Then
50810    If Not reg.KeyExists Then
50820     reg.CreateKey
50830    End If
50840    reg.SetRegistryValue "StampFontname", CStr(.StampFontname), REG_SZ
50850    Set reg = Nothing
50860    Exit Sub
50870   End If
50880   If UCase$(OptionName) = "STAMPFONTSIZE" Then
50890    If Not reg.KeyExists Then
50900     reg.CreateKey
50910    End If
50920    reg.SetRegistryValue "StampFontsize", CStr(.StampFontsize), REG_SZ
50930    Set reg = Nothing
50940    Exit Sub
50950   End If
50960   If UCase$(OptionName) = "STAMPOUTLINEFONTTHICKNESS" Then
50970    If Not reg.KeyExists Then
50980     reg.CreateKey
50990    End If
51000    reg.SetRegistryValue "StampOutlineFontthickness", CStr(.StampOutlineFontthickness), REG_SZ
51010    Set reg = Nothing
51020    Exit Sub
51030   End If
51040   If UCase$(OptionName) = "STAMPSTRING" Then
51050    If Not reg.KeyExists Then
51060     reg.CreateKey
51070    End If
51080    reg.SetRegistryValue "StampString", CStr(.StampString), REG_SZ
51090    Set reg = Nothing
51100    Exit Sub
51110   End If
51120   If UCase$(OptionName) = "STAMPUSEOUTLINEFONT" Then
51130    If Not reg.KeyExists Then
51140     reg.CreateKey
51150    End If
51160    reg.SetRegistryValue "StampUseOutlineFont", CStr(Abs(.StampUseOutlineFont)), REG_SZ
51170    Set reg = Nothing
51180    Exit Sub
51190   End If
51200   If UCase$(OptionName) = "STANDARDAUTHOR" Then
51210    If Not reg.KeyExists Then
51220     reg.CreateKey
51230    End If
51240    reg.SetRegistryValue "StandardAuthor", CStr(.StandardAuthor), REG_SZ
51250    Set reg = Nothing
51260    Exit Sub
51270   End If
51280   If UCase$(OptionName) = "STANDARDCREATIONDATE" Then
51290    If Not reg.KeyExists Then
51300     reg.CreateKey
51310    End If
51320    reg.SetRegistryValue "StandardCreationdate", CStr(.StandardCreationdate), REG_SZ
51330    Set reg = Nothing
51340    Exit Sub
51350   End If
51360   If UCase$(OptionName) = "STANDARDDATEFORMAT" Then
51370    If Not reg.KeyExists Then
51380     reg.CreateKey
51390    End If
51400    reg.SetRegistryValue "StandardDateformat", CStr(.StandardDateformat), REG_SZ
51410    Set reg = Nothing
51420    Exit Sub
51430   End If
51440   If UCase$(OptionName) = "STANDARDKEYWORDS" Then
51450    If Not reg.KeyExists Then
51460     reg.CreateKey
51470    End If
51480    reg.SetRegistryValue "StandardKeywords", CStr(.StandardKeywords), REG_SZ
51490    Set reg = Nothing
51500    Exit Sub
51510   End If
51520   If UCase$(OptionName) = "STANDARDMAILDOMAIN" Then
51530    If Not reg.KeyExists Then
51540     reg.CreateKey
51550    End If
51560    reg.SetRegistryValue "StandardMailDomain", CStr(.StandardMailDomain), REG_SZ
51570    Set reg = Nothing
51580    Exit Sub
51590   End If
51600   If UCase$(OptionName) = "STANDARDMODIFYDATE" Then
51610    If Not reg.KeyExists Then
51620     reg.CreateKey
51630    End If
51640    reg.SetRegistryValue "StandardModifydate", CStr(.StandardModifydate), REG_SZ
51650    Set reg = Nothing
51660    Exit Sub
51670   End If
51680   If UCase$(OptionName) = "STANDARDSAVEFORMAT" Then
51690    If Not reg.KeyExists Then
51700     reg.CreateKey
51710    End If
51720    reg.SetRegistryValue "StandardSaveformat", CStr(.StandardSaveformat), REG_SZ
51730    Set reg = Nothing
51740    Exit Sub
51750   End If
51760   If UCase$(OptionName) = "STANDARDSUBJECT" Then
51770    If Not reg.KeyExists Then
51780     reg.CreateKey
51790    End If
51800    reg.SetRegistryValue "StandardSubject", CStr(.StandardSubject), REG_SZ
51810    Set reg = Nothing
51820    Exit Sub
51830   End If
51840   If UCase$(OptionName) = "STANDARDTITLE" Then
51850    If Not reg.KeyExists Then
51860     reg.CreateKey
51870    End If
51880    reg.SetRegistryValue "StandardTitle", CStr(.StandardTitle), REG_SZ
51890    Set reg = Nothing
51900    Exit Sub
51910   End If
51920   If UCase$(OptionName) = "USECREATIONDATENOW" Then
51930    If Not reg.KeyExists Then
51940     reg.CreateKey
51950    End If
51960    reg.SetRegistryValue "UseCreationDateNow", CStr(Abs(.UseCreationDateNow)), REG_SZ
51970    Set reg = Nothing
51980    Exit Sub
51990   End If
52000   If UCase$(OptionName) = "USESTANDARDAUTHOR" Then
52010    If Not reg.KeyExists Then
52020     reg.CreateKey
52030    End If
52040    reg.SetRegistryValue "UseStandardAuthor", CStr(Abs(.UseStandardAuthor)), REG_SZ
52050    Set reg = Nothing
52060    Exit Sub
52070   End If
52080   reg.Subkey = "Printing\Formats\Bitmap\Colors"
52090   If UCase$(OptionName) = "BITMAPRESOLUTION" Then
52100    If Not reg.KeyExists Then
52110     reg.CreateKey
52120    End If
52130    reg.SetRegistryValue "BitmapResolution", CStr(.BitmapResolution), REG_SZ
52140    Set reg = Nothing
52150    Exit Sub
52160   End If
52170   If UCase$(OptionName) = "BMPCOLORSCOUNT" Then
52180    If Not reg.KeyExists Then
52190     reg.CreateKey
52200    End If
52210    reg.SetRegistryValue "BMPColorscount", CStr(.BMPColorscount), REG_SZ
52220    Set reg = Nothing
52230    Exit Sub
52240   End If
52250   If UCase$(OptionName) = "JPEGCOLORSCOUNT" Then
52260    If Not reg.KeyExists Then
52270     reg.CreateKey
52280    End If
52290    reg.SetRegistryValue "JPEGColorscount", CStr(.JPEGColorscount), REG_SZ
52300    Set reg = Nothing
52310    Exit Sub
52320   End If
52330   If UCase$(OptionName) = "JPEGQUALITY" Then
52340    If Not reg.KeyExists Then
52350     reg.CreateKey
52360    End If
52370    reg.SetRegistryValue "JPEGQuality", CStr(.JPEGQuality), REG_SZ
52380    Set reg = Nothing
52390    Exit Sub
52400   End If
52410   If UCase$(OptionName) = "PCXCOLORSCOUNT" Then
52420    If Not reg.KeyExists Then
52430     reg.CreateKey
52440    End If
52450    reg.SetRegistryValue "PCXColorscount", CStr(.PCXColorscount), REG_SZ
52460    Set reg = Nothing
52470    Exit Sub
52480   End If
52490   If UCase$(OptionName) = "PNGCOLORSCOUNT" Then
52500    If Not reg.KeyExists Then
52510     reg.CreateKey
52520    End If
52530    reg.SetRegistryValue "PNGColorscount", CStr(.PNGColorscount), REG_SZ
52540    Set reg = Nothing
52550    Exit Sub
52560   End If
52570   If UCase$(OptionName) = "TIFFCOLORSCOUNT" Then
52580    If Not reg.KeyExists Then
52590     reg.CreateKey
52600    End If
52610    reg.SetRegistryValue "TIFFColorscount", CStr(.TIFFColorscount), REG_SZ
52620    Set reg = Nothing
52630    Exit Sub
52640   End If
52650   reg.Subkey = "Printing\Formats\PDF\Colors"
52660   If UCase$(OptionName) = "PDFCOLORSCMYKTORGB" Then
52670    If Not reg.KeyExists Then
52680     reg.CreateKey
52690    End If
52700    reg.SetRegistryValue "PDFColorsCMYKToRGB", CStr(Abs(.PDFColorsCMYKToRGB)), REG_SZ
52710    Set reg = Nothing
52720    Exit Sub
52730   End If
52740   If UCase$(OptionName) = "PDFCOLORSCOLORMODEL" Then
52750    If Not reg.KeyExists Then
52760     reg.CreateKey
52770    End If
52780    reg.SetRegistryValue "PDFColorsColorModel", CStr(.PDFColorsColorModel), REG_SZ
52790    Set reg = Nothing
52800    Exit Sub
52810   End If
52820   If UCase$(OptionName) = "PDFCOLORSPRESERVEHALFTONE" Then
52830    If Not reg.KeyExists Then
52840     reg.CreateKey
52850    End If
52860    reg.SetRegistryValue "PDFColorsPreserveHalftone", CStr(Abs(.PDFColorsPreserveHalftone)), REG_SZ
52870    Set reg = Nothing
52880    Exit Sub
52890   End If
52900   If UCase$(OptionName) = "PDFCOLORSPRESERVEOVERPRINT" Then
52910    If Not reg.KeyExists Then
52920     reg.CreateKey
52930    End If
52940    reg.SetRegistryValue "PDFColorsPreserveOverprint", CStr(Abs(.PDFColorsPreserveOverprint)), REG_SZ
52950    Set reg = Nothing
52960    Exit Sub
52970   End If
52980   If UCase$(OptionName) = "PDFCOLORSPRESERVETRANSFER" Then
52990    If Not reg.KeyExists Then
53000     reg.CreateKey
53010    End If
53020    reg.SetRegistryValue "PDFColorsPreserveTransfer", CStr(Abs(.PDFColorsPreserveTransfer)), REG_SZ
53030    Set reg = Nothing
53040    Exit Sub
53050   End If
53060   reg.Subkey = "Printing\Formats\PDF\Compression"
53070   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSION" Then
53080    If Not reg.KeyExists Then
53090     reg.CreateKey
53100    End If
53110    reg.SetRegistryValue "PDFCompressionColorCompression", CStr(Abs(.PDFCompressionColorCompression)), REG_SZ
53120    Set reg = Nothing
53130    Exit Sub
53140   End If
53150   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE" Then
53160    If Not reg.KeyExists Then
53170     reg.CreateKey
53180    End If
53190    reg.SetRegistryValue "PDFCompressionColorCompressionChoice", CStr(.PDFCompressionColorCompressionChoice), REG_SZ
53200    Set reg = Nothing
53210    Exit Sub
53220   End If
53230   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR" Then
53240    If Not reg.KeyExists Then
53250     reg.CreateKey
53260    End If
53270   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
53280    Set reg = Nothing
53290    Exit Sub
53300   End If
53310   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR" Then
53320    If Not reg.KeyExists Then
53330     reg.CreateKey
53340    End If
53350   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
53360    Set reg = Nothing
53370    Exit Sub
53380   End If
53390   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR" Then
53400    If Not reg.KeyExists Then
53410     reg.CreateKey
53420    End If
53430   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
53440    Set reg = Nothing
53450    Exit Sub
53460   End If
53470   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR" Then
53480    If Not reg.KeyExists Then
53490     reg.CreateKey
53500    End If
53510   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
53520    Set reg = Nothing
53530    Exit Sub
53540   End If
53550   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR" Then
53560    If Not reg.KeyExists Then
53570     reg.CreateKey
53580    End If
53590   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
53600    Set reg = Nothing
53610    Exit Sub
53620   End If
53630   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESAMPLE" Then
53640    If Not reg.KeyExists Then
53650     reg.CreateKey
53660    End If
53670    reg.SetRegistryValue "PDFCompressionColorResample", CStr(Abs(.PDFCompressionColorResample)), REG_SZ
53680    Set reg = Nothing
53690    Exit Sub
53700   End If
53710   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESAMPLECHOICE" Then
53720    If Not reg.KeyExists Then
53730     reg.CreateKey
53740    End If
53750    reg.SetRegistryValue "PDFCompressionColorResampleChoice", CStr(.PDFCompressionColorResampleChoice), REG_SZ
53760    Set reg = Nothing
53770    Exit Sub
53780   End If
53790   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESOLUTION" Then
53800    If Not reg.KeyExists Then
53810     reg.CreateKey
53820    End If
53830    reg.SetRegistryValue "PDFCompressionColorResolution", CStr(.PDFCompressionColorResolution), REG_SZ
53840    Set reg = Nothing
53850    Exit Sub
53860   End If
53870   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSION" Then
53880    If Not reg.KeyExists Then
53890     reg.CreateKey
53900    End If
53910    reg.SetRegistryValue "PDFCompressionGreyCompression", CStr(Abs(.PDFCompressionGreyCompression)), REG_SZ
53920    Set reg = Nothing
53930    Exit Sub
53940   End If
53950   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE" Then
53960    If Not reg.KeyExists Then
53970     reg.CreateKey
53980    End If
53990    reg.SetRegistryValue "PDFCompressionGreyCompressionChoice", CStr(.PDFCompressionGreyCompressionChoice), REG_SZ
54000    Set reg = Nothing
54010    Exit Sub
54020   End If
54030   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR" Then
54040    If Not reg.KeyExists Then
54050     reg.CreateKey
54060    End If
54070   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
54080    Set reg = Nothing
54090    Exit Sub
54100   End If
54110   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR" Then
54120    If Not reg.KeyExists Then
54130     reg.CreateKey
54140    End If
54150   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
54160    Set reg = Nothing
54170    Exit Sub
54180   End If
54190   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR" Then
54200    If Not reg.KeyExists Then
54210     reg.CreateKey
54220    End If
54230   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
54240    Set reg = Nothing
54250    Exit Sub
54260   End If
54270   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR" Then
54280    If Not reg.KeyExists Then
54290     reg.CreateKey
54300    End If
54310   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
54320    Set reg = Nothing
54330    Exit Sub
54340   End If
54350   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR" Then
54360    If Not reg.KeyExists Then
54370     reg.CreateKey
54380    End If
54390   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
54400    Set reg = Nothing
54410    Exit Sub
54420   End If
54430   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESAMPLE" Then
54440    If Not reg.KeyExists Then
54450     reg.CreateKey
54460    End If
54470    reg.SetRegistryValue "PDFCompressionGreyResample", CStr(Abs(.PDFCompressionGreyResample)), REG_SZ
54480    Set reg = Nothing
54490    Exit Sub
54500   End If
54510   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESAMPLECHOICE" Then
54520    If Not reg.KeyExists Then
54530     reg.CreateKey
54540    End If
54550    reg.SetRegistryValue "PDFCompressionGreyResampleChoice", CStr(.PDFCompressionGreyResampleChoice), REG_SZ
54560    Set reg = Nothing
54570    Exit Sub
54580   End If
54590   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESOLUTION" Then
54600    If Not reg.KeyExists Then
54610     reg.CreateKey
54620    End If
54630    reg.SetRegistryValue "PDFCompressionGreyResolution", CStr(.PDFCompressionGreyResolution), REG_SZ
54640    Set reg = Nothing
54650    Exit Sub
54660   End If
54670   If UCase$(OptionName) = "PDFCOMPRESSIONMONOCOMPRESSION" Then
54680    If Not reg.KeyExists Then
54690     reg.CreateKey
54700    End If
54710    reg.SetRegistryValue "PDFCompressionMonoCompression", CStr(Abs(.PDFCompressionMonoCompression)), REG_SZ
54720    Set reg = Nothing
54730    Exit Sub
54740   End If
54750   If UCase$(OptionName) = "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE" Then
54760    If Not reg.KeyExists Then
54770     reg.CreateKey
54780    End If
54790    reg.SetRegistryValue "PDFCompressionMonoCompressionChoice", CStr(.PDFCompressionMonoCompressionChoice), REG_SZ
54800    Set reg = Nothing
54810    Exit Sub
54820   End If
54830   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESAMPLE" Then
54840    If Not reg.KeyExists Then
54850     reg.CreateKey
54860    End If
54870    reg.SetRegistryValue "PDFCompressionMonoResample", CStr(Abs(.PDFCompressionMonoResample)), REG_SZ
54880    Set reg = Nothing
54890    Exit Sub
54900   End If
54910   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESAMPLECHOICE" Then
54920    If Not reg.KeyExists Then
54930     reg.CreateKey
54940    End If
54950    reg.SetRegistryValue "PDFCompressionMonoResampleChoice", CStr(.PDFCompressionMonoResampleChoice), REG_SZ
54960    Set reg = Nothing
54970    Exit Sub
54980   End If
54990   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESOLUTION" Then
55000    If Not reg.KeyExists Then
55010     reg.CreateKey
55020    End If
55030    reg.SetRegistryValue "PDFCompressionMonoResolution", CStr(.PDFCompressionMonoResolution), REG_SZ
55040    Set reg = Nothing
55050    Exit Sub
55060   End If
55070   If UCase$(OptionName) = "PDFCOMPRESSIONTEXTCOMPRESSION" Then
55080    If Not reg.KeyExists Then
55090     reg.CreateKey
55100    End If
55110    reg.SetRegistryValue "PDFCompressionTextCompression", CStr(Abs(.PDFCompressionTextCompression)), REG_SZ
55120    Set reg = Nothing
55130    Exit Sub
55140   End If
55150   reg.Subkey = "Printing\Formats\PDF\Fonts"
55160   If UCase$(OptionName) = "PDFFONTSEMBEDALL" Then
55170    If Not reg.KeyExists Then
55180     reg.CreateKey
55190    End If
55200    reg.SetRegistryValue "PDFFontsEmbedAll", CStr(Abs(.PDFFontsEmbedAll)), REG_SZ
55210    Set reg = Nothing
55220    Exit Sub
55230   End If
55240   If UCase$(OptionName) = "PDFFONTSSUBSETFONTS" Then
55250    If Not reg.KeyExists Then
55260     reg.CreateKey
55270    End If
55280    reg.SetRegistryValue "PDFFontsSubSetFonts", CStr(Abs(.PDFFontsSubSetFonts)), REG_SZ
55290    Set reg = Nothing
55300    Exit Sub
55310   End If
55320   If UCase$(OptionName) = "PDFFONTSSUBSETFONTSPERCENT" Then
55330    If Not reg.KeyExists Then
55340     reg.CreateKey
55350    End If
55360    reg.SetRegistryValue "PDFFontsSubSetFontsPercent", CStr(.PDFFontsSubSetFontsPercent), REG_SZ
55370    Set reg = Nothing
55380    Exit Sub
55390   End If
55400   reg.Subkey = "Printing\Formats\PDF\General"
55410   If UCase$(OptionName) = "PDFGENERALASCII85" Then
55420    If Not reg.KeyExists Then
55430     reg.CreateKey
55440    End If
55450    reg.SetRegistryValue "PDFGeneralASCII85", CStr(Abs(.PDFGeneralASCII85)), REG_SZ
55460    Set reg = Nothing
55470    Exit Sub
55480   End If
55490   If UCase$(OptionName) = "PDFGENERALAUTOROTATE" Then
55500    If Not reg.KeyExists Then
55510     reg.CreateKey
55520    End If
55530    reg.SetRegistryValue "PDFGeneralAutorotate", CStr(.PDFGeneralAutorotate), REG_SZ
55540    Set reg = Nothing
55550    Exit Sub
55560   End If
55570   If UCase$(OptionName) = "PDFGENERALCOMPATIBILITY" Then
55580    If Not reg.KeyExists Then
55590     reg.CreateKey
55600    End If
55610    reg.SetRegistryValue "PDFGeneralCompatibility", CStr(.PDFGeneralCompatibility), REG_SZ
55620    Set reg = Nothing
55630    Exit Sub
55640   End If
55650   If UCase$(OptionName) = "PDFGENERALOVERPRINT" Then
55660    If Not reg.KeyExists Then
55670     reg.CreateKey
55680    End If
55690    reg.SetRegistryValue "PDFGeneralOverprint", CStr(.PDFGeneralOverprint), REG_SZ
55700    Set reg = Nothing
55710    Exit Sub
55720   End If
55730   If UCase$(OptionName) = "PDFGENERALRESOLUTION" Then
55740    If Not reg.KeyExists Then
55750     reg.CreateKey
55760    End If
55770    reg.SetRegistryValue "PDFGeneralResolution", CStr(.PDFGeneralResolution), REG_SZ
55780    Set reg = Nothing
55790    Exit Sub
55800   End If
55810   If UCase$(OptionName) = "PDFOPTIMIZE" Then
55820    If Not reg.KeyExists Then
55830     reg.CreateKey
55840    End If
55850    reg.SetRegistryValue "PDFOptimize", CStr(Abs(.PDFOptimize)), REG_SZ
55860    Set reg = Nothing
55870    Exit Sub
55880   End If
55890   reg.Subkey = "Printing\Formats\PDF\Security"
55900   If UCase$(OptionName) = "PDFALLOWASSEMBLY" Then
55910    If Not reg.KeyExists Then
55920     reg.CreateKey
55930    End If
55940    reg.SetRegistryValue "PDFAllowAssembly", CStr(Abs(.PDFAllowAssembly)), REG_SZ
55950    Set reg = Nothing
55960    Exit Sub
55970   End If
55980   If UCase$(OptionName) = "PDFALLOWDEGRADEDPRINTING" Then
55990    If Not reg.KeyExists Then
56000     reg.CreateKey
56010    End If
56020    reg.SetRegistryValue "PDFAllowDegradedPrinting", CStr(Abs(.PDFAllowDegradedPrinting)), REG_SZ
56030    Set reg = Nothing
56040    Exit Sub
56050   End If
56060   If UCase$(OptionName) = "PDFALLOWFILLIN" Then
56070    If Not reg.KeyExists Then
56080     reg.CreateKey
56090    End If
56100    reg.SetRegistryValue "PDFAllowFillIn", CStr(Abs(.PDFAllowFillIn)), REG_SZ
56110    Set reg = Nothing
56120    Exit Sub
56130   End If
56140   If UCase$(OptionName) = "PDFALLOWSCREENREADERS" Then
56150    If Not reg.KeyExists Then
56160     reg.CreateKey
56170    End If
56180    reg.SetRegistryValue "PDFAllowScreenReaders", CStr(Abs(.PDFAllowScreenReaders)), REG_SZ
56190    Set reg = Nothing
56200    Exit Sub
56210   End If
56220   If UCase$(OptionName) = "PDFDISALLOWCOPY" Then
56230    If Not reg.KeyExists Then
56240     reg.CreateKey
56250    End If
56260    reg.SetRegistryValue "PDFDisallowCopy", CStr(Abs(.PDFDisallowCopy)), REG_SZ
56270    Set reg = Nothing
56280    Exit Sub
56290   End If
56300   If UCase$(OptionName) = "PDFDISALLOWMODIFYANNOTATIONS" Then
56310    If Not reg.KeyExists Then
56320     reg.CreateKey
56330    End If
56340    reg.SetRegistryValue "PDFDisallowModifyAnnotations", CStr(Abs(.PDFDisallowModifyAnnotations)), REG_SZ
56350    Set reg = Nothing
56360    Exit Sub
56370   End If
56380   If UCase$(OptionName) = "PDFDISALLOWMODIFYCONTENTS" Then
56390    If Not reg.KeyExists Then
56400     reg.CreateKey
56410    End If
56420    reg.SetRegistryValue "PDFDisallowModifyContents", CStr(Abs(.PDFDisallowModifyContents)), REG_SZ
56430    Set reg = Nothing
56440    Exit Sub
56450   End If
56460   If UCase$(OptionName) = "PDFDISALLOWPRINTING" Then
56470    If Not reg.KeyExists Then
56480     reg.CreateKey
56490    End If
56500    reg.SetRegistryValue "PDFDisallowPrinting", CStr(Abs(.PDFDisallowPrinting)), REG_SZ
56510    Set reg = Nothing
56520    Exit Sub
56530   End If
56540   If UCase$(OptionName) = "PDFENCRYPTOR" Then
56550    If Not reg.KeyExists Then
56560     reg.CreateKey
56570    End If
56580    reg.SetRegistryValue "PDFEncryptor", CStr(.PDFEncryptor), REG_SZ
56590    Set reg = Nothing
56600    Exit Sub
56610   End If
56620   If UCase$(OptionName) = "PDFHIGHENCRYPTION" Then
56630    If Not reg.KeyExists Then
56640     reg.CreateKey
56650    End If
56660    reg.SetRegistryValue "PDFHighEncryption", CStr(Abs(.PDFHighEncryption)), REG_SZ
56670    Set reg = Nothing
56680    Exit Sub
56690   End If
56700   If UCase$(OptionName) = "PDFLOWENCRYPTION" Then
56710    If Not reg.KeyExists Then
56720     reg.CreateKey
56730    End If
56740    reg.SetRegistryValue "PDFLowEncryption", CStr(Abs(.PDFLowEncryption)), REG_SZ
56750    Set reg = Nothing
56760    Exit Sub
56770   End If
56780   If UCase$(OptionName) = "PDFOWNERPASS" Then
56790    If Not reg.KeyExists Then
56800     reg.CreateKey
56810    End If
56820    reg.SetRegistryValue "PDFOwnerPass", CStr(Abs(.PDFOwnerPass)), REG_SZ
56830    Set reg = Nothing
56840    Exit Sub
56850   End If
56860   If UCase$(OptionName) = "PDFOWNERPASSWORDSTRING" Then
56870    If Not reg.KeyExists Then
56880     reg.CreateKey
56890    End If
56900    reg.SetRegistryValue "PDFOwnerPasswordString", CStr(.PDFOwnerPasswordString), REG_SZ
56910    Set reg = Nothing
56920    Exit Sub
56930   End If
56940   If UCase$(OptionName) = "PDFUSERPASS" Then
56950    If Not reg.KeyExists Then
56960     reg.CreateKey
56970    End If
56980    reg.SetRegistryValue "PDFUserPass", CStr(Abs(.PDFUserPass)), REG_SZ
56990    Set reg = Nothing
57000    Exit Sub
57010   End If
57020   If UCase$(OptionName) = "PDFUSERPASSWORDSTRING" Then
57030    If Not reg.KeyExists Then
57040     reg.CreateKey
57050    End If
57060    reg.SetRegistryValue "PDFUserPasswordString", CStr(.PDFUserPasswordString), REG_SZ
57070    Set reg = Nothing
57080    Exit Sub
57090   End If
57100   If UCase$(OptionName) = "PDFUSESECURITY" Then
57110    If Not reg.KeyExists Then
57120     reg.CreateKey
57130    End If
57140    reg.SetRegistryValue "PDFUseSecurity", CStr(Abs(.PDFUseSecurity)), REG_SZ
57150    Set reg = Nothing
57160    Exit Sub
57170   End If
57180   reg.Subkey = "Printing\Formats\PS\LanguageLevel"
57190   If UCase$(OptionName) = "EPSLANGUAGELEVEL" Then
57200    If Not reg.KeyExists Then
57210     reg.CreateKey
57220    End If
57230    reg.SetRegistryValue "EPSLanguageLevel", CStr(.EPSLanguageLevel), REG_SZ
57240    Set reg = Nothing
57250    Exit Sub
57260   End If
57270   If UCase$(OptionName) = "PSLANGUAGELEVEL" Then
57280    If Not reg.KeyExists Then
57290     reg.CreateKey
57300    End If
57310    reg.SetRegistryValue "PSLanguageLevel", CStr(.PSLanguageLevel), REG_SZ
57320    Set reg = Nothing
57330    Exit Sub
57340   End If
57350   reg.Subkey = "Program"
57360   If UCase$(OptionName) = "ADDITIONALGHOSTSCRIPTPARAMETERS" Then
57370    If Not reg.KeyExists Then
57380     reg.CreateKey
57390    End If
57400    reg.SetRegistryValue "AdditionalGhostscriptParameters", CStr(.AdditionalGhostscriptParameters), REG_SZ
57410    Set reg = Nothing
57420    Exit Sub
57430   End If
57440   If UCase$(OptionName) = "ADDITIONALGHOSTSCRIPTSEARCHPATH" Then
57450    If Not reg.KeyExists Then
57460     reg.CreateKey
57470    End If
57480    reg.SetRegistryValue "AdditionalGhostscriptSearchpath", CStr(.AdditionalGhostscriptSearchpath), REG_SZ
57490    Set reg = Nothing
57500    Exit Sub
57510   End If
57520   If UCase$(OptionName) = "ADDWINDOWSFONTPATH" Then
57530    If Not reg.KeyExists Then
57540     reg.CreateKey
57550    End If
57560    reg.SetRegistryValue "AddWindowsFontpath", CStr(Abs(.AddWindowsFontpath)), REG_SZ
57570    Set reg = Nothing
57580    Exit Sub
57590   End If
57600   If UCase$(OptionName) = "AUTOSAVEDIRECTORY" Then
57610    If Not reg.KeyExists Then
57620     reg.CreateKey
57630    End If
57640    reg.SetRegistryValue "AutosaveDirectory", CStr(.AutosaveDirectory), REG_SZ
57650    Set reg = Nothing
57660    Exit Sub
57670   End If
57680   If UCase$(OptionName) = "AUTOSAVEFILENAME" Then
57690    If Not reg.KeyExists Then
57700     reg.CreateKey
57710    End If
57720    reg.SetRegistryValue "AutosaveFilename", CStr(.AutosaveFilename), REG_SZ
57730    Set reg = Nothing
57740    Exit Sub
57750   End If
57760   If UCase$(OptionName) = "AUTOSAVEFORMAT" Then
57770    If Not reg.KeyExists Then
57780     reg.CreateKey
57790    End If
57800    reg.SetRegistryValue "AutosaveFormat", CStr(.AutosaveFormat), REG_SZ
57810    Set reg = Nothing
57820    Exit Sub
57830   End If
57840   If UCase$(OptionName) = "AUTOSAVESTARTSTANDARDPROGRAM" Then
57850    If Not reg.KeyExists Then
57860     reg.CreateKey
57870    End If
57880    reg.SetRegistryValue "AutosaveStartStandardProgram", CStr(Abs(.AutosaveStartStandardProgram)), REG_SZ
57890    Set reg = Nothing
57900    Exit Sub
57910   End If
57920   If UCase$(OptionName) = "CLIENTCOMPUTERRESOLVEIPADDRESS" Then
57930    If Not reg.KeyExists Then
57940     reg.CreateKey
57950    End If
57960    reg.SetRegistryValue "ClientComputerResolveIPAddress", CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
57970    Set reg = Nothing
57980    Exit Sub
57990   End If
58000   If UCase$(OptionName) = "DISABLEEMAIL" Then
58010    If Not reg.KeyExists Then
58020     reg.CreateKey
58030    End If
58040    reg.SetRegistryValue "DisableEmail", CStr(Abs(.DisableEmail)), REG_SZ
58050    Set reg = Nothing
58060    Exit Sub
58070   End If
58080   If UCase$(OptionName) = "DONTUSEDOCUMENTSETTINGS" Then
58090    If Not reg.KeyExists Then
58100     reg.CreateKey
58110    End If
58120    reg.SetRegistryValue "DontUseDocumentSettings", CStr(Abs(.DontUseDocumentSettings)), REG_SZ
58130    Set reg = Nothing
58140    Exit Sub
58150   End If
58160   If UCase$(OptionName) = "FILENAMESUBSTITUTIONS" Then
58170    If Not reg.KeyExists Then
58180     reg.CreateKey
58190    End If
58200    reg.SetRegistryValue "FilenameSubstitutions", CStr(.FilenameSubstitutions), REG_SZ
58210    Set reg = Nothing
58220    Exit Sub
58230   End If
58240   If UCase$(OptionName) = "FILENAMESUBSTITUTIONSONLYINTITLE" Then
58250    If Not reg.KeyExists Then
58260     reg.CreateKey
58270    End If
58280    reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle", CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
58290    Set reg = Nothing
58300    Exit Sub
58310   End If
58320   If UCase$(OptionName) = "LANGUAGE" Then
58330    If Not reg.KeyExists Then
58340     reg.CreateKey
58350    End If
58360    reg.SetRegistryValue "Language", CStr(.Language), REG_SZ
58370    Set reg = Nothing
58380    Exit Sub
58390   End If
58400   If UCase$(OptionName) = "LASTSAVEDIRECTORY" Then
58410    If Not reg.KeyExists Then
58420     reg.CreateKey
58430    End If
58440    reg.SetRegistryValue "LastSaveDirectory", CStr(.LastSaveDirectory), REG_SZ
58450    Set reg = Nothing
58460    Exit Sub
58470   End If
58480   If UCase$(OptionName) = "LOGGING" Then
58490    If Not reg.KeyExists Then
58500     reg.CreateKey
58510    End If
58520    reg.SetRegistryValue "Logging", CStr(Abs(.Logging)), REG_SZ
58530    Set reg = Nothing
58540    Exit Sub
58550   End If
58560   If UCase$(OptionName) = "LOGLINES" Then
58570    If Not reg.KeyExists Then
58580     reg.CreateKey
58590    End If
58600    reg.SetRegistryValue "LogLines", CStr(.LogLines), REG_SZ
58610    Set reg = Nothing
58620    Exit Sub
58630   End If
58640   If UCase$(OptionName) = "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER" Then
58650    If Not reg.KeyExists Then
58660     reg.CreateKey
58670    End If
58680    reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter", CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
58690    Set reg = Nothing
58700    Exit Sub
58710   End If
58720   If UCase$(OptionName) = "NOPROCESSINGATSTARTUP" Then
58730    If Not reg.KeyExists Then
58740     reg.CreateKey
58750    End If
58760    reg.SetRegistryValue "NoProcessingAtStartup", CStr(Abs(.NoProcessingAtStartup)), REG_SZ
58770    Set reg = Nothing
58780    Exit Sub
58790   End If
58800   If UCase$(OptionName) = "NOPSCHECK" Then
58810    If Not reg.KeyExists Then
58820     reg.CreateKey
58830    End If
58840    reg.SetRegistryValue "NoPSCheck", CStr(Abs(.NoPSCheck)), REG_SZ
58850    Set reg = Nothing
58860    Exit Sub
58870   End If
58880   If UCase$(OptionName) = "OPTIONSDESIGN" Then
58890    If Not reg.KeyExists Then
58900     reg.CreateKey
58910    End If
58920    reg.SetRegistryValue "OptionsDesign", CStr(.OptionsDesign), REG_SZ
58930    Set reg = Nothing
58940    Exit Sub
58950   End If
58960   If UCase$(OptionName) = "OPTIONSENABLED" Then
58970    If Not reg.KeyExists Then
58980     reg.CreateKey
58990    End If
59000    reg.SetRegistryValue "OptionsEnabled", CStr(Abs(.OptionsEnabled)), REG_SZ
59010    Set reg = Nothing
59020    Exit Sub
59030   End If
59040   If UCase$(OptionName) = "OPTIONSVISIBLE" Then
59050    If Not reg.KeyExists Then
59060     reg.CreateKey
59070    End If
59080    reg.SetRegistryValue "OptionsVisible", CStr(Abs(.OptionsVisible)), REG_SZ
59090    Set reg = Nothing
59100    Exit Sub
59110   End If
59120   If UCase$(OptionName) = "PRINTAFTERSAVING" Then
59130    If Not reg.KeyExists Then
59140     reg.CreateKey
59150    End If
59160    reg.SetRegistryValue "PrintAfterSaving", CStr(Abs(.PrintAfterSaving)), REG_SZ
59170    Set reg = Nothing
59180    Exit Sub
59190   End If
59200   If UCase$(OptionName) = "PRINTAFTERSAVINGDUPLEX" Then
59210    If Not reg.KeyExists Then
59220     reg.CreateKey
59230    End If
59240    reg.SetRegistryValue "PrintAfterSavingDuplex", CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
59250    Set reg = Nothing
59260    Exit Sub
59270   End If
59280   If UCase$(OptionName) = "PRINTAFTERSAVINGNOCANCEL" Then
59290    If Not reg.KeyExists Then
59300     reg.CreateKey
59310    End If
59320    reg.SetRegistryValue "PrintAfterSavingNoCancel", CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
59330    Set reg = Nothing
59340    Exit Sub
59350   End If
59360   If UCase$(OptionName) = "PRINTAFTERSAVINGPRINTER" Then
59370    If Not reg.KeyExists Then
59380     reg.CreateKey
59390    End If
59400    reg.SetRegistryValue "PrintAfterSavingPrinter", CStr(.PrintAfterSavingPrinter), REG_SZ
59410    Set reg = Nothing
59420    Exit Sub
59430   End If
59440   If UCase$(OptionName) = "PRINTAFTERSAVINGQUERYUSER" Then
59450    If Not reg.KeyExists Then
59460     reg.CreateKey
59470    End If
59480    reg.SetRegistryValue "PrintAfterSavingQueryUser", CStr(.PrintAfterSavingQueryUser), REG_SZ
59490    Set reg = Nothing
59500    Exit Sub
59510   End If
59520   If UCase$(OptionName) = "PRINTAFTERSAVINGTUMBLE" Then
59530    If Not reg.KeyExists Then
59540     reg.CreateKey
59550    End If
59560    reg.SetRegistryValue "PrintAfterSavingTumble", CStr(.PrintAfterSavingTumble), REG_SZ
59570    Set reg = Nothing
59580    Exit Sub
59590   End If
59600   If UCase$(OptionName) = "PRINTERSTOP" Then
59610    If Not reg.KeyExists Then
59620     reg.CreateKey
59630    End If
59640    reg.SetRegistryValue "PrinterStop", CStr(Abs(.PrinterStop)), REG_SZ
59650    Set reg = Nothing
59660    Exit Sub
59670   End If
59680   If UCase$(OptionName) = "PRINTERTEMPPATH" Then
59690    If Not reg.KeyExists Then
59700     reg.CreateKey
59710    End If
59720    reg.SetRegistryValue "PrinterTemppath", CStr(.PrinterTemppath), REG_SZ
59730    Set reg = Nothing
59740    Exit Sub
59750   End If
59760   If UCase$(OptionName) = "PROCESSPRIORITY" Then
59770    If Not reg.KeyExists Then
59780     reg.CreateKey
59790    End If
59800    reg.SetRegistryValue "ProcessPriority", CStr(.ProcessPriority), REG_SZ
59810    Set reg = Nothing
59820    Exit Sub
59830   End If
59840   If UCase$(OptionName) = "PROGRAMFONT" Then
59850    If Not reg.KeyExists Then
59860     reg.CreateKey
59870    End If
59880    reg.SetRegistryValue "ProgramFont", CStr(.ProgramFont), REG_SZ
59890    Set reg = Nothing
59900    Exit Sub
59910   End If
59920   If UCase$(OptionName) = "PROGRAMFONTCHARSET" Then
59930    If Not reg.KeyExists Then
59940     reg.CreateKey
59950    End If
59960    reg.SetRegistryValue "ProgramFontCharset", CStr(.ProgramFontCharset), REG_SZ
59970    Set reg = Nothing
59980    Exit Sub
59990   End If
60000   If UCase$(OptionName) = "PROGRAMFONTSIZE" Then
60010    If Not reg.KeyExists Then
60020     reg.CreateKey
60030    End If
60040    reg.SetRegistryValue "ProgramFontSize", CStr(.ProgramFontSize), REG_SZ
60050    Set reg = Nothing
60060    Exit Sub
60070   End If
60080   If UCase$(OptionName) = "REMOVEALLKNOWNFILEEXTENSIONS" Then
60090    If Not reg.KeyExists Then
60100     reg.CreateKey
60110    End If
60120    reg.SetRegistryValue "RemoveAllKnownFileExtensions", CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
60130    Set reg = Nothing
60140    Exit Sub
60150   End If
60160   If UCase$(OptionName) = "REMOVESPACES" Then
60170    If Not reg.KeyExists Then
60180     reg.CreateKey
60190    End If
60200    reg.SetRegistryValue "RemoveSpaces", CStr(Abs(.RemoveSpaces)), REG_SZ
60210    Set reg = Nothing
60220    Exit Sub
60230   End If
60240   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVING" Then
60250    If Not reg.KeyExists Then
60260     reg.CreateKey
60270    End If
60280    reg.SetRegistryValue "RunProgramAfterSaving", CStr(Abs(.RunProgramAfterSaving)), REG_SZ
60290    Set reg = Nothing
60300    Exit Sub
60310   End If
60320   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGPROGRAMNAME" Then
60330    If Not reg.KeyExists Then
60340     reg.CreateKey
60350    End If
60360    reg.SetRegistryValue "RunProgramAfterSavingProgramname", CStr(.RunProgramAfterSavingProgramname), REG_SZ
60370    Set reg = Nothing
60380    Exit Sub
60390   End If
60400   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS" Then
60410    If Not reg.KeyExists Then
60420     reg.CreateKey
60430    End If
60440    reg.SetRegistryValue "RunProgramAfterSavingProgramParameters", CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
60450    Set reg = Nothing
60460    Exit Sub
60470   End If
60480   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGWAITUNTILREADY" Then
60490    If Not reg.KeyExists Then
60500     reg.CreateKey
60510    End If
60520    reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady", CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
60530    Set reg = Nothing
60540    Exit Sub
60550   End If
60560   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGWINDOWSTYLE" Then
60570    If Not reg.KeyExists Then
60580     reg.CreateKey
60590    End If
60600    reg.SetRegistryValue "RunProgramAfterSavingWindowstyle", CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
60610    Set reg = Nothing
60620    Exit Sub
60630   End If
60640   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVING" Then
60650    If Not reg.KeyExists Then
60660     reg.CreateKey
60670    End If
60680    reg.SetRegistryValue "RunProgramBeforeSaving", CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
60690    Set reg = Nothing
60700    Exit Sub
60710   End If
60720   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGPROGRAMNAME" Then
60730    If Not reg.KeyExists Then
60740     reg.CreateKey
60750    End If
60760    reg.SetRegistryValue "RunProgramBeforeSavingProgramname", CStr(.RunProgramBeforeSavingProgramname), REG_SZ
60770    Set reg = Nothing
60780    Exit Sub
60790   End If
60800   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS" Then
60810    If Not reg.KeyExists Then
60820     reg.CreateKey
60830    End If
60840    reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters", CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
60850    Set reg = Nothing
60860    Exit Sub
60870   End If
60880   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGWINDOWSTYLE" Then
60890    If Not reg.KeyExists Then
60900     reg.CreateKey
60910    End If
60920    reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle", CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
60930    Set reg = Nothing
60940    Exit Sub
60950   End If
60960   If UCase$(OptionName) = "SAVEFILENAME" Then
60970    If Not reg.KeyExists Then
60980     reg.CreateKey
60990    End If
61000    reg.SetRegistryValue "SaveFilename", CStr(.SaveFilename), REG_SZ
61010    Set reg = Nothing
61020    Exit Sub
61030   End If
61040   If UCase$(OptionName) = "SENDEMAILAFTERAUTOSAVING" Then
61050    If Not reg.KeyExists Then
61060     reg.CreateKey
61070    End If
61080    reg.SetRegistryValue "SendEmailAfterAutoSaving", CStr(Abs(.SendEmailAfterAutoSaving)), REG_SZ
61090    Set reg = Nothing
61100    Exit Sub
61110   End If
61120   If UCase$(OptionName) = "SENDMAILMETHOD" Then
61130    If Not reg.KeyExists Then
61140     reg.CreateKey
61150    End If
61160    reg.SetRegistryValue "SendMailMethod", CStr(.SendMailMethod), REG_SZ
61170    Set reg = Nothing
61180    Exit Sub
61190   End If
61200   If UCase$(OptionName) = "SHOWANIMATION" Then
61210    If Not reg.KeyExists Then
61220     reg.CreateKey
61230    End If
61240    reg.SetRegistryValue "ShowAnimation", CStr(Abs(.ShowAnimation)), REG_SZ
61250    Set reg = Nothing
61260    Exit Sub
61270   End If
61280   If UCase$(OptionName) = "STARTSTANDARDPROGRAM" Then
61290    If Not reg.KeyExists Then
61300     reg.CreateKey
61310    End If
61320    reg.SetRegistryValue "StartStandardProgram", CStr(Abs(.StartStandardProgram)), REG_SZ
61330    Set reg = Nothing
61340    Exit Sub
61350   End If
61360   If UCase$(OptionName) = "TOOLBARS" Then
61370    If Not reg.KeyExists Then
61380     reg.CreateKey
61390    End If
61400    reg.SetRegistryValue "Toolbars", CStr(.Toolbars), REG_SZ
61410    Set reg = Nothing
61420    Exit Sub
61430   End If
61440   If UCase$(OptionName) = "USEAUTOSAVE" Then
61450    If Not reg.KeyExists Then
61460     reg.CreateKey
61470    End If
61480    reg.SetRegistryValue "UseAutosave", CStr(Abs(.UseAutosave)), REG_SZ
61490    Set reg = Nothing
61500    Exit Sub
61510   End If
61520   If UCase$(OptionName) = "USEAUTOSAVEDIRECTORY" Then
61530    If Not reg.KeyExists Then
61540     reg.CreateKey
61550    End If
61560    reg.SetRegistryValue "UseAutosaveDirectory", CStr(Abs(.UseAutosaveDirectory)), REG_SZ
61570    Set reg = Nothing
61580    Exit Sub
61590   End If
61600  End With
61610  Set reg = Nothing
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
50210   reg.SetRegistryValue "DeviceHeightPoints", Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), REG_SZ
50220   reg.SetRegistryValue "DeviceWidthPoints", Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), REG_SZ
50230   reg.SetRegistryValue "OnePagePerFile", CStr(Abs(.OnePagePerFile)), REG_SZ
50240   reg.SetRegistryValue "Papersize", CStr(.Papersize), REG_SZ
50250   reg.SetRegistryValue "StampFontColor", CStr(.StampFontColor), REG_SZ
50260   reg.SetRegistryValue "StampFontname", CStr(.StampFontname), REG_SZ
50270   reg.SetRegistryValue "StampFontsize", CStr(.StampFontsize), REG_SZ
50280   reg.SetRegistryValue "StampOutlineFontthickness", CStr(.StampOutlineFontthickness), REG_SZ
50290   reg.SetRegistryValue "StampString", CStr(.StampString), REG_SZ
50300   reg.SetRegistryValue "StampUseOutlineFont", CStr(Abs(.StampUseOutlineFont)), REG_SZ
50310   reg.SetRegistryValue "StandardAuthor", CStr(.StandardAuthor), REG_SZ
50320   reg.SetRegistryValue "StandardCreationdate", CStr(.StandardCreationdate), REG_SZ
50330   reg.SetRegistryValue "StandardDateformat", CStr(.StandardDateformat), REG_SZ
50340   reg.SetRegistryValue "StandardKeywords", CStr(.StandardKeywords), REG_SZ
50350   reg.SetRegistryValue "StandardMailDomain", CStr(.StandardMailDomain), REG_SZ
50360   reg.SetRegistryValue "StandardModifydate", CStr(.StandardModifydate), REG_SZ
50370   reg.SetRegistryValue "StandardSaveformat", CStr(.StandardSaveformat), REG_SZ
50380   reg.SetRegistryValue "StandardSubject", CStr(.StandardSubject), REG_SZ
50390   reg.SetRegistryValue "StandardTitle", CStr(.StandardTitle), REG_SZ
50400   reg.SetRegistryValue "UseCreationDateNow", CStr(Abs(.UseCreationDateNow)), REG_SZ
50410   reg.SetRegistryValue "UseStandardAuthor", CStr(Abs(.UseStandardAuthor)), REG_SZ
50420   reg.Subkey = "Printing\Formats\Bitmap\Colors"
50430   If Not reg.KeyExists Then
50440    reg.CreateKey
50450   End If
50460   reg.SetRegistryValue "BitmapResolution", CStr(.BitmapResolution), REG_SZ
50470   reg.SetRegistryValue "BMPColorscount", CStr(.BMPColorscount), REG_SZ
50480   reg.SetRegistryValue "JPEGColorscount", CStr(.JPEGColorscount), REG_SZ
50490   reg.SetRegistryValue "JPEGQuality", CStr(.JPEGQuality), REG_SZ
50500   reg.SetRegistryValue "PCXColorscount", CStr(.PCXColorscount), REG_SZ
50510   reg.SetRegistryValue "PNGColorscount", CStr(.PNGColorscount), REG_SZ
50520   reg.SetRegistryValue "TIFFColorscount", CStr(.TIFFColorscount), REG_SZ
50530   reg.Subkey = "Printing\Formats\PDF\Colors"
50540   If Not reg.KeyExists Then
50550    reg.CreateKey
50560   End If
50570   reg.SetRegistryValue "PDFColorsCMYKToRGB", CStr(Abs(.PDFColorsCMYKToRGB)), REG_SZ
50580   reg.SetRegistryValue "PDFColorsColorModel", CStr(.PDFColorsColorModel), REG_SZ
50590   reg.SetRegistryValue "PDFColorsPreserveHalftone", CStr(Abs(.PDFColorsPreserveHalftone)), REG_SZ
50600   reg.SetRegistryValue "PDFColorsPreserveOverprint", CStr(Abs(.PDFColorsPreserveOverprint)), REG_SZ
50610   reg.SetRegistryValue "PDFColorsPreserveTransfer", CStr(Abs(.PDFColorsPreserveTransfer)), REG_SZ
50620   reg.Subkey = "Printing\Formats\PDF\Compression"
50630   If Not reg.KeyExists Then
50640    reg.CreateKey
50650   End If
50660   reg.SetRegistryValue "PDFCompressionColorCompression", CStr(Abs(.PDFCompressionColorCompression)), REG_SZ
50670   reg.SetRegistryValue "PDFCompressionColorCompressionChoice", CStr(.PDFCompressionColorCompressionChoice), REG_SZ
50680   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
50690   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
50700   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
50710   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
50720   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
50730   reg.SetRegistryValue "PDFCompressionColorResample", CStr(Abs(.PDFCompressionColorResample)), REG_SZ
50740   reg.SetRegistryValue "PDFCompressionColorResampleChoice", CStr(.PDFCompressionColorResampleChoice), REG_SZ
50750   reg.SetRegistryValue "PDFCompressionColorResolution", CStr(.PDFCompressionColorResolution), REG_SZ
50760   reg.SetRegistryValue "PDFCompressionGreyCompression", CStr(Abs(.PDFCompressionGreyCompression)), REG_SZ
50770   reg.SetRegistryValue "PDFCompressionGreyCompressionChoice", CStr(.PDFCompressionGreyCompressionChoice), REG_SZ
50780   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
50790   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
50800   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
50810   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
50820   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
50830   reg.SetRegistryValue "PDFCompressionGreyResample", CStr(Abs(.PDFCompressionGreyResample)), REG_SZ
50840   reg.SetRegistryValue "PDFCompressionGreyResampleChoice", CStr(.PDFCompressionGreyResampleChoice), REG_SZ
50850   reg.SetRegistryValue "PDFCompressionGreyResolution", CStr(.PDFCompressionGreyResolution), REG_SZ
50860   reg.SetRegistryValue "PDFCompressionMonoCompression", CStr(Abs(.PDFCompressionMonoCompression)), REG_SZ
50870   reg.SetRegistryValue "PDFCompressionMonoCompressionChoice", CStr(.PDFCompressionMonoCompressionChoice), REG_SZ
50880   reg.SetRegistryValue "PDFCompressionMonoResample", CStr(Abs(.PDFCompressionMonoResample)), REG_SZ
50890   reg.SetRegistryValue "PDFCompressionMonoResampleChoice", CStr(.PDFCompressionMonoResampleChoice), REG_SZ
50900   reg.SetRegistryValue "PDFCompressionMonoResolution", CStr(.PDFCompressionMonoResolution), REG_SZ
50910   reg.SetRegistryValue "PDFCompressionTextCompression", CStr(Abs(.PDFCompressionTextCompression)), REG_SZ
50920   reg.Subkey = "Printing\Formats\PDF\Fonts"
50930   If Not reg.KeyExists Then
50940    reg.CreateKey
50950   End If
50960   reg.SetRegistryValue "PDFFontsEmbedAll", CStr(Abs(.PDFFontsEmbedAll)), REG_SZ
50970   reg.SetRegistryValue "PDFFontsSubSetFonts", CStr(Abs(.PDFFontsSubSetFonts)), REG_SZ
50980   reg.SetRegistryValue "PDFFontsSubSetFontsPercent", CStr(.PDFFontsSubSetFontsPercent), REG_SZ
50990   reg.Subkey = "Printing\Formats\PDF\General"
51000   If Not reg.KeyExists Then
51010    reg.CreateKey
51020   End If
51030   reg.SetRegistryValue "PDFGeneralASCII85", CStr(Abs(.PDFGeneralASCII85)), REG_SZ
51040   reg.SetRegistryValue "PDFGeneralAutorotate", CStr(.PDFGeneralAutorotate), REG_SZ
51050   reg.SetRegistryValue "PDFGeneralCompatibility", CStr(.PDFGeneralCompatibility), REG_SZ
51060   reg.SetRegistryValue "PDFGeneralOverprint", CStr(.PDFGeneralOverprint), REG_SZ
51070   reg.SetRegistryValue "PDFGeneralResolution", CStr(.PDFGeneralResolution), REG_SZ
51080   reg.SetRegistryValue "PDFOptimize", CStr(Abs(.PDFOptimize)), REG_SZ
51090   reg.Subkey = "Printing\Formats\PDF\Security"
51100   If Not reg.KeyExists Then
51110    reg.CreateKey
51120   End If
51130   reg.SetRegistryValue "PDFAllowAssembly", CStr(Abs(.PDFAllowAssembly)), REG_SZ
51140   reg.SetRegistryValue "PDFAllowDegradedPrinting", CStr(Abs(.PDFAllowDegradedPrinting)), REG_SZ
51150   reg.SetRegistryValue "PDFAllowFillIn", CStr(Abs(.PDFAllowFillIn)), REG_SZ
51160   reg.SetRegistryValue "PDFAllowScreenReaders", CStr(Abs(.PDFAllowScreenReaders)), REG_SZ
51170   reg.SetRegistryValue "PDFDisallowCopy", CStr(Abs(.PDFDisallowCopy)), REG_SZ
51180   reg.SetRegistryValue "PDFDisallowModifyAnnotations", CStr(Abs(.PDFDisallowModifyAnnotations)), REG_SZ
51190   reg.SetRegistryValue "PDFDisallowModifyContents", CStr(Abs(.PDFDisallowModifyContents)), REG_SZ
51200   reg.SetRegistryValue "PDFDisallowPrinting", CStr(Abs(.PDFDisallowPrinting)), REG_SZ
51210   reg.SetRegistryValue "PDFEncryptor", CStr(.PDFEncryptor), REG_SZ
51220   reg.SetRegistryValue "PDFHighEncryption", CStr(Abs(.PDFHighEncryption)), REG_SZ
51230   reg.SetRegistryValue "PDFLowEncryption", CStr(Abs(.PDFLowEncryption)), REG_SZ
51240   reg.SetRegistryValue "PDFOwnerPass", CStr(Abs(.PDFOwnerPass)), REG_SZ
51250   reg.SetRegistryValue "PDFOwnerPasswordString", CStr(.PDFOwnerPasswordString), REG_SZ
51260   reg.SetRegistryValue "PDFUserPass", CStr(Abs(.PDFUserPass)), REG_SZ
51270   reg.SetRegistryValue "PDFUserPasswordString", CStr(.PDFUserPasswordString), REG_SZ
51280   reg.SetRegistryValue "PDFUseSecurity", CStr(Abs(.PDFUseSecurity)), REG_SZ
51290   reg.Subkey = "Printing\Formats\PS\LanguageLevel"
51300   If Not reg.KeyExists Then
51310    reg.CreateKey
51320   End If
51330   reg.SetRegistryValue "EPSLanguageLevel", CStr(.EPSLanguageLevel), REG_SZ
51340   reg.SetRegistryValue "PSLanguageLevel", CStr(.PSLanguageLevel), REG_SZ
51350   reg.Subkey = "Program"
51360   If Not reg.KeyExists Then
51370    reg.CreateKey
51380   End If
51390   reg.SetRegistryValue "AdditionalGhostscriptParameters", CStr(.AdditionalGhostscriptParameters), REG_SZ
51400   reg.SetRegistryValue "AdditionalGhostscriptSearchpath", CStr(.AdditionalGhostscriptSearchpath), REG_SZ
51410   reg.SetRegistryValue "AddWindowsFontpath", CStr(Abs(.AddWindowsFontpath)), REG_SZ
51420   reg.SetRegistryValue "AutosaveDirectory", CStr(.AutosaveDirectory), REG_SZ
51430   reg.SetRegistryValue "AutosaveFilename", CStr(.AutosaveFilename), REG_SZ
51440   reg.SetRegistryValue "AutosaveFormat", CStr(.AutosaveFormat), REG_SZ
51450   reg.SetRegistryValue "AutosaveStartStandardProgram", CStr(Abs(.AutosaveStartStandardProgram)), REG_SZ
51460   reg.SetRegistryValue "ClientComputerResolveIPAddress", CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
51470   reg.SetRegistryValue "DisableEmail", CStr(Abs(.DisableEmail)), REG_SZ
51480   reg.SetRegistryValue "DontUseDocumentSettings", CStr(Abs(.DontUseDocumentSettings)), REG_SZ
51490   reg.SetRegistryValue "FilenameSubstitutions", CStr(.FilenameSubstitutions), REG_SZ
51500   reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle", CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
51510   reg.SetRegistryValue "Language", CStr(.Language), REG_SZ
51520   reg.SetRegistryValue "LastSaveDirectory", CStr(.LastSaveDirectory), REG_SZ
51530   reg.SetRegistryValue "Logging", CStr(Abs(.Logging)), REG_SZ
51540   reg.SetRegistryValue "LogLines", CStr(.LogLines), REG_SZ
51550   reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter", CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
51560   reg.SetRegistryValue "NoProcessingAtStartup", CStr(Abs(.NoProcessingAtStartup)), REG_SZ
51570   reg.SetRegistryValue "NoPSCheck", CStr(Abs(.NoPSCheck)), REG_SZ
51580   reg.SetRegistryValue "OptionsDesign", CStr(.OptionsDesign), REG_SZ
51590   reg.SetRegistryValue "OptionsEnabled", CStr(Abs(.OptionsEnabled)), REG_SZ
51600   reg.SetRegistryValue "OptionsVisible", CStr(Abs(.OptionsVisible)), REG_SZ
51610   reg.SetRegistryValue "PrintAfterSaving", CStr(Abs(.PrintAfterSaving)), REG_SZ
51620   reg.SetRegistryValue "PrintAfterSavingDuplex", CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
51630   reg.SetRegistryValue "PrintAfterSavingNoCancel", CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
51640   reg.SetRegistryValue "PrintAfterSavingPrinter", CStr(.PrintAfterSavingPrinter), REG_SZ
51650   reg.SetRegistryValue "PrintAfterSavingQueryUser", CStr(.PrintAfterSavingQueryUser), REG_SZ
51660   reg.SetRegistryValue "PrintAfterSavingTumble", CStr(.PrintAfterSavingTumble), REG_SZ
51670   reg.SetRegistryValue "PrinterStop", CStr(Abs(.PrinterStop)), REG_SZ
51680   reg.SetRegistryValue "PrinterTemppath", CStr(.PrinterTemppath), REG_SZ
51690   reg.SetRegistryValue "ProcessPriority", CStr(.ProcessPriority), REG_SZ
51700   reg.SetRegistryValue "ProgramFont", CStr(.ProgramFont), REG_SZ
51710   reg.SetRegistryValue "ProgramFontCharset", CStr(.ProgramFontCharset), REG_SZ
51720   reg.SetRegistryValue "ProgramFontSize", CStr(.ProgramFontSize), REG_SZ
51730   reg.SetRegistryValue "RemoveAllKnownFileExtensions", CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
51740   reg.SetRegistryValue "RemoveSpaces", CStr(Abs(.RemoveSpaces)), REG_SZ
51750   reg.SetRegistryValue "RunProgramAfterSaving", CStr(Abs(.RunProgramAfterSaving)), REG_SZ
51760   reg.SetRegistryValue "RunProgramAfterSavingProgramname", CStr(.RunProgramAfterSavingProgramname), REG_SZ
51770   reg.SetRegistryValue "RunProgramAfterSavingProgramParameters", CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
51780   reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady", CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
51790   reg.SetRegistryValue "RunProgramAfterSavingWindowstyle", CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
51800   reg.SetRegistryValue "RunProgramBeforeSaving", CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
51810   reg.SetRegistryValue "RunProgramBeforeSavingProgramname", CStr(.RunProgramBeforeSavingProgramname), REG_SZ
51820   reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters", CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
51830   reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle", CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
51840   reg.SetRegistryValue "SaveFilename", CStr(.SaveFilename), REG_SZ
51850   reg.SetRegistryValue "SendEmailAfterAutoSaving", CStr(Abs(.SendEmailAfterAutoSaving)), REG_SZ
51860   reg.SetRegistryValue "SendMailMethod", CStr(.SendMailMethod), REG_SZ
51870   reg.SetRegistryValue "ShowAnimation", CStr(Abs(.ShowAnimation)), REG_SZ
51880   reg.SetRegistryValue "StartStandardProgram", CStr(Abs(.StartStandardProgram)), REG_SZ
51890   reg.SetRegistryValue "Toolbars", CStr(.Toolbars), REG_SZ
51900   reg.SetRegistryValue "UseAutosave", CStr(Abs(.UseAutosave)), REG_SZ
51910   reg.SetRegistryValue "UseAutosaveDirectory", CStr(Abs(.UseAutosaveDirectory)), REG_SZ
51920  End With
51930  Set reg = Nothing
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

Public Sub ShowOptions(Frm As Form, sOptions As tOptions)
 On Error Resume Next
 Dim i As Long, tList() As String, tStrA() As String, lsv As ListView
 With sOptions
  Frm.txtAdditionalGhostscriptParameters.Text = .AdditionalGhostscriptParameters
  Frm.txtAdditionalGhostscriptSearchpath.Text = .AdditionalGhostscriptSearchpath
  Frm.chkAddWindowsFontpath.Value = .AddWindowsFontpath
  Frm.txtAutosaveDirectory.Text = .AutosaveDirectory
  Frm.txtAutosaveFilename.Text = .AutosaveFilename
  Frm.cmbAutosaveFormat.ListIndex = .AutosaveFormat
  Frm.chkAutosaveStartStandardProgram.Value = .AutosaveStartStandardProgram
  Frm.txtBitmapResolution.Text = .BitmapResolution
  Frm.cmbBMPColors.ListIndex = .BMPColorscount
  Frm.txtGSbin.Text = .DirectoryGhostscriptBinaries
  Frm.txtGSfonts.Text = .DirectoryGhostscriptFonts
  Frm.txtGSlib.Text = .DirectoryGhostscriptLibraries
  Frm.txtGSResource.Text = .DirectoryGhostscriptResource
  Frm.cmbEPSLanguageLevel.ListIndex = .EPSLanguageLevel
  Set lsv = Frm.lsvFilenameSubst
  tList = Split(.FilenameSubstitutions, "\")
  For i = 0 To UBound(tList)
   If InStr(tList(i), "|") <= 0 Then
    tList(i) = tList(i) & "|"
   End If
   If UBound(Split(tList(i), "|")) = 1 Then
    tStrA = Split(tList(i), "|")
    lsv.ListItems.Add , , tStrA(0)
    lsv.ListItems(lsv.ListItems.Count).SubItems(1) = tStrA(1)
   End If
  Next i
  If lsv.ListItems.Count > 0 Then
   lsv.ListItems(1).Selected = True
   Frm.txtFilenameSubst(0).Text = lsv.ListItems(1).Text
   Frm.txtFilenameSubst(0).ToolTipText = Frm.txtFilenameSubst(0).Text
   Frm.txtFilenameSubst(1).Text = lsv.ListItems(1).SubItems(1)
   Frm.txtFilenameSubst(1).ToolTipText = Frm.txtFilenameSubst(1).Text
  End If
  Frm.chkFilenameSubst.Value = .FilenameSubstitutionsOnlyInTitle
  Frm.cmbJPEGColors.ListIndex = .JPEGColorscount
  Frm.txtJPEGQuality.Text = .JPEGQuality
  Frm.chkNoConfirmMessageSwitchingDefaultprinter = .NoConfirmMessageSwitchingDefaultprinter
  Frm.chkNoProcessingAtStartup = .NoProcessingAtStartup
  Frm.chkOnePagePerFile.Value = .OnePagePerFile
  Frm.cmbOptionsDesign.ListIndex = .OptionsDesign
  Frm.cmbPCXColors.ListIndex = .PCXColorscount
  Frm.chkAllowAssembly.Value = .PDFAllowAssembly
  Frm.chkAllowDegradedPrinting.Value = .PDFAllowDegradedPrinting
  Frm.chkAllowFillIn.Value = .PDFAllowFillIn
  Frm.chkAllowScreenReaders.Value = .PDFAllowScreenReaders
  Frm.chkPDFCMYKtoRGB.Value = .PDFColorsCMYKToRGB
  Frm.cmbPDFColorModel.ListIndex = .PDFColorsColorModel
  Frm.chkPDFPreserveHalftone.Value = .PDFColorsPreserveHalftone
  Frm.chkPDFPreserveOverprint.Value = .PDFColorsPreserveOverprint
  Frm.chkPDFPreserveTransfer.Value = .PDFColorsPreserveTransfer
  Frm.chkPDFColorComp.Value = .PDFCompressionColorCompression
  Frm.cmbPDFColorComp.ListIndex = .PDFCompressionColorCompressionChoice
  Frm.chkPDFColorResample.Value = .PDFCompressionColorResample
  Frm.cmbPDFColorResample.ListIndex = .PDFCompressionColorResampleChoice
  Frm.txtPDFColorRes.Text = .PDFCompressionColorResolution
  Frm.chkPDFGreyComp.Value = .PDFCompressionGreyCompression
  Frm.cmbPDFGreyComp.ListIndex = .PDFCompressionGreyCompressionChoice
  Frm.chkPDFGreyResample.Value = .PDFCompressionGreyResample
  Frm.cmbPDFGreyResample.ListIndex = .PDFCompressionGreyResampleChoice
  Frm.txtPDFGreyRes.Text = .PDFCompressionGreyResolution
  Frm.chkPDFMonoComp.Value = .PDFCompressionMonoCompression
  Frm.cmbPDFMonoComp.ListIndex = .PDFCompressionMonoCompressionChoice
  Frm.chkPDFMonoResample.Value = .PDFCompressionMonoResample
  Frm.cmbPDFMonoResample.ListIndex = .PDFCompressionMonoResampleChoice
  Frm.txtPDFMonoRes.Text = .PDFCompressionMonoResolution
  Frm.chkPDFTextComp.Value = .PDFCompressionTextCompression
  Frm.chkAllowCopy.Value = .PDFDisallowCopy
  Frm.chkAllowModifyAnnotations.Value = .PDFDisallowModifyAnnotations
  Frm.chkAllowModifyContents.Value = .PDFDisallowModifyContents
  Frm.chkAllowPrinting.Value = .PDFDisallowPrinting
  Frm.cmbPDFEncryptor.ItemData(Frm.cmbPDFEncryptor.ListIndex) = .PDFEncryptor
  Frm.chkPDFEmbedAll.Value = .PDFFontsEmbedAll
  Frm.chkPDFSubSetFonts.Value = .PDFFontsSubSetFonts
  Frm.txtPDFSubSetPerc.Text = .PDFFontsSubSetFontsPercent
  Frm.chkPDFASCII85.Value = .PDFGeneralASCII85
  Frm.cmbPDFRotate.ListIndex = .PDFGeneralAutorotate
  Frm.cmbPDFCompat.ListIndex = .PDFGeneralCompatibility
  Frm.cmbPDFOverprint.ListIndex = .PDFGeneralOverprint
  Frm.txtPDFRes.Text = .PDFGeneralResolution
  Frm.optEncHigh.Value = .PDFHighEncryption
  Frm.optEncLow.Value = .PDFLowEncryption
  Frm.chkPDFOptimize.Value = .PDFOptimize
  Frm.chkOwnerPass.Value = .PDFOwnerPass
  Frm.chkUserPass.Value = .PDFUserPass
  Frm.chkUseSecurity.Value = .PDFUseSecurity
  Frm.cmbPNGColors.ListIndex = .PNGColorscount
  Frm.chkPrintAfterSaving.Value = .PrintAfterSaving
  Frm.chkPrintAfterSavingDuplex.Value = .PrintAfterSavingDuplex
  Frm.chkPrintAfterSavingNoCancel.Value = .PrintAfterSavingNoCancel
  Frm.cmbPrintAfterSavingPrinter.Text = .PrintAfterSavingPrinter
  Frm.cmbPrintAfterSavingQueryUser.ListIndex = .PrintAfterSavingQueryUser
  Frm.cmbPrintAfterSavingTumble.ListIndex = .PrintAfterSavingTumble
  Frm.txtTemppath.Text = .PrinterTemppath
  Frm.sldProcessPriority.Value = .ProcessPriority
  For i = 0 To Frm.cmbFonts.ListCount - 1
    If UCase$(Frm.cmbFonts.List(i)) = UCase$(.ProgramFont) Then
     Frm.cmbFonts.ListIndex = i
     Exit For
    End If
  Next i
  Frm.cmbCharset.Text = .ProgramFontCharset
  Frm.cmbProgramFontSize.Text = .ProgramFontSize
  Frm.cmbPSLanguageLevel.ListIndex = .PSLanguageLevel
  Frm.chkSpaces.Value = .RemoveSpaces
  Frm.chkRunProgramAfterSaving.Value = .RunProgramAfterSaving
  Frm.cmbRunProgramAfterSavingProgramname.Text = .RunProgramAfterSavingProgramname
  Frm.txtRunProgramAfterSavingProgramParameters.Text = .RunProgramAfterSavingProgramParameters
  Frm.chkRunProgramAfterSavingWaitUntilReady.Value = .RunProgramAfterSavingWaitUntilReady
  Frm.cmbRunProgramAfterSavingWindowstyle.ListIndex = .RunProgramAfterSavingWindowstyle
  Frm.chkRunProgramBeforeSaving.Value = .RunProgramBeforeSaving
  Frm.cmbRunProgramBeforeSavingProgramname.Text = .RunProgramBeforeSavingProgramname
  Frm.txtRunProgramBeforeSavingProgramParameters.Text = .RunProgramBeforeSavingProgramParameters
  Frm.cmbRunProgramBeforeSavingWindowstyle.ListIndex = .RunProgramBeforeSavingWindowstyle
  Frm.txtSaveFilename.Text = .SaveFilename
  Frm.chkAutosaveSendEmail.Value = .SendEmailAfterAutoSaving
  Frm.cmbSendMailMethod.ListIndex = .SendMailMethod
  Frm.chkShowAnimation = .ShowAnimation
  Frm.txtStandardAuthor.Text = .StandardAuthor
  Frm.cmbStandardSaveformat.ListIndex = .StandardSaveformat
  Frm.cmbTIFFColors.ListIndex = .TIFFColorscount
  Frm.chkUseAutosave.Value = .UseAutosave
  Frm.chkUseAutosaveDirectory.Value = .UseAutosaveDirectory
  Frm.chkUseCreationDateNow.Value = .UseCreationDateNow
  Frm.chkUseStandardAuthor.Value = .UseStandardAuthor
 End With
End Sub

Public Sub GetOptions(Frm As Form, sOptions As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, tStr As String, lsv As ListView
50020  With sOptions
50030  .AdditionalGhostscriptParameters = Frm.txtAdditionalGhostscriptParameters.Text
50040  .AdditionalGhostscriptSearchpath = Frm.txtAdditionalGhostscriptSearchpath.Text
50050  .AddWindowsFontpath = Abs(Frm.chkAddWindowsFontpath.Value)
50060  .AutosaveDirectory = Frm.txtAutosaveDirectory.Text
50070  .AutosaveFilename = Frm.txtAutosaveFilename.Text
50080  .AutosaveFormat = Frm.cmbAutosaveFormat.ListIndex
50090  .AutosaveStartStandardProgram = Abs(Frm.chkAutosaveStartStandardProgram.Value)
50100  .BitmapResolution = Frm.txtBitmapResolution.Text
50110  .BMPColorscount = Frm.cmbBMPColors.ListIndex
50120  .DirectoryGhostscriptBinaries = Frm.txtGSbin.Text
50130  .DirectoryGhostscriptFonts = Frm.txtGSfonts.Text
50140  .DirectoryGhostscriptLibraries = Frm.txtGSlib.Text
50150  .DirectoryGhostscriptResource = Frm.txtGSResource.Text
50160  .EPSLanguageLevel = Frm.cmbEPSLanguageLevel.ListIndex
50170  tStr = ""
50180  Set lsv = Frm.lsvFilenameSubst
50190  For i = 1 To lsv.ListItems.Count
50200   If i < lsv.ListItems.Count Then
50210     tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1) & "\"
50220    Else
50230     tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1)
50240   End If
50250  Next i
50260  .FilenameSubstitutions = tStr
50270  .FilenameSubstitutionsOnlyInTitle = Abs(Frm.chkFilenameSubst.Value)
50280  .JPEGColorscount = Frm.cmbJPEGColors.ListIndex
50290  .JPEGQuality = Frm.txtJPEGQuality.Text
50300  .NoConfirmMessageSwitchingDefaultprinter = Abs(Frm.chkNoConfirmMessageSwitchingDefaultprinter)
50310  .NoProcessingAtStartup = Abs(Frm.chkNoProcessingAtStartup)
50320  .OnePagePerFile = Abs(Frm.chkOnePagePerFile.Value)
50330  .OptionsDesign = Frm.cmbOptionsDesign.ListIndex
50340  .PCXColorscount = Frm.cmbPCXColors.ListIndex
50350  .PDFAllowAssembly = Abs(Frm.chkAllowAssembly.Value)
50360  .PDFAllowDegradedPrinting = Abs(Frm.chkAllowDegradedPrinting.Value)
50370  .PDFAllowFillIn = Abs(Frm.chkAllowFillIn.Value)
50380  .PDFAllowScreenReaders = Abs(Frm.chkAllowScreenReaders.Value)
50390  .PDFColorsCMYKToRGB = Abs(Frm.chkPDFCMYKtoRGB.Value)
50400  .PDFColorsColorModel = Frm.cmbPDFColorModel.ListIndex
50410  .PDFColorsPreserveHalftone = Abs(Frm.chkPDFPreserveHalftone.Value)
50420  .PDFColorsPreserveOverprint = Abs(Frm.chkPDFPreserveOverprint.Value)
50430  .PDFColorsPreserveTransfer = Abs(Frm.chkPDFPreserveTransfer.Value)
50440  .PDFCompressionColorCompression = Abs(Frm.chkPDFColorComp.Value)
50450  .PDFCompressionColorCompressionChoice = Frm.cmbPDFColorComp.ListIndex
50460  .PDFCompressionColorResample = Abs(Frm.chkPDFColorResample.Value)
50470  .PDFCompressionColorResampleChoice = Frm.cmbPDFColorResample.ListIndex
50480  .PDFCompressionColorResolution = Frm.txtPDFColorRes.Text
50490  .PDFCompressionGreyCompression = Abs(Frm.chkPDFGreyComp.Value)
50500  .PDFCompressionGreyCompressionChoice = Frm.cmbPDFGreyComp.ListIndex
50510  .PDFCompressionGreyResample = Abs(Frm.chkPDFGreyResample.Value)
50520  .PDFCompressionGreyResampleChoice = Frm.cmbPDFGreyResample.ListIndex
50530  .PDFCompressionGreyResolution = Frm.txtPDFGreyRes.Text
50540  .PDFCompressionMonoCompression = Abs(Frm.chkPDFMonoComp.Value)
50550  .PDFCompressionMonoCompressionChoice = Frm.cmbPDFMonoComp.ListIndex
50560  .PDFCompressionMonoResample = Abs(Frm.chkPDFMonoResample.Value)
50570  .PDFCompressionMonoResampleChoice = Frm.cmbPDFMonoResample.ListIndex
50580  .PDFCompressionMonoResolution = Frm.txtPDFMonoRes.Text
50590  .PDFCompressionTextCompression = Abs(Frm.chkPDFTextComp.Value)
50600  .PDFDisallowCopy = Abs(Frm.chkAllowCopy.Value)
50610  .PDFDisallowModifyAnnotations = Abs(Frm.chkAllowModifyAnnotations.Value)
50620  .PDFDisallowModifyContents = Abs(Frm.chkAllowModifyContents.Value)
50630  .PDFDisallowPrinting = Abs(Frm.chkAllowPrinting.Value)
50640  If Frm.cmbPDFEncryptor.ListIndex < 0 Then
50650    .PDFEncryptor = 0
50660   Else
50670    .PDFEncryptor = Frm.cmbPDFEncryptor.ItemData(Frm.cmbPDFEncryptor.ListIndex)
50680  End If
50690  .PDFFontsEmbedAll = Abs(Frm.chkPDFEmbedAll.Value)
50700  .PDFFontsSubSetFonts = Abs(Frm.chkPDFSubSetFonts.Value)
50710  .PDFFontsSubSetFontsPercent = Frm.txtPDFSubSetPerc.Text
50720  .PDFGeneralASCII85 = Abs(Frm.chkPDFASCII85.Value)
50730  .PDFGeneralAutorotate = Frm.cmbPDFRotate.ListIndex
50740  .PDFGeneralCompatibility = Frm.cmbPDFCompat.ListIndex
50750  .PDFGeneralOverprint = Frm.cmbPDFOverprint.ListIndex
50760  .PDFGeneralResolution = Frm.txtPDFRes.Text
50770  .PDFHighEncryption = Abs(Frm.optEncHigh.Value)
50780  .PDFLowEncryption = Abs(Frm.optEncLow.Value)
50790  .PDFOptimize = Abs(Frm.chkPDFOptimize.Value)
50800  .PDFOwnerPass = Abs(Frm.chkOwnerPass.Value)
50810  .PDFUserPass = Abs(Frm.chkUserPass.Value)
50820  .PDFUseSecurity = Abs(Frm.chkUseSecurity.Value)
50830  .PNGColorscount = Frm.cmbPNGColors.ListIndex
50840  .PrintAfterSaving = Abs(Frm.chkPrintAfterSaving.Value)
50850  .PrintAfterSavingDuplex = Abs(Frm.chkPrintAfterSavingDuplex.Value)
50860  .PrintAfterSavingNoCancel = Abs(Frm.chkPrintAfterSavingNoCancel.Value)
50870  .PrintAfterSavingPrinter = Frm.cmbPrintAfterSavingPrinter.Text
50880  .PrintAfterSavingQueryUser = Frm.cmbPrintAfterSavingQueryUser.ListIndex
50890  .PrintAfterSavingTumble = Frm.cmbPrintAfterSavingTumble.ListIndex
50900  .PrinterTemppath = Frm.txtTemppath.Text
50910  .ProcessPriority = Frm.sldProcessPriority.Value
50920  .ProgramFont = Frm.cmbFonts.List(Frm.cmbFonts.ListIndex)
50930  .ProgramFontCharset = Frm.cmbCharset.Text
50940  .ProgramFontSize = Frm.cmbProgramFontSize.Text
50950  .PSLanguageLevel = Frm.cmbPSLanguageLevel.ListIndex
50960  .RemoveSpaces = Abs(Frm.chkSpaces.Value)
50970  .RunProgramAfterSaving = Abs(Frm.chkRunProgramAfterSaving.Value)
50980  .RunProgramAfterSavingProgramname = Frm.cmbRunProgramAfterSavingProgramname.Text
50990  .RunProgramAfterSavingProgramParameters = Frm.txtRunProgramAfterSavingProgramParameters.Text
51000  .RunProgramAfterSavingWaitUntilReady = Abs(Frm.chkRunProgramAfterSavingWaitUntilReady.Value)
51010  .RunProgramAfterSavingWindowstyle = Frm.cmbRunProgramAfterSavingWindowstyle.ListIndex
51020  .RunProgramBeforeSaving = Abs(Frm.chkRunProgramBeforeSaving.Value)
51030  .RunProgramBeforeSavingProgramname = Frm.cmbRunProgramBeforeSavingProgramname.Text
51040  .RunProgramBeforeSavingProgramParameters = Frm.txtRunProgramBeforeSavingProgramParameters.Text
51050  .RunProgramBeforeSavingWindowstyle = Frm.cmbRunProgramBeforeSavingWindowstyle.ListIndex
51060  .SaveFilename = Frm.txtSaveFilename.Text
51070  .SendEmailAfterAutoSaving = Abs(Frm.chkAutosaveSendEmail.Value)
51080  .SendMailMethod = Frm.cmbSendMailMethod.ListIndex
51090  .ShowAnimation = Abs(Frm.chkShowAnimation)
51100  .StandardAuthor = Frm.txtStandardAuthor.Text
51110  .StandardSaveformat = Frm.cmbStandardSaveformat.ListIndex
51120  .TIFFColorscount = Frm.cmbTIFFColors.ListIndex
51130  .UseAutosave = Abs(Frm.chkUseAutosave.Value)
51140  .UseAutosaveDirectory = Abs(Frm.chkUseAutosaveDirectory.Value)
51150  .UseCreationDateNow = Abs(Frm.chkUseCreationDateNow.Value)
51160  .UseStandardAuthor = Abs(Frm.chkUseStandardAuthor.Value)
51170  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "GetOptions")
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

