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
 StandardSaveformat As String
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
50130   .BitmapResolution = "150"
50140   .BMPColorscount = "1"
50150   .ClientComputerResolveIPAddress = "0"
50160   .DeviceHeightPoints = Replace$("-1", ".", GetDecimalChar)
50170   .DeviceWidthPoints = Replace$("-1", ".", GetDecimalChar)
50180   Set reg = New clsRegistry
50190   reg.hkey = HKEY_LOCAL_MACHINE
50200   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50210   .DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50220   Set reg = Nothing
50230   Set reg = New clsRegistry
50240   reg.hkey = HKEY_LOCAL_MACHINE
50250   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50260   .DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50270   Set reg = Nothing
50280   Set reg = New clsRegistry
50290   reg.hkey = HKEY_LOCAL_MACHINE
50300   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50310   .DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50320   Set reg = Nothing
50330   Set reg = New clsRegistry
50340   reg.hkey = HKEY_LOCAL_MACHINE
50350   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50360   .DirectoryGhostscriptResource = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryResource"))
50370   Set reg = Nothing
50380   .DisableEmail = "0"
50390   .DontUseDocumentSettings = "0"
50400   .EPSLanguageLevel = "2"
50410   .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
50420   .FilenameSubstitutionsOnlyInTitle = "1"
50430   .JPEGColorscount = "0"
50440   .JPEGQuality = "75"
50450   .Language = "english"
50460   If InstalledAsServer Then
50470     .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
50480    Else
50490     .LastSaveDirectory = "<MyFiles>"
50500   End If
50510   .Logging = "0"
50520   .LogLines = "100"
50530   .NoConfirmMessageSwitchingDefaultprinter = "0"
50540   .NoProcessingAtStartup = "0"
50550   .NoPSCheck = "0"
50560   .OnePagePerFile = "0"
50570   .OptionsDesign = "1"
50580   .OptionsEnabled = "1"
50590   .OptionsVisible = "1"
50600   .Papersize = vbNullString
50610   .PCXColorscount = "0"
50620   .PDFAllowAssembly = "0"
50630   .PDFAllowDegradedPrinting = "0"
50640   .PDFAllowFillIn = "0"
50650   .PDFAllowScreenReaders = "0"
50660   .PDFColorsCMYKToRGB = "0"
50670   .PDFColorsColorModel = "1"
50680   .PDFColorsPreserveHalftone = "0"
50690   .PDFColorsPreserveOverprint = "1"
50700   .PDFColorsPreserveTransfer = "1"
50710   .PDFCompressionColorCompression = "1"
50720   .PDFCompressionColorCompressionChoice = "0"
50730   .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
50740   .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
50750   .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
50760   .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
50770   .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
50780   .PDFCompressionColorResample = "0"
50790   .PDFCompressionColorResampleChoice = "0"
50800   .PDFCompressionColorResolution = "300"
50810   .PDFCompressionGreyCompression = "1"
50820   .PDFCompressionGreyCompressionChoice = "0"
50830   .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
50840   .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
50850   .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
50860   .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
50870   .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
50880   .PDFCompressionGreyResample = "0"
50890   .PDFCompressionGreyResampleChoice = "0"
50900   .PDFCompressionGreyResolution = "300"
50910   .PDFCompressionMonoCompression = "1"
50920   .PDFCompressionMonoCompressionChoice = "0"
50930   .PDFCompressionMonoResample = "0"
50940   .PDFCompressionMonoResampleChoice = "0"
50950   .PDFCompressionMonoResolution = "1200"
50960   .PDFCompressionTextCompression = "1"
50970   .PDFDisallowCopy = "1"
50980   .PDFDisallowModifyAnnotations = "0"
50990   .PDFDisallowModifyContents = "0"
51000   .PDFDisallowPrinting = "0"
51010   .PDFEncryptor = "0"
51020   .PDFFontsEmbedAll = "1"
51030   .PDFFontsSubSetFonts = "1"
51040   .PDFFontsSubSetFontsPercent = "100"
51050   .PDFGeneralASCII85 = "0"
51060   .PDFGeneralAutorotate = "2"
51070   .PDFGeneralCompatibility = "1"
51080   .PDFGeneralOverprint = "0"
51090   .PDFGeneralResolution = "600"
51100   .PDFHighEncryption = "0"
51110   .PDFLowEncryption = "1"
51120   .PDFOptimize = "0"
51130   .PDFOwnerPass = "0"
51140   .PDFOwnerPasswordString = vbNullString
51150   .PDFUserPass = "0"
51160   .PDFUserPasswordString = vbNullString
51170   .PDFUseSecurity = "0"
51180   .PNGColorscount = "0"
51190   .PrintAfterSaving = "0"
51200   .PrintAfterSavingDuplex = "0"
51210   .PrintAfterSavingNoCancel = "0"
51220   .PrintAfterSavingPrinter = vbNullString
51230   .PrintAfterSavingQueryUser = "0"
51240   .PrintAfterSavingTumble = "0"
51250   .PrinterStop = "0"
51260   If InstalledAsServer Then
51270     .PrinterTemppath = CompletePath(GetPDFCreatorApplicationPath) & "Temp\"
51280    Else
51290     .PrinterTemppath = "<Temp>PDFCreator\"
51300   End If
51310   .ProcessPriority = "1"
51320   .ProgramFont = "MS Sans Serif"
51330   .ProgramFontCharset = "0"
51340   .ProgramFontSize = "8"
51350   .PSLanguageLevel = "2"
51360   .RemoveAllKnownFileExtensions = "1"
51370   .RemoveSpaces = "1"
51380   .RunProgramAfterSaving = "0"
51390   .RunProgramAfterSavingProgramname = vbNullString
51400   .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
51410   .RunProgramAfterSavingWaitUntilReady = "1"
51420   .RunProgramAfterSavingWindowstyle = "1"
51430   .RunProgramBeforeSaving = "0"
51440   .RunProgramBeforeSavingProgramname = vbNullString
51450   .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
51460   .RunProgramBeforeSavingWindowstyle = "1"
51470   .SaveFilename = "<Title>"
51480   .SendEmailAfterAutoSaving = "0"
51490   .SendMailMethod = "0"
51500   .ShowAnimation = "1"
51510   .StampFontColor = "#FF0000"
51520   .StampFontname = "Arial"
51530   .StampFontsize = "48"
51540   .StampOutlineFontthickness = "0"
51550   .StampString = vbNullString
51560   .StampUseOutlineFont = "1"
51570   .StandardAuthor = vbNullString
51580   .StandardCreationdate = vbNullString
51590   .StandardDateformat = "YYYYMMDDHHNNSS"
51600   .StandardKeywords = vbNullString
51610   .StandardMailDomain = vbNullString
51620   .StandardModifydate = vbNullString
51630   .StandardSaveformat = "pdf"
51640   .StandardSubject = vbNullString
51650   .StandardTitle = vbNullString
51660   .StartStandardProgram = "1"
51670   .TIFFColorscount = "0"
51680   .Toolbars = "1"
51690   .UseAutosave = "0"
51700   .UseAutosaveDirectory = "1"
51710   .UseCreationDateNow = "0"
51720   .UseStandardAuthor = "0"
51730  End With
51740  If UseINI Then
51750    If Not IsWin9xMe Then
51760     myOptions = ReadOptionsINI(myOptions, CompletePath(GetDefaultAppData) & "PDFCreator.ini", False, False)
51770    End If
51780   Else
51790    If Not IsWin9xMe Then
51800     myOptions = ReadOptionsReg(myOptions, ".DEFAULT\Software\PDFCreator", HKEY_USERS, False, False)
51810    End If
51820  End If
51830  StandardOptions = myOptions
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
50780   tStr = hOpt.Retrieve("BitmapResolution")
50790   If IsNumeric(tStr) Then
50800     If CLng(tStr) >= 1 Then
50810       .BitmapResolution = CLng(tStr)
50820      Else
50830       If UseStandard Then
50840        .BitmapResolution = 150
50850       End If
50860     End If
50870    Else
50880     If UseStandard Then
50890      .BitmapResolution = 150
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
51060   tStr = hOpt.Retrieve("ClientComputerResolveIPAddress")
51070   If IsNumeric(tStr) Then
51080     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51090       .ClientComputerResolveIPAddress = CLng(tStr)
51100      Else
51110       If UseStandard Then
51120        .ClientComputerResolveIPAddress = 0
51130       End If
51140     End If
51150    Else
51160     If UseStandard Then
51170      .ClientComputerResolveIPAddress = 0
51180     End If
51190   End If
51200   tStr = hOpt.Retrieve("DeviceHeightPoints")
51210   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
51220     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= -1 Then
51230       .DeviceHeightPoints = CDbl(Replace$(tStr, ".", GetDecimalChar))
51240      Else
51250       If UseStandard Then
51260        .DeviceHeightPoints = Replace$("-1", ".", GetDecimalChar)
51270       End If
51280     End If
51290    Else
51300     If UseStandard Then
51310      .DeviceHeightPoints = Replace$("-1", ".", GetDecimalChar)
51320     End If
51330   End If
51340   tStr = hOpt.Retrieve("DeviceWidthPoints")
51350   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
51360     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= -1 Then
51370       .DeviceWidthPoints = CDbl(Replace$(tStr, ".", GetDecimalChar))
51380      Else
51390       If UseStandard Then
51400        .DeviceWidthPoints = Replace$("-1", ".", GetDecimalChar)
51410       End If
51420     End If
51430    Else
51440     If UseStandard Then
51450      .DeviceWidthPoints = Replace$("-1", ".", GetDecimalChar)
51460     End If
51470   End If
51480   tStr = hOpt.Retrieve("DirectoryGhostscriptBinaries")
51490   If LenB(Trim$(tStr)) > 0 Then
51500     .DirectoryGhostscriptBinaries = CompletePath(tStr)
51510    Else
51520     If UseStandard Then
51530      tStr = GetPDFCreatorApplicationPath
51540      .DirectoryGhostscriptBinaries = CompletePath(tStr)
51550     End If
51560   End If
51570   tStr = hOpt.Retrieve("DirectoryGhostscriptFonts")
51580   If LenB(Trim$(tStr)) > 0 Then
51590     .DirectoryGhostscriptFonts = CompletePath(tStr)
51600    Else
51610     If UseStandard Then
51620      tStr = GetPDFCreatorApplicationPath & "fonts"
51630      .DirectoryGhostscriptFonts = CompletePath(tStr)
51640     End If
51650   End If
51660   tStr = hOpt.Retrieve("DirectoryGhostscriptLibraries")
51670   If LenB(Trim$(tStr)) > 0 Then
51680     .DirectoryGhostscriptLibraries = CompletePath(tStr)
51690    Else
51700     If UseStandard Then
51710      tStr = GetPDFCreatorApplicationPath & "lib"
51720      .DirectoryGhostscriptLibraries = CompletePath(tStr)
51730     End If
51740   End If
51750   tStr = hOpt.Retrieve("DirectoryGhostscriptResource")
51760   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
51770     .DirectoryGhostscriptResource = ""
51780    Else
51790     If LenB(tStr) > 0 Then
51800      .DirectoryGhostscriptResource = tStr
51810     End If
51820   End If
51830   tStr = hOpt.Retrieve("DisableEmail")
51840   If IsNumeric(tStr) Then
51850     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51860       .DisableEmail = CLng(tStr)
51870      Else
51880       If UseStandard Then
51890        .DisableEmail = 0
51900       End If
51910     End If
51920    Else
51930     If UseStandard Then
51940      .DisableEmail = 0
51950     End If
51960   End If
51970   tStr = hOpt.Retrieve("DontUseDocumentSettings")
51980   If IsNumeric(tStr) Then
51990     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52000       .DontUseDocumentSettings = CLng(tStr)
52010      Else
52020       If UseStandard Then
52030        .DontUseDocumentSettings = 0
52040       End If
52050     End If
52060    Else
52070     If UseStandard Then
52080      .DontUseDocumentSettings = 0
52090     End If
52100   End If
52110   tStr = hOpt.Retrieve("EPSLanguageLevel")
52120   If IsNumeric(tStr) Then
52130     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
52140       .EPSLanguageLevel = CLng(tStr)
52150      Else
52160       If UseStandard Then
52170        .EPSLanguageLevel = 2
52180       End If
52190     End If
52200    Else
52210     If UseStandard Then
52220      .EPSLanguageLevel = 2
52230     End If
52240   End If
52250   tStr = hOpt.Retrieve("FilenameSubstitutions")
52260   If LenB(tStr) = 0 And LenB("Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt") > 0 And UseStandard Then
52270     .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
52280    Else
52290     If LenB(tStr) > 0 Then
52300      .FilenameSubstitutions = tStr
52310     End If
52320   End If
52330   tStr = hOpt.Retrieve("FilenameSubstitutionsOnlyInTitle")
52340   If IsNumeric(tStr) Then
52350     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52360       .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
52370      Else
52380       If UseStandard Then
52390        .FilenameSubstitutionsOnlyInTitle = 1
52400       End If
52410     End If
52420    Else
52430     If UseStandard Then
52440      .FilenameSubstitutionsOnlyInTitle = 1
52450     End If
52460   End If
52470   tStr = hOpt.Retrieve("JPEGColorscount")
52480   If IsNumeric(tStr) Then
52490     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
52500       .JPEGColorscount = CLng(tStr)
52510      Else
52520       If UseStandard Then
52530        .JPEGColorscount = 0
52540       End If
52550     End If
52560    Else
52570     If UseStandard Then
52580      .JPEGColorscount = 0
52590     End If
52600   End If
52610   tStr = hOpt.Retrieve("JPEGQuality")
52620   If IsNumeric(tStr) Then
52630     If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
52640       .JPEGQuality = CLng(tStr)
52650      Else
52660       If UseStandard Then
52670        .JPEGQuality = 75
52680       End If
52690     End If
52700    Else
52710     If UseStandard Then
52720      .JPEGQuality = 75
52730     End If
52740   End If
52750   tStr = hOpt.Retrieve("Language")
52760   If LenB(tStr) = 0 And LenB("english") > 0 And UseStandard Then
52770     .Language = "english"
52780    Else
52790     If LenB(tStr) > 0 Then
52800      .Language = tStr
52810     End If
52820   End If
52830   tStr = hOpt.Retrieve("LastSaveDirectory")
52840   If LenB(Trim$(tStr)) > 0 Then
52850     .LastSaveDirectory = CompletePath(tStr)
52860    Else
52870     If UseStandard Then
52880      If InstalledAsServer Then
52890        .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
52900       Else
52910        .LastSaveDirectory = "<MyFiles>"
52920      End If
52930     End If
52940   End If
52950   tStr = hOpt.Retrieve("Logging")
52960   If IsNumeric(tStr) Then
52970     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52980       .Logging = CLng(tStr)
52990      Else
53000       If UseStandard Then
53010        .Logging = 0
53020       End If
53030     End If
53040    Else
53050     If UseStandard Then
53060      .Logging = 0
53070     End If
53080   End If
53090   tStr = hOpt.Retrieve("LogLines")
53100   If IsNumeric(tStr) Then
53110     If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
53120       .LogLines = CLng(tStr)
53130      Else
53140       If UseStandard Then
53150        .LogLines = 100
53160       End If
53170     End If
53180    Else
53190     If UseStandard Then
53200      .LogLines = 100
53210     End If
53220   End If
53230   tStr = hOpt.Retrieve("NoConfirmMessageSwitchingDefaultprinter")
53240   If IsNumeric(tStr) Then
53250     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53260       .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
53270      Else
53280       If UseStandard Then
53290        .NoConfirmMessageSwitchingDefaultprinter = 0
53300       End If
53310     End If
53320    Else
53330     If UseStandard Then
53340      .NoConfirmMessageSwitchingDefaultprinter = 0
53350     End If
53360   End If
53370   tStr = hOpt.Retrieve("NoProcessingAtStartup")
53380   If IsNumeric(tStr) Then
53390     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53400       .NoProcessingAtStartup = CLng(tStr)
53410      Else
53420       If UseStandard Then
53430        .NoProcessingAtStartup = 0
53440       End If
53450     End If
53460    Else
53470     If UseStandard Then
53480      .NoProcessingAtStartup = 0
53490     End If
53500   End If
53510   tStr = hOpt.Retrieve("NoPSCheck")
53520   If IsNumeric(tStr) Then
53530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53540       .NoPSCheck = CLng(tStr)
53550      Else
53560       If UseStandard Then
53570        .NoPSCheck = 0
53580       End If
53590     End If
53600    Else
53610     If UseStandard Then
53620      .NoPSCheck = 0
53630     End If
53640   End If
53650   tStr = hOpt.Retrieve("OnePagePerFile")
53660   If IsNumeric(tStr) Then
53670     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53680       .OnePagePerFile = CLng(tStr)
53690      Else
53700       If UseStandard Then
53710        .OnePagePerFile = 0
53720       End If
53730     End If
53740    Else
53750     If UseStandard Then
53760      .OnePagePerFile = 0
53770     End If
53780   End If
53790   tStr = hOpt.Retrieve("OptionsDesign")
53800   If IsNumeric(tStr) Then
53810     If CLng(tStr) >= 1 And CLng(tStr) <= 2 Then
53820       .OptionsDesign = CLng(tStr)
53830      Else
53840       If UseStandard Then
53850        .OptionsDesign = 1
53860       End If
53870     End If
53880    Else
53890     If UseStandard Then
53900      .OptionsDesign = 1
53910     End If
53920   End If
53930   tStr = hOpt.Retrieve("OptionsEnabled")
53940   If IsNumeric(tStr) Then
53950     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53960       .OptionsEnabled = CLng(tStr)
53970      Else
53980       If UseStandard Then
53990        .OptionsEnabled = 1
54000       End If
54010     End If
54020    Else
54030     If UseStandard Then
54040      .OptionsEnabled = 1
54050     End If
54060   End If
54070   tStr = hOpt.Retrieve("OptionsVisible")
54080   If IsNumeric(tStr) Then
54090     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54100       .OptionsVisible = CLng(tStr)
54110      Else
54120       If UseStandard Then
54130        .OptionsVisible = 1
54140       End If
54150     End If
54160    Else
54170     If UseStandard Then
54180      .OptionsVisible = 1
54190     End If
54200   End If
54210   tStr = hOpt.Retrieve("Papersize")
54220   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
54230     .Papersize = ""
54240    Else
54250     If LenB(tStr) > 0 Then
54260      .Papersize = tStr
54270     End If
54280   End If
54290   tStr = hOpt.Retrieve("PCXColorscount")
54300   If IsNumeric(tStr) Then
54310     If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
54320       .PCXColorscount = CLng(tStr)
54330      Else
54340       If UseStandard Then
54350        .PCXColorscount = 0
54360       End If
54370     End If
54380    Else
54390     If UseStandard Then
54400      .PCXColorscount = 0
54410     End If
54420   End If
54430   tStr = hOpt.Retrieve("PDFAllowAssembly")
54440   If IsNumeric(tStr) Then
54450     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54460       .PDFAllowAssembly = CLng(tStr)
54470      Else
54480       If UseStandard Then
54490        .PDFAllowAssembly = 0
54500       End If
54510     End If
54520    Else
54530     If UseStandard Then
54540      .PDFAllowAssembly = 0
54550     End If
54560   End If
54570   tStr = hOpt.Retrieve("PDFAllowDegradedPrinting")
54580   If IsNumeric(tStr) Then
54590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54600       .PDFAllowDegradedPrinting = CLng(tStr)
54610      Else
54620       If UseStandard Then
54630        .PDFAllowDegradedPrinting = 0
54640       End If
54650     End If
54660    Else
54670     If UseStandard Then
54680      .PDFAllowDegradedPrinting = 0
54690     End If
54700   End If
54710   tStr = hOpt.Retrieve("PDFAllowFillIn")
54720   If IsNumeric(tStr) Then
54730     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54740       .PDFAllowFillIn = CLng(tStr)
54750      Else
54760       If UseStandard Then
54770        .PDFAllowFillIn = 0
54780       End If
54790     End If
54800    Else
54810     If UseStandard Then
54820      .PDFAllowFillIn = 0
54830     End If
54840   End If
54850   tStr = hOpt.Retrieve("PDFAllowScreenReaders")
54860   If IsNumeric(tStr) Then
54870     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54880       .PDFAllowScreenReaders = CLng(tStr)
54890      Else
54900       If UseStandard Then
54910        .PDFAllowScreenReaders = 0
54920       End If
54930     End If
54940    Else
54950     If UseStandard Then
54960      .PDFAllowScreenReaders = 0
54970     End If
54980   End If
54990   tStr = hOpt.Retrieve("PDFColorsCMYKToRGB")
55000   If IsNumeric(tStr) Then
55010     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55020       .PDFColorsCMYKToRGB = CLng(tStr)
55030      Else
55040       If UseStandard Then
55050        .PDFColorsCMYKToRGB = 0
55060       End If
55070     End If
55080    Else
55090     If UseStandard Then
55100      .PDFColorsCMYKToRGB = 0
55110     End If
55120   End If
55130   tStr = hOpt.Retrieve("PDFColorsColorModel")
55140   If IsNumeric(tStr) Then
55150     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
55160       .PDFColorsColorModel = CLng(tStr)
55170      Else
55180       If UseStandard Then
55190        .PDFColorsColorModel = 1
55200       End If
55210     End If
55220    Else
55230     If UseStandard Then
55240      .PDFColorsColorModel = 1
55250     End If
55260   End If
55270   tStr = hOpt.Retrieve("PDFColorsPreserveHalftone")
55280   If IsNumeric(tStr) Then
55290     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55300       .PDFColorsPreserveHalftone = CLng(tStr)
55310      Else
55320       If UseStandard Then
55330        .PDFColorsPreserveHalftone = 0
55340       End If
55350     End If
55360    Else
55370     If UseStandard Then
55380      .PDFColorsPreserveHalftone = 0
55390     End If
55400   End If
55410   tStr = hOpt.Retrieve("PDFColorsPreserveOverprint")
55420   If IsNumeric(tStr) Then
55430     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55440       .PDFColorsPreserveOverprint = CLng(tStr)
55450      Else
55460       If UseStandard Then
55470        .PDFColorsPreserveOverprint = 1
55480       End If
55490     End If
55500    Else
55510     If UseStandard Then
55520      .PDFColorsPreserveOverprint = 1
55530     End If
55540   End If
55550   tStr = hOpt.Retrieve("PDFColorsPreserveTransfer")
55560   If IsNumeric(tStr) Then
55570     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55580       .PDFColorsPreserveTransfer = CLng(tStr)
55590      Else
55600       If UseStandard Then
55610        .PDFColorsPreserveTransfer = 1
55620       End If
55630     End If
55640    Else
55650     If UseStandard Then
55660      .PDFColorsPreserveTransfer = 1
55670     End If
55680   End If
55690   tStr = hOpt.Retrieve("PDFCompressionColorCompression")
55700   If IsNumeric(tStr) Then
55710     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55720       .PDFCompressionColorCompression = CLng(tStr)
55730      Else
55740       If UseStandard Then
55750        .PDFCompressionColorCompression = 1
55760       End If
55770     End If
55780    Else
55790     If UseStandard Then
55800      .PDFCompressionColorCompression = 1
55810     End If
55820   End If
55830   tStr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
55840   If IsNumeric(tStr) Then
55850     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
55860       .PDFCompressionColorCompressionChoice = CLng(tStr)
55870      Else
55880       If UseStandard Then
55890        .PDFCompressionColorCompressionChoice = 0
55900       End If
55910     End If
55920    Else
55930     If UseStandard Then
55940      .PDFCompressionColorCompressionChoice = 0
55950     End If
55960   End If
55970   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGHighFactor")
55980   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
55990     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56000       .PDFCompressionColorCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56010      Else
56020       If UseStandard Then
56030        .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56040       End If
56050     End If
56060    Else
56070     If UseStandard Then
56080      .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56090     End If
56100   End If
56110   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGLowFactor")
56120   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56130     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56140       .PDFCompressionColorCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56150      Else
56160       If UseStandard Then
56170        .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56180       End If
56190     End If
56200    Else
56210     If UseStandard Then
56220      .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56230     End If
56240   End If
56250   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMaximumFactor")
56260   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56270     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56280       .PDFCompressionColorCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56290      Else
56300       If UseStandard Then
56310        .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56320       End If
56330     End If
56340    Else
56350     If UseStandard Then
56360      .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56370     End If
56380   End If
56390   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMediumFactor")
56400   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56410     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56420       .PDFCompressionColorCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56430      Else
56440       If UseStandard Then
56450        .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56460       End If
56470     End If
56480    Else
56490     If UseStandard Then
56500      .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56510     End If
56520   End If
56530   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMinimumFactor")
56540   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56550     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56560       .PDFCompressionColorCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56570      Else
56580       If UseStandard Then
56590        .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
56600       End If
56610     End If
56620    Else
56630     If UseStandard Then
56640      .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
56650     End If
56660   End If
56670   tStr = hOpt.Retrieve("PDFCompressionColorResample")
56680   If IsNumeric(tStr) Then
56690     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56700       .PDFCompressionColorResample = CLng(tStr)
56710      Else
56720       If UseStandard Then
56730        .PDFCompressionColorResample = 0
56740       End If
56750     End If
56760    Else
56770     If UseStandard Then
56780      .PDFCompressionColorResample = 0
56790     End If
56800   End If
56810   tStr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
56820   If IsNumeric(tStr) Then
56830     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
56840       .PDFCompressionColorResampleChoice = CLng(tStr)
56850      Else
56860       If UseStandard Then
56870        .PDFCompressionColorResampleChoice = 0
56880       End If
56890     End If
56900    Else
56910     If UseStandard Then
56920      .PDFCompressionColorResampleChoice = 0
56930     End If
56940   End If
56950   tStr = hOpt.Retrieve("PDFCompressionColorResolution")
56960   If IsNumeric(tStr) Then
56970     If CLng(tStr) >= 0 Then
56980       .PDFCompressionColorResolution = CLng(tStr)
56990      Else
57000       If UseStandard Then
57010        .PDFCompressionColorResolution = 300
57020       End If
57030     End If
57040    Else
57050     If UseStandard Then
57060      .PDFCompressionColorResolution = 300
57070     End If
57080   End If
57090   tStr = hOpt.Retrieve("PDFCompressionGreyCompression")
57100   If IsNumeric(tStr) Then
57110     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57120       .PDFCompressionGreyCompression = CLng(tStr)
57130      Else
57140       If UseStandard Then
57150        .PDFCompressionGreyCompression = 1
57160       End If
57170     End If
57180    Else
57190     If UseStandard Then
57200      .PDFCompressionGreyCompression = 1
57210     End If
57220   End If
57230   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
57240   If IsNumeric(tStr) Then
57250     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
57260       .PDFCompressionGreyCompressionChoice = CLng(tStr)
57270      Else
57280       If UseStandard Then
57290        .PDFCompressionGreyCompressionChoice = 0
57300       End If
57310     End If
57320    Else
57330     If UseStandard Then
57340      .PDFCompressionGreyCompressionChoice = 0
57350     End If
57360   End If
57370   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGHighFactor")
57380   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57390     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57400       .PDFCompressionGreyCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57410      Else
57420       If UseStandard Then
57430        .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
57440       End If
57450     End If
57460    Else
57470     If UseStandard Then
57480      .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
57490     End If
57500   End If
57510   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGLowFactor")
57520   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57530     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57540       .PDFCompressionGreyCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57550      Else
57560       If UseStandard Then
57570        .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
57580       End If
57590     End If
57600    Else
57610     If UseStandard Then
57620      .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
57630     End If
57640   End If
57650   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMaximumFactor")
57660   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57670     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57680       .PDFCompressionGreyCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57690      Else
57700       If UseStandard Then
57710        .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57720       End If
57730     End If
57740    Else
57750     If UseStandard Then
57760      .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57770     End If
57780   End If
57790   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMediumFactor")
57800   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57810     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57820       .PDFCompressionGreyCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57830      Else
57840       If UseStandard Then
57850        .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57860       End If
57870     End If
57880    Else
57890     If UseStandard Then
57900      .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57910     End If
57920   End If
57930   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMinimumFactor")
57940   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57950     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57960       .PDFCompressionGreyCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57970      Else
57980       If UseStandard Then
57990        .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58000       End If
58010     End If
58020    Else
58030     If UseStandard Then
58040      .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58050     End If
58060   End If
58070   tStr = hOpt.Retrieve("PDFCompressionGreyResample")
58080   If IsNumeric(tStr) Then
58090     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58100       .PDFCompressionGreyResample = CLng(tStr)
58110      Else
58120       If UseStandard Then
58130        .PDFCompressionGreyResample = 0
58140       End If
58150     End If
58160    Else
58170     If UseStandard Then
58180      .PDFCompressionGreyResample = 0
58190     End If
58200   End If
58210   tStr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
58220   If IsNumeric(tStr) Then
58230     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
58240       .PDFCompressionGreyResampleChoice = CLng(tStr)
58250      Else
58260       If UseStandard Then
58270        .PDFCompressionGreyResampleChoice = 0
58280       End If
58290     End If
58300    Else
58310     If UseStandard Then
58320      .PDFCompressionGreyResampleChoice = 0
58330     End If
58340   End If
58350   tStr = hOpt.Retrieve("PDFCompressionGreyResolution")
58360   If IsNumeric(tStr) Then
58370     If CLng(tStr) >= 0 Then
58380       .PDFCompressionGreyResolution = CLng(tStr)
58390      Else
58400       If UseStandard Then
58410        .PDFCompressionGreyResolution = 300
58420       End If
58430     End If
58440    Else
58450     If UseStandard Then
58460      .PDFCompressionGreyResolution = 300
58470     End If
58480   End If
58490   tStr = hOpt.Retrieve("PDFCompressionMonoCompression")
58500   If IsNumeric(tStr) Then
58510     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58520       .PDFCompressionMonoCompression = CLng(tStr)
58530      Else
58540       If UseStandard Then
58550        .PDFCompressionMonoCompression = 1
58560       End If
58570     End If
58580    Else
58590     If UseStandard Then
58600      .PDFCompressionMonoCompression = 1
58610     End If
58620   End If
58630   tStr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
58640   If IsNumeric(tStr) Then
58650     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
58660       .PDFCompressionMonoCompressionChoice = CLng(tStr)
58670      Else
58680       If UseStandard Then
58690        .PDFCompressionMonoCompressionChoice = 0
58700       End If
58710     End If
58720    Else
58730     If UseStandard Then
58740      .PDFCompressionMonoCompressionChoice = 0
58750     End If
58760   End If
58770   tStr = hOpt.Retrieve("PDFCompressionMonoResample")
58780   If IsNumeric(tStr) Then
58790     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58800       .PDFCompressionMonoResample = CLng(tStr)
58810      Else
58820       If UseStandard Then
58830        .PDFCompressionMonoResample = 0
58840       End If
58850     End If
58860    Else
58870     If UseStandard Then
58880      .PDFCompressionMonoResample = 0
58890     End If
58900   End If
58910   tStr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
58920   If IsNumeric(tStr) Then
58930     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
58940       .PDFCompressionMonoResampleChoice = CLng(tStr)
58950      Else
58960       If UseStandard Then
58970        .PDFCompressionMonoResampleChoice = 0
58980       End If
58990     End If
59000    Else
59010     If UseStandard Then
59020      .PDFCompressionMonoResampleChoice = 0
59030     End If
59040   End If
59050   tStr = hOpt.Retrieve("PDFCompressionMonoResolution")
59060   If IsNumeric(tStr) Then
59070     If CLng(tStr) >= 0 Then
59080       .PDFCompressionMonoResolution = CLng(tStr)
59090      Else
59100       If UseStandard Then
59110        .PDFCompressionMonoResolution = 1200
59120       End If
59130     End If
59140    Else
59150     If UseStandard Then
59160      .PDFCompressionMonoResolution = 1200
59170     End If
59180   End If
59190   tStr = hOpt.Retrieve("PDFCompressionTextCompression")
59200   If IsNumeric(tStr) Then
59210     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59220       .PDFCompressionTextCompression = CLng(tStr)
59230      Else
59240       If UseStandard Then
59250        .PDFCompressionTextCompression = 1
59260       End If
59270     End If
59280    Else
59290     If UseStandard Then
59300      .PDFCompressionTextCompression = 1
59310     End If
59320   End If
59330   tStr = hOpt.Retrieve("PDFDisallowCopy")
59340   If IsNumeric(tStr) Then
59350     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59360       .PDFDisallowCopy = CLng(tStr)
59370      Else
59380       If UseStandard Then
59390        .PDFDisallowCopy = 1
59400       End If
59410     End If
59420    Else
59430     If UseStandard Then
59440      .PDFDisallowCopy = 1
59450     End If
59460   End If
59470   tStr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
59480   If IsNumeric(tStr) Then
59490     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59500       .PDFDisallowModifyAnnotations = CLng(tStr)
59510      Else
59520       If UseStandard Then
59530        .PDFDisallowModifyAnnotations = 0
59540       End If
59550     End If
59560    Else
59570     If UseStandard Then
59580      .PDFDisallowModifyAnnotations = 0
59590     End If
59600   End If
59610   tStr = hOpt.Retrieve("PDFDisallowModifyContents")
59620   If IsNumeric(tStr) Then
59630     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59640       .PDFDisallowModifyContents = CLng(tStr)
59650      Else
59660       If UseStandard Then
59670        .PDFDisallowModifyContents = 0
59680       End If
59690     End If
59700    Else
59710     If UseStandard Then
59720      .PDFDisallowModifyContents = 0
59730     End If
59740   End If
59750   tStr = hOpt.Retrieve("PDFDisallowPrinting")
59760   If IsNumeric(tStr) Then
59770     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59780       .PDFDisallowPrinting = CLng(tStr)
59790      Else
59800       If UseStandard Then
59810        .PDFDisallowPrinting = 0
59820       End If
59830     End If
59840    Else
59850     If UseStandard Then
59860      .PDFDisallowPrinting = 0
59870     End If
59880   End If
59890   tStr = hOpt.Retrieve("PDFEncryptor")
59900   If IsNumeric(tStr) Then
59910     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
59920       .PDFEncryptor = CLng(tStr)
59930      Else
59940       If UseStandard Then
59950        .PDFEncryptor = 0
59960       End If
59970     End If
59980    Else
59990     If UseStandard Then
60000      .PDFEncryptor = 0
60010     End If
60020   End If
60030   tStr = hOpt.Retrieve("PDFFontsEmbedAll")
60040   If IsNumeric(tStr) Then
60050     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60060       .PDFFontsEmbedAll = CLng(tStr)
60070      Else
60080       If UseStandard Then
60090        .PDFFontsEmbedAll = 1
60100       End If
60110     End If
60120    Else
60130     If UseStandard Then
60140      .PDFFontsEmbedAll = 1
60150     End If
60160   End If
60170   tStr = hOpt.Retrieve("PDFFontsSubSetFonts")
60180   If IsNumeric(tStr) Then
60190     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60200       .PDFFontsSubSetFonts = CLng(tStr)
60210      Else
60220       If UseStandard Then
60230        .PDFFontsSubSetFonts = 1
60240       End If
60250     End If
60260    Else
60270     If UseStandard Then
60280      .PDFFontsSubSetFonts = 1
60290     End If
60300   End If
60310   tStr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
60320   If IsNumeric(tStr) Then
60330     If CLng(tStr) >= 0 Then
60340       .PDFFontsSubSetFontsPercent = CLng(tStr)
60350      Else
60360       If UseStandard Then
60370        .PDFFontsSubSetFontsPercent = 100
60380       End If
60390     End If
60400    Else
60410     If UseStandard Then
60420      .PDFFontsSubSetFontsPercent = 100
60430     End If
60440   End If
60450   tStr = hOpt.Retrieve("PDFGeneralASCII85")
60460   If IsNumeric(tStr) Then
60470     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60480       .PDFGeneralASCII85 = CLng(tStr)
60490      Else
60500       If UseStandard Then
60510        .PDFGeneralASCII85 = 0
60520       End If
60530     End If
60540    Else
60550     If UseStandard Then
60560      .PDFGeneralASCII85 = 0
60570     End If
60580   End If
60590   tStr = hOpt.Retrieve("PDFGeneralAutorotate")
60600   If IsNumeric(tStr) Then
60610     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
60620       .PDFGeneralAutorotate = CLng(tStr)
60630      Else
60640       If UseStandard Then
60650        .PDFGeneralAutorotate = 2
60660       End If
60670     End If
60680    Else
60690     If UseStandard Then
60700      .PDFGeneralAutorotate = 2
60710     End If
60720   End If
60730   tStr = hOpt.Retrieve("PDFGeneralCompatibility")
60740   If IsNumeric(tStr) Then
60750     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
60760       .PDFGeneralCompatibility = CLng(tStr)
60770      Else
60780       If UseStandard Then
60790        .PDFGeneralCompatibility = 1
60800       End If
60810     End If
60820    Else
60830     If UseStandard Then
60840      .PDFGeneralCompatibility = 1
60850     End If
60860   End If
60870   tStr = hOpt.Retrieve("PDFGeneralOverprint")
60880   If IsNumeric(tStr) Then
60890     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
60900       .PDFGeneralOverprint = CLng(tStr)
60910      Else
60920       If UseStandard Then
60930        .PDFGeneralOverprint = 0
60940       End If
60950     End If
60960    Else
60970     If UseStandard Then
60980      .PDFGeneralOverprint = 0
60990     End If
61000   End If
61010   tStr = hOpt.Retrieve("PDFGeneralResolution")
61020   If IsNumeric(tStr) Then
61030     If CLng(tStr) >= 0 Then
61040       .PDFGeneralResolution = CLng(tStr)
61050      Else
61060       If UseStandard Then
61070        .PDFGeneralResolution = 600
61080       End If
61090     End If
61100    Else
61110     If UseStandard Then
61120      .PDFGeneralResolution = 600
61130     End If
61140   End If
61150   tStr = hOpt.Retrieve("PDFHighEncryption")
61160   If IsNumeric(tStr) Then
61170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61180       .PDFHighEncryption = CLng(tStr)
61190      Else
61200       If UseStandard Then
61210        .PDFHighEncryption = 0
61220       End If
61230     End If
61240    Else
61250     If UseStandard Then
61260      .PDFHighEncryption = 0
61270     End If
61280   End If
61290   tStr = hOpt.Retrieve("PDFLowEncryption")
61300   If IsNumeric(tStr) Then
61310     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61320       .PDFLowEncryption = CLng(tStr)
61330      Else
61340       If UseStandard Then
61350        .PDFLowEncryption = 1
61360       End If
61370     End If
61380    Else
61390     If UseStandard Then
61400      .PDFLowEncryption = 1
61410     End If
61420   End If
61430   tStr = hOpt.Retrieve("PDFOptimize")
61440   If IsNumeric(tStr) Then
61450     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61460       .PDFOptimize = CLng(tStr)
61470      Else
61480       If UseStandard Then
61490        .PDFOptimize = 0
61500       End If
61510     End If
61520    Else
61530     If UseStandard Then
61540      .PDFOptimize = 0
61550     End If
61560   End If
61570   tStr = hOpt.Retrieve("PDFOwnerPass")
61580   If IsNumeric(tStr) Then
61590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61600       .PDFOwnerPass = CLng(tStr)
61610      Else
61620       If UseStandard Then
61630        .PDFOwnerPass = 0
61640       End If
61650     End If
61660    Else
61670     If UseStandard Then
61680      .PDFOwnerPass = 0
61690     End If
61700   End If
61710   tStr = hOpt.Retrieve("PDFOwnerPasswordString")
61720   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
61730     .PDFOwnerPasswordString = ""
61740    Else
61750     If LenB(tStr) > 0 Then
61760      .PDFOwnerPasswordString = tStr
61770     End If
61780   End If
61790   tStr = hOpt.Retrieve("PDFUserPass")
61800   If IsNumeric(tStr) Then
61810     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61820       .PDFUserPass = CLng(tStr)
61830      Else
61840       If UseStandard Then
61850        .PDFUserPass = 0
61860       End If
61870     End If
61880    Else
61890     If UseStandard Then
61900      .PDFUserPass = 0
61910     End If
61920   End If
61930   tStr = hOpt.Retrieve("PDFUserPasswordString")
61940   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
61950     .PDFUserPasswordString = ""
61960    Else
61970     If LenB(tStr) > 0 Then
61980      .PDFUserPasswordString = tStr
61990     End If
62000   End If
62010   tStr = hOpt.Retrieve("PDFUseSecurity")
62020   If IsNumeric(tStr) Then
62030     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62040       .PDFUseSecurity = CLng(tStr)
62050      Else
62060       If UseStandard Then
62070        .PDFUseSecurity = 0
62080       End If
62090     End If
62100    Else
62110     If UseStandard Then
62120      .PDFUseSecurity = 0
62130     End If
62140   End If
62150   tStr = hOpt.Retrieve("PNGColorscount")
62160   If IsNumeric(tStr) Then
62170     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
62180       .PNGColorscount = CLng(tStr)
62190      Else
62200       If UseStandard Then
62210        .PNGColorscount = 0
62220       End If
62230     End If
62240    Else
62250     If UseStandard Then
62260      .PNGColorscount = 0
62270     End If
62280   End If
62290   tStr = hOpt.Retrieve("PrintAfterSaving")
62300   If IsNumeric(tStr) Then
62310     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62320       .PrintAfterSaving = CLng(tStr)
62330      Else
62340       If UseStandard Then
62350        .PrintAfterSaving = 0
62360       End If
62370     End If
62380    Else
62390     If UseStandard Then
62400      .PrintAfterSaving = 0
62410     End If
62420   End If
62430   tStr = hOpt.Retrieve("PrintAfterSavingDuplex")
62440   If IsNumeric(tStr) Then
62450     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62460       .PrintAfterSavingDuplex = CLng(tStr)
62470      Else
62480       If UseStandard Then
62490        .PrintAfterSavingDuplex = 0
62500       End If
62510     End If
62520    Else
62530     If UseStandard Then
62540      .PrintAfterSavingDuplex = 0
62550     End If
62560   End If
62570   tStr = hOpt.Retrieve("PrintAfterSavingNoCancel")
62580   If IsNumeric(tStr) Then
62590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62600       .PrintAfterSavingNoCancel = CLng(tStr)
62610      Else
62620       If UseStandard Then
62630        .PrintAfterSavingNoCancel = 0
62640       End If
62650     End If
62660    Else
62670     If UseStandard Then
62680      .PrintAfterSavingNoCancel = 0
62690     End If
62700   End If
62710   tStr = hOpt.Retrieve("PrintAfterSavingPrinter")
62720   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62730     .PrintAfterSavingPrinter = ""
62740    Else
62750     If LenB(tStr) > 0 Then
62760      .PrintAfterSavingPrinter = tStr
62770     End If
62780   End If
62790   tStr = hOpt.Retrieve("PrintAfterSavingQueryUser")
62800   If IsNumeric(tStr) Then
62810     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
62820       .PrintAfterSavingQueryUser = CLng(tStr)
62830      Else
62840       If UseStandard Then
62850        .PrintAfterSavingQueryUser = 0
62860       End If
62870     End If
62880    Else
62890     If UseStandard Then
62900      .PrintAfterSavingQueryUser = 0
62910     End If
62920   End If
62930   tStr = hOpt.Retrieve("PrintAfterSavingTumble")
62940   If IsNumeric(tStr) Then
62950     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62960       .PrintAfterSavingTumble = CLng(tStr)
62970      Else
62980       If UseStandard Then
62990        .PrintAfterSavingTumble = 0
63000       End If
63010     End If
63020    Else
63030     If UseStandard Then
63040      .PrintAfterSavingTumble = 0
63050     End If
63060   End If
63070   tStr = hOpt.Retrieve("PrinterStop")
63080   If IsNumeric(tStr) Then
63090     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63100       .PrinterStop = CLng(tStr)
63110      Else
63120       If UseStandard Then
63130        .PrinterStop = 0
63140       End If
63150     End If
63160    Else
63170     If UseStandard Then
63180      .PrinterStop = 0
63190     End If
63200   End If
63210   tStr = hOpt.Retrieve("PrinterTemppath")
63220   WriteToSpecialLogfile "hOpt.Retrieve(""PrinterTemppath"")=" & tStr
63230   WriteToSpecialLogfile "Options.PrinterTemppath1=" & .PrinterTemppath
63240   If hkey1 = HKEY_USERS Then
63250     If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then
63260       .PrinterTemppath = tStr
63270      Else
63280       If UseStandard Then
63290         .PrinterTemppath = GetTempPath
63300        Else
63310         .PrinterTemppath = tStr
63320       End If
63330     End If
63340    Else
63350     If LenB(Trim$(tStr)) > 0 Then
63360      If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then
63370        .PrinterTemppath = tStr
63380       Else
63390        MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))
63400        If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then
63410          If UseStandard Then
63420            .PrinterTemppath = GetTempPath
63430           Else
63440            .PrinterTemppath = ""
63450            If NoMsg = False Then
63460             MsgBox "PrinterTemppath: '" & tStr & "' = '" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & "'" & _
             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
63480            End If
63490          End If
63500         Else
63510          .PrinterTemppath = tStr
63520        End If
63530      End If
63540     End If
63550   End If
63560   WriteToSpecialLogfile "Options.PrinterTemppath2=" & .PrinterTemppath
63570   tStr = hOpt.Retrieve("ProcessPriority")
63580   If IsNumeric(tStr) Then
63590     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
63600       .ProcessPriority = CLng(tStr)
63610      Else
63620       If UseStandard Then
63630        .ProcessPriority = 1
63640       End If
63650     End If
63660    Else
63670     If UseStandard Then
63680      .ProcessPriority = 1
63690     End If
63700   End If
63710   tStr = hOpt.Retrieve("ProgramFont")
63720   If LenB(tStr) = 0 And LenB("MS Sans Serif") > 0 And UseStandard Then
63730     .ProgramFont = "MS Sans Serif"
63740    Else
63750     If LenB(tStr) > 0 Then
63760      .ProgramFont = tStr
63770     End If
63780   End If
63790   tStr = hOpt.Retrieve("ProgramFontCharset")
63800   If IsNumeric(tStr) Then
63810     If CLng(tStr) >= 0 Then
63820       .ProgramFontCharset = CLng(tStr)
63830      Else
63840       If UseStandard Then
63850        .ProgramFontCharset = 0
63860       End If
63870     End If
63880    Else
63890     If UseStandard Then
63900      .ProgramFontCharset = 0
63910     End If
63920   End If
63930   tStr = hOpt.Retrieve("ProgramFontSize")
63940   If IsNumeric(tStr) Then
63950     If CLng(tStr) >= 1 And CLng(tStr) <= 72 Then
63960       .ProgramFontSize = CLng(tStr)
63970      Else
63980       If UseStandard Then
63990        .ProgramFontSize = 8
64000       End If
64010     End If
64020    Else
64030     If UseStandard Then
64040      .ProgramFontSize = 8
64050     End If
64060   End If
64070   tStr = hOpt.Retrieve("PSLanguageLevel")
64080   If IsNumeric(tStr) Then
64090     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
64100       .PSLanguageLevel = CLng(tStr)
64110      Else
64120       If UseStandard Then
64130        .PSLanguageLevel = 2
64140       End If
64150     End If
64160    Else
64170     If UseStandard Then
64180      .PSLanguageLevel = 2
64190     End If
64200   End If
64210   tStr = hOpt.Retrieve("RemoveAllKnownFileExtensions")
64220   If IsNumeric(tStr) Then
64230     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64240       .RemoveAllKnownFileExtensions = CLng(tStr)
64250      Else
64260       If UseStandard Then
64270        .RemoveAllKnownFileExtensions = 1
64280       End If
64290     End If
64300    Else
64310     If UseStandard Then
64320      .RemoveAllKnownFileExtensions = 1
64330     End If
64340   End If
64350   tStr = hOpt.Retrieve("RemoveSpaces")
64360   If IsNumeric(tStr) Then
64370     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64380       .RemoveSpaces = CLng(tStr)
64390      Else
64400       If UseStandard Then
64410        .RemoveSpaces = 1
64420       End If
64430     End If
64440    Else
64450     If UseStandard Then
64460      .RemoveSpaces = 1
64470     End If
64480   End If
64490   tStr = hOpt.Retrieve("RunProgramAfterSaving")
64500   If IsNumeric(tStr) Then
64510     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64520       .RunProgramAfterSaving = CLng(tStr)
64530      Else
64540       If UseStandard Then
64550        .RunProgramAfterSaving = 0
64560       End If
64570     End If
64580    Else
64590     If UseStandard Then
64600      .RunProgramAfterSaving = 0
64610     End If
64620   End If
64630   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramname")
64640   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64650     .RunProgramAfterSavingProgramname = ""
64660    Else
64670     If LenB(tStr) > 0 Then
64680      .RunProgramAfterSavingProgramname = tStr
64690     End If
64700   End If
64710   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramParameters")
64720   If LenB(tStr) = 0 And LenB("""<OutputFilename>""") > 0 And UseStandard Then
64730     .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
64740    Else
64750     If LenB(tStr) > 0 Then
64760      .RunProgramAfterSavingProgramParameters = tStr
64770     End If
64780   End If
64790   tStr = hOpt.Retrieve("RunProgramAfterSavingWaitUntilReady")
64800   If IsNumeric(tStr) Then
64810     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64820       .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
64830      Else
64840       If UseStandard Then
64850        .RunProgramAfterSavingWaitUntilReady = 1
64860       End If
64870     End If
64880    Else
64890     If UseStandard Then
64900      .RunProgramAfterSavingWaitUntilReady = 1
64910     End If
64920   End If
64930   tStr = hOpt.Retrieve("RunProgramAfterSavingWindowstyle")
64940   If IsNumeric(tStr) Then
64950     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
64960       .RunProgramAfterSavingWindowstyle = CLng(tStr)
64970      Else
64980       If UseStandard Then
64990        .RunProgramAfterSavingWindowstyle = 1
65000       End If
65010     End If
65020    Else
65030     If UseStandard Then
65040      .RunProgramAfterSavingWindowstyle = 1
65050     End If
65060   End If
65070   tStr = hOpt.Retrieve("RunProgramBeforeSaving")
65080   If IsNumeric(tStr) Then
65090     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65100       .RunProgramBeforeSaving = CLng(tStr)
65110      Else
65120       If UseStandard Then
65130        .RunProgramBeforeSaving = 0
65140       End If
65150     End If
65160    Else
65170     If UseStandard Then
65180      .RunProgramBeforeSaving = 0
65190     End If
65200   End If
65210   tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramname")
65220   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
65230     .RunProgramBeforeSavingProgramname = ""
65240    Else
65250     If LenB(tStr) > 0 Then
65260      .RunProgramBeforeSavingProgramname = tStr
65270     End If
65280   End If
65290   tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramParameters")
65300   If LenB(tStr) = 0 And LenB("""<TempFilename>""") > 0 And UseStandard Then
65310     .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
65320    Else
65330     If LenB(tStr) > 0 Then
65340      .RunProgramBeforeSavingProgramParameters = tStr
65350     End If
65360   End If
65370   tStr = hOpt.Retrieve("RunProgramBeforeSavingWindowstyle")
65380   If IsNumeric(tStr) Then
65390     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
65400       .RunProgramBeforeSavingWindowstyle = CLng(tStr)
65410      Else
65420       If UseStandard Then
65430        .RunProgramBeforeSavingWindowstyle = 1
65440       End If
65450     End If
65460    Else
65470     If UseStandard Then
65480      .RunProgramBeforeSavingWindowstyle = 1
65490     End If
65500   End If
65510   tStr = hOpt.Retrieve("SaveFilename")
65520   If LenB(tStr) = 0 And LenB("<Title>") > 0 And UseStandard Then
65530     .SaveFilename = "<Title>"
65540    Else
65550     If LenB(tStr) > 0 Then
65560      .SaveFilename = tStr
65570     End If
65580   End If
65590   tStr = hOpt.Retrieve("SendEmailAfterAutoSaving")
65600   If IsNumeric(tStr) Then
65610     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65620       .SendEmailAfterAutoSaving = CLng(tStr)
65630      Else
65640       If UseStandard Then
65650        .SendEmailAfterAutoSaving = 0
65660       End If
65670     End If
65680    Else
65690     If UseStandard Then
65700      .SendEmailAfterAutoSaving = 0
65710     End If
65720   End If
65730   tStr = hOpt.Retrieve("SendMailMethod")
65740   If IsNumeric(tStr) Then
65750     If CLng(tStr) >= 0 Then
65760       .SendMailMethod = CLng(tStr)
65770      Else
65780       If UseStandard Then
65790        .SendMailMethod = 0
65800       End If
65810     End If
65820    Else
65830     If UseStandard Then
65840      .SendMailMethod = 0
65850     End If
65860   End If
65870   tStr = hOpt.Retrieve("ShowAnimation")
65880   If IsNumeric(tStr) Then
65890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65900       .ShowAnimation = CLng(tStr)
65910      Else
65920       If UseStandard Then
65930        .ShowAnimation = 1
65940       End If
65950     End If
65960    Else
65970     If UseStandard Then
65980      .ShowAnimation = 1
65990     End If
66000   End If
66010   tStr = hOpt.Retrieve("StampFontColor")
66020   If LenB(tStr) = 0 And LenB("#FF0000") > 0 And UseStandard Then
66030     .StampFontColor = "#FF0000"
66040    Else
66050     If LenB(tStr) > 0 Then
66060      .StampFontColor = tStr
66070     End If
66080   End If
66090   tStr = hOpt.Retrieve("StampFontname")
66100   If LenB(tStr) = 0 And LenB("Arial") > 0 And UseStandard Then
66110     .StampFontname = "Arial"
66120    Else
66130     If LenB(tStr) > 0 Then
66140      .StampFontname = tStr
66150     End If
66160   End If
66170   tStr = hOpt.Retrieve("StampFontsize")
66180   If IsNumeric(tStr) Then
66190     If CLng(tStr) >= 1 Then
66200       .StampFontsize = CLng(tStr)
66210      Else
66220       If UseStandard Then
66230        .StampFontsize = 48
66240       End If
66250     End If
66260    Else
66270     If UseStandard Then
66280      .StampFontsize = 48
66290     End If
66300   End If
66310   tStr = hOpt.Retrieve("StampOutlineFontthickness")
66320   If IsNumeric(tStr) Then
66330     If CLng(tStr) >= 0 Then
66340       .StampOutlineFontthickness = CLng(tStr)
66350      Else
66360       If UseStandard Then
66370        .StampOutlineFontthickness = 0
66380       End If
66390     End If
66400    Else
66410     If UseStandard Then
66420      .StampOutlineFontthickness = 0
66430     End If
66440   End If
66450   tStr = hOpt.Retrieve("StampString")
66460   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
66470     .StampString = ""
66480    Else
66490     If LenB(tStr) > 0 Then
66500      .StampString = tStr
66510     End If
66520   End If
66530   tStr = hOpt.Retrieve("StampUseOutlineFont")
66540   If IsNumeric(tStr) Then
66550     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66560       .StampUseOutlineFont = CLng(tStr)
66570      Else
66580       If UseStandard Then
66590        .StampUseOutlineFont = 1
66600       End If
66610     End If
66620    Else
66630     If UseStandard Then
66640      .StampUseOutlineFont = 1
66650     End If
66660   End If
66670   tStr = hOpt.Retrieve("StandardAuthor")
66680   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
66690     .StandardAuthor = ""
66700    Else
66710     If LenB(tStr) > 0 Then
66720      .StandardAuthor = tStr
66730     End If
66740   End If
66750   tStr = hOpt.Retrieve("StandardCreationdate")
66760   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
66770     .StandardCreationdate = ""
66780    Else
66790     If LenB(tStr) > 0 Then
66800      .StandardCreationdate = tStr
66810     End If
66820   End If
66830   tStr = hOpt.Retrieve("StandardDateformat")
66840   If LenB(tStr) = 0 And LenB("YYYYMMDDHHNNSS") > 0 And UseStandard Then
66850     .StandardDateformat = "YYYYMMDDHHNNSS"
66860    Else
66870     If LenB(tStr) > 0 Then
66880      .StandardDateformat = tStr
66890     End If
66900   End If
66910   tStr = hOpt.Retrieve("StandardKeywords")
66920   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
66930     .StandardKeywords = ""
66940    Else
66950     If LenB(tStr) > 0 Then
66960      .StandardKeywords = tStr
66970     End If
66980   End If
66990   tStr = hOpt.Retrieve("StandardMailDomain")
67000   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67010     .StandardMailDomain = ""
67020    Else
67030     If LenB(tStr) > 0 Then
67040      .StandardMailDomain = tStr
67050     End If
67060   End If
67070   tStr = hOpt.Retrieve("StandardModifydate")
67080   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67090     .StandardModifydate = ""
67100    Else
67110     If LenB(tStr) > 0 Then
67120      .StandardModifydate = tStr
67130     End If
67140   End If
67150   tStr = hOpt.Retrieve("StandardSaveformat")
67160   If LenB(tStr) = 0 And LenB("pdf") > 0 And UseStandard Then
67170     .StandardSaveformat = "pdf"
67180    Else
67190     If LenB(tStr) > 0 Then
67200      .StandardSaveformat = tStr
67210     End If
67220   End If
67230   tStr = hOpt.Retrieve("StandardSubject")
67240   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67250     .StandardSubject = ""
67260    Else
67270     If LenB(tStr) > 0 Then
67280      .StandardSubject = tStr
67290     End If
67300   End If
67310   tStr = hOpt.Retrieve("StandardTitle")
67320   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67330     .StandardTitle = ""
67340    Else
67350     If LenB(tStr) > 0 Then
67360      .StandardTitle = tStr
67370     End If
67380   End If
67390   tStr = hOpt.Retrieve("StartStandardProgram")
67400   If IsNumeric(tStr) Then
67410     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67420       .StartStandardProgram = CLng(tStr)
67430      Else
67440       If UseStandard Then
67450        .StartStandardProgram = 1
67460       End If
67470     End If
67480    Else
67490     If UseStandard Then
67500      .StartStandardProgram = 1
67510     End If
67520   End If
67530   tStr = hOpt.Retrieve("TIFFColorscount")
67540   If IsNumeric(tStr) Then
67550     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
67560       .TIFFColorscount = CLng(tStr)
67570      Else
67580       If UseStandard Then
67590        .TIFFColorscount = 0
67600       End If
67610     End If
67620    Else
67630     If UseStandard Then
67640      .TIFFColorscount = 0
67650     End If
67660   End If
67670   tStr = hOpt.Retrieve("Toolbars")
67680   If IsNumeric(tStr) Then
67690     If CLng(tStr) >= 0 Then
67700       .Toolbars = CLng(tStr)
67710      Else
67720       If UseStandard Then
67730        .Toolbars = 1
67740       End If
67750     End If
67760    Else
67770     If UseStandard Then
67780      .Toolbars = 1
67790     End If
67800   End If
67810   tStr = hOpt.Retrieve("UseAutosave")
67820   If IsNumeric(tStr) Then
67830     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67840       .UseAutosave = CLng(tStr)
67850      Else
67860       If UseStandard Then
67870        .UseAutosave = 0
67880       End If
67890     End If
67900    Else
67910     If UseStandard Then
67920      .UseAutosave = 0
67930     End If
67940   End If
67950   tStr = hOpt.Retrieve("UseAutosaveDirectory")
67960   If IsNumeric(tStr) Then
67970     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67980       .UseAutosaveDirectory = CLng(tStr)
67990      Else
68000       If UseStandard Then
68010        .UseAutosaveDirectory = 1
68020       End If
68030     End If
68040    Else
68050     If UseStandard Then
68060      .UseAutosaveDirectory = 1
68070     End If
68080   End If
68090   tStr = hOpt.Retrieve("UseCreationDateNow")
68100   If IsNumeric(tStr) Then
68110     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68120       .UseCreationDateNow = CLng(tStr)
68130      Else
68140       If UseStandard Then
68150        .UseCreationDateNow = 0
68160       End If
68170     End If
68180    Else
68190     If UseStandard Then
68200      .UseCreationDateNow = 0
68210     End If
68220   End If
68230   tStr = hOpt.Retrieve("UseStandardAuthor")
68240   If IsNumeric(tStr) Then
68250     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68260       .UseStandardAuthor = CLng(tStr)
68270      Else
68280       If UseStandard Then
68290        .UseStandardAuthor = 0
68300       End If
68310     End If
68320    Else
68330     If UseStandard Then
68340      .UseStandardAuthor = 0
68350     End If
68360   End If
68370  End With
68380  Set ini = Nothing
68390  ReadOptionsINI = myOptions
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
50160   Case "BITMAPRESOLUTION": ini.SaveKey CStr(.BitmapResolution), "BitmapResolution"
50170   Case "BMPCOLORSCOUNT": ini.SaveKey CStr(.BMPColorscount), "BMPColorscount"
50180   Case "CLIENTCOMPUTERRESOLVEIPADDRESS": ini.SaveKey CStr(Abs(.ClientComputerResolveIPAddress)), "ClientComputerResolveIPAddress"
50190   Case "DEVICEHEIGHTPOINTS": ini.SaveKey Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), "DeviceHeightPoints"
50200   Case "DEVICEWIDTHPOINTS": ini.SaveKey Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), "DeviceWidthPoints"
50210   Case "DIRECTORYGHOSTSCRIPTBINARIES": ini.SaveKey CStr(.DirectoryGhostscriptBinaries), "DirectoryGhostscriptBinaries"
50220   Case "DIRECTORYGHOSTSCRIPTFONTS": ini.SaveKey CStr(.DirectoryGhostscriptFonts), "DirectoryGhostscriptFonts"
50230   Case "DIRECTORYGHOSTSCRIPTLIBRARIES": ini.SaveKey CStr(.DirectoryGhostscriptLibraries), "DirectoryGhostscriptLibraries"
50240   Case "DIRECTORYGHOSTSCRIPTRESOURCE": ini.SaveKey CStr(.DirectoryGhostscriptResource), "DirectoryGhostscriptResource"
50250   Case "DISABLEEMAIL": ini.SaveKey CStr(Abs(.DisableEmail)), "DisableEmail"
50260   Case "DONTUSEDOCUMENTSETTINGS": ini.SaveKey CStr(Abs(.DontUseDocumentSettings)), "DontUseDocumentSettings"
50270   Case "EPSLANGUAGELEVEL": ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
50280   Case "FILENAMESUBSTITUTIONS": ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
50290   Case "FILENAMESUBSTITUTIONSONLYINTITLE": ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
50300   Case "JPEGCOLORSCOUNT": ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50310   Case "JPEGQUALITY": ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50320   Case "LANGUAGE": ini.SaveKey CStr(.Language), "Language"
50330   Case "LASTSAVEDIRECTORY": ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50340   Case "LOGGING": ini.SaveKey CStr(Abs(.Logging)), "Logging"
50350   Case "LOGLINES": ini.SaveKey CStr(.LogLines), "LogLines"
50360   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50370   Case "NOPROCESSINGATSTARTUP": ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
50380   Case "NOPSCHECK": ini.SaveKey CStr(Abs(.NoPSCheck)), "NoPSCheck"
50390   Case "ONEPAGEPERFILE": ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50400   Case "OPTIONSDESIGN": ini.SaveKey CStr(.OptionsDesign), "OptionsDesign"
50410   Case "OPTIONSENABLED": ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50420   Case "OPTIONSVISIBLE": ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50430   Case "PAPERSIZE": ini.SaveKey CStr(.Papersize), "Papersize"
50440   Case "PCXCOLORSCOUNT": ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50450   Case "PDFALLOWASSEMBLY": ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50460   Case "PDFALLOWDEGRADEDPRINTING": ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50470   Case "PDFALLOWFILLIN": ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50480   Case "PDFALLOWSCREENREADERS": ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50490   Case "PDFCOLORSCMYKTORGB": ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50500   Case "PDFCOLORSCOLORMODEL": ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50510   Case "PDFCOLORSPRESERVEHALFTONE": ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50520   Case "PDFCOLORSPRESERVEOVERPRINT": ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50530   Case "PDFCOLORSPRESERVETRANSFER": ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50540   Case "PDFCOMPRESSIONCOLORCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50550   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50560   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGHighFactor"
50570   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGLowFactor"
50580   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMaximumFactor"
50590   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMediumFactor"
50600   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMinimumFactor"
50610   Case "PDFCOMPRESSIONCOLORRESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50620   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50630   Case "PDFCOMPRESSIONCOLORRESOLUTION": ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50640   Case "PDFCOMPRESSIONGREYCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50650   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50660   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGHighFactor"
50670   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGLowFactor"
50680   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMaximumFactor"
50690   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMediumFactor"
50700   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMinimumFactor"
50710   Case "PDFCOMPRESSIONGREYRESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50720   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50730   Case "PDFCOMPRESSIONGREYRESOLUTION": ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50740   Case "PDFCOMPRESSIONMONOCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50750   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50760   Case "PDFCOMPRESSIONMONORESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50770   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50780   Case "PDFCOMPRESSIONMONORESOLUTION": ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50790   Case "PDFCOMPRESSIONTEXTCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50800   Case "PDFDISALLOWCOPY": ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50810   Case "PDFDISALLOWMODIFYANNOTATIONS": ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50820   Case "PDFDISALLOWMODIFYCONTENTS": ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50830   Case "PDFDISALLOWPRINTING": ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50840   Case "PDFENCRYPTOR": ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50850   Case "PDFFONTSEMBEDALL": ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50860   Case "PDFFONTSSUBSETFONTS": ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50870   Case "PDFFONTSSUBSETFONTSPERCENT": ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50880   Case "PDFGENERALASCII85": ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50890   Case "PDFGENERALAUTOROTATE": ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50900   Case "PDFGENERALCOMPATIBILITY": ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50910   Case "PDFGENERALOVERPRINT": ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50920   Case "PDFGENERALRESOLUTION": ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
50930   Case "PDFHIGHENCRYPTION": ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
50940   Case "PDFLOWENCRYPTION": ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
50950   Case "PDFOPTIMIZE": ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
50960   Case "PDFOWNERPASS": ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
50970   Case "PDFOWNERPASSWORDSTRING": ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
50980   Case "PDFUSERPASS": ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
50990   Case "PDFUSERPASSWORDSTRING": ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
51000   Case "PDFUSESECURITY": ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
51010   Case "PNGCOLORSCOUNT": ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
51020   Case "PRINTAFTERSAVING": ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
51030   Case "PRINTAFTERSAVINGDUPLEX": ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
51040   Case "PRINTAFTERSAVINGNOCANCEL": ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
51050   Case "PRINTAFTERSAVINGPRINTER": ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
51060   Case "PRINTAFTERSAVINGQUERYUSER": ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
51070   Case "PRINTAFTERSAVINGTUMBLE": ini.SaveKey CStr(Abs(.PrintAfterSavingTumble)), "PrintAfterSavingTumble"
51080   Case "PRINTERSTOP": ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
51090   Case "PRINTERTEMPPATH": ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
51100   Case "PROCESSPRIORITY": ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
51110   Case "PROGRAMFONT": ini.SaveKey CStr(.ProgramFont), "ProgramFont"
51120   Case "PROGRAMFONTCHARSET": ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
51130   Case "PROGRAMFONTSIZE": ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
51140   Case "PSLANGUAGELEVEL": ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
51150   Case "REMOVEALLKNOWNFILEEXTENSIONS": ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
51160   Case "REMOVESPACES": ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
51170   Case "RUNPROGRAMAFTERSAVING": ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
51180   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
51190   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
51200   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51210   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51220   Case "RUNPROGRAMBEFORESAVING": ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51230   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51240   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51250   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51260   Case "SAVEFILENAME": ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51270   Case "SENDEMAILAFTERAUTOSAVING": ini.SaveKey CStr(Abs(.SendEmailAfterAutoSaving)), "SendEmailAfterAutoSaving"
51280   Case "SENDMAILMETHOD": ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51290   Case "SHOWANIMATION": ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51300   Case "STAMPFONTCOLOR": ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51310   Case "STAMPFONTNAME": ini.SaveKey CStr(.StampFontname), "StampFontname"
51320   Case "STAMPFONTSIZE": ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51330   Case "STAMPOUTLINEFONTTHICKNESS": ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51340   Case "STAMPSTRING": ini.SaveKey CStr(.StampString), "StampString"
51350   Case "STAMPUSEOUTLINEFONT": ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51360   Case "STANDARDAUTHOR": ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51370   Case "STANDARDCREATIONDATE": ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51380   Case "STANDARDDATEFORMAT": ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51390   Case "STANDARDKEYWORDS": ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51400   Case "STANDARDMAILDOMAIN": ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51410   Case "STANDARDMODIFYDATE": ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51420   Case "STANDARDSAVEFORMAT": ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51430   Case "STANDARDSUBJECT": ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51440   Case "STANDARDTITLE": ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51450   Case "STARTSTANDARDPROGRAM": ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51460   Case "TIFFCOLORSCOUNT": ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51470   Case "TOOLBARS": ini.SaveKey CStr(.Toolbars), "Toolbars"
51480   Case "USEAUTOSAVE": ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51490   Case "USEAUTOSAVEDIRECTORY": ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51500   Case "USECREATIONDATENOW": ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51510   Case "USESTANDARDAUTHOR": ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51520   End Select
51530  End With
51540  Set ini = Nothing
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
50150   ini.SaveKey CStr(.BitmapResolution), "BitmapResolution"
50160   ini.SaveKey CStr(.BMPColorscount), "BMPColorscount"
50170   ini.SaveKey CStr(Abs(.ClientComputerResolveIPAddress)), "ClientComputerResolveIPAddress"
50180   ini.SaveKey Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), "DeviceHeightPoints"
50190   ini.SaveKey Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), "DeviceWidthPoints"
50200   ini.SaveKey CStr(.DirectoryGhostscriptBinaries), "DirectoryGhostscriptBinaries"
50210   ini.SaveKey CStr(.DirectoryGhostscriptFonts), "DirectoryGhostscriptFonts"
50220   ini.SaveKey CStr(.DirectoryGhostscriptLibraries), "DirectoryGhostscriptLibraries"
50230   ini.SaveKey CStr(.DirectoryGhostscriptResource), "DirectoryGhostscriptResource"
50240   ini.SaveKey CStr(Abs(.DisableEmail)), "DisableEmail"
50250   ini.SaveKey CStr(Abs(.DontUseDocumentSettings)), "DontUseDocumentSettings"
50260   ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
50270   ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
50280   ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
50290   ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50300   ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50310   ini.SaveKey CStr(.Language), "Language"
50320   ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50330   ini.SaveKey CStr(Abs(.Logging)), "Logging"
50340   ini.SaveKey CStr(.LogLines), "LogLines"
50350   ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50360   ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
50370   ini.SaveKey CStr(Abs(.NoPSCheck)), "NoPSCheck"
50380   ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50390   ini.SaveKey CStr(.OptionsDesign), "OptionsDesign"
50400   ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50410   ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50420   ini.SaveKey CStr(.Papersize), "Papersize"
50430   ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50440   ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50450   ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50460   ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50470   ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50480   ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50490   ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50500   ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50510   ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50520   ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50530   ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50540   ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50550   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGHighFactor"
50560   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGLowFactor"
50570   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMaximumFactor"
50580   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMediumFactor"
50590   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMinimumFactor"
50600   ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50610   ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50620   ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50630   ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50640   ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50650   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGHighFactor"
50660   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGLowFactor"
50670   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMaximumFactor"
50680   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMediumFactor"
50690   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMinimumFactor"
50700   ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50710   ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50720   ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50730   ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50740   ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50750   ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50760   ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50770   ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50780   ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50790   ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50800   ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50810   ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50820   ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50830   ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50840   ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50850   ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50860   ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50870   ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50880   ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50890   ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50900   ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50910   ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
50920   ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
50930   ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
50940   ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
50950   ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
50960   ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
50970   ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
50980   ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
50990   ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
51000   ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
51010   ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
51020   ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
51030   ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
51040   ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
51050   ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
51060   ini.SaveKey CStr(Abs(.PrintAfterSavingTumble)), "PrintAfterSavingTumble"
51070   ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
51080   ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
51090   ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
51100   ini.SaveKey CStr(.ProgramFont), "ProgramFont"
51110   ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
51120   ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
51130   ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
51140   ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
51150   ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
51160   ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
51170   ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
51180   ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
51190   ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51200   ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51210   ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51220   ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51230   ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51240   ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51250   ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51260   ini.SaveKey CStr(Abs(.SendEmailAfterAutoSaving)), "SendEmailAfterAutoSaving"
51270   ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51280   ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51290   ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51300   ini.SaveKey CStr(.StampFontname), "StampFontname"
51310   ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51320   ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51330   ini.SaveKey CStr(.StampString), "StampString"
51340   ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51350   ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51360   ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51370   ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51380   ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51390   ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51400   ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51410   ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51420   ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51430   ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51440   ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51450   ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51460   ini.SaveKey CStr(.Toolbars), "Toolbars"
51470   ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51480   ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51490   ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51500   ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51510  End With
51520  Set ini = Nothing
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
52080   If LenB(tStr) = 0 And LenB("pdf") > 0 And UseStandard Then
52090     .StandardSaveformat = "pdf"
52100    Else
52110     If LenB(tStr) > 0 Then
52120      .StandardSaveformat = tStr
52130     End If
52140   End If
52150   tStr = reg.GetRegistryValue("StandardSubject")
52160   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
52170     .StandardSubject = ""
52180    Else
52190     If LenB(tStr) > 0 Then
52200      .StandardSubject = tStr
52210     End If
52220   End If
52230   tStr = reg.GetRegistryValue("StandardTitle")
52240   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
52250     .StandardTitle = ""
52260    Else
52270     If LenB(tStr) > 0 Then
52280      .StandardTitle = tStr
52290     End If
52300   End If
52310   tStr = reg.GetRegistryValue("UseCreationDateNow")
52320   If IsNumeric(tStr) Then
52330     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52340       .UseCreationDateNow = CLng(tStr)
52350      Else
52360       If UseStandard Then
52370        .UseCreationDateNow = 0
52380       End If
52390     End If
52400    Else
52410     If UseStandard Then
52420      .UseCreationDateNow = 0
52430     End If
52440   End If
52450   tStr = reg.GetRegistryValue("UseStandardAuthor")
52460   If IsNumeric(tStr) Then
52470     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52480       .UseStandardAuthor = CLng(tStr)
52490      Else
52500       If UseStandard Then
52510        .UseStandardAuthor = 0
52520       End If
52530     End If
52540    Else
52550     If UseStandard Then
52560      .UseStandardAuthor = 0
52570     End If
52580   End If
52590   reg.Subkey = "Printing\Formats\Bitmap\Colors"
52600   tStr = reg.GetRegistryValue("BitmapResolution")
52610   If IsNumeric(tStr) Then
52620     If CLng(tStr) >= 1 Then
52630       .BitmapResolution = CLng(tStr)
52640      Else
52650       If UseStandard Then
52660        .BitmapResolution = 150
52670       End If
52680     End If
52690    Else
52700     If UseStandard Then
52710      .BitmapResolution = 150
52720     End If
52730   End If
52740   tStr = reg.GetRegistryValue("BMPColorscount")
52750   If IsNumeric(tStr) Then
52760     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
52770       .BMPColorscount = CLng(tStr)
52780      Else
52790       If UseStandard Then
52800        .BMPColorscount = 1
52810       End If
52820     End If
52830    Else
52840     If UseStandard Then
52850      .BMPColorscount = 1
52860     End If
52870   End If
52880   tStr = reg.GetRegistryValue("JPEGColorscount")
52890   If IsNumeric(tStr) Then
52900     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
52910       .JPEGColorscount = CLng(tStr)
52920      Else
52930       If UseStandard Then
52940        .JPEGColorscount = 0
52950       End If
52960     End If
52970    Else
52980     If UseStandard Then
52990      .JPEGColorscount = 0
53000     End If
53010   End If
53020   tStr = reg.GetRegistryValue("JPEGQuality")
53030   If IsNumeric(tStr) Then
53040     If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
53050       .JPEGQuality = CLng(tStr)
53060      Else
53070       If UseStandard Then
53080        .JPEGQuality = 75
53090       End If
53100     End If
53110    Else
53120     If UseStandard Then
53130      .JPEGQuality = 75
53140     End If
53150   End If
53160   tStr = reg.GetRegistryValue("PCXColorscount")
53170   If IsNumeric(tStr) Then
53180     If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
53190       .PCXColorscount = CLng(tStr)
53200      Else
53210       If UseStandard Then
53220        .PCXColorscount = 0
53230       End If
53240     End If
53250    Else
53260     If UseStandard Then
53270      .PCXColorscount = 0
53280     End If
53290   End If
53300   tStr = reg.GetRegistryValue("PNGColorscount")
53310   If IsNumeric(tStr) Then
53320     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
53330       .PNGColorscount = CLng(tStr)
53340      Else
53350       If UseStandard Then
53360        .PNGColorscount = 0
53370       End If
53380     End If
53390    Else
53400     If UseStandard Then
53410      .PNGColorscount = 0
53420     End If
53430   End If
53440   tStr = reg.GetRegistryValue("TIFFColorscount")
53450   If IsNumeric(tStr) Then
53460     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
53470       .TIFFColorscount = CLng(tStr)
53480      Else
53490       If UseStandard Then
53500        .TIFFColorscount = 0
53510       End If
53520     End If
53530    Else
53540     If UseStandard Then
53550      .TIFFColorscount = 0
53560     End If
53570   End If
53580   reg.Subkey = "Printing\Formats\PDF\Colors"
53590   tStr = reg.GetRegistryValue("PDFColorsCMYKToRGB")
53600   If IsNumeric(tStr) Then
53610     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53620       .PDFColorsCMYKToRGB = CLng(tStr)
53630      Else
53640       If UseStandard Then
53650        .PDFColorsCMYKToRGB = 0
53660       End If
53670     End If
53680    Else
53690     If UseStandard Then
53700      .PDFColorsCMYKToRGB = 0
53710     End If
53720   End If
53730   tStr = reg.GetRegistryValue("PDFColorsColorModel")
53740   If IsNumeric(tStr) Then
53750     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
53760       .PDFColorsColorModel = CLng(tStr)
53770      Else
53780       If UseStandard Then
53790        .PDFColorsColorModel = 1
53800       End If
53810     End If
53820    Else
53830     If UseStandard Then
53840      .PDFColorsColorModel = 1
53850     End If
53860   End If
53870   tStr = reg.GetRegistryValue("PDFColorsPreserveHalftone")
53880   If IsNumeric(tStr) Then
53890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53900       .PDFColorsPreserveHalftone = CLng(tStr)
53910      Else
53920       If UseStandard Then
53930        .PDFColorsPreserveHalftone = 0
53940       End If
53950     End If
53960    Else
53970     If UseStandard Then
53980      .PDFColorsPreserveHalftone = 0
53990     End If
54000   End If
54010   tStr = reg.GetRegistryValue("PDFColorsPreserveOverprint")
54020   If IsNumeric(tStr) Then
54030     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54040       .PDFColorsPreserveOverprint = CLng(tStr)
54050      Else
54060       If UseStandard Then
54070        .PDFColorsPreserveOverprint = 1
54080       End If
54090     End If
54100    Else
54110     If UseStandard Then
54120      .PDFColorsPreserveOverprint = 1
54130     End If
54140   End If
54150   tStr = reg.GetRegistryValue("PDFColorsPreserveTransfer")
54160   If IsNumeric(tStr) Then
54170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54180       .PDFColorsPreserveTransfer = CLng(tStr)
54190      Else
54200       If UseStandard Then
54210        .PDFColorsPreserveTransfer = 1
54220       End If
54230     End If
54240    Else
54250     If UseStandard Then
54260      .PDFColorsPreserveTransfer = 1
54270     End If
54280   End If
54290   reg.Subkey = "Printing\Formats\PDF\Compression"
54300   tStr = reg.GetRegistryValue("PDFCompressionColorCompression")
54310   If IsNumeric(tStr) Then
54320     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54330       .PDFCompressionColorCompression = CLng(tStr)
54340      Else
54350       If UseStandard Then
54360        .PDFCompressionColorCompression = 1
54370       End If
54380     End If
54390    Else
54400     If UseStandard Then
54410      .PDFCompressionColorCompression = 1
54420     End If
54430   End If
54440   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionChoice")
54450   If IsNumeric(tStr) Then
54460     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
54470       .PDFCompressionColorCompressionChoice = CLng(tStr)
54480      Else
54490       If UseStandard Then
54500        .PDFCompressionColorCompressionChoice = 0
54510       End If
54520     End If
54530    Else
54540     If UseStandard Then
54550      .PDFCompressionColorCompressionChoice = 0
54560     End If
54570   End If
54580   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGHighFactor")
54590   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
54600     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
54610       .PDFCompressionColorCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
54620      Else
54630       If UseStandard Then
54640        .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
54650       End If
54660     End If
54670    Else
54680     If UseStandard Then
54690      .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
54700     End If
54710   End If
54720   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGLowFactor")
54730   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
54740     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
54750       .PDFCompressionColorCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
54760      Else
54770       If UseStandard Then
54780        .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
54790       End If
54800     End If
54810    Else
54820     If UseStandard Then
54830      .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
54840     End If
54850   End If
54860   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMaximumFactor")
54870   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
54880     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
54890       .PDFCompressionColorCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
54900      Else
54910       If UseStandard Then
54920        .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
54930       End If
54940     End If
54950    Else
54960     If UseStandard Then
54970      .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
54980     End If
54990   End If
55000   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMediumFactor")
55010   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
55020     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
55030       .PDFCompressionColorCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
55040      Else
55050       If UseStandard Then
55060        .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
55070       End If
55080     End If
55090    Else
55100     If UseStandard Then
55110      .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
55120     End If
55130   End If
55140   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMinimumFactor")
55150   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
55160     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
55170       .PDFCompressionColorCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
55180      Else
55190       If UseStandard Then
55200        .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
55210       End If
55220     End If
55230    Else
55240     If UseStandard Then
55250      .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
55260     End If
55270   End If
55280   tStr = reg.GetRegistryValue("PDFCompressionColorResample")
55290   If IsNumeric(tStr) Then
55300     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55310       .PDFCompressionColorResample = CLng(tStr)
55320      Else
55330       If UseStandard Then
55340        .PDFCompressionColorResample = 0
55350       End If
55360     End If
55370    Else
55380     If UseStandard Then
55390      .PDFCompressionColorResample = 0
55400     End If
55410   End If
55420   tStr = reg.GetRegistryValue("PDFCompressionColorResampleChoice")
55430   If IsNumeric(tStr) Then
55440     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
55450       .PDFCompressionColorResampleChoice = CLng(tStr)
55460      Else
55470       If UseStandard Then
55480        .PDFCompressionColorResampleChoice = 0
55490       End If
55500     End If
55510    Else
55520     If UseStandard Then
55530      .PDFCompressionColorResampleChoice = 0
55540     End If
55550   End If
55560   tStr = reg.GetRegistryValue("PDFCompressionColorResolution")
55570   If IsNumeric(tStr) Then
55580     If CLng(tStr) >= 0 Then
55590       .PDFCompressionColorResolution = CLng(tStr)
55600      Else
55610       If UseStandard Then
55620        .PDFCompressionColorResolution = 300
55630       End If
55640     End If
55650    Else
55660     If UseStandard Then
55670      .PDFCompressionColorResolution = 300
55680     End If
55690   End If
55700   tStr = reg.GetRegistryValue("PDFCompressionGreyCompression")
55710   If IsNumeric(tStr) Then
55720     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55730       .PDFCompressionGreyCompression = CLng(tStr)
55740      Else
55750       If UseStandard Then
55760        .PDFCompressionGreyCompression = 1
55770       End If
55780     End If
55790    Else
55800     If UseStandard Then
55810      .PDFCompressionGreyCompression = 1
55820     End If
55830   End If
55840   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionChoice")
55850   If IsNumeric(tStr) Then
55860     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
55870       .PDFCompressionGreyCompressionChoice = CLng(tStr)
55880      Else
55890       If UseStandard Then
55900        .PDFCompressionGreyCompressionChoice = 0
55910       End If
55920     End If
55930    Else
55940     If UseStandard Then
55950      .PDFCompressionGreyCompressionChoice = 0
55960     End If
55970   End If
55980   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGHighFactor")
55990   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56000     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56010       .PDFCompressionGreyCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56020      Else
56030       If UseStandard Then
56040        .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56050       End If
56060     End If
56070    Else
56080     If UseStandard Then
56090      .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56100     End If
56110   End If
56120   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGLowFactor")
56130   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56140     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56150       .PDFCompressionGreyCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56160      Else
56170       If UseStandard Then
56180        .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56190       End If
56200     End If
56210    Else
56220     If UseStandard Then
56230      .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56240     End If
56250   End If
56260   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMaximumFactor")
56270   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56280     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56290       .PDFCompressionGreyCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56300      Else
56310       If UseStandard Then
56320        .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56330       End If
56340     End If
56350    Else
56360     If UseStandard Then
56370      .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56380     End If
56390   End If
56400   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMediumFactor")
56410   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56420     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56430       .PDFCompressionGreyCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56440      Else
56450       If UseStandard Then
56460        .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56470       End If
56480     End If
56490    Else
56500     If UseStandard Then
56510      .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
56520     End If
56530   End If
56540   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMinimumFactor")
56550   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56560     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56570       .PDFCompressionGreyCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56580      Else
56590       If UseStandard Then
56600        .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
56610       End If
56620     End If
56630    Else
56640     If UseStandard Then
56650      .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
56660     End If
56670   End If
56680   tStr = reg.GetRegistryValue("PDFCompressionGreyResample")
56690   If IsNumeric(tStr) Then
56700     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56710       .PDFCompressionGreyResample = CLng(tStr)
56720      Else
56730       If UseStandard Then
56740        .PDFCompressionGreyResample = 0
56750       End If
56760     End If
56770    Else
56780     If UseStandard Then
56790      .PDFCompressionGreyResample = 0
56800     End If
56810   End If
56820   tStr = reg.GetRegistryValue("PDFCompressionGreyResampleChoice")
56830   If IsNumeric(tStr) Then
56840     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
56850       .PDFCompressionGreyResampleChoice = CLng(tStr)
56860      Else
56870       If UseStandard Then
56880        .PDFCompressionGreyResampleChoice = 0
56890       End If
56900     End If
56910    Else
56920     If UseStandard Then
56930      .PDFCompressionGreyResampleChoice = 0
56940     End If
56950   End If
56960   tStr = reg.GetRegistryValue("PDFCompressionGreyResolution")
56970   If IsNumeric(tStr) Then
56980     If CLng(tStr) >= 0 Then
56990       .PDFCompressionGreyResolution = CLng(tStr)
57000      Else
57010       If UseStandard Then
57020        .PDFCompressionGreyResolution = 300
57030       End If
57040     End If
57050    Else
57060     If UseStandard Then
57070      .PDFCompressionGreyResolution = 300
57080     End If
57090   End If
57100   tStr = reg.GetRegistryValue("PDFCompressionMonoCompression")
57110   If IsNumeric(tStr) Then
57120     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57130       .PDFCompressionMonoCompression = CLng(tStr)
57140      Else
57150       If UseStandard Then
57160        .PDFCompressionMonoCompression = 1
57170       End If
57180     End If
57190    Else
57200     If UseStandard Then
57210      .PDFCompressionMonoCompression = 1
57220     End If
57230   End If
57240   tStr = reg.GetRegistryValue("PDFCompressionMonoCompressionChoice")
57250   If IsNumeric(tStr) Then
57260     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
57270       .PDFCompressionMonoCompressionChoice = CLng(tStr)
57280      Else
57290       If UseStandard Then
57300        .PDFCompressionMonoCompressionChoice = 0
57310       End If
57320     End If
57330    Else
57340     If UseStandard Then
57350      .PDFCompressionMonoCompressionChoice = 0
57360     End If
57370   End If
57380   tStr = reg.GetRegistryValue("PDFCompressionMonoResample")
57390   If IsNumeric(tStr) Then
57400     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57410       .PDFCompressionMonoResample = CLng(tStr)
57420      Else
57430       If UseStandard Then
57440        .PDFCompressionMonoResample = 0
57450       End If
57460     End If
57470    Else
57480     If UseStandard Then
57490      .PDFCompressionMonoResample = 0
57500     End If
57510   End If
57520   tStr = reg.GetRegistryValue("PDFCompressionMonoResampleChoice")
57530   If IsNumeric(tStr) Then
57540     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
57550       .PDFCompressionMonoResampleChoice = CLng(tStr)
57560      Else
57570       If UseStandard Then
57580        .PDFCompressionMonoResampleChoice = 0
57590       End If
57600     End If
57610    Else
57620     If UseStandard Then
57630      .PDFCompressionMonoResampleChoice = 0
57640     End If
57650   End If
57660   tStr = reg.GetRegistryValue("PDFCompressionMonoResolution")
57670   If IsNumeric(tStr) Then
57680     If CLng(tStr) >= 0 Then
57690       .PDFCompressionMonoResolution = CLng(tStr)
57700      Else
57710       If UseStandard Then
57720        .PDFCompressionMonoResolution = 1200
57730       End If
57740     End If
57750    Else
57760     If UseStandard Then
57770      .PDFCompressionMonoResolution = 1200
57780     End If
57790   End If
57800   tStr = reg.GetRegistryValue("PDFCompressionTextCompression")
57810   If IsNumeric(tStr) Then
57820     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57830       .PDFCompressionTextCompression = CLng(tStr)
57840      Else
57850       If UseStandard Then
57860        .PDFCompressionTextCompression = 1
57870       End If
57880     End If
57890    Else
57900     If UseStandard Then
57910      .PDFCompressionTextCompression = 1
57920     End If
57930   End If
57940   reg.Subkey = "Printing\Formats\PDF\Fonts"
57950   tStr = reg.GetRegistryValue("PDFFontsEmbedAll")
57960   If IsNumeric(tStr) Then
57970     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57980       .PDFFontsEmbedAll = CLng(tStr)
57990      Else
58000       If UseStandard Then
58010        .PDFFontsEmbedAll = 1
58020       End If
58030     End If
58040    Else
58050     If UseStandard Then
58060      .PDFFontsEmbedAll = 1
58070     End If
58080   End If
58090   tStr = reg.GetRegistryValue("PDFFontsSubSetFonts")
58100   If IsNumeric(tStr) Then
58110     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58120       .PDFFontsSubSetFonts = CLng(tStr)
58130      Else
58140       If UseStandard Then
58150        .PDFFontsSubSetFonts = 1
58160       End If
58170     End If
58180    Else
58190     If UseStandard Then
58200      .PDFFontsSubSetFonts = 1
58210     End If
58220   End If
58230   tStr = reg.GetRegistryValue("PDFFontsSubSetFontsPercent")
58240   If IsNumeric(tStr) Then
58250     If CLng(tStr) >= 0 Then
58260       .PDFFontsSubSetFontsPercent = CLng(tStr)
58270      Else
58280       If UseStandard Then
58290        .PDFFontsSubSetFontsPercent = 100
58300       End If
58310     End If
58320    Else
58330     If UseStandard Then
58340      .PDFFontsSubSetFontsPercent = 100
58350     End If
58360   End If
58370   reg.Subkey = "Printing\Formats\PDF\General"
58380   tStr = reg.GetRegistryValue("PDFGeneralASCII85")
58390   If IsNumeric(tStr) Then
58400     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58410       .PDFGeneralASCII85 = CLng(tStr)
58420      Else
58430       If UseStandard Then
58440        .PDFGeneralASCII85 = 0
58450       End If
58460     End If
58470    Else
58480     If UseStandard Then
58490      .PDFGeneralASCII85 = 0
58500     End If
58510   End If
58520   tStr = reg.GetRegistryValue("PDFGeneralAutorotate")
58530   If IsNumeric(tStr) Then
58540     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
58550       .PDFGeneralAutorotate = CLng(tStr)
58560      Else
58570       If UseStandard Then
58580        .PDFGeneralAutorotate = 2
58590       End If
58600     End If
58610    Else
58620     If UseStandard Then
58630      .PDFGeneralAutorotate = 2
58640     End If
58650   End If
58660   tStr = reg.GetRegistryValue("PDFGeneralCompatibility")
58670   If IsNumeric(tStr) Then
58680     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
58690       .PDFGeneralCompatibility = CLng(tStr)
58700      Else
58710       If UseStandard Then
58720        .PDFGeneralCompatibility = 1
58730       End If
58740     End If
58750    Else
58760     If UseStandard Then
58770      .PDFGeneralCompatibility = 1
58780     End If
58790   End If
58800   tStr = reg.GetRegistryValue("PDFGeneralOverprint")
58810   If IsNumeric(tStr) Then
58820     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
58830       .PDFGeneralOverprint = CLng(tStr)
58840      Else
58850       If UseStandard Then
58860        .PDFGeneralOverprint = 0
58870       End If
58880     End If
58890    Else
58900     If UseStandard Then
58910      .PDFGeneralOverprint = 0
58920     End If
58930   End If
58940   tStr = reg.GetRegistryValue("PDFGeneralResolution")
58950   If IsNumeric(tStr) Then
58960     If CLng(tStr) >= 0 Then
58970       .PDFGeneralResolution = CLng(tStr)
58980      Else
58990       If UseStandard Then
59000        .PDFGeneralResolution = 600
59010       End If
59020     End If
59030    Else
59040     If UseStandard Then
59050      .PDFGeneralResolution = 600
59060     End If
59070   End If
59080   tStr = reg.GetRegistryValue("PDFOptimize")
59090   If IsNumeric(tStr) Then
59100     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59110       .PDFOptimize = CLng(tStr)
59120      Else
59130       If UseStandard Then
59140        .PDFOptimize = 0
59150       End If
59160     End If
59170    Else
59180     If UseStandard Then
59190      .PDFOptimize = 0
59200     End If
59210   End If
59220   reg.Subkey = "Printing\Formats\PDF\Security"
59230   tStr = reg.GetRegistryValue("PDFAllowAssembly")
59240   If IsNumeric(tStr) Then
59250     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59260       .PDFAllowAssembly = CLng(tStr)
59270      Else
59280       If UseStandard Then
59290        .PDFAllowAssembly = 0
59300       End If
59310     End If
59320    Else
59330     If UseStandard Then
59340      .PDFAllowAssembly = 0
59350     End If
59360   End If
59370   tStr = reg.GetRegistryValue("PDFAllowDegradedPrinting")
59380   If IsNumeric(tStr) Then
59390     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59400       .PDFAllowDegradedPrinting = CLng(tStr)
59410      Else
59420       If UseStandard Then
59430        .PDFAllowDegradedPrinting = 0
59440       End If
59450     End If
59460    Else
59470     If UseStandard Then
59480      .PDFAllowDegradedPrinting = 0
59490     End If
59500   End If
59510   tStr = reg.GetRegistryValue("PDFAllowFillIn")
59520   If IsNumeric(tStr) Then
59530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59540       .PDFAllowFillIn = CLng(tStr)
59550      Else
59560       If UseStandard Then
59570        .PDFAllowFillIn = 0
59580       End If
59590     End If
59600    Else
59610     If UseStandard Then
59620      .PDFAllowFillIn = 0
59630     End If
59640   End If
59650   tStr = reg.GetRegistryValue("PDFAllowScreenReaders")
59660   If IsNumeric(tStr) Then
59670     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59680       .PDFAllowScreenReaders = CLng(tStr)
59690      Else
59700       If UseStandard Then
59710        .PDFAllowScreenReaders = 0
59720       End If
59730     End If
59740    Else
59750     If UseStandard Then
59760      .PDFAllowScreenReaders = 0
59770     End If
59780   End If
59790   tStr = reg.GetRegistryValue("PDFDisallowCopy")
59800   If IsNumeric(tStr) Then
59810     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59820       .PDFDisallowCopy = CLng(tStr)
59830      Else
59840       If UseStandard Then
59850        .PDFDisallowCopy = 1
59860       End If
59870     End If
59880    Else
59890     If UseStandard Then
59900      .PDFDisallowCopy = 1
59910     End If
59920   End If
59930   tStr = reg.GetRegistryValue("PDFDisallowModifyAnnotations")
59940   If IsNumeric(tStr) Then
59950     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59960       .PDFDisallowModifyAnnotations = CLng(tStr)
59970      Else
59980       If UseStandard Then
59990        .PDFDisallowModifyAnnotations = 0
60000       End If
60010     End If
60020    Else
60030     If UseStandard Then
60040      .PDFDisallowModifyAnnotations = 0
60050     End If
60060   End If
60070   tStr = reg.GetRegistryValue("PDFDisallowModifyContents")
60080   If IsNumeric(tStr) Then
60090     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60100       .PDFDisallowModifyContents = CLng(tStr)
60110      Else
60120       If UseStandard Then
60130        .PDFDisallowModifyContents = 0
60140       End If
60150     End If
60160    Else
60170     If UseStandard Then
60180      .PDFDisallowModifyContents = 0
60190     End If
60200   End If
60210   tStr = reg.GetRegistryValue("PDFDisallowPrinting")
60220   If IsNumeric(tStr) Then
60230     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60240       .PDFDisallowPrinting = CLng(tStr)
60250      Else
60260       If UseStandard Then
60270        .PDFDisallowPrinting = 0
60280       End If
60290     End If
60300    Else
60310     If UseStandard Then
60320      .PDFDisallowPrinting = 0
60330     End If
60340   End If
60350   tStr = reg.GetRegistryValue("PDFEncryptor")
60360   If IsNumeric(tStr) Then
60370     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
60380       .PDFEncryptor = CLng(tStr)
60390      Else
60400       If UseStandard Then
60410        .PDFEncryptor = 0
60420       End If
60430     End If
60440    Else
60450     If UseStandard Then
60460      .PDFEncryptor = 0
60470     End If
60480   End If
60490   tStr = reg.GetRegistryValue("PDFHighEncryption")
60500   If IsNumeric(tStr) Then
60510     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60520       .PDFHighEncryption = CLng(tStr)
60530      Else
60540       If UseStandard Then
60550        .PDFHighEncryption = 0
60560       End If
60570     End If
60580    Else
60590     If UseStandard Then
60600      .PDFHighEncryption = 0
60610     End If
60620   End If
60630   tStr = reg.GetRegistryValue("PDFLowEncryption")
60640   If IsNumeric(tStr) Then
60650     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60660       .PDFLowEncryption = CLng(tStr)
60670      Else
60680       If UseStandard Then
60690        .PDFLowEncryption = 1
60700       End If
60710     End If
60720    Else
60730     If UseStandard Then
60740      .PDFLowEncryption = 1
60750     End If
60760   End If
60770   tStr = reg.GetRegistryValue("PDFOwnerPass")
60780   If IsNumeric(tStr) Then
60790     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60800       .PDFOwnerPass = CLng(tStr)
60810      Else
60820       If UseStandard Then
60830        .PDFOwnerPass = 0
60840       End If
60850     End If
60860    Else
60870     If UseStandard Then
60880      .PDFOwnerPass = 0
60890     End If
60900   End If
60910   tStr = reg.GetRegistryValue("PDFOwnerPasswordString")
60920   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
60930     .PDFOwnerPasswordString = ""
60940    Else
60950     If LenB(tStr) > 0 Then
60960      .PDFOwnerPasswordString = tStr
60970     End If
60980   End If
60990   tStr = reg.GetRegistryValue("PDFUserPass")
61000   If IsNumeric(tStr) Then
61010     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61020       .PDFUserPass = CLng(tStr)
61030      Else
61040       If UseStandard Then
61050        .PDFUserPass = 0
61060       End If
61070     End If
61080    Else
61090     If UseStandard Then
61100      .PDFUserPass = 0
61110     End If
61120   End If
61130   tStr = reg.GetRegistryValue("PDFUserPasswordString")
61140   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
61150     .PDFUserPasswordString = ""
61160    Else
61170     If LenB(tStr) > 0 Then
61180      .PDFUserPasswordString = tStr
61190     End If
61200   End If
61210   tStr = reg.GetRegistryValue("PDFUseSecurity")
61220   If IsNumeric(tStr) Then
61230     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61240       .PDFUseSecurity = CLng(tStr)
61250      Else
61260       If UseStandard Then
61270        .PDFUseSecurity = 0
61280       End If
61290     End If
61300    Else
61310     If UseStandard Then
61320      .PDFUseSecurity = 0
61330     End If
61340   End If
61350   reg.Subkey = "Printing\Formats\PS\LanguageLevel"
61360   tStr = reg.GetRegistryValue("EPSLanguageLevel")
61370   If IsNumeric(tStr) Then
61380     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
61390       .EPSLanguageLevel = CLng(tStr)
61400      Else
61410       If UseStandard Then
61420        .EPSLanguageLevel = 2
61430       End If
61440     End If
61450    Else
61460     If UseStandard Then
61470      .EPSLanguageLevel = 2
61480     End If
61490   End If
61500   tStr = reg.GetRegistryValue("PSLanguageLevel")
61510   If IsNumeric(tStr) Then
61520     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
61530       .PSLanguageLevel = CLng(tStr)
61540      Else
61550       If UseStandard Then
61560        .PSLanguageLevel = 2
61570       End If
61580     End If
61590    Else
61600     If UseStandard Then
61610      .PSLanguageLevel = 2
61620     End If
61630   End If
61640   reg.Subkey = "Program"
61650   tStr = reg.GetRegistryValue("AdditionalGhostscriptParameters")
61660   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
61670     .AdditionalGhostscriptParameters = ""
61680    Else
61690     If LenB(tStr) > 0 Then
61700      .AdditionalGhostscriptParameters = tStr
61710     End If
61720   End If
61730   tStr = reg.GetRegistryValue("AdditionalGhostscriptSearchpath")
61740   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
61750     .AdditionalGhostscriptSearchpath = ""
61760    Else
61770     If LenB(tStr) > 0 Then
61780      .AdditionalGhostscriptSearchpath = tStr
61790     End If
61800   End If
61810   tStr = reg.GetRegistryValue("AddWindowsFontpath")
61820   If IsNumeric(tStr) Then
61830     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61840       .AddWindowsFontpath = CLng(tStr)
61850      Else
61860       If UseStandard Then
61870        .AddWindowsFontpath = 1
61880       End If
61890     End If
61900    Else
61910     If UseStandard Then
61920      .AddWindowsFontpath = 1
61930     End If
61940   End If
61950   tStr = reg.GetRegistryValue("AutosaveDirectory")
61960   If LenB(Trim$(tStr)) > 0 Then
61970     .AutosaveDirectory = CompletePath(tStr)
61980    Else
61990     If UseStandard Then
62000      If InstalledAsServer Then
62010        .AutosaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
62020       Else
62030        .AutosaveDirectory = "<MyFiles>"
62040      End If
62050     End If
62060   End If
62070   tStr = reg.GetRegistryValue("AutosaveFilename")
62080   If LenB(tStr) = 0 And LenB("<DateTime>") > 0 And UseStandard Then
62090     .AutosaveFilename = "<DateTime>"
62100    Else
62110     If LenB(tStr) > 0 Then
62120      .AutosaveFilename = tStr
62130     End If
62140   End If
62150   tStr = reg.GetRegistryValue("AutosaveFormat")
62160   If IsNumeric(tStr) Then
62170     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
62180       .AutosaveFormat = CLng(tStr)
62190      Else
62200       If UseStandard Then
62210        .AutosaveFormat = 0
62220       End If
62230     End If
62240    Else
62250     If UseStandard Then
62260      .AutosaveFormat = 0
62270     End If
62280   End If
62290   tStr = reg.GetRegistryValue("ClientComputerResolveIPAddress")
62300   If IsNumeric(tStr) Then
62310     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62320       .ClientComputerResolveIPAddress = CLng(tStr)
62330      Else
62340       If UseStandard Then
62350        .ClientComputerResolveIPAddress = 0
62360       End If
62370     End If
62380    Else
62390     If UseStandard Then
62400      .ClientComputerResolveIPAddress = 0
62410     End If
62420   End If
62430   tStr = reg.GetRegistryValue("DisableEmail")
62440   If IsNumeric(tStr) Then
62450     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62460       .DisableEmail = CLng(tStr)
62470      Else
62480       If UseStandard Then
62490        .DisableEmail = 0
62500       End If
62510     End If
62520    Else
62530     If UseStandard Then
62540      .DisableEmail = 0
62550     End If
62560   End If
62570   tStr = reg.GetRegistryValue("DontUseDocumentSettings")
62580   If IsNumeric(tStr) Then
62590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62600       .DontUseDocumentSettings = CLng(tStr)
62610      Else
62620       If UseStandard Then
62630        .DontUseDocumentSettings = 0
62640       End If
62650     End If
62660    Else
62670     If UseStandard Then
62680      .DontUseDocumentSettings = 0
62690     End If
62700   End If
62710   tStr = reg.GetRegistryValue("FilenameSubstitutions")
62720   If LenB(tStr) = 0 And LenB("Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt") > 0 And UseStandard Then
62730     .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
62740    Else
62750     If LenB(tStr) > 0 Then
62760      .FilenameSubstitutions = tStr
62770     End If
62780   End If
62790   tStr = reg.GetRegistryValue("FilenameSubstitutionsOnlyInTitle")
62800   If IsNumeric(tStr) Then
62810     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62820       .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
62830      Else
62840       If UseStandard Then
62850        .FilenameSubstitutionsOnlyInTitle = 1
62860       End If
62870     End If
62880    Else
62890     If UseStandard Then
62900      .FilenameSubstitutionsOnlyInTitle = 1
62910     End If
62920   End If
62930   tStr = reg.GetRegistryValue("Language")
62940   If LenB(tStr) = 0 And LenB("english") > 0 And UseStandard Then
62950     .Language = "english"
62960    Else
62970     If LenB(tStr) > 0 Then
62980      .Language = tStr
62990     End If
63000   End If
63010   tStr = reg.GetRegistryValue("LastSaveDirectory")
63020   If LenB(Trim$(tStr)) > 0 Then
63030     .LastSaveDirectory = CompletePath(tStr)
63040    Else
63050     If UseStandard Then
63060      If InstalledAsServer Then
63070        .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
63080       Else
63090        .LastSaveDirectory = "<MyFiles>"
63100      End If
63110     End If
63120   End If
63130   tStr = reg.GetRegistryValue("Logging")
63140   If IsNumeric(tStr) Then
63150     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63160       .Logging = CLng(tStr)
63170      Else
63180       If UseStandard Then
63190        .Logging = 0
63200       End If
63210     End If
63220    Else
63230     If UseStandard Then
63240      .Logging = 0
63250     End If
63260   End If
63270   tStr = reg.GetRegistryValue("LogLines")
63280   If IsNumeric(tStr) Then
63290     If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
63300       .LogLines = CLng(tStr)
63310      Else
63320       If UseStandard Then
63330        .LogLines = 100
63340       End If
63350     End If
63360    Else
63370     If UseStandard Then
63380      .LogLines = 100
63390     End If
63400   End If
63410   tStr = reg.GetRegistryValue("NoConfirmMessageSwitchingDefaultprinter")
63420   If IsNumeric(tStr) Then
63430     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63440       .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
63450      Else
63460       If UseStandard Then
63470        .NoConfirmMessageSwitchingDefaultprinter = 0
63480       End If
63490     End If
63500    Else
63510     If UseStandard Then
63520      .NoConfirmMessageSwitchingDefaultprinter = 0
63530     End If
63540   End If
63550   tStr = reg.GetRegistryValue("NoProcessingAtStartup")
63560   If IsNumeric(tStr) Then
63570     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63580       .NoProcessingAtStartup = CLng(tStr)
63590      Else
63600       If UseStandard Then
63610        .NoProcessingAtStartup = 0
63620       End If
63630     End If
63640    Else
63650     If UseStandard Then
63660      .NoProcessingAtStartup = 0
63670     End If
63680   End If
63690   tStr = reg.GetRegistryValue("NoPSCheck")
63700   If IsNumeric(tStr) Then
63710     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63720       .NoPSCheck = CLng(tStr)
63730      Else
63740       If UseStandard Then
63750        .NoPSCheck = 0
63760       End If
63770     End If
63780    Else
63790     If UseStandard Then
63800      .NoPSCheck = 0
63810     End If
63820   End If
63830   tStr = reg.GetRegistryValue("OptionsDesign")
63840   If IsNumeric(tStr) Then
63850     If CLng(tStr) >= 1 And CLng(tStr) <= 2 Then
63860       .OptionsDesign = CLng(tStr)
63870      Else
63880       If UseStandard Then
63890        .OptionsDesign = 1
63900       End If
63910     End If
63920    Else
63930     If UseStandard Then
63940      .OptionsDesign = 1
63950     End If
63960   End If
63970   tStr = reg.GetRegistryValue("OptionsEnabled")
63980   If IsNumeric(tStr) Then
63990     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64000       .OptionsEnabled = CLng(tStr)
64010      Else
64020       If UseStandard Then
64030        .OptionsEnabled = 1
64040       End If
64050     End If
64060    Else
64070     If UseStandard Then
64080      .OptionsEnabled = 1
64090     End If
64100   End If
64110   tStr = reg.GetRegistryValue("OptionsVisible")
64120   If IsNumeric(tStr) Then
64130     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64140       .OptionsVisible = CLng(tStr)
64150      Else
64160       If UseStandard Then
64170        .OptionsVisible = 1
64180       End If
64190     End If
64200    Else
64210     If UseStandard Then
64220      .OptionsVisible = 1
64230     End If
64240   End If
64250   tStr = reg.GetRegistryValue("PrintAfterSaving")
64260   If IsNumeric(tStr) Then
64270     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64280       .PrintAfterSaving = CLng(tStr)
64290      Else
64300       If UseStandard Then
64310        .PrintAfterSaving = 0
64320       End If
64330     End If
64340    Else
64350     If UseStandard Then
64360      .PrintAfterSaving = 0
64370     End If
64380   End If
64390   tStr = reg.GetRegistryValue("PrintAfterSavingDuplex")
64400   If IsNumeric(tStr) Then
64410     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64420       .PrintAfterSavingDuplex = CLng(tStr)
64430      Else
64440       If UseStandard Then
64450        .PrintAfterSavingDuplex = 0
64460       End If
64470     End If
64480    Else
64490     If UseStandard Then
64500      .PrintAfterSavingDuplex = 0
64510     End If
64520   End If
64530   tStr = reg.GetRegistryValue("PrintAfterSavingNoCancel")
64540   If IsNumeric(tStr) Then
64550     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64560       .PrintAfterSavingNoCancel = CLng(tStr)
64570      Else
64580       If UseStandard Then
64590        .PrintAfterSavingNoCancel = 0
64600       End If
64610     End If
64620    Else
64630     If UseStandard Then
64640      .PrintAfterSavingNoCancel = 0
64650     End If
64660   End If
64670   tStr = reg.GetRegistryValue("PrintAfterSavingPrinter")
64680   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64690     .PrintAfterSavingPrinter = ""
64700    Else
64710     If LenB(tStr) > 0 Then
64720      .PrintAfterSavingPrinter = tStr
64730     End If
64740   End If
64750   tStr = reg.GetRegistryValue("PrintAfterSavingQueryUser")
64760   If IsNumeric(tStr) Then
64770     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
64780       .PrintAfterSavingQueryUser = CLng(tStr)
64790      Else
64800       If UseStandard Then
64810        .PrintAfterSavingQueryUser = 0
64820       End If
64830     End If
64840    Else
64850     If UseStandard Then
64860      .PrintAfterSavingQueryUser = 0
64870     End If
64880   End If
64890   tStr = reg.GetRegistryValue("PrintAfterSavingTumble")
64900   If IsNumeric(tStr) Then
64910     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64920       .PrintAfterSavingTumble = CLng(tStr)
64930      Else
64940       If UseStandard Then
64950        .PrintAfterSavingTumble = 0
64960       End If
64970     End If
64980    Else
64990     If UseStandard Then
65000      .PrintAfterSavingTumble = 0
65010     End If
65020   End If
65030   tStr = reg.GetRegistryValue("PrinterStop")
65040   If IsNumeric(tStr) Then
65050     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65060       .PrinterStop = CLng(tStr)
65070      Else
65080       If UseStandard Then
65090        .PrinterStop = 0
65100       End If
65110     End If
65120    Else
65130     If UseStandard Then
65140      .PrinterStop = 0
65150     End If
65160   End If
65170   tStr = reg.GetRegistryValue("PrinterTemppath")
65180   WriteToSpecialLogfile "reg.GetRegistryValue(""PrinterTemppath"")=" & tStr
65190   WriteToSpecialLogfile "Options.PrinterTemppath1=" & .PrinterTemppath
65200   If hkey1 = HKEY_USERS Then
65210     If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then
65220       .PrinterTemppath = tStr
65230      Else
65240       If UseStandard Then
65250         .PrinterTemppath = GetTempPath
65260        Else
65270         .PrinterTemppath = tStr
65280       End If
65290     End If
65300    Else
65310     If LenB(Trim$(tStr)) > 0 Then
65320      If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then
65330        .PrinterTemppath = tStr
65340       Else
65350        MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))
65360        If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then
65370          If UseStandard Then
65380            .PrinterTemppath = GetTempPath
65390           Else
65400            .PrinterTemppath = ""
65410            If NoMsg = False Then
65420             MsgBox "PrinterTemppath: '" & tStr & "' = '" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & "'" & _
             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
65440            End If
65450          End If
65460         Else
65470          .PrinterTemppath = tStr
65480        End If
65490      End If
65500     End If
65510   End If
65520   WriteToSpecialLogfile "Options.PrinterTemppath2=" & .PrinterTemppath
65530   tStr = reg.GetRegistryValue("ProcessPriority")
65540   If IsNumeric(tStr) Then
65550     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
65560       .ProcessPriority = CLng(tStr)
65570      Else
65580       If UseStandard Then
65590        .ProcessPriority = 1
65600       End If
65610     End If
65620    Else
65630     If UseStandard Then
65640      .ProcessPriority = 1
65650     End If
65660   End If
65670   tStr = reg.GetRegistryValue("ProgramFont")
65680   If LenB(tStr) = 0 And LenB("MS Sans Serif") > 0 And UseStandard Then
65690     .ProgramFont = "MS Sans Serif"
65700    Else
65710     If LenB(tStr) > 0 Then
65720      .ProgramFont = tStr
65730     End If
65740   End If
65750   tStr = reg.GetRegistryValue("ProgramFontCharset")
65760   If IsNumeric(tStr) Then
65770     If CLng(tStr) >= 0 Then
65780       .ProgramFontCharset = CLng(tStr)
65790      Else
65800       If UseStandard Then
65810        .ProgramFontCharset = 0
65820       End If
65830     End If
65840    Else
65850     If UseStandard Then
65860      .ProgramFontCharset = 0
65870     End If
65880   End If
65890   tStr = reg.GetRegistryValue("ProgramFontSize")
65900   If IsNumeric(tStr) Then
65910     If CLng(tStr) >= 1 And CLng(tStr) <= 72 Then
65920       .ProgramFontSize = CLng(tStr)
65930      Else
65940       If UseStandard Then
65950        .ProgramFontSize = 8
65960       End If
65970     End If
65980    Else
65990     If UseStandard Then
66000      .ProgramFontSize = 8
66010     End If
66020   End If
66030   tStr = reg.GetRegistryValue("RemoveAllKnownFileExtensions")
66040   If IsNumeric(tStr) Then
66050     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66060       .RemoveAllKnownFileExtensions = CLng(tStr)
66070      Else
66080       If UseStandard Then
66090        .RemoveAllKnownFileExtensions = 1
66100       End If
66110     End If
66120    Else
66130     If UseStandard Then
66140      .RemoveAllKnownFileExtensions = 1
66150     End If
66160   End If
66170   tStr = reg.GetRegistryValue("RemoveSpaces")
66180   If IsNumeric(tStr) Then
66190     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66200       .RemoveSpaces = CLng(tStr)
66210      Else
66220       If UseStandard Then
66230        .RemoveSpaces = 1
66240       End If
66250     End If
66260    Else
66270     If UseStandard Then
66280      .RemoveSpaces = 1
66290     End If
66300   End If
66310   tStr = reg.GetRegistryValue("RunProgramAfterSaving")
66320   If IsNumeric(tStr) Then
66330     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66340       .RunProgramAfterSaving = CLng(tStr)
66350      Else
66360       If UseStandard Then
66370        .RunProgramAfterSaving = 0
66380       End If
66390     End If
66400    Else
66410     If UseStandard Then
66420      .RunProgramAfterSaving = 0
66430     End If
66440   End If
66450   tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramname")
66460   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
66470     .RunProgramAfterSavingProgramname = ""
66480    Else
66490     If LenB(tStr) > 0 Then
66500      .RunProgramAfterSavingProgramname = tStr
66510     End If
66520   End If
66530   tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramParameters")
66540   If LenB(tStr) = 0 And LenB("""<OutputFilename>""") > 0 And UseStandard Then
66550     .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
66560    Else
66570     If LenB(tStr) > 0 Then
66580      .RunProgramAfterSavingProgramParameters = tStr
66590     End If
66600   End If
66610   tStr = reg.GetRegistryValue("RunProgramAfterSavingWaitUntilReady")
66620   If IsNumeric(tStr) Then
66630     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66640       .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
66650      Else
66660       If UseStandard Then
66670        .RunProgramAfterSavingWaitUntilReady = 1
66680       End If
66690     End If
66700    Else
66710     If UseStandard Then
66720      .RunProgramAfterSavingWaitUntilReady = 1
66730     End If
66740   End If
66750   tStr = reg.GetRegistryValue("RunProgramAfterSavingWindowstyle")
66760   If IsNumeric(tStr) Then
66770     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
66780       .RunProgramAfterSavingWindowstyle = CLng(tStr)
66790      Else
66800       If UseStandard Then
66810        .RunProgramAfterSavingWindowstyle = 1
66820       End If
66830     End If
66840    Else
66850     If UseStandard Then
66860      .RunProgramAfterSavingWindowstyle = 1
66870     End If
66880   End If
66890   tStr = reg.GetRegistryValue("RunProgramBeforeSaving")
66900   If IsNumeric(tStr) Then
66910     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66920       .RunProgramBeforeSaving = CLng(tStr)
66930      Else
66940       If UseStandard Then
66950        .RunProgramBeforeSaving = 0
66960       End If
66970     End If
66980    Else
66990     If UseStandard Then
67000      .RunProgramBeforeSaving = 0
67010     End If
67020   End If
67030   tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramname")
67040   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67050     .RunProgramBeforeSavingProgramname = ""
67060    Else
67070     If LenB(tStr) > 0 Then
67080      .RunProgramBeforeSavingProgramname = tStr
67090     End If
67100   End If
67110   tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramParameters")
67120   If LenB(tStr) = 0 And LenB("""<TempFilename>""") > 0 And UseStandard Then
67130     .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
67140    Else
67150     If LenB(tStr) > 0 Then
67160      .RunProgramBeforeSavingProgramParameters = tStr
67170     End If
67180   End If
67190   tStr = reg.GetRegistryValue("RunProgramBeforeSavingWindowstyle")
67200   If IsNumeric(tStr) Then
67210     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
67220       .RunProgramBeforeSavingWindowstyle = CLng(tStr)
67230      Else
67240       If UseStandard Then
67250        .RunProgramBeforeSavingWindowstyle = 1
67260       End If
67270     End If
67280    Else
67290     If UseStandard Then
67300      .RunProgramBeforeSavingWindowstyle = 1
67310     End If
67320   End If
67330   tStr = reg.GetRegistryValue("SaveFilename")
67340   If LenB(tStr) = 0 And LenB("<Title>") > 0 And UseStandard Then
67350     .SaveFilename = "<Title>"
67360    Else
67370     If LenB(tStr) > 0 Then
67380      .SaveFilename = tStr
67390     End If
67400   End If
67410   tStr = reg.GetRegistryValue("SendEmailAfterAutoSaving")
67420   If IsNumeric(tStr) Then
67430     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67440       .SendEmailAfterAutoSaving = CLng(tStr)
67450      Else
67460       If UseStandard Then
67470        .SendEmailAfterAutoSaving = 0
67480       End If
67490     End If
67500    Else
67510     If UseStandard Then
67520      .SendEmailAfterAutoSaving = 0
67530     End If
67540   End If
67550   tStr = reg.GetRegistryValue("SendMailMethod")
67560   If IsNumeric(tStr) Then
67570     If CLng(tStr) >= 0 Then
67580       .SendMailMethod = CLng(tStr)
67590      Else
67600       If UseStandard Then
67610        .SendMailMethod = 0
67620       End If
67630     End If
67640    Else
67650     If UseStandard Then
67660      .SendMailMethod = 0
67670     End If
67680   End If
67690   tStr = reg.GetRegistryValue("ShowAnimation")
67700   If IsNumeric(tStr) Then
67710     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67720       .ShowAnimation = CLng(tStr)
67730      Else
67740       If UseStandard Then
67750        .ShowAnimation = 1
67760       End If
67770     End If
67780    Else
67790     If UseStandard Then
67800      .ShowAnimation = 1
67810     End If
67820   End If
67830   tStr = reg.GetRegistryValue("StartStandardProgram")
67840   If IsNumeric(tStr) Then
67850     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67860       .StartStandardProgram = CLng(tStr)
67870      Else
67880       If UseStandard Then
67890        .StartStandardProgram = 1
67900       End If
67910     End If
67920    Else
67930     If UseStandard Then
67940      .StartStandardProgram = 1
67950     End If
67960   End If
67970   tStr = reg.GetRegistryValue("Toolbars")
67980   If IsNumeric(tStr) Then
67990     If CLng(tStr) >= 0 Then
68000       .Toolbars = CLng(tStr)
68010      Else
68020       If UseStandard Then
68030        .Toolbars = 1
68040       End If
68050     End If
68060    Else
68070     If UseStandard Then
68080      .Toolbars = 1
68090     End If
68100   End If
68110   tStr = reg.GetRegistryValue("UseAutosave")
68120   If IsNumeric(tStr) Then
68130     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68140       .UseAutosave = CLng(tStr)
68150      Else
68160       If UseStandard Then
68170        .UseAutosave = 0
68180       End If
68190     End If
68200    Else
68210     If UseStandard Then
68220      .UseAutosave = 0
68230     End If
68240   End If
68250   tStr = reg.GetRegistryValue("UseAutosaveDirectory")
68260   If IsNumeric(tStr) Then
68270     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68280       .UseAutosaveDirectory = CLng(tStr)
68290      Else
68300       If UseStandard Then
68310        .UseAutosaveDirectory = 1
68320       End If
68330     End If
68340    Else
68350     If UseStandard Then
68360      .UseAutosaveDirectory = 1
68370     End If
68380   End If
68390  End With
68400  Set reg = Nothing
68410  ReadOptionsReg = myOptions
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
57840   If UCase$(OptionName) = "CLIENTCOMPUTERRESOLVEIPADDRESS" Then
57850    If Not reg.KeyExists Then
57860     reg.CreateKey
57870    End If
57880    reg.SetRegistryValue "ClientComputerResolveIPAddress", CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
57890    Set reg = Nothing
57900    Exit Sub
57910   End If
57920   If UCase$(OptionName) = "DISABLEEMAIL" Then
57930    If Not reg.KeyExists Then
57940     reg.CreateKey
57950    End If
57960    reg.SetRegistryValue "DisableEmail", CStr(Abs(.DisableEmail)), REG_SZ
57970    Set reg = Nothing
57980    Exit Sub
57990   End If
58000   If UCase$(OptionName) = "DONTUSEDOCUMENTSETTINGS" Then
58010    If Not reg.KeyExists Then
58020     reg.CreateKey
58030    End If
58040    reg.SetRegistryValue "DontUseDocumentSettings", CStr(Abs(.DontUseDocumentSettings)), REG_SZ
58050    Set reg = Nothing
58060    Exit Sub
58070   End If
58080   If UCase$(OptionName) = "FILENAMESUBSTITUTIONS" Then
58090    If Not reg.KeyExists Then
58100     reg.CreateKey
58110    End If
58120    reg.SetRegistryValue "FilenameSubstitutions", CStr(.FilenameSubstitutions), REG_SZ
58130    Set reg = Nothing
58140    Exit Sub
58150   End If
58160   If UCase$(OptionName) = "FILENAMESUBSTITUTIONSONLYINTITLE" Then
58170    If Not reg.KeyExists Then
58180     reg.CreateKey
58190    End If
58200    reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle", CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
58210    Set reg = Nothing
58220    Exit Sub
58230   End If
58240   If UCase$(OptionName) = "LANGUAGE" Then
58250    If Not reg.KeyExists Then
58260     reg.CreateKey
58270    End If
58280    reg.SetRegistryValue "Language", CStr(.Language), REG_SZ
58290    Set reg = Nothing
58300    Exit Sub
58310   End If
58320   If UCase$(OptionName) = "LASTSAVEDIRECTORY" Then
58330    If Not reg.KeyExists Then
58340     reg.CreateKey
58350    End If
58360    reg.SetRegistryValue "LastSaveDirectory", CStr(.LastSaveDirectory), REG_SZ
58370    Set reg = Nothing
58380    Exit Sub
58390   End If
58400   If UCase$(OptionName) = "LOGGING" Then
58410    If Not reg.KeyExists Then
58420     reg.CreateKey
58430    End If
58440    reg.SetRegistryValue "Logging", CStr(Abs(.Logging)), REG_SZ
58450    Set reg = Nothing
58460    Exit Sub
58470   End If
58480   If UCase$(OptionName) = "LOGLINES" Then
58490    If Not reg.KeyExists Then
58500     reg.CreateKey
58510    End If
58520    reg.SetRegistryValue "LogLines", CStr(.LogLines), REG_SZ
58530    Set reg = Nothing
58540    Exit Sub
58550   End If
58560   If UCase$(OptionName) = "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER" Then
58570    If Not reg.KeyExists Then
58580     reg.CreateKey
58590    End If
58600    reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter", CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
58610    Set reg = Nothing
58620    Exit Sub
58630   End If
58640   If UCase$(OptionName) = "NOPROCESSINGATSTARTUP" Then
58650    If Not reg.KeyExists Then
58660     reg.CreateKey
58670    End If
58680    reg.SetRegistryValue "NoProcessingAtStartup", CStr(Abs(.NoProcessingAtStartup)), REG_SZ
58690    Set reg = Nothing
58700    Exit Sub
58710   End If
58720   If UCase$(OptionName) = "NOPSCHECK" Then
58730    If Not reg.KeyExists Then
58740     reg.CreateKey
58750    End If
58760    reg.SetRegistryValue "NoPSCheck", CStr(Abs(.NoPSCheck)), REG_SZ
58770    Set reg = Nothing
58780    Exit Sub
58790   End If
58800   If UCase$(OptionName) = "OPTIONSDESIGN" Then
58810    If Not reg.KeyExists Then
58820     reg.CreateKey
58830    End If
58840    reg.SetRegistryValue "OptionsDesign", CStr(.OptionsDesign), REG_SZ
58850    Set reg = Nothing
58860    Exit Sub
58870   End If
58880   If UCase$(OptionName) = "OPTIONSENABLED" Then
58890    If Not reg.KeyExists Then
58900     reg.CreateKey
58910    End If
58920    reg.SetRegistryValue "OptionsEnabled", CStr(Abs(.OptionsEnabled)), REG_SZ
58930    Set reg = Nothing
58940    Exit Sub
58950   End If
58960   If UCase$(OptionName) = "OPTIONSVISIBLE" Then
58970    If Not reg.KeyExists Then
58980     reg.CreateKey
58990    End If
59000    reg.SetRegistryValue "OptionsVisible", CStr(Abs(.OptionsVisible)), REG_SZ
59010    Set reg = Nothing
59020    Exit Sub
59030   End If
59040   If UCase$(OptionName) = "PRINTAFTERSAVING" Then
59050    If Not reg.KeyExists Then
59060     reg.CreateKey
59070    End If
59080    reg.SetRegistryValue "PrintAfterSaving", CStr(Abs(.PrintAfterSaving)), REG_SZ
59090    Set reg = Nothing
59100    Exit Sub
59110   End If
59120   If UCase$(OptionName) = "PRINTAFTERSAVINGDUPLEX" Then
59130    If Not reg.KeyExists Then
59140     reg.CreateKey
59150    End If
59160    reg.SetRegistryValue "PrintAfterSavingDuplex", CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
59170    Set reg = Nothing
59180    Exit Sub
59190   End If
59200   If UCase$(OptionName) = "PRINTAFTERSAVINGNOCANCEL" Then
59210    If Not reg.KeyExists Then
59220     reg.CreateKey
59230    End If
59240    reg.SetRegistryValue "PrintAfterSavingNoCancel", CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
59250    Set reg = Nothing
59260    Exit Sub
59270   End If
59280   If UCase$(OptionName) = "PRINTAFTERSAVINGPRINTER" Then
59290    If Not reg.KeyExists Then
59300     reg.CreateKey
59310    End If
59320    reg.SetRegistryValue "PrintAfterSavingPrinter", CStr(.PrintAfterSavingPrinter), REG_SZ
59330    Set reg = Nothing
59340    Exit Sub
59350   End If
59360   If UCase$(OptionName) = "PRINTAFTERSAVINGQUERYUSER" Then
59370    If Not reg.KeyExists Then
59380     reg.CreateKey
59390    End If
59400    reg.SetRegistryValue "PrintAfterSavingQueryUser", CStr(.PrintAfterSavingQueryUser), REG_SZ
59410    Set reg = Nothing
59420    Exit Sub
59430   End If
59440   If UCase$(OptionName) = "PRINTAFTERSAVINGTUMBLE" Then
59450    If Not reg.KeyExists Then
59460     reg.CreateKey
59470    End If
59480    reg.SetRegistryValue "PrintAfterSavingTumble", CStr(Abs(.PrintAfterSavingTumble)), REG_SZ
59490    Set reg = Nothing
59500    Exit Sub
59510   End If
59520   If UCase$(OptionName) = "PRINTERSTOP" Then
59530    If Not reg.KeyExists Then
59540     reg.CreateKey
59550    End If
59560    reg.SetRegistryValue "PrinterStop", CStr(Abs(.PrinterStop)), REG_SZ
59570    Set reg = Nothing
59580    Exit Sub
59590   End If
59600   If UCase$(OptionName) = "PRINTERTEMPPATH" Then
59610    If Not reg.KeyExists Then
59620     reg.CreateKey
59630    End If
59640    reg.SetRegistryValue "PrinterTemppath", CStr(.PrinterTemppath), REG_SZ
59650    Set reg = Nothing
59660    Exit Sub
59670   End If
59680   If UCase$(OptionName) = "PROCESSPRIORITY" Then
59690    If Not reg.KeyExists Then
59700     reg.CreateKey
59710    End If
59720    reg.SetRegistryValue "ProcessPriority", CStr(.ProcessPriority), REG_SZ
59730    Set reg = Nothing
59740    Exit Sub
59750   End If
59760   If UCase$(OptionName) = "PROGRAMFONT" Then
59770    If Not reg.KeyExists Then
59780     reg.CreateKey
59790    End If
59800    reg.SetRegistryValue "ProgramFont", CStr(.ProgramFont), REG_SZ
59810    Set reg = Nothing
59820    Exit Sub
59830   End If
59840   If UCase$(OptionName) = "PROGRAMFONTCHARSET" Then
59850    If Not reg.KeyExists Then
59860     reg.CreateKey
59870    End If
59880    reg.SetRegistryValue "ProgramFontCharset", CStr(.ProgramFontCharset), REG_SZ
59890    Set reg = Nothing
59900    Exit Sub
59910   End If
59920   If UCase$(OptionName) = "PROGRAMFONTSIZE" Then
59930    If Not reg.KeyExists Then
59940     reg.CreateKey
59950    End If
59960    reg.SetRegistryValue "ProgramFontSize", CStr(.ProgramFontSize), REG_SZ
59970    Set reg = Nothing
59980    Exit Sub
59990   End If
60000   If UCase$(OptionName) = "REMOVEALLKNOWNFILEEXTENSIONS" Then
60010    If Not reg.KeyExists Then
60020     reg.CreateKey
60030    End If
60040    reg.SetRegistryValue "RemoveAllKnownFileExtensions", CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
60050    Set reg = Nothing
60060    Exit Sub
60070   End If
60080   If UCase$(OptionName) = "REMOVESPACES" Then
60090    If Not reg.KeyExists Then
60100     reg.CreateKey
60110    End If
60120    reg.SetRegistryValue "RemoveSpaces", CStr(Abs(.RemoveSpaces)), REG_SZ
60130    Set reg = Nothing
60140    Exit Sub
60150   End If
60160   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVING" Then
60170    If Not reg.KeyExists Then
60180     reg.CreateKey
60190    End If
60200    reg.SetRegistryValue "RunProgramAfterSaving", CStr(Abs(.RunProgramAfterSaving)), REG_SZ
60210    Set reg = Nothing
60220    Exit Sub
60230   End If
60240   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGPROGRAMNAME" Then
60250    If Not reg.KeyExists Then
60260     reg.CreateKey
60270    End If
60280    reg.SetRegistryValue "RunProgramAfterSavingProgramname", CStr(.RunProgramAfterSavingProgramname), REG_SZ
60290    Set reg = Nothing
60300    Exit Sub
60310   End If
60320   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS" Then
60330    If Not reg.KeyExists Then
60340     reg.CreateKey
60350    End If
60360    reg.SetRegistryValue "RunProgramAfterSavingProgramParameters", CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
60370    Set reg = Nothing
60380    Exit Sub
60390   End If
60400   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGWAITUNTILREADY" Then
60410    If Not reg.KeyExists Then
60420     reg.CreateKey
60430    End If
60440    reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady", CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
60450    Set reg = Nothing
60460    Exit Sub
60470   End If
60480   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGWINDOWSTYLE" Then
60490    If Not reg.KeyExists Then
60500     reg.CreateKey
60510    End If
60520    reg.SetRegistryValue "RunProgramAfterSavingWindowstyle", CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
60530    Set reg = Nothing
60540    Exit Sub
60550   End If
60560   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVING" Then
60570    If Not reg.KeyExists Then
60580     reg.CreateKey
60590    End If
60600    reg.SetRegistryValue "RunProgramBeforeSaving", CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
60610    Set reg = Nothing
60620    Exit Sub
60630   End If
60640   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGPROGRAMNAME" Then
60650    If Not reg.KeyExists Then
60660     reg.CreateKey
60670    End If
60680    reg.SetRegistryValue "RunProgramBeforeSavingProgramname", CStr(.RunProgramBeforeSavingProgramname), REG_SZ
60690    Set reg = Nothing
60700    Exit Sub
60710   End If
60720   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS" Then
60730    If Not reg.KeyExists Then
60740     reg.CreateKey
60750    End If
60760    reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters", CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
60770    Set reg = Nothing
60780    Exit Sub
60790   End If
60800   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGWINDOWSTYLE" Then
60810    If Not reg.KeyExists Then
60820     reg.CreateKey
60830    End If
60840    reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle", CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
60850    Set reg = Nothing
60860    Exit Sub
60870   End If
60880   If UCase$(OptionName) = "SAVEFILENAME" Then
60890    If Not reg.KeyExists Then
60900     reg.CreateKey
60910    End If
60920    reg.SetRegistryValue "SaveFilename", CStr(.SaveFilename), REG_SZ
60930    Set reg = Nothing
60940    Exit Sub
60950   End If
60960   If UCase$(OptionName) = "SENDEMAILAFTERAUTOSAVING" Then
60970    If Not reg.KeyExists Then
60980     reg.CreateKey
60990    End If
61000    reg.SetRegistryValue "SendEmailAfterAutoSaving", CStr(Abs(.SendEmailAfterAutoSaving)), REG_SZ
61010    Set reg = Nothing
61020    Exit Sub
61030   End If
61040   If UCase$(OptionName) = "SENDMAILMETHOD" Then
61050    If Not reg.KeyExists Then
61060     reg.CreateKey
61070    End If
61080    reg.SetRegistryValue "SendMailMethod", CStr(.SendMailMethod), REG_SZ
61090    Set reg = Nothing
61100    Exit Sub
61110   End If
61120   If UCase$(OptionName) = "SHOWANIMATION" Then
61130    If Not reg.KeyExists Then
61140     reg.CreateKey
61150    End If
61160    reg.SetRegistryValue "ShowAnimation", CStr(Abs(.ShowAnimation)), REG_SZ
61170    Set reg = Nothing
61180    Exit Sub
61190   End If
61200   If UCase$(OptionName) = "STARTSTANDARDPROGRAM" Then
61210    If Not reg.KeyExists Then
61220     reg.CreateKey
61230    End If
61240    reg.SetRegistryValue "StartStandardProgram", CStr(Abs(.StartStandardProgram)), REG_SZ
61250    Set reg = Nothing
61260    Exit Sub
61270   End If
61280   If UCase$(OptionName) = "TOOLBARS" Then
61290    If Not reg.KeyExists Then
61300     reg.CreateKey
61310    End If
61320    reg.SetRegistryValue "Toolbars", CStr(.Toolbars), REG_SZ
61330    Set reg = Nothing
61340    Exit Sub
61350   End If
61360   If UCase$(OptionName) = "USEAUTOSAVE" Then
61370    If Not reg.KeyExists Then
61380     reg.CreateKey
61390    End If
61400    reg.SetRegistryValue "UseAutosave", CStr(Abs(.UseAutosave)), REG_SZ
61410    Set reg = Nothing
61420    Exit Sub
61430   End If
61440   If UCase$(OptionName) = "USEAUTOSAVEDIRECTORY" Then
61450    If Not reg.KeyExists Then
61460     reg.CreateKey
61470    End If
61480    reg.SetRegistryValue "UseAutosaveDirectory", CStr(Abs(.UseAutosaveDirectory)), REG_SZ
61490    Set reg = Nothing
61500    Exit Sub
61510   End If
61520  End With
61530  Set reg = Nothing
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
51450   reg.SetRegistryValue "ClientComputerResolveIPAddress", CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
51460   reg.SetRegistryValue "DisableEmail", CStr(Abs(.DisableEmail)), REG_SZ
51470   reg.SetRegistryValue "DontUseDocumentSettings", CStr(Abs(.DontUseDocumentSettings)), REG_SZ
51480   reg.SetRegistryValue "FilenameSubstitutions", CStr(.FilenameSubstitutions), REG_SZ
51490   reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle", CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
51500   reg.SetRegistryValue "Language", CStr(.Language), REG_SZ
51510   reg.SetRegistryValue "LastSaveDirectory", CStr(.LastSaveDirectory), REG_SZ
51520   reg.SetRegistryValue "Logging", CStr(Abs(.Logging)), REG_SZ
51530   reg.SetRegistryValue "LogLines", CStr(.LogLines), REG_SZ
51540   reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter", CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
51550   reg.SetRegistryValue "NoProcessingAtStartup", CStr(Abs(.NoProcessingAtStartup)), REG_SZ
51560   reg.SetRegistryValue "NoPSCheck", CStr(Abs(.NoPSCheck)), REG_SZ
51570   reg.SetRegistryValue "OptionsDesign", CStr(.OptionsDesign), REG_SZ
51580   reg.SetRegistryValue "OptionsEnabled", CStr(Abs(.OptionsEnabled)), REG_SZ
51590   reg.SetRegistryValue "OptionsVisible", CStr(Abs(.OptionsVisible)), REG_SZ
51600   reg.SetRegistryValue "PrintAfterSaving", CStr(Abs(.PrintAfterSaving)), REG_SZ
51610   reg.SetRegistryValue "PrintAfterSavingDuplex", CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
51620   reg.SetRegistryValue "PrintAfterSavingNoCancel", CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
51630   reg.SetRegistryValue "PrintAfterSavingPrinter", CStr(.PrintAfterSavingPrinter), REG_SZ
51640   reg.SetRegistryValue "PrintAfterSavingQueryUser", CStr(.PrintAfterSavingQueryUser), REG_SZ
51650   reg.SetRegistryValue "PrintAfterSavingTumble", CStr(Abs(.PrintAfterSavingTumble)), REG_SZ
51660   reg.SetRegistryValue "PrinterStop", CStr(Abs(.PrinterStop)), REG_SZ
51670   reg.SetRegistryValue "PrinterTemppath", CStr(.PrinterTemppath), REG_SZ
51680   reg.SetRegistryValue "ProcessPriority", CStr(.ProcessPriority), REG_SZ
51690   reg.SetRegistryValue "ProgramFont", CStr(.ProgramFont), REG_SZ
51700   reg.SetRegistryValue "ProgramFontCharset", CStr(.ProgramFontCharset), REG_SZ
51710   reg.SetRegistryValue "ProgramFontSize", CStr(.ProgramFontSize), REG_SZ
51720   reg.SetRegistryValue "RemoveAllKnownFileExtensions", CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
51730   reg.SetRegistryValue "RemoveSpaces", CStr(Abs(.RemoveSpaces)), REG_SZ
51740   reg.SetRegistryValue "RunProgramAfterSaving", CStr(Abs(.RunProgramAfterSaving)), REG_SZ
51750   reg.SetRegistryValue "RunProgramAfterSavingProgramname", CStr(.RunProgramAfterSavingProgramname), REG_SZ
51760   reg.SetRegistryValue "RunProgramAfterSavingProgramParameters", CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
51770   reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady", CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
51780   reg.SetRegistryValue "RunProgramAfterSavingWindowstyle", CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
51790   reg.SetRegistryValue "RunProgramBeforeSaving", CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
51800   reg.SetRegistryValue "RunProgramBeforeSavingProgramname", CStr(.RunProgramBeforeSavingProgramname), REG_SZ
51810   reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters", CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
51820   reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle", CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
51830   reg.SetRegistryValue "SaveFilename", CStr(.SaveFilename), REG_SZ
51840   reg.SetRegistryValue "SendEmailAfterAutoSaving", CStr(Abs(.SendEmailAfterAutoSaving)), REG_SZ
51850   reg.SetRegistryValue "SendMailMethod", CStr(.SendMailMethod), REG_SZ
51860   reg.SetRegistryValue "ShowAnimation", CStr(Abs(.ShowAnimation)), REG_SZ
51870   reg.SetRegistryValue "StartStandardProgram", CStr(Abs(.StartStandardProgram)), REG_SZ
51880   reg.SetRegistryValue "Toolbars", CStr(.Toolbars), REG_SZ
51890   reg.SetRegistryValue "UseAutosave", CStr(Abs(.UseAutosave)), REG_SZ
51900   reg.SetRegistryValue "UseAutosaveDirectory", CStr(Abs(.UseAutosaveDirectory)), REG_SZ
51910  End With
51920  Set reg = Nothing
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
  Frm.txtAutosaveDirectory.Text = .AutosaveDirectory
  Frm.txtAutosaveFilename.Text = .AutosaveFilename
  Frm.cmbAutosaveFormat.ListIndex = .AutosaveFormat
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
  Frm.chkOwnerPass.Value = .PDFOwnerPass
  Frm.chkUserPass.Value = .PDFUserPass
  Frm.chkUseSecurity.Value = .PDFUseSecurity
  Frm.cmbPNGColors.ListIndex = .PNGColorscount
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
  Frm.txtSaveFilename.Text = .SaveFilename
  Frm.txtStandardAuthor.Text = .StandardAuthor
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
50030  .AutosaveDirectory = Frm.txtAutosaveDirectory.Text
50040  .AutosaveFilename = Frm.txtAutosaveFilename.Text
50050  .AutosaveFormat = Frm.cmbAutosaveFormat.ListIndex
50060  .BitmapResolution = Frm.txtBitmapResolution.Text
50070  .BMPColorscount = Frm.cmbBMPColors.ListIndex
50080  .DirectoryGhostscriptBinaries = Frm.txtGSbin.Text
50090  .DirectoryGhostscriptFonts = Frm.txtGSfonts.Text
50100  .DirectoryGhostscriptLibraries = Frm.txtGSlib.Text
50110  .DirectoryGhostscriptResource = Frm.txtGSResource.Text
50120  .EPSLanguageLevel = Frm.cmbEPSLanguageLevel.ListIndex
50130  tStr = ""
50140  Set lsv = Frm.lsvFilenameSubst
50150  For i = 1 To lsv.ListItems.Count
50160   If i < lsv.ListItems.Count Then
50170     tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1) & "\"
50180    Else
50190     tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1)
50200   End If
50210  Next i
50220  .FilenameSubstitutions = tStr
50230  .FilenameSubstitutionsOnlyInTitle = Abs(Frm.chkFilenameSubst.Value)
50240  .JPEGColorscount = Frm.cmbJPEGColors.ListIndex
50250  .JPEGQuality = Frm.txtJPEGQuality.Text
50260  .NoConfirmMessageSwitchingDefaultprinter = Abs(Frm.chkNoConfirmMessageSwitchingDefaultprinter)
50270  .PCXColorscount = Frm.cmbPCXColors.ListIndex
50280  .PDFAllowAssembly = Abs(Frm.chkAllowAssembly.Value)
50290  .PDFAllowDegradedPrinting = Abs(Frm.chkAllowDegradedPrinting.Value)
50300  .PDFAllowFillIn = Abs(Frm.chkAllowFillIn.Value)
50310  .PDFAllowScreenReaders = Abs(Frm.chkAllowScreenReaders.Value)
50320  .PDFColorsCMYKToRGB = Abs(Frm.chkPDFCMYKtoRGB.Value)
50330  .PDFColorsColorModel = Frm.cmbPDFColorModel.ListIndex
50340  .PDFColorsPreserveHalftone = Abs(Frm.chkPDFPreserveHalftone.Value)
50350  .PDFColorsPreserveOverprint = Abs(Frm.chkPDFPreserveOverprint.Value)
50360  .PDFColorsPreserveTransfer = Abs(Frm.chkPDFPreserveTransfer.Value)
50370  .PDFCompressionColorCompression = Abs(Frm.chkPDFColorComp.Value)
50380  .PDFCompressionColorCompressionChoice = Frm.cmbPDFColorComp.ListIndex
50390  .PDFCompressionColorResample = Abs(Frm.chkPDFColorResample.Value)
50400  .PDFCompressionColorResampleChoice = Frm.cmbPDFColorResample.ListIndex
50410  .PDFCompressionColorResolution = Frm.txtPDFColorRes.Text
50420  .PDFCompressionGreyCompression = Abs(Frm.chkPDFGreyComp.Value)
50430  .PDFCompressionGreyCompressionChoice = Frm.cmbPDFGreyComp.ListIndex
50440  .PDFCompressionGreyResample = Abs(Frm.chkPDFGreyResample.Value)
50450  .PDFCompressionGreyResampleChoice = Frm.cmbPDFGreyResample.ListIndex
50460  .PDFCompressionGreyResolution = Frm.txtPDFGreyRes.Text
50470  .PDFCompressionMonoCompression = Abs(Frm.chkPDFMonoComp.Value)
50480  .PDFCompressionMonoCompressionChoice = Frm.cmbPDFMonoComp.ListIndex
50490  .PDFCompressionMonoResample = Abs(Frm.chkPDFMonoResample.Value)
50500  .PDFCompressionMonoResampleChoice = Frm.cmbPDFMonoResample.ListIndex
50510  .PDFCompressionMonoResolution = Frm.txtPDFMonoRes.Text
50520  .PDFCompressionTextCompression = Abs(Frm.chkPDFTextComp.Value)
50530  .PDFDisallowCopy = Abs(Frm.chkAllowCopy.Value)
50540  .PDFDisallowModifyAnnotations = Abs(Frm.chkAllowModifyAnnotations.Value)
50550  .PDFDisallowModifyContents = Abs(Frm.chkAllowModifyContents.Value)
50560  .PDFDisallowPrinting = Abs(Frm.chkAllowPrinting.Value)
50570  If Frm.cmbPDFEncryptor.ListIndex < 0 Then
50580    .PDFEncryptor = 0
50590   Else
50600    .PDFEncryptor = Frm.cmbPDFEncryptor.ItemData(Frm.cmbPDFEncryptor.ListIndex)
50610  End If
50620  .PDFFontsEmbedAll = Abs(Frm.chkPDFEmbedAll.Value)
50630  .PDFFontsSubSetFonts = Abs(Frm.chkPDFSubSetFonts.Value)
50640  .PDFFontsSubSetFontsPercent = Frm.txtPDFSubSetPerc.Text
50650  .PDFGeneralASCII85 = Abs(Frm.chkPDFASCII85.Value)
50660  .PDFGeneralAutorotate = Frm.cmbPDFRotate.ListIndex
50670  .PDFGeneralCompatibility = Frm.cmbPDFCompat.ListIndex
50680  .PDFGeneralOverprint = Frm.cmbPDFOverprint.ListIndex
50690  .PDFGeneralResolution = Frm.txtPDFRes.Text
50700  .PDFHighEncryption = Abs(Frm.optEncHigh.Value)
50710  .PDFLowEncryption = Abs(Frm.optEncLow.Value)
50720  .PDFOwnerPass = Abs(Frm.chkOwnerPass.Value)
50730  .PDFUserPass = Abs(Frm.chkUserPass.Value)
50740  .PDFUseSecurity = Abs(Frm.chkUseSecurity.Value)
50750  .PNGColorscount = Frm.cmbPNGColors.ListIndex
50760  .PrinterTemppath = Frm.txtTemppath.Text
50770  .ProcessPriority = Frm.sldProcessPriority.Value
50780  .ProgramFont = Frm.cmbFonts.List(Frm.cmbFonts.ListIndex)
50790  .ProgramFontCharset = Frm.cmbCharset.Text
50800  .ProgramFontSize = Frm.cmbProgramFontSize.Text
50810  .PSLanguageLevel = Frm.cmbPSLanguageLevel.ListIndex
50820  .RemoveSpaces = Abs(Frm.chkSpaces.Value)
50830  .SaveFilename = Frm.txtSaveFilename.Text
50840  .StandardAuthor = Frm.txtStandardAuthor.Text
50850  .TIFFColorscount = Frm.cmbTIFFColors.ListIndex
50860  .UseAutosave = Abs(Frm.chkUseAutosave.Value)
50870  .UseAutosaveDirectory = Abs(Frm.chkUseAutosaveDirectory.Value)
50880  .UseCreationDateNow = Abs(Frm.chkUseCreationDateNow.Value)
50890  .UseStandardAuthor = Abs(Frm.chkUseStandardAuthor.Value)
50900  End With
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

