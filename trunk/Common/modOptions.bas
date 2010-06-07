Attribute VB_Name = "modOptions"
Option Explicit

' Automatically generated with DeveloperTool by Frank Heindörfer
' 2003 - 2007
' Email: thesmilyface@users.sourceforge.net

Public Type tOptions
 AdditionalGhostscriptParameters As String
 AdditionalGhostscriptSearchpath As String
 AddWindowsFontpath As Long
 AllowSpecialGSCharsInFilenames As Long
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
 PDFCompressionColorCompressionJPEGManualFactor As Double
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
 PDFCompressionGreyCompressionJPEGManualFactor As Double
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
 PDFSigningSignatureOnPage As Long
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
 PrintAfterSavingBitsPerPixel As Long
 PrintAfterSavingDuplex As Long
 PrintAfterSavingMaxResolution As Long
 PrintAfterSavingMaxResolutionEnabled As Long
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
 SVGResolution As Long
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

Public Options As tOptions, Options1 As tOptions

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
50070   .AllowSpecialGSCharsInFilenames = "1"
50080   If InstalledAsServer Then
50090     .AutosaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
50100    Else
50110     .AutosaveDirectory = "<MyFiles>"
50120   End If
50130   .AutosaveFilename = "<DateTime>"
50140   .AutosaveFormat = "0"
50150   .AutosaveStartStandardProgram = "0"
50160   .BMPColorscount = "1"
50170   .BMPResolution = "150"
50180   .ClientComputerResolveIPAddress = "0"
50190   .Counter = "0"
50200   .DeviceHeightPoints = Replace$("842", ".", GetDecimalChar)
50210   .DeviceWidthPoints = Replace$("595", ".", GetDecimalChar)
50220   .DirectoryGhostscriptBinaries = vbNullString
50230   .DirectoryGhostscriptFonts = vbNullString
50240   .DirectoryGhostscriptLibraries = vbNullString
50250   .DirectoryGhostscriptResource = vbNullString
50260   .DisableEmail = "0"
50270   .DontUseDocumentSettings = "0"
50280   .EPSLanguageLevel = "2"
50290   .FilenameSubstitutions = "Microsoft Word - \.docx\.doc\Microsoft Excel - \.xlsx\.xls\Microsoft PowerPoint - \.pptx\.ppt"
50300   .FilenameSubstitutionsOnlyInTitle = "1"
50310   .JPEGColorscount = "0"
50320   .JPEGQuality = "75"
50330   .JPEGResolution = "150"
50340   .Language = "english"
50350   If InstalledAsServer Then
50360     .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
50370    Else
50380     .LastSaveDirectory = "<MyFiles>"
50390   End If
50400   .LastUpdateCheck = vbNullString
50410   .Logging = "0"
50420   .LogLines = "100"
50430   .NoConfirmMessageSwitchingDefaultprinter = "0"
50440   .NoProcessingAtStartup = "0"
50450   .NoPSCheck = "0"
50460   .OnePagePerFile = "0"
50470   .OptionsDesign = "0"
50480   .OptionsEnabled = "1"
50490   .OptionsVisible = "1"
50500   .Papersize = "a4"
50510   .PCLColorsCount = "0"
50520   .PCLResolution = "150"
50530   .PCXColorscount = "0"
50540   .PCXResolution = "150"
50550   .PDFAllowAssembly = "0"
50560   .PDFAllowDegradedPrinting = "0"
50570   .PDFAllowFillIn = "0"
50580   .PDFAllowScreenReaders = "0"
50590   .PDFColorsCMYKToRGB = "0"
50600   .PDFColorsColorModel = "1"
50610   .PDFColorsPreserveHalftone = "0"
50620   .PDFColorsPreserveOverprint = "1"
50630   .PDFColorsPreserveTransfer = "1"
50640   .PDFCompressionColorCompression = "1"
50650   .PDFCompressionColorCompressionChoice = "0"
50660   .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
50670   .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
50680   .PDFCompressionColorCompressionJPEGManualFactor = Replace$("3.0", ".", GetDecimalChar)
50690   .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
50700   .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
50710   .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
50720   .PDFCompressionColorResample = "0"
50730   .PDFCompressionColorResampleChoice = "0"
50740   .PDFCompressionColorResolution = "300"
50750   .PDFCompressionGreyCompression = "1"
50760   .PDFCompressionGreyCompressionChoice = "0"
50770   .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
50780   .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
50790   .PDFCompressionGreyCompressionJPEGManualFactor = Replace$("3.0", ".", GetDecimalChar)
50800   .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
50810   .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
50820   .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
50830   .PDFCompressionGreyResample = "0"
50840   .PDFCompressionGreyResampleChoice = "0"
50850   .PDFCompressionGreyResolution = "300"
50860   .PDFCompressionMonoCompression = "1"
50870   .PDFCompressionMonoCompressionChoice = "0"
50880   .PDFCompressionMonoResample = "0"
50890   .PDFCompressionMonoResampleChoice = "0"
50900   .PDFCompressionMonoResolution = "1200"
50910   .PDFCompressionTextCompression = "1"
50920   .PDFDisallowCopy = "1"
50930   .PDFDisallowModifyAnnotations = "0"
50940   .PDFDisallowModifyContents = "0"
50950   .PDFDisallowPrinting = "0"
50960   .PDFEncryptor = "0"
50970   .PDFFontsEmbedAll = "1"
50980   .PDFFontsSubSetFonts = "1"
50990   .PDFFontsSubSetFontsPercent = "100"
51000   .PDFGeneralASCII85 = "0"
51010   .PDFGeneralAutorotate = "2"
51020   .PDFGeneralCompatibility = "2"
51030   .PDFGeneralDefault = "0"
51040   .PDFGeneralOverprint = "0"
51050   .PDFGeneralResolution = "600"
51060   .PDFHighEncryption = "0"
51070   .PDFLowEncryption = "1"
51080   .PDFOptimize = "0"
51090   .PDFOwnerPass = "0"
51100   .PDFOwnerPasswordString = vbNullString
51110   .PDFSigningMultiSignature = "0"
51120   .PDFSigningPFXFile = vbNullString
51130   .PDFSigningPFXFilePassword = vbNullString
51140   .PDFSigningSignatureContact = vbNullString
51150   .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
51160   .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
51170   .PDFSigningSignatureLocation = vbNullString
51180   .PDFSigningSignatureOnPage = "1"
51190   .PDFSigningSignatureReason = vbNullString
51200   .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
51210   .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
51220   .PDFSigningSignatureVisible = "0"
51230   .PDFSigningSignPDF = "0"
51240   .PDFUpdateMetadata = "1"
51250   .PDFUserPass = "0"
51260   .PDFUserPasswordString = vbNullString
51270   .PDFUseSecurity = "0"
51280   .PNGColorscount = "0"
51290   .PNGResolution = "150"
51300   .PrintAfterSaving = "0"
51310   .PrintAfterSavingBitsPerPixel = "2"
51320   .PrintAfterSavingDuplex = "0"
51330   .PrintAfterSavingMaxResolution = "600"
51340   .PrintAfterSavingMaxResolutionEnabled = "0"
51350   .PrintAfterSavingNoCancel = "0"
51360   .PrintAfterSavingPrinter = vbNullString
51370   .PrintAfterSavingQueryUser = "0"
51380   .PrintAfterSavingTumble = "0"
51390   .PrinterStop = "0"
51400   If InstalledAsServer Then
51410     .PrinterTemppath = CompletePath(GetPDFCreatorApplicationPath) & "Temp\"
51420    Else
51430     .PrinterTemppath = "<Temp>PDFCreator\"
51440   End If
51450   .ProcessPriority = "1"
51460   .ProgramFont = "MS Sans Serif"
51470   .ProgramFontCharset = "0"
51480   .ProgramFontSize = "8"
51490   .PSDColorsCount = "0"
51500   .PSDResolution = "150"
51510   .PSLanguageLevel = "2"
51520   .RAWColorsCount = "0"
51530   .RAWResolution = "150"
51540   .RemoveAllKnownFileExtensions = "1"
51550   .RemoveSpaces = "1"
51560   .RunProgramAfterSaving = "0"
51570   .RunProgramAfterSavingProgramname = vbNullString
51580   .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
51590   .RunProgramAfterSavingWaitUntilReady = "1"
51600   .RunProgramAfterSavingWindowstyle = "1"
51610   .RunProgramBeforeSaving = "0"
51620   .RunProgramBeforeSavingProgramname = vbNullString
51630   .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
51640   .RunProgramBeforeSavingWindowstyle = "1"
51650   .SaveFilename = "<Title>"
51660   .SendEmailAfterAutoSaving = "0"
51670   .SendMailMethod = "0"
51680   .ShowAnimation = "1"
51690   .StampFontColor = "#FF0000"
51700   .StampFontname = "Arial"
51710   .StampFontsize = "48"
51720   .StampOutlineFontthickness = "0"
51730   .StampString = vbNullString
51740   .StampUseOutlineFont = "1"
51750   .StandardAuthor = vbNullString
51760   .StandardCreationdate = vbNullString
51770   .StandardDateformat = "YYYYMMDDHHNNSS"
51780   .StandardKeywords = vbNullString
51790   .StandardMailDomain = vbNullString
51800   .StandardModifydate = vbNullString
51810   .StandardSaveformat = "0"
51820   .StandardSubject = vbNullString
51830   .StandardTitle = vbNullString
51840   .StartStandardProgram = "1"
51850   .SVGResolution = "72"
51860   .TIFFColorscount = "0"
51870   .TIFFResolution = "150"
51880   .Toolbars = "1"
51890   .UpdateInterval = "2"
51900   .UseAutosave = "0"
51910   .UseAutosaveDirectory = "1"
51920   .UseCreationDateNow = "0"
51930   .UseCustomPaperSize = "0"
51940   .UseFixPapersize = "0"
51950   .UseStandardAuthor = "0"
51960  End With
51970  If Not IsWin9xMe Then
51980   myOptions = ReadOptionsReg(myOptions, ".DEFAULT\Software\PDFCreator", HKEY_USERS, False, False)
51990  End If
52000  StandardOptions = myOptions
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

Public Function ReadOptions(Optional NoMsg As Boolean = False, Optional hProfile As hkey = HKEY_CURRENT_USER, Optional ProfileName As String = "") As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim myOptions As tOptions, Str1 As String
50020  Dim tStr As String
50030
50040  ProfileName = Trim$(ProfileName)
50050  If LenB(ProfileName) > 0 Then
50060   tStr = "_" & ProfileName
50070  End If
50080  If InstalledAsServer Then
50090    WriteToSpecialLogfile "Reg-Read options: HKEY_LOCAL_MACHINE"
50100    If LenB(ProfileName) > 0 Then
50110      myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator\Profiles\" & ProfileName, HKEY_LOCAL_MACHINE, NoMsg)
50120     Else
50130      myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", HKEY_LOCAL_MACHINE, NoMsg)
50140    End If
50150   Else
50160    If Not IsWin9xMe Then
50170      WriteToSpecialLogfile "Reg-Read options: HKEY_USERS"
50180      myOptions = ReadOptionsReg(myOptions, ".DEFAULT\Software\PDFCreator", HKEY_USERS, NoMsg)
50190      WriteToSpecialLogfile "Reg-Read options: HKEY_CURRENT_USER [" & hProfile & "]"
50200      If LenB(ProfileName) > 0 Then
50210        myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator\Profiles\" & ProfileName, hProfile, NoMsg, False)
50220       Else
50230        myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", hProfile, NoMsg, False)
50240      End If
50250     Else
50260      WriteToSpecialLogfile "Reg-Read options: HKEY_CURRENT_USER [" & hProfile & "]"
50270      If LenB(ProfileName) > 0 Then
50280        myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator\Profiles\" & ProfileName, hProfile, NoMsg)
50290       Else
50300        myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", hProfile, NoMsg)
50310      End If
50320    End If
50330    WriteToSpecialLogfile "Reg-Read options: HKEY_LOCAL_MACHINE"
50340    If LenB(ProfileName) > 0 Then
50350      myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator\Profiles\" & ProfileName, HKEY_LOCAL_MACHINE, NoMsg, False)
50360     Else
50370      myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", HKEY_LOCAL_MACHINE, NoMsg, False)
50380    End If
50390  End If
50400  Str1 = "7777772E706466666F7267652E6F7267"
50410  myOptions = CorrectOptionsAfterLoading(myOptions)
50420  ReadOptions = myOptions
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
50440   tStr = hOpt.Retrieve("AllowSpecialGSCharsInFilenames")
50450   If IsNumeric(tStr) Then
50460     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
50470       .AllowSpecialGSCharsInFilenames = CLng(tStr)
50480      Else
50490       If UseStandard Then
50500        .AllowSpecialGSCharsInFilenames = 1
50510       End If
50520     End If
50530    Else
50540     If UseStandard Then
50550      .AllowSpecialGSCharsInFilenames = 1
50560     End If
50570   End If
50580   tStr = hOpt.Retrieve("AutosaveDirectory")
50590   If LenB(Trim$(tStr)) > 0 Then
50600     .AutosaveDirectory = CompletePath(tStr)
50610    Else
50620     If UseStandard Then
50630      If InstalledAsServer Then
50640        .AutosaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
50650       Else
50660        .AutosaveDirectory = "<MyFiles>"
50670      End If
50680     End If
50690   End If
50700   tStr = hOpt.Retrieve("AutosaveFilename")
50710   If LenB(tStr) = 0 And LenB("<DateTime>") > 0 And UseStandard Then
50720     .AutosaveFilename = "<DateTime>"
50730    Else
50740     If LenB(tStr) > 0 Then
50750      .AutosaveFilename = tStr
50760     End If
50770   End If
50780   tStr = hOpt.Retrieve("AutosaveFormat")
50790   If IsNumeric(tStr) Then
50800     If CLng(tStr) >= 0 And CLng(tStr) <= 14 Then
50810       .AutosaveFormat = CLng(tStr)
50820      Else
50830       If UseStandard Then
50840        .AutosaveFormat = 0
50850       End If
50860     End If
50870    Else
50880     If UseStandard Then
50890      .AutosaveFormat = 0
50900     End If
50910   End If
50920   tStr = hOpt.Retrieve("AutosaveStartStandardProgram")
50930   If IsNumeric(tStr) Then
50940     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
50950       .AutosaveStartStandardProgram = CLng(tStr)
50960      Else
50970       If UseStandard Then
50980        .AutosaveStartStandardProgram = 0
50990       End If
51000     End If
51010    Else
51020     If UseStandard Then
51030      .AutosaveStartStandardProgram = 0
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
51200   tStr = hOpt.Retrieve("BMPResolution")
51210   If IsNumeric(tStr) Then
51220     If CLng(tStr) >= 1 Then
51230       .BMPResolution = CLng(tStr)
51240      Else
51250       If UseStandard Then
51260        .BMPResolution = 150
51270       End If
51280     End If
51290    Else
51300     If UseStandard Then
51310      .BMPResolution = 150
51320     End If
51330   End If
51340   tStr = hOpt.Retrieve("ClientComputerResolveIPAddress")
51350   If IsNumeric(tStr) Then
51360     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51370       .ClientComputerResolveIPAddress = CLng(tStr)
51380      Else
51390       If UseStandard Then
51400        .ClientComputerResolveIPAddress = 0
51410       End If
51420     End If
51430    Else
51440     If UseStandard Then
51450      .ClientComputerResolveIPAddress = 0
51460     End If
51470   End If
51480   tStr = hOpt.Retrieve("Counter")
51490   If IsNumeric(tStr) Then
51500     If CCur(tStr) >= 0 And CCur(tStr) <= 922337203685477# Then
51510       .Counter = CCur(tStr)
51520      Else
51530       If UseStandard Then
51540        .Counter = 0
51550       End If
51560     End If
51570    Else
51580     If UseStandard Then
51590      .Counter = 0
51600     End If
51610   End If
51620   tStr = hOpt.Retrieve("DeviceHeightPoints")
51630   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
51640     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 1 Then
51650       .DeviceHeightPoints = CDbl(Replace$(tStr, ".", GetDecimalChar))
51660      Else
51670       If UseStandard Then
51680        .DeviceHeightPoints = Replace$("842", ".", GetDecimalChar)
51690       End If
51700     End If
51710    Else
51720     If UseStandard Then
51730      .DeviceHeightPoints = Replace$("842", ".", GetDecimalChar)
51740     End If
51750   End If
51760   tStr = hOpt.Retrieve("DeviceWidthPoints")
51770   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
51780     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 1 Then
51790       .DeviceWidthPoints = CDbl(Replace$(tStr, ".", GetDecimalChar))
51800      Else
51810       If UseStandard Then
51820        .DeviceWidthPoints = Replace$("595", ".", GetDecimalChar)
51830       End If
51840     End If
51850    Else
51860     If UseStandard Then
51870      .DeviceWidthPoints = Replace$("595", ".", GetDecimalChar)
51880     End If
51890   End If
51900   tStr = hOpt.Retrieve("DirectoryGhostscriptBinaries")
51910   If LenB(Trim$(tStr)) > 0 Then
51920     .DirectoryGhostscriptBinaries = CompletePath(tStr)
51930    Else
51940     If UseStandard Then
51950      tStr = GetPDFCreatorApplicationPath
51960      .DirectoryGhostscriptBinaries = CompletePath(tStr)
51970     End If
51980   End If
51990   tStr = hOpt.Retrieve("DirectoryGhostscriptFonts")
52000   If LenB(Trim$(tStr)) > 0 Then
52010     .DirectoryGhostscriptFonts = CompletePath(tStr)
52020    Else
52030     If UseStandard Then
52040      tStr = GetPDFCreatorApplicationPath & "fonts"
52050      .DirectoryGhostscriptFonts = CompletePath(tStr)
52060     End If
52070   End If
52080   tStr = hOpt.Retrieve("DirectoryGhostscriptLibraries")
52090   If LenB(Trim$(tStr)) > 0 Then
52100     .DirectoryGhostscriptLibraries = CompletePath(tStr)
52110    Else
52120     If UseStandard Then
52130      tStr = GetPDFCreatorApplicationPath & "lib"
52140      .DirectoryGhostscriptLibraries = CompletePath(tStr)
52150     End If
52160   End If
52170   tStr = hOpt.Retrieve("DirectoryGhostscriptResource")
52180   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
52190     .DirectoryGhostscriptResource = ""
52200    Else
52210     If LenB(tStr) > 0 Then
52220      .DirectoryGhostscriptResource = tStr
52230     End If
52240   End If
52250   tStr = hOpt.Retrieve("DisableEmail")
52260   If IsNumeric(tStr) Then
52270     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52280       .DisableEmail = CLng(tStr)
52290      Else
52300       If UseStandard Then
52310        .DisableEmail = 0
52320       End If
52330     End If
52340    Else
52350     If UseStandard Then
52360      .DisableEmail = 0
52370     End If
52380   End If
52390   tStr = hOpt.Retrieve("DontUseDocumentSettings")
52400   If IsNumeric(tStr) Then
52410     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52420       .DontUseDocumentSettings = CLng(tStr)
52430      Else
52440       If UseStandard Then
52450        .DontUseDocumentSettings = 0
52460       End If
52470     End If
52480    Else
52490     If UseStandard Then
52500      .DontUseDocumentSettings = 0
52510     End If
52520   End If
52530   tStr = hOpt.Retrieve("EPSLanguageLevel")
52540   If IsNumeric(tStr) Then
52550     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
52560       .EPSLanguageLevel = CLng(tStr)
52570      Else
52580       If UseStandard Then
52590        .EPSLanguageLevel = 2
52600       End If
52610     End If
52620    Else
52630     If UseStandard Then
52640      .EPSLanguageLevel = 2
52650     End If
52660   End If
52670   tStr = hOpt.Retrieve("FilenameSubstitutions")
52680   If LenB(tStr) = 0 And LenB("Microsoft Word - \.docx\.doc\Microsoft Excel - \.xlsx\.xls\Microsoft PowerPoint - \.pptx\.ppt") > 0 And UseStandard Then
52690     .FilenameSubstitutions = "Microsoft Word - \.docx\.doc\Microsoft Excel - \.xlsx\.xls\Microsoft PowerPoint - \.pptx\.ppt"
52700    Else
52710     If LenB(tStr) > 0 Then
52720      .FilenameSubstitutions = tStr
52730     End If
52740   End If
52750   tStr = hOpt.Retrieve("FilenameSubstitutionsOnlyInTitle")
52760   If IsNumeric(tStr) Then
52770     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52780       .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
52790      Else
52800       If UseStandard Then
52810        .FilenameSubstitutionsOnlyInTitle = 1
52820       End If
52830     End If
52840    Else
52850     If UseStandard Then
52860      .FilenameSubstitutionsOnlyInTitle = 1
52870     End If
52880   End If
52890   tStr = hOpt.Retrieve("JPEGColorscount")
52900   If IsNumeric(tStr) Then
52910     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
52920       .JPEGColorscount = CLng(tStr)
52930      Else
52940       If UseStandard Then
52950        .JPEGColorscount = 0
52960       End If
52970     End If
52980    Else
52990     If UseStandard Then
53000      .JPEGColorscount = 0
53010     End If
53020   End If
53030   tStr = hOpt.Retrieve("JPEGQuality")
53040   If IsNumeric(tStr) Then
53050     If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
53060       .JPEGQuality = CLng(tStr)
53070      Else
53080       If UseStandard Then
53090        .JPEGQuality = 75
53100       End If
53110     End If
53120    Else
53130     If UseStandard Then
53140      .JPEGQuality = 75
53150     End If
53160   End If
53170   tStr = hOpt.Retrieve("JPEGResolution")
53180   If IsNumeric(tStr) Then
53190     If CLng(tStr) >= 1 Then
53200       .JPEGResolution = CLng(tStr)
53210      Else
53220       If UseStandard Then
53230        .JPEGResolution = 150
53240       End If
53250     End If
53260    Else
53270     If UseStandard Then
53280      .JPEGResolution = 150
53290     End If
53300   End If
53310   tStr = hOpt.Retrieve("Language")
53320   If LenB(tStr) = 0 And LenB("english") > 0 And UseStandard Then
53330     .Language = "english"
53340    Else
53350     If LenB(tStr) > 0 Then
53360      .Language = tStr
53370     End If
53380   End If
53390   tStr = hOpt.Retrieve("LastSaveDirectory")
53400   If LenB(Trim$(tStr)) > 0 Then
53410     .LastSaveDirectory = CompletePath(tStr)
53420    Else
53430     If UseStandard Then
53440      If InstalledAsServer Then
53450        .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
53460       Else
53470        .LastSaveDirectory = "<MyFiles>"
53480      End If
53490     End If
53500   End If
53510   tStr = hOpt.Retrieve("LastUpdateCheck")
53520   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
53530     .LastUpdateCheck = ""
53540    Else
53550     If LenB(tStr) > 0 Then
53560      .LastUpdateCheck = tStr
53570     End If
53580   End If
53590   tStr = hOpt.Retrieve("Logging")
53600   If IsNumeric(tStr) Then
53610     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53620       .Logging = CLng(tStr)
53630      Else
53640       If UseStandard Then
53650        .Logging = 0
53660       End If
53670     End If
53680    Else
53690     If UseStandard Then
53700      .Logging = 0
53710     End If
53720   End If
53730   tStr = hOpt.Retrieve("LogLines")
53740   If IsNumeric(tStr) Then
53750     If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
53760       .LogLines = CLng(tStr)
53770      Else
53780       If UseStandard Then
53790        .LogLines = 100
53800       End If
53810     End If
53820    Else
53830     If UseStandard Then
53840      .LogLines = 100
53850     End If
53860   End If
53870   tStr = hOpt.Retrieve("NoConfirmMessageSwitchingDefaultprinter")
53880   If IsNumeric(tStr) Then
53890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53900       .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
53910      Else
53920       If UseStandard Then
53930        .NoConfirmMessageSwitchingDefaultprinter = 0
53940       End If
53950     End If
53960    Else
53970     If UseStandard Then
53980      .NoConfirmMessageSwitchingDefaultprinter = 0
53990     End If
54000   End If
54010   tStr = hOpt.Retrieve("NoProcessingAtStartup")
54020   If IsNumeric(tStr) Then
54030     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54040       .NoProcessingAtStartup = CLng(tStr)
54050      Else
54060       If UseStandard Then
54070        .NoProcessingAtStartup = 0
54080       End If
54090     End If
54100    Else
54110     If UseStandard Then
54120      .NoProcessingAtStartup = 0
54130     End If
54140   End If
54150   tStr = hOpt.Retrieve("NoPSCheck")
54160   If IsNumeric(tStr) Then
54170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54180       .NoPSCheck = CLng(tStr)
54190      Else
54200       If UseStandard Then
54210        .NoPSCheck = 0
54220       End If
54230     End If
54240    Else
54250     If UseStandard Then
54260      .NoPSCheck = 0
54270     End If
54280   End If
54290   tStr = hOpt.Retrieve("OnePagePerFile")
54300   If IsNumeric(tStr) Then
54310     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54320       .OnePagePerFile = CLng(tStr)
54330      Else
54340       If UseStandard Then
54350        .OnePagePerFile = 0
54360       End If
54370     End If
54380    Else
54390     If UseStandard Then
54400      .OnePagePerFile = 0
54410     End If
54420   End If
54430   tStr = hOpt.Retrieve("OptionsDesign")
54440   If IsNumeric(tStr) Then
54450     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
54460       .OptionsDesign = CLng(tStr)
54470      Else
54480       If UseStandard Then
54490        .OptionsDesign = 0
54500       End If
54510     End If
54520    Else
54530     If UseStandard Then
54540      .OptionsDesign = 0
54550     End If
54560   End If
54570   tStr = hOpt.Retrieve("OptionsEnabled")
54580   If IsNumeric(tStr) Then
54590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54600       .OptionsEnabled = CLng(tStr)
54610      Else
54620       If UseStandard Then
54630        .OptionsEnabled = 1
54640       End If
54650     End If
54660    Else
54670     If UseStandard Then
54680      .OptionsEnabled = 1
54690     End If
54700   End If
54710   tStr = hOpt.Retrieve("OptionsVisible")
54720   If IsNumeric(tStr) Then
54730     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54740       .OptionsVisible = CLng(tStr)
54750      Else
54760       If UseStandard Then
54770        .OptionsVisible = 1
54780       End If
54790     End If
54800    Else
54810     If UseStandard Then
54820      .OptionsVisible = 1
54830     End If
54840   End If
54850   tStr = hOpt.Retrieve("Papersize")
54860   If LenB(tStr) = 0 And LenB("a4") > 0 And UseStandard Then
54870     .Papersize = "a4"
54880    Else
54890     If LenB(tStr) > 0 Then
54900      .Papersize = tStr
54910     End If
54920   End If
54930   tStr = hOpt.Retrieve("PCLColorsCount")
54940   If IsNumeric(tStr) Then
54950     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
54960       .PCLColorsCount = CLng(tStr)
54970      Else
54980       If UseStandard Then
54990        .PCLColorsCount = 0
55000       End If
55010     End If
55020    Else
55030     If UseStandard Then
55040      .PCLColorsCount = 0
55050     End If
55060   End If
55070   tStr = hOpt.Retrieve("PCLResolution")
55080   If IsNumeric(tStr) Then
55090     If CLng(tStr) >= 1 Then
55100       .PCLResolution = CLng(tStr)
55110      Else
55120       If UseStandard Then
55130        .PCLResolution = 150
55140       End If
55150     End If
55160    Else
55170     If UseStandard Then
55180      .PCLResolution = 150
55190     End If
55200   End If
55210   tStr = hOpt.Retrieve("PCXColorscount")
55220   If IsNumeric(tStr) Then
55230     If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
55240       .PCXColorscount = CLng(tStr)
55250      Else
55260       If UseStandard Then
55270        .PCXColorscount = 0
55280       End If
55290     End If
55300    Else
55310     If UseStandard Then
55320      .PCXColorscount = 0
55330     End If
55340   End If
55350   tStr = hOpt.Retrieve("PCXResolution")
55360   If IsNumeric(tStr) Then
55370     If CLng(tStr) >= 1 Then
55380       .PCXResolution = CLng(tStr)
55390      Else
55400       If UseStandard Then
55410        .PCXResolution = 150
55420       End If
55430     End If
55440    Else
55450     If UseStandard Then
55460      .PCXResolution = 150
55470     End If
55480   End If
55490   tStr = hOpt.Retrieve("PDFAllowAssembly")
55500   If IsNumeric(tStr) Then
55510     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55520       .PDFAllowAssembly = CLng(tStr)
55530      Else
55540       If UseStandard Then
55550        .PDFAllowAssembly = 0
55560       End If
55570     End If
55580    Else
55590     If UseStandard Then
55600      .PDFAllowAssembly = 0
55610     End If
55620   End If
55630   tStr = hOpt.Retrieve("PDFAllowDegradedPrinting")
55640   If IsNumeric(tStr) Then
55650     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55660       .PDFAllowDegradedPrinting = CLng(tStr)
55670      Else
55680       If UseStandard Then
55690        .PDFAllowDegradedPrinting = 0
55700       End If
55710     End If
55720    Else
55730     If UseStandard Then
55740      .PDFAllowDegradedPrinting = 0
55750     End If
55760   End If
55770   tStr = hOpt.Retrieve("PDFAllowFillIn")
55780   If IsNumeric(tStr) Then
55790     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55800       .PDFAllowFillIn = CLng(tStr)
55810      Else
55820       If UseStandard Then
55830        .PDFAllowFillIn = 0
55840       End If
55850     End If
55860    Else
55870     If UseStandard Then
55880      .PDFAllowFillIn = 0
55890     End If
55900   End If
55910   tStr = hOpt.Retrieve("PDFAllowScreenReaders")
55920   If IsNumeric(tStr) Then
55930     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55940       .PDFAllowScreenReaders = CLng(tStr)
55950      Else
55960       If UseStandard Then
55970        .PDFAllowScreenReaders = 0
55980       End If
55990     End If
56000    Else
56010     If UseStandard Then
56020      .PDFAllowScreenReaders = 0
56030     End If
56040   End If
56050   tStr = hOpt.Retrieve("PDFColorsCMYKToRGB")
56060   If IsNumeric(tStr) Then
56070     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56080       .PDFColorsCMYKToRGB = CLng(tStr)
56090      Else
56100       If UseStandard Then
56110        .PDFColorsCMYKToRGB = 0
56120       End If
56130     End If
56140    Else
56150     If UseStandard Then
56160      .PDFColorsCMYKToRGB = 0
56170     End If
56180   End If
56190   tStr = hOpt.Retrieve("PDFColorsColorModel")
56200   If IsNumeric(tStr) Then
56210     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
56220       .PDFColorsColorModel = CLng(tStr)
56230      Else
56240       If UseStandard Then
56250        .PDFColorsColorModel = 1
56260       End If
56270     End If
56280    Else
56290     If UseStandard Then
56300      .PDFColorsColorModel = 1
56310     End If
56320   End If
56330   tStr = hOpt.Retrieve("PDFColorsPreserveHalftone")
56340   If IsNumeric(tStr) Then
56350     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56360       .PDFColorsPreserveHalftone = CLng(tStr)
56370      Else
56380       If UseStandard Then
56390        .PDFColorsPreserveHalftone = 0
56400       End If
56410     End If
56420    Else
56430     If UseStandard Then
56440      .PDFColorsPreserveHalftone = 0
56450     End If
56460   End If
56470   tStr = hOpt.Retrieve("PDFColorsPreserveOverprint")
56480   If IsNumeric(tStr) Then
56490     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56500       .PDFColorsPreserveOverprint = CLng(tStr)
56510      Else
56520       If UseStandard Then
56530        .PDFColorsPreserveOverprint = 1
56540       End If
56550     End If
56560    Else
56570     If UseStandard Then
56580      .PDFColorsPreserveOverprint = 1
56590     End If
56600   End If
56610   tStr = hOpt.Retrieve("PDFColorsPreserveTransfer")
56620   If IsNumeric(tStr) Then
56630     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56640       .PDFColorsPreserveTransfer = CLng(tStr)
56650      Else
56660       If UseStandard Then
56670        .PDFColorsPreserveTransfer = 1
56680       End If
56690     End If
56700    Else
56710     If UseStandard Then
56720      .PDFColorsPreserveTransfer = 1
56730     End If
56740   End If
56750   tStr = hOpt.Retrieve("PDFCompressionColorCompression")
56760   If IsNumeric(tStr) Then
56770     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56780       .PDFCompressionColorCompression = CLng(tStr)
56790      Else
56800       If UseStandard Then
56810        .PDFCompressionColorCompression = 1
56820       End If
56830     End If
56840    Else
56850     If UseStandard Then
56860      .PDFCompressionColorCompression = 1
56870     End If
56880   End If
56890   tStr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
56900   If IsNumeric(tStr) Then
56910     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
56920       .PDFCompressionColorCompressionChoice = CLng(tStr)
56930      Else
56940       If UseStandard Then
56950        .PDFCompressionColorCompressionChoice = 0
56960       End If
56970     End If
56980    Else
56990     If UseStandard Then
57000      .PDFCompressionColorCompressionChoice = 0
57010     End If
57020   End If
57030   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGHighFactor")
57040   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57050     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57060       .PDFCompressionColorCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57070      Else
57080       If UseStandard Then
57090        .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
57100       End If
57110     End If
57120    Else
57130     If UseStandard Then
57140      .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
57150     End If
57160   End If
57170   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGLowFactor")
57180   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57190     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57200       .PDFCompressionColorCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57210      Else
57220       If UseStandard Then
57230        .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
57240       End If
57250     End If
57260    Else
57270     If UseStandard Then
57280      .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
57290     End If
57300   End If
57310   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGManualFactor")
57320   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57330     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57340       .PDFCompressionColorCompressionJPEGManualFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57350      Else
57360       If UseStandard Then
57370        .PDFCompressionColorCompressionJPEGManualFactor = Replace$("3.0", ".", GetDecimalChar)
57380       End If
57390     End If
57400    Else
57410     If UseStandard Then
57420      .PDFCompressionColorCompressionJPEGManualFactor = Replace$("3.0", ".", GetDecimalChar)
57430     End If
57440   End If
57450   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMaximumFactor")
57460   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57470     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57480       .PDFCompressionColorCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57490      Else
57500       If UseStandard Then
57510        .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57520       End If
57530     End If
57540    Else
57550     If UseStandard Then
57560      .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57570     End If
57580   End If
57590   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMediumFactor")
57600   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57610     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57620       .PDFCompressionColorCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57630      Else
57640       If UseStandard Then
57650        .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57660       End If
57670     End If
57680    Else
57690     If UseStandard Then
57700      .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57710     End If
57720   End If
57730   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMinimumFactor")
57740   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57750     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57760       .PDFCompressionColorCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57770      Else
57780       If UseStandard Then
57790        .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57800       End If
57810     End If
57820    Else
57830     If UseStandard Then
57840      .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57850     End If
57860   End If
57870   tStr = hOpt.Retrieve("PDFCompressionColorResample")
57880   If IsNumeric(tStr) Then
57890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57900       .PDFCompressionColorResample = CLng(tStr)
57910      Else
57920       If UseStandard Then
57930        .PDFCompressionColorResample = 0
57940       End If
57950     End If
57960    Else
57970     If UseStandard Then
57980      .PDFCompressionColorResample = 0
57990     End If
58000   End If
58010   tStr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
58020   If IsNumeric(tStr) Then
58030     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
58040       .PDFCompressionColorResampleChoice = CLng(tStr)
58050      Else
58060       If UseStandard Then
58070        .PDFCompressionColorResampleChoice = 0
58080       End If
58090     End If
58100    Else
58110     If UseStandard Then
58120      .PDFCompressionColorResampleChoice = 0
58130     End If
58140   End If
58150   tStr = hOpt.Retrieve("PDFCompressionColorResolution")
58160   If IsNumeric(tStr) Then
58170     If CLng(tStr) >= 0 Then
58180       .PDFCompressionColorResolution = CLng(tStr)
58190      Else
58200       If UseStandard Then
58210        .PDFCompressionColorResolution = 300
58220       End If
58230     End If
58240    Else
58250     If UseStandard Then
58260      .PDFCompressionColorResolution = 300
58270     End If
58280   End If
58290   tStr = hOpt.Retrieve("PDFCompressionGreyCompression")
58300   If IsNumeric(tStr) Then
58310     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58320       .PDFCompressionGreyCompression = CLng(tStr)
58330      Else
58340       If UseStandard Then
58350        .PDFCompressionGreyCompression = 1
58360       End If
58370     End If
58380    Else
58390     If UseStandard Then
58400      .PDFCompressionGreyCompression = 1
58410     End If
58420   End If
58430   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
58440   If IsNumeric(tStr) Then
58450     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
58460       .PDFCompressionGreyCompressionChoice = CLng(tStr)
58470      Else
58480       If UseStandard Then
58490        .PDFCompressionGreyCompressionChoice = 0
58500       End If
58510     End If
58520    Else
58530     If UseStandard Then
58540      .PDFCompressionGreyCompressionChoice = 0
58550     End If
58560   End If
58570   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGHighFactor")
58580   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58590     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58600       .PDFCompressionGreyCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58610      Else
58620       If UseStandard Then
58630        .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
58640       End If
58650     End If
58660    Else
58670     If UseStandard Then
58680      .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
58690     End If
58700   End If
58710   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGLowFactor")
58720   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58730     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58740       .PDFCompressionGreyCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58750      Else
58760       If UseStandard Then
58770        .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58780       End If
58790     End If
58800    Else
58810     If UseStandard Then
58820      .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58830     End If
58840   End If
58850   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGManualFactor")
58860   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58870     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58880       .PDFCompressionGreyCompressionJPEGManualFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58890      Else
58900       If UseStandard Then
58910        .PDFCompressionGreyCompressionJPEGManualFactor = Replace$("3.0", ".", GetDecimalChar)
58920       End If
58930     End If
58940    Else
58950     If UseStandard Then
58960      .PDFCompressionGreyCompressionJPEGManualFactor = Replace$("3.0", ".", GetDecimalChar)
58970     End If
58980   End If
58990   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMaximumFactor")
59000   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
59010     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
59020       .PDFCompressionGreyCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
59030      Else
59040       If UseStandard Then
59050        .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
59060       End If
59070     End If
59080    Else
59090     If UseStandard Then
59100      .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
59110     End If
59120   End If
59130   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMediumFactor")
59140   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
59150     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
59160       .PDFCompressionGreyCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
59170      Else
59180       If UseStandard Then
59190        .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
59200       End If
59210     End If
59220    Else
59230     If UseStandard Then
59240      .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
59250     End If
59260   End If
59270   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMinimumFactor")
59280   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
59290     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
59300       .PDFCompressionGreyCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
59310      Else
59320       If UseStandard Then
59330        .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
59340       End If
59350     End If
59360    Else
59370     If UseStandard Then
59380      .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
59390     End If
59400   End If
59410   tStr = hOpt.Retrieve("PDFCompressionGreyResample")
59420   If IsNumeric(tStr) Then
59430     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59440       .PDFCompressionGreyResample = CLng(tStr)
59450      Else
59460       If UseStandard Then
59470        .PDFCompressionGreyResample = 0
59480       End If
59490     End If
59500    Else
59510     If UseStandard Then
59520      .PDFCompressionGreyResample = 0
59530     End If
59540   End If
59550   tStr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
59560   If IsNumeric(tStr) Then
59570     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
59580       .PDFCompressionGreyResampleChoice = CLng(tStr)
59590      Else
59600       If UseStandard Then
59610        .PDFCompressionGreyResampleChoice = 0
59620       End If
59630     End If
59640    Else
59650     If UseStandard Then
59660      .PDFCompressionGreyResampleChoice = 0
59670     End If
59680   End If
59690   tStr = hOpt.Retrieve("PDFCompressionGreyResolution")
59700   If IsNumeric(tStr) Then
59710     If CLng(tStr) >= 0 Then
59720       .PDFCompressionGreyResolution = CLng(tStr)
59730      Else
59740       If UseStandard Then
59750        .PDFCompressionGreyResolution = 300
59760       End If
59770     End If
59780    Else
59790     If UseStandard Then
59800      .PDFCompressionGreyResolution = 300
59810     End If
59820   End If
59830   tStr = hOpt.Retrieve("PDFCompressionMonoCompression")
59840   If IsNumeric(tStr) Then
59850     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59860       .PDFCompressionMonoCompression = CLng(tStr)
59870      Else
59880       If UseStandard Then
59890        .PDFCompressionMonoCompression = 1
59900       End If
59910     End If
59920    Else
59930     If UseStandard Then
59940      .PDFCompressionMonoCompression = 1
59950     End If
59960   End If
59970   tStr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
59980   If IsNumeric(tStr) Then
59990     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
60000       .PDFCompressionMonoCompressionChoice = CLng(tStr)
60010      Else
60020       If UseStandard Then
60030        .PDFCompressionMonoCompressionChoice = 0
60040       End If
60050     End If
60060    Else
60070     If UseStandard Then
60080      .PDFCompressionMonoCompressionChoice = 0
60090     End If
60100   End If
60110   tStr = hOpt.Retrieve("PDFCompressionMonoResample")
60120   If IsNumeric(tStr) Then
60130     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60140       .PDFCompressionMonoResample = CLng(tStr)
60150      Else
60160       If UseStandard Then
60170        .PDFCompressionMonoResample = 0
60180       End If
60190     End If
60200    Else
60210     If UseStandard Then
60220      .PDFCompressionMonoResample = 0
60230     End If
60240   End If
60250   tStr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
60260   If IsNumeric(tStr) Then
60270     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
60280       .PDFCompressionMonoResampleChoice = CLng(tStr)
60290      Else
60300       If UseStandard Then
60310        .PDFCompressionMonoResampleChoice = 0
60320       End If
60330     End If
60340    Else
60350     If UseStandard Then
60360      .PDFCompressionMonoResampleChoice = 0
60370     End If
60380   End If
60390   tStr = hOpt.Retrieve("PDFCompressionMonoResolution")
60400   If IsNumeric(tStr) Then
60410     If CLng(tStr) >= 0 Then
60420       .PDFCompressionMonoResolution = CLng(tStr)
60430      Else
60440       If UseStandard Then
60450        .PDFCompressionMonoResolution = 1200
60460       End If
60470     End If
60480    Else
60490     If UseStandard Then
60500      .PDFCompressionMonoResolution = 1200
60510     End If
60520   End If
60530   tStr = hOpt.Retrieve("PDFCompressionTextCompression")
60540   If IsNumeric(tStr) Then
60550     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60560       .PDFCompressionTextCompression = CLng(tStr)
60570      Else
60580       If UseStandard Then
60590        .PDFCompressionTextCompression = 1
60600       End If
60610     End If
60620    Else
60630     If UseStandard Then
60640      .PDFCompressionTextCompression = 1
60650     End If
60660   End If
60670   tStr = hOpt.Retrieve("PDFDisallowCopy")
60680   If IsNumeric(tStr) Then
60690     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60700       .PDFDisallowCopy = CLng(tStr)
60710      Else
60720       If UseStandard Then
60730        .PDFDisallowCopy = 1
60740       End If
60750     End If
60760    Else
60770     If UseStandard Then
60780      .PDFDisallowCopy = 1
60790     End If
60800   End If
60810   tStr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
60820   If IsNumeric(tStr) Then
60830     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60840       .PDFDisallowModifyAnnotations = CLng(tStr)
60850      Else
60860       If UseStandard Then
60870        .PDFDisallowModifyAnnotations = 0
60880       End If
60890     End If
60900    Else
60910     If UseStandard Then
60920      .PDFDisallowModifyAnnotations = 0
60930     End If
60940   End If
60950   tStr = hOpt.Retrieve("PDFDisallowModifyContents")
60960   If IsNumeric(tStr) Then
60970     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60980       .PDFDisallowModifyContents = CLng(tStr)
60990      Else
61000       If UseStandard Then
61010        .PDFDisallowModifyContents = 0
61020       End If
61030     End If
61040    Else
61050     If UseStandard Then
61060      .PDFDisallowModifyContents = 0
61070     End If
61080   End If
61090   tStr = hOpt.Retrieve("PDFDisallowPrinting")
61100   If IsNumeric(tStr) Then
61110     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61120       .PDFDisallowPrinting = CLng(tStr)
61130      Else
61140       If UseStandard Then
61150        .PDFDisallowPrinting = 0
61160       End If
61170     End If
61180    Else
61190     If UseStandard Then
61200      .PDFDisallowPrinting = 0
61210     End If
61220   End If
61230   tStr = hOpt.Retrieve("PDFEncryptor")
61240   If IsNumeric(tStr) Then
61250     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
61260       .PDFEncryptor = CLng(tStr)
61270      Else
61280       If UseStandard Then
61290        .PDFEncryptor = 0
61300       End If
61310     End If
61320    Else
61330     If UseStandard Then
61340      .PDFEncryptor = 0
61350     End If
61360   End If
61370   tStr = hOpt.Retrieve("PDFFontsEmbedAll")
61380   If IsNumeric(tStr) Then
61390     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61400       .PDFFontsEmbedAll = CLng(tStr)
61410      Else
61420       If UseStandard Then
61430        .PDFFontsEmbedAll = 1
61440       End If
61450     End If
61460    Else
61470     If UseStandard Then
61480      .PDFFontsEmbedAll = 1
61490     End If
61500   End If
61510   tStr = hOpt.Retrieve("PDFFontsSubSetFonts")
61520   If IsNumeric(tStr) Then
61530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61540       .PDFFontsSubSetFonts = CLng(tStr)
61550      Else
61560       If UseStandard Then
61570        .PDFFontsSubSetFonts = 1
61580       End If
61590     End If
61600    Else
61610     If UseStandard Then
61620      .PDFFontsSubSetFonts = 1
61630     End If
61640   End If
61650   tStr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
61660   If IsNumeric(tStr) Then
61670     If CLng(tStr) >= 0 Then
61680       .PDFFontsSubSetFontsPercent = CLng(tStr)
61690      Else
61700       If UseStandard Then
61710        .PDFFontsSubSetFontsPercent = 100
61720       End If
61730     End If
61740    Else
61750     If UseStandard Then
61760      .PDFFontsSubSetFontsPercent = 100
61770     End If
61780   End If
61790   tStr = hOpt.Retrieve("PDFGeneralASCII85")
61800   If IsNumeric(tStr) Then
61810     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61820       .PDFGeneralASCII85 = CLng(tStr)
61830      Else
61840       If UseStandard Then
61850        .PDFGeneralASCII85 = 0
61860       End If
61870     End If
61880    Else
61890     If UseStandard Then
61900      .PDFGeneralASCII85 = 0
61910     End If
61920   End If
61930   tStr = hOpt.Retrieve("PDFGeneralAutorotate")
61940   If IsNumeric(tStr) Then
61950     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
61960       .PDFGeneralAutorotate = CLng(tStr)
61970      Else
61980       If UseStandard Then
61990        .PDFGeneralAutorotate = 2
62000       End If
62010     End If
62020    Else
62030     If UseStandard Then
62040      .PDFGeneralAutorotate = 2
62050     End If
62060   End If
62070   tStr = hOpt.Retrieve("PDFGeneralCompatibility")
62080   If IsNumeric(tStr) Then
62090     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
62100       .PDFGeneralCompatibility = CLng(tStr)
62110      Else
62120       If UseStandard Then
62130        .PDFGeneralCompatibility = 2
62140       End If
62150     End If
62160    Else
62170     If UseStandard Then
62180      .PDFGeneralCompatibility = 2
62190     End If
62200   End If
62210   tStr = hOpt.Retrieve("PDFGeneralDefault")
62220   If IsNumeric(tStr) Then
62230     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
62240       .PDFGeneralDefault = CLng(tStr)
62250      Else
62260       If UseStandard Then
62270        .PDFGeneralDefault = 0
62280       End If
62290     End If
62300    Else
62310     If UseStandard Then
62320      .PDFGeneralDefault = 0
62330     End If
62340   End If
62350   tStr = hOpt.Retrieve("PDFGeneralOverprint")
62360   If IsNumeric(tStr) Then
62370     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
62380       .PDFGeneralOverprint = CLng(tStr)
62390      Else
62400       If UseStandard Then
62410        .PDFGeneralOverprint = 0
62420       End If
62430     End If
62440    Else
62450     If UseStandard Then
62460      .PDFGeneralOverprint = 0
62470     End If
62480   End If
62490   tStr = hOpt.Retrieve("PDFGeneralResolution")
62500   If IsNumeric(tStr) Then
62510     If CLng(tStr) >= 0 Then
62520       .PDFGeneralResolution = CLng(tStr)
62530      Else
62540       If UseStandard Then
62550        .PDFGeneralResolution = 600
62560       End If
62570     End If
62580    Else
62590     If UseStandard Then
62600      .PDFGeneralResolution = 600
62610     End If
62620   End If
62630   tStr = hOpt.Retrieve("PDFHighEncryption")
62640   If IsNumeric(tStr) Then
62650     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62660       .PDFHighEncryption = CLng(tStr)
62670      Else
62680       If UseStandard Then
62690        .PDFHighEncryption = 0
62700       End If
62710     End If
62720    Else
62730     If UseStandard Then
62740      .PDFHighEncryption = 0
62750     End If
62760   End If
62770   tStr = hOpt.Retrieve("PDFLowEncryption")
62780   If IsNumeric(tStr) Then
62790     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62800       .PDFLowEncryption = CLng(tStr)
62810      Else
62820       If UseStandard Then
62830        .PDFLowEncryption = 1
62840       End If
62850     End If
62860    Else
62870     If UseStandard Then
62880      .PDFLowEncryption = 1
62890     End If
62900   End If
62910   tStr = hOpt.Retrieve("PDFOptimize")
62920   If IsNumeric(tStr) Then
62930     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62940       .PDFOptimize = CLng(tStr)
62950      Else
62960       If UseStandard Then
62970        .PDFOptimize = 0
62980       End If
62990     End If
63000    Else
63010     If UseStandard Then
63020      .PDFOptimize = 0
63030     End If
63040   End If
63050   tStr = hOpt.Retrieve("PDFOwnerPass")
63060   If IsNumeric(tStr) Then
63070     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63080       .PDFOwnerPass = CLng(tStr)
63090      Else
63100       If UseStandard Then
63110        .PDFOwnerPass = 0
63120       End If
63130     End If
63140    Else
63150     If UseStandard Then
63160      .PDFOwnerPass = 0
63170     End If
63180   End If
63190   tStr = hOpt.Retrieve("PDFOwnerPasswordString")
63200   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63210     .PDFOwnerPasswordString = ""
63220    Else
63230     If LenB(tStr) > 0 Then
63240      .PDFOwnerPasswordString = tStr
63250     End If
63260   End If
63270   tStr = hOpt.Retrieve("PDFSigningMultiSignature")
63280   If IsNumeric(tStr) Then
63290     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63300       .PDFSigningMultiSignature = CLng(tStr)
63310      Else
63320       If UseStandard Then
63330        .PDFSigningMultiSignature = 0
63340       End If
63350     End If
63360    Else
63370     If UseStandard Then
63380      .PDFSigningMultiSignature = 0
63390     End If
63400   End If
63410   tStr = hOpt.Retrieve("PDFSigningPFXFile")
63420   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63430     .PDFSigningPFXFile = ""
63440    Else
63450     If LenB(tStr) > 0 Then
63460      .PDFSigningPFXFile = tStr
63470     End If
63480   End If
63490   tStr = hOpt.Retrieve("PDFSigningPFXFilePassword")
63500   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63510     .PDFSigningPFXFilePassword = ""
63520    Else
63530     If LenB(tStr) > 0 Then
63540      .PDFSigningPFXFilePassword = tStr
63550     End If
63560   End If
63570   tStr = hOpt.Retrieve("PDFSigningSignatureContact")
63580   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63590     .PDFSigningSignatureContact = ""
63600    Else
63610     If LenB(tStr) > 0 Then
63620      .PDFSigningSignatureContact = tStr
63630     End If
63640   End If
63650   tStr = hOpt.Retrieve("PDFSigningSignatureLeftX")
63660   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63670     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63680       .PDFSigningSignatureLeftX = CDbl(Replace$(tStr, ".", GetDecimalChar))
63690      Else
63700       If UseStandard Then
63710        .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
63720       End If
63730     End If
63740    Else
63750     If UseStandard Then
63760      .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
63770     End If
63780   End If
63790   tStr = hOpt.Retrieve("PDFSigningSignatureLeftY")
63800   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63810     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63820       .PDFSigningSignatureLeftY = CDbl(Replace$(tStr, ".", GetDecimalChar))
63830      Else
63840       If UseStandard Then
63850        .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
63860       End If
63870     End If
63880    Else
63890     If UseStandard Then
63900      .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
63910     End If
63920   End If
63930   tStr = hOpt.Retrieve("PDFSigningSignatureLocation")
63940   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63950     .PDFSigningSignatureLocation = ""
63960    Else
63970     If LenB(tStr) > 0 Then
63980      .PDFSigningSignatureLocation = tStr
63990     End If
64000   End If
64010   tStr = hOpt.Retrieve("PDFSigningSignatureOnPage")
64020   If IsNumeric(tStr) Then
64030     If CLng(tStr) >= 0 Then
64040       .PDFSigningSignatureOnPage = CLng(tStr)
64050      Else
64060       If UseStandard Then
64070        .PDFSigningSignatureOnPage = 1
64080       End If
64090     End If
64100    Else
64110     If UseStandard Then
64120      .PDFSigningSignatureOnPage = 1
64130     End If
64140   End If
64150   tStr = hOpt.Retrieve("PDFSigningSignatureReason")
64160   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64170     .PDFSigningSignatureReason = ""
64180    Else
64190     If LenB(tStr) > 0 Then
64200      .PDFSigningSignatureReason = tStr
64210     End If
64220   End If
64230   tStr = hOpt.Retrieve("PDFSigningSignatureRightX")
64240   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
64250     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
64260       .PDFSigningSignatureRightX = CDbl(Replace$(tStr, ".", GetDecimalChar))
64270      Else
64280       If UseStandard Then
64290        .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
64300       End If
64310     End If
64320    Else
64330     If UseStandard Then
64340      .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
64350     End If
64360   End If
64370   tStr = hOpt.Retrieve("PDFSigningSignatureRightY")
64380   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
64390     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
64400       .PDFSigningSignatureRightY = CDbl(Replace$(tStr, ".", GetDecimalChar))
64410      Else
64420       If UseStandard Then
64430        .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
64440       End If
64450     End If
64460    Else
64470     If UseStandard Then
64480      .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
64490     End If
64500   End If
64510   tStr = hOpt.Retrieve("PDFSigningSignatureVisible")
64520   If IsNumeric(tStr) Then
64530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64540       .PDFSigningSignatureVisible = CLng(tStr)
64550      Else
64560       If UseStandard Then
64570        .PDFSigningSignatureVisible = 0
64580       End If
64590     End If
64600    Else
64610     If UseStandard Then
64620      .PDFSigningSignatureVisible = 0
64630     End If
64640   End If
64650   tStr = hOpt.Retrieve("PDFSigningSignPDF")
64660   If IsNumeric(tStr) Then
64670     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64680       .PDFSigningSignPDF = CLng(tStr)
64690      Else
64700       If UseStandard Then
64710        .PDFSigningSignPDF = 0
64720       End If
64730     End If
64740    Else
64750     If UseStandard Then
64760      .PDFSigningSignPDF = 0
64770     End If
64780   End If
64790   tStr = hOpt.Retrieve("PDFUpdateMetadata")
64800   If IsNumeric(tStr) Then
64810     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
64820       .PDFUpdateMetadata = CLng(tStr)
64830      Else
64840       If UseStandard Then
64850        .PDFUpdateMetadata = 1
64860       End If
64870     End If
64880    Else
64890     If UseStandard Then
64900      .PDFUpdateMetadata = 1
64910     End If
64920   End If
64930   tStr = hOpt.Retrieve("PDFUserPass")
64940   If IsNumeric(tStr) Then
64950     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64960       .PDFUserPass = CLng(tStr)
64970      Else
64980       If UseStandard Then
64990        .PDFUserPass = 0
65000       End If
65010     End If
65020    Else
65030     If UseStandard Then
65040      .PDFUserPass = 0
65050     End If
65060   End If
65070   tStr = hOpt.Retrieve("PDFUserPasswordString")
65080   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
65090     .PDFUserPasswordString = ""
65100    Else
65110     If LenB(tStr) > 0 Then
65120      .PDFUserPasswordString = tStr
65130     End If
65140   End If
65150   tStr = hOpt.Retrieve("PDFUseSecurity")
65160   If IsNumeric(tStr) Then
65170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65180       .PDFUseSecurity = CLng(tStr)
65190      Else
65200       If UseStandard Then
65210        .PDFUseSecurity = 0
65220       End If
65230     End If
65240    Else
65250     If UseStandard Then
65260      .PDFUseSecurity = 0
65270     End If
65280   End If
65290   tStr = hOpt.Retrieve("PNGColorscount")
65300   If IsNumeric(tStr) Then
65310     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
65320       .PNGColorscount = CLng(tStr)
65330      Else
65340       If UseStandard Then
65350        .PNGColorscount = 0
65360       End If
65370     End If
65380    Else
65390     If UseStandard Then
65400      .PNGColorscount = 0
65410     End If
65420   End If
65430   tStr = hOpt.Retrieve("PNGResolution")
65440   If IsNumeric(tStr) Then
65450     If CLng(tStr) >= 1 Then
65460       .PNGResolution = CLng(tStr)
65470      Else
65480       If UseStandard Then
65490        .PNGResolution = 150
65500       End If
65510     End If
65520    Else
65530     If UseStandard Then
65540      .PNGResolution = 150
65550     End If
65560   End If
65570   tStr = hOpt.Retrieve("PrintAfterSaving")
65580   If IsNumeric(tStr) Then
65590     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65600       .PrintAfterSaving = CLng(tStr)
65610      Else
65620       If UseStandard Then
65630        .PrintAfterSaving = 0
65640       End If
65650     End If
65660    Else
65670     If UseStandard Then
65680      .PrintAfterSaving = 0
65690     End If
65700   End If
65710   tStr = hOpt.Retrieve("PrintAfterSavingBitsPerPixel")
65720   If IsNumeric(tStr) Then
65730     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
65740       .PrintAfterSavingBitsPerPixel = CLng(tStr)
65750      Else
65760       If UseStandard Then
65770        .PrintAfterSavingBitsPerPixel = 2
65780       End If
65790     End If
65800    Else
65810     If UseStandard Then
65820      .PrintAfterSavingBitsPerPixel = 2
65830     End If
65840   End If
65850   tStr = hOpt.Retrieve("PrintAfterSavingDuplex")
65860   If IsNumeric(tStr) Then
65870     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65880       .PrintAfterSavingDuplex = CLng(tStr)
65890      Else
65900       If UseStandard Then
65910        .PrintAfterSavingDuplex = 0
65920       End If
65930     End If
65940    Else
65950     If UseStandard Then
65960      .PrintAfterSavingDuplex = 0
65970     End If
65980   End If
65990   tStr = hOpt.Retrieve("PrintAfterSavingMaxResolution")
66000   If IsNumeric(tStr) Then
66010     If CLng(tStr) >= 72 Then
66020       .PrintAfterSavingMaxResolution = CLng(tStr)
66030      Else
66040       If UseStandard Then
66050        .PrintAfterSavingMaxResolution = 600
66060       End If
66070     End If
66080    Else
66090     If UseStandard Then
66100      .PrintAfterSavingMaxResolution = 600
66110     End If
66120   End If
66130   tStr = hOpt.Retrieve("PrintAfterSavingMaxResolutionEnabled")
66140   If IsNumeric(tStr) Then
66150     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66160       .PrintAfterSavingMaxResolutionEnabled = CLng(tStr)
66170      Else
66180       If UseStandard Then
66190        .PrintAfterSavingMaxResolutionEnabled = 0
66200       End If
66210     End If
66220    Else
66230     If UseStandard Then
66240      .PrintAfterSavingMaxResolutionEnabled = 0
66250     End If
66260   End If
66270   tStr = hOpt.Retrieve("PrintAfterSavingNoCancel")
66280   If IsNumeric(tStr) Then
66290     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66300       .PrintAfterSavingNoCancel = CLng(tStr)
66310      Else
66320       If UseStandard Then
66330        .PrintAfterSavingNoCancel = 0
66340       End If
66350     End If
66360    Else
66370     If UseStandard Then
66380      .PrintAfterSavingNoCancel = 0
66390     End If
66400   End If
66410   tStr = hOpt.Retrieve("PrintAfterSavingPrinter")
66420   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
66430     .PrintAfterSavingPrinter = ""
66440    Else
66450     If LenB(tStr) > 0 Then
66460      .PrintAfterSavingPrinter = tStr
66470     End If
66480   End If
66490   tStr = hOpt.Retrieve("PrintAfterSavingQueryUser")
66500   If IsNumeric(tStr) Then
66510     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
66520       .PrintAfterSavingQueryUser = CLng(tStr)
66530      Else
66540       If UseStandard Then
66550        .PrintAfterSavingQueryUser = 0
66560       End If
66570     End If
66580    Else
66590     If UseStandard Then
66600      .PrintAfterSavingQueryUser = 0
66610     End If
66620   End If
66630   tStr = hOpt.Retrieve("PrintAfterSavingTumble")
66640   If IsNumeric(tStr) Then
66650     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
66660       .PrintAfterSavingTumble = CLng(tStr)
66670      Else
66680       If UseStandard Then
66690        .PrintAfterSavingTumble = 0
66700       End If
66710     End If
66720    Else
66730     If UseStandard Then
66740      .PrintAfterSavingTumble = 0
66750     End If
66760   End If
66770   tStr = hOpt.Retrieve("PrinterStop")
66780   If IsNumeric(tStr) Then
66790     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66800       .PrinterStop = CLng(tStr)
66810      Else
66820       If UseStandard Then
66830        .PrinterStop = 0
66840       End If
66850     End If
66860    Else
66870     If UseStandard Then
66880      .PrinterStop = 0
66890     End If
66900   End If
66910   tStr = hOpt.Retrieve("PrinterTemppath")
66920   WriteToSpecialLogfile "hOpt.Retrieve(""PrinterTemppath"")=" & tStr
66930   WriteToSpecialLogfile "Options.PrinterTemppath1=" & .PrinterTemppath
66940   If hkey1 = HKEY_USERS Then
66950     If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then
66960       .PrinterTemppath = tStr
66970      Else
66980       If UseStandard Then
66990         .PrinterTemppath = GetTempPath
67000        Else
67010         .PrinterTemppath = tStr
67020       End If
67030     End If
67040    Else
67050     If LenB(Trim$(tStr)) > 0 Then
67060      If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then
67070        .PrinterTemppath = tStr
67080       Else
67090        MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))
67100        If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then
67110          If UseStandard Then
67120            .PrinterTemppath = GetTempPath
67130           Else
67140            .PrinterTemppath = ""
67150            If NoMsg = False Then
67160             MsgBox "PrinterTemppath: '" & tStr & "' = '" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & "'" & _
             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
67180            End If
67190          End If
67200         Else
67210          .PrinterTemppath = tStr
67220        End If
67230      End If
67240     End If
67250   End If
67260   WriteToSpecialLogfile "Options.PrinterTemppath2=" & .PrinterTemppath
67270   tStr = hOpt.Retrieve("ProcessPriority")
67280   If IsNumeric(tStr) Then
67290     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
67300       .ProcessPriority = CLng(tStr)
67310      Else
67320       If UseStandard Then
67330        .ProcessPriority = 1
67340       End If
67350     End If
67360    Else
67370     If UseStandard Then
67380      .ProcessPriority = 1
67390     End If
67400   End If
67410   tStr = hOpt.Retrieve("ProgramFont")
67420   If LenB(tStr) = 0 And LenB("MS Sans Serif") > 0 And UseStandard Then
67430     .ProgramFont = "MS Sans Serif"
67440    Else
67450     If LenB(tStr) > 0 Then
67460      .ProgramFont = tStr
67470     End If
67480   End If
67490   tStr = hOpt.Retrieve("ProgramFontCharset")
67500   If IsNumeric(tStr) Then
67510     If CLng(tStr) >= 0 Then
67520       .ProgramFontCharset = CLng(tStr)
67530      Else
67540       If UseStandard Then
67550        .ProgramFontCharset = 0
67560       End If
67570     End If
67580    Else
67590     If UseStandard Then
67600      .ProgramFontCharset = 0
67610     End If
67620   End If
67630   tStr = hOpt.Retrieve("ProgramFontSize")
67640   If IsNumeric(tStr) Then
67650     If CLng(tStr) >= 6 And CLng(tStr) <= 72 Then
67660       .ProgramFontSize = CLng(tStr)
67670      Else
67680       If UseStandard Then
67690        .ProgramFontSize = 8
67700       End If
67710     End If
67720    Else
67730     If UseStandard Then
67740      .ProgramFontSize = 8
67750     End If
67760   End If
67770   tStr = hOpt.Retrieve("PSDColorsCount")
67780   If IsNumeric(tStr) Then
67790     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
67800       .PSDColorsCount = CLng(tStr)
67810      Else
67820       If UseStandard Then
67830        .PSDColorsCount = 0
67840       End If
67850     End If
67860    Else
67870     If UseStandard Then
67880      .PSDColorsCount = 0
67890     End If
67900   End If
67910   tStr = hOpt.Retrieve("PSDResolution")
67920   If IsNumeric(tStr) Then
67930     If CLng(tStr) >= 1 Then
67940       .PSDResolution = CLng(tStr)
67950      Else
67960       If UseStandard Then
67970        .PSDResolution = 150
67980       End If
67990     End If
68000    Else
68010     If UseStandard Then
68020      .PSDResolution = 150
68030     End If
68040   End If
68050   tStr = hOpt.Retrieve("PSLanguageLevel")
68060   If IsNumeric(tStr) Then
68070     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
68080       .PSLanguageLevel = CLng(tStr)
68090      Else
68100       If UseStandard Then
68110        .PSLanguageLevel = 2
68120       End If
68130     End If
68140    Else
68150     If UseStandard Then
68160      .PSLanguageLevel = 2
68170     End If
68180   End If
68190   tStr = hOpt.Retrieve("RAWColorsCount")
68200   If IsNumeric(tStr) Then
68210     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
68220       .RAWColorsCount = CLng(tStr)
68230      Else
68240       If UseStandard Then
68250        .RAWColorsCount = 0
68260       End If
68270     End If
68280    Else
68290     If UseStandard Then
68300      .RAWColorsCount = 0
68310     End If
68320   End If
68330   tStr = hOpt.Retrieve("RAWResolution")
68340   If IsNumeric(tStr) Then
68350     If CLng(tStr) >= 1 Then
68360       .RAWResolution = CLng(tStr)
68370      Else
68380       If UseStandard Then
68390        .RAWResolution = 150
68400       End If
68410     End If
68420    Else
68430     If UseStandard Then
68440      .RAWResolution = 150
68450     End If
68460   End If
68470   tStr = hOpt.Retrieve("RemoveAllKnownFileExtensions")
68480   If IsNumeric(tStr) Then
68490     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68500       .RemoveAllKnownFileExtensions = CLng(tStr)
68510      Else
68520       If UseStandard Then
68530        .RemoveAllKnownFileExtensions = 1
68540       End If
68550     End If
68560    Else
68570     If UseStandard Then
68580      .RemoveAllKnownFileExtensions = 1
68590     End If
68600   End If
68610   tStr = hOpt.Retrieve("RemoveSpaces")
68620   If IsNumeric(tStr) Then
68630     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68640       .RemoveSpaces = CLng(tStr)
68650      Else
68660       If UseStandard Then
68670        .RemoveSpaces = 1
68680       End If
68690     End If
68700    Else
68710     If UseStandard Then
68720      .RemoveSpaces = 1
68730     End If
68740   End If
68750   tStr = hOpt.Retrieve("RunProgramAfterSaving")
68760   If IsNumeric(tStr) Then
68770     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68780       .RunProgramAfterSaving = CLng(tStr)
68790      Else
68800       If UseStandard Then
68810        .RunProgramAfterSaving = 0
68820       End If
68830     End If
68840    Else
68850     If UseStandard Then
68860      .RunProgramAfterSaving = 0
68870     End If
68880   End If
68890   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramname")
68900   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
68910     .RunProgramAfterSavingProgramname = ""
68920    Else
68930     If LenB(tStr) > 0 Then
68940      .RunProgramAfterSavingProgramname = tStr
68950     End If
68960   End If
68970   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramParameters")
68980   If LenB(tStr) = 0 And LenB("""<OutputFilename>""") > 0 And UseStandard Then
68990     .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
69000    Else
69010     If LenB(tStr) > 0 Then
69020      .RunProgramAfterSavingProgramParameters = tStr
69030     End If
69040   End If
69050   tStr = hOpt.Retrieve("RunProgramAfterSavingWaitUntilReady")
69060   If IsNumeric(tStr) Then
69070     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69080       .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
69090      Else
69100       If UseStandard Then
69110        .RunProgramAfterSavingWaitUntilReady = 1
69120       End If
69130     End If
69140    Else
69150     If UseStandard Then
69160      .RunProgramAfterSavingWaitUntilReady = 1
69170     End If
69180   End If
69190   tStr = hOpt.Retrieve("RunProgramAfterSavingWindowstyle")
69200   If IsNumeric(tStr) Then
69210     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
69220       .RunProgramAfterSavingWindowstyle = CLng(tStr)
69230      Else
69240       If UseStandard Then
69250        .RunProgramAfterSavingWindowstyle = 1
69260       End If
69270     End If
69280    Else
69290     If UseStandard Then
69300      .RunProgramAfterSavingWindowstyle = 1
69310     End If
69320   End If
69330   tStr = hOpt.Retrieve("RunProgramBeforeSaving")
69340   If IsNumeric(tStr) Then
69350     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69360       .RunProgramBeforeSaving = CLng(tStr)
69370      Else
69380       If UseStandard Then
69390        .RunProgramBeforeSaving = 0
69400       End If
69410     End If
69420    Else
69430     If UseStandard Then
69440      .RunProgramBeforeSaving = 0
69450     End If
69460   End If
69470   tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramname")
69480   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69490     .RunProgramBeforeSavingProgramname = ""
69500    Else
69510     If LenB(tStr) > 0 Then
69520      .RunProgramBeforeSavingProgramname = tStr
69530     End If
69540   End If
69550   tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramParameters")
69560   If LenB(tStr) = 0 And LenB("""<TempFilename>""") > 0 And UseStandard Then
69570     .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
69580    Else
69590     If LenB(tStr) > 0 Then
69600      .RunProgramBeforeSavingProgramParameters = tStr
69610     End If
69620   End If
69630   tStr = hOpt.Retrieve("RunProgramBeforeSavingWindowstyle")
69640   If IsNumeric(tStr) Then
69650     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
69660       .RunProgramBeforeSavingWindowstyle = CLng(tStr)
69670      Else
69680       If UseStandard Then
69690        .RunProgramBeforeSavingWindowstyle = 1
69700       End If
69710     End If
69720    Else
69730     If UseStandard Then
69740      .RunProgramBeforeSavingWindowstyle = 1
69750     End If
69760   End If
69770   tStr = hOpt.Retrieve("SaveFilename")
69780   If LenB(tStr) = 0 And LenB("<Title>") > 0 And UseStandard Then
69790     .SaveFilename = "<Title>"
69800    Else
69810     If LenB(tStr) > 0 Then
69820      .SaveFilename = tStr
69830     End If
69840   End If
69850   tStr = hOpt.Retrieve("SendEmailAfterAutoSaving")
69860   If IsNumeric(tStr) Then
69870     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69880       .SendEmailAfterAutoSaving = CLng(tStr)
69890      Else
69900       If UseStandard Then
69910        .SendEmailAfterAutoSaving = 0
69920       End If
69930     End If
69940    Else
69950     If UseStandard Then
69960      .SendEmailAfterAutoSaving = 0
69970     End If
69980   End If
69990   tStr = hOpt.Retrieve("SendMailMethod")
70000   If IsNumeric(tStr) Then
70010     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
70020       .SendMailMethod = CLng(tStr)
70030      Else
70040       If UseStandard Then
70050        .SendMailMethod = 0
70060       End If
70070     End If
70080    Else
70090     If UseStandard Then
70100      .SendMailMethod = 0
70110     End If
70120   End If
70130   tStr = hOpt.Retrieve("ShowAnimation")
70140   If IsNumeric(tStr) Then
70150     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70160       .ShowAnimation = CLng(tStr)
70170      Else
70180       If UseStandard Then
70190        .ShowAnimation = 1
70200       End If
70210     End If
70220    Else
70230     If UseStandard Then
70240      .ShowAnimation = 1
70250     End If
70260   End If
70270   tStr = hOpt.Retrieve("StampFontColor")
70280   If LenB(tStr) = 0 And LenB("#FF0000") > 0 And UseStandard Then
70290     .StampFontColor = "#FF0000"
70300    Else
70310     If LenB(tStr) > 0 Then
70320      .StampFontColor = tStr
70330     End If
70340   End If
70350   tStr = hOpt.Retrieve("StampFontname")
70360   If LenB(tStr) = 0 And LenB("Arial") > 0 And UseStandard Then
70370     .StampFontname = "Arial"
70380    Else
70390     If LenB(tStr) > 0 Then
70400      .StampFontname = tStr
70410     End If
70420   End If
70430   tStr = hOpt.Retrieve("StampFontsize")
70440   If IsNumeric(tStr) Then
70450     If CLng(tStr) >= 1 Then
70460       .StampFontsize = CLng(tStr)
70470      Else
70480       If UseStandard Then
70490        .StampFontsize = 48
70500       End If
70510     End If
70520    Else
70530     If UseStandard Then
70540      .StampFontsize = 48
70550     End If
70560   End If
70570   tStr = hOpt.Retrieve("StampOutlineFontthickness")
70580   If IsNumeric(tStr) Then
70590     If CLng(tStr) >= 0 Then
70600       .StampOutlineFontthickness = CLng(tStr)
70610      Else
70620       If UseStandard Then
70630        .StampOutlineFontthickness = 0
70640       End If
70650     End If
70660    Else
70670     If UseStandard Then
70680      .StampOutlineFontthickness = 0
70690     End If
70700   End If
70710   tStr = hOpt.Retrieve("StampString")
70720   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70730     .StampString = ""
70740    Else
70750     If LenB(tStr) > 0 Then
70760      .StampString = tStr
70770     End If
70780   End If
70790   tStr = hOpt.Retrieve("StampUseOutlineFont")
70800   If IsNumeric(tStr) Then
70810     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70820       .StampUseOutlineFont = CLng(tStr)
70830      Else
70840       If UseStandard Then
70850        .StampUseOutlineFont = 1
70860       End If
70870     End If
70880    Else
70890     If UseStandard Then
70900      .StampUseOutlineFont = 1
70910     End If
70920   End If
70930   tStr = hOpt.Retrieve("StandardAuthor")
70940   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70950     .StandardAuthor = ""
70960    Else
70970     If LenB(tStr) > 0 Then
70980      .StandardAuthor = tStr
70990     End If
71000   End If
71010   tStr = hOpt.Retrieve("StandardCreationdate")
71020   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
71030     .StandardCreationdate = ""
71040    Else
71050     If LenB(tStr) > 0 Then
71060      .StandardCreationdate = tStr
71070     End If
71080   End If
71090   tStr = hOpt.Retrieve("StandardDateformat")
71100   If LenB(tStr) = 0 And LenB("YYYYMMDDHHNNSS") > 0 And UseStandard Then
71110     .StandardDateformat = "YYYYMMDDHHNNSS"
71120    Else
71130     If LenB(tStr) > 0 Then
71140      .StandardDateformat = tStr
71150     End If
71160   End If
71170   tStr = hOpt.Retrieve("StandardKeywords")
71180   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
71190     .StandardKeywords = ""
71200    Else
71210     If LenB(tStr) > 0 Then
71220      .StandardKeywords = tStr
71230     End If
71240   End If
71250   tStr = hOpt.Retrieve("StandardMailDomain")
71260   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
71270     .StandardMailDomain = ""
71280    Else
71290     If LenB(tStr) > 0 Then
71300      .StandardMailDomain = tStr
71310     End If
71320   End If
71330   tStr = hOpt.Retrieve("StandardModifydate")
71340   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
71350     .StandardModifydate = ""
71360    Else
71370     If LenB(tStr) > 0 Then
71380      .StandardModifydate = tStr
71390     End If
71400   End If
71410   tStr = hOpt.Retrieve("StandardSaveformat")
71420   If IsNumeric(tStr) Then
71430     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
71440       .StandardSaveformat = CLng(tStr)
71450      Else
71460       If UseStandard Then
71470        .StandardSaveformat = 0
71480       End If
71490     End If
71500    Else
71510     If UseStandard Then
71520      .StandardSaveformat = 0
71530     End If
71540   End If
71550   tStr = hOpt.Retrieve("StandardSubject")
71560   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
71570     .StandardSubject = ""
71580    Else
71590     If LenB(tStr) > 0 Then
71600      .StandardSubject = tStr
71610     End If
71620   End If
71630   tStr = hOpt.Retrieve("StandardTitle")
71640   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
71650     .StandardTitle = ""
71660    Else
71670     If LenB(tStr) > 0 Then
71680      .StandardTitle = tStr
71690     End If
71700   End If
71710   tStr = hOpt.Retrieve("StartStandardProgram")
71720   If IsNumeric(tStr) Then
71730     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71740       .StartStandardProgram = CLng(tStr)
71750      Else
71760       If UseStandard Then
71770        .StartStandardProgram = 1
71780       End If
71790     End If
71800    Else
71810     If UseStandard Then
71820      .StartStandardProgram = 1
71830     End If
71840   End If
71850   tStr = hOpt.Retrieve("SVGResolution")
71860   If IsNumeric(tStr) Then
71870     If CLng(tStr) >= 1 Then
71880       .SVGResolution = CLng(tStr)
71890      Else
71900       If UseStandard Then
71910        .SVGResolution = 72
71920       End If
71930     End If
71940    Else
71950     If UseStandard Then
71960      .SVGResolution = 72
71970     End If
71980   End If
71990   tStr = hOpt.Retrieve("TIFFColorscount")
72000   If IsNumeric(tStr) Then
72010     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
72020       .TIFFColorscount = CLng(tStr)
72030      Else
72040       If UseStandard Then
72050        .TIFFColorscount = 0
72060       End If
72070     End If
72080    Else
72090     If UseStandard Then
72100      .TIFFColorscount = 0
72110     End If
72120   End If
72130   tStr = hOpt.Retrieve("TIFFResolution")
72140   If IsNumeric(tStr) Then
72150     If CLng(tStr) >= 1 Then
72160       .TIFFResolution = CLng(tStr)
72170      Else
72180       If UseStandard Then
72190        .TIFFResolution = 150
72200       End If
72210     End If
72220    Else
72230     If UseStandard Then
72240      .TIFFResolution = 150
72250     End If
72260   End If
72270   tStr = hOpt.Retrieve("Toolbars")
72280   If IsNumeric(tStr) Then
72290     If CLng(tStr) >= 0 Then
72300       .Toolbars = CLng(tStr)
72310      Else
72320       If UseStandard Then
72330        .Toolbars = 1
72340       End If
72350     End If
72360    Else
72370     If UseStandard Then
72380      .Toolbars = 1
72390     End If
72400   End If
72410   tStr = hOpt.Retrieve("UpdateInterval")
72420   If IsNumeric(tStr) Then
72430     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
72440       .UpdateInterval = CLng(tStr)
72450      Else
72460       If UseStandard Then
72470        .UpdateInterval = 2
72480       End If
72490     End If
72500    Else
72510     If UseStandard Then
72520      .UpdateInterval = 2
72530     End If
72540   End If
72550   tStr = hOpt.Retrieve("UseAutosave")
72560   If IsNumeric(tStr) Then
72570     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
72580       .UseAutosave = CLng(tStr)
72590      Else
72600       If UseStandard Then
72610        .UseAutosave = 0
72620       End If
72630     End If
72640    Else
72650     If UseStandard Then
72660      .UseAutosave = 0
72670     End If
72680   End If
72690   tStr = hOpt.Retrieve("UseAutosaveDirectory")
72700   If IsNumeric(tStr) Then
72710     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
72720       .UseAutosaveDirectory = CLng(tStr)
72730      Else
72740       If UseStandard Then
72750        .UseAutosaveDirectory = 1
72760       End If
72770     End If
72780    Else
72790     If UseStandard Then
72800      .UseAutosaveDirectory = 1
72810     End If
72820   End If
72830   tStr = hOpt.Retrieve("UseCreationDateNow")
72840   If IsNumeric(tStr) Then
72850     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
72860       .UseCreationDateNow = CLng(tStr)
72870      Else
72880       If UseStandard Then
72890        .UseCreationDateNow = 0
72900       End If
72910     End If
72920    Else
72930     If UseStandard Then
72940      .UseCreationDateNow = 0
72950     End If
72960   End If
72970   tStr = hOpt.Retrieve("UseCustomPaperSize")
72980   If LenB(tStr) = 0 And LenB("0") > 0 And UseStandard Then
72990     .UseCustomPaperSize = "0"
73000    Else
73010     If LenB(tStr) > 0 Then
73020      .UseCustomPaperSize = tStr
73030     End If
73040   End If
73050   tStr = hOpt.Retrieve("UseFixPapersize")
73060   If IsNumeric(tStr) Then
73070     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
73080       .UseFixPapersize = CLng(tStr)
73090      Else
73100       If UseStandard Then
73110        .UseFixPapersize = 0
73120       End If
73130     End If
73140    Else
73150     If UseStandard Then
73160      .UseFixPapersize = 0
73170     End If
73180   End If
73190   tStr = hOpt.Retrieve("UseStandardAuthor")
73200   If IsNumeric(tStr) Then
73210     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
73220       .UseStandardAuthor = CLng(tStr)
73230      Else
73240       If UseStandard Then
73250        .UseStandardAuthor = 0
73260       End If
73270     End If
73280    Else
73290     If UseStandard Then
73300      .UseStandardAuthor = 0
73310     End If
73320   End If
73330  End With
73340  Set ini = Nothing
73350  ReadOptionsINI = myOptions
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

Public Sub SaveOptions(sOptions As tOptions, Optional ProfileName As String = "")
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  CorrectOptionsBeforeSaving
50020
50030  ProfileName = Trim$(ProfileName)
50040
50050  If InstalledAsServer Then
50060    SaveOptionsREG sOptions, HKEY_LOCAL_MACHINE, ProfileName
50070   Else
50080    If HKLMProfileExists(ProfileName) = False Then
50090     SaveOptionsREG sOptions, , ProfileName
50100    End If
50110  End If
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

Public Sub SaveOption(sOptions As tOptions, OptionName As String, Optional ProfileName As String = "")
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If InstalledAsServer Then
50020    SaveOptionREG sOptions, OptionName, HKEY_LOCAL_MACHINE, ProfileName
50030   Else
50040    SaveOptionREG sOptions, OptionName, , ProfileName
50050  End If
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
50130   Case "ALLOWSPECIALGSCHARSINFILENAMES": ini.SaveKey CStr(Abs(.AllowSpecialGSCharsInFilenames)), "AllowSpecialGSCharsInFilenames"
50140   Case "AUTOSAVEDIRECTORY": ini.SaveKey CStr(.AutosaveDirectory), "AutosaveDirectory"
50150   Case "AUTOSAVEFILENAME": ini.SaveKey CStr(.AutosaveFilename), "AutosaveFilename"
50160   Case "AUTOSAVEFORMAT": ini.SaveKey CStr(.AutosaveFormat), "AutosaveFormat"
50170   Case "AUTOSAVESTARTSTANDARDPROGRAM": ini.SaveKey CStr(Abs(.AutosaveStartStandardProgram)), "AutosaveStartStandardProgram"
50180   Case "BMPCOLORSCOUNT": ini.SaveKey CStr(.BMPColorscount), "BMPColorscount"
50190   Case "BMPRESOLUTION": ini.SaveKey CStr(.BMPResolution), "BMPResolution"
50200   Case "CLIENTCOMPUTERRESOLVEIPADDRESS": ini.SaveKey CStr(Abs(.ClientComputerResolveIPAddress)), "ClientComputerResolveIPAddress"
50210   Case "COUNTER": ini.SaveKey CStr(.Counter), "Counter"
50220   Case "DEVICEHEIGHTPOINTS": ini.SaveKey Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), "DeviceHeightPoints"
50230   Case "DEVICEWIDTHPOINTS": ini.SaveKey Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), "DeviceWidthPoints"
50240   Case "DIRECTORYGHOSTSCRIPTBINARIES": ini.SaveKey CStr(.DirectoryGhostscriptBinaries), "DirectoryGhostscriptBinaries"
50250   Case "DIRECTORYGHOSTSCRIPTFONTS": ini.SaveKey CStr(.DirectoryGhostscriptFonts), "DirectoryGhostscriptFonts"
50260   Case "DIRECTORYGHOSTSCRIPTLIBRARIES": ini.SaveKey CStr(.DirectoryGhostscriptLibraries), "DirectoryGhostscriptLibraries"
50270   Case "DIRECTORYGHOSTSCRIPTRESOURCE": ini.SaveKey CStr(.DirectoryGhostscriptResource), "DirectoryGhostscriptResource"
50280   Case "DISABLEEMAIL": ini.SaveKey CStr(Abs(.DisableEmail)), "DisableEmail"
50290   Case "DONTUSEDOCUMENTSETTINGS": ini.SaveKey CStr(Abs(.DontUseDocumentSettings)), "DontUseDocumentSettings"
50300   Case "EPSLANGUAGELEVEL": ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
50310   Case "FILENAMESUBSTITUTIONS": ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
50320   Case "FILENAMESUBSTITUTIONSONLYINTITLE": ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
50330   Case "JPEGCOLORSCOUNT": ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50340   Case "JPEGQUALITY": ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50350   Case "JPEGRESOLUTION": ini.SaveKey CStr(.JPEGResolution), "JPEGResolution"
50360   Case "LANGUAGE": ini.SaveKey CStr(.Language), "Language"
50370   Case "LASTSAVEDIRECTORY": ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50380   Case "LASTUPDATECHECK": ini.SaveKey CStr(.LastUpdateCheck), "LastUpdateCheck"
50390   Case "LOGGING": ini.SaveKey CStr(Abs(.Logging)), "Logging"
50400   Case "LOGLINES": ini.SaveKey CStr(.LogLines), "LogLines"
50410   Case "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER": ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50420   Case "NOPROCESSINGATSTARTUP": ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
50430   Case "NOPSCHECK": ini.SaveKey CStr(Abs(.NoPSCheck)), "NoPSCheck"
50440   Case "ONEPAGEPERFILE": ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50450   Case "OPTIONSDESIGN": ini.SaveKey CStr(.OptionsDesign), "OptionsDesign"
50460   Case "OPTIONSENABLED": ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50470   Case "OPTIONSVISIBLE": ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50480   Case "PAPERSIZE": ini.SaveKey CStr(.Papersize), "Papersize"
50490   Case "PCLCOLORSCOUNT": ini.SaveKey CStr(.PCLColorsCount), "PCLColorsCount"
50500   Case "PCLRESOLUTION": ini.SaveKey CStr(.PCLResolution), "PCLResolution"
50510   Case "PCXCOLORSCOUNT": ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50520   Case "PCXRESOLUTION": ini.SaveKey CStr(.PCXResolution), "PCXResolution"
50530   Case "PDFALLOWASSEMBLY": ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50540   Case "PDFALLOWDEGRADEDPRINTING": ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50550   Case "PDFALLOWFILLIN": ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50560   Case "PDFALLOWSCREENREADERS": ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50570   Case "PDFCOLORSCMYKTORGB": ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50580   Case "PDFCOLORSCOLORMODEL": ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50590   Case "PDFCOLORSPRESERVEHALFTONE": ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50600   Case "PDFCOLORSPRESERVEOVERPRINT": ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50610   Case "PDFCOLORSPRESERVETRANSFER": ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50620   Case "PDFCOMPRESSIONCOLORCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50630   Case "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50640   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGHighFactor"
50650   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGLowFactor"
50660   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMANUALFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGManualFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGManualFactor"
50670   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMaximumFactor"
50680   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMediumFactor"
50690   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMinimumFactor"
50700   Case "PDFCOMPRESSIONCOLORRESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50710   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50720   Case "PDFCOMPRESSIONCOLORRESOLUTION": ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50730   Case "PDFCOMPRESSIONGREYCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50740   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50750   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGHighFactor"
50760   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGLowFactor"
50770   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMANUALFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGManualFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGManualFactor"
50780   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMaximumFactor"
50790   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMediumFactor"
50800   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMinimumFactor"
50810   Case "PDFCOMPRESSIONGREYRESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50820   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50830   Case "PDFCOMPRESSIONGREYRESOLUTION": ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50840   Case "PDFCOMPRESSIONMONOCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50850   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50860   Case "PDFCOMPRESSIONMONORESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50870   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50880   Case "PDFCOMPRESSIONMONORESOLUTION": ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50890   Case "PDFCOMPRESSIONTEXTCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50900   Case "PDFDISALLOWCOPY": ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50910   Case "PDFDISALLOWMODIFYANNOTATIONS": ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50920   Case "PDFDISALLOWMODIFYCONTENTS": ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50930   Case "PDFDISALLOWPRINTING": ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50940   Case "PDFENCRYPTOR": ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50950   Case "PDFFONTSEMBEDALL": ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50960   Case "PDFFONTSSUBSETFONTS": ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50970   Case "PDFFONTSSUBSETFONTSPERCENT": ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50980   Case "PDFGENERALASCII85": ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50990   Case "PDFGENERALAUTOROTATE": ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
51000   Case "PDFGENERALCOMPATIBILITY": ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
51010   Case "PDFGENERALDEFAULT": ini.SaveKey CStr(.PDFGeneralDefault), "PDFGeneralDefault"
51020   Case "PDFGENERALOVERPRINT": ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
51030   Case "PDFGENERALRESOLUTION": ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
51040   Case "PDFHIGHENCRYPTION": ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
51050   Case "PDFLOWENCRYPTION": ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
51060   Case "PDFOPTIMIZE": ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
51070   Case "PDFOWNERPASS": ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
51080   Case "PDFOWNERPASSWORDSTRING": ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
51090   Case "PDFSIGNINGMULTISIGNATURE": ini.SaveKey CStr(Abs(.PDFSigningMultiSignature)), "PDFSigningMultiSignature"
51100   Case "PDFSIGNINGPFXFILE": ini.SaveKey CStr(.PDFSigningPFXFile), "PDFSigningPFXFile"
51110   Case "PDFSIGNINGPFXFILEPASSWORD": ini.SaveKey CStr(.PDFSigningPFXFilePassword), "PDFSigningPFXFilePassword"
51120   Case "PDFSIGNINGSIGNATURECONTACT": ini.SaveKey CStr(.PDFSigningSignatureContact), "PDFSigningSignatureContact"
51130   Case "PDFSIGNINGSIGNATURELEFTX": ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), "PDFSigningSignatureLeftX"
51140   Case "PDFSIGNINGSIGNATURELEFTY": ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), "PDFSigningSignatureLeftY"
51150   Case "PDFSIGNINGSIGNATURELOCATION": ini.SaveKey CStr(.PDFSigningSignatureLocation), "PDFSigningSignatureLocation"
51160   Case "PDFSIGNINGSIGNATUREONPAGE": ini.SaveKey CStr(.PDFSigningSignatureOnPage), "PDFSigningSignatureOnPage"
51170   Case "PDFSIGNINGSIGNATUREREASON": ini.SaveKey CStr(.PDFSigningSignatureReason), "PDFSigningSignatureReason"
51180   Case "PDFSIGNINGSIGNATURERIGHTX": ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), "PDFSigningSignatureRightX"
51190   Case "PDFSIGNINGSIGNATURERIGHTY": ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), "PDFSigningSignatureRightY"
51200   Case "PDFSIGNINGSIGNATUREVISIBLE": ini.SaveKey CStr(Abs(.PDFSigningSignatureVisible)), "PDFSigningSignatureVisible"
51210   Case "PDFSIGNINGSIGNPDF": ini.SaveKey CStr(Abs(.PDFSigningSignPDF)), "PDFSigningSignPDF"
51220   Case "PDFUPDATEMETADATA": ini.SaveKey CStr(.PDFUpdateMetadata), "PDFUpdateMetadata"
51230   Case "PDFUSERPASS": ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
51240   Case "PDFUSERPASSWORDSTRING": ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
51250   Case "PDFUSESECURITY": ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
51260   Case "PNGCOLORSCOUNT": ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
51270   Case "PNGRESOLUTION": ini.SaveKey CStr(.PNGResolution), "PNGResolution"
51280   Case "PRINTAFTERSAVING": ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
51290   Case "PRINTAFTERSAVINGBITSPERPIXEL": ini.SaveKey CStr(.PrintAfterSavingBitsPerPixel), "PrintAfterSavingBitsPerPixel"
51300   Case "PRINTAFTERSAVINGDUPLEX": ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
51310   Case "PRINTAFTERSAVINGMAXRESOLUTION": ini.SaveKey CStr(.PrintAfterSavingMaxResolution), "PrintAfterSavingMaxResolution"
51320   Case "PRINTAFTERSAVINGMAXRESOLUTIONENABLED": ini.SaveKey CStr(Abs(.PrintAfterSavingMaxResolutionEnabled)), "PrintAfterSavingMaxResolutionEnabled"
51330   Case "PRINTAFTERSAVINGNOCANCEL": ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
51340   Case "PRINTAFTERSAVINGPRINTER": ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
51350   Case "PRINTAFTERSAVINGQUERYUSER": ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
51360   Case "PRINTAFTERSAVINGTUMBLE": ini.SaveKey CStr(.PrintAfterSavingTumble), "PrintAfterSavingTumble"
51370   Case "PRINTERSTOP": ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
51380   Case "PRINTERTEMPPATH": ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
51390   Case "PROCESSPRIORITY": ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
51400   Case "PROGRAMFONT": ini.SaveKey CStr(.ProgramFont), "ProgramFont"
51410   Case "PROGRAMFONTCHARSET": ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
51420   Case "PROGRAMFONTSIZE": ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
51430   Case "PSDCOLORSCOUNT": ini.SaveKey CStr(.PSDColorsCount), "PSDColorsCount"
51440   Case "PSDRESOLUTION": ini.SaveKey CStr(.PSDResolution), "PSDResolution"
51450   Case "PSLANGUAGELEVEL": ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
51460   Case "RAWCOLORSCOUNT": ini.SaveKey CStr(.RAWColorsCount), "RAWColorsCount"
51470   Case "RAWRESOLUTION": ini.SaveKey CStr(.RAWResolution), "RAWResolution"
51480   Case "REMOVEALLKNOWNFILEEXTENSIONS": ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
51490   Case "REMOVESPACES": ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
51500   Case "RUNPROGRAMAFTERSAVING": ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
51510   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
51520   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
51530   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51540   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51550   Case "RUNPROGRAMBEFORESAVING": ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51560   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51570   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51580   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51590   Case "SAVEFILENAME": ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51600   Case "SENDEMAILAFTERAUTOSAVING": ini.SaveKey CStr(Abs(.SendEmailAfterAutoSaving)), "SendEmailAfterAutoSaving"
51610   Case "SENDMAILMETHOD": ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51620   Case "SHOWANIMATION": ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51630   Case "STAMPFONTCOLOR": ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51640   Case "STAMPFONTNAME": ini.SaveKey CStr(.StampFontname), "StampFontname"
51650   Case "STAMPFONTSIZE": ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51660   Case "STAMPOUTLINEFONTTHICKNESS": ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51670   Case "STAMPSTRING": ini.SaveKey CStr(.StampString), "StampString"
51680   Case "STAMPUSEOUTLINEFONT": ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51690   Case "STANDARDAUTHOR": ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51700   Case "STANDARDCREATIONDATE": ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51710   Case "STANDARDDATEFORMAT": ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51720   Case "STANDARDKEYWORDS": ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51730   Case "STANDARDMAILDOMAIN": ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51740   Case "STANDARDMODIFYDATE": ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51750   Case "STANDARDSAVEFORMAT": ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51760   Case "STANDARDSUBJECT": ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51770   Case "STANDARDTITLE": ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51780   Case "STARTSTANDARDPROGRAM": ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51790   Case "SVGRESOLUTION": ini.SaveKey CStr(.SVGResolution), "SVGResolution"
51800   Case "TIFFCOLORSCOUNT": ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51810   Case "TIFFRESOLUTION": ini.SaveKey CStr(.TIFFResolution), "TIFFResolution"
51820   Case "TOOLBARS": ini.SaveKey CStr(.Toolbars), "Toolbars"
51830   Case "UPDATEINTERVAL": ini.SaveKey CStr(.UpdateInterval), "UpdateInterval"
51840   Case "USEAUTOSAVE": ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51850   Case "USEAUTOSAVEDIRECTORY": ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51860   Case "USECREATIONDATENOW": ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51870   Case "USECUSTOMPAPERSIZE": ini.SaveKey CStr(.UseCustomPaperSize), "UseCustomPaperSize"
51880   Case "USEFIXPAPERSIZE": ini.SaveKey CStr(Abs(.UseFixPapersize)), "UseFixPapersize"
51890   Case "USESTANDARDAUTHOR": ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51900   End Select
51910  End With
51920  Set ini = Nothing
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
50120   ini.SaveKey CStr(Abs(.AllowSpecialGSCharsInFilenames)), "AllowSpecialGSCharsInFilenames"
50130   ini.SaveKey CStr(.AutosaveDirectory), "AutosaveDirectory"
50140   ini.SaveKey CStr(.AutosaveFilename), "AutosaveFilename"
50150   ini.SaveKey CStr(.AutosaveFormat), "AutosaveFormat"
50160   ini.SaveKey CStr(Abs(.AutosaveStartStandardProgram)), "AutosaveStartStandardProgram"
50170   ini.SaveKey CStr(.BMPColorscount), "BMPColorscount"
50180   ini.SaveKey CStr(.BMPResolution), "BMPResolution"
50190   ini.SaveKey CStr(Abs(.ClientComputerResolveIPAddress)), "ClientComputerResolveIPAddress"
50200   ini.SaveKey CStr(.Counter), "Counter"
50210   ini.SaveKey Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), "DeviceHeightPoints"
50220   ini.SaveKey Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), "DeviceWidthPoints"
50230   ini.SaveKey CStr(.DirectoryGhostscriptBinaries), "DirectoryGhostscriptBinaries"
50240   ini.SaveKey CStr(.DirectoryGhostscriptFonts), "DirectoryGhostscriptFonts"
50250   ini.SaveKey CStr(.DirectoryGhostscriptLibraries), "DirectoryGhostscriptLibraries"
50260   ini.SaveKey CStr(.DirectoryGhostscriptResource), "DirectoryGhostscriptResource"
50270   ini.SaveKey CStr(Abs(.DisableEmail)), "DisableEmail"
50280   ini.SaveKey CStr(Abs(.DontUseDocumentSettings)), "DontUseDocumentSettings"
50290   ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
50300   ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
50310   ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
50320   ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50330   ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50340   ini.SaveKey CStr(.JPEGResolution), "JPEGResolution"
50350   ini.SaveKey CStr(.Language), "Language"
50360   ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50370   ini.SaveKey CStr(.LastUpdateCheck), "LastUpdateCheck"
50380   ini.SaveKey CStr(Abs(.Logging)), "Logging"
50390   ini.SaveKey CStr(.LogLines), "LogLines"
50400   ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50410   ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
50420   ini.SaveKey CStr(Abs(.NoPSCheck)), "NoPSCheck"
50430   ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50440   ini.SaveKey CStr(.OptionsDesign), "OptionsDesign"
50450   ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50460   ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50470   ini.SaveKey CStr(.Papersize), "Papersize"
50480   ini.SaveKey CStr(.PCLColorsCount), "PCLColorsCount"
50490   ini.SaveKey CStr(.PCLResolution), "PCLResolution"
50500   ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50510   ini.SaveKey CStr(.PCXResolution), "PCXResolution"
50520   ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50530   ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50540   ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50550   ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50560   ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50570   ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50580   ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50590   ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50600   ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50610   ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50620   ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50630   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGHighFactor"
50640   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGLowFactor"
50650   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGManualFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGManualFactor"
50660   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMaximumFactor"
50670   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMediumFactor"
50680   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMinimumFactor"
50690   ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50700   ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50710   ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50720   ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50730   ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50740   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGHighFactor"
50750   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGLowFactor"
50760   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGManualFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGManualFactor"
50770   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMaximumFactor"
50780   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMediumFactor"
50790   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMinimumFactor"
50800   ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50810   ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50820   ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50830   ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50840   ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50850   ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50860   ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50870   ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50880   ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50890   ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50900   ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50910   ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50920   ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50930   ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50940   ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50950   ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50960   ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50970   ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50980   ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50990   ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
51000   ini.SaveKey CStr(.PDFGeneralDefault), "PDFGeneralDefault"
51010   ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
51020   ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
51030   ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
51040   ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
51050   ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
51060   ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
51070   ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
51080   ini.SaveKey CStr(Abs(.PDFSigningMultiSignature)), "PDFSigningMultiSignature"
51090   ini.SaveKey CStr(.PDFSigningPFXFile), "PDFSigningPFXFile"
51100   ini.SaveKey CStr(.PDFSigningPFXFilePassword), "PDFSigningPFXFilePassword"
51110   ini.SaveKey CStr(.PDFSigningSignatureContact), "PDFSigningSignatureContact"
51120   ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), "PDFSigningSignatureLeftX"
51130   ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), "PDFSigningSignatureLeftY"
51140   ini.SaveKey CStr(.PDFSigningSignatureLocation), "PDFSigningSignatureLocation"
51150   ini.SaveKey CStr(.PDFSigningSignatureOnPage), "PDFSigningSignatureOnPage"
51160   ini.SaveKey CStr(.PDFSigningSignatureReason), "PDFSigningSignatureReason"
51170   ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), "PDFSigningSignatureRightX"
51180   ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), "PDFSigningSignatureRightY"
51190   ini.SaveKey CStr(Abs(.PDFSigningSignatureVisible)), "PDFSigningSignatureVisible"
51200   ini.SaveKey CStr(Abs(.PDFSigningSignPDF)), "PDFSigningSignPDF"
51210   ini.SaveKey CStr(.PDFUpdateMetadata), "PDFUpdateMetadata"
51220   ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
51230   ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
51240   ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
51250   ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
51260   ini.SaveKey CStr(.PNGResolution), "PNGResolution"
51270   ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
51280   ini.SaveKey CStr(.PrintAfterSavingBitsPerPixel), "PrintAfterSavingBitsPerPixel"
51290   ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
51300   ini.SaveKey CStr(.PrintAfterSavingMaxResolution), "PrintAfterSavingMaxResolution"
51310   ini.SaveKey CStr(Abs(.PrintAfterSavingMaxResolutionEnabled)), "PrintAfterSavingMaxResolutionEnabled"
51320   ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
51330   ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
51340   ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
51350   ini.SaveKey CStr(.PrintAfterSavingTumble), "PrintAfterSavingTumble"
51360   ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
51370   ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
51380   ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
51390   ini.SaveKey CStr(.ProgramFont), "ProgramFont"
51400   ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
51410   ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
51420   ini.SaveKey CStr(.PSDColorsCount), "PSDColorsCount"
51430   ini.SaveKey CStr(.PSDResolution), "PSDResolution"
51440   ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
51450   ini.SaveKey CStr(.RAWColorsCount), "RAWColorsCount"
51460   ini.SaveKey CStr(.RAWResolution), "RAWResolution"
51470   ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
51480   ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
51490   ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
51500   ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
51510   ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
51520   ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51530   ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51540   ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51550   ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51560   ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51570   ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51580   ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51590   ini.SaveKey CStr(Abs(.SendEmailAfterAutoSaving)), "SendEmailAfterAutoSaving"
51600   ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51610   ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51620   ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51630   ini.SaveKey CStr(.StampFontname), "StampFontname"
51640   ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51650   ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51660   ini.SaveKey CStr(.StampString), "StampString"
51670   ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51680   ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51690   ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51700   ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51710   ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51720   ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51730   ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51740   ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51750   ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51760   ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51770   ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51780   ini.SaveKey CStr(.SVGResolution), "SVGResolution"
51790   ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51800   ini.SaveKey CStr(.TIFFResolution), "TIFFResolution"
51810   ini.SaveKey CStr(.Toolbars), "Toolbars"
51820   ini.SaveKey CStr(.UpdateInterval), "UpdateInterval"
51830   ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51840   ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51850   ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51860   ini.SaveKey CStr(.UseCustomPaperSize), "UseCustomPaperSize"
51870   ini.SaveKey CStr(Abs(.UseFixPapersize)), "UseFixPapersize"
51880   ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51890  End With
51900  Set ini = Nothing
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
50060   reg.SubKey = "Ghostscript"
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
50420   reg.SubKey = "Printing"
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
53010   reg.SubKey = "Printing\Formats\Bitmap\Colors"
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
55120   tStr = reg.GetRegistryValue("SVGResolution")
55130   If IsNumeric(tStr) Then
55140     If CLng(tStr) >= 1 Then
55150       .SVGResolution = CLng(tStr)
55160      Else
55170       If UseStandard Then
55180        .SVGResolution = 72
55190       End If
55200     End If
55210    Else
55220     If UseStandard Then
55230      .SVGResolution = 72
55240     End If
55250   End If
55260   tStr = reg.GetRegistryValue("TIFFColorscount")
55270   If IsNumeric(tStr) Then
55280     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
55290       .TIFFColorscount = CLng(tStr)
55300      Else
55310       If UseStandard Then
55320        .TIFFColorscount = 0
55330       End If
55340     End If
55350    Else
55360     If UseStandard Then
55370      .TIFFColorscount = 0
55380     End If
55390   End If
55400   tStr = reg.GetRegistryValue("TIFFResolution")
55410   If IsNumeric(tStr) Then
55420     If CLng(tStr) >= 1 Then
55430       .TIFFResolution = CLng(tStr)
55440      Else
55450       If UseStandard Then
55460        .TIFFResolution = 150
55470       End If
55480     End If
55490    Else
55500     If UseStandard Then
55510      .TIFFResolution = 150
55520     End If
55530   End If
55540   reg.SubKey = "Printing\Formats\PDF\Colors"
55550   tStr = reg.GetRegistryValue("PDFColorsCMYKToRGB")
55560   If IsNumeric(tStr) Then
55570     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55580       .PDFColorsCMYKToRGB = CLng(tStr)
55590      Else
55600       If UseStandard Then
55610        .PDFColorsCMYKToRGB = 0
55620       End If
55630     End If
55640    Else
55650     If UseStandard Then
55660      .PDFColorsCMYKToRGB = 0
55670     End If
55680   End If
55690   tStr = reg.GetRegistryValue("PDFColorsColorModel")
55700   If IsNumeric(tStr) Then
55710     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
55720       .PDFColorsColorModel = CLng(tStr)
55730      Else
55740       If UseStandard Then
55750        .PDFColorsColorModel = 1
55760       End If
55770     End If
55780    Else
55790     If UseStandard Then
55800      .PDFColorsColorModel = 1
55810     End If
55820   End If
55830   tStr = reg.GetRegistryValue("PDFColorsPreserveHalftone")
55840   If IsNumeric(tStr) Then
55850     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55860       .PDFColorsPreserveHalftone = CLng(tStr)
55870      Else
55880       If UseStandard Then
55890        .PDFColorsPreserveHalftone = 0
55900       End If
55910     End If
55920    Else
55930     If UseStandard Then
55940      .PDFColorsPreserveHalftone = 0
55950     End If
55960   End If
55970   tStr = reg.GetRegistryValue("PDFColorsPreserveOverprint")
55980   If IsNumeric(tStr) Then
55990     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56000       .PDFColorsPreserveOverprint = CLng(tStr)
56010      Else
56020       If UseStandard Then
56030        .PDFColorsPreserveOverprint = 1
56040       End If
56050     End If
56060    Else
56070     If UseStandard Then
56080      .PDFColorsPreserveOverprint = 1
56090     End If
56100   End If
56110   tStr = reg.GetRegistryValue("PDFColorsPreserveTransfer")
56120   If IsNumeric(tStr) Then
56130     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56140       .PDFColorsPreserveTransfer = CLng(tStr)
56150      Else
56160       If UseStandard Then
56170        .PDFColorsPreserveTransfer = 1
56180       End If
56190     End If
56200    Else
56210     If UseStandard Then
56220      .PDFColorsPreserveTransfer = 1
56230     End If
56240   End If
56250   reg.SubKey = "Printing\Formats\PDF\Compression"
56260   tStr = reg.GetRegistryValue("PDFCompressionColorCompression")
56270   If IsNumeric(tStr) Then
56280     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
56290       .PDFCompressionColorCompression = CLng(tStr)
56300      Else
56310       If UseStandard Then
56320        .PDFCompressionColorCompression = 1
56330       End If
56340     End If
56350    Else
56360     If UseStandard Then
56370      .PDFCompressionColorCompression = 1
56380     End If
56390   End If
56400   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionChoice")
56410   If IsNumeric(tStr) Then
56420     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
56430       .PDFCompressionColorCompressionChoice = CLng(tStr)
56440      Else
56450       If UseStandard Then
56460        .PDFCompressionColorCompressionChoice = 0
56470       End If
56480     End If
56490    Else
56500     If UseStandard Then
56510      .PDFCompressionColorCompressionChoice = 0
56520     End If
56530   End If
56540   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGHighFactor")
56550   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56560     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56570       .PDFCompressionColorCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56580      Else
56590       If UseStandard Then
56600        .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56610       End If
56620     End If
56630    Else
56640     If UseStandard Then
56650      .PDFCompressionColorCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
56660     End If
56670   End If
56680   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGLowFactor")
56690   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56700     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56710       .PDFCompressionColorCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56720      Else
56730       If UseStandard Then
56740        .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56750       End If
56760     End If
56770    Else
56780     If UseStandard Then
56790      .PDFCompressionColorCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
56800     End If
56810   End If
56820   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGManualFactor")
56830   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56840     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56850       .PDFCompressionColorCompressionJPEGManualFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56860      Else
56870       If UseStandard Then
56880        .PDFCompressionColorCompressionJPEGManualFactor = Replace$("3.0", ".", GetDecimalChar)
56890       End If
56900     End If
56910    Else
56920     If UseStandard Then
56930      .PDFCompressionColorCompressionJPEGManualFactor = Replace$("3.0", ".", GetDecimalChar)
56940     End If
56950   End If
56960   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMaximumFactor")
56970   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56980     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56990       .PDFCompressionColorCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57000      Else
57010       If UseStandard Then
57020        .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57030       End If
57040     End If
57050    Else
57060     If UseStandard Then
57070      .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57080     End If
57090   End If
57100   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMediumFactor")
57110   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57120     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57130       .PDFCompressionColorCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57140      Else
57150       If UseStandard Then
57160        .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57170       End If
57180     End If
57190    Else
57200     If UseStandard Then
57210      .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57220     End If
57230   End If
57240   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMinimumFactor")
57250   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57260     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57270       .PDFCompressionColorCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57280      Else
57290       If UseStandard Then
57300        .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57310       End If
57320     End If
57330    Else
57340     If UseStandard Then
57350      .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57360     End If
57370   End If
57380   tStr = reg.GetRegistryValue("PDFCompressionColorResample")
57390   If IsNumeric(tStr) Then
57400     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57410       .PDFCompressionColorResample = CLng(tStr)
57420      Else
57430       If UseStandard Then
57440        .PDFCompressionColorResample = 0
57450       End If
57460     End If
57470    Else
57480     If UseStandard Then
57490      .PDFCompressionColorResample = 0
57500     End If
57510   End If
57520   tStr = reg.GetRegistryValue("PDFCompressionColorResampleChoice")
57530   If IsNumeric(tStr) Then
57540     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
57550       .PDFCompressionColorResampleChoice = CLng(tStr)
57560      Else
57570       If UseStandard Then
57580        .PDFCompressionColorResampleChoice = 0
57590       End If
57600     End If
57610    Else
57620     If UseStandard Then
57630      .PDFCompressionColorResampleChoice = 0
57640     End If
57650   End If
57660   tStr = reg.GetRegistryValue("PDFCompressionColorResolution")
57670   If IsNumeric(tStr) Then
57680     If CLng(tStr) >= 0 Then
57690       .PDFCompressionColorResolution = CLng(tStr)
57700      Else
57710       If UseStandard Then
57720        .PDFCompressionColorResolution = 300
57730       End If
57740     End If
57750    Else
57760     If UseStandard Then
57770      .PDFCompressionColorResolution = 300
57780     End If
57790   End If
57800   tStr = reg.GetRegistryValue("PDFCompressionGreyCompression")
57810   If IsNumeric(tStr) Then
57820     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57830       .PDFCompressionGreyCompression = CLng(tStr)
57840      Else
57850       If UseStandard Then
57860        .PDFCompressionGreyCompression = 1
57870       End If
57880     End If
57890    Else
57900     If UseStandard Then
57910      .PDFCompressionGreyCompression = 1
57920     End If
57930   End If
57940   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionChoice")
57950   If IsNumeric(tStr) Then
57960     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
57970       .PDFCompressionGreyCompressionChoice = CLng(tStr)
57980      Else
57990       If UseStandard Then
58000        .PDFCompressionGreyCompressionChoice = 0
58010       End If
58020     End If
58030    Else
58040     If UseStandard Then
58050      .PDFCompressionGreyCompressionChoice = 0
58060     End If
58070   End If
58080   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGHighFactor")
58090   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58100     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58110       .PDFCompressionGreyCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58120      Else
58130       If UseStandard Then
58140        .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
58150       End If
58160     End If
58170    Else
58180     If UseStandard Then
58190      .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
58200     End If
58210   End If
58220   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGLowFactor")
58230   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58240     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58250       .PDFCompressionGreyCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58260      Else
58270       If UseStandard Then
58280        .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58290       End If
58300     End If
58310    Else
58320     If UseStandard Then
58330      .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58340     End If
58350   End If
58360   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGManualFactor")
58370   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58380     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58390       .PDFCompressionGreyCompressionJPEGManualFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58400      Else
58410       If UseStandard Then
58420        .PDFCompressionGreyCompressionJPEGManualFactor = Replace$("3.0", ".", GetDecimalChar)
58430       End If
58440     End If
58450    Else
58460     If UseStandard Then
58470      .PDFCompressionGreyCompressionJPEGManualFactor = Replace$("3.0", ".", GetDecimalChar)
58480     End If
58490   End If
58500   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMaximumFactor")
58510   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58520     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58530       .PDFCompressionGreyCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58540      Else
58550       If UseStandard Then
58560        .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
58570       End If
58580     End If
58590    Else
58600     If UseStandard Then
58610      .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
58620     End If
58630   End If
58640   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMediumFactor")
58650   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58660     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58670       .PDFCompressionGreyCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58680      Else
58690       If UseStandard Then
58700        .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58710       End If
58720     End If
58730    Else
58740     If UseStandard Then
58750      .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58760     End If
58770   End If
58780   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMinimumFactor")
58790   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58800     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58810       .PDFCompressionGreyCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58820      Else
58830       If UseStandard Then
58840        .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58850       End If
58860     End If
58870    Else
58880     If UseStandard Then
58890      .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58900     End If
58910   End If
58920   tStr = reg.GetRegistryValue("PDFCompressionGreyResample")
58930   If IsNumeric(tStr) Then
58940     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58950       .PDFCompressionGreyResample = CLng(tStr)
58960      Else
58970       If UseStandard Then
58980        .PDFCompressionGreyResample = 0
58990       End If
59000     End If
59010    Else
59020     If UseStandard Then
59030      .PDFCompressionGreyResample = 0
59040     End If
59050   End If
59060   tStr = reg.GetRegistryValue("PDFCompressionGreyResampleChoice")
59070   If IsNumeric(tStr) Then
59080     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
59090       .PDFCompressionGreyResampleChoice = CLng(tStr)
59100      Else
59110       If UseStandard Then
59120        .PDFCompressionGreyResampleChoice = 0
59130       End If
59140     End If
59150    Else
59160     If UseStandard Then
59170      .PDFCompressionGreyResampleChoice = 0
59180     End If
59190   End If
59200   tStr = reg.GetRegistryValue("PDFCompressionGreyResolution")
59210   If IsNumeric(tStr) Then
59220     If CLng(tStr) >= 0 Then
59230       .PDFCompressionGreyResolution = CLng(tStr)
59240      Else
59250       If UseStandard Then
59260        .PDFCompressionGreyResolution = 300
59270       End If
59280     End If
59290    Else
59300     If UseStandard Then
59310      .PDFCompressionGreyResolution = 300
59320     End If
59330   End If
59340   tStr = reg.GetRegistryValue("PDFCompressionMonoCompression")
59350   If IsNumeric(tStr) Then
59360     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59370       .PDFCompressionMonoCompression = CLng(tStr)
59380      Else
59390       If UseStandard Then
59400        .PDFCompressionMonoCompression = 1
59410       End If
59420     End If
59430    Else
59440     If UseStandard Then
59450      .PDFCompressionMonoCompression = 1
59460     End If
59470   End If
59480   tStr = reg.GetRegistryValue("PDFCompressionMonoCompressionChoice")
59490   If IsNumeric(tStr) Then
59500     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
59510       .PDFCompressionMonoCompressionChoice = CLng(tStr)
59520      Else
59530       If UseStandard Then
59540        .PDFCompressionMonoCompressionChoice = 0
59550       End If
59560     End If
59570    Else
59580     If UseStandard Then
59590      .PDFCompressionMonoCompressionChoice = 0
59600     End If
59610   End If
59620   tStr = reg.GetRegistryValue("PDFCompressionMonoResample")
59630   If IsNumeric(tStr) Then
59640     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59650       .PDFCompressionMonoResample = CLng(tStr)
59660      Else
59670       If UseStandard Then
59680        .PDFCompressionMonoResample = 0
59690       End If
59700     End If
59710    Else
59720     If UseStandard Then
59730      .PDFCompressionMonoResample = 0
59740     End If
59750   End If
59760   tStr = reg.GetRegistryValue("PDFCompressionMonoResampleChoice")
59770   If IsNumeric(tStr) Then
59780     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
59790       .PDFCompressionMonoResampleChoice = CLng(tStr)
59800      Else
59810       If UseStandard Then
59820        .PDFCompressionMonoResampleChoice = 0
59830       End If
59840     End If
59850    Else
59860     If UseStandard Then
59870      .PDFCompressionMonoResampleChoice = 0
59880     End If
59890   End If
59900   tStr = reg.GetRegistryValue("PDFCompressionMonoResolution")
59910   If IsNumeric(tStr) Then
59920     If CLng(tStr) >= 0 Then
59930       .PDFCompressionMonoResolution = CLng(tStr)
59940      Else
59950       If UseStandard Then
59960        .PDFCompressionMonoResolution = 1200
59970       End If
59980     End If
59990    Else
60000     If UseStandard Then
60010      .PDFCompressionMonoResolution = 1200
60020     End If
60030   End If
60040   tStr = reg.GetRegistryValue("PDFCompressionTextCompression")
60050   If IsNumeric(tStr) Then
60060     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60070       .PDFCompressionTextCompression = CLng(tStr)
60080      Else
60090       If UseStandard Then
60100        .PDFCompressionTextCompression = 1
60110       End If
60120     End If
60130    Else
60140     If UseStandard Then
60150      .PDFCompressionTextCompression = 1
60160     End If
60170   End If
60180   reg.SubKey = "Printing\Formats\PDF\Fonts"
60190   tStr = reg.GetRegistryValue("PDFFontsEmbedAll")
60200   If IsNumeric(tStr) Then
60210     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60220       .PDFFontsEmbedAll = CLng(tStr)
60230      Else
60240       If UseStandard Then
60250        .PDFFontsEmbedAll = 1
60260       End If
60270     End If
60280    Else
60290     If UseStandard Then
60300      .PDFFontsEmbedAll = 1
60310     End If
60320   End If
60330   tStr = reg.GetRegistryValue("PDFFontsSubSetFonts")
60340   If IsNumeric(tStr) Then
60350     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60360       .PDFFontsSubSetFonts = CLng(tStr)
60370      Else
60380       If UseStandard Then
60390        .PDFFontsSubSetFonts = 1
60400       End If
60410     End If
60420    Else
60430     If UseStandard Then
60440      .PDFFontsSubSetFonts = 1
60450     End If
60460   End If
60470   tStr = reg.GetRegistryValue("PDFFontsSubSetFontsPercent")
60480   If IsNumeric(tStr) Then
60490     If CLng(tStr) >= 0 Then
60500       .PDFFontsSubSetFontsPercent = CLng(tStr)
60510      Else
60520       If UseStandard Then
60530        .PDFFontsSubSetFontsPercent = 100
60540       End If
60550     End If
60560    Else
60570     If UseStandard Then
60580      .PDFFontsSubSetFontsPercent = 100
60590     End If
60600   End If
60610   reg.SubKey = "Printing\Formats\PDF\General"
60620   tStr = reg.GetRegistryValue("PDFGeneralASCII85")
60630   If IsNumeric(tStr) Then
60640     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60650       .PDFGeneralASCII85 = CLng(tStr)
60660      Else
60670       If UseStandard Then
60680        .PDFGeneralASCII85 = 0
60690       End If
60700     End If
60710    Else
60720     If UseStandard Then
60730      .PDFGeneralASCII85 = 0
60740     End If
60750   End If
60760   tStr = reg.GetRegistryValue("PDFGeneralAutorotate")
60770   If IsNumeric(tStr) Then
60780     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
60790       .PDFGeneralAutorotate = CLng(tStr)
60800      Else
60810       If UseStandard Then
60820        .PDFGeneralAutorotate = 2
60830       End If
60840     End If
60850    Else
60860     If UseStandard Then
60870      .PDFGeneralAutorotate = 2
60880     End If
60890   End If
60900   tStr = reg.GetRegistryValue("PDFGeneralCompatibility")
60910   If IsNumeric(tStr) Then
60920     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
60930       .PDFGeneralCompatibility = CLng(tStr)
60940      Else
60950       If UseStandard Then
60960        .PDFGeneralCompatibility = 2
60970       End If
60980     End If
60990    Else
61000     If UseStandard Then
61010      .PDFGeneralCompatibility = 2
61020     End If
61030   End If
61040   tStr = reg.GetRegistryValue("PDFGeneralDefault")
61050   If IsNumeric(tStr) Then
61060     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
61070       .PDFGeneralDefault = CLng(tStr)
61080      Else
61090       If UseStandard Then
61100        .PDFGeneralDefault = 0
61110       End If
61120     End If
61130    Else
61140     If UseStandard Then
61150      .PDFGeneralDefault = 0
61160     End If
61170   End If
61180   tStr = reg.GetRegistryValue("PDFGeneralOverprint")
61190   If IsNumeric(tStr) Then
61200     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
61210       .PDFGeneralOverprint = CLng(tStr)
61220      Else
61230       If UseStandard Then
61240        .PDFGeneralOverprint = 0
61250       End If
61260     End If
61270    Else
61280     If UseStandard Then
61290      .PDFGeneralOverprint = 0
61300     End If
61310   End If
61320   tStr = reg.GetRegistryValue("PDFGeneralResolution")
61330   If IsNumeric(tStr) Then
61340     If CLng(tStr) >= 0 Then
61350       .PDFGeneralResolution = CLng(tStr)
61360      Else
61370       If UseStandard Then
61380        .PDFGeneralResolution = 600
61390       End If
61400     End If
61410    Else
61420     If UseStandard Then
61430      .PDFGeneralResolution = 600
61440     End If
61450   End If
61460   tStr = reg.GetRegistryValue("PDFOptimize")
61470   If IsNumeric(tStr) Then
61480     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61490       .PDFOptimize = CLng(tStr)
61500      Else
61510       If UseStandard Then
61520        .PDFOptimize = 0
61530       End If
61540     End If
61550    Else
61560     If UseStandard Then
61570      .PDFOptimize = 0
61580     End If
61590   End If
61600   tStr = reg.GetRegistryValue("PDFUpdateMetadata")
61610   If IsNumeric(tStr) Then
61620     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
61630       .PDFUpdateMetadata = CLng(tStr)
61640      Else
61650       If UseStandard Then
61660        .PDFUpdateMetadata = 1
61670       End If
61680     End If
61690    Else
61700     If UseStandard Then
61710      .PDFUpdateMetadata = 1
61720     End If
61730   End If
61740   reg.SubKey = "Printing\Formats\PDF\Security"
61750   tStr = reg.GetRegistryValue("PDFAllowAssembly")
61760   If IsNumeric(tStr) Then
61770     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61780       .PDFAllowAssembly = CLng(tStr)
61790      Else
61800       If UseStandard Then
61810        .PDFAllowAssembly = 0
61820       End If
61830     End If
61840    Else
61850     If UseStandard Then
61860      .PDFAllowAssembly = 0
61870     End If
61880   End If
61890   tStr = reg.GetRegistryValue("PDFAllowDegradedPrinting")
61900   If IsNumeric(tStr) Then
61910     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61920       .PDFAllowDegradedPrinting = CLng(tStr)
61930      Else
61940       If UseStandard Then
61950        .PDFAllowDegradedPrinting = 0
61960       End If
61970     End If
61980    Else
61990     If UseStandard Then
62000      .PDFAllowDegradedPrinting = 0
62010     End If
62020   End If
62030   tStr = reg.GetRegistryValue("PDFAllowFillIn")
62040   If IsNumeric(tStr) Then
62050     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62060       .PDFAllowFillIn = CLng(tStr)
62070      Else
62080       If UseStandard Then
62090        .PDFAllowFillIn = 0
62100       End If
62110     End If
62120    Else
62130     If UseStandard Then
62140      .PDFAllowFillIn = 0
62150     End If
62160   End If
62170   tStr = reg.GetRegistryValue("PDFAllowScreenReaders")
62180   If IsNumeric(tStr) Then
62190     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62200       .PDFAllowScreenReaders = CLng(tStr)
62210      Else
62220       If UseStandard Then
62230        .PDFAllowScreenReaders = 0
62240       End If
62250     End If
62260    Else
62270     If UseStandard Then
62280      .PDFAllowScreenReaders = 0
62290     End If
62300   End If
62310   tStr = reg.GetRegistryValue("PDFDisallowCopy")
62320   If IsNumeric(tStr) Then
62330     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62340       .PDFDisallowCopy = CLng(tStr)
62350      Else
62360       If UseStandard Then
62370        .PDFDisallowCopy = 1
62380       End If
62390     End If
62400    Else
62410     If UseStandard Then
62420      .PDFDisallowCopy = 1
62430     End If
62440   End If
62450   tStr = reg.GetRegistryValue("PDFDisallowModifyAnnotations")
62460   If IsNumeric(tStr) Then
62470     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62480       .PDFDisallowModifyAnnotations = CLng(tStr)
62490      Else
62500       If UseStandard Then
62510        .PDFDisallowModifyAnnotations = 0
62520       End If
62530     End If
62540    Else
62550     If UseStandard Then
62560      .PDFDisallowModifyAnnotations = 0
62570     End If
62580   End If
62590   tStr = reg.GetRegistryValue("PDFDisallowModifyContents")
62600   If IsNumeric(tStr) Then
62610     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62620       .PDFDisallowModifyContents = CLng(tStr)
62630      Else
62640       If UseStandard Then
62650        .PDFDisallowModifyContents = 0
62660       End If
62670     End If
62680    Else
62690     If UseStandard Then
62700      .PDFDisallowModifyContents = 0
62710     End If
62720   End If
62730   tStr = reg.GetRegistryValue("PDFDisallowPrinting")
62740   If IsNumeric(tStr) Then
62750     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62760       .PDFDisallowPrinting = CLng(tStr)
62770      Else
62780       If UseStandard Then
62790        .PDFDisallowPrinting = 0
62800       End If
62810     End If
62820    Else
62830     If UseStandard Then
62840      .PDFDisallowPrinting = 0
62850     End If
62860   End If
62870   tStr = reg.GetRegistryValue("PDFEncryptor")
62880   If IsNumeric(tStr) Then
62890     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
62900       .PDFEncryptor = CLng(tStr)
62910      Else
62920       If UseStandard Then
62930        .PDFEncryptor = 0
62940       End If
62950     End If
62960    Else
62970     If UseStandard Then
62980      .PDFEncryptor = 0
62990     End If
63000   End If
63010   tStr = reg.GetRegistryValue("PDFHighEncryption")
63020   If IsNumeric(tStr) Then
63030     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63040       .PDFHighEncryption = CLng(tStr)
63050      Else
63060       If UseStandard Then
63070        .PDFHighEncryption = 0
63080       End If
63090     End If
63100    Else
63110     If UseStandard Then
63120      .PDFHighEncryption = 0
63130     End If
63140   End If
63150   tStr = reg.GetRegistryValue("PDFLowEncryption")
63160   If IsNumeric(tStr) Then
63170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63180       .PDFLowEncryption = CLng(tStr)
63190      Else
63200       If UseStandard Then
63210        .PDFLowEncryption = 1
63220       End If
63230     End If
63240    Else
63250     If UseStandard Then
63260      .PDFLowEncryption = 1
63270     End If
63280   End If
63290   tStr = reg.GetRegistryValue("PDFOwnerPass")
63300   If IsNumeric(tStr) Then
63310     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63320       .PDFOwnerPass = CLng(tStr)
63330      Else
63340       If UseStandard Then
63350        .PDFOwnerPass = 0
63360       End If
63370     End If
63380    Else
63390     If UseStandard Then
63400      .PDFOwnerPass = 0
63410     End If
63420   End If
63430   tStr = reg.GetRegistryValue("PDFOwnerPasswordString")
63440   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63450     .PDFOwnerPasswordString = ""
63460    Else
63470     If LenB(tStr) > 0 Then
63480      .PDFOwnerPasswordString = tStr
63490     End If
63500   End If
63510   tStr = reg.GetRegistryValue("PDFUserPass")
63520   If IsNumeric(tStr) Then
63530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63540       .PDFUserPass = CLng(tStr)
63550      Else
63560       If UseStandard Then
63570        .PDFUserPass = 0
63580       End If
63590     End If
63600    Else
63610     If UseStandard Then
63620      .PDFUserPass = 0
63630     End If
63640   End If
63650   tStr = reg.GetRegistryValue("PDFUserPasswordString")
63660   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63670     .PDFUserPasswordString = ""
63680    Else
63690     If LenB(tStr) > 0 Then
63700      .PDFUserPasswordString = tStr
63710     End If
63720   End If
63730   tStr = reg.GetRegistryValue("PDFUseSecurity")
63740   If IsNumeric(tStr) Then
63750     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63760       .PDFUseSecurity = CLng(tStr)
63770      Else
63780       If UseStandard Then
63790        .PDFUseSecurity = 0
63800       End If
63810     End If
63820    Else
63830     If UseStandard Then
63840      .PDFUseSecurity = 0
63850     End If
63860   End If
63870   reg.SubKey = "Printing\Formats\PDF\Signing"
63880   tStr = reg.GetRegistryValue("PDFSigningMultiSignature")
63890   If IsNumeric(tStr) Then
63900     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63910       .PDFSigningMultiSignature = CLng(tStr)
63920      Else
63930       If UseStandard Then
63940        .PDFSigningMultiSignature = 0
63950       End If
63960     End If
63970    Else
63980     If UseStandard Then
63990      .PDFSigningMultiSignature = 0
64000     End If
64010   End If
64020   tStr = reg.GetRegistryValue("PDFSigningPFXFile")
64030   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64040     .PDFSigningPFXFile = ""
64050    Else
64060     If LenB(tStr) > 0 Then
64070      .PDFSigningPFXFile = tStr
64080     End If
64090   End If
64100   tStr = reg.GetRegistryValue("PDFSigningPFXFilePassword")
64110   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64120     .PDFSigningPFXFilePassword = ""
64130    Else
64140     If LenB(tStr) > 0 Then
64150      .PDFSigningPFXFilePassword = tStr
64160     End If
64170   End If
64180   tStr = reg.GetRegistryValue("PDFSigningSignatureContact")
64190   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64200     .PDFSigningSignatureContact = ""
64210    Else
64220     If LenB(tStr) > 0 Then
64230      .PDFSigningSignatureContact = tStr
64240     End If
64250   End If
64260   tStr = reg.GetRegistryValue("PDFSigningSignatureLeftX")
64270   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
64280     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
64290       .PDFSigningSignatureLeftX = CDbl(Replace$(tStr, ".", GetDecimalChar))
64300      Else
64310       If UseStandard Then
64320        .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
64330       End If
64340     End If
64350    Else
64360     If UseStandard Then
64370      .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
64380     End If
64390   End If
64400   tStr = reg.GetRegistryValue("PDFSigningSignatureLeftY")
64410   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
64420     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
64430       .PDFSigningSignatureLeftY = CDbl(Replace$(tStr, ".", GetDecimalChar))
64440      Else
64450       If UseStandard Then
64460        .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
64470       End If
64480     End If
64490    Else
64500     If UseStandard Then
64510      .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
64520     End If
64530   End If
64540   tStr = reg.GetRegistryValue("PDFSigningSignatureLocation")
64550   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64560     .PDFSigningSignatureLocation = ""
64570    Else
64580     If LenB(tStr) > 0 Then
64590      .PDFSigningSignatureLocation = tStr
64600     End If
64610   End If
64620   tStr = reg.GetRegistryValue("PDFSigningSignatureOnPage")
64630   If IsNumeric(tStr) Then
64640     If CLng(tStr) >= 0 Then
64650       .PDFSigningSignatureOnPage = CLng(tStr)
64660      Else
64670       If UseStandard Then
64680        .PDFSigningSignatureOnPage = 1
64690       End If
64700     End If
64710    Else
64720     If UseStandard Then
64730      .PDFSigningSignatureOnPage = 1
64740     End If
64750   End If
64760   tStr = reg.GetRegistryValue("PDFSigningSignatureReason")
64770   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64780     .PDFSigningSignatureReason = ""
64790    Else
64800     If LenB(tStr) > 0 Then
64810      .PDFSigningSignatureReason = tStr
64820     End If
64830   End If
64840   tStr = reg.GetRegistryValue("PDFSigningSignatureRightX")
64850   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
64860     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
64870       .PDFSigningSignatureRightX = CDbl(Replace$(tStr, ".", GetDecimalChar))
64880      Else
64890       If UseStandard Then
64900        .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
64910       End If
64920     End If
64930    Else
64940     If UseStandard Then
64950      .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
64960     End If
64970   End If
64980   tStr = reg.GetRegistryValue("PDFSigningSignatureRightY")
64990   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
65000     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
65010       .PDFSigningSignatureRightY = CDbl(Replace$(tStr, ".", GetDecimalChar))
65020      Else
65030       If UseStandard Then
65040        .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
65050       End If
65060     End If
65070    Else
65080     If UseStandard Then
65090      .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
65100     End If
65110   End If
65120   tStr = reg.GetRegistryValue("PDFSigningSignatureVisible")
65130   If IsNumeric(tStr) Then
65140     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65150       .PDFSigningSignatureVisible = CLng(tStr)
65160      Else
65170       If UseStandard Then
65180        .PDFSigningSignatureVisible = 0
65190       End If
65200     End If
65210    Else
65220     If UseStandard Then
65230      .PDFSigningSignatureVisible = 0
65240     End If
65250   End If
65260   tStr = reg.GetRegistryValue("PDFSigningSignPDF")
65270   If IsNumeric(tStr) Then
65280     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65290       .PDFSigningSignPDF = CLng(tStr)
65300      Else
65310       If UseStandard Then
65320        .PDFSigningSignPDF = 0
65330       End If
65340     End If
65350    Else
65360     If UseStandard Then
65370      .PDFSigningSignPDF = 0
65380     End If
65390   End If
65400   reg.SubKey = "Printing\Formats\PS\LanguageLevel"
65410   tStr = reg.GetRegistryValue("EPSLanguageLevel")
65420   If IsNumeric(tStr) Then
65430     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
65440       .EPSLanguageLevel = CLng(tStr)
65450      Else
65460       If UseStandard Then
65470        .EPSLanguageLevel = 2
65480       End If
65490     End If
65500    Else
65510     If UseStandard Then
65520      .EPSLanguageLevel = 2
65530     End If
65540   End If
65550   tStr = reg.GetRegistryValue("PSLanguageLevel")
65560   If IsNumeric(tStr) Then
65570     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
65580       .PSLanguageLevel = CLng(tStr)
65590      Else
65600       If UseStandard Then
65610        .PSLanguageLevel = 2
65620       End If
65630     End If
65640    Else
65650     If UseStandard Then
65660      .PSLanguageLevel = 2
65670     End If
65680   End If
65690   reg.SubKey = "Program"
65700   tStr = reg.GetRegistryValue("AdditionalGhostscriptParameters")
65710   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
65720     .AdditionalGhostscriptParameters = ""
65730    Else
65740     If LenB(tStr) > 0 Then
65750      .AdditionalGhostscriptParameters = tStr
65760     End If
65770   End If
65780   tStr = reg.GetRegistryValue("AdditionalGhostscriptSearchpath")
65790   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
65800     .AdditionalGhostscriptSearchpath = ""
65810    Else
65820     If LenB(tStr) > 0 Then
65830      .AdditionalGhostscriptSearchpath = tStr
65840     End If
65850   End If
65860   tStr = reg.GetRegistryValue("AddWindowsFontpath")
65870   If IsNumeric(tStr) Then
65880     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65890       .AddWindowsFontpath = CLng(tStr)
65900      Else
65910       If UseStandard Then
65920        .AddWindowsFontpath = 1
65930       End If
65940     End If
65950    Else
65960     If UseStandard Then
65970      .AddWindowsFontpath = 1
65980     End If
65990   End If
66000   tStr = reg.GetRegistryValue("AllowSpecialGSCharsInFilenames")
66010   If IsNumeric(tStr) Then
66020     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66030       .AllowSpecialGSCharsInFilenames = CLng(tStr)
66040      Else
66050       If UseStandard Then
66060        .AllowSpecialGSCharsInFilenames = 1
66070       End If
66080     End If
66090    Else
66100     If UseStandard Then
66110      .AllowSpecialGSCharsInFilenames = 1
66120     End If
66130   End If
66140   tStr = reg.GetRegistryValue("AutosaveDirectory")
66150   If LenB(Trim$(tStr)) > 0 Then
66160     .AutosaveDirectory = CompletePath(tStr)
66170    Else
66180     If UseStandard Then
66190      If InstalledAsServer Then
66200        .AutosaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
66210       Else
66220        .AutosaveDirectory = "<MyFiles>"
66230      End If
66240     End If
66250   End If
66260   tStr = reg.GetRegistryValue("AutosaveFilename")
66270   If LenB(tStr) = 0 And LenB("<DateTime>") > 0 And UseStandard Then
66280     .AutosaveFilename = "<DateTime>"
66290    Else
66300     If LenB(tStr) > 0 Then
66310      .AutosaveFilename = tStr
66320     End If
66330   End If
66340   tStr = reg.GetRegistryValue("AutosaveFormat")
66350   If IsNumeric(tStr) Then
66360     If CLng(tStr) >= 0 And CLng(tStr) <= 14 Then
66370       .AutosaveFormat = CLng(tStr)
66380      Else
66390       If UseStandard Then
66400        .AutosaveFormat = 0
66410       End If
66420     End If
66430    Else
66440     If UseStandard Then
66450      .AutosaveFormat = 0
66460     End If
66470   End If
66480   tStr = reg.GetRegistryValue("AutosaveStartStandardProgram")
66490   If IsNumeric(tStr) Then
66500     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66510       .AutosaveStartStandardProgram = CLng(tStr)
66520      Else
66530       If UseStandard Then
66540        .AutosaveStartStandardProgram = 0
66550       End If
66560     End If
66570    Else
66580     If UseStandard Then
66590      .AutosaveStartStandardProgram = 0
66600     End If
66610   End If
66620   tStr = reg.GetRegistryValue("ClientComputerResolveIPAddress")
66630   If IsNumeric(tStr) Then
66640     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66650       .ClientComputerResolveIPAddress = CLng(tStr)
66660      Else
66670       If UseStandard Then
66680        .ClientComputerResolveIPAddress = 0
66690       End If
66700     End If
66710    Else
66720     If UseStandard Then
66730      .ClientComputerResolveIPAddress = 0
66740     End If
66750   End If
66760   tStr = reg.GetRegistryValue("DisableEmail")
66770   If IsNumeric(tStr) Then
66780     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66790       .DisableEmail = CLng(tStr)
66800      Else
66810       If UseStandard Then
66820        .DisableEmail = 0
66830       End If
66840     End If
66850    Else
66860     If UseStandard Then
66870      .DisableEmail = 0
66880     End If
66890   End If
66900   tStr = reg.GetRegistryValue("DontUseDocumentSettings")
66910   If IsNumeric(tStr) Then
66920     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66930       .DontUseDocumentSettings = CLng(tStr)
66940      Else
66950       If UseStandard Then
66960        .DontUseDocumentSettings = 0
66970       End If
66980     End If
66990    Else
67000     If UseStandard Then
67010      .DontUseDocumentSettings = 0
67020     End If
67030   End If
67040   tStr = reg.GetRegistryValue("FilenameSubstitutions")
67050   If LenB(tStr) = 0 And LenB("Microsoft Word - \.docx\.doc\Microsoft Excel - \.xlsx\.xls\Microsoft PowerPoint - \.pptx\.ppt") > 0 And UseStandard Then
67060     .FilenameSubstitutions = "Microsoft Word - \.docx\.doc\Microsoft Excel - \.xlsx\.xls\Microsoft PowerPoint - \.pptx\.ppt"
67070    Else
67080     If LenB(tStr) > 0 Then
67090      .FilenameSubstitutions = tStr
67100     End If
67110   End If
67120   tStr = reg.GetRegistryValue("FilenameSubstitutionsOnlyInTitle")
67130   If IsNumeric(tStr) Then
67140     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67150       .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
67160      Else
67170       If UseStandard Then
67180        .FilenameSubstitutionsOnlyInTitle = 1
67190       End If
67200     End If
67210    Else
67220     If UseStandard Then
67230      .FilenameSubstitutionsOnlyInTitle = 1
67240     End If
67250   End If
67260   tStr = reg.GetRegistryValue("Language")
67270   If LenB(tStr) = 0 And LenB("english") > 0 And UseStandard Then
67280     .Language = "english"
67290    Else
67300     If LenB(tStr) > 0 Then
67310      .Language = tStr
67320     End If
67330   End If
67340   tStr = reg.GetRegistryValue("LastSaveDirectory")
67350   If LenB(Trim$(tStr)) > 0 Then
67360     .LastSaveDirectory = CompletePath(tStr)
67370    Else
67380     If UseStandard Then
67390      If InstalledAsServer Then
67400        .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
67410       Else
67420        .LastSaveDirectory = "<MyFiles>"
67430      End If
67440     End If
67450   End If
67460   tStr = reg.GetRegistryValue("LastUpdateCheck")
67470   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67480     .LastUpdateCheck = ""
67490    Else
67500     If LenB(tStr) > 0 Then
67510      .LastUpdateCheck = tStr
67520     End If
67530   End If
67540   tStr = reg.GetRegistryValue("Logging")
67550   If IsNumeric(tStr) Then
67560     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67570       .Logging = CLng(tStr)
67580      Else
67590       If UseStandard Then
67600        .Logging = 0
67610       End If
67620     End If
67630    Else
67640     If UseStandard Then
67650      .Logging = 0
67660     End If
67670   End If
67680   tStr = reg.GetRegistryValue("LogLines")
67690   If IsNumeric(tStr) Then
67700     If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
67710       .LogLines = CLng(tStr)
67720      Else
67730       If UseStandard Then
67740        .LogLines = 100
67750       End If
67760     End If
67770    Else
67780     If UseStandard Then
67790      .LogLines = 100
67800     End If
67810   End If
67820   tStr = reg.GetRegistryValue("NoConfirmMessageSwitchingDefaultprinter")
67830   If IsNumeric(tStr) Then
67840     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67850       .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
67860      Else
67870       If UseStandard Then
67880        .NoConfirmMessageSwitchingDefaultprinter = 0
67890       End If
67900     End If
67910    Else
67920     If UseStandard Then
67930      .NoConfirmMessageSwitchingDefaultprinter = 0
67940     End If
67950   End If
67960   tStr = reg.GetRegistryValue("NoProcessingAtStartup")
67970   If IsNumeric(tStr) Then
67980     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67990       .NoProcessingAtStartup = CLng(tStr)
68000      Else
68010       If UseStandard Then
68020        .NoProcessingAtStartup = 0
68030       End If
68040     End If
68050    Else
68060     If UseStandard Then
68070      .NoProcessingAtStartup = 0
68080     End If
68090   End If
68100   tStr = reg.GetRegistryValue("NoPSCheck")
68110   If IsNumeric(tStr) Then
68120     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68130       .NoPSCheck = CLng(tStr)
68140      Else
68150       If UseStandard Then
68160        .NoPSCheck = 0
68170       End If
68180     End If
68190    Else
68200     If UseStandard Then
68210      .NoPSCheck = 0
68220     End If
68230   End If
68240   tStr = reg.GetRegistryValue("OptionsDesign")
68250   If IsNumeric(tStr) Then
68260     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
68270       .OptionsDesign = CLng(tStr)
68280      Else
68290       If UseStandard Then
68300        .OptionsDesign = 0
68310       End If
68320     End If
68330    Else
68340     If UseStandard Then
68350      .OptionsDesign = 0
68360     End If
68370   End If
68380   tStr = reg.GetRegistryValue("OptionsEnabled")
68390   If IsNumeric(tStr) Then
68400     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68410       .OptionsEnabled = CLng(tStr)
68420      Else
68430       If UseStandard Then
68440        .OptionsEnabled = 1
68450       End If
68460     End If
68470    Else
68480     If UseStandard Then
68490      .OptionsEnabled = 1
68500     End If
68510   End If
68520   tStr = reg.GetRegistryValue("OptionsVisible")
68530   If IsNumeric(tStr) Then
68540     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68550       .OptionsVisible = CLng(tStr)
68560      Else
68570       If UseStandard Then
68580        .OptionsVisible = 1
68590       End If
68600     End If
68610    Else
68620     If UseStandard Then
68630      .OptionsVisible = 1
68640     End If
68650   End If
68660   tStr = reg.GetRegistryValue("PrintAfterSaving")
68670   If IsNumeric(tStr) Then
68680     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68690       .PrintAfterSaving = CLng(tStr)
68700      Else
68710       If UseStandard Then
68720        .PrintAfterSaving = 0
68730       End If
68740     End If
68750    Else
68760     If UseStandard Then
68770      .PrintAfterSaving = 0
68780     End If
68790   End If
68800   tStr = reg.GetRegistryValue("PrintAfterSavingBitsPerPixel")
68810   If IsNumeric(tStr) Then
68820     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
68830       .PrintAfterSavingBitsPerPixel = CLng(tStr)
68840      Else
68850       If UseStandard Then
68860        .PrintAfterSavingBitsPerPixel = 2
68870       End If
68880     End If
68890    Else
68900     If UseStandard Then
68910      .PrintAfterSavingBitsPerPixel = 2
68920     End If
68930   End If
68940   tStr = reg.GetRegistryValue("PrintAfterSavingDuplex")
68950   If IsNumeric(tStr) Then
68960     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68970       .PrintAfterSavingDuplex = CLng(tStr)
68980      Else
68990       If UseStandard Then
69000        .PrintAfterSavingDuplex = 0
69010       End If
69020     End If
69030    Else
69040     If UseStandard Then
69050      .PrintAfterSavingDuplex = 0
69060     End If
69070   End If
69080   tStr = reg.GetRegistryValue("PrintAfterSavingMaxResolution")
69090   If IsNumeric(tStr) Then
69100     If CLng(tStr) >= 72 Then
69110       .PrintAfterSavingMaxResolution = CLng(tStr)
69120      Else
69130       If UseStandard Then
69140        .PrintAfterSavingMaxResolution = 600
69150       End If
69160     End If
69170    Else
69180     If UseStandard Then
69190      .PrintAfterSavingMaxResolution = 600
69200     End If
69210   End If
69220   tStr = reg.GetRegistryValue("PrintAfterSavingMaxResolutionEnabled")
69230   If IsNumeric(tStr) Then
69240     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69250       .PrintAfterSavingMaxResolutionEnabled = CLng(tStr)
69260      Else
69270       If UseStandard Then
69280        .PrintAfterSavingMaxResolutionEnabled = 0
69290       End If
69300     End If
69310    Else
69320     If UseStandard Then
69330      .PrintAfterSavingMaxResolutionEnabled = 0
69340     End If
69350   End If
69360   tStr = reg.GetRegistryValue("PrintAfterSavingNoCancel")
69370   If IsNumeric(tStr) Then
69380     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69390       .PrintAfterSavingNoCancel = CLng(tStr)
69400      Else
69410       If UseStandard Then
69420        .PrintAfterSavingNoCancel = 0
69430       End If
69440     End If
69450    Else
69460     If UseStandard Then
69470      .PrintAfterSavingNoCancel = 0
69480     End If
69490   End If
69500   tStr = reg.GetRegistryValue("PrintAfterSavingPrinter")
69510   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69520     .PrintAfterSavingPrinter = ""
69530    Else
69540     If LenB(tStr) > 0 Then
69550      .PrintAfterSavingPrinter = tStr
69560     End If
69570   End If
69580   tStr = reg.GetRegistryValue("PrintAfterSavingQueryUser")
69590   If IsNumeric(tStr) Then
69600     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
69610       .PrintAfterSavingQueryUser = CLng(tStr)
69620      Else
69630       If UseStandard Then
69640        .PrintAfterSavingQueryUser = 0
69650       End If
69660     End If
69670    Else
69680     If UseStandard Then
69690      .PrintAfterSavingQueryUser = 0
69700     End If
69710   End If
69720   tStr = reg.GetRegistryValue("PrintAfterSavingTumble")
69730   If IsNumeric(tStr) Then
69740     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
69750       .PrintAfterSavingTumble = CLng(tStr)
69760      Else
69770       If UseStandard Then
69780        .PrintAfterSavingTumble = 0
69790       End If
69800     End If
69810    Else
69820     If UseStandard Then
69830      .PrintAfterSavingTumble = 0
69840     End If
69850   End If
69860   tStr = reg.GetRegistryValue("PrinterStop")
69870   If IsNumeric(tStr) Then
69880     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69890       .PrinterStop = CLng(tStr)
69900      Else
69910       If UseStandard Then
69920        .PrinterStop = 0
69930       End If
69940     End If
69950    Else
69960     If UseStandard Then
69970      .PrinterStop = 0
69980     End If
69990   End If
70000   tStr = reg.GetRegistryValue("PrinterTemppath")
70010   WriteToSpecialLogfile "reg.GetRegistryValue(""PrinterTemppath"")=" & tStr
70020   WriteToSpecialLogfile "Options.PrinterTemppath1=" & .PrinterTemppath
70030   If hkey1 = HKEY_USERS Then
70040     If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then
70050       .PrinterTemppath = tStr
70060      Else
70070       If UseStandard Then
70080         .PrinterTemppath = GetTempPath
70090        Else
70100         .PrinterTemppath = tStr
70110       End If
70120     End If
70130    Else
70140     If LenB(Trim$(tStr)) > 0 Then
70150      If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then
70160        .PrinterTemppath = tStr
70170       Else
70180        MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))
70190        If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then
70200          If UseStandard Then
70210            .PrinterTemppath = GetTempPath
70220           Else
70230            .PrinterTemppath = ""
70240            If NoMsg = False Then
70250             MsgBox "PrinterTemppath: '" & tStr & "' = '" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & "'" & _
             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
70270            End If
70280          End If
70290         Else
70300          .PrinterTemppath = tStr
70310        End If
70320      End If
70330     End If
70340   End If
70350   WriteToSpecialLogfile "Options.PrinterTemppath2=" & .PrinterTemppath
70360   tStr = reg.GetRegistryValue("ProcessPriority")
70370   If IsNumeric(tStr) Then
70380     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
70390       .ProcessPriority = CLng(tStr)
70400      Else
70410       If UseStandard Then
70420        .ProcessPriority = 1
70430       End If
70440     End If
70450    Else
70460     If UseStandard Then
70470      .ProcessPriority = 1
70480     End If
70490   End If
70500   tStr = reg.GetRegistryValue("ProgramFont")
70510   If LenB(tStr) = 0 And LenB("MS Sans Serif") > 0 And UseStandard Then
70520     .ProgramFont = "MS Sans Serif"
70530    Else
70540     If LenB(tStr) > 0 Then
70550      .ProgramFont = tStr
70560     End If
70570   End If
70580   tStr = reg.GetRegistryValue("ProgramFontCharset")
70590   If IsNumeric(tStr) Then
70600     If CLng(tStr) >= 0 Then
70610       .ProgramFontCharset = CLng(tStr)
70620      Else
70630       If UseStandard Then
70640        .ProgramFontCharset = 0
70650       End If
70660     End If
70670    Else
70680     If UseStandard Then
70690      .ProgramFontCharset = 0
70700     End If
70710   End If
70720   tStr = reg.GetRegistryValue("ProgramFontSize")
70730   If IsNumeric(tStr) Then
70740     If CLng(tStr) >= 6 And CLng(tStr) <= 72 Then
70750       .ProgramFontSize = CLng(tStr)
70760      Else
70770       If UseStandard Then
70780        .ProgramFontSize = 8
70790       End If
70800     End If
70810    Else
70820     If UseStandard Then
70830      .ProgramFontSize = 8
70840     End If
70850   End If
70860   tStr = reg.GetRegistryValue("RemoveAllKnownFileExtensions")
70870   If IsNumeric(tStr) Then
70880     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70890       .RemoveAllKnownFileExtensions = CLng(tStr)
70900      Else
70910       If UseStandard Then
70920        .RemoveAllKnownFileExtensions = 1
70930       End If
70940     End If
70950    Else
70960     If UseStandard Then
70970      .RemoveAllKnownFileExtensions = 1
70980     End If
70990   End If
71000   tStr = reg.GetRegistryValue("RemoveSpaces")
71010   If IsNumeric(tStr) Then
71020     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71030       .RemoveSpaces = CLng(tStr)
71040      Else
71050       If UseStandard Then
71060        .RemoveSpaces = 1
71070       End If
71080     End If
71090    Else
71100     If UseStandard Then
71110      .RemoveSpaces = 1
71120     End If
71130   End If
71140   tStr = reg.GetRegistryValue("RunProgramAfterSaving")
71150   If IsNumeric(tStr) Then
71160     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71170       .RunProgramAfterSaving = CLng(tStr)
71180      Else
71190       If UseStandard Then
71200        .RunProgramAfterSaving = 0
71210       End If
71220     End If
71230    Else
71240     If UseStandard Then
71250      .RunProgramAfterSaving = 0
71260     End If
71270   End If
71280   tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramname")
71290   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
71300     .RunProgramAfterSavingProgramname = ""
71310    Else
71320     If LenB(tStr) > 0 Then
71330      .RunProgramAfterSavingProgramname = tStr
71340     End If
71350   End If
71360   tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramParameters")
71370   If LenB(tStr) = 0 And LenB("""<OutputFilename>""") > 0 And UseStandard Then
71380     .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
71390    Else
71400     If LenB(tStr) > 0 Then
71410      .RunProgramAfterSavingProgramParameters = tStr
71420     End If
71430   End If
71440   tStr = reg.GetRegistryValue("RunProgramAfterSavingWaitUntilReady")
71450   If IsNumeric(tStr) Then
71460     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71470       .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
71480      Else
71490       If UseStandard Then
71500        .RunProgramAfterSavingWaitUntilReady = 1
71510       End If
71520     End If
71530    Else
71540     If UseStandard Then
71550      .RunProgramAfterSavingWaitUntilReady = 1
71560     End If
71570   End If
71580   tStr = reg.GetRegistryValue("RunProgramAfterSavingWindowstyle")
71590   If IsNumeric(tStr) Then
71600     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
71610       .RunProgramAfterSavingWindowstyle = CLng(tStr)
71620      Else
71630       If UseStandard Then
71640        .RunProgramAfterSavingWindowstyle = 1
71650       End If
71660     End If
71670    Else
71680     If UseStandard Then
71690      .RunProgramAfterSavingWindowstyle = 1
71700     End If
71710   End If
71720   tStr = reg.GetRegistryValue("RunProgramBeforeSaving")
71730   If IsNumeric(tStr) Then
71740     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71750       .RunProgramBeforeSaving = CLng(tStr)
71760      Else
71770       If UseStandard Then
71780        .RunProgramBeforeSaving = 0
71790       End If
71800     End If
71810    Else
71820     If UseStandard Then
71830      .RunProgramBeforeSaving = 0
71840     End If
71850   End If
71860   tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramname")
71870   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
71880     .RunProgramBeforeSavingProgramname = ""
71890    Else
71900     If LenB(tStr) > 0 Then
71910      .RunProgramBeforeSavingProgramname = tStr
71920     End If
71930   End If
71940   tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramParameters")
71950   If LenB(tStr) = 0 And LenB("""<TempFilename>""") > 0 And UseStandard Then
71960     .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
71970    Else
71980     If LenB(tStr) > 0 Then
71990      .RunProgramBeforeSavingProgramParameters = tStr
72000     End If
72010   End If
72020   tStr = reg.GetRegistryValue("RunProgramBeforeSavingWindowstyle")
72030   If IsNumeric(tStr) Then
72040     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
72050       .RunProgramBeforeSavingWindowstyle = CLng(tStr)
72060      Else
72070       If UseStandard Then
72080        .RunProgramBeforeSavingWindowstyle = 1
72090       End If
72100     End If
72110    Else
72120     If UseStandard Then
72130      .RunProgramBeforeSavingWindowstyle = 1
72140     End If
72150   End If
72160   tStr = reg.GetRegistryValue("SaveFilename")
72170   If LenB(tStr) = 0 And LenB("<Title>") > 0 And UseStandard Then
72180     .SaveFilename = "<Title>"
72190    Else
72200     If LenB(tStr) > 0 Then
72210      .SaveFilename = tStr
72220     End If
72230   End If
72240   tStr = reg.GetRegistryValue("SendEmailAfterAutoSaving")
72250   If IsNumeric(tStr) Then
72260     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
72270       .SendEmailAfterAutoSaving = CLng(tStr)
72280      Else
72290       If UseStandard Then
72300        .SendEmailAfterAutoSaving = 0
72310       End If
72320     End If
72330    Else
72340     If UseStandard Then
72350      .SendEmailAfterAutoSaving = 0
72360     End If
72370   End If
72380   tStr = reg.GetRegistryValue("SendMailMethod")
72390   If IsNumeric(tStr) Then
72400     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
72410       .SendMailMethod = CLng(tStr)
72420      Else
72430       If UseStandard Then
72440        .SendMailMethod = 0
72450       End If
72460     End If
72470    Else
72480     If UseStandard Then
72490      .SendMailMethod = 0
72500     End If
72510   End If
72520   tStr = reg.GetRegistryValue("ShowAnimation")
72530   If IsNumeric(tStr) Then
72540     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
72550       .ShowAnimation = CLng(tStr)
72560      Else
72570       If UseStandard Then
72580        .ShowAnimation = 1
72590       End If
72600     End If
72610    Else
72620     If UseStandard Then
72630      .ShowAnimation = 1
72640     End If
72650   End If
72660   tStr = reg.GetRegistryValue("StartStandardProgram")
72670   If IsNumeric(tStr) Then
72680     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
72690       .StartStandardProgram = CLng(tStr)
72700      Else
72710       If UseStandard Then
72720        .StartStandardProgram = 1
72730       End If
72740     End If
72750    Else
72760     If UseStandard Then
72770      .StartStandardProgram = 1
72780     End If
72790   End If
72800   tStr = reg.GetRegistryValue("Toolbars")
72810   If IsNumeric(tStr) Then
72820     If CLng(tStr) >= 0 Then
72830       .Toolbars = CLng(tStr)
72840      Else
72850       If UseStandard Then
72860        .Toolbars = 1
72870       End If
72880     End If
72890    Else
72900     If UseStandard Then
72910      .Toolbars = 1
72920     End If
72930   End If
72940   tStr = reg.GetRegistryValue("UpdateInterval")
72950   If IsNumeric(tStr) Then
72960     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
72970       .UpdateInterval = CLng(tStr)
72980      Else
72990       If UseStandard Then
73000        .UpdateInterval = 2
73010       End If
73020     End If
73030    Else
73040     If UseStandard Then
73050      .UpdateInterval = 2
73060     End If
73070   End If
73080   tStr = reg.GetRegistryValue("UseAutosave")
73090   If IsNumeric(tStr) Then
73100     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
73110       .UseAutosave = CLng(tStr)
73120      Else
73130       If UseStandard Then
73140        .UseAutosave = 0
73150       End If
73160     End If
73170    Else
73180     If UseStandard Then
73190      .UseAutosave = 0
73200     End If
73210   End If
73220   tStr = reg.GetRegistryValue("UseAutosaveDirectory")
73230   If IsNumeric(tStr) Then
73240     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
73250       .UseAutosaveDirectory = CLng(tStr)
73260      Else
73270       If UseStandard Then
73280        .UseAutosaveDirectory = 1
73290       End If
73300     End If
73310    Else
73320     If UseStandard Then
73330      .UseAutosaveDirectory = 1
73340     End If
73350   End If
73360  End With
73370  Set reg = Nothing
73380  ReadOptionsReg = myOptions
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

Public Sub SaveOptionREG(sOptions As tOptions, OptionName As String, Optional hkey1 As hkey = HKEY_CURRENT_USER, Optional ProfileName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = hkey1
50040
50050  ProfileName = Trim$(ProfileName)
50060
50070  If LenB(ProfileName) > 0 Then
50080    reg.KeyRoot = "Software\PDFCreator\Profiles\" & ProfileName
50090   Else
50100    reg.KeyRoot = "Software\PDFCreator"
50110  End If
50120
50130  With sOptions
50140   reg.SubKey = "Ghostscript"
50150   If UCase$(OptionName) = "DIRECTORYGHOSTSCRIPTBINARIES" Then
50160    If Not reg.KeyExists Then
50170     reg.CreateKey
50180    End If
50190    reg.SetRegistryValue "DirectoryGhostscriptBinaries", CStr(.DirectoryGhostscriptBinaries), REG_SZ
50200    Set reg = Nothing
50210    Exit Sub
50220   End If
50230   If UCase$(OptionName) = "DIRECTORYGHOSTSCRIPTFONTS" Then
50240    If Not reg.KeyExists Then
50250     reg.CreateKey
50260    End If
50270    reg.SetRegistryValue "DirectoryGhostscriptFonts", CStr(.DirectoryGhostscriptFonts), REG_SZ
50280    Set reg = Nothing
50290    Exit Sub
50300   End If
50310   If UCase$(OptionName) = "DIRECTORYGHOSTSCRIPTLIBRARIES" Then
50320    If Not reg.KeyExists Then
50330     reg.CreateKey
50340    End If
50350    reg.SetRegistryValue "DirectoryGhostscriptLibraries", CStr(.DirectoryGhostscriptLibraries), REG_SZ
50360    Set reg = Nothing
50370    Exit Sub
50380   End If
50390   If UCase$(OptionName) = "DIRECTORYGHOSTSCRIPTRESOURCE" Then
50400    If Not reg.KeyExists Then
50410     reg.CreateKey
50420    End If
50430    reg.SetRegistryValue "DirectoryGhostscriptResource", CStr(.DirectoryGhostscriptResource), REG_SZ
50440    Set reg = Nothing
50450    Exit Sub
50460   End If
50470   reg.SubKey = "Printing"
50480   If UCase$(OptionName) = "COUNTER" Then
50490    If Not reg.KeyExists Then
50500     reg.CreateKey
50510    End If
50520    reg.SetRegistryValue "Counter", CStr(.Counter), REG_SZ
50530    Set reg = Nothing
50540    Exit Sub
50550   End If
50560   If UCase$(OptionName) = "DEVICEHEIGHTPOINTS" Then
50570    If Not reg.KeyExists Then
50580     reg.CreateKey
50590    End If
50600   reg.SetRegistryValue "DeviceHeightPoints", Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), REG_SZ
50610    Set reg = Nothing
50620    Exit Sub
50630   End If
50640   If UCase$(OptionName) = "DEVICEWIDTHPOINTS" Then
50650    If Not reg.KeyExists Then
50660     reg.CreateKey
50670    End If
50680   reg.SetRegistryValue "DeviceWidthPoints", Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), REG_SZ
50690    Set reg = Nothing
50700    Exit Sub
50710   End If
50720   If UCase$(OptionName) = "ONEPAGEPERFILE" Then
50730    If Not reg.KeyExists Then
50740     reg.CreateKey
50750    End If
50760    reg.SetRegistryValue "OnePagePerFile", CStr(Abs(.OnePagePerFile)), REG_SZ
50770    Set reg = Nothing
50780    Exit Sub
50790   End If
50800   If UCase$(OptionName) = "PAPERSIZE" Then
50810    If Not reg.KeyExists Then
50820     reg.CreateKey
50830    End If
50840    reg.SetRegistryValue "Papersize", CStr(.Papersize), REG_SZ
50850    Set reg = Nothing
50860    Exit Sub
50870   End If
50880   If UCase$(OptionName) = "STAMPFONTCOLOR" Then
50890    If Not reg.KeyExists Then
50900     reg.CreateKey
50910    End If
50920    reg.SetRegistryValue "StampFontColor", CStr(.StampFontColor), REG_SZ
50930    Set reg = Nothing
50940    Exit Sub
50950   End If
50960   If UCase$(OptionName) = "STAMPFONTNAME" Then
50970    If Not reg.KeyExists Then
50980     reg.CreateKey
50990    End If
51000    reg.SetRegistryValue "StampFontname", CStr(.StampFontname), REG_SZ
51010    Set reg = Nothing
51020    Exit Sub
51030   End If
51040   If UCase$(OptionName) = "STAMPFONTSIZE" Then
51050    If Not reg.KeyExists Then
51060     reg.CreateKey
51070    End If
51080    reg.SetRegistryValue "StampFontsize", CStr(.StampFontsize), REG_SZ
51090    Set reg = Nothing
51100    Exit Sub
51110   End If
51120   If UCase$(OptionName) = "STAMPOUTLINEFONTTHICKNESS" Then
51130    If Not reg.KeyExists Then
51140     reg.CreateKey
51150    End If
51160    reg.SetRegistryValue "StampOutlineFontthickness", CStr(.StampOutlineFontthickness), REG_SZ
51170    Set reg = Nothing
51180    Exit Sub
51190   End If
51200   If UCase$(OptionName) = "STAMPSTRING" Then
51210    If Not reg.KeyExists Then
51220     reg.CreateKey
51230    End If
51240    reg.SetRegistryValue "StampString", CStr(.StampString), REG_SZ
51250    Set reg = Nothing
51260    Exit Sub
51270   End If
51280   If UCase$(OptionName) = "STAMPUSEOUTLINEFONT" Then
51290    If Not reg.KeyExists Then
51300     reg.CreateKey
51310    End If
51320    reg.SetRegistryValue "StampUseOutlineFont", CStr(Abs(.StampUseOutlineFont)), REG_SZ
51330    Set reg = Nothing
51340    Exit Sub
51350   End If
51360   If UCase$(OptionName) = "STANDARDAUTHOR" Then
51370    If Not reg.KeyExists Then
51380     reg.CreateKey
51390    End If
51400    reg.SetRegistryValue "StandardAuthor", CStr(.StandardAuthor), REG_SZ
51410    Set reg = Nothing
51420    Exit Sub
51430   End If
51440   If UCase$(OptionName) = "STANDARDCREATIONDATE" Then
51450    If Not reg.KeyExists Then
51460     reg.CreateKey
51470    End If
51480    reg.SetRegistryValue "StandardCreationdate", CStr(.StandardCreationdate), REG_SZ
51490    Set reg = Nothing
51500    Exit Sub
51510   End If
51520   If UCase$(OptionName) = "STANDARDDATEFORMAT" Then
51530    If Not reg.KeyExists Then
51540     reg.CreateKey
51550    End If
51560    reg.SetRegistryValue "StandardDateformat", CStr(.StandardDateformat), REG_SZ
51570    Set reg = Nothing
51580    Exit Sub
51590   End If
51600   If UCase$(OptionName) = "STANDARDKEYWORDS" Then
51610    If Not reg.KeyExists Then
51620     reg.CreateKey
51630    End If
51640    reg.SetRegistryValue "StandardKeywords", CStr(.StandardKeywords), REG_SZ
51650    Set reg = Nothing
51660    Exit Sub
51670   End If
51680   If UCase$(OptionName) = "STANDARDMAILDOMAIN" Then
51690    If Not reg.KeyExists Then
51700     reg.CreateKey
51710    End If
51720    reg.SetRegistryValue "StandardMailDomain", CStr(.StandardMailDomain), REG_SZ
51730    Set reg = Nothing
51740    Exit Sub
51750   End If
51760   If UCase$(OptionName) = "STANDARDMODIFYDATE" Then
51770    If Not reg.KeyExists Then
51780     reg.CreateKey
51790    End If
51800    reg.SetRegistryValue "StandardModifydate", CStr(.StandardModifydate), REG_SZ
51810    Set reg = Nothing
51820    Exit Sub
51830   End If
51840   If UCase$(OptionName) = "STANDARDSAVEFORMAT" Then
51850    If Not reg.KeyExists Then
51860     reg.CreateKey
51870    End If
51880    reg.SetRegistryValue "StandardSaveformat", CStr(.StandardSaveformat), REG_SZ
51890    Set reg = Nothing
51900    Exit Sub
51910   End If
51920   If UCase$(OptionName) = "STANDARDSUBJECT" Then
51930    If Not reg.KeyExists Then
51940     reg.CreateKey
51950    End If
51960    reg.SetRegistryValue "StandardSubject", CStr(.StandardSubject), REG_SZ
51970    Set reg = Nothing
51980    Exit Sub
51990   End If
52000   If UCase$(OptionName) = "STANDARDTITLE" Then
52010    If Not reg.KeyExists Then
52020     reg.CreateKey
52030    End If
52040    reg.SetRegistryValue "StandardTitle", CStr(.StandardTitle), REG_SZ
52050    Set reg = Nothing
52060    Exit Sub
52070   End If
52080   If UCase$(OptionName) = "USECREATIONDATENOW" Then
52090    If Not reg.KeyExists Then
52100     reg.CreateKey
52110    End If
52120    reg.SetRegistryValue "UseCreationDateNow", CStr(Abs(.UseCreationDateNow)), REG_SZ
52130    Set reg = Nothing
52140    Exit Sub
52150   End If
52160   If UCase$(OptionName) = "USECUSTOMPAPERSIZE" Then
52170    If Not reg.KeyExists Then
52180     reg.CreateKey
52190    End If
52200    reg.SetRegistryValue "UseCustomPaperSize", CStr(.UseCustomPaperSize), REG_SZ
52210    Set reg = Nothing
52220    Exit Sub
52230   End If
52240   If UCase$(OptionName) = "USEFIXPAPERSIZE" Then
52250    If Not reg.KeyExists Then
52260     reg.CreateKey
52270    End If
52280    reg.SetRegistryValue "UseFixPapersize", CStr(Abs(.UseFixPapersize)), REG_SZ
52290    Set reg = Nothing
52300    Exit Sub
52310   End If
52320   If UCase$(OptionName) = "USESTANDARDAUTHOR" Then
52330    If Not reg.KeyExists Then
52340     reg.CreateKey
52350    End If
52360    reg.SetRegistryValue "UseStandardAuthor", CStr(Abs(.UseStandardAuthor)), REG_SZ
52370    Set reg = Nothing
52380    Exit Sub
52390   End If
52400   reg.SubKey = "Printing\Formats\Bitmap\Colors"
52410   If UCase$(OptionName) = "BMPCOLORSCOUNT" Then
52420    If Not reg.KeyExists Then
52430     reg.CreateKey
52440    End If
52450    reg.SetRegistryValue "BMPColorscount", CStr(.BMPColorscount), REG_SZ
52460    Set reg = Nothing
52470    Exit Sub
52480   End If
52490   If UCase$(OptionName) = "BMPRESOLUTION" Then
52500    If Not reg.KeyExists Then
52510     reg.CreateKey
52520    End If
52530    reg.SetRegistryValue "BMPResolution", CStr(.BMPResolution), REG_SZ
52540    Set reg = Nothing
52550    Exit Sub
52560   End If
52570   If UCase$(OptionName) = "JPEGCOLORSCOUNT" Then
52580    If Not reg.KeyExists Then
52590     reg.CreateKey
52600    End If
52610    reg.SetRegistryValue "JPEGColorscount", CStr(.JPEGColorscount), REG_SZ
52620    Set reg = Nothing
52630    Exit Sub
52640   End If
52650   If UCase$(OptionName) = "JPEGQUALITY" Then
52660    If Not reg.KeyExists Then
52670     reg.CreateKey
52680    End If
52690    reg.SetRegistryValue "JPEGQuality", CStr(.JPEGQuality), REG_SZ
52700    Set reg = Nothing
52710    Exit Sub
52720   End If
52730   If UCase$(OptionName) = "JPEGRESOLUTION" Then
52740    If Not reg.KeyExists Then
52750     reg.CreateKey
52760    End If
52770    reg.SetRegistryValue "JPEGResolution", CStr(.JPEGResolution), REG_SZ
52780    Set reg = Nothing
52790    Exit Sub
52800   End If
52810   If UCase$(OptionName) = "PCLCOLORSCOUNT" Then
52820    If Not reg.KeyExists Then
52830     reg.CreateKey
52840    End If
52850    reg.SetRegistryValue "PCLColorsCount", CStr(.PCLColorsCount), REG_SZ
52860    Set reg = Nothing
52870    Exit Sub
52880   End If
52890   If UCase$(OptionName) = "PCLRESOLUTION" Then
52900    If Not reg.KeyExists Then
52910     reg.CreateKey
52920    End If
52930    reg.SetRegistryValue "PCLResolution", CStr(.PCLResolution), REG_SZ
52940    Set reg = Nothing
52950    Exit Sub
52960   End If
52970   If UCase$(OptionName) = "PCXCOLORSCOUNT" Then
52980    If Not reg.KeyExists Then
52990     reg.CreateKey
53000    End If
53010    reg.SetRegistryValue "PCXColorscount", CStr(.PCXColorscount), REG_SZ
53020    Set reg = Nothing
53030    Exit Sub
53040   End If
53050   If UCase$(OptionName) = "PCXRESOLUTION" Then
53060    If Not reg.KeyExists Then
53070     reg.CreateKey
53080    End If
53090    reg.SetRegistryValue "PCXResolution", CStr(.PCXResolution), REG_SZ
53100    Set reg = Nothing
53110    Exit Sub
53120   End If
53130   If UCase$(OptionName) = "PNGCOLORSCOUNT" Then
53140    If Not reg.KeyExists Then
53150     reg.CreateKey
53160    End If
53170    reg.SetRegistryValue "PNGColorscount", CStr(.PNGColorscount), REG_SZ
53180    Set reg = Nothing
53190    Exit Sub
53200   End If
53210   If UCase$(OptionName) = "PNGRESOLUTION" Then
53220    If Not reg.KeyExists Then
53230     reg.CreateKey
53240    End If
53250    reg.SetRegistryValue "PNGResolution", CStr(.PNGResolution), REG_SZ
53260    Set reg = Nothing
53270    Exit Sub
53280   End If
53290   If UCase$(OptionName) = "PSDCOLORSCOUNT" Then
53300    If Not reg.KeyExists Then
53310     reg.CreateKey
53320    End If
53330    reg.SetRegistryValue "PSDColorsCount", CStr(.PSDColorsCount), REG_SZ
53340    Set reg = Nothing
53350    Exit Sub
53360   End If
53370   If UCase$(OptionName) = "PSDRESOLUTION" Then
53380    If Not reg.KeyExists Then
53390     reg.CreateKey
53400    End If
53410    reg.SetRegistryValue "PSDResolution", CStr(.PSDResolution), REG_SZ
53420    Set reg = Nothing
53430    Exit Sub
53440   End If
53450   If UCase$(OptionName) = "RAWCOLORSCOUNT" Then
53460    If Not reg.KeyExists Then
53470     reg.CreateKey
53480    End If
53490    reg.SetRegistryValue "RAWColorsCount", CStr(.RAWColorsCount), REG_SZ
53500    Set reg = Nothing
53510    Exit Sub
53520   End If
53530   If UCase$(OptionName) = "RAWRESOLUTION" Then
53540    If Not reg.KeyExists Then
53550     reg.CreateKey
53560    End If
53570    reg.SetRegistryValue "RAWResolution", CStr(.RAWResolution), REG_SZ
53580    Set reg = Nothing
53590    Exit Sub
53600   End If
53610   If UCase$(OptionName) = "SVGRESOLUTION" Then
53620    If Not reg.KeyExists Then
53630     reg.CreateKey
53640    End If
53650    reg.SetRegistryValue "SVGResolution", CStr(.SVGResolution), REG_SZ
53660    Set reg = Nothing
53670    Exit Sub
53680   End If
53690   If UCase$(OptionName) = "TIFFCOLORSCOUNT" Then
53700    If Not reg.KeyExists Then
53710     reg.CreateKey
53720    End If
53730    reg.SetRegistryValue "TIFFColorscount", CStr(.TIFFColorscount), REG_SZ
53740    Set reg = Nothing
53750    Exit Sub
53760   End If
53770   If UCase$(OptionName) = "TIFFRESOLUTION" Then
53780    If Not reg.KeyExists Then
53790     reg.CreateKey
53800    End If
53810    reg.SetRegistryValue "TIFFResolution", CStr(.TIFFResolution), REG_SZ
53820    Set reg = Nothing
53830    Exit Sub
53840   End If
53850   reg.SubKey = "Printing\Formats\PDF\Colors"
53860   If UCase$(OptionName) = "PDFCOLORSCMYKTORGB" Then
53870    If Not reg.KeyExists Then
53880     reg.CreateKey
53890    End If
53900    reg.SetRegistryValue "PDFColorsCMYKToRGB", CStr(Abs(.PDFColorsCMYKToRGB)), REG_SZ
53910    Set reg = Nothing
53920    Exit Sub
53930   End If
53940   If UCase$(OptionName) = "PDFCOLORSCOLORMODEL" Then
53950    If Not reg.KeyExists Then
53960     reg.CreateKey
53970    End If
53980    reg.SetRegistryValue "PDFColorsColorModel", CStr(.PDFColorsColorModel), REG_SZ
53990    Set reg = Nothing
54000    Exit Sub
54010   End If
54020   If UCase$(OptionName) = "PDFCOLORSPRESERVEHALFTONE" Then
54030    If Not reg.KeyExists Then
54040     reg.CreateKey
54050    End If
54060    reg.SetRegistryValue "PDFColorsPreserveHalftone", CStr(Abs(.PDFColorsPreserveHalftone)), REG_SZ
54070    Set reg = Nothing
54080    Exit Sub
54090   End If
54100   If UCase$(OptionName) = "PDFCOLORSPRESERVEOVERPRINT" Then
54110    If Not reg.KeyExists Then
54120     reg.CreateKey
54130    End If
54140    reg.SetRegistryValue "PDFColorsPreserveOverprint", CStr(Abs(.PDFColorsPreserveOverprint)), REG_SZ
54150    Set reg = Nothing
54160    Exit Sub
54170   End If
54180   If UCase$(OptionName) = "PDFCOLORSPRESERVETRANSFER" Then
54190    If Not reg.KeyExists Then
54200     reg.CreateKey
54210    End If
54220    reg.SetRegistryValue "PDFColorsPreserveTransfer", CStr(Abs(.PDFColorsPreserveTransfer)), REG_SZ
54230    Set reg = Nothing
54240    Exit Sub
54250   End If
54260   reg.SubKey = "Printing\Formats\PDF\Compression"
54270   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSION" Then
54280    If Not reg.KeyExists Then
54290     reg.CreateKey
54300    End If
54310    reg.SetRegistryValue "PDFCompressionColorCompression", CStr(Abs(.PDFCompressionColorCompression)), REG_SZ
54320    Set reg = Nothing
54330    Exit Sub
54340   End If
54350   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE" Then
54360    If Not reg.KeyExists Then
54370     reg.CreateKey
54380    End If
54390    reg.SetRegistryValue "PDFCompressionColorCompressionChoice", CStr(.PDFCompressionColorCompressionChoice), REG_SZ
54400    Set reg = Nothing
54410    Exit Sub
54420   End If
54430   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR" Then
54440    If Not reg.KeyExists Then
54450     reg.CreateKey
54460    End If
54470   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
54480    Set reg = Nothing
54490    Exit Sub
54500   End If
54510   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR" Then
54520    If Not reg.KeyExists Then
54530     reg.CreateKey
54540    End If
54550   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
54560    Set reg = Nothing
54570    Exit Sub
54580   End If
54590   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMANUALFACTOR" Then
54600    If Not reg.KeyExists Then
54610     reg.CreateKey
54620    End If
54630   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGManualFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGManualFactor), GetDecimalChar, "."), REG_SZ
54640    Set reg = Nothing
54650    Exit Sub
54660   End If
54670   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR" Then
54680    If Not reg.KeyExists Then
54690     reg.CreateKey
54700    End If
54710   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
54720    Set reg = Nothing
54730    Exit Sub
54740   End If
54750   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR" Then
54760    If Not reg.KeyExists Then
54770     reg.CreateKey
54780    End If
54790   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
54800    Set reg = Nothing
54810    Exit Sub
54820   End If
54830   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR" Then
54840    If Not reg.KeyExists Then
54850     reg.CreateKey
54860    End If
54870   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
54880    Set reg = Nothing
54890    Exit Sub
54900   End If
54910   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESAMPLE" Then
54920    If Not reg.KeyExists Then
54930     reg.CreateKey
54940    End If
54950    reg.SetRegistryValue "PDFCompressionColorResample", CStr(Abs(.PDFCompressionColorResample)), REG_SZ
54960    Set reg = Nothing
54970    Exit Sub
54980   End If
54990   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESAMPLECHOICE" Then
55000    If Not reg.KeyExists Then
55010     reg.CreateKey
55020    End If
55030    reg.SetRegistryValue "PDFCompressionColorResampleChoice", CStr(.PDFCompressionColorResampleChoice), REG_SZ
55040    Set reg = Nothing
55050    Exit Sub
55060   End If
55070   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESOLUTION" Then
55080    If Not reg.KeyExists Then
55090     reg.CreateKey
55100    End If
55110    reg.SetRegistryValue "PDFCompressionColorResolution", CStr(.PDFCompressionColorResolution), REG_SZ
55120    Set reg = Nothing
55130    Exit Sub
55140   End If
55150   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSION" Then
55160    If Not reg.KeyExists Then
55170     reg.CreateKey
55180    End If
55190    reg.SetRegistryValue "PDFCompressionGreyCompression", CStr(Abs(.PDFCompressionGreyCompression)), REG_SZ
55200    Set reg = Nothing
55210    Exit Sub
55220   End If
55230   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE" Then
55240    If Not reg.KeyExists Then
55250     reg.CreateKey
55260    End If
55270    reg.SetRegistryValue "PDFCompressionGreyCompressionChoice", CStr(.PDFCompressionGreyCompressionChoice), REG_SZ
55280    Set reg = Nothing
55290    Exit Sub
55300   End If
55310   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR" Then
55320    If Not reg.KeyExists Then
55330     reg.CreateKey
55340    End If
55350   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
55360    Set reg = Nothing
55370    Exit Sub
55380   End If
55390   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR" Then
55400    If Not reg.KeyExists Then
55410     reg.CreateKey
55420    End If
55430   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
55440    Set reg = Nothing
55450    Exit Sub
55460   End If
55470   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMANUALFACTOR" Then
55480    If Not reg.KeyExists Then
55490     reg.CreateKey
55500    End If
55510   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGManualFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGManualFactor), GetDecimalChar, "."), REG_SZ
55520    Set reg = Nothing
55530    Exit Sub
55540   End If
55550   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR" Then
55560    If Not reg.KeyExists Then
55570     reg.CreateKey
55580    End If
55590   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
55600    Set reg = Nothing
55610    Exit Sub
55620   End If
55630   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR" Then
55640    If Not reg.KeyExists Then
55650     reg.CreateKey
55660    End If
55670   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
55680    Set reg = Nothing
55690    Exit Sub
55700   End If
55710   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR" Then
55720    If Not reg.KeyExists Then
55730     reg.CreateKey
55740    End If
55750   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
55760    Set reg = Nothing
55770    Exit Sub
55780   End If
55790   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESAMPLE" Then
55800    If Not reg.KeyExists Then
55810     reg.CreateKey
55820    End If
55830    reg.SetRegistryValue "PDFCompressionGreyResample", CStr(Abs(.PDFCompressionGreyResample)), REG_SZ
55840    Set reg = Nothing
55850    Exit Sub
55860   End If
55870   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESAMPLECHOICE" Then
55880    If Not reg.KeyExists Then
55890     reg.CreateKey
55900    End If
55910    reg.SetRegistryValue "PDFCompressionGreyResampleChoice", CStr(.PDFCompressionGreyResampleChoice), REG_SZ
55920    Set reg = Nothing
55930    Exit Sub
55940   End If
55950   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESOLUTION" Then
55960    If Not reg.KeyExists Then
55970     reg.CreateKey
55980    End If
55990    reg.SetRegistryValue "PDFCompressionGreyResolution", CStr(.PDFCompressionGreyResolution), REG_SZ
56000    Set reg = Nothing
56010    Exit Sub
56020   End If
56030   If UCase$(OptionName) = "PDFCOMPRESSIONMONOCOMPRESSION" Then
56040    If Not reg.KeyExists Then
56050     reg.CreateKey
56060    End If
56070    reg.SetRegistryValue "PDFCompressionMonoCompression", CStr(Abs(.PDFCompressionMonoCompression)), REG_SZ
56080    Set reg = Nothing
56090    Exit Sub
56100   End If
56110   If UCase$(OptionName) = "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE" Then
56120    If Not reg.KeyExists Then
56130     reg.CreateKey
56140    End If
56150    reg.SetRegistryValue "PDFCompressionMonoCompressionChoice", CStr(.PDFCompressionMonoCompressionChoice), REG_SZ
56160    Set reg = Nothing
56170    Exit Sub
56180   End If
56190   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESAMPLE" Then
56200    If Not reg.KeyExists Then
56210     reg.CreateKey
56220    End If
56230    reg.SetRegistryValue "PDFCompressionMonoResample", CStr(Abs(.PDFCompressionMonoResample)), REG_SZ
56240    Set reg = Nothing
56250    Exit Sub
56260   End If
56270   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESAMPLECHOICE" Then
56280    If Not reg.KeyExists Then
56290     reg.CreateKey
56300    End If
56310    reg.SetRegistryValue "PDFCompressionMonoResampleChoice", CStr(.PDFCompressionMonoResampleChoice), REG_SZ
56320    Set reg = Nothing
56330    Exit Sub
56340   End If
56350   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESOLUTION" Then
56360    If Not reg.KeyExists Then
56370     reg.CreateKey
56380    End If
56390    reg.SetRegistryValue "PDFCompressionMonoResolution", CStr(.PDFCompressionMonoResolution), REG_SZ
56400    Set reg = Nothing
56410    Exit Sub
56420   End If
56430   If UCase$(OptionName) = "PDFCOMPRESSIONTEXTCOMPRESSION" Then
56440    If Not reg.KeyExists Then
56450     reg.CreateKey
56460    End If
56470    reg.SetRegistryValue "PDFCompressionTextCompression", CStr(Abs(.PDFCompressionTextCompression)), REG_SZ
56480    Set reg = Nothing
56490    Exit Sub
56500   End If
56510   reg.SubKey = "Printing\Formats\PDF\Fonts"
56520   If UCase$(OptionName) = "PDFFONTSEMBEDALL" Then
56530    If Not reg.KeyExists Then
56540     reg.CreateKey
56550    End If
56560    reg.SetRegistryValue "PDFFontsEmbedAll", CStr(Abs(.PDFFontsEmbedAll)), REG_SZ
56570    Set reg = Nothing
56580    Exit Sub
56590   End If
56600   If UCase$(OptionName) = "PDFFONTSSUBSETFONTS" Then
56610    If Not reg.KeyExists Then
56620     reg.CreateKey
56630    End If
56640    reg.SetRegistryValue "PDFFontsSubSetFonts", CStr(Abs(.PDFFontsSubSetFonts)), REG_SZ
56650    Set reg = Nothing
56660    Exit Sub
56670   End If
56680   If UCase$(OptionName) = "PDFFONTSSUBSETFONTSPERCENT" Then
56690    If Not reg.KeyExists Then
56700     reg.CreateKey
56710    End If
56720    reg.SetRegistryValue "PDFFontsSubSetFontsPercent", CStr(.PDFFontsSubSetFontsPercent), REG_SZ
56730    Set reg = Nothing
56740    Exit Sub
56750   End If
56760   reg.SubKey = "Printing\Formats\PDF\General"
56770   If UCase$(OptionName) = "PDFGENERALASCII85" Then
56780    If Not reg.KeyExists Then
56790     reg.CreateKey
56800    End If
56810    reg.SetRegistryValue "PDFGeneralASCII85", CStr(Abs(.PDFGeneralASCII85)), REG_SZ
56820    Set reg = Nothing
56830    Exit Sub
56840   End If
56850   If UCase$(OptionName) = "PDFGENERALAUTOROTATE" Then
56860    If Not reg.KeyExists Then
56870     reg.CreateKey
56880    End If
56890    reg.SetRegistryValue "PDFGeneralAutorotate", CStr(.PDFGeneralAutorotate), REG_SZ
56900    Set reg = Nothing
56910    Exit Sub
56920   End If
56930   If UCase$(OptionName) = "PDFGENERALCOMPATIBILITY" Then
56940    If Not reg.KeyExists Then
56950     reg.CreateKey
56960    End If
56970    reg.SetRegistryValue "PDFGeneralCompatibility", CStr(.PDFGeneralCompatibility), REG_SZ
56980    Set reg = Nothing
56990    Exit Sub
57000   End If
57010   If UCase$(OptionName) = "PDFGENERALDEFAULT" Then
57020    If Not reg.KeyExists Then
57030     reg.CreateKey
57040    End If
57050    reg.SetRegistryValue "PDFGeneralDefault", CStr(.PDFGeneralDefault), REG_SZ
57060    Set reg = Nothing
57070    Exit Sub
57080   End If
57090   If UCase$(OptionName) = "PDFGENERALOVERPRINT" Then
57100    If Not reg.KeyExists Then
57110     reg.CreateKey
57120    End If
57130    reg.SetRegistryValue "PDFGeneralOverprint", CStr(.PDFGeneralOverprint), REG_SZ
57140    Set reg = Nothing
57150    Exit Sub
57160   End If
57170   If UCase$(OptionName) = "PDFGENERALRESOLUTION" Then
57180    If Not reg.KeyExists Then
57190     reg.CreateKey
57200    End If
57210    reg.SetRegistryValue "PDFGeneralResolution", CStr(.PDFGeneralResolution), REG_SZ
57220    Set reg = Nothing
57230    Exit Sub
57240   End If
57250   If UCase$(OptionName) = "PDFOPTIMIZE" Then
57260    If Not reg.KeyExists Then
57270     reg.CreateKey
57280    End If
57290    reg.SetRegistryValue "PDFOptimize", CStr(Abs(.PDFOptimize)), REG_SZ
57300    Set reg = Nothing
57310    Exit Sub
57320   End If
57330   If UCase$(OptionName) = "PDFUPDATEMETADATA" Then
57340    If Not reg.KeyExists Then
57350     reg.CreateKey
57360    End If
57370    reg.SetRegistryValue "PDFUpdateMetadata", CStr(.PDFUpdateMetadata), REG_SZ
57380    Set reg = Nothing
57390    Exit Sub
57400   End If
57410   reg.SubKey = "Printing\Formats\PDF\Security"
57420   If UCase$(OptionName) = "PDFALLOWASSEMBLY" Then
57430    If Not reg.KeyExists Then
57440     reg.CreateKey
57450    End If
57460    reg.SetRegistryValue "PDFAllowAssembly", CStr(Abs(.PDFAllowAssembly)), REG_SZ
57470    Set reg = Nothing
57480    Exit Sub
57490   End If
57500   If UCase$(OptionName) = "PDFALLOWDEGRADEDPRINTING" Then
57510    If Not reg.KeyExists Then
57520     reg.CreateKey
57530    End If
57540    reg.SetRegistryValue "PDFAllowDegradedPrinting", CStr(Abs(.PDFAllowDegradedPrinting)), REG_SZ
57550    Set reg = Nothing
57560    Exit Sub
57570   End If
57580   If UCase$(OptionName) = "PDFALLOWFILLIN" Then
57590    If Not reg.KeyExists Then
57600     reg.CreateKey
57610    End If
57620    reg.SetRegistryValue "PDFAllowFillIn", CStr(Abs(.PDFAllowFillIn)), REG_SZ
57630    Set reg = Nothing
57640    Exit Sub
57650   End If
57660   If UCase$(OptionName) = "PDFALLOWSCREENREADERS" Then
57670    If Not reg.KeyExists Then
57680     reg.CreateKey
57690    End If
57700    reg.SetRegistryValue "PDFAllowScreenReaders", CStr(Abs(.PDFAllowScreenReaders)), REG_SZ
57710    Set reg = Nothing
57720    Exit Sub
57730   End If
57740   If UCase$(OptionName) = "PDFDISALLOWCOPY" Then
57750    If Not reg.KeyExists Then
57760     reg.CreateKey
57770    End If
57780    reg.SetRegistryValue "PDFDisallowCopy", CStr(Abs(.PDFDisallowCopy)), REG_SZ
57790    Set reg = Nothing
57800    Exit Sub
57810   End If
57820   If UCase$(OptionName) = "PDFDISALLOWMODIFYANNOTATIONS" Then
57830    If Not reg.KeyExists Then
57840     reg.CreateKey
57850    End If
57860    reg.SetRegistryValue "PDFDisallowModifyAnnotations", CStr(Abs(.PDFDisallowModifyAnnotations)), REG_SZ
57870    Set reg = Nothing
57880    Exit Sub
57890   End If
57900   If UCase$(OptionName) = "PDFDISALLOWMODIFYCONTENTS" Then
57910    If Not reg.KeyExists Then
57920     reg.CreateKey
57930    End If
57940    reg.SetRegistryValue "PDFDisallowModifyContents", CStr(Abs(.PDFDisallowModifyContents)), REG_SZ
57950    Set reg = Nothing
57960    Exit Sub
57970   End If
57980   If UCase$(OptionName) = "PDFDISALLOWPRINTING" Then
57990    If Not reg.KeyExists Then
58000     reg.CreateKey
58010    End If
58020    reg.SetRegistryValue "PDFDisallowPrinting", CStr(Abs(.PDFDisallowPrinting)), REG_SZ
58030    Set reg = Nothing
58040    Exit Sub
58050   End If
58060   If UCase$(OptionName) = "PDFENCRYPTOR" Then
58070    If Not reg.KeyExists Then
58080     reg.CreateKey
58090    End If
58100    reg.SetRegistryValue "PDFEncryptor", CStr(.PDFEncryptor), REG_SZ
58110    Set reg = Nothing
58120    Exit Sub
58130   End If
58140   If UCase$(OptionName) = "PDFHIGHENCRYPTION" Then
58150    If Not reg.KeyExists Then
58160     reg.CreateKey
58170    End If
58180    reg.SetRegistryValue "PDFHighEncryption", CStr(Abs(.PDFHighEncryption)), REG_SZ
58190    Set reg = Nothing
58200    Exit Sub
58210   End If
58220   If UCase$(OptionName) = "PDFLOWENCRYPTION" Then
58230    If Not reg.KeyExists Then
58240     reg.CreateKey
58250    End If
58260    reg.SetRegistryValue "PDFLowEncryption", CStr(Abs(.PDFLowEncryption)), REG_SZ
58270    Set reg = Nothing
58280    Exit Sub
58290   End If
58300   If UCase$(OptionName) = "PDFOWNERPASS" Then
58310    If Not reg.KeyExists Then
58320     reg.CreateKey
58330    End If
58340    reg.SetRegistryValue "PDFOwnerPass", CStr(Abs(.PDFOwnerPass)), REG_SZ
58350    Set reg = Nothing
58360    Exit Sub
58370   End If
58380   If UCase$(OptionName) = "PDFOWNERPASSWORDSTRING" Then
58390    If Not reg.KeyExists Then
58400     reg.CreateKey
58410    End If
58420    reg.SetRegistryValue "PDFOwnerPasswordString", CStr(.PDFOwnerPasswordString), REG_SZ
58430    Set reg = Nothing
58440    Exit Sub
58450   End If
58460   If UCase$(OptionName) = "PDFUSERPASS" Then
58470    If Not reg.KeyExists Then
58480     reg.CreateKey
58490    End If
58500    reg.SetRegistryValue "PDFUserPass", CStr(Abs(.PDFUserPass)), REG_SZ
58510    Set reg = Nothing
58520    Exit Sub
58530   End If
58540   If UCase$(OptionName) = "PDFUSERPASSWORDSTRING" Then
58550    If Not reg.KeyExists Then
58560     reg.CreateKey
58570    End If
58580    reg.SetRegistryValue "PDFUserPasswordString", CStr(.PDFUserPasswordString), REG_SZ
58590    Set reg = Nothing
58600    Exit Sub
58610   End If
58620   If UCase$(OptionName) = "PDFUSESECURITY" Then
58630    If Not reg.KeyExists Then
58640     reg.CreateKey
58650    End If
58660    reg.SetRegistryValue "PDFUseSecurity", CStr(Abs(.PDFUseSecurity)), REG_SZ
58670    Set reg = Nothing
58680    Exit Sub
58690   End If
58700   reg.SubKey = "Printing\Formats\PDF\Signing"
58710   If UCase$(OptionName) = "PDFSIGNINGMULTISIGNATURE" Then
58720    If Not reg.KeyExists Then
58730     reg.CreateKey
58740    End If
58750    reg.SetRegistryValue "PDFSigningMultiSignature", CStr(Abs(.PDFSigningMultiSignature)), REG_SZ
58760    Set reg = Nothing
58770    Exit Sub
58780   End If
58790   If UCase$(OptionName) = "PDFSIGNINGPFXFILE" Then
58800    If Not reg.KeyExists Then
58810     reg.CreateKey
58820    End If
58830    reg.SetRegistryValue "PDFSigningPFXFile", CStr(.PDFSigningPFXFile), REG_SZ
58840    Set reg = Nothing
58850    Exit Sub
58860   End If
58870   If UCase$(OptionName) = "PDFSIGNINGPFXFILEPASSWORD" Then
58880    If Not reg.KeyExists Then
58890     reg.CreateKey
58900    End If
58910    reg.SetRegistryValue "PDFSigningPFXFilePassword", CStr(.PDFSigningPFXFilePassword), REG_SZ
58920    Set reg = Nothing
58930    Exit Sub
58940   End If
58950   If UCase$(OptionName) = "PDFSIGNINGSIGNATURECONTACT" Then
58960    If Not reg.KeyExists Then
58970     reg.CreateKey
58980    End If
58990    reg.SetRegistryValue "PDFSigningSignatureContact", CStr(.PDFSigningSignatureContact), REG_SZ
59000    Set reg = Nothing
59010    Exit Sub
59020   End If
59030   If UCase$(OptionName) = "PDFSIGNINGSIGNATURELEFTX" Then
59040    If Not reg.KeyExists Then
59050     reg.CreateKey
59060    End If
59070   reg.SetRegistryValue "PDFSigningSignatureLeftX", Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), REG_SZ
59080    Set reg = Nothing
59090    Exit Sub
59100   End If
59110   If UCase$(OptionName) = "PDFSIGNINGSIGNATURELEFTY" Then
59120    If Not reg.KeyExists Then
59130     reg.CreateKey
59140    End If
59150   reg.SetRegistryValue "PDFSigningSignatureLeftY", Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), REG_SZ
59160    Set reg = Nothing
59170    Exit Sub
59180   End If
59190   If UCase$(OptionName) = "PDFSIGNINGSIGNATURELOCATION" Then
59200    If Not reg.KeyExists Then
59210     reg.CreateKey
59220    End If
59230    reg.SetRegistryValue "PDFSigningSignatureLocation", CStr(.PDFSigningSignatureLocation), REG_SZ
59240    Set reg = Nothing
59250    Exit Sub
59260   End If
59270   If UCase$(OptionName) = "PDFSIGNINGSIGNATUREONPAGE" Then
59280    If Not reg.KeyExists Then
59290     reg.CreateKey
59300    End If
59310    reg.SetRegistryValue "PDFSigningSignatureOnPage", CStr(.PDFSigningSignatureOnPage), REG_SZ
59320    Set reg = Nothing
59330    Exit Sub
59340   End If
59350   If UCase$(OptionName) = "PDFSIGNINGSIGNATUREREASON" Then
59360    If Not reg.KeyExists Then
59370     reg.CreateKey
59380    End If
59390    reg.SetRegistryValue "PDFSigningSignatureReason", CStr(.PDFSigningSignatureReason), REG_SZ
59400    Set reg = Nothing
59410    Exit Sub
59420   End If
59430   If UCase$(OptionName) = "PDFSIGNINGSIGNATURERIGHTX" Then
59440    If Not reg.KeyExists Then
59450     reg.CreateKey
59460    End If
59470   reg.SetRegistryValue "PDFSigningSignatureRightX", Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), REG_SZ
59480    Set reg = Nothing
59490    Exit Sub
59500   End If
59510   If UCase$(OptionName) = "PDFSIGNINGSIGNATURERIGHTY" Then
59520    If Not reg.KeyExists Then
59530     reg.CreateKey
59540    End If
59550   reg.SetRegistryValue "PDFSigningSignatureRightY", Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), REG_SZ
59560    Set reg = Nothing
59570    Exit Sub
59580   End If
59590   If UCase$(OptionName) = "PDFSIGNINGSIGNATUREVISIBLE" Then
59600    If Not reg.KeyExists Then
59610     reg.CreateKey
59620    End If
59630    reg.SetRegistryValue "PDFSigningSignatureVisible", CStr(Abs(.PDFSigningSignatureVisible)), REG_SZ
59640    Set reg = Nothing
59650    Exit Sub
59660   End If
59670   If UCase$(OptionName) = "PDFSIGNINGSIGNPDF" Then
59680    If Not reg.KeyExists Then
59690     reg.CreateKey
59700    End If
59710    reg.SetRegistryValue "PDFSigningSignPDF", CStr(Abs(.PDFSigningSignPDF)), REG_SZ
59720    Set reg = Nothing
59730    Exit Sub
59740   End If
59750   reg.SubKey = "Printing\Formats\PS\LanguageLevel"
59760   If UCase$(OptionName) = "EPSLANGUAGELEVEL" Then
59770    If Not reg.KeyExists Then
59780     reg.CreateKey
59790    End If
59800    reg.SetRegistryValue "EPSLanguageLevel", CStr(.EPSLanguageLevel), REG_SZ
59810    Set reg = Nothing
59820    Exit Sub
59830   End If
59840   If UCase$(OptionName) = "PSLANGUAGELEVEL" Then
59850    If Not reg.KeyExists Then
59860     reg.CreateKey
59870    End If
59880    reg.SetRegistryValue "PSLanguageLevel", CStr(.PSLanguageLevel), REG_SZ
59890    Set reg = Nothing
59900    Exit Sub
59910   End If
59920   reg.SubKey = "Program"
59930   If UCase$(OptionName) = "ADDITIONALGHOSTSCRIPTPARAMETERS" Then
59940    If Not reg.KeyExists Then
59950     reg.CreateKey
59960    End If
59970    reg.SetRegistryValue "AdditionalGhostscriptParameters", CStr(.AdditionalGhostscriptParameters), REG_SZ
59980    Set reg = Nothing
59990    Exit Sub
60000   End If
60010   If UCase$(OptionName) = "ADDITIONALGHOSTSCRIPTSEARCHPATH" Then
60020    If Not reg.KeyExists Then
60030     reg.CreateKey
60040    End If
60050    reg.SetRegistryValue "AdditionalGhostscriptSearchpath", CStr(.AdditionalGhostscriptSearchpath), REG_SZ
60060    Set reg = Nothing
60070    Exit Sub
60080   End If
60090   If UCase$(OptionName) = "ADDWINDOWSFONTPATH" Then
60100    If Not reg.KeyExists Then
60110     reg.CreateKey
60120    End If
60130    reg.SetRegistryValue "AddWindowsFontpath", CStr(Abs(.AddWindowsFontpath)), REG_SZ
60140    Set reg = Nothing
60150    Exit Sub
60160   End If
60170   If UCase$(OptionName) = "ALLOWSPECIALGSCHARSINFILENAMES" Then
60180    If Not reg.KeyExists Then
60190     reg.CreateKey
60200    End If
60210    reg.SetRegistryValue "AllowSpecialGSCharsInFilenames", CStr(Abs(.AllowSpecialGSCharsInFilenames)), REG_SZ
60220    Set reg = Nothing
60230    Exit Sub
60240   End If
60250   If UCase$(OptionName) = "AUTOSAVEDIRECTORY" Then
60260    If Not reg.KeyExists Then
60270     reg.CreateKey
60280    End If
60290    reg.SetRegistryValue "AutosaveDirectory", CStr(.AutosaveDirectory), REG_SZ
60300    Set reg = Nothing
60310    Exit Sub
60320   End If
60330   If UCase$(OptionName) = "AUTOSAVEFILENAME" Then
60340    If Not reg.KeyExists Then
60350     reg.CreateKey
60360    End If
60370    reg.SetRegistryValue "AutosaveFilename", CStr(.AutosaveFilename), REG_SZ
60380    Set reg = Nothing
60390    Exit Sub
60400   End If
60410   If UCase$(OptionName) = "AUTOSAVEFORMAT" Then
60420    If Not reg.KeyExists Then
60430     reg.CreateKey
60440    End If
60450    reg.SetRegistryValue "AutosaveFormat", CStr(.AutosaveFormat), REG_SZ
60460    Set reg = Nothing
60470    Exit Sub
60480   End If
60490   If UCase$(OptionName) = "AUTOSAVESTARTSTANDARDPROGRAM" Then
60500    If Not reg.KeyExists Then
60510     reg.CreateKey
60520    End If
60530    reg.SetRegistryValue "AutosaveStartStandardProgram", CStr(Abs(.AutosaveStartStandardProgram)), REG_SZ
60540    Set reg = Nothing
60550    Exit Sub
60560   End If
60570   If UCase$(OptionName) = "CLIENTCOMPUTERRESOLVEIPADDRESS" Then
60580    If Not reg.KeyExists Then
60590     reg.CreateKey
60600    End If
60610    reg.SetRegistryValue "ClientComputerResolveIPAddress", CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
60620    Set reg = Nothing
60630    Exit Sub
60640   End If
60650   If UCase$(OptionName) = "DISABLEEMAIL" Then
60660    If Not reg.KeyExists Then
60670     reg.CreateKey
60680    End If
60690    reg.SetRegistryValue "DisableEmail", CStr(Abs(.DisableEmail)), REG_SZ
60700    Set reg = Nothing
60710    Exit Sub
60720   End If
60730   If UCase$(OptionName) = "DONTUSEDOCUMENTSETTINGS" Then
60740    If Not reg.KeyExists Then
60750     reg.CreateKey
60760    End If
60770    reg.SetRegistryValue "DontUseDocumentSettings", CStr(Abs(.DontUseDocumentSettings)), REG_SZ
60780    Set reg = Nothing
60790    Exit Sub
60800   End If
60810   If UCase$(OptionName) = "FILENAMESUBSTITUTIONS" Then
60820    If Not reg.KeyExists Then
60830     reg.CreateKey
60840    End If
60850    reg.SetRegistryValue "FilenameSubstitutions", CStr(.FilenameSubstitutions), REG_SZ
60860    Set reg = Nothing
60870    Exit Sub
60880   End If
60890   If UCase$(OptionName) = "FILENAMESUBSTITUTIONSONLYINTITLE" Then
60900    If Not reg.KeyExists Then
60910     reg.CreateKey
60920    End If
60930    reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle", CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
60940    Set reg = Nothing
60950    Exit Sub
60960   End If
60970   If UCase$(OptionName) = "LANGUAGE" Then
60980    If Not reg.KeyExists Then
60990     reg.CreateKey
61000    End If
61010    reg.SetRegistryValue "Language", CStr(.Language), REG_SZ
61020    Set reg = Nothing
61030    Exit Sub
61040   End If
61050   If UCase$(OptionName) = "LASTSAVEDIRECTORY" Then
61060    If Not reg.KeyExists Then
61070     reg.CreateKey
61080    End If
61090    reg.SetRegistryValue "LastSaveDirectory", CStr(.LastSaveDirectory), REG_SZ
61100    Set reg = Nothing
61110    Exit Sub
61120   End If
61130   If UCase$(OptionName) = "LASTUPDATECHECK" Then
61140    If Not reg.KeyExists Then
61150     reg.CreateKey
61160    End If
61170    reg.SetRegistryValue "LastUpdateCheck", CStr(.LastUpdateCheck), REG_SZ
61180    Set reg = Nothing
61190    Exit Sub
61200   End If
61210   If UCase$(OptionName) = "LOGGING" Then
61220    If Not reg.KeyExists Then
61230     reg.CreateKey
61240    End If
61250    reg.SetRegistryValue "Logging", CStr(Abs(.Logging)), REG_SZ
61260    Set reg = Nothing
61270    Exit Sub
61280   End If
61290   If UCase$(OptionName) = "LOGLINES" Then
61300    If Not reg.KeyExists Then
61310     reg.CreateKey
61320    End If
61330    reg.SetRegistryValue "LogLines", CStr(.LogLines), REG_SZ
61340    Set reg = Nothing
61350    Exit Sub
61360   End If
61370   If UCase$(OptionName) = "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER" Then
61380    If Not reg.KeyExists Then
61390     reg.CreateKey
61400    End If
61410    reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter", CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
61420    Set reg = Nothing
61430    Exit Sub
61440   End If
61450   If UCase$(OptionName) = "NOPROCESSINGATSTARTUP" Then
61460    If Not reg.KeyExists Then
61470     reg.CreateKey
61480    End If
61490    reg.SetRegistryValue "NoProcessingAtStartup", CStr(Abs(.NoProcessingAtStartup)), REG_SZ
61500    Set reg = Nothing
61510    Exit Sub
61520   End If
61530   If UCase$(OptionName) = "NOPSCHECK" Then
61540    If Not reg.KeyExists Then
61550     reg.CreateKey
61560    End If
61570    reg.SetRegistryValue "NoPSCheck", CStr(Abs(.NoPSCheck)), REG_SZ
61580    Set reg = Nothing
61590    Exit Sub
61600   End If
61610   If UCase$(OptionName) = "OPTIONSDESIGN" Then
61620    If Not reg.KeyExists Then
61630     reg.CreateKey
61640    End If
61650    reg.SetRegistryValue "OptionsDesign", CStr(.OptionsDesign), REG_SZ
61660    Set reg = Nothing
61670    Exit Sub
61680   End If
61690   If UCase$(OptionName) = "OPTIONSENABLED" Then
61700    If Not reg.KeyExists Then
61710     reg.CreateKey
61720    End If
61730    reg.SetRegistryValue "OptionsEnabled", CStr(Abs(.OptionsEnabled)), REG_SZ
61740    Set reg = Nothing
61750    Exit Sub
61760   End If
61770   If UCase$(OptionName) = "OPTIONSVISIBLE" Then
61780    If Not reg.KeyExists Then
61790     reg.CreateKey
61800    End If
61810    reg.SetRegistryValue "OptionsVisible", CStr(Abs(.OptionsVisible)), REG_SZ
61820    Set reg = Nothing
61830    Exit Sub
61840   End If
61850   If UCase$(OptionName) = "PRINTAFTERSAVING" Then
61860    If Not reg.KeyExists Then
61870     reg.CreateKey
61880    End If
61890    reg.SetRegistryValue "PrintAfterSaving", CStr(Abs(.PrintAfterSaving)), REG_SZ
61900    Set reg = Nothing
61910    Exit Sub
61920   End If
61930   If UCase$(OptionName) = "PRINTAFTERSAVINGBITSPERPIXEL" Then
61940    If Not reg.KeyExists Then
61950     reg.CreateKey
61960    End If
61970    reg.SetRegistryValue "PrintAfterSavingBitsPerPixel", CStr(.PrintAfterSavingBitsPerPixel), REG_SZ
61980    Set reg = Nothing
61990    Exit Sub
62000   End If
62010   If UCase$(OptionName) = "PRINTAFTERSAVINGDUPLEX" Then
62020    If Not reg.KeyExists Then
62030     reg.CreateKey
62040    End If
62050    reg.SetRegistryValue "PrintAfterSavingDuplex", CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
62060    Set reg = Nothing
62070    Exit Sub
62080   End If
62090   If UCase$(OptionName) = "PRINTAFTERSAVINGMAXRESOLUTION" Then
62100    If Not reg.KeyExists Then
62110     reg.CreateKey
62120    End If
62130    reg.SetRegistryValue "PrintAfterSavingMaxResolution", CStr(.PrintAfterSavingMaxResolution), REG_SZ
62140    Set reg = Nothing
62150    Exit Sub
62160   End If
62170   If UCase$(OptionName) = "PRINTAFTERSAVINGMAXRESOLUTIONENABLED" Then
62180    If Not reg.KeyExists Then
62190     reg.CreateKey
62200    End If
62210    reg.SetRegistryValue "PrintAfterSavingMaxResolutionEnabled", CStr(Abs(.PrintAfterSavingMaxResolutionEnabled)), REG_SZ
62220    Set reg = Nothing
62230    Exit Sub
62240   End If
62250   If UCase$(OptionName) = "PRINTAFTERSAVINGNOCANCEL" Then
62260    If Not reg.KeyExists Then
62270     reg.CreateKey
62280    End If
62290    reg.SetRegistryValue "PrintAfterSavingNoCancel", CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
62300    Set reg = Nothing
62310    Exit Sub
62320   End If
62330   If UCase$(OptionName) = "PRINTAFTERSAVINGPRINTER" Then
62340    If Not reg.KeyExists Then
62350     reg.CreateKey
62360    End If
62370    reg.SetRegistryValue "PrintAfterSavingPrinter", CStr(.PrintAfterSavingPrinter), REG_SZ
62380    Set reg = Nothing
62390    Exit Sub
62400   End If
62410   If UCase$(OptionName) = "PRINTAFTERSAVINGQUERYUSER" Then
62420    If Not reg.KeyExists Then
62430     reg.CreateKey
62440    End If
62450    reg.SetRegistryValue "PrintAfterSavingQueryUser", CStr(.PrintAfterSavingQueryUser), REG_SZ
62460    Set reg = Nothing
62470    Exit Sub
62480   End If
62490   If UCase$(OptionName) = "PRINTAFTERSAVINGTUMBLE" Then
62500    If Not reg.KeyExists Then
62510     reg.CreateKey
62520    End If
62530    reg.SetRegistryValue "PrintAfterSavingTumble", CStr(.PrintAfterSavingTumble), REG_SZ
62540    Set reg = Nothing
62550    Exit Sub
62560   End If
62570   If UCase$(OptionName) = "PRINTERSTOP" Then
62580    If Not reg.KeyExists Then
62590     reg.CreateKey
62600    End If
62610    reg.SetRegistryValue "PrinterStop", CStr(Abs(.PrinterStop)), REG_SZ
62620    Set reg = Nothing
62630    Exit Sub
62640   End If
62650   If UCase$(OptionName) = "PRINTERTEMPPATH" Then
62660    If Not reg.KeyExists Then
62670     reg.CreateKey
62680    End If
62690    reg.SetRegistryValue "PrinterTemppath", CStr(.PrinterTemppath), REG_SZ
62700    Set reg = Nothing
62710    Exit Sub
62720   End If
62730   If UCase$(OptionName) = "PROCESSPRIORITY" Then
62740    If Not reg.KeyExists Then
62750     reg.CreateKey
62760    End If
62770    reg.SetRegistryValue "ProcessPriority", CStr(.ProcessPriority), REG_SZ
62780    Set reg = Nothing
62790    Exit Sub
62800   End If
62810   If UCase$(OptionName) = "PROGRAMFONT" Then
62820    If Not reg.KeyExists Then
62830     reg.CreateKey
62840    End If
62850    reg.SetRegistryValue "ProgramFont", CStr(.ProgramFont), REG_SZ
62860    Set reg = Nothing
62870    Exit Sub
62880   End If
62890   If UCase$(OptionName) = "PROGRAMFONTCHARSET" Then
62900    If Not reg.KeyExists Then
62910     reg.CreateKey
62920    End If
62930    reg.SetRegistryValue "ProgramFontCharset", CStr(.ProgramFontCharset), REG_SZ
62940    Set reg = Nothing
62950    Exit Sub
62960   End If
62970   If UCase$(OptionName) = "PROGRAMFONTSIZE" Then
62980    If Not reg.KeyExists Then
62990     reg.CreateKey
63000    End If
63010    reg.SetRegistryValue "ProgramFontSize", CStr(.ProgramFontSize), REG_SZ
63020    Set reg = Nothing
63030    Exit Sub
63040   End If
63050   If UCase$(OptionName) = "REMOVEALLKNOWNFILEEXTENSIONS" Then
63060    If Not reg.KeyExists Then
63070     reg.CreateKey
63080    End If
63090    reg.SetRegistryValue "RemoveAllKnownFileExtensions", CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
63100    Set reg = Nothing
63110    Exit Sub
63120   End If
63130   If UCase$(OptionName) = "REMOVESPACES" Then
63140    If Not reg.KeyExists Then
63150     reg.CreateKey
63160    End If
63170    reg.SetRegistryValue "RemoveSpaces", CStr(Abs(.RemoveSpaces)), REG_SZ
63180    Set reg = Nothing
63190    Exit Sub
63200   End If
63210   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVING" Then
63220    If Not reg.KeyExists Then
63230     reg.CreateKey
63240    End If
63250    reg.SetRegistryValue "RunProgramAfterSaving", CStr(Abs(.RunProgramAfterSaving)), REG_SZ
63260    Set reg = Nothing
63270    Exit Sub
63280   End If
63290   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGPROGRAMNAME" Then
63300    If Not reg.KeyExists Then
63310     reg.CreateKey
63320    End If
63330    reg.SetRegistryValue "RunProgramAfterSavingProgramname", CStr(.RunProgramAfterSavingProgramname), REG_SZ
63340    Set reg = Nothing
63350    Exit Sub
63360   End If
63370   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS" Then
63380    If Not reg.KeyExists Then
63390     reg.CreateKey
63400    End If
63410    reg.SetRegistryValue "RunProgramAfterSavingProgramParameters", CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
63420    Set reg = Nothing
63430    Exit Sub
63440   End If
63450   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGWAITUNTILREADY" Then
63460    If Not reg.KeyExists Then
63470     reg.CreateKey
63480    End If
63490    reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady", CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
63500    Set reg = Nothing
63510    Exit Sub
63520   End If
63530   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGWINDOWSTYLE" Then
63540    If Not reg.KeyExists Then
63550     reg.CreateKey
63560    End If
63570    reg.SetRegistryValue "RunProgramAfterSavingWindowstyle", CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
63580    Set reg = Nothing
63590    Exit Sub
63600   End If
63610   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVING" Then
63620    If Not reg.KeyExists Then
63630     reg.CreateKey
63640    End If
63650    reg.SetRegistryValue "RunProgramBeforeSaving", CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
63660    Set reg = Nothing
63670    Exit Sub
63680   End If
63690   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGPROGRAMNAME" Then
63700    If Not reg.KeyExists Then
63710     reg.CreateKey
63720    End If
63730    reg.SetRegistryValue "RunProgramBeforeSavingProgramname", CStr(.RunProgramBeforeSavingProgramname), REG_SZ
63740    Set reg = Nothing
63750    Exit Sub
63760   End If
63770   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS" Then
63780    If Not reg.KeyExists Then
63790     reg.CreateKey
63800    End If
63810    reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters", CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
63820    Set reg = Nothing
63830    Exit Sub
63840   End If
63850   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGWINDOWSTYLE" Then
63860    If Not reg.KeyExists Then
63870     reg.CreateKey
63880    End If
63890    reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle", CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
63900    Set reg = Nothing
63910    Exit Sub
63920   End If
63930   If UCase$(OptionName) = "SAVEFILENAME" Then
63940    If Not reg.KeyExists Then
63950     reg.CreateKey
63960    End If
63970    reg.SetRegistryValue "SaveFilename", CStr(.SaveFilename), REG_SZ
63980    Set reg = Nothing
63990    Exit Sub
64000   End If
64010   If UCase$(OptionName) = "SENDEMAILAFTERAUTOSAVING" Then
64020    If Not reg.KeyExists Then
64030     reg.CreateKey
64040    End If
64050    reg.SetRegistryValue "SendEmailAfterAutoSaving", CStr(Abs(.SendEmailAfterAutoSaving)), REG_SZ
64060    Set reg = Nothing
64070    Exit Sub
64080   End If
64090   If UCase$(OptionName) = "SENDMAILMETHOD" Then
64100    If Not reg.KeyExists Then
64110     reg.CreateKey
64120    End If
64130    reg.SetRegistryValue "SendMailMethod", CStr(.SendMailMethod), REG_SZ
64140    Set reg = Nothing
64150    Exit Sub
64160   End If
64170   If UCase$(OptionName) = "SHOWANIMATION" Then
64180    If Not reg.KeyExists Then
64190     reg.CreateKey
64200    End If
64210    reg.SetRegistryValue "ShowAnimation", CStr(Abs(.ShowAnimation)), REG_SZ
64220    Set reg = Nothing
64230    Exit Sub
64240   End If
64250   If UCase$(OptionName) = "STARTSTANDARDPROGRAM" Then
64260    If Not reg.KeyExists Then
64270     reg.CreateKey
64280    End If
64290    reg.SetRegistryValue "StartStandardProgram", CStr(Abs(.StartStandardProgram)), REG_SZ
64300    Set reg = Nothing
64310    Exit Sub
64320   End If
64330   If UCase$(OptionName) = "TOOLBARS" Then
64340    If Not reg.KeyExists Then
64350     reg.CreateKey
64360    End If
64370    reg.SetRegistryValue "Toolbars", CStr(.Toolbars), REG_SZ
64380    Set reg = Nothing
64390    Exit Sub
64400   End If
64410   If UCase$(OptionName) = "UPDATEINTERVAL" Then
64420    If Not reg.KeyExists Then
64430     reg.CreateKey
64440    End If
64450    reg.SetRegistryValue "UpdateInterval", CStr(.UpdateInterval), REG_SZ
64460    Set reg = Nothing
64470    Exit Sub
64480   End If
64490   If UCase$(OptionName) = "USEAUTOSAVE" Then
64500    If Not reg.KeyExists Then
64510     reg.CreateKey
64520    End If
64530    reg.SetRegistryValue "UseAutosave", CStr(Abs(.UseAutosave)), REG_SZ
64540    Set reg = Nothing
64550    Exit Sub
64560   End If
64570   If UCase$(OptionName) = "USEAUTOSAVEDIRECTORY" Then
64580    If Not reg.KeyExists Then
64590     reg.CreateKey
64600    End If
64610    reg.SetRegistryValue "UseAutosaveDirectory", CStr(Abs(.UseAutosaveDirectory)), REG_SZ
64620    Set reg = Nothing
64630    Exit Sub
64640   End If
64650  End With
64660  Set reg = Nothing
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

Public Sub SaveOptionsREG(sOptions As tOptions, Optional hkey1 As hkey = HKEY_CURRENT_USER, Optional ProfileName As String = "")
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = hkey1
50040
50050  ProfileName = Trim$(ProfileName)
50060
50070  If LenB(ProfileName) > 0 Then
50080    reg.KeyRoot = "Software\PDFCreator\Profiles\" & ProfileName
50090   Else
50100    reg.KeyRoot = "Software\PDFCreator"
50110  End If
50120
50130  If Not reg.KeyExists Then
50140   reg.CreateKey
50150  End If
50160  With sOptions
50170   reg.SubKey = "Ghostscript"
50180   If Not reg.KeyExists Then
50190    reg.CreateKey
50200   End If
50210   reg.SetRegistryValue "DirectoryGhostscriptBinaries", CStr(.DirectoryGhostscriptBinaries), REG_SZ
50220   reg.SetRegistryValue "DirectoryGhostscriptFonts", CStr(.DirectoryGhostscriptFonts), REG_SZ
50230   reg.SetRegistryValue "DirectoryGhostscriptLibraries", CStr(.DirectoryGhostscriptLibraries), REG_SZ
50240   reg.SetRegistryValue "DirectoryGhostscriptResource", CStr(.DirectoryGhostscriptResource), REG_SZ
50250   reg.SubKey = "Printing"
50260   If Not reg.KeyExists Then
50270    reg.CreateKey
50280   End If
50290   reg.SetRegistryValue "Counter", CStr(.Counter), REG_SZ
50300   reg.SetRegistryValue "DeviceHeightPoints", Replace$(CStr(.DeviceHeightPoints), GetDecimalChar, "."), REG_SZ
50310   reg.SetRegistryValue "DeviceWidthPoints", Replace$(CStr(.DeviceWidthPoints), GetDecimalChar, "."), REG_SZ
50320   reg.SetRegistryValue "OnePagePerFile", CStr(Abs(.OnePagePerFile)), REG_SZ
50330   reg.SetRegistryValue "Papersize", CStr(.Papersize), REG_SZ
50340   reg.SetRegistryValue "StampFontColor", CStr(.StampFontColor), REG_SZ
50350   reg.SetRegistryValue "StampFontname", CStr(.StampFontname), REG_SZ
50360   reg.SetRegistryValue "StampFontsize", CStr(.StampFontsize), REG_SZ
50370   reg.SetRegistryValue "StampOutlineFontthickness", CStr(.StampOutlineFontthickness), REG_SZ
50380   reg.SetRegistryValue "StampString", CStr(.StampString), REG_SZ
50390   reg.SetRegistryValue "StampUseOutlineFont", CStr(Abs(.StampUseOutlineFont)), REG_SZ
50400   reg.SetRegistryValue "StandardAuthor", CStr(.StandardAuthor), REG_SZ
50410   reg.SetRegistryValue "StandardCreationdate", CStr(.StandardCreationdate), REG_SZ
50420   reg.SetRegistryValue "StandardDateformat", CStr(.StandardDateformat), REG_SZ
50430   reg.SetRegistryValue "StandardKeywords", CStr(.StandardKeywords), REG_SZ
50440   reg.SetRegistryValue "StandardMailDomain", CStr(.StandardMailDomain), REG_SZ
50450   reg.SetRegistryValue "StandardModifydate", CStr(.StandardModifydate), REG_SZ
50460   reg.SetRegistryValue "StandardSaveformat", CStr(.StandardSaveformat), REG_SZ
50470   reg.SetRegistryValue "StandardSubject", CStr(.StandardSubject), REG_SZ
50480   reg.SetRegistryValue "StandardTitle", CStr(.StandardTitle), REG_SZ
50490   reg.SetRegistryValue "UseCreationDateNow", CStr(Abs(.UseCreationDateNow)), REG_SZ
50500   reg.SetRegistryValue "UseCustomPaperSize", CStr(.UseCustomPaperSize), REG_SZ
50510   reg.SetRegistryValue "UseFixPapersize", CStr(Abs(.UseFixPapersize)), REG_SZ
50520   reg.SetRegistryValue "UseStandardAuthor", CStr(Abs(.UseStandardAuthor)), REG_SZ
50530   reg.SubKey = "Printing\Formats\Bitmap\Colors"
50540   If Not reg.KeyExists Then
50550    reg.CreateKey
50560   End If
50570   reg.SetRegistryValue "BMPColorscount", CStr(.BMPColorscount), REG_SZ
50580   reg.SetRegistryValue "BMPResolution", CStr(.BMPResolution), REG_SZ
50590   reg.SetRegistryValue "JPEGColorscount", CStr(.JPEGColorscount), REG_SZ
50600   reg.SetRegistryValue "JPEGQuality", CStr(.JPEGQuality), REG_SZ
50610   reg.SetRegistryValue "JPEGResolution", CStr(.JPEGResolution), REG_SZ
50620   reg.SetRegistryValue "PCLColorsCount", CStr(.PCLColorsCount), REG_SZ
50630   reg.SetRegistryValue "PCLResolution", CStr(.PCLResolution), REG_SZ
50640   reg.SetRegistryValue "PCXColorscount", CStr(.PCXColorscount), REG_SZ
50650   reg.SetRegistryValue "PCXResolution", CStr(.PCXResolution), REG_SZ
50660   reg.SetRegistryValue "PNGColorscount", CStr(.PNGColorscount), REG_SZ
50670   reg.SetRegistryValue "PNGResolution", CStr(.PNGResolution), REG_SZ
50680   reg.SetRegistryValue "PSDColorsCount", CStr(.PSDColorsCount), REG_SZ
50690   reg.SetRegistryValue "PSDResolution", CStr(.PSDResolution), REG_SZ
50700   reg.SetRegistryValue "RAWColorsCount", CStr(.RAWColorsCount), REG_SZ
50710   reg.SetRegistryValue "RAWResolution", CStr(.RAWResolution), REG_SZ
50720   reg.SetRegistryValue "SVGResolution", CStr(.SVGResolution), REG_SZ
50730   reg.SetRegistryValue "TIFFColorscount", CStr(.TIFFColorscount), REG_SZ
50740   reg.SetRegistryValue "TIFFResolution", CStr(.TIFFResolution), REG_SZ
50750   reg.SubKey = "Printing\Formats\PDF\Colors"
50760   If Not reg.KeyExists Then
50770    reg.CreateKey
50780   End If
50790   reg.SetRegistryValue "PDFColorsCMYKToRGB", CStr(Abs(.PDFColorsCMYKToRGB)), REG_SZ
50800   reg.SetRegistryValue "PDFColorsColorModel", CStr(.PDFColorsColorModel), REG_SZ
50810   reg.SetRegistryValue "PDFColorsPreserveHalftone", CStr(Abs(.PDFColorsPreserveHalftone)), REG_SZ
50820   reg.SetRegistryValue "PDFColorsPreserveOverprint", CStr(Abs(.PDFColorsPreserveOverprint)), REG_SZ
50830   reg.SetRegistryValue "PDFColorsPreserveTransfer", CStr(Abs(.PDFColorsPreserveTransfer)), REG_SZ
50840   reg.SubKey = "Printing\Formats\PDF\Compression"
50850   If Not reg.KeyExists Then
50860    reg.CreateKey
50870   End If
50880   reg.SetRegistryValue "PDFCompressionColorCompression", CStr(Abs(.PDFCompressionColorCompression)), REG_SZ
50890   reg.SetRegistryValue "PDFCompressionColorCompressionChoice", CStr(.PDFCompressionColorCompressionChoice), REG_SZ
50900   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
50910   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
50920   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGManualFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGManualFactor), GetDecimalChar, "."), REG_SZ
50930   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
50940   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
50950   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
50960   reg.SetRegistryValue "PDFCompressionColorResample", CStr(Abs(.PDFCompressionColorResample)), REG_SZ
50970   reg.SetRegistryValue "PDFCompressionColorResampleChoice", CStr(.PDFCompressionColorResampleChoice), REG_SZ
50980   reg.SetRegistryValue "PDFCompressionColorResolution", CStr(.PDFCompressionColorResolution), REG_SZ
50990   reg.SetRegistryValue "PDFCompressionGreyCompression", CStr(Abs(.PDFCompressionGreyCompression)), REG_SZ
51000   reg.SetRegistryValue "PDFCompressionGreyCompressionChoice", CStr(.PDFCompressionGreyCompressionChoice), REG_SZ
51010   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
51020   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
51030   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGManualFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGManualFactor), GetDecimalChar, "."), REG_SZ
51040   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
51050   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
51060   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
51070   reg.SetRegistryValue "PDFCompressionGreyResample", CStr(Abs(.PDFCompressionGreyResample)), REG_SZ
51080   reg.SetRegistryValue "PDFCompressionGreyResampleChoice", CStr(.PDFCompressionGreyResampleChoice), REG_SZ
51090   reg.SetRegistryValue "PDFCompressionGreyResolution", CStr(.PDFCompressionGreyResolution), REG_SZ
51100   reg.SetRegistryValue "PDFCompressionMonoCompression", CStr(Abs(.PDFCompressionMonoCompression)), REG_SZ
51110   reg.SetRegistryValue "PDFCompressionMonoCompressionChoice", CStr(.PDFCompressionMonoCompressionChoice), REG_SZ
51120   reg.SetRegistryValue "PDFCompressionMonoResample", CStr(Abs(.PDFCompressionMonoResample)), REG_SZ
51130   reg.SetRegistryValue "PDFCompressionMonoResampleChoice", CStr(.PDFCompressionMonoResampleChoice), REG_SZ
51140   reg.SetRegistryValue "PDFCompressionMonoResolution", CStr(.PDFCompressionMonoResolution), REG_SZ
51150   reg.SetRegistryValue "PDFCompressionTextCompression", CStr(Abs(.PDFCompressionTextCompression)), REG_SZ
51160   reg.SubKey = "Printing\Formats\PDF\Fonts"
51170   If Not reg.KeyExists Then
51180    reg.CreateKey
51190   End If
51200   reg.SetRegistryValue "PDFFontsEmbedAll", CStr(Abs(.PDFFontsEmbedAll)), REG_SZ
51210   reg.SetRegistryValue "PDFFontsSubSetFonts", CStr(Abs(.PDFFontsSubSetFonts)), REG_SZ
51220   reg.SetRegistryValue "PDFFontsSubSetFontsPercent", CStr(.PDFFontsSubSetFontsPercent), REG_SZ
51230   reg.SubKey = "Printing\Formats\PDF\General"
51240   If Not reg.KeyExists Then
51250    reg.CreateKey
51260   End If
51270   reg.SetRegistryValue "PDFGeneralASCII85", CStr(Abs(.PDFGeneralASCII85)), REG_SZ
51280   reg.SetRegistryValue "PDFGeneralAutorotate", CStr(.PDFGeneralAutorotate), REG_SZ
51290   reg.SetRegistryValue "PDFGeneralCompatibility", CStr(.PDFGeneralCompatibility), REG_SZ
51300   reg.SetRegistryValue "PDFGeneralDefault", CStr(.PDFGeneralDefault), REG_SZ
51310   reg.SetRegistryValue "PDFGeneralOverprint", CStr(.PDFGeneralOverprint), REG_SZ
51320   reg.SetRegistryValue "PDFGeneralResolution", CStr(.PDFGeneralResolution), REG_SZ
51330   reg.SetRegistryValue "PDFOptimize", CStr(Abs(.PDFOptimize)), REG_SZ
51340   reg.SetRegistryValue "PDFUpdateMetadata", CStr(.PDFUpdateMetadata), REG_SZ
51350   reg.SubKey = "Printing\Formats\PDF\Security"
51360   If Not reg.KeyExists Then
51370    reg.CreateKey
51380   End If
51390   reg.SetRegistryValue "PDFAllowAssembly", CStr(Abs(.PDFAllowAssembly)), REG_SZ
51400   reg.SetRegistryValue "PDFAllowDegradedPrinting", CStr(Abs(.PDFAllowDegradedPrinting)), REG_SZ
51410   reg.SetRegistryValue "PDFAllowFillIn", CStr(Abs(.PDFAllowFillIn)), REG_SZ
51420   reg.SetRegistryValue "PDFAllowScreenReaders", CStr(Abs(.PDFAllowScreenReaders)), REG_SZ
51430   reg.SetRegistryValue "PDFDisallowCopy", CStr(Abs(.PDFDisallowCopy)), REG_SZ
51440   reg.SetRegistryValue "PDFDisallowModifyAnnotations", CStr(Abs(.PDFDisallowModifyAnnotations)), REG_SZ
51450   reg.SetRegistryValue "PDFDisallowModifyContents", CStr(Abs(.PDFDisallowModifyContents)), REG_SZ
51460   reg.SetRegistryValue "PDFDisallowPrinting", CStr(Abs(.PDFDisallowPrinting)), REG_SZ
51470   reg.SetRegistryValue "PDFEncryptor", CStr(.PDFEncryptor), REG_SZ
51480   reg.SetRegistryValue "PDFHighEncryption", CStr(Abs(.PDFHighEncryption)), REG_SZ
51490   reg.SetRegistryValue "PDFLowEncryption", CStr(Abs(.PDFLowEncryption)), REG_SZ
51500   reg.SetRegistryValue "PDFOwnerPass", CStr(Abs(.PDFOwnerPass)), REG_SZ
51510   reg.SetRegistryValue "PDFOwnerPasswordString", CStr(.PDFOwnerPasswordString), REG_SZ
51520   reg.SetRegistryValue "PDFUserPass", CStr(Abs(.PDFUserPass)), REG_SZ
51530   reg.SetRegistryValue "PDFUserPasswordString", CStr(.PDFUserPasswordString), REG_SZ
51540   reg.SetRegistryValue "PDFUseSecurity", CStr(Abs(.PDFUseSecurity)), REG_SZ
51550   reg.SubKey = "Printing\Formats\PDF\Signing"
51560   If Not reg.KeyExists Then
51570    reg.CreateKey
51580   End If
51590   reg.SetRegistryValue "PDFSigningMultiSignature", CStr(Abs(.PDFSigningMultiSignature)), REG_SZ
51600   reg.SetRegistryValue "PDFSigningPFXFile", CStr(.PDFSigningPFXFile), REG_SZ
51610   reg.SetRegistryValue "PDFSigningPFXFilePassword", CStr(.PDFSigningPFXFilePassword), REG_SZ
51620   reg.SetRegistryValue "PDFSigningSignatureContact", CStr(.PDFSigningSignatureContact), REG_SZ
51630   reg.SetRegistryValue "PDFSigningSignatureLeftX", Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), REG_SZ
51640   reg.SetRegistryValue "PDFSigningSignatureLeftY", Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), REG_SZ
51650   reg.SetRegistryValue "PDFSigningSignatureLocation", CStr(.PDFSigningSignatureLocation), REG_SZ
51660   reg.SetRegistryValue "PDFSigningSignatureOnPage", CStr(.PDFSigningSignatureOnPage), REG_SZ
51670   reg.SetRegistryValue "PDFSigningSignatureReason", CStr(.PDFSigningSignatureReason), REG_SZ
51680   reg.SetRegistryValue "PDFSigningSignatureRightX", Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), REG_SZ
51690   reg.SetRegistryValue "PDFSigningSignatureRightY", Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), REG_SZ
51700   reg.SetRegistryValue "PDFSigningSignatureVisible", CStr(Abs(.PDFSigningSignatureVisible)), REG_SZ
51710   reg.SetRegistryValue "PDFSigningSignPDF", CStr(Abs(.PDFSigningSignPDF)), REG_SZ
51720   reg.SubKey = "Printing\Formats\PS\LanguageLevel"
51730   If Not reg.KeyExists Then
51740    reg.CreateKey
51750   End If
51760   reg.SetRegistryValue "EPSLanguageLevel", CStr(.EPSLanguageLevel), REG_SZ
51770   reg.SetRegistryValue "PSLanguageLevel", CStr(.PSLanguageLevel), REG_SZ
51780   reg.SubKey = "Program"
51790   If Not reg.KeyExists Then
51800    reg.CreateKey
51810   End If
51820   reg.SetRegistryValue "AdditionalGhostscriptParameters", CStr(.AdditionalGhostscriptParameters), REG_SZ
51830   reg.SetRegistryValue "AdditionalGhostscriptSearchpath", CStr(.AdditionalGhostscriptSearchpath), REG_SZ
51840   reg.SetRegistryValue "AddWindowsFontpath", CStr(Abs(.AddWindowsFontpath)), REG_SZ
51850   reg.SetRegistryValue "AllowSpecialGSCharsInFilenames", CStr(Abs(.AllowSpecialGSCharsInFilenames)), REG_SZ
51860   reg.SetRegistryValue "AutosaveDirectory", CStr(.AutosaveDirectory), REG_SZ
51870   reg.SetRegistryValue "AutosaveFilename", CStr(.AutosaveFilename), REG_SZ
51880   reg.SetRegistryValue "AutosaveFormat", CStr(.AutosaveFormat), REG_SZ
51890   reg.SetRegistryValue "AutosaveStartStandardProgram", CStr(Abs(.AutosaveStartStandardProgram)), REG_SZ
51900   reg.SetRegistryValue "ClientComputerResolveIPAddress", CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
51910   reg.SetRegistryValue "DisableEmail", CStr(Abs(.DisableEmail)), REG_SZ
51920   reg.SetRegistryValue "DontUseDocumentSettings", CStr(Abs(.DontUseDocumentSettings)), REG_SZ
51930   reg.SetRegistryValue "FilenameSubstitutions", CStr(.FilenameSubstitutions), REG_SZ
51940   reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle", CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
51950   reg.SetRegistryValue "Language", CStr(.Language), REG_SZ
51960   reg.SetRegistryValue "LastSaveDirectory", CStr(.LastSaveDirectory), REG_SZ
51970   reg.SetRegistryValue "LastUpdateCheck", CStr(.LastUpdateCheck), REG_SZ
51980   reg.SetRegistryValue "Logging", CStr(Abs(.Logging)), REG_SZ
51990   reg.SetRegistryValue "LogLines", CStr(.LogLines), REG_SZ
52000   reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter", CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
52010   reg.SetRegistryValue "NoProcessingAtStartup", CStr(Abs(.NoProcessingAtStartup)), REG_SZ
52020   reg.SetRegistryValue "NoPSCheck", CStr(Abs(.NoPSCheck)), REG_SZ
52030   reg.SetRegistryValue "OptionsDesign", CStr(.OptionsDesign), REG_SZ
52040   reg.SetRegistryValue "OptionsEnabled", CStr(Abs(.OptionsEnabled)), REG_SZ
52050   reg.SetRegistryValue "OptionsVisible", CStr(Abs(.OptionsVisible)), REG_SZ
52060   reg.SetRegistryValue "PrintAfterSaving", CStr(Abs(.PrintAfterSaving)), REG_SZ
52070   reg.SetRegistryValue "PrintAfterSavingBitsPerPixel", CStr(.PrintAfterSavingBitsPerPixel), REG_SZ
52080   reg.SetRegistryValue "PrintAfterSavingDuplex", CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
52090   reg.SetRegistryValue "PrintAfterSavingMaxResolution", CStr(.PrintAfterSavingMaxResolution), REG_SZ
52100   reg.SetRegistryValue "PrintAfterSavingMaxResolutionEnabled", CStr(Abs(.PrintAfterSavingMaxResolutionEnabled)), REG_SZ
52110   reg.SetRegistryValue "PrintAfterSavingNoCancel", CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
52120   reg.SetRegistryValue "PrintAfterSavingPrinter", CStr(.PrintAfterSavingPrinter), REG_SZ
52130   reg.SetRegistryValue "PrintAfterSavingQueryUser", CStr(.PrintAfterSavingQueryUser), REG_SZ
52140   reg.SetRegistryValue "PrintAfterSavingTumble", CStr(.PrintAfterSavingTumble), REG_SZ
52150   reg.SetRegistryValue "PrinterStop", CStr(Abs(.PrinterStop)), REG_SZ
52160   reg.SetRegistryValue "PrinterTemppath", CStr(.PrinterTemppath), REG_SZ
52170   reg.SetRegistryValue "ProcessPriority", CStr(.ProcessPriority), REG_SZ
52180   reg.SetRegistryValue "ProgramFont", CStr(.ProgramFont), REG_SZ
52190   reg.SetRegistryValue "ProgramFontCharset", CStr(.ProgramFontCharset), REG_SZ
52200   reg.SetRegistryValue "ProgramFontSize", CStr(.ProgramFontSize), REG_SZ
52210   reg.SetRegistryValue "RemoveAllKnownFileExtensions", CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
52220   reg.SetRegistryValue "RemoveSpaces", CStr(Abs(.RemoveSpaces)), REG_SZ
52230   reg.SetRegistryValue "RunProgramAfterSaving", CStr(Abs(.RunProgramAfterSaving)), REG_SZ
52240   reg.SetRegistryValue "RunProgramAfterSavingProgramname", CStr(.RunProgramAfterSavingProgramname), REG_SZ
52250   reg.SetRegistryValue "RunProgramAfterSavingProgramParameters", CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
52260   reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady", CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
52270   reg.SetRegistryValue "RunProgramAfterSavingWindowstyle", CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
52280   reg.SetRegistryValue "RunProgramBeforeSaving", CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
52290   reg.SetRegistryValue "RunProgramBeforeSavingProgramname", CStr(.RunProgramBeforeSavingProgramname), REG_SZ
52300   reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters", CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
52310   reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle", CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
52320   reg.SetRegistryValue "SaveFilename", CStr(.SaveFilename), REG_SZ
52330   reg.SetRegistryValue "SendEmailAfterAutoSaving", CStr(Abs(.SendEmailAfterAutoSaving)), REG_SZ
52340   reg.SetRegistryValue "SendMailMethod", CStr(.SendMailMethod), REG_SZ
52350   reg.SetRegistryValue "ShowAnimation", CStr(Abs(.ShowAnimation)), REG_SZ
52360   reg.SetRegistryValue "StartStandardProgram", CStr(Abs(.StartStandardProgram)), REG_SZ
52370   reg.SetRegistryValue "Toolbars", CStr(.Toolbars), REG_SZ
52380   reg.SetRegistryValue "UpdateInterval", CStr(.UpdateInterval), REG_SZ
52390   reg.SetRegistryValue "UseAutosave", CStr(Abs(.UseAutosave)), REG_SZ
52400   reg.SetRegistryValue "UseAutosaveDirectory", CStr(Abs(.UseAutosaveDirectory)), REG_SZ
52410  End With
52420  Set reg = Nothing
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
50020  SaveOption Options, "Language"
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
50010  Dim sLanguage As String, lang As String
50020  If InstalledAsServer Then
50030    sLanguage = ReadLanguageFromOptionsReg(sLanguage, "Software\PDFCreator", HKEY_LOCAL_MACHINE)
50040   Else
50050    If Not IsWin9xMe Then
50060      sLanguage = ReadLanguageFromOptionsReg(sLanguage, ".DEFAULT\Software\PDFCreator", HKEY_USERS)
50070      sLanguage = ReadLanguageFromOptionsReg(sLanguage, "Software\PDFCreator", hProfile, False)
50080     Else
50090      sLanguage = ReadLanguageFromOptionsReg(sLanguage, "Software\PDFCreator", hProfile)
50100    End If
50110    sLanguage = ReadLanguageFromOptionsReg(sLanguage, "Software\PDFCreator", HKEY_LOCAL_MACHINE, False)
50120  End If
50130  SplitPath sLanguage, , , , lang
50140  Options.Language = sLanguage
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
50050   .SubKey = "Program"
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

