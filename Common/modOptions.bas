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
50680   .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
50690   .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
50700   .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
50710   .PDFCompressionColorResample = "0"
50720   .PDFCompressionColorResampleChoice = "0"
50730   .PDFCompressionColorResolution = "300"
50740   .PDFCompressionGreyCompression = "1"
50750   .PDFCompressionGreyCompressionChoice = "0"
50760   .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
50770   .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
50780   .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
50790   .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
50800   .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
50810   .PDFCompressionGreyResample = "0"
50820   .PDFCompressionGreyResampleChoice = "0"
50830   .PDFCompressionGreyResolution = "300"
50840   .PDFCompressionMonoCompression = "1"
50850   .PDFCompressionMonoCompressionChoice = "0"
50860   .PDFCompressionMonoResample = "0"
50870   .PDFCompressionMonoResampleChoice = "0"
50880   .PDFCompressionMonoResolution = "1200"
50890   .PDFCompressionTextCompression = "1"
50900   .PDFDisallowCopy = "1"
50910   .PDFDisallowModifyAnnotations = "0"
50920   .PDFDisallowModifyContents = "0"
50930   .PDFDisallowPrinting = "0"
50940   .PDFEncryptor = "0"
50950   .PDFFontsEmbedAll = "1"
50960   .PDFFontsSubSetFonts = "1"
50970   .PDFFontsSubSetFontsPercent = "100"
50980   .PDFGeneralASCII85 = "0"
50990   .PDFGeneralAutorotate = "2"
51000   .PDFGeneralCompatibility = "2"
51010   .PDFGeneralDefault = "0"
51020   .PDFGeneralOverprint = "0"
51030   .PDFGeneralResolution = "600"
51040   .PDFHighEncryption = "0"
51050   .PDFLowEncryption = "1"
51060   .PDFOptimize = "0"
51070   .PDFOwnerPass = "0"
51080   .PDFOwnerPasswordString = vbNullString
51090   .PDFSigningMultiSignature = "0"
51100   .PDFSigningPFXFile = vbNullString
51110   .PDFSigningPFXFilePassword = vbNullString
51120   .PDFSigningSignatureContact = vbNullString
51130   .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
51140   .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
51150   .PDFSigningSignatureLocation = vbNullString
51160   .PDFSigningSignatureReason = vbNullString
51170   .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
51180   .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
51190   .PDFSigningSignatureVisible = "0"
51200   .PDFSigningSignPDF = "0"
51210   .PDFUpdateMetadata = "1"
51220   .PDFUserPass = "0"
51230   .PDFUserPasswordString = vbNullString
51240   .PDFUseSecurity = "0"
51250   .PNGColorscount = "0"
51260   .PNGResolution = "150"
51270   .PrintAfterSaving = "0"
51280   .PrintAfterSavingDuplex = "0"
51290   .PrintAfterSavingNoCancel = "0"
51300   .PrintAfterSavingPrinter = vbNullString
51310   .PrintAfterSavingQueryUser = "0"
51320   .PrintAfterSavingTumble = "0"
51330   .PrinterStop = "0"
51340   If InstalledAsServer Then
51350     .PrinterTemppath = CompletePath(GetPDFCreatorApplicationPath) & "Temp\"
51360    Else
51370     .PrinterTemppath = "<Temp>PDFCreator\"
51380   End If
51390   .ProcessPriority = "1"
51400   .ProgramFont = "MS Sans Serif"
51410   .ProgramFontCharset = "0"
51420   .ProgramFontSize = "8"
51430   .PSDColorsCount = "0"
51440   .PSDResolution = "150"
51450   .PSLanguageLevel = "2"
51460   .RAWColorsCount = "0"
51470   .RAWResolution = "150"
51480   .RemoveAllKnownFileExtensions = "1"
51490   .RemoveSpaces = "1"
51500   .RunProgramAfterSaving = "0"
51510   .RunProgramAfterSavingProgramname = vbNullString
51520   .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
51530   .RunProgramAfterSavingWaitUntilReady = "1"
51540   .RunProgramAfterSavingWindowstyle = "1"
51550   .RunProgramBeforeSaving = "0"
51560   .RunProgramBeforeSavingProgramname = vbNullString
51570   .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
51580   .RunProgramBeforeSavingWindowstyle = "1"
51590   .SaveFilename = "<Title>"
51600   .SendEmailAfterAutoSaving = "0"
51610   .SendMailMethod = "0"
51620   .ShowAnimation = "1"
51630   .StampFontColor = "#FF0000"
51640   .StampFontname = "Arial"
51650   .StampFontsize = "48"
51660   .StampOutlineFontthickness = "0"
51670   .StampString = vbNullString
51680   .StampUseOutlineFont = "1"
51690   .StandardAuthor = vbNullString
51700   .StandardCreationdate = vbNullString
51710   .StandardDateformat = "YYYYMMDDHHNNSS"
51720   .StandardKeywords = vbNullString
51730   .StandardMailDomain = vbNullString
51740   .StandardModifydate = vbNullString
51750   .StandardSaveformat = "0"
51760   .StandardSubject = vbNullString
51770   .StandardTitle = vbNullString
51780   .StartStandardProgram = "1"
51790   .SVGResolution = "72"
51800   .TIFFColorscount = "0"
51810   .TIFFResolution = "150"
51820   .Toolbars = "1"
51830   .UpdateInterval = "2"
51840   .UseAutosave = "0"
51850   .UseAutosaveDirectory = "1"
51860   .UseCreationDateNow = "0"
51870   .UseCustomPaperSize = "0"
51880   .UseFixPapersize = "0"
51890   .UseStandardAuthor = "0"
51900  End With
51910  StandardOptions = myOptions
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

Public Function ReadOptions(Optional NoMsg As Boolean = False, Optional hProfile As hkey = HKEY_CURRENT_USER, Optional Profilename As String = "") As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim myOptions As tOptions, Str1 As String
50020  Dim tStr As String
50030
50040  Profilename = Trim$(Profilename)
50050  If LenB(Profilename) > 0 Then
50060   tStr = "_" & Profilename
50070  End If
50080
50090  If InstalledAsServer Then
50100    WriteToSpecialLogfile "Reg-Read options: HKEY_LOCAL_MACHINE"
50110    If LenB(Profilename) > 0 Then
50120      myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator\Profiles\" & Profilename, HKEY_LOCAL_MACHINE, NoMsg)
50130     Else
50140      myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", HKEY_LOCAL_MACHINE, NoMsg)
50150    End If
50160   Else
50170    If Not IsWin9xMe Then
50180      WriteToSpecialLogfile "Reg-Read options: HKEY_USERS"
50190      myOptions = ReadOptionsReg(myOptions, ".DEFAULT\Software\PDFCreator", HKEY_USERS, NoMsg)
50200      WriteToSpecialLogfile "Reg-Read options: HKEY_CURRENT_USER [" & hProfile & "]"
50210      If LenB(Profilename) > 0 Then
50220        myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator\Profiles\" & Profilename, hProfile, NoMsg, False)
50230       Else
50240        myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", hProfile, NoMsg, False)
50250      End If
50260     Else
50270      If LenB(Profilename) > 0 Then
50280        myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator\Profiles\" & Profilename, hProfile, NoMsg, False)
50290       Else
50300        myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", hProfile, NoMsg, False)
50310      End If
50320    End If
50330    WriteToSpecialLogfile "Reg-Read options: HKEY_LOCAL_MACHINE"
50340    If LenB(Profilename) > 0 Then
50350      myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator\Profiles\" & Profilename, HKEY_LOCAL_MACHINE, NoMsg, False)
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
56910     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
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
57310   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMaximumFactor")
57320   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57330     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57340       .PDFCompressionColorCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57350      Else
57360       If UseStandard Then
57370        .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57380       End If
57390     End If
57400    Else
57410     If UseStandard Then
57420      .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
57430     End If
57440   End If
57450   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMediumFactor")
57460   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57470     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57480       .PDFCompressionColorCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57490      Else
57500       If UseStandard Then
57510        .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57520       End If
57530     End If
57540    Else
57550     If UseStandard Then
57560      .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57570     End If
57580   End If
57590   tStr = hOpt.Retrieve("PDFCompressionColorCompressionJPEGMinimumFactor")
57600   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57610     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57620       .PDFCompressionColorCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57630      Else
57640       If UseStandard Then
57650        .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57660       End If
57670     End If
57680    Else
57690     If UseStandard Then
57700      .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57710     End If
57720   End If
57730   tStr = hOpt.Retrieve("PDFCompressionColorResample")
57740   If IsNumeric(tStr) Then
57750     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57760       .PDFCompressionColorResample = CLng(tStr)
57770      Else
57780       If UseStandard Then
57790        .PDFCompressionColorResample = 0
57800       End If
57810     End If
57820    Else
57830     If UseStandard Then
57840      .PDFCompressionColorResample = 0
57850     End If
57860   End If
57870   tStr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
57880   If IsNumeric(tStr) Then
57890     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
57900       .PDFCompressionColorResampleChoice = CLng(tStr)
57910      Else
57920       If UseStandard Then
57930        .PDFCompressionColorResampleChoice = 0
57940       End If
57950     End If
57960    Else
57970     If UseStandard Then
57980      .PDFCompressionColorResampleChoice = 0
57990     End If
58000   End If
58010   tStr = hOpt.Retrieve("PDFCompressionColorResolution")
58020   If IsNumeric(tStr) Then
58030     If CLng(tStr) >= 0 Then
58040       .PDFCompressionColorResolution = CLng(tStr)
58050      Else
58060       If UseStandard Then
58070        .PDFCompressionColorResolution = 300
58080       End If
58090     End If
58100    Else
58110     If UseStandard Then
58120      .PDFCompressionColorResolution = 300
58130     End If
58140   End If
58150   tStr = hOpt.Retrieve("PDFCompressionGreyCompression")
58160   If IsNumeric(tStr) Then
58170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58180       .PDFCompressionGreyCompression = CLng(tStr)
58190      Else
58200       If UseStandard Then
58210        .PDFCompressionGreyCompression = 1
58220       End If
58230     End If
58240    Else
58250     If UseStandard Then
58260      .PDFCompressionGreyCompression = 1
58270     End If
58280   End If
58290   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
58300   If IsNumeric(tStr) Then
58310     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
58320       .PDFCompressionGreyCompressionChoice = CLng(tStr)
58330      Else
58340       If UseStandard Then
58350        .PDFCompressionGreyCompressionChoice = 0
58360       End If
58370     End If
58380    Else
58390     If UseStandard Then
58400      .PDFCompressionGreyCompressionChoice = 0
58410     End If
58420   End If
58430   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGHighFactor")
58440   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58450     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58460       .PDFCompressionGreyCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58470      Else
58480       If UseStandard Then
58490        .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
58500       End If
58510     End If
58520    Else
58530     If UseStandard Then
58540      .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
58550     End If
58560   End If
58570   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGLowFactor")
58580   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58590     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58600       .PDFCompressionGreyCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58610      Else
58620       If UseStandard Then
58630        .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58640       End If
58650     End If
58660    Else
58670     If UseStandard Then
58680      .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58690     End If
58700   End If
58710   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMaximumFactor")
58720   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58730     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58740       .PDFCompressionGreyCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58750      Else
58760       If UseStandard Then
58770        .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
58780       End If
58790     End If
58800    Else
58810     If UseStandard Then
58820      .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
58830     End If
58840   End If
58850   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMediumFactor")
58860   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58870     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58880       .PDFCompressionGreyCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58890      Else
58900       If UseStandard Then
58910        .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58920       End If
58930     End If
58940    Else
58950     If UseStandard Then
58960      .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58970     End If
58980   End If
58990   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionJPEGMinimumFactor")
59000   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
59010     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
59020       .PDFCompressionGreyCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
59030      Else
59040       If UseStandard Then
59050        .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
59060       End If
59070     End If
59080    Else
59090     If UseStandard Then
59100      .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
59110     End If
59120   End If
59130   tStr = hOpt.Retrieve("PDFCompressionGreyResample")
59140   If IsNumeric(tStr) Then
59150     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59160       .PDFCompressionGreyResample = CLng(tStr)
59170      Else
59180       If UseStandard Then
59190        .PDFCompressionGreyResample = 0
59200       End If
59210     End If
59220    Else
59230     If UseStandard Then
59240      .PDFCompressionGreyResample = 0
59250     End If
59260   End If
59270   tStr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
59280   If IsNumeric(tStr) Then
59290     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
59300       .PDFCompressionGreyResampleChoice = CLng(tStr)
59310      Else
59320       If UseStandard Then
59330        .PDFCompressionGreyResampleChoice = 0
59340       End If
59350     End If
59360    Else
59370     If UseStandard Then
59380      .PDFCompressionGreyResampleChoice = 0
59390     End If
59400   End If
59410   tStr = hOpt.Retrieve("PDFCompressionGreyResolution")
59420   If IsNumeric(tStr) Then
59430     If CLng(tStr) >= 0 Then
59440       .PDFCompressionGreyResolution = CLng(tStr)
59450      Else
59460       If UseStandard Then
59470        .PDFCompressionGreyResolution = 300
59480       End If
59490     End If
59500    Else
59510     If UseStandard Then
59520      .PDFCompressionGreyResolution = 300
59530     End If
59540   End If
59550   tStr = hOpt.Retrieve("PDFCompressionMonoCompression")
59560   If IsNumeric(tStr) Then
59570     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59580       .PDFCompressionMonoCompression = CLng(tStr)
59590      Else
59600       If UseStandard Then
59610        .PDFCompressionMonoCompression = 1
59620       End If
59630     End If
59640    Else
59650     If UseStandard Then
59660      .PDFCompressionMonoCompression = 1
59670     End If
59680   End If
59690   tStr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
59700   If IsNumeric(tStr) Then
59710     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
59720       .PDFCompressionMonoCompressionChoice = CLng(tStr)
59730      Else
59740       If UseStandard Then
59750        .PDFCompressionMonoCompressionChoice = 0
59760       End If
59770     End If
59780    Else
59790     If UseStandard Then
59800      .PDFCompressionMonoCompressionChoice = 0
59810     End If
59820   End If
59830   tStr = hOpt.Retrieve("PDFCompressionMonoResample")
59840   If IsNumeric(tStr) Then
59850     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59860       .PDFCompressionMonoResample = CLng(tStr)
59870      Else
59880       If UseStandard Then
59890        .PDFCompressionMonoResample = 0
59900       End If
59910     End If
59920    Else
59930     If UseStandard Then
59940      .PDFCompressionMonoResample = 0
59950     End If
59960   End If
59970   tStr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
59980   If IsNumeric(tStr) Then
59990     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
60000       .PDFCompressionMonoResampleChoice = CLng(tStr)
60010      Else
60020       If UseStandard Then
60030        .PDFCompressionMonoResampleChoice = 0
60040       End If
60050     End If
60060    Else
60070     If UseStandard Then
60080      .PDFCompressionMonoResampleChoice = 0
60090     End If
60100   End If
60110   tStr = hOpt.Retrieve("PDFCompressionMonoResolution")
60120   If IsNumeric(tStr) Then
60130     If CLng(tStr) >= 0 Then
60140       .PDFCompressionMonoResolution = CLng(tStr)
60150      Else
60160       If UseStandard Then
60170        .PDFCompressionMonoResolution = 1200
60180       End If
60190     End If
60200    Else
60210     If UseStandard Then
60220      .PDFCompressionMonoResolution = 1200
60230     End If
60240   End If
60250   tStr = hOpt.Retrieve("PDFCompressionTextCompression")
60260   If IsNumeric(tStr) Then
60270     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60280       .PDFCompressionTextCompression = CLng(tStr)
60290      Else
60300       If UseStandard Then
60310        .PDFCompressionTextCompression = 1
60320       End If
60330     End If
60340    Else
60350     If UseStandard Then
60360      .PDFCompressionTextCompression = 1
60370     End If
60380   End If
60390   tStr = hOpt.Retrieve("PDFDisallowCopy")
60400   If IsNumeric(tStr) Then
60410     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60420       .PDFDisallowCopy = CLng(tStr)
60430      Else
60440       If UseStandard Then
60450        .PDFDisallowCopy = 1
60460       End If
60470     End If
60480    Else
60490     If UseStandard Then
60500      .PDFDisallowCopy = 1
60510     End If
60520   End If
60530   tStr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
60540   If IsNumeric(tStr) Then
60550     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60560       .PDFDisallowModifyAnnotations = CLng(tStr)
60570      Else
60580       If UseStandard Then
60590        .PDFDisallowModifyAnnotations = 0
60600       End If
60610     End If
60620    Else
60630     If UseStandard Then
60640      .PDFDisallowModifyAnnotations = 0
60650     End If
60660   End If
60670   tStr = hOpt.Retrieve("PDFDisallowModifyContents")
60680   If IsNumeric(tStr) Then
60690     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60700       .PDFDisallowModifyContents = CLng(tStr)
60710      Else
60720       If UseStandard Then
60730        .PDFDisallowModifyContents = 0
60740       End If
60750     End If
60760    Else
60770     If UseStandard Then
60780      .PDFDisallowModifyContents = 0
60790     End If
60800   End If
60810   tStr = hOpt.Retrieve("PDFDisallowPrinting")
60820   If IsNumeric(tStr) Then
60830     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60840       .PDFDisallowPrinting = CLng(tStr)
60850      Else
60860       If UseStandard Then
60870        .PDFDisallowPrinting = 0
60880       End If
60890     End If
60900    Else
60910     If UseStandard Then
60920      .PDFDisallowPrinting = 0
60930     End If
60940   End If
60950   tStr = hOpt.Retrieve("PDFEncryptor")
60960   If IsNumeric(tStr) Then
60970     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
60980       .PDFEncryptor = CLng(tStr)
60990      Else
61000       If UseStandard Then
61010        .PDFEncryptor = 0
61020       End If
61030     End If
61040    Else
61050     If UseStandard Then
61060      .PDFEncryptor = 0
61070     End If
61080   End If
61090   tStr = hOpt.Retrieve("PDFFontsEmbedAll")
61100   If IsNumeric(tStr) Then
61110     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61120       .PDFFontsEmbedAll = CLng(tStr)
61130      Else
61140       If UseStandard Then
61150        .PDFFontsEmbedAll = 1
61160       End If
61170     End If
61180    Else
61190     If UseStandard Then
61200      .PDFFontsEmbedAll = 1
61210     End If
61220   End If
61230   tStr = hOpt.Retrieve("PDFFontsSubSetFonts")
61240   If IsNumeric(tStr) Then
61250     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61260       .PDFFontsSubSetFonts = CLng(tStr)
61270      Else
61280       If UseStandard Then
61290        .PDFFontsSubSetFonts = 1
61300       End If
61310     End If
61320    Else
61330     If UseStandard Then
61340      .PDFFontsSubSetFonts = 1
61350     End If
61360   End If
61370   tStr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
61380   If IsNumeric(tStr) Then
61390     If CLng(tStr) >= 0 Then
61400       .PDFFontsSubSetFontsPercent = CLng(tStr)
61410      Else
61420       If UseStandard Then
61430        .PDFFontsSubSetFontsPercent = 100
61440       End If
61450     End If
61460    Else
61470     If UseStandard Then
61480      .PDFFontsSubSetFontsPercent = 100
61490     End If
61500   End If
61510   tStr = hOpt.Retrieve("PDFGeneralASCII85")
61520   If IsNumeric(tStr) Then
61530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61540       .PDFGeneralASCII85 = CLng(tStr)
61550      Else
61560       If UseStandard Then
61570        .PDFGeneralASCII85 = 0
61580       End If
61590     End If
61600    Else
61610     If UseStandard Then
61620      .PDFGeneralASCII85 = 0
61630     End If
61640   End If
61650   tStr = hOpt.Retrieve("PDFGeneralAutorotate")
61660   If IsNumeric(tStr) Then
61670     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
61680       .PDFGeneralAutorotate = CLng(tStr)
61690      Else
61700       If UseStandard Then
61710        .PDFGeneralAutorotate = 2
61720       End If
61730     End If
61740    Else
61750     If UseStandard Then
61760      .PDFGeneralAutorotate = 2
61770     End If
61780   End If
61790   tStr = hOpt.Retrieve("PDFGeneralCompatibility")
61800   If IsNumeric(tStr) Then
61810     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
61820       .PDFGeneralCompatibility = CLng(tStr)
61830      Else
61840       If UseStandard Then
61850        .PDFGeneralCompatibility = 2
61860       End If
61870     End If
61880    Else
61890     If UseStandard Then
61900      .PDFGeneralCompatibility = 2
61910     End If
61920   End If
61930   tStr = hOpt.Retrieve("PDFGeneralDefault")
61940   If IsNumeric(tStr) Then
61950     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
61960       .PDFGeneralDefault = CLng(tStr)
61970      Else
61980       If UseStandard Then
61990        .PDFGeneralDefault = 0
62000       End If
62010     End If
62020    Else
62030     If UseStandard Then
62040      .PDFGeneralDefault = 0
62050     End If
62060   End If
62070   tStr = hOpt.Retrieve("PDFGeneralOverprint")
62080   If IsNumeric(tStr) Then
62090     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
62100       .PDFGeneralOverprint = CLng(tStr)
62110      Else
62120       If UseStandard Then
62130        .PDFGeneralOverprint = 0
62140       End If
62150     End If
62160    Else
62170     If UseStandard Then
62180      .PDFGeneralOverprint = 0
62190     End If
62200   End If
62210   tStr = hOpt.Retrieve("PDFGeneralResolution")
62220   If IsNumeric(tStr) Then
62230     If CLng(tStr) >= 0 Then
62240       .PDFGeneralResolution = CLng(tStr)
62250      Else
62260       If UseStandard Then
62270        .PDFGeneralResolution = 600
62280       End If
62290     End If
62300    Else
62310     If UseStandard Then
62320      .PDFGeneralResolution = 600
62330     End If
62340   End If
62350   tStr = hOpt.Retrieve("PDFHighEncryption")
62360   If IsNumeric(tStr) Then
62370     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62380       .PDFHighEncryption = CLng(tStr)
62390      Else
62400       If UseStandard Then
62410        .PDFHighEncryption = 0
62420       End If
62430     End If
62440    Else
62450     If UseStandard Then
62460      .PDFHighEncryption = 0
62470     End If
62480   End If
62490   tStr = hOpt.Retrieve("PDFLowEncryption")
62500   If IsNumeric(tStr) Then
62510     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62520       .PDFLowEncryption = CLng(tStr)
62530      Else
62540       If UseStandard Then
62550        .PDFLowEncryption = 1
62560       End If
62570     End If
62580    Else
62590     If UseStandard Then
62600      .PDFLowEncryption = 1
62610     End If
62620   End If
62630   tStr = hOpt.Retrieve("PDFOptimize")
62640   If IsNumeric(tStr) Then
62650     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62660       .PDFOptimize = CLng(tStr)
62670      Else
62680       If UseStandard Then
62690        .PDFOptimize = 0
62700       End If
62710     End If
62720    Else
62730     If UseStandard Then
62740      .PDFOptimize = 0
62750     End If
62760   End If
62770   tStr = hOpt.Retrieve("PDFOwnerPass")
62780   If IsNumeric(tStr) Then
62790     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62800       .PDFOwnerPass = CLng(tStr)
62810      Else
62820       If UseStandard Then
62830        .PDFOwnerPass = 0
62840       End If
62850     End If
62860    Else
62870     If UseStandard Then
62880      .PDFOwnerPass = 0
62890     End If
62900   End If
62910   tStr = hOpt.Retrieve("PDFOwnerPasswordString")
62920   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
62930     .PDFOwnerPasswordString = ""
62940    Else
62950     If LenB(tStr) > 0 Then
62960      .PDFOwnerPasswordString = tStr
62970     End If
62980   End If
62990   tStr = hOpt.Retrieve("PDFSigningMultiSignature")
63000   If IsNumeric(tStr) Then
63010     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63020       .PDFSigningMultiSignature = CLng(tStr)
63030      Else
63040       If UseStandard Then
63050        .PDFSigningMultiSignature = 0
63060       End If
63070     End If
63080    Else
63090     If UseStandard Then
63100      .PDFSigningMultiSignature = 0
63110     End If
63120   End If
63130   tStr = hOpt.Retrieve("PDFSigningPFXFile")
63140   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63150     .PDFSigningPFXFile = ""
63160    Else
63170     If LenB(tStr) > 0 Then
63180      .PDFSigningPFXFile = tStr
63190     End If
63200   End If
63210   tStr = hOpt.Retrieve("PDFSigningPFXFilePassword")
63220   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63230     .PDFSigningPFXFilePassword = ""
63240    Else
63250     If LenB(tStr) > 0 Then
63260      .PDFSigningPFXFilePassword = tStr
63270     End If
63280   End If
63290   tStr = hOpt.Retrieve("PDFSigningSignatureContact")
63300   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63310     .PDFSigningSignatureContact = ""
63320    Else
63330     If LenB(tStr) > 0 Then
63340      .PDFSigningSignatureContact = tStr
63350     End If
63360   End If
63370   tStr = hOpt.Retrieve("PDFSigningSignatureLeftX")
63380   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63390     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63400       .PDFSigningSignatureLeftX = CDbl(Replace$(tStr, ".", GetDecimalChar))
63410      Else
63420       If UseStandard Then
63430        .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
63440       End If
63450     End If
63460    Else
63470     If UseStandard Then
63480      .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
63490     End If
63500   End If
63510   tStr = hOpt.Retrieve("PDFSigningSignatureLeftY")
63520   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63530     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63540       .PDFSigningSignatureLeftY = CDbl(Replace$(tStr, ".", GetDecimalChar))
63550      Else
63560       If UseStandard Then
63570        .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
63580       End If
63590     End If
63600    Else
63610     If UseStandard Then
63620      .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
63630     End If
63640   End If
63650   tStr = hOpt.Retrieve("PDFSigningSignatureLocation")
63660   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63670     .PDFSigningSignatureLocation = ""
63680    Else
63690     If LenB(tStr) > 0 Then
63700      .PDFSigningSignatureLocation = tStr
63710     End If
63720   End If
63730   tStr = hOpt.Retrieve("PDFSigningSignatureReason")
63740   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63750     .PDFSigningSignatureReason = ""
63760    Else
63770     If LenB(tStr) > 0 Then
63780      .PDFSigningSignatureReason = tStr
63790     End If
63800   End If
63810   tStr = hOpt.Retrieve("PDFSigningSignatureRightX")
63820   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63830     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63840       .PDFSigningSignatureRightX = CDbl(Replace$(tStr, ".", GetDecimalChar))
63850      Else
63860       If UseStandard Then
63870        .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
63880       End If
63890     End If
63900    Else
63910     If UseStandard Then
63920      .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
63930     End If
63940   End If
63950   tStr = hOpt.Retrieve("PDFSigningSignatureRightY")
63960   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
63970     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
63980       .PDFSigningSignatureRightY = CDbl(Replace$(tStr, ".", GetDecimalChar))
63990      Else
64000       If UseStandard Then
64010        .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
64020       End If
64030     End If
64040    Else
64050     If UseStandard Then
64060      .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
64070     End If
64080   End If
64090   tStr = hOpt.Retrieve("PDFSigningSignatureVisible")
64100   If IsNumeric(tStr) Then
64110     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64120       .PDFSigningSignatureVisible = CLng(tStr)
64130      Else
64140       If UseStandard Then
64150        .PDFSigningSignatureVisible = 0
64160       End If
64170     End If
64180    Else
64190     If UseStandard Then
64200      .PDFSigningSignatureVisible = 0
64210     End If
64220   End If
64230   tStr = hOpt.Retrieve("PDFSigningSignPDF")
64240   If IsNumeric(tStr) Then
64250     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64260       .PDFSigningSignPDF = CLng(tStr)
64270      Else
64280       If UseStandard Then
64290        .PDFSigningSignPDF = 0
64300       End If
64310     End If
64320    Else
64330     If UseStandard Then
64340      .PDFSigningSignPDF = 0
64350     End If
64360   End If
64370   tStr = hOpt.Retrieve("PDFUpdateMetadata")
64380   If IsNumeric(tStr) Then
64390     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
64400       .PDFUpdateMetadata = CLng(tStr)
64410      Else
64420       If UseStandard Then
64430        .PDFUpdateMetadata = 1
64440       End If
64450     End If
64460    Else
64470     If UseStandard Then
64480      .PDFUpdateMetadata = 1
64490     End If
64500   End If
64510   tStr = hOpt.Retrieve("PDFUserPass")
64520   If IsNumeric(tStr) Then
64530     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64540       .PDFUserPass = CLng(tStr)
64550      Else
64560       If UseStandard Then
64570        .PDFUserPass = 0
64580       End If
64590     End If
64600    Else
64610     If UseStandard Then
64620      .PDFUserPass = 0
64630     End If
64640   End If
64650   tStr = hOpt.Retrieve("PDFUserPasswordString")
64660   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64670     .PDFUserPasswordString = ""
64680    Else
64690     If LenB(tStr) > 0 Then
64700      .PDFUserPasswordString = tStr
64710     End If
64720   End If
64730   tStr = hOpt.Retrieve("PDFUseSecurity")
64740   If IsNumeric(tStr) Then
64750     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64760       .PDFUseSecurity = CLng(tStr)
64770      Else
64780       If UseStandard Then
64790        .PDFUseSecurity = 0
64800       End If
64810     End If
64820    Else
64830     If UseStandard Then
64840      .PDFUseSecurity = 0
64850     End If
64860   End If
64870   tStr = hOpt.Retrieve("PNGColorscount")
64880   If IsNumeric(tStr) Then
64890     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
64900       .PNGColorscount = CLng(tStr)
64910      Else
64920       If UseStandard Then
64930        .PNGColorscount = 0
64940       End If
64950     End If
64960    Else
64970     If UseStandard Then
64980      .PNGColorscount = 0
64990     End If
65000   End If
65010   tStr = hOpt.Retrieve("PNGResolution")
65020   If IsNumeric(tStr) Then
65030     If CLng(tStr) >= 1 Then
65040       .PNGResolution = CLng(tStr)
65050      Else
65060       If UseStandard Then
65070        .PNGResolution = 150
65080       End If
65090     End If
65100    Else
65110     If UseStandard Then
65120      .PNGResolution = 150
65130     End If
65140   End If
65150   tStr = hOpt.Retrieve("PrintAfterSaving")
65160   If IsNumeric(tStr) Then
65170     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65180       .PrintAfterSaving = CLng(tStr)
65190      Else
65200       If UseStandard Then
65210        .PrintAfterSaving = 0
65220       End If
65230     End If
65240    Else
65250     If UseStandard Then
65260      .PrintAfterSaving = 0
65270     End If
65280   End If
65290   tStr = hOpt.Retrieve("PrintAfterSavingDuplex")
65300   If IsNumeric(tStr) Then
65310     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65320       .PrintAfterSavingDuplex = CLng(tStr)
65330      Else
65340       If UseStandard Then
65350        .PrintAfterSavingDuplex = 0
65360       End If
65370     End If
65380    Else
65390     If UseStandard Then
65400      .PrintAfterSavingDuplex = 0
65410     End If
65420   End If
65430   tStr = hOpt.Retrieve("PrintAfterSavingNoCancel")
65440   If IsNumeric(tStr) Then
65450     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65460       .PrintAfterSavingNoCancel = CLng(tStr)
65470      Else
65480       If UseStandard Then
65490        .PrintAfterSavingNoCancel = 0
65500       End If
65510     End If
65520    Else
65530     If UseStandard Then
65540      .PrintAfterSavingNoCancel = 0
65550     End If
65560   End If
65570   tStr = hOpt.Retrieve("PrintAfterSavingPrinter")
65580   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
65590     .PrintAfterSavingPrinter = ""
65600    Else
65610     If LenB(tStr) > 0 Then
65620      .PrintAfterSavingPrinter = tStr
65630     End If
65640   End If
65650   tStr = hOpt.Retrieve("PrintAfterSavingQueryUser")
65660   If IsNumeric(tStr) Then
65670     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
65680       .PrintAfterSavingQueryUser = CLng(tStr)
65690      Else
65700       If UseStandard Then
65710        .PrintAfterSavingQueryUser = 0
65720       End If
65730     End If
65740    Else
65750     If UseStandard Then
65760      .PrintAfterSavingQueryUser = 0
65770     End If
65780   End If
65790   tStr = hOpt.Retrieve("PrintAfterSavingTumble")
65800   If IsNumeric(tStr) Then
65810     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
65820       .PrintAfterSavingTumble = CLng(tStr)
65830      Else
65840       If UseStandard Then
65850        .PrintAfterSavingTumble = 0
65860       End If
65870     End If
65880    Else
65890     If UseStandard Then
65900      .PrintAfterSavingTumble = 0
65910     End If
65920   End If
65930   tStr = hOpt.Retrieve("PrinterStop")
65940   If IsNumeric(tStr) Then
65950     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65960       .PrinterStop = CLng(tStr)
65970      Else
65980       If UseStandard Then
65990        .PrinterStop = 0
66000       End If
66010     End If
66020    Else
66030     If UseStandard Then
66040      .PrinterStop = 0
66050     End If
66060   End If
66070   tStr = hOpt.Retrieve("PrinterTemppath")
66080   WriteToSpecialLogfile "hOpt.Retrieve(""PrinterTemppath"")=" & tStr
66090   WriteToSpecialLogfile "Options.PrinterTemppath1=" & .PrinterTemppath
66100   If hkey1 = HKEY_USERS Then
66110     If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then
66120       .PrinterTemppath = tStr
66130      Else
66140       If UseStandard Then
66150         .PrinterTemppath = GetTempPath
66160        Else
66170         .PrinterTemppath = tStr
66180       End If
66190     End If
66200    Else
66210     If LenB(Trim$(tStr)) > 0 Then
66220      If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then
66230        .PrinterTemppath = tStr
66240       Else
66250        MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))
66260        If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then
66270          If UseStandard Then
66280            .PrinterTemppath = GetTempPath
66290           Else
66300            .PrinterTemppath = ""
66310            If NoMsg = False Then
66320             MsgBox "PrinterTemppath: '" & tStr & "' = '" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & "'" & _
             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
66340            End If
66350          End If
66360         Else
66370          .PrinterTemppath = tStr
66380        End If
66390      End If
66400     End If
66410   End If
66420   WriteToSpecialLogfile "Options.PrinterTemppath2=" & .PrinterTemppath
66430   tStr = hOpt.Retrieve("ProcessPriority")
66440   If IsNumeric(tStr) Then
66450     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
66460       .ProcessPriority = CLng(tStr)
66470      Else
66480       If UseStandard Then
66490        .ProcessPriority = 1
66500       End If
66510     End If
66520    Else
66530     If UseStandard Then
66540      .ProcessPriority = 1
66550     End If
66560   End If
66570   tStr = hOpt.Retrieve("ProgramFont")
66580   If LenB(tStr) = 0 And LenB("MS Sans Serif") > 0 And UseStandard Then
66590     .ProgramFont = "MS Sans Serif"
66600    Else
66610     If LenB(tStr) > 0 Then
66620      .ProgramFont = tStr
66630     End If
66640   End If
66650   tStr = hOpt.Retrieve("ProgramFontCharset")
66660   If IsNumeric(tStr) Then
66670     If CLng(tStr) >= 0 Then
66680       .ProgramFontCharset = CLng(tStr)
66690      Else
66700       If UseStandard Then
66710        .ProgramFontCharset = 0
66720       End If
66730     End If
66740    Else
66750     If UseStandard Then
66760      .ProgramFontCharset = 0
66770     End If
66780   End If
66790   tStr = hOpt.Retrieve("ProgramFontSize")
66800   If IsNumeric(tStr) Then
66810     If CLng(tStr) >= 6 And CLng(tStr) <= 72 Then
66820       .ProgramFontSize = CLng(tStr)
66830      Else
66840       If UseStandard Then
66850        .ProgramFontSize = 8
66860       End If
66870     End If
66880    Else
66890     If UseStandard Then
66900      .ProgramFontSize = 8
66910     End If
66920   End If
66930   tStr = hOpt.Retrieve("PSDColorsCount")
66940   If IsNumeric(tStr) Then
66950     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
66960       .PSDColorsCount = CLng(tStr)
66970      Else
66980       If UseStandard Then
66990        .PSDColorsCount = 0
67000       End If
67010     End If
67020    Else
67030     If UseStandard Then
67040      .PSDColorsCount = 0
67050     End If
67060   End If
67070   tStr = hOpt.Retrieve("PSDResolution")
67080   If IsNumeric(tStr) Then
67090     If CLng(tStr) >= 1 Then
67100       .PSDResolution = CLng(tStr)
67110      Else
67120       If UseStandard Then
67130        .PSDResolution = 150
67140       End If
67150     End If
67160    Else
67170     If UseStandard Then
67180      .PSDResolution = 150
67190     End If
67200   End If
67210   tStr = hOpt.Retrieve("PSLanguageLevel")
67220   If IsNumeric(tStr) Then
67230     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
67240       .PSLanguageLevel = CLng(tStr)
67250      Else
67260       If UseStandard Then
67270        .PSLanguageLevel = 2
67280       End If
67290     End If
67300    Else
67310     If UseStandard Then
67320      .PSLanguageLevel = 2
67330     End If
67340   End If
67350   tStr = hOpt.Retrieve("RAWColorsCount")
67360   If IsNumeric(tStr) Then
67370     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
67380       .RAWColorsCount = CLng(tStr)
67390      Else
67400       If UseStandard Then
67410        .RAWColorsCount = 0
67420       End If
67430     End If
67440    Else
67450     If UseStandard Then
67460      .RAWColorsCount = 0
67470     End If
67480   End If
67490   tStr = hOpt.Retrieve("RAWResolution")
67500   If IsNumeric(tStr) Then
67510     If CLng(tStr) >= 1 Then
67520       .RAWResolution = CLng(tStr)
67530      Else
67540       If UseStandard Then
67550        .RAWResolution = 150
67560       End If
67570     End If
67580    Else
67590     If UseStandard Then
67600      .RAWResolution = 150
67610     End If
67620   End If
67630   tStr = hOpt.Retrieve("RemoveAllKnownFileExtensions")
67640   If IsNumeric(tStr) Then
67650     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67660       .RemoveAllKnownFileExtensions = CLng(tStr)
67670      Else
67680       If UseStandard Then
67690        .RemoveAllKnownFileExtensions = 1
67700       End If
67710     End If
67720    Else
67730     If UseStandard Then
67740      .RemoveAllKnownFileExtensions = 1
67750     End If
67760   End If
67770   tStr = hOpt.Retrieve("RemoveSpaces")
67780   If IsNumeric(tStr) Then
67790     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67800       .RemoveSpaces = CLng(tStr)
67810      Else
67820       If UseStandard Then
67830        .RemoveSpaces = 1
67840       End If
67850     End If
67860    Else
67870     If UseStandard Then
67880      .RemoveSpaces = 1
67890     End If
67900   End If
67910   tStr = hOpt.Retrieve("RunProgramAfterSaving")
67920   If IsNumeric(tStr) Then
67930     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67940       .RunProgramAfterSaving = CLng(tStr)
67950      Else
67960       If UseStandard Then
67970        .RunProgramAfterSaving = 0
67980       End If
67990     End If
68000    Else
68010     If UseStandard Then
68020      .RunProgramAfterSaving = 0
68030     End If
68040   End If
68050   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramname")
68060   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
68070     .RunProgramAfterSavingProgramname = ""
68080    Else
68090     If LenB(tStr) > 0 Then
68100      .RunProgramAfterSavingProgramname = tStr
68110     End If
68120   End If
68130   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramParameters")
68140   If LenB(tStr) = 0 And LenB("""<OutputFilename>""") > 0 And UseStandard Then
68150     .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
68160    Else
68170     If LenB(tStr) > 0 Then
68180      .RunProgramAfterSavingProgramParameters = tStr
68190     End If
68200   End If
68210   tStr = hOpt.Retrieve("RunProgramAfterSavingWaitUntilReady")
68220   If IsNumeric(tStr) Then
68230     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68240       .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
68250      Else
68260       If UseStandard Then
68270        .RunProgramAfterSavingWaitUntilReady = 1
68280       End If
68290     End If
68300    Else
68310     If UseStandard Then
68320      .RunProgramAfterSavingWaitUntilReady = 1
68330     End If
68340   End If
68350   tStr = hOpt.Retrieve("RunProgramAfterSavingWindowstyle")
68360   If IsNumeric(tStr) Then
68370     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
68380       .RunProgramAfterSavingWindowstyle = CLng(tStr)
68390      Else
68400       If UseStandard Then
68410        .RunProgramAfterSavingWindowstyle = 1
68420       End If
68430     End If
68440    Else
68450     If UseStandard Then
68460      .RunProgramAfterSavingWindowstyle = 1
68470     End If
68480   End If
68490   tStr = hOpt.Retrieve("RunProgramBeforeSaving")
68500   If IsNumeric(tStr) Then
68510     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68520       .RunProgramBeforeSaving = CLng(tStr)
68530      Else
68540       If UseStandard Then
68550        .RunProgramBeforeSaving = 0
68560       End If
68570     End If
68580    Else
68590     If UseStandard Then
68600      .RunProgramBeforeSaving = 0
68610     End If
68620   End If
68630   tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramname")
68640   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
68650     .RunProgramBeforeSavingProgramname = ""
68660    Else
68670     If LenB(tStr) > 0 Then
68680      .RunProgramBeforeSavingProgramname = tStr
68690     End If
68700   End If
68710   tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramParameters")
68720   If LenB(tStr) = 0 And LenB("""<TempFilename>""") > 0 And UseStandard Then
68730     .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
68740    Else
68750     If LenB(tStr) > 0 Then
68760      .RunProgramBeforeSavingProgramParameters = tStr
68770     End If
68780   End If
68790   tStr = hOpt.Retrieve("RunProgramBeforeSavingWindowstyle")
68800   If IsNumeric(tStr) Then
68810     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
68820       .RunProgramBeforeSavingWindowstyle = CLng(tStr)
68830      Else
68840       If UseStandard Then
68850        .RunProgramBeforeSavingWindowstyle = 1
68860       End If
68870     End If
68880    Else
68890     If UseStandard Then
68900      .RunProgramBeforeSavingWindowstyle = 1
68910     End If
68920   End If
68930   tStr = hOpt.Retrieve("SaveFilename")
68940   If LenB(tStr) = 0 And LenB("<Title>") > 0 And UseStandard Then
68950     .SaveFilename = "<Title>"
68960    Else
68970     If LenB(tStr) > 0 Then
68980      .SaveFilename = tStr
68990     End If
69000   End If
69010   tStr = hOpt.Retrieve("SendEmailAfterAutoSaving")
69020   If IsNumeric(tStr) Then
69030     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69040       .SendEmailAfterAutoSaving = CLng(tStr)
69050      Else
69060       If UseStandard Then
69070        .SendEmailAfterAutoSaving = 0
69080       End If
69090     End If
69100    Else
69110     If UseStandard Then
69120      .SendEmailAfterAutoSaving = 0
69130     End If
69140   End If
69150   tStr = hOpt.Retrieve("SendMailMethod")
69160   If IsNumeric(tStr) Then
69170     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
69180       .SendMailMethod = CLng(tStr)
69190      Else
69200       If UseStandard Then
69210        .SendMailMethod = 0
69220       End If
69230     End If
69240    Else
69250     If UseStandard Then
69260      .SendMailMethod = 0
69270     End If
69280   End If
69290   tStr = hOpt.Retrieve("ShowAnimation")
69300   If IsNumeric(tStr) Then
69310     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69320       .ShowAnimation = CLng(tStr)
69330      Else
69340       If UseStandard Then
69350        .ShowAnimation = 1
69360       End If
69370     End If
69380    Else
69390     If UseStandard Then
69400      .ShowAnimation = 1
69410     End If
69420   End If
69430   tStr = hOpt.Retrieve("StampFontColor")
69440   If LenB(tStr) = 0 And LenB("#FF0000") > 0 And UseStandard Then
69450     .StampFontColor = "#FF0000"
69460    Else
69470     If LenB(tStr) > 0 Then
69480      .StampFontColor = tStr
69490     End If
69500   End If
69510   tStr = hOpt.Retrieve("StampFontname")
69520   If LenB(tStr) = 0 And LenB("Arial") > 0 And UseStandard Then
69530     .StampFontname = "Arial"
69540    Else
69550     If LenB(tStr) > 0 Then
69560      .StampFontname = tStr
69570     End If
69580   End If
69590   tStr = hOpt.Retrieve("StampFontsize")
69600   If IsNumeric(tStr) Then
69610     If CLng(tStr) >= 1 Then
69620       .StampFontsize = CLng(tStr)
69630      Else
69640       If UseStandard Then
69650        .StampFontsize = 48
69660       End If
69670     End If
69680    Else
69690     If UseStandard Then
69700      .StampFontsize = 48
69710     End If
69720   End If
69730   tStr = hOpt.Retrieve("StampOutlineFontthickness")
69740   If IsNumeric(tStr) Then
69750     If CLng(tStr) >= 0 Then
69760       .StampOutlineFontthickness = CLng(tStr)
69770      Else
69780       If UseStandard Then
69790        .StampOutlineFontthickness = 0
69800       End If
69810     End If
69820    Else
69830     If UseStandard Then
69840      .StampOutlineFontthickness = 0
69850     End If
69860   End If
69870   tStr = hOpt.Retrieve("StampString")
69880   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
69890     .StampString = ""
69900    Else
69910     If LenB(tStr) > 0 Then
69920      .StampString = tStr
69930     End If
69940   End If
69950   tStr = hOpt.Retrieve("StampUseOutlineFont")
69960   If IsNumeric(tStr) Then
69970     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69980       .StampUseOutlineFont = CLng(tStr)
69990      Else
70000       If UseStandard Then
70010        .StampUseOutlineFont = 1
70020       End If
70030     End If
70040    Else
70050     If UseStandard Then
70060      .StampUseOutlineFont = 1
70070     End If
70080   End If
70090   tStr = hOpt.Retrieve("StandardAuthor")
70100   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70110     .StandardAuthor = ""
70120    Else
70130     If LenB(tStr) > 0 Then
70140      .StandardAuthor = tStr
70150     End If
70160   End If
70170   tStr = hOpt.Retrieve("StandardCreationdate")
70180   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70190     .StandardCreationdate = ""
70200    Else
70210     If LenB(tStr) > 0 Then
70220      .StandardCreationdate = tStr
70230     End If
70240   End If
70250   tStr = hOpt.Retrieve("StandardDateformat")
70260   If LenB(tStr) = 0 And LenB("YYYYMMDDHHNNSS") > 0 And UseStandard Then
70270     .StandardDateformat = "YYYYMMDDHHNNSS"
70280    Else
70290     If LenB(tStr) > 0 Then
70300      .StandardDateformat = tStr
70310     End If
70320   End If
70330   tStr = hOpt.Retrieve("StandardKeywords")
70340   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70350     .StandardKeywords = ""
70360    Else
70370     If LenB(tStr) > 0 Then
70380      .StandardKeywords = tStr
70390     End If
70400   End If
70410   tStr = hOpt.Retrieve("StandardMailDomain")
70420   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70430     .StandardMailDomain = ""
70440    Else
70450     If LenB(tStr) > 0 Then
70460      .StandardMailDomain = tStr
70470     End If
70480   End If
70490   tStr = hOpt.Retrieve("StandardModifydate")
70500   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70510     .StandardModifydate = ""
70520    Else
70530     If LenB(tStr) > 0 Then
70540      .StandardModifydate = tStr
70550     End If
70560   End If
70570   tStr = hOpt.Retrieve("StandardSaveformat")
70580   If IsNumeric(tStr) Then
70590     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
70600       .StandardSaveformat = CLng(tStr)
70610      Else
70620       If UseStandard Then
70630        .StandardSaveformat = 0
70640       End If
70650     End If
70660    Else
70670     If UseStandard Then
70680      .StandardSaveformat = 0
70690     End If
70700   End If
70710   tStr = hOpt.Retrieve("StandardSubject")
70720   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70730     .StandardSubject = ""
70740    Else
70750     If LenB(tStr) > 0 Then
70760      .StandardSubject = tStr
70770     End If
70780   End If
70790   tStr = hOpt.Retrieve("StandardTitle")
70800   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70810     .StandardTitle = ""
70820    Else
70830     If LenB(tStr) > 0 Then
70840      .StandardTitle = tStr
70850     End If
70860   End If
70870   tStr = hOpt.Retrieve("StartStandardProgram")
70880   If IsNumeric(tStr) Then
70890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70900       .StartStandardProgram = CLng(tStr)
70910      Else
70920       If UseStandard Then
70930        .StartStandardProgram = 1
70940       End If
70950     End If
70960    Else
70970     If UseStandard Then
70980      .StartStandardProgram = 1
70990     End If
71000   End If
71010   tStr = hOpt.Retrieve("SVGResolution")
71020   If IsNumeric(tStr) Then
71030     If CLng(tStr) >= 1 Then
71040       .SVGResolution = CLng(tStr)
71050      Else
71060       If UseStandard Then
71070        .SVGResolution = 72
71080       End If
71090     End If
71100    Else
71110     If UseStandard Then
71120      .SVGResolution = 72
71130     End If
71140   End If
71150   tStr = hOpt.Retrieve("TIFFColorscount")
71160   If IsNumeric(tStr) Then
71170     If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
71180       .TIFFColorscount = CLng(tStr)
71190      Else
71200       If UseStandard Then
71210        .TIFFColorscount = 0
71220       End If
71230     End If
71240    Else
71250     If UseStandard Then
71260      .TIFFColorscount = 0
71270     End If
71280   End If
71290   tStr = hOpt.Retrieve("TIFFResolution")
71300   If IsNumeric(tStr) Then
71310     If CLng(tStr) >= 1 Then
71320       .TIFFResolution = CLng(tStr)
71330      Else
71340       If UseStandard Then
71350        .TIFFResolution = 150
71360       End If
71370     End If
71380    Else
71390     If UseStandard Then
71400      .TIFFResolution = 150
71410     End If
71420   End If
71430   tStr = hOpt.Retrieve("Toolbars")
71440   If IsNumeric(tStr) Then
71450     If CLng(tStr) >= 0 Then
71460       .Toolbars = CLng(tStr)
71470      Else
71480       If UseStandard Then
71490        .Toolbars = 1
71500       End If
71510     End If
71520    Else
71530     If UseStandard Then
71540      .Toolbars = 1
71550     End If
71560   End If
71570   tStr = hOpt.Retrieve("UpdateInterval")
71580   If IsNumeric(tStr) Then
71590     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
71600       .UpdateInterval = CLng(tStr)
71610      Else
71620       If UseStandard Then
71630        .UpdateInterval = 2
71640       End If
71650     End If
71660    Else
71670     If UseStandard Then
71680      .UpdateInterval = 2
71690     End If
71700   End If
71710   tStr = hOpt.Retrieve("UseAutosave")
71720   If IsNumeric(tStr) Then
71730     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71740       .UseAutosave = CLng(tStr)
71750      Else
71760       If UseStandard Then
71770        .UseAutosave = 0
71780       End If
71790     End If
71800    Else
71810     If UseStandard Then
71820      .UseAutosave = 0
71830     End If
71840   End If
71850   tStr = hOpt.Retrieve("UseAutosaveDirectory")
71860   If IsNumeric(tStr) Then
71870     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71880       .UseAutosaveDirectory = CLng(tStr)
71890      Else
71900       If UseStandard Then
71910        .UseAutosaveDirectory = 1
71920       End If
71930     End If
71940    Else
71950     If UseStandard Then
71960      .UseAutosaveDirectory = 1
71970     End If
71980   End If
71990   tStr = hOpt.Retrieve("UseCreationDateNow")
72000   If IsNumeric(tStr) Then
72010     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
72020       .UseCreationDateNow = CLng(tStr)
72030      Else
72040       If UseStandard Then
72050        .UseCreationDateNow = 0
72060       End If
72070     End If
72080    Else
72090     If UseStandard Then
72100      .UseCreationDateNow = 0
72110     End If
72120   End If
72130   tStr = hOpt.Retrieve("UseCustomPaperSize")
72140   If LenB(tStr) = 0 And LenB("0") > 0 And UseStandard Then
72150     .UseCustomPaperSize = "0"
72160    Else
72170     If LenB(tStr) > 0 Then
72180      .UseCustomPaperSize = tStr
72190     End If
72200   End If
72210   tStr = hOpt.Retrieve("UseFixPapersize")
72220   If IsNumeric(tStr) Then
72230     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
72240       .UseFixPapersize = CLng(tStr)
72250      Else
72260       If UseStandard Then
72270        .UseFixPapersize = 0
72280       End If
72290     End If
72300    Else
72310     If UseStandard Then
72320      .UseFixPapersize = 0
72330     End If
72340   End If
72350   tStr = hOpt.Retrieve("UseStandardAuthor")
72360   If IsNumeric(tStr) Then
72370     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
72380       .UseStandardAuthor = CLng(tStr)
72390      Else
72400       If UseStandard Then
72410        .UseStandardAuthor = 0
72420       End If
72430     End If
72440    Else
72450     If UseStandard Then
72460      .UseStandardAuthor = 0
72470     End If
72480   End If
72490  End With
72500  Set ini = Nothing
72510  ReadOptionsINI = myOptions
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

Public Function ProfileExists(ByVal Profilename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Profiles As Collection, i As Long
50020  Set Profiles = GetProfiles
50030  For i = 1 To Profiles.Count
50040   If LCase$(Profiles(i)) = LCase$(Profilename) Then
50050    ProfileExists = True
50060    Exit Function
50070   End If
50080  Next i
50090  ProfileExists = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "ProfileExists")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetProfiles() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020
50030  Set reg = New clsRegistry
50040  reg.KeyRoot = "Software\PDFCreator\Profiles\"
50050
50060  If InstalledAsServer Then
50070    reg.hkey = HKEY_LOCAL_MACHINE
50080   Else
50090    reg.hkey = HKEY_CURRENT_USER
50100  End If
50110  Set GetProfiles = reg.EnumRegistryKeys(reg.hkey, reg.KeyRoot)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "GetProfiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetPrinterDefaultProfile(Printername As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, PrinterProfiles As Collection, i As Long
50020
50030  Set reg = New clsRegistry
50040  reg.KeyRoot = "Software\PDFCreator\Printers\"
50050
50060  If InstalledAsServer Then
50070    reg.hkey = HKEY_LOCAL_MACHINE
50080   Else
50090    reg.hkey = HKEY_CURRENT_USER
50100  End If
50110  Set PrinterProfiles = reg.EnumRegistryValues(reg.hkey, "Software\PDFCreator\Printers\")
50120  For i = 1 To PrinterProfiles.Count
50130   If UCase$(Trim$(PrinterProfiles(i)(0))) = UCase$(Trim$(Printername)) Then
50140    GetPrinterDefaultProfile = PrinterProfiles(i)(1)
50150    Exit For
50160   End If
50170  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "GetPrinterDefaultProfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub DeleteProfile(Profilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030
50040  Profilename = Trim$(Profilename)
50050
50060  reg.KeyRoot = "Software\PDFCreator\Profiles"
50070
50080  If InstalledAsServer Then
50090    reg.hkey = HKEY_LOCAL_MACHINE
50100   Else
50110    reg.hkey = HKEY_CURRENT_USER
50120  End If
50130  reg.DeleteKeyWithSubkeys Profilename
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "DeleteProfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SaveOptions(sOptions As tOptions, Optional Profilename As String = "")
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String
50020
50030  CorrectOptionsBeforeSaving
50040
50050  Profilename = Trim$(Profilename)
50060  If LenB(Profilename) > 0 Then
50070   tStr = "_" & Profilename
50080  End If
50090
50100  If InstalledAsServer Then
50110    SaveOptionsREG sOptions, HKEY_LOCAL_MACHINE, Profilename
50120   Else
50130    SaveOptionsREG sOptions, , Profilename
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
50020    SaveOptionREG sOptions, OptionName, HKEY_LOCAL_MACHINE
50030   Else
50040    SaveOptionREG sOptions, OptionName
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
50660   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMaximumFactor"
50670   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMediumFactor"
50680   Case "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMinimumFactor"
50690   Case "PDFCOMPRESSIONCOLORRESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50700   Case "PDFCOMPRESSIONCOLORRESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50710   Case "PDFCOMPRESSIONCOLORRESOLUTION": ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50720   Case "PDFCOMPRESSIONGREYCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50730   Case "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50740   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGHighFactor"
50750   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGLowFactor"
50760   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMaximumFactor"
50770   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMediumFactor"
50780   Case "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR": ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMinimumFactor"
50790   Case "PDFCOMPRESSIONGREYRESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50800   Case "PDFCOMPRESSIONGREYRESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50810   Case "PDFCOMPRESSIONGREYRESOLUTION": ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50820   Case "PDFCOMPRESSIONMONOCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50830   Case "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE": ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50840   Case "PDFCOMPRESSIONMONORESAMPLE": ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50850   Case "PDFCOMPRESSIONMONORESAMPLECHOICE": ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50860   Case "PDFCOMPRESSIONMONORESOLUTION": ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50870   Case "PDFCOMPRESSIONTEXTCOMPRESSION": ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50880   Case "PDFDISALLOWCOPY": ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50890   Case "PDFDISALLOWMODIFYANNOTATIONS": ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50900   Case "PDFDISALLOWMODIFYCONTENTS": ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50910   Case "PDFDISALLOWPRINTING": ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50920   Case "PDFENCRYPTOR": ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50930   Case "PDFFONTSEMBEDALL": ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50940   Case "PDFFONTSSUBSETFONTS": ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50950   Case "PDFFONTSSUBSETFONTSPERCENT": ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50960   Case "PDFGENERALASCII85": ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50970   Case "PDFGENERALAUTOROTATE": ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50980   Case "PDFGENERALCOMPATIBILITY": ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50990   Case "PDFGENERALDEFAULT": ini.SaveKey CStr(.PDFGeneralDefault), "PDFGeneralDefault"
51000   Case "PDFGENERALOVERPRINT": ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
51010   Case "PDFGENERALRESOLUTION": ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
51020   Case "PDFHIGHENCRYPTION": ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
51030   Case "PDFLOWENCRYPTION": ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
51040   Case "PDFOPTIMIZE": ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
51050   Case "PDFOWNERPASS": ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
51060   Case "PDFOWNERPASSWORDSTRING": ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
51070   Case "PDFSIGNINGMULTISIGNATURE": ini.SaveKey CStr(Abs(.PDFSigningMultiSignature)), "PDFSigningMultiSignature"
51080   Case "PDFSIGNINGPFXFILE": ini.SaveKey CStr(.PDFSigningPFXFile), "PDFSigningPFXFile"
51090   Case "PDFSIGNINGPFXFILEPASSWORD": ini.SaveKey CStr(.PDFSigningPFXFilePassword), "PDFSigningPFXFilePassword"
51100   Case "PDFSIGNINGSIGNATURECONTACT": ini.SaveKey CStr(.PDFSigningSignatureContact), "PDFSigningSignatureContact"
51110   Case "PDFSIGNINGSIGNATURELEFTX": ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), "PDFSigningSignatureLeftX"
51120   Case "PDFSIGNINGSIGNATURELEFTY": ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), "PDFSigningSignatureLeftY"
51130   Case "PDFSIGNINGSIGNATURELOCATION": ini.SaveKey CStr(.PDFSigningSignatureLocation), "PDFSigningSignatureLocation"
51140   Case "PDFSIGNINGSIGNATUREREASON": ini.SaveKey CStr(.PDFSigningSignatureReason), "PDFSigningSignatureReason"
51150   Case "PDFSIGNINGSIGNATURERIGHTX": ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), "PDFSigningSignatureRightX"
51160   Case "PDFSIGNINGSIGNATURERIGHTY": ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), "PDFSigningSignatureRightY"
51170   Case "PDFSIGNINGSIGNATUREVISIBLE": ini.SaveKey CStr(Abs(.PDFSigningSignatureVisible)), "PDFSigningSignatureVisible"
51180   Case "PDFSIGNINGSIGNPDF": ini.SaveKey CStr(Abs(.PDFSigningSignPDF)), "PDFSigningSignPDF"
51190   Case "PDFUPDATEMETADATA": ini.SaveKey CStr(.PDFUpdateMetadata), "PDFUpdateMetadata"
51200   Case "PDFUSERPASS": ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
51210   Case "PDFUSERPASSWORDSTRING": ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
51220   Case "PDFUSESECURITY": ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
51230   Case "PNGCOLORSCOUNT": ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
51240   Case "PNGRESOLUTION": ini.SaveKey CStr(.PNGResolution), "PNGResolution"
51250   Case "PRINTAFTERSAVING": ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
51260   Case "PRINTAFTERSAVINGDUPLEX": ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
51270   Case "PRINTAFTERSAVINGNOCANCEL": ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
51280   Case "PRINTAFTERSAVINGPRINTER": ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
51290   Case "PRINTAFTERSAVINGQUERYUSER": ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
51300   Case "PRINTAFTERSAVINGTUMBLE": ini.SaveKey CStr(.PrintAfterSavingTumble), "PrintAfterSavingTumble"
51310   Case "PRINTERSTOP": ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
51320   Case "PRINTERTEMPPATH": ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
51330   Case "PROCESSPRIORITY": ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
51340   Case "PROGRAMFONT": ini.SaveKey CStr(.ProgramFont), "ProgramFont"
51350   Case "PROGRAMFONTCHARSET": ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
51360   Case "PROGRAMFONTSIZE": ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
51370   Case "PSDCOLORSCOUNT": ini.SaveKey CStr(.PSDColorsCount), "PSDColorsCount"
51380   Case "PSDRESOLUTION": ini.SaveKey CStr(.PSDResolution), "PSDResolution"
51390   Case "PSLANGUAGELEVEL": ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
51400   Case "RAWCOLORSCOUNT": ini.SaveKey CStr(.RAWColorsCount), "RAWColorsCount"
51410   Case "RAWRESOLUTION": ini.SaveKey CStr(.RAWResolution), "RAWResolution"
51420   Case "REMOVEALLKNOWNFILEEXTENSIONS": ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
51430   Case "REMOVESPACES": ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
51440   Case "RUNPROGRAMAFTERSAVING": ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
51450   Case "RUNPROGRAMAFTERSAVINGPROGRAMNAME": ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
51460   Case "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS": ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
51470   Case "RUNPROGRAMAFTERSAVINGWAITUNTILREADY": ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51480   Case "RUNPROGRAMAFTERSAVINGWINDOWSTYLE": ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51490   Case "RUNPROGRAMBEFORESAVING": ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51500   Case "RUNPROGRAMBEFORESAVINGPROGRAMNAME": ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51510   Case "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS": ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51520   Case "RUNPROGRAMBEFORESAVINGWINDOWSTYLE": ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51530   Case "SAVEFILENAME": ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51540   Case "SENDEMAILAFTERAUTOSAVING": ini.SaveKey CStr(Abs(.SendEmailAfterAutoSaving)), "SendEmailAfterAutoSaving"
51550   Case "SENDMAILMETHOD": ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51560   Case "SHOWANIMATION": ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51570   Case "STAMPFONTCOLOR": ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51580   Case "STAMPFONTNAME": ini.SaveKey CStr(.StampFontname), "StampFontname"
51590   Case "STAMPFONTSIZE": ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51600   Case "STAMPOUTLINEFONTTHICKNESS": ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51610   Case "STAMPSTRING": ini.SaveKey CStr(.StampString), "StampString"
51620   Case "STAMPUSEOUTLINEFONT": ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51630   Case "STANDARDAUTHOR": ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51640   Case "STANDARDCREATIONDATE": ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51650   Case "STANDARDDATEFORMAT": ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51660   Case "STANDARDKEYWORDS": ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51670   Case "STANDARDMAILDOMAIN": ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51680   Case "STANDARDMODIFYDATE": ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51690   Case "STANDARDSAVEFORMAT": ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51700   Case "STANDARDSUBJECT": ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51710   Case "STANDARDTITLE": ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51720   Case "STARTSTANDARDPROGRAM": ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51730   Case "SVGRESOLUTION": ini.SaveKey CStr(.SVGResolution), "SVGResolution"
51740   Case "TIFFCOLORSCOUNT": ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51750   Case "TIFFRESOLUTION": ini.SaveKey CStr(.TIFFResolution), "TIFFResolution"
51760   Case "TOOLBARS": ini.SaveKey CStr(.Toolbars), "Toolbars"
51770   Case "UPDATEINTERVAL": ini.SaveKey CStr(.UpdateInterval), "UpdateInterval"
51780   Case "USEAUTOSAVE": ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51790   Case "USEAUTOSAVEDIRECTORY": ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51800   Case "USECREATIONDATENOW": ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51810   Case "USECUSTOMPAPERSIZE": ini.SaveKey CStr(.UseCustomPaperSize), "UseCustomPaperSize"
51820   Case "USEFIXPAPERSIZE": ini.SaveKey CStr(Abs(.UseFixPapersize)), "UseFixPapersize"
51830   Case "USESTANDARDAUTHOR": ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51840   End Select
51850  End With
51860  Set ini = Nothing
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
50650   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMaximumFactor"
50660   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMediumFactor"
50670   ini.SaveKey Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionColorCompressionJPEGMinimumFactor"
50680   ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50690   ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50700   ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50710   ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50720   ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50730   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGHighFactor"
50740   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGLowFactor"
50750   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMaximumFactor"
50760   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMediumFactor"
50770   ini.SaveKey Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), "PDFCompressionGreyCompressionJPEGMinimumFactor"
50780   ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50790   ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50800   ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50810   ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50820   ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50830   ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50840   ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50850   ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50860   ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50870   ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50880   ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50890   ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50900   ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50910   ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50920   ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50930   ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50940   ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50950   ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50960   ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50970   ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50980   ini.SaveKey CStr(.PDFGeneralDefault), "PDFGeneralDefault"
50990   ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
51000   ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
51010   ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
51020   ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
51030   ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
51040   ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
51050   ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
51060   ini.SaveKey CStr(Abs(.PDFSigningMultiSignature)), "PDFSigningMultiSignature"
51070   ini.SaveKey CStr(.PDFSigningPFXFile), "PDFSigningPFXFile"
51080   ini.SaveKey CStr(.PDFSigningPFXFilePassword), "PDFSigningPFXFilePassword"
51090   ini.SaveKey CStr(.PDFSigningSignatureContact), "PDFSigningSignatureContact"
51100   ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), "PDFSigningSignatureLeftX"
51110   ini.SaveKey Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), "PDFSigningSignatureLeftY"
51120   ini.SaveKey CStr(.PDFSigningSignatureLocation), "PDFSigningSignatureLocation"
51130   ini.SaveKey CStr(.PDFSigningSignatureReason), "PDFSigningSignatureReason"
51140   ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), "PDFSigningSignatureRightX"
51150   ini.SaveKey Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), "PDFSigningSignatureRightY"
51160   ini.SaveKey CStr(Abs(.PDFSigningSignatureVisible)), "PDFSigningSignatureVisible"
51170   ini.SaveKey CStr(Abs(.PDFSigningSignPDF)), "PDFSigningSignPDF"
51180   ini.SaveKey CStr(.PDFUpdateMetadata), "PDFUpdateMetadata"
51190   ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
51200   ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
51210   ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
51220   ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
51230   ini.SaveKey CStr(.PNGResolution), "PNGResolution"
51240   ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
51250   ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
51260   ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
51270   ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
51280   ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
51290   ini.SaveKey CStr(.PrintAfterSavingTumble), "PrintAfterSavingTumble"
51300   ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
51310   ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
51320   ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
51330   ini.SaveKey CStr(.ProgramFont), "ProgramFont"
51340   ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
51350   ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
51360   ini.SaveKey CStr(.PSDColorsCount), "PSDColorsCount"
51370   ini.SaveKey CStr(.PSDResolution), "PSDResolution"
51380   ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
51390   ini.SaveKey CStr(.RAWColorsCount), "RAWColorsCount"
51400   ini.SaveKey CStr(.RAWResolution), "RAWResolution"
51410   ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
51420   ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
51430   ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
51440   ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
51450   ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
51460   ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51470   ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51480   ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51490   ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51500   ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51510   ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51520   ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51530   ini.SaveKey CStr(Abs(.SendEmailAfterAutoSaving)), "SendEmailAfterAutoSaving"
51540   ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51550   ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51560   ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51570   ini.SaveKey CStr(.StampFontname), "StampFontname"
51580   ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51590   ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51600   ini.SaveKey CStr(.StampString), "StampString"
51610   ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51620   ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51630   ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51640   ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51650   ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51660   ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51670   ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51680   ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51690   ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51700   ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51710   ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51720   ini.SaveKey CStr(.SVGResolution), "SVGResolution"
51730   ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51740   ini.SaveKey CStr(.TIFFResolution), "TIFFResolution"
51750   ini.SaveKey CStr(.Toolbars), "Toolbars"
51760   ini.SaveKey CStr(.UpdateInterval), "UpdateInterval"
51770   ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51780   ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51790   ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51800   ini.SaveKey CStr(.UseCustomPaperSize), "UseCustomPaperSize"
51810   ini.SaveKey CStr(Abs(.UseFixPapersize)), "UseFixPapersize"
51820   ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51830  End With
51840  Set ini = Nothing
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
56420     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
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
56820   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMaximumFactor")
56830   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56840     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56850       .PDFCompressionColorCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
56860      Else
56870       If UseStandard Then
56880        .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56890       End If
56900     End If
56910    Else
56920     If UseStandard Then
56930      .PDFCompressionColorCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
56940     End If
56950   End If
56960   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMediumFactor")
56970   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
56980     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
56990       .PDFCompressionColorCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57000      Else
57010       If UseStandard Then
57020        .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57030       End If
57040     End If
57050    Else
57060     If UseStandard Then
57070      .PDFCompressionColorCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
57080     End If
57090   End If
57100   tStr = reg.GetRegistryValue("PDFCompressionColorCompressionJPEGMinimumFactor")
57110   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57120     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57130       .PDFCompressionColorCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57140      Else
57150       If UseStandard Then
57160        .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57170       End If
57180     End If
57190    Else
57200     If UseStandard Then
57210      .PDFCompressionColorCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
57220     End If
57230   End If
57240   tStr = reg.GetRegistryValue("PDFCompressionColorResample")
57250   If IsNumeric(tStr) Then
57260     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57270       .PDFCompressionColorResample = CLng(tStr)
57280      Else
57290       If UseStandard Then
57300        .PDFCompressionColorResample = 0
57310       End If
57320     End If
57330    Else
57340     If UseStandard Then
57350      .PDFCompressionColorResample = 0
57360     End If
57370   End If
57380   tStr = reg.GetRegistryValue("PDFCompressionColorResampleChoice")
57390   If IsNumeric(tStr) Then
57400     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
57410       .PDFCompressionColorResampleChoice = CLng(tStr)
57420      Else
57430       If UseStandard Then
57440        .PDFCompressionColorResampleChoice = 0
57450       End If
57460     End If
57470    Else
57480     If UseStandard Then
57490      .PDFCompressionColorResampleChoice = 0
57500     End If
57510   End If
57520   tStr = reg.GetRegistryValue("PDFCompressionColorResolution")
57530   If IsNumeric(tStr) Then
57540     If CLng(tStr) >= 0 Then
57550       .PDFCompressionColorResolution = CLng(tStr)
57560      Else
57570       If UseStandard Then
57580        .PDFCompressionColorResolution = 300
57590       End If
57600     End If
57610    Else
57620     If UseStandard Then
57630      .PDFCompressionColorResolution = 300
57640     End If
57650   End If
57660   tStr = reg.GetRegistryValue("PDFCompressionGreyCompression")
57670   If IsNumeric(tStr) Then
57680     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
57690       .PDFCompressionGreyCompression = CLng(tStr)
57700      Else
57710       If UseStandard Then
57720        .PDFCompressionGreyCompression = 1
57730       End If
57740     End If
57750    Else
57760     If UseStandard Then
57770      .PDFCompressionGreyCompression = 1
57780     End If
57790   End If
57800   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionChoice")
57810   If IsNumeric(tStr) Then
57820     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
57830       .PDFCompressionGreyCompressionChoice = CLng(tStr)
57840      Else
57850       If UseStandard Then
57860        .PDFCompressionGreyCompressionChoice = 0
57870       End If
57880     End If
57890    Else
57900     If UseStandard Then
57910      .PDFCompressionGreyCompressionChoice = 0
57920     End If
57930   End If
57940   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGHighFactor")
57950   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
57960     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
57970       .PDFCompressionGreyCompressionJPEGHighFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
57980      Else
57990       If UseStandard Then
58000        .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
58010       End If
58020     End If
58030    Else
58040     If UseStandard Then
58050      .PDFCompressionGreyCompressionJPEGHighFactor = Replace$("0.9", ".", GetDecimalChar)
58060     End If
58070   End If
58080   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGLowFactor")
58090   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58100     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58110       .PDFCompressionGreyCompressionJPEGLowFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58120      Else
58130       If UseStandard Then
58140        .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58150       End If
58160     End If
58170    Else
58180     If UseStandard Then
58190      .PDFCompressionGreyCompressionJPEGLowFactor = Replace$("0.25", ".", GetDecimalChar)
58200     End If
58210   End If
58220   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMaximumFactor")
58230   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58240     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58250       .PDFCompressionGreyCompressionJPEGMaximumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58260      Else
58270       If UseStandard Then
58280        .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
58290       End If
58300     End If
58310    Else
58320     If UseStandard Then
58330      .PDFCompressionGreyCompressionJPEGMaximumFactor = Replace$("2", ".", GetDecimalChar)
58340     End If
58350   End If
58360   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMediumFactor")
58370   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58380     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58390       .PDFCompressionGreyCompressionJPEGMediumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58400      Else
58410       If UseStandard Then
58420        .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58430       End If
58440     End If
58450    Else
58460     If UseStandard Then
58470      .PDFCompressionGreyCompressionJPEGMediumFactor = Replace$("0.5", ".", GetDecimalChar)
58480     End If
58490   End If
58500   tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionJPEGMinimumFactor")
58510   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
58520     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
58530       .PDFCompressionGreyCompressionJPEGMinimumFactor = CDbl(Replace$(tStr, ".", GetDecimalChar))
58540      Else
58550       If UseStandard Then
58560        .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58570       End If
58580     End If
58590    Else
58600     If UseStandard Then
58610      .PDFCompressionGreyCompressionJPEGMinimumFactor = Replace$("0.1", ".", GetDecimalChar)
58620     End If
58630   End If
58640   tStr = reg.GetRegistryValue("PDFCompressionGreyResample")
58650   If IsNumeric(tStr) Then
58660     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
58670       .PDFCompressionGreyResample = CLng(tStr)
58680      Else
58690       If UseStandard Then
58700        .PDFCompressionGreyResample = 0
58710       End If
58720     End If
58730    Else
58740     If UseStandard Then
58750      .PDFCompressionGreyResample = 0
58760     End If
58770   End If
58780   tStr = reg.GetRegistryValue("PDFCompressionGreyResampleChoice")
58790   If IsNumeric(tStr) Then
58800     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
58810       .PDFCompressionGreyResampleChoice = CLng(tStr)
58820      Else
58830       If UseStandard Then
58840        .PDFCompressionGreyResampleChoice = 0
58850       End If
58860     End If
58870    Else
58880     If UseStandard Then
58890      .PDFCompressionGreyResampleChoice = 0
58900     End If
58910   End If
58920   tStr = reg.GetRegistryValue("PDFCompressionGreyResolution")
58930   If IsNumeric(tStr) Then
58940     If CLng(tStr) >= 0 Then
58950       .PDFCompressionGreyResolution = CLng(tStr)
58960      Else
58970       If UseStandard Then
58980        .PDFCompressionGreyResolution = 300
58990       End If
59000     End If
59010    Else
59020     If UseStandard Then
59030      .PDFCompressionGreyResolution = 300
59040     End If
59050   End If
59060   tStr = reg.GetRegistryValue("PDFCompressionMonoCompression")
59070   If IsNumeric(tStr) Then
59080     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59090       .PDFCompressionMonoCompression = CLng(tStr)
59100      Else
59110       If UseStandard Then
59120        .PDFCompressionMonoCompression = 1
59130       End If
59140     End If
59150    Else
59160     If UseStandard Then
59170      .PDFCompressionMonoCompression = 1
59180     End If
59190   End If
59200   tStr = reg.GetRegistryValue("PDFCompressionMonoCompressionChoice")
59210   If IsNumeric(tStr) Then
59220     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
59230       .PDFCompressionMonoCompressionChoice = CLng(tStr)
59240      Else
59250       If UseStandard Then
59260        .PDFCompressionMonoCompressionChoice = 0
59270       End If
59280     End If
59290    Else
59300     If UseStandard Then
59310      .PDFCompressionMonoCompressionChoice = 0
59320     End If
59330   End If
59340   tStr = reg.GetRegistryValue("PDFCompressionMonoResample")
59350   If IsNumeric(tStr) Then
59360     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59370       .PDFCompressionMonoResample = CLng(tStr)
59380      Else
59390       If UseStandard Then
59400        .PDFCompressionMonoResample = 0
59410       End If
59420     End If
59430    Else
59440     If UseStandard Then
59450      .PDFCompressionMonoResample = 0
59460     End If
59470   End If
59480   tStr = reg.GetRegistryValue("PDFCompressionMonoResampleChoice")
59490   If IsNumeric(tStr) Then
59500     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
59510       .PDFCompressionMonoResampleChoice = CLng(tStr)
59520      Else
59530       If UseStandard Then
59540        .PDFCompressionMonoResampleChoice = 0
59550       End If
59560     End If
59570    Else
59580     If UseStandard Then
59590      .PDFCompressionMonoResampleChoice = 0
59600     End If
59610   End If
59620   tStr = reg.GetRegistryValue("PDFCompressionMonoResolution")
59630   If IsNumeric(tStr) Then
59640     If CLng(tStr) >= 0 Then
59650       .PDFCompressionMonoResolution = CLng(tStr)
59660      Else
59670       If UseStandard Then
59680        .PDFCompressionMonoResolution = 1200
59690       End If
59700     End If
59710    Else
59720     If UseStandard Then
59730      .PDFCompressionMonoResolution = 1200
59740     End If
59750   End If
59760   tStr = reg.GetRegistryValue("PDFCompressionTextCompression")
59770   If IsNumeric(tStr) Then
59780     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59790       .PDFCompressionTextCompression = CLng(tStr)
59800      Else
59810       If UseStandard Then
59820        .PDFCompressionTextCompression = 1
59830       End If
59840     End If
59850    Else
59860     If UseStandard Then
59870      .PDFCompressionTextCompression = 1
59880     End If
59890   End If
59900   reg.SubKey = "Printing\Formats\PDF\Fonts"
59910   tStr = reg.GetRegistryValue("PDFFontsEmbedAll")
59920   If IsNumeric(tStr) Then
59930     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
59940       .PDFFontsEmbedAll = CLng(tStr)
59950      Else
59960       If UseStandard Then
59970        .PDFFontsEmbedAll = 1
59980       End If
59990     End If
60000    Else
60010     If UseStandard Then
60020      .PDFFontsEmbedAll = 1
60030     End If
60040   End If
60050   tStr = reg.GetRegistryValue("PDFFontsSubSetFonts")
60060   If IsNumeric(tStr) Then
60070     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60080       .PDFFontsSubSetFonts = CLng(tStr)
60090      Else
60100       If UseStandard Then
60110        .PDFFontsSubSetFonts = 1
60120       End If
60130     End If
60140    Else
60150     If UseStandard Then
60160      .PDFFontsSubSetFonts = 1
60170     End If
60180   End If
60190   tStr = reg.GetRegistryValue("PDFFontsSubSetFontsPercent")
60200   If IsNumeric(tStr) Then
60210     If CLng(tStr) >= 0 Then
60220       .PDFFontsSubSetFontsPercent = CLng(tStr)
60230      Else
60240       If UseStandard Then
60250        .PDFFontsSubSetFontsPercent = 100
60260       End If
60270     End If
60280    Else
60290     If UseStandard Then
60300      .PDFFontsSubSetFontsPercent = 100
60310     End If
60320   End If
60330   reg.SubKey = "Printing\Formats\PDF\General"
60340   tStr = reg.GetRegistryValue("PDFGeneralASCII85")
60350   If IsNumeric(tStr) Then
60360     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
60370       .PDFGeneralASCII85 = CLng(tStr)
60380      Else
60390       If UseStandard Then
60400        .PDFGeneralASCII85 = 0
60410       End If
60420     End If
60430    Else
60440     If UseStandard Then
60450      .PDFGeneralASCII85 = 0
60460     End If
60470   End If
60480   tStr = reg.GetRegistryValue("PDFGeneralAutorotate")
60490   If IsNumeric(tStr) Then
60500     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
60510       .PDFGeneralAutorotate = CLng(tStr)
60520      Else
60530       If UseStandard Then
60540        .PDFGeneralAutorotate = 2
60550       End If
60560     End If
60570    Else
60580     If UseStandard Then
60590      .PDFGeneralAutorotate = 2
60600     End If
60610   End If
60620   tStr = reg.GetRegistryValue("PDFGeneralCompatibility")
60630   If IsNumeric(tStr) Then
60640     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
60650       .PDFGeneralCompatibility = CLng(tStr)
60660      Else
60670       If UseStandard Then
60680        .PDFGeneralCompatibility = 2
60690       End If
60700     End If
60710    Else
60720     If UseStandard Then
60730      .PDFGeneralCompatibility = 2
60740     End If
60750   End If
60760   tStr = reg.GetRegistryValue("PDFGeneralDefault")
60770   If IsNumeric(tStr) Then
60780     If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
60790       .PDFGeneralDefault = CLng(tStr)
60800      Else
60810       If UseStandard Then
60820        .PDFGeneralDefault = 0
60830       End If
60840     End If
60850    Else
60860     If UseStandard Then
60870      .PDFGeneralDefault = 0
60880     End If
60890   End If
60900   tStr = reg.GetRegistryValue("PDFGeneralOverprint")
60910   If IsNumeric(tStr) Then
60920     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
60930       .PDFGeneralOverprint = CLng(tStr)
60940      Else
60950       If UseStandard Then
60960        .PDFGeneralOverprint = 0
60970       End If
60980     End If
60990    Else
61000     If UseStandard Then
61010      .PDFGeneralOverprint = 0
61020     End If
61030   End If
61040   tStr = reg.GetRegistryValue("PDFGeneralResolution")
61050   If IsNumeric(tStr) Then
61060     If CLng(tStr) >= 0 Then
61070       .PDFGeneralResolution = CLng(tStr)
61080      Else
61090       If UseStandard Then
61100        .PDFGeneralResolution = 600
61110       End If
61120     End If
61130    Else
61140     If UseStandard Then
61150      .PDFGeneralResolution = 600
61160     End If
61170   End If
61180   tStr = reg.GetRegistryValue("PDFOptimize")
61190   If IsNumeric(tStr) Then
61200     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61210       .PDFOptimize = CLng(tStr)
61220      Else
61230       If UseStandard Then
61240        .PDFOptimize = 0
61250       End If
61260     End If
61270    Else
61280     If UseStandard Then
61290      .PDFOptimize = 0
61300     End If
61310   End If
61320   tStr = reg.GetRegistryValue("PDFUpdateMetadata")
61330   If IsNumeric(tStr) Then
61340     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
61350       .PDFUpdateMetadata = CLng(tStr)
61360      Else
61370       If UseStandard Then
61380        .PDFUpdateMetadata = 1
61390       End If
61400     End If
61410    Else
61420     If UseStandard Then
61430      .PDFUpdateMetadata = 1
61440     End If
61450   End If
61460   reg.SubKey = "Printing\Formats\PDF\Security"
61470   tStr = reg.GetRegistryValue("PDFAllowAssembly")
61480   If IsNumeric(tStr) Then
61490     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61500       .PDFAllowAssembly = CLng(tStr)
61510      Else
61520       If UseStandard Then
61530        .PDFAllowAssembly = 0
61540       End If
61550     End If
61560    Else
61570     If UseStandard Then
61580      .PDFAllowAssembly = 0
61590     End If
61600   End If
61610   tStr = reg.GetRegistryValue("PDFAllowDegradedPrinting")
61620   If IsNumeric(tStr) Then
61630     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61640       .PDFAllowDegradedPrinting = CLng(tStr)
61650      Else
61660       If UseStandard Then
61670        .PDFAllowDegradedPrinting = 0
61680       End If
61690     End If
61700    Else
61710     If UseStandard Then
61720      .PDFAllowDegradedPrinting = 0
61730     End If
61740   End If
61750   tStr = reg.GetRegistryValue("PDFAllowFillIn")
61760   If IsNumeric(tStr) Then
61770     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61780       .PDFAllowFillIn = CLng(tStr)
61790      Else
61800       If UseStandard Then
61810        .PDFAllowFillIn = 0
61820       End If
61830     End If
61840    Else
61850     If UseStandard Then
61860      .PDFAllowFillIn = 0
61870     End If
61880   End If
61890   tStr = reg.GetRegistryValue("PDFAllowScreenReaders")
61900   If IsNumeric(tStr) Then
61910     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
61920       .PDFAllowScreenReaders = CLng(tStr)
61930      Else
61940       If UseStandard Then
61950        .PDFAllowScreenReaders = 0
61960       End If
61970     End If
61980    Else
61990     If UseStandard Then
62000      .PDFAllowScreenReaders = 0
62010     End If
62020   End If
62030   tStr = reg.GetRegistryValue("PDFDisallowCopy")
62040   If IsNumeric(tStr) Then
62050     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62060       .PDFDisallowCopy = CLng(tStr)
62070      Else
62080       If UseStandard Then
62090        .PDFDisallowCopy = 1
62100       End If
62110     End If
62120    Else
62130     If UseStandard Then
62140      .PDFDisallowCopy = 1
62150     End If
62160   End If
62170   tStr = reg.GetRegistryValue("PDFDisallowModifyAnnotations")
62180   If IsNumeric(tStr) Then
62190     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62200       .PDFDisallowModifyAnnotations = CLng(tStr)
62210      Else
62220       If UseStandard Then
62230        .PDFDisallowModifyAnnotations = 0
62240       End If
62250     End If
62260    Else
62270     If UseStandard Then
62280      .PDFDisallowModifyAnnotations = 0
62290     End If
62300   End If
62310   tStr = reg.GetRegistryValue("PDFDisallowModifyContents")
62320   If IsNumeric(tStr) Then
62330     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62340       .PDFDisallowModifyContents = CLng(tStr)
62350      Else
62360       If UseStandard Then
62370        .PDFDisallowModifyContents = 0
62380       End If
62390     End If
62400    Else
62410     If UseStandard Then
62420      .PDFDisallowModifyContents = 0
62430     End If
62440   End If
62450   tStr = reg.GetRegistryValue("PDFDisallowPrinting")
62460   If IsNumeric(tStr) Then
62470     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62480       .PDFDisallowPrinting = CLng(tStr)
62490      Else
62500       If UseStandard Then
62510        .PDFDisallowPrinting = 0
62520       End If
62530     End If
62540    Else
62550     If UseStandard Then
62560      .PDFDisallowPrinting = 0
62570     End If
62580   End If
62590   tStr = reg.GetRegistryValue("PDFEncryptor")
62600   If IsNumeric(tStr) Then
62610     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
62620       .PDFEncryptor = CLng(tStr)
62630      Else
62640       If UseStandard Then
62650        .PDFEncryptor = 0
62660       End If
62670     End If
62680    Else
62690     If UseStandard Then
62700      .PDFEncryptor = 0
62710     End If
62720   End If
62730   tStr = reg.GetRegistryValue("PDFHighEncryption")
62740   If IsNumeric(tStr) Then
62750     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62760       .PDFHighEncryption = CLng(tStr)
62770      Else
62780       If UseStandard Then
62790        .PDFHighEncryption = 0
62800       End If
62810     End If
62820    Else
62830     If UseStandard Then
62840      .PDFHighEncryption = 0
62850     End If
62860   End If
62870   tStr = reg.GetRegistryValue("PDFLowEncryption")
62880   If IsNumeric(tStr) Then
62890     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
62900       .PDFLowEncryption = CLng(tStr)
62910      Else
62920       If UseStandard Then
62930        .PDFLowEncryption = 1
62940       End If
62950     End If
62960    Else
62970     If UseStandard Then
62980      .PDFLowEncryption = 1
62990     End If
63000   End If
63010   tStr = reg.GetRegistryValue("PDFOwnerPass")
63020   If IsNumeric(tStr) Then
63030     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63040       .PDFOwnerPass = CLng(tStr)
63050      Else
63060       If UseStandard Then
63070        .PDFOwnerPass = 0
63080       End If
63090     End If
63100    Else
63110     If UseStandard Then
63120      .PDFOwnerPass = 0
63130     End If
63140   End If
63150   tStr = reg.GetRegistryValue("PDFOwnerPasswordString")
63160   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63170     .PDFOwnerPasswordString = ""
63180    Else
63190     If LenB(tStr) > 0 Then
63200      .PDFOwnerPasswordString = tStr
63210     End If
63220   End If
63230   tStr = reg.GetRegistryValue("PDFUserPass")
63240   If IsNumeric(tStr) Then
63250     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63260       .PDFUserPass = CLng(tStr)
63270      Else
63280       If UseStandard Then
63290        .PDFUserPass = 0
63300       End If
63310     End If
63320    Else
63330     If UseStandard Then
63340      .PDFUserPass = 0
63350     End If
63360   End If
63370   tStr = reg.GetRegistryValue("PDFUserPasswordString")
63380   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63390     .PDFUserPasswordString = ""
63400    Else
63410     If LenB(tStr) > 0 Then
63420      .PDFUserPasswordString = tStr
63430     End If
63440   End If
63450   tStr = reg.GetRegistryValue("PDFUseSecurity")
63460   If IsNumeric(tStr) Then
63470     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63480       .PDFUseSecurity = CLng(tStr)
63490      Else
63500       If UseStandard Then
63510        .PDFUseSecurity = 0
63520       End If
63530     End If
63540    Else
63550     If UseStandard Then
63560      .PDFUseSecurity = 0
63570     End If
63580   End If
63590   reg.SubKey = "Printing\Formats\PDF\Signing"
63600   tStr = reg.GetRegistryValue("PDFSigningMultiSignature")
63610   If IsNumeric(tStr) Then
63620     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
63630       .PDFSigningMultiSignature = CLng(tStr)
63640      Else
63650       If UseStandard Then
63660        .PDFSigningMultiSignature = 0
63670       End If
63680     End If
63690    Else
63700     If UseStandard Then
63710      .PDFSigningMultiSignature = 0
63720     End If
63730   End If
63740   tStr = reg.GetRegistryValue("PDFSigningPFXFile")
63750   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63760     .PDFSigningPFXFile = ""
63770    Else
63780     If LenB(tStr) > 0 Then
63790      .PDFSigningPFXFile = tStr
63800     End If
63810   End If
63820   tStr = reg.GetRegistryValue("PDFSigningPFXFilePassword")
63830   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63840     .PDFSigningPFXFilePassword = ""
63850    Else
63860     If LenB(tStr) > 0 Then
63870      .PDFSigningPFXFilePassword = tStr
63880     End If
63890   End If
63900   tStr = reg.GetRegistryValue("PDFSigningSignatureContact")
63910   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
63920     .PDFSigningSignatureContact = ""
63930    Else
63940     If LenB(tStr) > 0 Then
63950      .PDFSigningSignatureContact = tStr
63960     End If
63970   End If
63980   tStr = reg.GetRegistryValue("PDFSigningSignatureLeftX")
63990   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
64000     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
64010       .PDFSigningSignatureLeftX = CDbl(Replace$(tStr, ".", GetDecimalChar))
64020      Else
64030       If UseStandard Then
64040        .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
64050       End If
64060     End If
64070    Else
64080     If UseStandard Then
64090      .PDFSigningSignatureLeftX = Replace$("100", ".", GetDecimalChar)
64100     End If
64110   End If
64120   tStr = reg.GetRegistryValue("PDFSigningSignatureLeftY")
64130   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
64140     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
64150       .PDFSigningSignatureLeftY = CDbl(Replace$(tStr, ".", GetDecimalChar))
64160      Else
64170       If UseStandard Then
64180        .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
64190       End If
64200     End If
64210    Else
64220     If UseStandard Then
64230      .PDFSigningSignatureLeftY = Replace$("100", ".", GetDecimalChar)
64240     End If
64250   End If
64260   tStr = reg.GetRegistryValue("PDFSigningSignatureLocation")
64270   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64280     .PDFSigningSignatureLocation = ""
64290    Else
64300     If LenB(tStr) > 0 Then
64310      .PDFSigningSignatureLocation = tStr
64320     End If
64330   End If
64340   tStr = reg.GetRegistryValue("PDFSigningSignatureReason")
64350   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
64360     .PDFSigningSignatureReason = ""
64370    Else
64380     If LenB(tStr) > 0 Then
64390      .PDFSigningSignatureReason = tStr
64400     End If
64410   End If
64420   tStr = reg.GetRegistryValue("PDFSigningSignatureRightX")
64430   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
64440     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
64450       .PDFSigningSignatureRightX = CDbl(Replace$(tStr, ".", GetDecimalChar))
64460      Else
64470       If UseStandard Then
64480        .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
64490       End If
64500     End If
64510    Else
64520     If UseStandard Then
64530      .PDFSigningSignatureRightX = Replace$("200", ".", GetDecimalChar)
64540     End If
64550   End If
64560   tStr = reg.GetRegistryValue("PDFSigningSignatureRightY")
64570   If IsNumeric(Replace$(tStr, ".", GetDecimalChar)) Then
64580     If CDbl(Replace$(tStr, ".", GetDecimalChar)) >= 0 Then
64590       .PDFSigningSignatureRightY = CDbl(Replace$(tStr, ".", GetDecimalChar))
64600      Else
64610       If UseStandard Then
64620        .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
64630       End If
64640     End If
64650    Else
64660     If UseStandard Then
64670      .PDFSigningSignatureRightY = Replace$("200", ".", GetDecimalChar)
64680     End If
64690   End If
64700   tStr = reg.GetRegistryValue("PDFSigningSignatureVisible")
64710   If IsNumeric(tStr) Then
64720     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64730       .PDFSigningSignatureVisible = CLng(tStr)
64740      Else
64750       If UseStandard Then
64760        .PDFSigningSignatureVisible = 0
64770       End If
64780     End If
64790    Else
64800     If UseStandard Then
64810      .PDFSigningSignatureVisible = 0
64820     End If
64830   End If
64840   tStr = reg.GetRegistryValue("PDFSigningSignPDF")
64850   If IsNumeric(tStr) Then
64860     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
64870       .PDFSigningSignPDF = CLng(tStr)
64880      Else
64890       If UseStandard Then
64900        .PDFSigningSignPDF = 0
64910       End If
64920     End If
64930    Else
64940     If UseStandard Then
64950      .PDFSigningSignPDF = 0
64960     End If
64970   End If
64980   reg.SubKey = "Printing\Formats\PS\LanguageLevel"
64990   tStr = reg.GetRegistryValue("EPSLanguageLevel")
65000   If IsNumeric(tStr) Then
65010     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
65020       .EPSLanguageLevel = CLng(tStr)
65030      Else
65040       If UseStandard Then
65050        .EPSLanguageLevel = 2
65060       End If
65070     End If
65080    Else
65090     If UseStandard Then
65100      .EPSLanguageLevel = 2
65110     End If
65120   End If
65130   tStr = reg.GetRegistryValue("PSLanguageLevel")
65140   If IsNumeric(tStr) Then
65150     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
65160       .PSLanguageLevel = CLng(tStr)
65170      Else
65180       If UseStandard Then
65190        .PSLanguageLevel = 2
65200       End If
65210     End If
65220    Else
65230     If UseStandard Then
65240      .PSLanguageLevel = 2
65250     End If
65260   End If
65270   reg.SubKey = "Program"
65280   tStr = reg.GetRegistryValue("AdditionalGhostscriptParameters")
65290   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
65300     .AdditionalGhostscriptParameters = ""
65310    Else
65320     If LenB(tStr) > 0 Then
65330      .AdditionalGhostscriptParameters = tStr
65340     End If
65350   End If
65360   tStr = reg.GetRegistryValue("AdditionalGhostscriptSearchpath")
65370   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
65380     .AdditionalGhostscriptSearchpath = ""
65390    Else
65400     If LenB(tStr) > 0 Then
65410      .AdditionalGhostscriptSearchpath = tStr
65420     End If
65430   End If
65440   tStr = reg.GetRegistryValue("AddWindowsFontpath")
65450   If IsNumeric(tStr) Then
65460     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65470       .AddWindowsFontpath = CLng(tStr)
65480      Else
65490       If UseStandard Then
65500        .AddWindowsFontpath = 1
65510       End If
65520     End If
65530    Else
65540     If UseStandard Then
65550      .AddWindowsFontpath = 1
65560     End If
65570   End If
65580   tStr = reg.GetRegistryValue("AllowSpecialGSCharsInFilenames")
65590   If IsNumeric(tStr) Then
65600     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
65610       .AllowSpecialGSCharsInFilenames = CLng(tStr)
65620      Else
65630       If UseStandard Then
65640        .AllowSpecialGSCharsInFilenames = 1
65650       End If
65660     End If
65670    Else
65680     If UseStandard Then
65690      .AllowSpecialGSCharsInFilenames = 1
65700     End If
65710   End If
65720   tStr = reg.GetRegistryValue("AutosaveDirectory")
65730   If LenB(Trim$(tStr)) > 0 Then
65740     .AutosaveDirectory = CompletePath(tStr)
65750    Else
65760     If UseStandard Then
65770      If InstalledAsServer Then
65780        .AutosaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
65790       Else
65800        .AutosaveDirectory = "<MyFiles>"
65810      End If
65820     End If
65830   End If
65840   tStr = reg.GetRegistryValue("AutosaveFilename")
65850   If LenB(tStr) = 0 And LenB("<DateTime>") > 0 And UseStandard Then
65860     .AutosaveFilename = "<DateTime>"
65870    Else
65880     If LenB(tStr) > 0 Then
65890      .AutosaveFilename = tStr
65900     End If
65910   End If
65920   tStr = reg.GetRegistryValue("AutosaveFormat")
65930   If IsNumeric(tStr) Then
65940     If CLng(tStr) >= 0 And CLng(tStr) <= 14 Then
65950       .AutosaveFormat = CLng(tStr)
65960      Else
65970       If UseStandard Then
65980        .AutosaveFormat = 0
65990       End If
66000     End If
66010    Else
66020     If UseStandard Then
66030      .AutosaveFormat = 0
66040     End If
66050   End If
66060   tStr = reg.GetRegistryValue("AutosaveStartStandardProgram")
66070   If IsNumeric(tStr) Then
66080     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66090       .AutosaveStartStandardProgram = CLng(tStr)
66100      Else
66110       If UseStandard Then
66120        .AutosaveStartStandardProgram = 0
66130       End If
66140     End If
66150    Else
66160     If UseStandard Then
66170      .AutosaveStartStandardProgram = 0
66180     End If
66190   End If
66200   tStr = reg.GetRegistryValue("ClientComputerResolveIPAddress")
66210   If IsNumeric(tStr) Then
66220     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66230       .ClientComputerResolveIPAddress = CLng(tStr)
66240      Else
66250       If UseStandard Then
66260        .ClientComputerResolveIPAddress = 0
66270       End If
66280     End If
66290    Else
66300     If UseStandard Then
66310      .ClientComputerResolveIPAddress = 0
66320     End If
66330   End If
66340   tStr = reg.GetRegistryValue("DisableEmail")
66350   If IsNumeric(tStr) Then
66360     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66370       .DisableEmail = CLng(tStr)
66380      Else
66390       If UseStandard Then
66400        .DisableEmail = 0
66410       End If
66420     End If
66430    Else
66440     If UseStandard Then
66450      .DisableEmail = 0
66460     End If
66470   End If
66480   tStr = reg.GetRegistryValue("DontUseDocumentSettings")
66490   If IsNumeric(tStr) Then
66500     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66510       .DontUseDocumentSettings = CLng(tStr)
66520      Else
66530       If UseStandard Then
66540        .DontUseDocumentSettings = 0
66550       End If
66560     End If
66570    Else
66580     If UseStandard Then
66590      .DontUseDocumentSettings = 0
66600     End If
66610   End If
66620   tStr = reg.GetRegistryValue("FilenameSubstitutions")
66630   If LenB(tStr) = 0 And LenB("Microsoft Word - \.docx\.doc\Microsoft Excel - \.xlsx\.xls\Microsoft PowerPoint - \.pptx\.ppt") > 0 And UseStandard Then
66640     .FilenameSubstitutions = "Microsoft Word - \.docx\.doc\Microsoft Excel - \.xlsx\.xls\Microsoft PowerPoint - \.pptx\.ppt"
66650    Else
66660     If LenB(tStr) > 0 Then
66670      .FilenameSubstitutions = tStr
66680     End If
66690   End If
66700   tStr = reg.GetRegistryValue("FilenameSubstitutionsOnlyInTitle")
66710   If IsNumeric(tStr) Then
66720     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
66730       .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
66740      Else
66750       If UseStandard Then
66760        .FilenameSubstitutionsOnlyInTitle = 1
66770       End If
66780     End If
66790    Else
66800     If UseStandard Then
66810      .FilenameSubstitutionsOnlyInTitle = 1
66820     End If
66830   End If
66840   tStr = reg.GetRegistryValue("Language")
66850   If LenB(tStr) = 0 And LenB("english") > 0 And UseStandard Then
66860     .Language = "english"
66870    Else
66880     If LenB(tStr) > 0 Then
66890      .Language = tStr
66900     End If
66910   End If
66920   tStr = reg.GetRegistryValue("LastSaveDirectory")
66930   If LenB(Trim$(tStr)) > 0 Then
66940     .LastSaveDirectory = CompletePath(tStr)
66950    Else
66960     If UseStandard Then
66970      If InstalledAsServer Then
66980        .LastSaveDirectory = "C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"
66990       Else
67000        .LastSaveDirectory = "<MyFiles>"
67010      End If
67020     End If
67030   End If
67040   tStr = reg.GetRegistryValue("LastUpdateCheck")
67050   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
67060     .LastUpdateCheck = ""
67070    Else
67080     If LenB(tStr) > 0 Then
67090      .LastUpdateCheck = tStr
67100     End If
67110   End If
67120   tStr = reg.GetRegistryValue("Logging")
67130   If IsNumeric(tStr) Then
67140     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67150       .Logging = CLng(tStr)
67160      Else
67170       If UseStandard Then
67180        .Logging = 0
67190       End If
67200     End If
67210    Else
67220     If UseStandard Then
67230      .Logging = 0
67240     End If
67250   End If
67260   tStr = reg.GetRegistryValue("LogLines")
67270   If IsNumeric(tStr) Then
67280     If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
67290       .LogLines = CLng(tStr)
67300      Else
67310       If UseStandard Then
67320        .LogLines = 100
67330       End If
67340     End If
67350    Else
67360     If UseStandard Then
67370      .LogLines = 100
67380     End If
67390   End If
67400   tStr = reg.GetRegistryValue("NoConfirmMessageSwitchingDefaultprinter")
67410   If IsNumeric(tStr) Then
67420     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67430       .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
67440      Else
67450       If UseStandard Then
67460        .NoConfirmMessageSwitchingDefaultprinter = 0
67470       End If
67480     End If
67490    Else
67500     If UseStandard Then
67510      .NoConfirmMessageSwitchingDefaultprinter = 0
67520     End If
67530   End If
67540   tStr = reg.GetRegistryValue("NoProcessingAtStartup")
67550   If IsNumeric(tStr) Then
67560     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67570       .NoProcessingAtStartup = CLng(tStr)
67580      Else
67590       If UseStandard Then
67600        .NoProcessingAtStartup = 0
67610       End If
67620     End If
67630    Else
67640     If UseStandard Then
67650      .NoProcessingAtStartup = 0
67660     End If
67670   End If
67680   tStr = reg.GetRegistryValue("NoPSCheck")
67690   If IsNumeric(tStr) Then
67700     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67710       .NoPSCheck = CLng(tStr)
67720      Else
67730       If UseStandard Then
67740        .NoPSCheck = 0
67750       End If
67760     End If
67770    Else
67780     If UseStandard Then
67790      .NoPSCheck = 0
67800     End If
67810   End If
67820   tStr = reg.GetRegistryValue("OptionsDesign")
67830   If IsNumeric(tStr) Then
67840     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
67850       .OptionsDesign = CLng(tStr)
67860      Else
67870       If UseStandard Then
67880        .OptionsDesign = 0
67890       End If
67900     End If
67910    Else
67920     If UseStandard Then
67930      .OptionsDesign = 0
67940     End If
67950   End If
67960   tStr = reg.GetRegistryValue("OptionsEnabled")
67970   If IsNumeric(tStr) Then
67980     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
67990       .OptionsEnabled = CLng(tStr)
68000      Else
68010       If UseStandard Then
68020        .OptionsEnabled = 1
68030       End If
68040     End If
68050    Else
68060     If UseStandard Then
68070      .OptionsEnabled = 1
68080     End If
68090   End If
68100   tStr = reg.GetRegistryValue("OptionsVisible")
68110   If IsNumeric(tStr) Then
68120     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68130       .OptionsVisible = CLng(tStr)
68140      Else
68150       If UseStandard Then
68160        .OptionsVisible = 1
68170       End If
68180     End If
68190    Else
68200     If UseStandard Then
68210      .OptionsVisible = 1
68220     End If
68230   End If
68240   tStr = reg.GetRegistryValue("PrintAfterSaving")
68250   If IsNumeric(tStr) Then
68260     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68270       .PrintAfterSaving = CLng(tStr)
68280      Else
68290       If UseStandard Then
68300        .PrintAfterSaving = 0
68310       End If
68320     End If
68330    Else
68340     If UseStandard Then
68350      .PrintAfterSaving = 0
68360     End If
68370   End If
68380   tStr = reg.GetRegistryValue("PrintAfterSavingDuplex")
68390   If IsNumeric(tStr) Then
68400     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68410       .PrintAfterSavingDuplex = CLng(tStr)
68420      Else
68430       If UseStandard Then
68440        .PrintAfterSavingDuplex = 0
68450       End If
68460     End If
68470    Else
68480     If UseStandard Then
68490      .PrintAfterSavingDuplex = 0
68500     End If
68510   End If
68520   tStr = reg.GetRegistryValue("PrintAfterSavingNoCancel")
68530   If IsNumeric(tStr) Then
68540     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
68550       .PrintAfterSavingNoCancel = CLng(tStr)
68560      Else
68570       If UseStandard Then
68580        .PrintAfterSavingNoCancel = 0
68590       End If
68600     End If
68610    Else
68620     If UseStandard Then
68630      .PrintAfterSavingNoCancel = 0
68640     End If
68650   End If
68660   tStr = reg.GetRegistryValue("PrintAfterSavingPrinter")
68670   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
68680     .PrintAfterSavingPrinter = ""
68690    Else
68700     If LenB(tStr) > 0 Then
68710      .PrintAfterSavingPrinter = tStr
68720     End If
68730   End If
68740   tStr = reg.GetRegistryValue("PrintAfterSavingQueryUser")
68750   If IsNumeric(tStr) Then
68760     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
68770       .PrintAfterSavingQueryUser = CLng(tStr)
68780      Else
68790       If UseStandard Then
68800        .PrintAfterSavingQueryUser = 0
68810       End If
68820     End If
68830    Else
68840     If UseStandard Then
68850      .PrintAfterSavingQueryUser = 0
68860     End If
68870   End If
68880   tStr = reg.GetRegistryValue("PrintAfterSavingTumble")
68890   If IsNumeric(tStr) Then
68900     If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
68910       .PrintAfterSavingTumble = CLng(tStr)
68920      Else
68930       If UseStandard Then
68940        .PrintAfterSavingTumble = 0
68950       End If
68960     End If
68970    Else
68980     If UseStandard Then
68990      .PrintAfterSavingTumble = 0
69000     End If
69010   End If
69020   tStr = reg.GetRegistryValue("PrinterStop")
69030   If IsNumeric(tStr) Then
69040     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
69050       .PrinterStop = CLng(tStr)
69060      Else
69070       If UseStandard Then
69080        .PrinterStop = 0
69090       End If
69100     End If
69110    Else
69120     If UseStandard Then
69130      .PrinterStop = 0
69140     End If
69150   End If
69160   tStr = reg.GetRegistryValue("PrinterTemppath")
69170   WriteToSpecialLogfile "reg.GetRegistryValue(""PrinterTemppath"")=" & tStr
69180   WriteToSpecialLogfile "Options.PrinterTemppath1=" & .PrinterTemppath
69190   If hkey1 = HKEY_USERS Then
69200     If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then
69210       .PrinterTemppath = tStr
69220      Else
69230       If UseStandard Then
69240         .PrinterTemppath = GetTempPath
69250        Else
69260         .PrinterTemppath = tStr
69270       End If
69280     End If
69290    Else
69300     If LenB(Trim$(tStr)) > 0 Then
69310      If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then
69320        .PrinterTemppath = tStr
69330       Else
69340        MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))
69350        If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then
69360          If UseStandard Then
69370            .PrinterTemppath = GetTempPath
69380           Else
69390            .PrinterTemppath = ""
69400            If NoMsg = False Then
69410             MsgBox "PrinterTemppath: '" & tStr & "' = '" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & "'" & _
             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
69430            End If
69440          End If
69450         Else
69460          .PrinterTemppath = tStr
69470        End If
69480      End If
69490     End If
69500   End If
69510   WriteToSpecialLogfile "Options.PrinterTemppath2=" & .PrinterTemppath
69520   tStr = reg.GetRegistryValue("ProcessPriority")
69530   If IsNumeric(tStr) Then
69540     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
69550       .ProcessPriority = CLng(tStr)
69560      Else
69570       If UseStandard Then
69580        .ProcessPriority = 1
69590       End If
69600     End If
69610    Else
69620     If UseStandard Then
69630      .ProcessPriority = 1
69640     End If
69650   End If
69660   tStr = reg.GetRegistryValue("ProgramFont")
69670   If LenB(tStr) = 0 And LenB("MS Sans Serif") > 0 And UseStandard Then
69680     .ProgramFont = "MS Sans Serif"
69690    Else
69700     If LenB(tStr) > 0 Then
69710      .ProgramFont = tStr
69720     End If
69730   End If
69740   tStr = reg.GetRegistryValue("ProgramFontCharset")
69750   If IsNumeric(tStr) Then
69760     If CLng(tStr) >= 0 Then
69770       .ProgramFontCharset = CLng(tStr)
69780      Else
69790       If UseStandard Then
69800        .ProgramFontCharset = 0
69810       End If
69820     End If
69830    Else
69840     If UseStandard Then
69850      .ProgramFontCharset = 0
69860     End If
69870   End If
69880   tStr = reg.GetRegistryValue("ProgramFontSize")
69890   If IsNumeric(tStr) Then
69900     If CLng(tStr) >= 6 And CLng(tStr) <= 72 Then
69910       .ProgramFontSize = CLng(tStr)
69920      Else
69930       If UseStandard Then
69940        .ProgramFontSize = 8
69950       End If
69960     End If
69970    Else
69980     If UseStandard Then
69990      .ProgramFontSize = 8
70000     End If
70010   End If
70020   tStr = reg.GetRegistryValue("RemoveAllKnownFileExtensions")
70030   If IsNumeric(tStr) Then
70040     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70050       .RemoveAllKnownFileExtensions = CLng(tStr)
70060      Else
70070       If UseStandard Then
70080        .RemoveAllKnownFileExtensions = 1
70090       End If
70100     End If
70110    Else
70120     If UseStandard Then
70130      .RemoveAllKnownFileExtensions = 1
70140     End If
70150   End If
70160   tStr = reg.GetRegistryValue("RemoveSpaces")
70170   If IsNumeric(tStr) Then
70180     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70190       .RemoveSpaces = CLng(tStr)
70200      Else
70210       If UseStandard Then
70220        .RemoveSpaces = 1
70230       End If
70240     End If
70250    Else
70260     If UseStandard Then
70270      .RemoveSpaces = 1
70280     End If
70290   End If
70300   tStr = reg.GetRegistryValue("RunProgramAfterSaving")
70310   If IsNumeric(tStr) Then
70320     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70330       .RunProgramAfterSaving = CLng(tStr)
70340      Else
70350       If UseStandard Then
70360        .RunProgramAfterSaving = 0
70370       End If
70380     End If
70390    Else
70400     If UseStandard Then
70410      .RunProgramAfterSaving = 0
70420     End If
70430   End If
70440   tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramname")
70450   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
70460     .RunProgramAfterSavingProgramname = ""
70470    Else
70480     If LenB(tStr) > 0 Then
70490      .RunProgramAfterSavingProgramname = tStr
70500     End If
70510   End If
70520   tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramParameters")
70530   If LenB(tStr) = 0 And LenB("""<OutputFilename>""") > 0 And UseStandard Then
70540     .RunProgramAfterSavingProgramParameters = """<OutputFilename>"""
70550    Else
70560     If LenB(tStr) > 0 Then
70570      .RunProgramAfterSavingProgramParameters = tStr
70580     End If
70590   End If
70600   tStr = reg.GetRegistryValue("RunProgramAfterSavingWaitUntilReady")
70610   If IsNumeric(tStr) Then
70620     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70630       .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
70640      Else
70650       If UseStandard Then
70660        .RunProgramAfterSavingWaitUntilReady = 1
70670       End If
70680     End If
70690    Else
70700     If UseStandard Then
70710      .RunProgramAfterSavingWaitUntilReady = 1
70720     End If
70730   End If
70740   tStr = reg.GetRegistryValue("RunProgramAfterSavingWindowstyle")
70750   If IsNumeric(tStr) Then
70760     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
70770       .RunProgramAfterSavingWindowstyle = CLng(tStr)
70780      Else
70790       If UseStandard Then
70800        .RunProgramAfterSavingWindowstyle = 1
70810       End If
70820     End If
70830    Else
70840     If UseStandard Then
70850      .RunProgramAfterSavingWindowstyle = 1
70860     End If
70870   End If
70880   tStr = reg.GetRegistryValue("RunProgramBeforeSaving")
70890   If IsNumeric(tStr) Then
70900     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
70910       .RunProgramBeforeSaving = CLng(tStr)
70920      Else
70930       If UseStandard Then
70940        .RunProgramBeforeSaving = 0
70950       End If
70960     End If
70970    Else
70980     If UseStandard Then
70990      .RunProgramBeforeSaving = 0
71000     End If
71010   End If
71020   tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramname")
71030   If LenB(tStr) = 0 And LenB("") > 0 And UseStandard Then
71040     .RunProgramBeforeSavingProgramname = ""
71050    Else
71060     If LenB(tStr) > 0 Then
71070      .RunProgramBeforeSavingProgramname = tStr
71080     End If
71090   End If
71100   tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramParameters")
71110   If LenB(tStr) = 0 And LenB("""<TempFilename>""") > 0 And UseStandard Then
71120     .RunProgramBeforeSavingProgramParameters = """<TempFilename>"""
71130    Else
71140     If LenB(tStr) > 0 Then
71150      .RunProgramBeforeSavingProgramParameters = tStr
71160     End If
71170   End If
71180   tStr = reg.GetRegistryValue("RunProgramBeforeSavingWindowstyle")
71190   If IsNumeric(tStr) Then
71200     If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
71210       .RunProgramBeforeSavingWindowstyle = CLng(tStr)
71220      Else
71230       If UseStandard Then
71240        .RunProgramBeforeSavingWindowstyle = 1
71250       End If
71260     End If
71270    Else
71280     If UseStandard Then
71290      .RunProgramBeforeSavingWindowstyle = 1
71300     End If
71310   End If
71320   tStr = reg.GetRegistryValue("SaveFilename")
71330   If LenB(tStr) = 0 And LenB("<Title>") > 0 And UseStandard Then
71340     .SaveFilename = "<Title>"
71350    Else
71360     If LenB(tStr) > 0 Then
71370      .SaveFilename = tStr
71380     End If
71390   End If
71400   tStr = reg.GetRegistryValue("SendEmailAfterAutoSaving")
71410   If IsNumeric(tStr) Then
71420     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71430       .SendEmailAfterAutoSaving = CLng(tStr)
71440      Else
71450       If UseStandard Then
71460        .SendEmailAfterAutoSaving = 0
71470       End If
71480     End If
71490    Else
71500     If UseStandard Then
71510      .SendEmailAfterAutoSaving = 0
71520     End If
71530   End If
71540   tStr = reg.GetRegistryValue("SendMailMethod")
71550   If IsNumeric(tStr) Then
71560     If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
71570       .SendMailMethod = CLng(tStr)
71580      Else
71590       If UseStandard Then
71600        .SendMailMethod = 0
71610       End If
71620     End If
71630    Else
71640     If UseStandard Then
71650      .SendMailMethod = 0
71660     End If
71670   End If
71680   tStr = reg.GetRegistryValue("ShowAnimation")
71690   If IsNumeric(tStr) Then
71700     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71710       .ShowAnimation = CLng(tStr)
71720      Else
71730       If UseStandard Then
71740        .ShowAnimation = 1
71750       End If
71760     End If
71770    Else
71780     If UseStandard Then
71790      .ShowAnimation = 1
71800     End If
71810   End If
71820   tStr = reg.GetRegistryValue("StartStandardProgram")
71830   If IsNumeric(tStr) Then
71840     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
71850       .StartStandardProgram = CLng(tStr)
71860      Else
71870       If UseStandard Then
71880        .StartStandardProgram = 1
71890       End If
71900     End If
71910    Else
71920     If UseStandard Then
71930      .StartStandardProgram = 1
71940     End If
71950   End If
71960   tStr = reg.GetRegistryValue("Toolbars")
71970   If IsNumeric(tStr) Then
71980     If CLng(tStr) >= 0 Then
71990       .Toolbars = CLng(tStr)
72000      Else
72010       If UseStandard Then
72020        .Toolbars = 1
72030       End If
72040     End If
72050    Else
72060     If UseStandard Then
72070      .Toolbars = 1
72080     End If
72090   End If
72100   tStr = reg.GetRegistryValue("UpdateInterval")
72110   If IsNumeric(tStr) Then
72120     If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
72130       .UpdateInterval = CLng(tStr)
72140      Else
72150       If UseStandard Then
72160        .UpdateInterval = 2
72170       End If
72180     End If
72190    Else
72200     If UseStandard Then
72210      .UpdateInterval = 2
72220     End If
72230   End If
72240   tStr = reg.GetRegistryValue("UseAutosave")
72250   If IsNumeric(tStr) Then
72260     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
72270       .UseAutosave = CLng(tStr)
72280      Else
72290       If UseStandard Then
72300        .UseAutosave = 0
72310       End If
72320     End If
72330    Else
72340     If UseStandard Then
72350      .UseAutosave = 0
72360     End If
72370   End If
72380   tStr = reg.GetRegistryValue("UseAutosaveDirectory")
72390   If IsNumeric(tStr) Then
72400     If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
72410       .UseAutosaveDirectory = CLng(tStr)
72420      Else
72430       If UseStandard Then
72440        .UseAutosaveDirectory = 1
72450       End If
72460     End If
72470    Else
72480     If UseStandard Then
72490      .UseAutosaveDirectory = 1
72500     End If
72510   End If
72520  End With
72530  Set reg = Nothing
72540  ReadOptionsReg = myOptions
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
50060   reg.SubKey = "Ghostscript"
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
50390   reg.SubKey = "Printing"
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
52320   reg.SubKey = "Printing\Formats\Bitmap\Colors"
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
53530   If UCase$(OptionName) = "SVGRESOLUTION" Then
53540    If Not reg.KeyExists Then
53550     reg.CreateKey
53560    End If
53570    reg.SetRegistryValue "SVGResolution", CStr(.SVGResolution), REG_SZ
53580    Set reg = Nothing
53590    Exit Sub
53600   End If
53610   If UCase$(OptionName) = "TIFFCOLORSCOUNT" Then
53620    If Not reg.KeyExists Then
53630     reg.CreateKey
53640    End If
53650    reg.SetRegistryValue "TIFFColorscount", CStr(.TIFFColorscount), REG_SZ
53660    Set reg = Nothing
53670    Exit Sub
53680   End If
53690   If UCase$(OptionName) = "TIFFRESOLUTION" Then
53700    If Not reg.KeyExists Then
53710     reg.CreateKey
53720    End If
53730    reg.SetRegistryValue "TIFFResolution", CStr(.TIFFResolution), REG_SZ
53740    Set reg = Nothing
53750    Exit Sub
53760   End If
53770   reg.SubKey = "Printing\Formats\PDF\Colors"
53780   If UCase$(OptionName) = "PDFCOLORSCMYKTORGB" Then
53790    If Not reg.KeyExists Then
53800     reg.CreateKey
53810    End If
53820    reg.SetRegistryValue "PDFColorsCMYKToRGB", CStr(Abs(.PDFColorsCMYKToRGB)), REG_SZ
53830    Set reg = Nothing
53840    Exit Sub
53850   End If
53860   If UCase$(OptionName) = "PDFCOLORSCOLORMODEL" Then
53870    If Not reg.KeyExists Then
53880     reg.CreateKey
53890    End If
53900    reg.SetRegistryValue "PDFColorsColorModel", CStr(.PDFColorsColorModel), REG_SZ
53910    Set reg = Nothing
53920    Exit Sub
53930   End If
53940   If UCase$(OptionName) = "PDFCOLORSPRESERVEHALFTONE" Then
53950    If Not reg.KeyExists Then
53960     reg.CreateKey
53970    End If
53980    reg.SetRegistryValue "PDFColorsPreserveHalftone", CStr(Abs(.PDFColorsPreserveHalftone)), REG_SZ
53990    Set reg = Nothing
54000    Exit Sub
54010   End If
54020   If UCase$(OptionName) = "PDFCOLORSPRESERVEOVERPRINT" Then
54030    If Not reg.KeyExists Then
54040     reg.CreateKey
54050    End If
54060    reg.SetRegistryValue "PDFColorsPreserveOverprint", CStr(Abs(.PDFColorsPreserveOverprint)), REG_SZ
54070    Set reg = Nothing
54080    Exit Sub
54090   End If
54100   If UCase$(OptionName) = "PDFCOLORSPRESERVETRANSFER" Then
54110    If Not reg.KeyExists Then
54120     reg.CreateKey
54130    End If
54140    reg.SetRegistryValue "PDFColorsPreserveTransfer", CStr(Abs(.PDFColorsPreserveTransfer)), REG_SZ
54150    Set reg = Nothing
54160    Exit Sub
54170   End If
54180   reg.SubKey = "Printing\Formats\PDF\Compression"
54190   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSION" Then
54200    If Not reg.KeyExists Then
54210     reg.CreateKey
54220    End If
54230    reg.SetRegistryValue "PDFCompressionColorCompression", CStr(Abs(.PDFCompressionColorCompression)), REG_SZ
54240    Set reg = Nothing
54250    Exit Sub
54260   End If
54270   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONCHOICE" Then
54280    If Not reg.KeyExists Then
54290     reg.CreateKey
54300    End If
54310    reg.SetRegistryValue "PDFCompressionColorCompressionChoice", CStr(.PDFCompressionColorCompressionChoice), REG_SZ
54320    Set reg = Nothing
54330    Exit Sub
54340   End If
54350   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGHIGHFACTOR" Then
54360    If Not reg.KeyExists Then
54370     reg.CreateKey
54380    End If
54390   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
54400    Set reg = Nothing
54410    Exit Sub
54420   End If
54430   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGLOWFACTOR" Then
54440    If Not reg.KeyExists Then
54450     reg.CreateKey
54460    End If
54470   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
54480    Set reg = Nothing
54490    Exit Sub
54500   End If
54510   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMAXIMUMFACTOR" Then
54520    If Not reg.KeyExists Then
54530     reg.CreateKey
54540    End If
54550   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
54560    Set reg = Nothing
54570    Exit Sub
54580   End If
54590   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMEDIUMFACTOR" Then
54600    If Not reg.KeyExists Then
54610     reg.CreateKey
54620    End If
54630   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
54640    Set reg = Nothing
54650    Exit Sub
54660   End If
54670   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORCOMPRESSIONJPEGMINIMUMFACTOR" Then
54680    If Not reg.KeyExists Then
54690     reg.CreateKey
54700    End If
54710   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
54720    Set reg = Nothing
54730    Exit Sub
54740   End If
54750   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESAMPLE" Then
54760    If Not reg.KeyExists Then
54770     reg.CreateKey
54780    End If
54790    reg.SetRegistryValue "PDFCompressionColorResample", CStr(Abs(.PDFCompressionColorResample)), REG_SZ
54800    Set reg = Nothing
54810    Exit Sub
54820   End If
54830   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESAMPLECHOICE" Then
54840    If Not reg.KeyExists Then
54850     reg.CreateKey
54860    End If
54870    reg.SetRegistryValue "PDFCompressionColorResampleChoice", CStr(.PDFCompressionColorResampleChoice), REG_SZ
54880    Set reg = Nothing
54890    Exit Sub
54900   End If
54910   If UCase$(OptionName) = "PDFCOMPRESSIONCOLORRESOLUTION" Then
54920    If Not reg.KeyExists Then
54930     reg.CreateKey
54940    End If
54950    reg.SetRegistryValue "PDFCompressionColorResolution", CStr(.PDFCompressionColorResolution), REG_SZ
54960    Set reg = Nothing
54970    Exit Sub
54980   End If
54990   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSION" Then
55000    If Not reg.KeyExists Then
55010     reg.CreateKey
55020    End If
55030    reg.SetRegistryValue "PDFCompressionGreyCompression", CStr(Abs(.PDFCompressionGreyCompression)), REG_SZ
55040    Set reg = Nothing
55050    Exit Sub
55060   End If
55070   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONCHOICE" Then
55080    If Not reg.KeyExists Then
55090     reg.CreateKey
55100    End If
55110    reg.SetRegistryValue "PDFCompressionGreyCompressionChoice", CStr(.PDFCompressionGreyCompressionChoice), REG_SZ
55120    Set reg = Nothing
55130    Exit Sub
55140   End If
55150   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGHIGHFACTOR" Then
55160    If Not reg.KeyExists Then
55170     reg.CreateKey
55180    End If
55190   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
55200    Set reg = Nothing
55210    Exit Sub
55220   End If
55230   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGLOWFACTOR" Then
55240    If Not reg.KeyExists Then
55250     reg.CreateKey
55260    End If
55270   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
55280    Set reg = Nothing
55290    Exit Sub
55300   End If
55310   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMAXIMUMFACTOR" Then
55320    If Not reg.KeyExists Then
55330     reg.CreateKey
55340    End If
55350   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
55360    Set reg = Nothing
55370    Exit Sub
55380   End If
55390   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMEDIUMFACTOR" Then
55400    If Not reg.KeyExists Then
55410     reg.CreateKey
55420    End If
55430   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
55440    Set reg = Nothing
55450    Exit Sub
55460   End If
55470   If UCase$(OptionName) = "PDFCOMPRESSIONGREYCOMPRESSIONJPEGMINIMUMFACTOR" Then
55480    If Not reg.KeyExists Then
55490     reg.CreateKey
55500    End If
55510   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
55520    Set reg = Nothing
55530    Exit Sub
55540   End If
55550   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESAMPLE" Then
55560    If Not reg.KeyExists Then
55570     reg.CreateKey
55580    End If
55590    reg.SetRegistryValue "PDFCompressionGreyResample", CStr(Abs(.PDFCompressionGreyResample)), REG_SZ
55600    Set reg = Nothing
55610    Exit Sub
55620   End If
55630   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESAMPLECHOICE" Then
55640    If Not reg.KeyExists Then
55650     reg.CreateKey
55660    End If
55670    reg.SetRegistryValue "PDFCompressionGreyResampleChoice", CStr(.PDFCompressionGreyResampleChoice), REG_SZ
55680    Set reg = Nothing
55690    Exit Sub
55700   End If
55710   If UCase$(OptionName) = "PDFCOMPRESSIONGREYRESOLUTION" Then
55720    If Not reg.KeyExists Then
55730     reg.CreateKey
55740    End If
55750    reg.SetRegistryValue "PDFCompressionGreyResolution", CStr(.PDFCompressionGreyResolution), REG_SZ
55760    Set reg = Nothing
55770    Exit Sub
55780   End If
55790   If UCase$(OptionName) = "PDFCOMPRESSIONMONOCOMPRESSION" Then
55800    If Not reg.KeyExists Then
55810     reg.CreateKey
55820    End If
55830    reg.SetRegistryValue "PDFCompressionMonoCompression", CStr(Abs(.PDFCompressionMonoCompression)), REG_SZ
55840    Set reg = Nothing
55850    Exit Sub
55860   End If
55870   If UCase$(OptionName) = "PDFCOMPRESSIONMONOCOMPRESSIONCHOICE" Then
55880    If Not reg.KeyExists Then
55890     reg.CreateKey
55900    End If
55910    reg.SetRegistryValue "PDFCompressionMonoCompressionChoice", CStr(.PDFCompressionMonoCompressionChoice), REG_SZ
55920    Set reg = Nothing
55930    Exit Sub
55940   End If
55950   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESAMPLE" Then
55960    If Not reg.KeyExists Then
55970     reg.CreateKey
55980    End If
55990    reg.SetRegistryValue "PDFCompressionMonoResample", CStr(Abs(.PDFCompressionMonoResample)), REG_SZ
56000    Set reg = Nothing
56010    Exit Sub
56020   End If
56030   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESAMPLECHOICE" Then
56040    If Not reg.KeyExists Then
56050     reg.CreateKey
56060    End If
56070    reg.SetRegistryValue "PDFCompressionMonoResampleChoice", CStr(.PDFCompressionMonoResampleChoice), REG_SZ
56080    Set reg = Nothing
56090    Exit Sub
56100   End If
56110   If UCase$(OptionName) = "PDFCOMPRESSIONMONORESOLUTION" Then
56120    If Not reg.KeyExists Then
56130     reg.CreateKey
56140    End If
56150    reg.SetRegistryValue "PDFCompressionMonoResolution", CStr(.PDFCompressionMonoResolution), REG_SZ
56160    Set reg = Nothing
56170    Exit Sub
56180   End If
56190   If UCase$(OptionName) = "PDFCOMPRESSIONTEXTCOMPRESSION" Then
56200    If Not reg.KeyExists Then
56210     reg.CreateKey
56220    End If
56230    reg.SetRegistryValue "PDFCompressionTextCompression", CStr(Abs(.PDFCompressionTextCompression)), REG_SZ
56240    Set reg = Nothing
56250    Exit Sub
56260   End If
56270   reg.SubKey = "Printing\Formats\PDF\Fonts"
56280   If UCase$(OptionName) = "PDFFONTSEMBEDALL" Then
56290    If Not reg.KeyExists Then
56300     reg.CreateKey
56310    End If
56320    reg.SetRegistryValue "PDFFontsEmbedAll", CStr(Abs(.PDFFontsEmbedAll)), REG_SZ
56330    Set reg = Nothing
56340    Exit Sub
56350   End If
56360   If UCase$(OptionName) = "PDFFONTSSUBSETFONTS" Then
56370    If Not reg.KeyExists Then
56380     reg.CreateKey
56390    End If
56400    reg.SetRegistryValue "PDFFontsSubSetFonts", CStr(Abs(.PDFFontsSubSetFonts)), REG_SZ
56410    Set reg = Nothing
56420    Exit Sub
56430   End If
56440   If UCase$(OptionName) = "PDFFONTSSUBSETFONTSPERCENT" Then
56450    If Not reg.KeyExists Then
56460     reg.CreateKey
56470    End If
56480    reg.SetRegistryValue "PDFFontsSubSetFontsPercent", CStr(.PDFFontsSubSetFontsPercent), REG_SZ
56490    Set reg = Nothing
56500    Exit Sub
56510   End If
56520   reg.SubKey = "Printing\Formats\PDF\General"
56530   If UCase$(OptionName) = "PDFGENERALASCII85" Then
56540    If Not reg.KeyExists Then
56550     reg.CreateKey
56560    End If
56570    reg.SetRegistryValue "PDFGeneralASCII85", CStr(Abs(.PDFGeneralASCII85)), REG_SZ
56580    Set reg = Nothing
56590    Exit Sub
56600   End If
56610   If UCase$(OptionName) = "PDFGENERALAUTOROTATE" Then
56620    If Not reg.KeyExists Then
56630     reg.CreateKey
56640    End If
56650    reg.SetRegistryValue "PDFGeneralAutorotate", CStr(.PDFGeneralAutorotate), REG_SZ
56660    Set reg = Nothing
56670    Exit Sub
56680   End If
56690   If UCase$(OptionName) = "PDFGENERALCOMPATIBILITY" Then
56700    If Not reg.KeyExists Then
56710     reg.CreateKey
56720    End If
56730    reg.SetRegistryValue "PDFGeneralCompatibility", CStr(.PDFGeneralCompatibility), REG_SZ
56740    Set reg = Nothing
56750    Exit Sub
56760   End If
56770   If UCase$(OptionName) = "PDFGENERALDEFAULT" Then
56780    If Not reg.KeyExists Then
56790     reg.CreateKey
56800    End If
56810    reg.SetRegistryValue "PDFGeneralDefault", CStr(.PDFGeneralDefault), REG_SZ
56820    Set reg = Nothing
56830    Exit Sub
56840   End If
56850   If UCase$(OptionName) = "PDFGENERALOVERPRINT" Then
56860    If Not reg.KeyExists Then
56870     reg.CreateKey
56880    End If
56890    reg.SetRegistryValue "PDFGeneralOverprint", CStr(.PDFGeneralOverprint), REG_SZ
56900    Set reg = Nothing
56910    Exit Sub
56920   End If
56930   If UCase$(OptionName) = "PDFGENERALRESOLUTION" Then
56940    If Not reg.KeyExists Then
56950     reg.CreateKey
56960    End If
56970    reg.SetRegistryValue "PDFGeneralResolution", CStr(.PDFGeneralResolution), REG_SZ
56980    Set reg = Nothing
56990    Exit Sub
57000   End If
57010   If UCase$(OptionName) = "PDFOPTIMIZE" Then
57020    If Not reg.KeyExists Then
57030     reg.CreateKey
57040    End If
57050    reg.SetRegistryValue "PDFOptimize", CStr(Abs(.PDFOptimize)), REG_SZ
57060    Set reg = Nothing
57070    Exit Sub
57080   End If
57090   If UCase$(OptionName) = "PDFUPDATEMETADATA" Then
57100    If Not reg.KeyExists Then
57110     reg.CreateKey
57120    End If
57130    reg.SetRegistryValue "PDFUpdateMetadata", CStr(.PDFUpdateMetadata), REG_SZ
57140    Set reg = Nothing
57150    Exit Sub
57160   End If
57170   reg.SubKey = "Printing\Formats\PDF\Security"
57180   If UCase$(OptionName) = "PDFALLOWASSEMBLY" Then
57190    If Not reg.KeyExists Then
57200     reg.CreateKey
57210    End If
57220    reg.SetRegistryValue "PDFAllowAssembly", CStr(Abs(.PDFAllowAssembly)), REG_SZ
57230    Set reg = Nothing
57240    Exit Sub
57250   End If
57260   If UCase$(OptionName) = "PDFALLOWDEGRADEDPRINTING" Then
57270    If Not reg.KeyExists Then
57280     reg.CreateKey
57290    End If
57300    reg.SetRegistryValue "PDFAllowDegradedPrinting", CStr(Abs(.PDFAllowDegradedPrinting)), REG_SZ
57310    Set reg = Nothing
57320    Exit Sub
57330   End If
57340   If UCase$(OptionName) = "PDFALLOWFILLIN" Then
57350    If Not reg.KeyExists Then
57360     reg.CreateKey
57370    End If
57380    reg.SetRegistryValue "PDFAllowFillIn", CStr(Abs(.PDFAllowFillIn)), REG_SZ
57390    Set reg = Nothing
57400    Exit Sub
57410   End If
57420   If UCase$(OptionName) = "PDFALLOWSCREENREADERS" Then
57430    If Not reg.KeyExists Then
57440     reg.CreateKey
57450    End If
57460    reg.SetRegistryValue "PDFAllowScreenReaders", CStr(Abs(.PDFAllowScreenReaders)), REG_SZ
57470    Set reg = Nothing
57480    Exit Sub
57490   End If
57500   If UCase$(OptionName) = "PDFDISALLOWCOPY" Then
57510    If Not reg.KeyExists Then
57520     reg.CreateKey
57530    End If
57540    reg.SetRegistryValue "PDFDisallowCopy", CStr(Abs(.PDFDisallowCopy)), REG_SZ
57550    Set reg = Nothing
57560    Exit Sub
57570   End If
57580   If UCase$(OptionName) = "PDFDISALLOWMODIFYANNOTATIONS" Then
57590    If Not reg.KeyExists Then
57600     reg.CreateKey
57610    End If
57620    reg.SetRegistryValue "PDFDisallowModifyAnnotations", CStr(Abs(.PDFDisallowModifyAnnotations)), REG_SZ
57630    Set reg = Nothing
57640    Exit Sub
57650   End If
57660   If UCase$(OptionName) = "PDFDISALLOWMODIFYCONTENTS" Then
57670    If Not reg.KeyExists Then
57680     reg.CreateKey
57690    End If
57700    reg.SetRegistryValue "PDFDisallowModifyContents", CStr(Abs(.PDFDisallowModifyContents)), REG_SZ
57710    Set reg = Nothing
57720    Exit Sub
57730   End If
57740   If UCase$(OptionName) = "PDFDISALLOWPRINTING" Then
57750    If Not reg.KeyExists Then
57760     reg.CreateKey
57770    End If
57780    reg.SetRegistryValue "PDFDisallowPrinting", CStr(Abs(.PDFDisallowPrinting)), REG_SZ
57790    Set reg = Nothing
57800    Exit Sub
57810   End If
57820   If UCase$(OptionName) = "PDFENCRYPTOR" Then
57830    If Not reg.KeyExists Then
57840     reg.CreateKey
57850    End If
57860    reg.SetRegistryValue "PDFEncryptor", CStr(.PDFEncryptor), REG_SZ
57870    Set reg = Nothing
57880    Exit Sub
57890   End If
57900   If UCase$(OptionName) = "PDFHIGHENCRYPTION" Then
57910    If Not reg.KeyExists Then
57920     reg.CreateKey
57930    End If
57940    reg.SetRegistryValue "PDFHighEncryption", CStr(Abs(.PDFHighEncryption)), REG_SZ
57950    Set reg = Nothing
57960    Exit Sub
57970   End If
57980   If UCase$(OptionName) = "PDFLOWENCRYPTION" Then
57990    If Not reg.KeyExists Then
58000     reg.CreateKey
58010    End If
58020    reg.SetRegistryValue "PDFLowEncryption", CStr(Abs(.PDFLowEncryption)), REG_SZ
58030    Set reg = Nothing
58040    Exit Sub
58050   End If
58060   If UCase$(OptionName) = "PDFOWNERPASS" Then
58070    If Not reg.KeyExists Then
58080     reg.CreateKey
58090    End If
58100    reg.SetRegistryValue "PDFOwnerPass", CStr(Abs(.PDFOwnerPass)), REG_SZ
58110    Set reg = Nothing
58120    Exit Sub
58130   End If
58140   If UCase$(OptionName) = "PDFOWNERPASSWORDSTRING" Then
58150    If Not reg.KeyExists Then
58160     reg.CreateKey
58170    End If
58180    reg.SetRegistryValue "PDFOwnerPasswordString", CStr(.PDFOwnerPasswordString), REG_SZ
58190    Set reg = Nothing
58200    Exit Sub
58210   End If
58220   If UCase$(OptionName) = "PDFUSERPASS" Then
58230    If Not reg.KeyExists Then
58240     reg.CreateKey
58250    End If
58260    reg.SetRegistryValue "PDFUserPass", CStr(Abs(.PDFUserPass)), REG_SZ
58270    Set reg = Nothing
58280    Exit Sub
58290   End If
58300   If UCase$(OptionName) = "PDFUSERPASSWORDSTRING" Then
58310    If Not reg.KeyExists Then
58320     reg.CreateKey
58330    End If
58340    reg.SetRegistryValue "PDFUserPasswordString", CStr(.PDFUserPasswordString), REG_SZ
58350    Set reg = Nothing
58360    Exit Sub
58370   End If
58380   If UCase$(OptionName) = "PDFUSESECURITY" Then
58390    If Not reg.KeyExists Then
58400     reg.CreateKey
58410    End If
58420    reg.SetRegistryValue "PDFUseSecurity", CStr(Abs(.PDFUseSecurity)), REG_SZ
58430    Set reg = Nothing
58440    Exit Sub
58450   End If
58460   reg.SubKey = "Printing\Formats\PDF\Signing"
58470   If UCase$(OptionName) = "PDFSIGNINGMULTISIGNATURE" Then
58480    If Not reg.KeyExists Then
58490     reg.CreateKey
58500    End If
58510    reg.SetRegistryValue "PDFSigningMultiSignature", CStr(Abs(.PDFSigningMultiSignature)), REG_SZ
58520    Set reg = Nothing
58530    Exit Sub
58540   End If
58550   If UCase$(OptionName) = "PDFSIGNINGPFXFILE" Then
58560    If Not reg.KeyExists Then
58570     reg.CreateKey
58580    End If
58590    reg.SetRegistryValue "PDFSigningPFXFile", CStr(.PDFSigningPFXFile), REG_SZ
58600    Set reg = Nothing
58610    Exit Sub
58620   End If
58630   If UCase$(OptionName) = "PDFSIGNINGPFXFILEPASSWORD" Then
58640    If Not reg.KeyExists Then
58650     reg.CreateKey
58660    End If
58670    reg.SetRegistryValue "PDFSigningPFXFilePassword", CStr(.PDFSigningPFXFilePassword), REG_SZ
58680    Set reg = Nothing
58690    Exit Sub
58700   End If
58710   If UCase$(OptionName) = "PDFSIGNINGSIGNATURECONTACT" Then
58720    If Not reg.KeyExists Then
58730     reg.CreateKey
58740    End If
58750    reg.SetRegistryValue "PDFSigningSignatureContact", CStr(.PDFSigningSignatureContact), REG_SZ
58760    Set reg = Nothing
58770    Exit Sub
58780   End If
58790   If UCase$(OptionName) = "PDFSIGNINGSIGNATURELEFTX" Then
58800    If Not reg.KeyExists Then
58810     reg.CreateKey
58820    End If
58830   reg.SetRegistryValue "PDFSigningSignatureLeftX", Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), REG_SZ
58840    Set reg = Nothing
58850    Exit Sub
58860   End If
58870   If UCase$(OptionName) = "PDFSIGNINGSIGNATURELEFTY" Then
58880    If Not reg.KeyExists Then
58890     reg.CreateKey
58900    End If
58910   reg.SetRegistryValue "PDFSigningSignatureLeftY", Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), REG_SZ
58920    Set reg = Nothing
58930    Exit Sub
58940   End If
58950   If UCase$(OptionName) = "PDFSIGNINGSIGNATURELOCATION" Then
58960    If Not reg.KeyExists Then
58970     reg.CreateKey
58980    End If
58990    reg.SetRegistryValue "PDFSigningSignatureLocation", CStr(.PDFSigningSignatureLocation), REG_SZ
59000    Set reg = Nothing
59010    Exit Sub
59020   End If
59030   If UCase$(OptionName) = "PDFSIGNINGSIGNATUREREASON" Then
59040    If Not reg.KeyExists Then
59050     reg.CreateKey
59060    End If
59070    reg.SetRegistryValue "PDFSigningSignatureReason", CStr(.PDFSigningSignatureReason), REG_SZ
59080    Set reg = Nothing
59090    Exit Sub
59100   End If
59110   If UCase$(OptionName) = "PDFSIGNINGSIGNATURERIGHTX" Then
59120    If Not reg.KeyExists Then
59130     reg.CreateKey
59140    End If
59150   reg.SetRegistryValue "PDFSigningSignatureRightX", Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), REG_SZ
59160    Set reg = Nothing
59170    Exit Sub
59180   End If
59190   If UCase$(OptionName) = "PDFSIGNINGSIGNATURERIGHTY" Then
59200    If Not reg.KeyExists Then
59210     reg.CreateKey
59220    End If
59230   reg.SetRegistryValue "PDFSigningSignatureRightY", Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), REG_SZ
59240    Set reg = Nothing
59250    Exit Sub
59260   End If
59270   If UCase$(OptionName) = "PDFSIGNINGSIGNATUREVISIBLE" Then
59280    If Not reg.KeyExists Then
59290     reg.CreateKey
59300    End If
59310    reg.SetRegistryValue "PDFSigningSignatureVisible", CStr(Abs(.PDFSigningSignatureVisible)), REG_SZ
59320    Set reg = Nothing
59330    Exit Sub
59340   End If
59350   If UCase$(OptionName) = "PDFSIGNINGSIGNPDF" Then
59360    If Not reg.KeyExists Then
59370     reg.CreateKey
59380    End If
59390    reg.SetRegistryValue "PDFSigningSignPDF", CStr(Abs(.PDFSigningSignPDF)), REG_SZ
59400    Set reg = Nothing
59410    Exit Sub
59420   End If
59430   reg.SubKey = "Printing\Formats\PS\LanguageLevel"
59440   If UCase$(OptionName) = "EPSLANGUAGELEVEL" Then
59450    If Not reg.KeyExists Then
59460     reg.CreateKey
59470    End If
59480    reg.SetRegistryValue "EPSLanguageLevel", CStr(.EPSLanguageLevel), REG_SZ
59490    Set reg = Nothing
59500    Exit Sub
59510   End If
59520   If UCase$(OptionName) = "PSLANGUAGELEVEL" Then
59530    If Not reg.KeyExists Then
59540     reg.CreateKey
59550    End If
59560    reg.SetRegistryValue "PSLanguageLevel", CStr(.PSLanguageLevel), REG_SZ
59570    Set reg = Nothing
59580    Exit Sub
59590   End If
59600   reg.SubKey = "Program"
59610   If UCase$(OptionName) = "ADDITIONALGHOSTSCRIPTPARAMETERS" Then
59620    If Not reg.KeyExists Then
59630     reg.CreateKey
59640    End If
59650    reg.SetRegistryValue "AdditionalGhostscriptParameters", CStr(.AdditionalGhostscriptParameters), REG_SZ
59660    Set reg = Nothing
59670    Exit Sub
59680   End If
59690   If UCase$(OptionName) = "ADDITIONALGHOSTSCRIPTSEARCHPATH" Then
59700    If Not reg.KeyExists Then
59710     reg.CreateKey
59720    End If
59730    reg.SetRegistryValue "AdditionalGhostscriptSearchpath", CStr(.AdditionalGhostscriptSearchpath), REG_SZ
59740    Set reg = Nothing
59750    Exit Sub
59760   End If
59770   If UCase$(OptionName) = "ADDWINDOWSFONTPATH" Then
59780    If Not reg.KeyExists Then
59790     reg.CreateKey
59800    End If
59810    reg.SetRegistryValue "AddWindowsFontpath", CStr(Abs(.AddWindowsFontpath)), REG_SZ
59820    Set reg = Nothing
59830    Exit Sub
59840   End If
59850   If UCase$(OptionName) = "ALLOWSPECIALGSCHARSINFILENAMES" Then
59860    If Not reg.KeyExists Then
59870     reg.CreateKey
59880    End If
59890    reg.SetRegistryValue "AllowSpecialGSCharsInFilenames", CStr(Abs(.AllowSpecialGSCharsInFilenames)), REG_SZ
59900    Set reg = Nothing
59910    Exit Sub
59920   End If
59930   If UCase$(OptionName) = "AUTOSAVEDIRECTORY" Then
59940    If Not reg.KeyExists Then
59950     reg.CreateKey
59960    End If
59970    reg.SetRegistryValue "AutosaveDirectory", CStr(.AutosaveDirectory), REG_SZ
59980    Set reg = Nothing
59990    Exit Sub
60000   End If
60010   If UCase$(OptionName) = "AUTOSAVEFILENAME" Then
60020    If Not reg.KeyExists Then
60030     reg.CreateKey
60040    End If
60050    reg.SetRegistryValue "AutosaveFilename", CStr(.AutosaveFilename), REG_SZ
60060    Set reg = Nothing
60070    Exit Sub
60080   End If
60090   If UCase$(OptionName) = "AUTOSAVEFORMAT" Then
60100    If Not reg.KeyExists Then
60110     reg.CreateKey
60120    End If
60130    reg.SetRegistryValue "AutosaveFormat", CStr(.AutosaveFormat), REG_SZ
60140    Set reg = Nothing
60150    Exit Sub
60160   End If
60170   If UCase$(OptionName) = "AUTOSAVESTARTSTANDARDPROGRAM" Then
60180    If Not reg.KeyExists Then
60190     reg.CreateKey
60200    End If
60210    reg.SetRegistryValue "AutosaveStartStandardProgram", CStr(Abs(.AutosaveStartStandardProgram)), REG_SZ
60220    Set reg = Nothing
60230    Exit Sub
60240   End If
60250   If UCase$(OptionName) = "CLIENTCOMPUTERRESOLVEIPADDRESS" Then
60260    If Not reg.KeyExists Then
60270     reg.CreateKey
60280    End If
60290    reg.SetRegistryValue "ClientComputerResolveIPAddress", CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
60300    Set reg = Nothing
60310    Exit Sub
60320   End If
60330   If UCase$(OptionName) = "DISABLEEMAIL" Then
60340    If Not reg.KeyExists Then
60350     reg.CreateKey
60360    End If
60370    reg.SetRegistryValue "DisableEmail", CStr(Abs(.DisableEmail)), REG_SZ
60380    Set reg = Nothing
60390    Exit Sub
60400   End If
60410   If UCase$(OptionName) = "DONTUSEDOCUMENTSETTINGS" Then
60420    If Not reg.KeyExists Then
60430     reg.CreateKey
60440    End If
60450    reg.SetRegistryValue "DontUseDocumentSettings", CStr(Abs(.DontUseDocumentSettings)), REG_SZ
60460    Set reg = Nothing
60470    Exit Sub
60480   End If
60490   If UCase$(OptionName) = "FILENAMESUBSTITUTIONS" Then
60500    If Not reg.KeyExists Then
60510     reg.CreateKey
60520    End If
60530    reg.SetRegistryValue "FilenameSubstitutions", CStr(.FilenameSubstitutions), REG_SZ
60540    Set reg = Nothing
60550    Exit Sub
60560   End If
60570   If UCase$(OptionName) = "FILENAMESUBSTITUTIONSONLYINTITLE" Then
60580    If Not reg.KeyExists Then
60590     reg.CreateKey
60600    End If
60610    reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle", CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
60620    Set reg = Nothing
60630    Exit Sub
60640   End If
60650   If UCase$(OptionName) = "LANGUAGE" Then
60660    If Not reg.KeyExists Then
60670     reg.CreateKey
60680    End If
60690    reg.SetRegistryValue "Language", CStr(.Language), REG_SZ
60700    Set reg = Nothing
60710    Exit Sub
60720   End If
60730   If UCase$(OptionName) = "LASTSAVEDIRECTORY" Then
60740    If Not reg.KeyExists Then
60750     reg.CreateKey
60760    End If
60770    reg.SetRegistryValue "LastSaveDirectory", CStr(.LastSaveDirectory), REG_SZ
60780    Set reg = Nothing
60790    Exit Sub
60800   End If
60810   If UCase$(OptionName) = "LASTUPDATECHECK" Then
60820    If Not reg.KeyExists Then
60830     reg.CreateKey
60840    End If
60850    reg.SetRegistryValue "LastUpdateCheck", CStr(.LastUpdateCheck), REG_SZ
60860    Set reg = Nothing
60870    Exit Sub
60880   End If
60890   If UCase$(OptionName) = "LOGGING" Then
60900    If Not reg.KeyExists Then
60910     reg.CreateKey
60920    End If
60930    reg.SetRegistryValue "Logging", CStr(Abs(.Logging)), REG_SZ
60940    Set reg = Nothing
60950    Exit Sub
60960   End If
60970   If UCase$(OptionName) = "LOGLINES" Then
60980    If Not reg.KeyExists Then
60990     reg.CreateKey
61000    End If
61010    reg.SetRegistryValue "LogLines", CStr(.LogLines), REG_SZ
61020    Set reg = Nothing
61030    Exit Sub
61040   End If
61050   If UCase$(OptionName) = "NOCONFIRMMESSAGESWITCHINGDEFAULTPRINTER" Then
61060    If Not reg.KeyExists Then
61070     reg.CreateKey
61080    End If
61090    reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter", CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
61100    Set reg = Nothing
61110    Exit Sub
61120   End If
61130   If UCase$(OptionName) = "NOPROCESSINGATSTARTUP" Then
61140    If Not reg.KeyExists Then
61150     reg.CreateKey
61160    End If
61170    reg.SetRegistryValue "NoProcessingAtStartup", CStr(Abs(.NoProcessingAtStartup)), REG_SZ
61180    Set reg = Nothing
61190    Exit Sub
61200   End If
61210   If UCase$(OptionName) = "NOPSCHECK" Then
61220    If Not reg.KeyExists Then
61230     reg.CreateKey
61240    End If
61250    reg.SetRegistryValue "NoPSCheck", CStr(Abs(.NoPSCheck)), REG_SZ
61260    Set reg = Nothing
61270    Exit Sub
61280   End If
61290   If UCase$(OptionName) = "OPTIONSDESIGN" Then
61300    If Not reg.KeyExists Then
61310     reg.CreateKey
61320    End If
61330    reg.SetRegistryValue "OptionsDesign", CStr(.OptionsDesign), REG_SZ
61340    Set reg = Nothing
61350    Exit Sub
61360   End If
61370   If UCase$(OptionName) = "OPTIONSENABLED" Then
61380    If Not reg.KeyExists Then
61390     reg.CreateKey
61400    End If
61410    reg.SetRegistryValue "OptionsEnabled", CStr(Abs(.OptionsEnabled)), REG_SZ
61420    Set reg = Nothing
61430    Exit Sub
61440   End If
61450   If UCase$(OptionName) = "OPTIONSVISIBLE" Then
61460    If Not reg.KeyExists Then
61470     reg.CreateKey
61480    End If
61490    reg.SetRegistryValue "OptionsVisible", CStr(Abs(.OptionsVisible)), REG_SZ
61500    Set reg = Nothing
61510    Exit Sub
61520   End If
61530   If UCase$(OptionName) = "PRINTAFTERSAVING" Then
61540    If Not reg.KeyExists Then
61550     reg.CreateKey
61560    End If
61570    reg.SetRegistryValue "PrintAfterSaving", CStr(Abs(.PrintAfterSaving)), REG_SZ
61580    Set reg = Nothing
61590    Exit Sub
61600   End If
61610   If UCase$(OptionName) = "PRINTAFTERSAVINGDUPLEX" Then
61620    If Not reg.KeyExists Then
61630     reg.CreateKey
61640    End If
61650    reg.SetRegistryValue "PrintAfterSavingDuplex", CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
61660    Set reg = Nothing
61670    Exit Sub
61680   End If
61690   If UCase$(OptionName) = "PRINTAFTERSAVINGNOCANCEL" Then
61700    If Not reg.KeyExists Then
61710     reg.CreateKey
61720    End If
61730    reg.SetRegistryValue "PrintAfterSavingNoCancel", CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
61740    Set reg = Nothing
61750    Exit Sub
61760   End If
61770   If UCase$(OptionName) = "PRINTAFTERSAVINGPRINTER" Then
61780    If Not reg.KeyExists Then
61790     reg.CreateKey
61800    End If
61810    reg.SetRegistryValue "PrintAfterSavingPrinter", CStr(.PrintAfterSavingPrinter), REG_SZ
61820    Set reg = Nothing
61830    Exit Sub
61840   End If
61850   If UCase$(OptionName) = "PRINTAFTERSAVINGQUERYUSER" Then
61860    If Not reg.KeyExists Then
61870     reg.CreateKey
61880    End If
61890    reg.SetRegistryValue "PrintAfterSavingQueryUser", CStr(.PrintAfterSavingQueryUser), REG_SZ
61900    Set reg = Nothing
61910    Exit Sub
61920   End If
61930   If UCase$(OptionName) = "PRINTAFTERSAVINGTUMBLE" Then
61940    If Not reg.KeyExists Then
61950     reg.CreateKey
61960    End If
61970    reg.SetRegistryValue "PrintAfterSavingTumble", CStr(.PrintAfterSavingTumble), REG_SZ
61980    Set reg = Nothing
61990    Exit Sub
62000   End If
62010   If UCase$(OptionName) = "PRINTERSTOP" Then
62020    If Not reg.KeyExists Then
62030     reg.CreateKey
62040    End If
62050    reg.SetRegistryValue "PrinterStop", CStr(Abs(.PrinterStop)), REG_SZ
62060    Set reg = Nothing
62070    Exit Sub
62080   End If
62090   If UCase$(OptionName) = "PRINTERTEMPPATH" Then
62100    If Not reg.KeyExists Then
62110     reg.CreateKey
62120    End If
62130    reg.SetRegistryValue "PrinterTemppath", CStr(.PrinterTemppath), REG_SZ
62140    Set reg = Nothing
62150    Exit Sub
62160   End If
62170   If UCase$(OptionName) = "PROCESSPRIORITY" Then
62180    If Not reg.KeyExists Then
62190     reg.CreateKey
62200    End If
62210    reg.SetRegistryValue "ProcessPriority", CStr(.ProcessPriority), REG_SZ
62220    Set reg = Nothing
62230    Exit Sub
62240   End If
62250   If UCase$(OptionName) = "PROGRAMFONT" Then
62260    If Not reg.KeyExists Then
62270     reg.CreateKey
62280    End If
62290    reg.SetRegistryValue "ProgramFont", CStr(.ProgramFont), REG_SZ
62300    Set reg = Nothing
62310    Exit Sub
62320   End If
62330   If UCase$(OptionName) = "PROGRAMFONTCHARSET" Then
62340    If Not reg.KeyExists Then
62350     reg.CreateKey
62360    End If
62370    reg.SetRegistryValue "ProgramFontCharset", CStr(.ProgramFontCharset), REG_SZ
62380    Set reg = Nothing
62390    Exit Sub
62400   End If
62410   If UCase$(OptionName) = "PROGRAMFONTSIZE" Then
62420    If Not reg.KeyExists Then
62430     reg.CreateKey
62440    End If
62450    reg.SetRegistryValue "ProgramFontSize", CStr(.ProgramFontSize), REG_SZ
62460    Set reg = Nothing
62470    Exit Sub
62480   End If
62490   If UCase$(OptionName) = "REMOVEALLKNOWNFILEEXTENSIONS" Then
62500    If Not reg.KeyExists Then
62510     reg.CreateKey
62520    End If
62530    reg.SetRegistryValue "RemoveAllKnownFileExtensions", CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
62540    Set reg = Nothing
62550    Exit Sub
62560   End If
62570   If UCase$(OptionName) = "REMOVESPACES" Then
62580    If Not reg.KeyExists Then
62590     reg.CreateKey
62600    End If
62610    reg.SetRegistryValue "RemoveSpaces", CStr(Abs(.RemoveSpaces)), REG_SZ
62620    Set reg = Nothing
62630    Exit Sub
62640   End If
62650   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVING" Then
62660    If Not reg.KeyExists Then
62670     reg.CreateKey
62680    End If
62690    reg.SetRegistryValue "RunProgramAfterSaving", CStr(Abs(.RunProgramAfterSaving)), REG_SZ
62700    Set reg = Nothing
62710    Exit Sub
62720   End If
62730   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGPROGRAMNAME" Then
62740    If Not reg.KeyExists Then
62750     reg.CreateKey
62760    End If
62770    reg.SetRegistryValue "RunProgramAfterSavingProgramname", CStr(.RunProgramAfterSavingProgramname), REG_SZ
62780    Set reg = Nothing
62790    Exit Sub
62800   End If
62810   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGPROGRAMPARAMETERS" Then
62820    If Not reg.KeyExists Then
62830     reg.CreateKey
62840    End If
62850    reg.SetRegistryValue "RunProgramAfterSavingProgramParameters", CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
62860    Set reg = Nothing
62870    Exit Sub
62880   End If
62890   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGWAITUNTILREADY" Then
62900    If Not reg.KeyExists Then
62910     reg.CreateKey
62920    End If
62930    reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady", CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
62940    Set reg = Nothing
62950    Exit Sub
62960   End If
62970   If UCase$(OptionName) = "RUNPROGRAMAFTERSAVINGWINDOWSTYLE" Then
62980    If Not reg.KeyExists Then
62990     reg.CreateKey
63000    End If
63010    reg.SetRegistryValue "RunProgramAfterSavingWindowstyle", CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
63020    Set reg = Nothing
63030    Exit Sub
63040   End If
63050   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVING" Then
63060    If Not reg.KeyExists Then
63070     reg.CreateKey
63080    End If
63090    reg.SetRegistryValue "RunProgramBeforeSaving", CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
63100    Set reg = Nothing
63110    Exit Sub
63120   End If
63130   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGPROGRAMNAME" Then
63140    If Not reg.KeyExists Then
63150     reg.CreateKey
63160    End If
63170    reg.SetRegistryValue "RunProgramBeforeSavingProgramname", CStr(.RunProgramBeforeSavingProgramname), REG_SZ
63180    Set reg = Nothing
63190    Exit Sub
63200   End If
63210   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGPROGRAMPARAMETERS" Then
63220    If Not reg.KeyExists Then
63230     reg.CreateKey
63240    End If
63250    reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters", CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
63260    Set reg = Nothing
63270    Exit Sub
63280   End If
63290   If UCase$(OptionName) = "RUNPROGRAMBEFORESAVINGWINDOWSTYLE" Then
63300    If Not reg.KeyExists Then
63310     reg.CreateKey
63320    End If
63330    reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle", CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
63340    Set reg = Nothing
63350    Exit Sub
63360   End If
63370   If UCase$(OptionName) = "SAVEFILENAME" Then
63380    If Not reg.KeyExists Then
63390     reg.CreateKey
63400    End If
63410    reg.SetRegistryValue "SaveFilename", CStr(.SaveFilename), REG_SZ
63420    Set reg = Nothing
63430    Exit Sub
63440   End If
63450   If UCase$(OptionName) = "SENDEMAILAFTERAUTOSAVING" Then
63460    If Not reg.KeyExists Then
63470     reg.CreateKey
63480    End If
63490    reg.SetRegistryValue "SendEmailAfterAutoSaving", CStr(Abs(.SendEmailAfterAutoSaving)), REG_SZ
63500    Set reg = Nothing
63510    Exit Sub
63520   End If
63530   If UCase$(OptionName) = "SENDMAILMETHOD" Then
63540    If Not reg.KeyExists Then
63550     reg.CreateKey
63560    End If
63570    reg.SetRegistryValue "SendMailMethod", CStr(.SendMailMethod), REG_SZ
63580    Set reg = Nothing
63590    Exit Sub
63600   End If
63610   If UCase$(OptionName) = "SHOWANIMATION" Then
63620    If Not reg.KeyExists Then
63630     reg.CreateKey
63640    End If
63650    reg.SetRegistryValue "ShowAnimation", CStr(Abs(.ShowAnimation)), REG_SZ
63660    Set reg = Nothing
63670    Exit Sub
63680   End If
63690   If UCase$(OptionName) = "STARTSTANDARDPROGRAM" Then
63700    If Not reg.KeyExists Then
63710     reg.CreateKey
63720    End If
63730    reg.SetRegistryValue "StartStandardProgram", CStr(Abs(.StartStandardProgram)), REG_SZ
63740    Set reg = Nothing
63750    Exit Sub
63760   End If
63770   If UCase$(OptionName) = "TOOLBARS" Then
63780    If Not reg.KeyExists Then
63790     reg.CreateKey
63800    End If
63810    reg.SetRegistryValue "Toolbars", CStr(.Toolbars), REG_SZ
63820    Set reg = Nothing
63830    Exit Sub
63840   End If
63850   If UCase$(OptionName) = "UPDATEINTERVAL" Then
63860    If Not reg.KeyExists Then
63870     reg.CreateKey
63880    End If
63890    reg.SetRegistryValue "UpdateInterval", CStr(.UpdateInterval), REG_SZ
63900    Set reg = Nothing
63910    Exit Sub
63920   End If
63930   If UCase$(OptionName) = "USEAUTOSAVE" Then
63940    If Not reg.KeyExists Then
63950     reg.CreateKey
63960    End If
63970    reg.SetRegistryValue "UseAutosave", CStr(Abs(.UseAutosave)), REG_SZ
63980    Set reg = Nothing
63990    Exit Sub
64000   End If
64010   If UCase$(OptionName) = "USEAUTOSAVEDIRECTORY" Then
64020    If Not reg.KeyExists Then
64030     reg.CreateKey
64040    End If
64050    reg.SetRegistryValue "UseAutosaveDirectory", CStr(Abs(.UseAutosaveDirectory)), REG_SZ
64060    Set reg = Nothing
64070    Exit Sub
64080   End If
64090  End With
64100  Set reg = Nothing
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

Public Sub SaveOptionsREG(sOptions As tOptions, Optional hkey1 As hkey = HKEY_CURRENT_USER, Optional Profilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = hkey1
50040
50050  Profilename = Trim$(Profilename)
50060
50070  If LenB(Profilename) > 0 Then
50080   reg.KeyRoot = "Software\PDFCreator\Profiles\" & Profilename
50090  Else
50100   reg.KeyRoot = "Software\PDFCreator"
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
50920   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
50930   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
50940   reg.SetRegistryValue "PDFCompressionColorCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
50950   reg.SetRegistryValue "PDFCompressionColorResample", CStr(Abs(.PDFCompressionColorResample)), REG_SZ
50960   reg.SetRegistryValue "PDFCompressionColorResampleChoice", CStr(.PDFCompressionColorResampleChoice), REG_SZ
50970   reg.SetRegistryValue "PDFCompressionColorResolution", CStr(.PDFCompressionColorResolution), REG_SZ
50980   reg.SetRegistryValue "PDFCompressionGreyCompression", CStr(Abs(.PDFCompressionGreyCompression)), REG_SZ
50990   reg.SetRegistryValue "PDFCompressionGreyCompressionChoice", CStr(.PDFCompressionGreyCompressionChoice), REG_SZ
51000   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGHighFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, "."), REG_SZ
51010   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGLowFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, "."), REG_SZ
51020   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMaximumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, "."), REG_SZ
51030   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMediumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, "."), REG_SZ
51040   reg.SetRegistryValue "PDFCompressionGreyCompressionJPEGMinimumFactor", Replace$(CStr(.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, "."), REG_SZ
51050   reg.SetRegistryValue "PDFCompressionGreyResample", CStr(Abs(.PDFCompressionGreyResample)), REG_SZ
51060   reg.SetRegistryValue "PDFCompressionGreyResampleChoice", CStr(.PDFCompressionGreyResampleChoice), REG_SZ
51070   reg.SetRegistryValue "PDFCompressionGreyResolution", CStr(.PDFCompressionGreyResolution), REG_SZ
51080   reg.SetRegistryValue "PDFCompressionMonoCompression", CStr(Abs(.PDFCompressionMonoCompression)), REG_SZ
51090   reg.SetRegistryValue "PDFCompressionMonoCompressionChoice", CStr(.PDFCompressionMonoCompressionChoice), REG_SZ
51100   reg.SetRegistryValue "PDFCompressionMonoResample", CStr(Abs(.PDFCompressionMonoResample)), REG_SZ
51110   reg.SetRegistryValue "PDFCompressionMonoResampleChoice", CStr(.PDFCompressionMonoResampleChoice), REG_SZ
51120   reg.SetRegistryValue "PDFCompressionMonoResolution", CStr(.PDFCompressionMonoResolution), REG_SZ
51130   reg.SetRegistryValue "PDFCompressionTextCompression", CStr(Abs(.PDFCompressionTextCompression)), REG_SZ
51140   reg.SubKey = "Printing\Formats\PDF\Fonts"
51150   If Not reg.KeyExists Then
51160    reg.CreateKey
51170   End If
51180   reg.SetRegistryValue "PDFFontsEmbedAll", CStr(Abs(.PDFFontsEmbedAll)), REG_SZ
51190   reg.SetRegistryValue "PDFFontsSubSetFonts", CStr(Abs(.PDFFontsSubSetFonts)), REG_SZ
51200   reg.SetRegistryValue "PDFFontsSubSetFontsPercent", CStr(.PDFFontsSubSetFontsPercent), REG_SZ
51210   reg.SubKey = "Printing\Formats\PDF\General"
51220   If Not reg.KeyExists Then
51230    reg.CreateKey
51240   End If
51250   reg.SetRegistryValue "PDFGeneralASCII85", CStr(Abs(.PDFGeneralASCII85)), REG_SZ
51260   reg.SetRegistryValue "PDFGeneralAutorotate", CStr(.PDFGeneralAutorotate), REG_SZ
51270   reg.SetRegistryValue "PDFGeneralCompatibility", CStr(.PDFGeneralCompatibility), REG_SZ
51280   reg.SetRegistryValue "PDFGeneralDefault", CStr(.PDFGeneralDefault), REG_SZ
51290   reg.SetRegistryValue "PDFGeneralOverprint", CStr(.PDFGeneralOverprint), REG_SZ
51300   reg.SetRegistryValue "PDFGeneralResolution", CStr(.PDFGeneralResolution), REG_SZ
51310   reg.SetRegistryValue "PDFOptimize", CStr(Abs(.PDFOptimize)), REG_SZ
51320   reg.SetRegistryValue "PDFUpdateMetadata", CStr(.PDFUpdateMetadata), REG_SZ
51330   reg.SubKey = "Printing\Formats\PDF\Security"
51340   If Not reg.KeyExists Then
51350    reg.CreateKey
51360   End If
51370   reg.SetRegistryValue "PDFAllowAssembly", CStr(Abs(.PDFAllowAssembly)), REG_SZ
51380   reg.SetRegistryValue "PDFAllowDegradedPrinting", CStr(Abs(.PDFAllowDegradedPrinting)), REG_SZ
51390   reg.SetRegistryValue "PDFAllowFillIn", CStr(Abs(.PDFAllowFillIn)), REG_SZ
51400   reg.SetRegistryValue "PDFAllowScreenReaders", CStr(Abs(.PDFAllowScreenReaders)), REG_SZ
51410   reg.SetRegistryValue "PDFDisallowCopy", CStr(Abs(.PDFDisallowCopy)), REG_SZ
51420   reg.SetRegistryValue "PDFDisallowModifyAnnotations", CStr(Abs(.PDFDisallowModifyAnnotations)), REG_SZ
51430   reg.SetRegistryValue "PDFDisallowModifyContents", CStr(Abs(.PDFDisallowModifyContents)), REG_SZ
51440   reg.SetRegistryValue "PDFDisallowPrinting", CStr(Abs(.PDFDisallowPrinting)), REG_SZ
51450   reg.SetRegistryValue "PDFEncryptor", CStr(.PDFEncryptor), REG_SZ
51460   reg.SetRegistryValue "PDFHighEncryption", CStr(Abs(.PDFHighEncryption)), REG_SZ
51470   reg.SetRegistryValue "PDFLowEncryption", CStr(Abs(.PDFLowEncryption)), REG_SZ
51480   reg.SetRegistryValue "PDFOwnerPass", CStr(Abs(.PDFOwnerPass)), REG_SZ
51490   reg.SetRegistryValue "PDFOwnerPasswordString", CStr(.PDFOwnerPasswordString), REG_SZ
51500   reg.SetRegistryValue "PDFUserPass", CStr(Abs(.PDFUserPass)), REG_SZ
51510   reg.SetRegistryValue "PDFUserPasswordString", CStr(.PDFUserPasswordString), REG_SZ
51520   reg.SetRegistryValue "PDFUseSecurity", CStr(Abs(.PDFUseSecurity)), REG_SZ
51530   reg.SubKey = "Printing\Formats\PDF\Signing"
51540   If Not reg.KeyExists Then
51550    reg.CreateKey
51560   End If
51570   reg.SetRegistryValue "PDFSigningMultiSignature", CStr(Abs(.PDFSigningMultiSignature)), REG_SZ
51580   reg.SetRegistryValue "PDFSigningPFXFile", CStr(.PDFSigningPFXFile), REG_SZ
51590   reg.SetRegistryValue "PDFSigningPFXFilePassword", CStr(.PDFSigningPFXFilePassword), REG_SZ
51600   reg.SetRegistryValue "PDFSigningSignatureContact", CStr(.PDFSigningSignatureContact), REG_SZ
51610   reg.SetRegistryValue "PDFSigningSignatureLeftX", Replace$(CStr(.PDFSigningSignatureLeftX), GetDecimalChar, "."), REG_SZ
51620   reg.SetRegistryValue "PDFSigningSignatureLeftY", Replace$(CStr(.PDFSigningSignatureLeftY), GetDecimalChar, "."), REG_SZ
51630   reg.SetRegistryValue "PDFSigningSignatureLocation", CStr(.PDFSigningSignatureLocation), REG_SZ
51640   reg.SetRegistryValue "PDFSigningSignatureReason", CStr(.PDFSigningSignatureReason), REG_SZ
51650   reg.SetRegistryValue "PDFSigningSignatureRightX", Replace$(CStr(.PDFSigningSignatureRightX), GetDecimalChar, "."), REG_SZ
51660   reg.SetRegistryValue "PDFSigningSignatureRightY", Replace$(CStr(.PDFSigningSignatureRightY), GetDecimalChar, "."), REG_SZ
51670   reg.SetRegistryValue "PDFSigningSignatureVisible", CStr(Abs(.PDFSigningSignatureVisible)), REG_SZ
51680   reg.SetRegistryValue "PDFSigningSignPDF", CStr(Abs(.PDFSigningSignPDF)), REG_SZ
51690   reg.SubKey = "Printing\Formats\PS\LanguageLevel"
51700   If Not reg.KeyExists Then
51710    reg.CreateKey
51720   End If
51730   reg.SetRegistryValue "EPSLanguageLevel", CStr(.EPSLanguageLevel), REG_SZ
51740   reg.SetRegistryValue "PSLanguageLevel", CStr(.PSLanguageLevel), REG_SZ
51750   reg.SubKey = "Program"
51760   If Not reg.KeyExists Then
51770    reg.CreateKey
51780   End If
51790   reg.SetRegistryValue "AdditionalGhostscriptParameters", CStr(.AdditionalGhostscriptParameters), REG_SZ
51800   reg.SetRegistryValue "AdditionalGhostscriptSearchpath", CStr(.AdditionalGhostscriptSearchpath), REG_SZ
51810   reg.SetRegistryValue "AddWindowsFontpath", CStr(Abs(.AddWindowsFontpath)), REG_SZ
51820   reg.SetRegistryValue "AllowSpecialGSCharsInFilenames", CStr(Abs(.AllowSpecialGSCharsInFilenames)), REG_SZ
51830   reg.SetRegistryValue "AutosaveDirectory", CStr(.AutosaveDirectory), REG_SZ
51840   reg.SetRegistryValue "AutosaveFilename", CStr(.AutosaveFilename), REG_SZ
51850   reg.SetRegistryValue "AutosaveFormat", CStr(.AutosaveFormat), REG_SZ
51860   reg.SetRegistryValue "AutosaveStartStandardProgram", CStr(Abs(.AutosaveStartStandardProgram)), REG_SZ
51870   reg.SetRegistryValue "ClientComputerResolveIPAddress", CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
51880   reg.SetRegistryValue "DisableEmail", CStr(Abs(.DisableEmail)), REG_SZ
51890   reg.SetRegistryValue "DontUseDocumentSettings", CStr(Abs(.DontUseDocumentSettings)), REG_SZ
51900   reg.SetRegistryValue "FilenameSubstitutions", CStr(.FilenameSubstitutions), REG_SZ
51910   reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle", CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
51920   reg.SetRegistryValue "Language", CStr(.Language), REG_SZ
51930   reg.SetRegistryValue "LastSaveDirectory", CStr(.LastSaveDirectory), REG_SZ
51940   reg.SetRegistryValue "LastUpdateCheck", CStr(.LastUpdateCheck), REG_SZ
51950   reg.SetRegistryValue "Logging", CStr(Abs(.Logging)), REG_SZ
51960   reg.SetRegistryValue "LogLines", CStr(.LogLines), REG_SZ
51970   reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter", CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
51980   reg.SetRegistryValue "NoProcessingAtStartup", CStr(Abs(.NoProcessingAtStartup)), REG_SZ
51990   reg.SetRegistryValue "NoPSCheck", CStr(Abs(.NoPSCheck)), REG_SZ
52000   reg.SetRegistryValue "OptionsDesign", CStr(.OptionsDesign), REG_SZ
52010   reg.SetRegistryValue "OptionsEnabled", CStr(Abs(.OptionsEnabled)), REG_SZ
52020   reg.SetRegistryValue "OptionsVisible", CStr(Abs(.OptionsVisible)), REG_SZ
52030   reg.SetRegistryValue "PrintAfterSaving", CStr(Abs(.PrintAfterSaving)), REG_SZ
52040   reg.SetRegistryValue "PrintAfterSavingDuplex", CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
52050   reg.SetRegistryValue "PrintAfterSavingNoCancel", CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
52060   reg.SetRegistryValue "PrintAfterSavingPrinter", CStr(.PrintAfterSavingPrinter), REG_SZ
52070   reg.SetRegistryValue "PrintAfterSavingQueryUser", CStr(.PrintAfterSavingQueryUser), REG_SZ
52080   reg.SetRegistryValue "PrintAfterSavingTumble", CStr(.PrintAfterSavingTumble), REG_SZ
52090   reg.SetRegistryValue "PrinterStop", CStr(Abs(.PrinterStop)), REG_SZ
52100   reg.SetRegistryValue "PrinterTemppath", CStr(.PrinterTemppath), REG_SZ
52110   reg.SetRegistryValue "ProcessPriority", CStr(.ProcessPriority), REG_SZ
52120   reg.SetRegistryValue "ProgramFont", CStr(.ProgramFont), REG_SZ
52130   reg.SetRegistryValue "ProgramFontCharset", CStr(.ProgramFontCharset), REG_SZ
52140   reg.SetRegistryValue "ProgramFontSize", CStr(.ProgramFontSize), REG_SZ
52150   reg.SetRegistryValue "RemoveAllKnownFileExtensions", CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
52160   reg.SetRegistryValue "RemoveSpaces", CStr(Abs(.RemoveSpaces)), REG_SZ
52170   reg.SetRegistryValue "RunProgramAfterSaving", CStr(Abs(.RunProgramAfterSaving)), REG_SZ
52180   reg.SetRegistryValue "RunProgramAfterSavingProgramname", CStr(.RunProgramAfterSavingProgramname), REG_SZ
52190   reg.SetRegistryValue "RunProgramAfterSavingProgramParameters", CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
52200   reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady", CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
52210   reg.SetRegistryValue "RunProgramAfterSavingWindowstyle", CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
52220   reg.SetRegistryValue "RunProgramBeforeSaving", CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
52230   reg.SetRegistryValue "RunProgramBeforeSavingProgramname", CStr(.RunProgramBeforeSavingProgramname), REG_SZ
52240   reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters", CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
52250   reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle", CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
52260   reg.SetRegistryValue "SaveFilename", CStr(.SaveFilename), REG_SZ
52270   reg.SetRegistryValue "SendEmailAfterAutoSaving", CStr(Abs(.SendEmailAfterAutoSaving)), REG_SZ
52280   reg.SetRegistryValue "SendMailMethod", CStr(.SendMailMethod), REG_SZ
52290   reg.SetRegistryValue "ShowAnimation", CStr(Abs(.ShowAnimation)), REG_SZ
52300   reg.SetRegistryValue "StartStandardProgram", CStr(Abs(.StartStandardProgram)), REG_SZ
52310   reg.SetRegistryValue "Toolbars", CStr(.Toolbars), REG_SZ
52320   reg.SetRegistryValue "UpdateInterval", CStr(.UpdateInterval), REG_SZ
52330   reg.SetRegistryValue "UseAutosave", CStr(Abs(.UseAutosave)), REG_SZ
52340   reg.SetRegistryValue "UseAutosaveDirectory", CStr(Abs(.UseAutosaveDirectory)), REG_SZ
52350  End With
52360  Set reg = Nothing
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
50130  Options.Language = sLanguage
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
