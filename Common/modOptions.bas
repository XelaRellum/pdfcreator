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
 PDFCompressionColorResample As Long
 PDFCompressionColorResampleChoice As Long
 PDFCompressionColorResolution As Long
 PDFCompressionGreyCompression As Long
 PDFCompressionGreyCompressionChoice As Long
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
50030   .AdditionalGhostscriptParameters = " "
50040   .AdditionalGhostscriptSearchpath = " "
50050   .AddWindowsFontpath = "1"
50060   .AutosaveDirectory = " "
50070   .AutosaveFilename = "<DateTime>"
50080   .AutosaveFormat = "0"
50090   .BitmapResolution = "150"
50100   .BMPColorscount = "1"
50110   .ClientComputerResolveIPAddress = "0"
50120   .DeviceHeightPoints = "-1"
50130   .DeviceWidthPoints = "-1"
50140   Set reg = New clsRegistry
50150   reg.hkey = HKEY_LOCAL_MACHINE
50160   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50170   .DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50180   Set reg = Nothing
50190   Set reg = New clsRegistry
50200   reg.hkey = HKEY_LOCAL_MACHINE
50210   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50220   .DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50230   Set reg = Nothing
50240   Set reg = New clsRegistry
50250   reg.hkey = HKEY_LOCAL_MACHINE
50260   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50270   .DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50280   Set reg = Nothing
50290   Set reg = New clsRegistry
50300   reg.hkey = HKEY_LOCAL_MACHINE
50310   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50320   .DirectoryGhostscriptResource = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryResource"))
50330   Set reg = Nothing
50340   .DisableEmail = "0"
50350   .DontUseDocumentSettings = "0"
50360   .EPSLanguageLevel = "2"
50370   .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
50380   .FilenameSubstitutionsOnlyInTitle = "1"
50390   .JPEGColorscount = "0"
50400   .JPEGQuality = "75"
50410   .Language = "english"
50420   .LastSaveDirectory = GetMyFiles
50430   .Logging = "0"
50440   .LogLines = "100"
50450   .NoConfirmMessageSwitchingDefaultprinter = "0"
50460   .NoProcessingAtStartup = "0"
50470   .OnePagePerFile = "0"
50480   .OptionsDesign = "1"
50490   .OptionsEnabled = "1"
50500   .OptionsVisible = "1"
50510   .Papersize = " "
50520   .PCXColorscount = "0"
50530   .PDFAllowAssembly = "0"
50540   .PDFAllowDegradedPrinting = "0"
50550   .PDFAllowFillIn = "0"
50560   .PDFAllowScreenReaders = "0"
50570   .PDFColorsCMYKToRGB = "1"
50580   .PDFColorsColorModel = "1"
50590   .PDFColorsPreserveHalftone = "0"
50600   .PDFColorsPreserveOverprint = "1"
50610   .PDFColorsPreserveTransfer = "1"
50620   .PDFCompressionColorCompression = "1"
50630   .PDFCompressionColorCompressionChoice = "0"
50640   .PDFCompressionColorResample = "0"
50650   .PDFCompressionColorResampleChoice = "0"
50660   .PDFCompressionColorResolution = "300"
50670   .PDFCompressionGreyCompression = "1"
50680   .PDFCompressionGreyCompressionChoice = "0"
50690   .PDFCompressionGreyResample = "0"
50700   .PDFCompressionGreyResampleChoice = "0"
50710   .PDFCompressionGreyResolution = "300"
50720   .PDFCompressionMonoCompression = "1"
50730   .PDFCompressionMonoCompressionChoice = "0"
50740   .PDFCompressionMonoResample = "0"
50750   .PDFCompressionMonoResampleChoice = "0"
50760   .PDFCompressionMonoResolution = "1200"
50770   .PDFCompressionTextCompression = "1"
50780   .PDFDisallowCopy = "1"
50790   .PDFDisallowModifyAnnotations = "0"
50800   .PDFDisallowModifyContents = "0"
50810   .PDFDisallowPrinting = "0"
50820   .PDFEncryptor = "0"
50830   .PDFFontsEmbedAll = "1"
50840   .PDFFontsSubSetFonts = "1"
50850   .PDFFontsSubSetFontsPercent = "100"
50860   .PDFGeneralASCII85 = "0"
50870   .PDFGeneralAutorotate = "2"
50880   .PDFGeneralCompatibility = "1"
50890   .PDFGeneralOverprint = "0"
50900   .PDFGeneralResolution = "600"
50910   .PDFHighEncryption = "0"
50920   .PDFLowEncryption = "1"
50930   .PDFOptimize = "0"
50940   .PDFOwnerPass = "0"
50950   .PDFOwnerPasswordString = " "
50960   .PDFUserPass = "0"
50970   .PDFUserPasswordString = " "
50980   .PDFUseSecurity = "0"
50990   .PNGColorscount = "0"
51000   .PrintAfterSaving = "0"
51010   .PrintAfterSavingDuplex = "0"
51020   .PrintAfterSavingNoCancel = "0"
51030   .PrintAfterSavingPrinter = " "
51040   .PrintAfterSavingQueryUser = "0"
51050   .PrintAfterSavingTumble = "0"
51060   .PrinterStop = "0"
51070   .PrinterTemppath = GetTempPath
51080   .ProcessPriority = "1"
51090   .ProgramFont = "MS Sans Serif"
51100   .ProgramFontCharset = "0"
51110   .ProgramFontSize = "8"
51120   .PSLanguageLevel = "2"
51130   .RemoveAllKnownFileExtensions = "1"
51140   .RemoveSpaces = "1"
51150   .RunProgramAfterSaving = "0"
51160   .RunProgramAfterSavingProgramname = " "
51170   .RunProgramAfterSavingProgramParameters = " "
51180   .RunProgramAfterSavingWaitUntilReady = "1"
51190   .RunProgramAfterSavingWindowstyle = "1"
51200   .RunProgramBeforeSaving = "0"
51210   .RunProgramBeforeSavingProgramname = " "
51220   .RunProgramBeforeSavingProgramParameters = " "
51230   .RunProgramBeforeSavingWindowstyle = "1"
51240   .SaveFilename = "<Title>"
51250   .SendMailMethod = "0"
51260   .ShowAnimation = "1"
51270   .StampFontColor = "#FF0000"
51280   .StampFontname = "Arial"
51290   .StampFontsize = "48"
51300   .StampOutlineFontthickness = "0"
51310   .StampString = " "
51320   .StampUseOutlineFont = "1"
51330   .StandardAuthor = " "
51340   .StandardCreationdate = " "
51350   .StandardDateformat = "YYYYMMDDHHNNSS"
51360   .StandardKeywords = " "
51370   .StandardMailDomain = " "
51380   .StandardModifydate = " "
51390   .StandardSaveformat = "pdf"
51400   .StandardSubject = " "
51410   .StandardTitle = " "
51420   .StartStandardProgram = "1"
51430   .TIFFColorscount = "0"
51440   .Toolbars = "1"
51450   .UseAutosave = "0"
51460   .UseAutosaveDirectory = "1"
51470   .UseCreationDateNow = "0"
51480   .UseStandardAuthor = "0"
51490  End With
51500  StandardOptions = myOptions
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
50040      myOptions = ReadOptionsINI(myOptions, CompletePath(GetCommonAppData) & "PDFCreator.ini", NoMsg)
50050     Else
50060      myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", HKEY_LOCAL_MACHINE, NoMsg)
50070    End If
50080   Else
50090    If UseINI Then
50100      If Not IsWin9xMe Then
50110        myOptions = ReadOptionsINI(myOptions, CompletePath(GetDefaultAppData) & "PDFCreator.ini", NoMsg)
50120        myOptions = ReadOptionsINI(myOptions, PDFCreatorINIFile, NoMsg, False)
50130       Else
50140        myOptions = ReadOptionsINI(myOptions, PDFCreatorINIFile, NoMsg)
50150      End If
50160      myOptions = ReadOptionsINI(myOptions, CompletePath(GetCommonAppData) & "PDFCreator.ini", NoMsg, False)
50170     Else
50180      If Not IsWin9xMe Then
50190        myOptions = ReadOptionsReg(myOptions, ".DEFAULT\Software\PDFCreator", HKEY_USERS, NoMsg)
50200        myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", hProfile, NoMsg, False)
50210       Else
50220        myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", hProfile, NoMsg)
50230      End If
50240      myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", HKEY_LOCAL_MACHINE, NoMsg, False)
50250    End If
50260  End If
50270  ReadOptions = myOptions
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

Public Function ReadOptionsINI(myOptions As tOptions, PDFCreatorINIFile As String, Optional NoMsg As Boolean = False, Optional UseStandard As Boolean = True) As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI, tstr As String, hOpt As New clsHash
50020  Set ini = New clsINI
50030  ini.Filename = PDFCreatorINIFile
50040  ini.Section = "Options"
50050  If ini.CheckIniFile = False Then
50060   ReadOptionsINI = StandardOptions
50070   Exit Function
50080  End If
50090  ReadINISection PDFCreatorINIFile, "Options", hOpt
50100  With myOptions
50110   tstr = hOpt.Retrieve("AdditionalGhostscriptParameters")
50120   If LenB(tstr) = 0 And LenB("") > 0 Then
50130     If UseStandard Then
50140      .AdditionalGhostscriptParameters = " "
50150     End If
50160    Else
50170     If LenB(tstr) > 0 Then
50180      .AdditionalGhostscriptParameters = tstr
50190     End If
50200   End If
50210   tstr = hOpt.Retrieve("AdditionalGhostscriptSearchpath")
50220   If LenB(tstr) = 0 And LenB("") > 0 Then
50230     If UseStandard Then
50240      .AdditionalGhostscriptSearchpath = " "
50250     End If
50260    Else
50270     If LenB(tstr) > 0 Then
50280      .AdditionalGhostscriptSearchpath = tstr
50290     End If
50300   End If
50310   tstr = hOpt.Retrieve("AddWindowsFontpath")
50320   If IsNumeric(tstr) Then
50330     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
50340       .AddWindowsFontpath = CLng(tstr)
50350      Else
50360       If UseStandard Then
50370        .AddWindowsFontpath = 1
50380       End If
50390     End If
50400    Else
50410     If UseStandard Then
50420      .AddWindowsFontpath = 1
50430     End If
50440   End If
50450   tstr = hOpt.Retrieve("AutosaveDirectory")
50460   If LenB(Trim$(tstr)) > 0 Then
50470     .AutosaveDirectory = CompletePath(tstr)
50480    Else
50490     If UseStandard Then
50500      tstr = GetMyFiles
50510      .AutosaveDirectory = CompletePath(tstr)
50520     End If
50530   End If
50540   tstr = hOpt.Retrieve("AutosaveFilename")
50550   If LenB(tstr) = 0 And LenB("<DateTime>") > 0 Then
50560     If UseStandard Then
50570      .AutosaveFilename = "<DateTime>"
50580     End If
50590    Else
50600     If LenB(tstr) > 0 Then
50610      .AutosaveFilename = tstr
50620     End If
50630   End If
50640   tstr = hOpt.Retrieve("AutosaveFormat")
50650   If IsNumeric(tstr) Then
50660     If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
50670       .AutosaveFormat = CLng(tstr)
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
50780   tstr = hOpt.Retrieve("BitmapResolution")
50790   If IsNumeric(tstr) Then
50800     If CLng(tstr) >= 1 Then
50810       .BitmapResolution = CLng(tstr)
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
50920   tstr = hOpt.Retrieve("BMPColorscount")
50930   If IsNumeric(tstr) Then
50940     If CLng(tstr) >= 0 And CLng(tstr) <= 6 Then
50950       .BMPColorscount = CLng(tstr)
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
51060   tstr = hOpt.Retrieve("ClientComputerResolveIPAddress")
51070   If IsNumeric(tstr) Then
51080     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51090       .ClientComputerResolveIPAddress = CLng(tstr)
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
51200   tstr = hOpt.Retrieve("DeviceHeightPoints")
51210   If IsNumeric(tstr) Then
51220     If CDbl(tstr) >= -1 Then
51230       .DeviceHeightPoints = CDbl(tstr)
51240      Else
51250       If UseStandard Then
51260        .DeviceHeightPoints = -1
51270       End If
51280     End If
51290    Else
51300     If UseStandard Then
51310      .DeviceHeightPoints = -1
51320     End If
51330   End If
51340   tstr = hOpt.Retrieve("DeviceWidthPoints")
51350   If IsNumeric(tstr) Then
51360     If CDbl(tstr) >= -1 Then
51370       .DeviceWidthPoints = CDbl(tstr)
51380      Else
51390       If UseStandard Then
51400        .DeviceWidthPoints = -1
51410       End If
51420     End If
51430    Else
51440     If UseStandard Then
51450      .DeviceWidthPoints = -1
51460     End If
51470   End If
51480   tstr = hOpt.Retrieve("DirectoryGhostscriptBinaries")
51490   If LenB(Trim$(tstr)) > 0 Then
51500     .DirectoryGhostscriptBinaries = CompletePath(tstr)
51510    Else
51520     If UseStandard Then
51530      tstr = App.Path
51540      .DirectoryGhostscriptBinaries = CompletePath(tstr)
51550     End If
51560   End If
51570   tstr = hOpt.Retrieve("DirectoryGhostscriptFonts")
51580   If LenB(Trim$(tstr)) > 0 Then
51590     .DirectoryGhostscriptFonts = CompletePath(tstr)
51600    Else
51610     If UseStandard Then
51620      tstr = App.Path & "\fonts"
51630      .DirectoryGhostscriptFonts = CompletePath(tstr)
51640     End If
51650   End If
51660   tstr = hOpt.Retrieve("DirectoryGhostscriptLibraries")
51670   If LenB(Trim$(tstr)) > 0 Then
51680     .DirectoryGhostscriptLibraries = CompletePath(tstr)
51690    Else
51700     If UseStandard Then
51710      tstr = App.Path & "\lib"
51720      .DirectoryGhostscriptLibraries = CompletePath(tstr)
51730     End If
51740   End If
51750   tstr = hOpt.Retrieve("DirectoryGhostscriptResource")
51760   If LenB(tstr) = 0 And LenB("") > 0 Then
51770     If UseStandard Then
51780      .DirectoryGhostscriptResource = " "
51790     End If
51800    Else
51810     If LenB(tstr) > 0 Then
51820      .DirectoryGhostscriptResource = tstr
51830     End If
51840   End If
51850   tstr = hOpt.Retrieve("DisableEmail")
51860   If IsNumeric(tstr) Then
51870     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51880       .DisableEmail = CLng(tstr)
51890      Else
51900       If UseStandard Then
51910        .DisableEmail = 0
51920       End If
51930     End If
51940    Else
51950     If UseStandard Then
51960      .DisableEmail = 0
51970     End If
51980   End If
51990   tstr = hOpt.Retrieve("DontUseDocumentSettings")
52000   If IsNumeric(tstr) Then
52010     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52020       .DontUseDocumentSettings = CLng(tstr)
52030      Else
52040       If UseStandard Then
52050        .DontUseDocumentSettings = 0
52060       End If
52070     End If
52080    Else
52090     If UseStandard Then
52100      .DontUseDocumentSettings = 0
52110     End If
52120   End If
52130   tstr = hOpt.Retrieve("EPSLanguageLevel")
52140   If IsNumeric(tstr) Then
52150     If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
52160       .EPSLanguageLevel = CLng(tstr)
52170      Else
52180       If UseStandard Then
52190        .EPSLanguageLevel = 2
52200       End If
52210     End If
52220    Else
52230     If UseStandard Then
52240      .EPSLanguageLevel = 2
52250     End If
52260   End If
52270   tstr = hOpt.Retrieve("FilenameSubstitutions")
52280   If LenB(tstr) = 0 And LenB("Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt") > 0 Then
52290     If UseStandard Then
52300      .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
52310     End If
52320    Else
52330     If LenB(tstr) > 0 Then
52340      .FilenameSubstitutions = tstr
52350     End If
52360   End If
52370   tstr = hOpt.Retrieve("FilenameSubstitutionsOnlyInTitle")
52380   If IsNumeric(tstr) Then
52390     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52400       .FilenameSubstitutionsOnlyInTitle = CLng(tstr)
52410      Else
52420       If UseStandard Then
52430        .FilenameSubstitutionsOnlyInTitle = 1
52440       End If
52450     End If
52460    Else
52470     If UseStandard Then
52480      .FilenameSubstitutionsOnlyInTitle = 1
52490     End If
52500   End If
52510   tstr = hOpt.Retrieve("JPEGColorscount")
52520   If IsNumeric(tstr) Then
52530     If CLng(tstr) >= 0 And CLng(tstr) <= 1 Then
52540       .JPEGColorscount = CLng(tstr)
52550      Else
52560       If UseStandard Then
52570        .JPEGColorscount = 0
52580       End If
52590     End If
52600    Else
52610     If UseStandard Then
52620      .JPEGColorscount = 0
52630     End If
52640   End If
52650   tstr = hOpt.Retrieve("JPEGQuality")
52660   If IsNumeric(tstr) Then
52670     If CLng(tstr) >= 0 And CLng(tstr) <= 100 Then
52680       .JPEGQuality = CLng(tstr)
52690      Else
52700       If UseStandard Then
52710        .JPEGQuality = 75
52720       End If
52730     End If
52740    Else
52750     If UseStandard Then
52760      .JPEGQuality = 75
52770     End If
52780   End If
52790   tstr = hOpt.Retrieve("Language")
52800   If LenB(tstr) = 0 And LenB("english") > 0 Then
52810     If UseStandard Then
52820      .Language = "english"
52830     End If
52840    Else
52850     If LenB(tstr) > 0 Then
52860      .Language = tstr
52870     End If
52880   End If
52890   tstr = hOpt.Retrieve("LastSaveDirectory")
52900   If LenB(Trim$(tstr)) > 0 Then
52910     .LastSaveDirectory = CompletePath(tstr)
52920    Else
52930     If UseStandard Then
52940      tstr = GetMyFiles
52950      .LastSaveDirectory = CompletePath(tstr)
52960     End If
52970   End If
52980   tstr = hOpt.Retrieve("Logging")
52990   If IsNumeric(tstr) Then
53000     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53010       .Logging = CLng(tstr)
53020      Else
53030       If UseStandard Then
53040        .Logging = 0
53050       End If
53060     End If
53070    Else
53080     If UseStandard Then
53090      .Logging = 0
53100     End If
53110   End If
53120   tstr = hOpt.Retrieve("LogLines")
53130   If IsNumeric(tstr) Then
53140     If CLng(tstr) >= 100 And CLng(tstr) <= 1000 Then
53150       .LogLines = CLng(tstr)
53160      Else
53170       If UseStandard Then
53180        .LogLines = 100
53190       End If
53200     End If
53210    Else
53220     If UseStandard Then
53230      .LogLines = 100
53240     End If
53250   End If
53260   tstr = hOpt.Retrieve("NoConfirmMessageSwitchingDefaultprinter")
53270   If IsNumeric(tstr) Then
53280     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53290       .NoConfirmMessageSwitchingDefaultprinter = CLng(tstr)
53300      Else
53310       If UseStandard Then
53320        .NoConfirmMessageSwitchingDefaultprinter = 0
53330       End If
53340     End If
53350    Else
53360     If UseStandard Then
53370      .NoConfirmMessageSwitchingDefaultprinter = 0
53380     End If
53390   End If
53400   tstr = hOpt.Retrieve("NoProcessingAtStartup")
53410   If IsNumeric(tstr) Then
53420     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53430       .NoProcessingAtStartup = CLng(tstr)
53440      Else
53450       If UseStandard Then
53460        .NoProcessingAtStartup = 0
53470       End If
53480     End If
53490    Else
53500     If UseStandard Then
53510      .NoProcessingAtStartup = 0
53520     End If
53530   End If
53540   tstr = hOpt.Retrieve("OnePagePerFile")
53550   If IsNumeric(tstr) Then
53560     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53570       .OnePagePerFile = CLng(tstr)
53580      Else
53590       If UseStandard Then
53600        .OnePagePerFile = 0
53610       End If
53620     End If
53630    Else
53640     If UseStandard Then
53650      .OnePagePerFile = 0
53660     End If
53670   End If
53680   tstr = hOpt.Retrieve("OptionsDesign")
53690   If IsNumeric(tstr) Then
53700     If CLng(tstr) >= 1 And CLng(tstr) <= 2 Then
53710       .OptionsDesign = CLng(tstr)
53720      Else
53730       If UseStandard Then
53740        .OptionsDesign = 1
53750       End If
53760     End If
53770    Else
53780     If UseStandard Then
53790      .OptionsDesign = 1
53800     End If
53810   End If
53820   tstr = hOpt.Retrieve("OptionsEnabled")
53830   If IsNumeric(tstr) Then
53840     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53850       .OptionsEnabled = CLng(tstr)
53860      Else
53870       If UseStandard Then
53880        .OptionsEnabled = 1
53890       End If
53900     End If
53910    Else
53920     If UseStandard Then
53930      .OptionsEnabled = 1
53940     End If
53950   End If
53960   tstr = hOpt.Retrieve("OptionsVisible")
53970   If IsNumeric(tstr) Then
53980     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53990       .OptionsVisible = CLng(tstr)
54000      Else
54010       If UseStandard Then
54020        .OptionsVisible = 1
54030       End If
54040     End If
54050    Else
54060     If UseStandard Then
54070      .OptionsVisible = 1
54080     End If
54090   End If
54100   tstr = hOpt.Retrieve("Papersize")
54110   If LenB(tstr) = 0 And LenB("") > 0 Then
54120     If UseStandard Then
54130      .Papersize = " "
54140     End If
54150    Else
54160     If LenB(tstr) > 0 Then
54170      .Papersize = tstr
54180     End If
54190   End If
54200   tstr = hOpt.Retrieve("PCXColorscount")
54210   If IsNumeric(tstr) Then
54220     If CLng(tstr) >= 0 And CLng(tstr) <= 5 Then
54230       .PCXColorscount = CLng(tstr)
54240      Else
54250       If UseStandard Then
54260        .PCXColorscount = 0
54270       End If
54280     End If
54290    Else
54300     If UseStandard Then
54310      .PCXColorscount = 0
54320     End If
54330   End If
54340   tstr = hOpt.Retrieve("PDFAllowAssembly")
54350   If IsNumeric(tstr) Then
54360     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54370       .PDFAllowAssembly = CLng(tstr)
54380      Else
54390       If UseStandard Then
54400        .PDFAllowAssembly = 0
54410       End If
54420     End If
54430    Else
54440     If UseStandard Then
54450      .PDFAllowAssembly = 0
54460     End If
54470   End If
54480   tstr = hOpt.Retrieve("PDFAllowDegradedPrinting")
54490   If IsNumeric(tstr) Then
54500     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54510       .PDFAllowDegradedPrinting = CLng(tstr)
54520      Else
54530       If UseStandard Then
54540        .PDFAllowDegradedPrinting = 0
54550       End If
54560     End If
54570    Else
54580     If UseStandard Then
54590      .PDFAllowDegradedPrinting = 0
54600     End If
54610   End If
54620   tstr = hOpt.Retrieve("PDFAllowFillIn")
54630   If IsNumeric(tstr) Then
54640     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54650       .PDFAllowFillIn = CLng(tstr)
54660      Else
54670       If UseStandard Then
54680        .PDFAllowFillIn = 0
54690       End If
54700     End If
54710    Else
54720     If UseStandard Then
54730      .PDFAllowFillIn = 0
54740     End If
54750   End If
54760   tstr = hOpt.Retrieve("PDFAllowScreenReaders")
54770   If IsNumeric(tstr) Then
54780     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54790       .PDFAllowScreenReaders = CLng(tstr)
54800      Else
54810       If UseStandard Then
54820        .PDFAllowScreenReaders = 0
54830       End If
54840     End If
54850    Else
54860     If UseStandard Then
54870      .PDFAllowScreenReaders = 0
54880     End If
54890   End If
54900   tstr = hOpt.Retrieve("PDFColorsCMYKToRGB")
54910   If IsNumeric(tstr) Then
54920     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54930       .PDFColorsCMYKToRGB = CLng(tstr)
54940      Else
54950       If UseStandard Then
54960        .PDFColorsCMYKToRGB = 1
54970       End If
54980     End If
54990    Else
55000     If UseStandard Then
55010      .PDFColorsCMYKToRGB = 1
55020     End If
55030   End If
55040   tstr = hOpt.Retrieve("PDFColorsColorModel")
55050   If IsNumeric(tstr) Then
55060     If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
55070       .PDFColorsColorModel = CLng(tstr)
55080      Else
55090       If UseStandard Then
55100        .PDFColorsColorModel = 1
55110       End If
55120     End If
55130    Else
55140     If UseStandard Then
55150      .PDFColorsColorModel = 1
55160     End If
55170   End If
55180   tstr = hOpt.Retrieve("PDFColorsPreserveHalftone")
55190   If IsNumeric(tstr) Then
55200     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55210       .PDFColorsPreserveHalftone = CLng(tstr)
55220      Else
55230       If UseStandard Then
55240        .PDFColorsPreserveHalftone = 0
55250       End If
55260     End If
55270    Else
55280     If UseStandard Then
55290      .PDFColorsPreserveHalftone = 0
55300     End If
55310   End If
55320   tstr = hOpt.Retrieve("PDFColorsPreserveOverprint")
55330   If IsNumeric(tstr) Then
55340     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55350       .PDFColorsPreserveOverprint = CLng(tstr)
55360      Else
55370       If UseStandard Then
55380        .PDFColorsPreserveOverprint = 1
55390       End If
55400     End If
55410    Else
55420     If UseStandard Then
55430      .PDFColorsPreserveOverprint = 1
55440     End If
55450   End If
55460   tstr = hOpt.Retrieve("PDFColorsPreserveTransfer")
55470   If IsNumeric(tstr) Then
55480     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55490       .PDFColorsPreserveTransfer = CLng(tstr)
55500      Else
55510       If UseStandard Then
55520        .PDFColorsPreserveTransfer = 1
55530       End If
55540     End If
55550    Else
55560     If UseStandard Then
55570      .PDFColorsPreserveTransfer = 1
55580     End If
55590   End If
55600   tstr = hOpt.Retrieve("PDFCompressionColorCompression")
55610   If IsNumeric(tstr) Then
55620     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55630       .PDFCompressionColorCompression = CLng(tstr)
55640      Else
55650       If UseStandard Then
55660        .PDFCompressionColorCompression = 1
55670       End If
55680     End If
55690    Else
55700     If UseStandard Then
55710      .PDFCompressionColorCompression = 1
55720     End If
55730   End If
55740   tstr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
55750   If IsNumeric(tstr) Then
55760     If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
55770       .PDFCompressionColorCompressionChoice = CLng(tstr)
55780      Else
55790       If UseStandard Then
55800        .PDFCompressionColorCompressionChoice = 0
55810       End If
55820     End If
55830    Else
55840     If UseStandard Then
55850      .PDFCompressionColorCompressionChoice = 0
55860     End If
55870   End If
55880   tstr = hOpt.Retrieve("PDFCompressionColorResample")
55890   If IsNumeric(tstr) Then
55900     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55910       .PDFCompressionColorResample = CLng(tstr)
55920      Else
55930       If UseStandard Then
55940        .PDFCompressionColorResample = 0
55950       End If
55960     End If
55970    Else
55980     If UseStandard Then
55990      .PDFCompressionColorResample = 0
56000     End If
56010   End If
56020   tstr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
56030   If IsNumeric(tstr) Then
56040     If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
56050       .PDFCompressionColorResampleChoice = CLng(tstr)
56060      Else
56070       If UseStandard Then
56080        .PDFCompressionColorResampleChoice = 0
56090       End If
56100     End If
56110    Else
56120     If UseStandard Then
56130      .PDFCompressionColorResampleChoice = 0
56140     End If
56150   End If
56160   tstr = hOpt.Retrieve("PDFCompressionColorResolution")
56170   If IsNumeric(tstr) Then
56180     If CLng(tstr) >= 0 Then
56190       .PDFCompressionColorResolution = CLng(tstr)
56200      Else
56210       If UseStandard Then
56220        .PDFCompressionColorResolution = 300
56230       End If
56240     End If
56250    Else
56260     If UseStandard Then
56270      .PDFCompressionColorResolution = 300
56280     End If
56290   End If
56300   tstr = hOpt.Retrieve("PDFCompressionGreyCompression")
56310   If IsNumeric(tstr) Then
56320     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
56330       .PDFCompressionGreyCompression = CLng(tstr)
56340      Else
56350       If UseStandard Then
56360        .PDFCompressionGreyCompression = 1
56370       End If
56380     End If
56390    Else
56400     If UseStandard Then
56410      .PDFCompressionGreyCompression = 1
56420     End If
56430   End If
56440   tstr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
56450   If IsNumeric(tstr) Then
56460     If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
56470       .PDFCompressionGreyCompressionChoice = CLng(tstr)
56480      Else
56490       If UseStandard Then
56500        .PDFCompressionGreyCompressionChoice = 0
56510       End If
56520     End If
56530    Else
56540     If UseStandard Then
56550      .PDFCompressionGreyCompressionChoice = 0
56560     End If
56570   End If
56580   tstr = hOpt.Retrieve("PDFCompressionGreyResample")
56590   If IsNumeric(tstr) Then
56600     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
56610       .PDFCompressionGreyResample = CLng(tstr)
56620      Else
56630       If UseStandard Then
56640        .PDFCompressionGreyResample = 0
56650       End If
56660     End If
56670    Else
56680     If UseStandard Then
56690      .PDFCompressionGreyResample = 0
56700     End If
56710   End If
56720   tstr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
56730   If IsNumeric(tstr) Then
56740     If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
56750       .PDFCompressionGreyResampleChoice = CLng(tstr)
56760      Else
56770       If UseStandard Then
56780        .PDFCompressionGreyResampleChoice = 0
56790       End If
56800     End If
56810    Else
56820     If UseStandard Then
56830      .PDFCompressionGreyResampleChoice = 0
56840     End If
56850   End If
56860   tstr = hOpt.Retrieve("PDFCompressionGreyResolution")
56870   If IsNumeric(tstr) Then
56880     If CLng(tstr) >= 0 Then
56890       .PDFCompressionGreyResolution = CLng(tstr)
56900      Else
56910       If UseStandard Then
56920        .PDFCompressionGreyResolution = 300
56930       End If
56940     End If
56950    Else
56960     If UseStandard Then
56970      .PDFCompressionGreyResolution = 300
56980     End If
56990   End If
57000   tstr = hOpt.Retrieve("PDFCompressionMonoCompression")
57010   If IsNumeric(tstr) Then
57020     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
57030       .PDFCompressionMonoCompression = CLng(tstr)
57040      Else
57050       If UseStandard Then
57060        .PDFCompressionMonoCompression = 1
57070       End If
57080     End If
57090    Else
57100     If UseStandard Then
57110      .PDFCompressionMonoCompression = 1
57120     End If
57130   End If
57140   tstr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
57150   If IsNumeric(tstr) Then
57160     If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
57170       .PDFCompressionMonoCompressionChoice = CLng(tstr)
57180      Else
57190       If UseStandard Then
57200        .PDFCompressionMonoCompressionChoice = 0
57210       End If
57220     End If
57230    Else
57240     If UseStandard Then
57250      .PDFCompressionMonoCompressionChoice = 0
57260     End If
57270   End If
57280   tstr = hOpt.Retrieve("PDFCompressionMonoResample")
57290   If IsNumeric(tstr) Then
57300     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
57310       .PDFCompressionMonoResample = CLng(tstr)
57320      Else
57330       If UseStandard Then
57340        .PDFCompressionMonoResample = 0
57350       End If
57360     End If
57370    Else
57380     If UseStandard Then
57390      .PDFCompressionMonoResample = 0
57400     End If
57410   End If
57420   tstr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
57430   If IsNumeric(tstr) Then
57440     If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
57450       .PDFCompressionMonoResampleChoice = CLng(tstr)
57460      Else
57470       If UseStandard Then
57480        .PDFCompressionMonoResampleChoice = 0
57490       End If
57500     End If
57510    Else
57520     If UseStandard Then
57530      .PDFCompressionMonoResampleChoice = 0
57540     End If
57550   End If
57560   tstr = hOpt.Retrieve("PDFCompressionMonoResolution")
57570   If IsNumeric(tstr) Then
57580     If CLng(tstr) >= 0 Then
57590       .PDFCompressionMonoResolution = CLng(tstr)
57600      Else
57610       If UseStandard Then
57620        .PDFCompressionMonoResolution = 1200
57630       End If
57640     End If
57650    Else
57660     If UseStandard Then
57670      .PDFCompressionMonoResolution = 1200
57680     End If
57690   End If
57700   tstr = hOpt.Retrieve("PDFCompressionTextCompression")
57710   If IsNumeric(tstr) Then
57720     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
57730       .PDFCompressionTextCompression = CLng(tstr)
57740      Else
57750       If UseStandard Then
57760        .PDFCompressionTextCompression = 1
57770       End If
57780     End If
57790    Else
57800     If UseStandard Then
57810      .PDFCompressionTextCompression = 1
57820     End If
57830   End If
57840   tstr = hOpt.Retrieve("PDFDisallowCopy")
57850   If IsNumeric(tstr) Then
57860     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
57870       .PDFDisallowCopy = CLng(tstr)
57880      Else
57890       If UseStandard Then
57900        .PDFDisallowCopy = 1
57910       End If
57920     End If
57930    Else
57940     If UseStandard Then
57950      .PDFDisallowCopy = 1
57960     End If
57970   End If
57980   tstr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
57990   If IsNumeric(tstr) Then
58000     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
58010       .PDFDisallowModifyAnnotations = CLng(tstr)
58020      Else
58030       If UseStandard Then
58040        .PDFDisallowModifyAnnotations = 0
58050       End If
58060     End If
58070    Else
58080     If UseStandard Then
58090      .PDFDisallowModifyAnnotations = 0
58100     End If
58110   End If
58120   tstr = hOpt.Retrieve("PDFDisallowModifyContents")
58130   If IsNumeric(tstr) Then
58140     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
58150       .PDFDisallowModifyContents = CLng(tstr)
58160      Else
58170       If UseStandard Then
58180        .PDFDisallowModifyContents = 0
58190       End If
58200     End If
58210    Else
58220     If UseStandard Then
58230      .PDFDisallowModifyContents = 0
58240     End If
58250   End If
58260   tstr = hOpt.Retrieve("PDFDisallowPrinting")
58270   If IsNumeric(tstr) Then
58280     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
58290       .PDFDisallowPrinting = CLng(tstr)
58300      Else
58310       If UseStandard Then
58320        .PDFDisallowPrinting = 0
58330       End If
58340     End If
58350    Else
58360     If UseStandard Then
58370      .PDFDisallowPrinting = 0
58380     End If
58390   End If
58400   tstr = hOpt.Retrieve("PDFEncryptor")
58410   If IsNumeric(tstr) Then
58420     If CLng(tstr) >= 0 And CLng(tstr) <= 1 Then
58430       .PDFEncryptor = CLng(tstr)
58440      Else
58450       If UseStandard Then
58460        .PDFEncryptor = 0
58470       End If
58480     End If
58490    Else
58500     If UseStandard Then
58510      .PDFEncryptor = 0
58520     End If
58530   End If
58540   tstr = hOpt.Retrieve("PDFFontsEmbedAll")
58550   If IsNumeric(tstr) Then
58560     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
58570       .PDFFontsEmbedAll = CLng(tstr)
58580      Else
58590       If UseStandard Then
58600        .PDFFontsEmbedAll = 1
58610       End If
58620     End If
58630    Else
58640     If UseStandard Then
58650      .PDFFontsEmbedAll = 1
58660     End If
58670   End If
58680   tstr = hOpt.Retrieve("PDFFontsSubSetFonts")
58690   If IsNumeric(tstr) Then
58700     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
58710       .PDFFontsSubSetFonts = CLng(tstr)
58720      Else
58730       If UseStandard Then
58740        .PDFFontsSubSetFonts = 1
58750       End If
58760     End If
58770    Else
58780     If UseStandard Then
58790      .PDFFontsSubSetFonts = 1
58800     End If
58810   End If
58820   tstr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
58830   If IsNumeric(tstr) Then
58840     If CLng(tstr) >= 0 Then
58850       .PDFFontsSubSetFontsPercent = CLng(tstr)
58860      Else
58870       If UseStandard Then
58880        .PDFFontsSubSetFontsPercent = 100
58890       End If
58900     End If
58910    Else
58920     If UseStandard Then
58930      .PDFFontsSubSetFontsPercent = 100
58940     End If
58950   End If
58960   tstr = hOpt.Retrieve("PDFGeneralASCII85")
58970   If IsNumeric(tstr) Then
58980     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
58990       .PDFGeneralASCII85 = CLng(tstr)
59000      Else
59010       If UseStandard Then
59020        .PDFGeneralASCII85 = 0
59030       End If
59040     End If
59050    Else
59060     If UseStandard Then
59070      .PDFGeneralASCII85 = 0
59080     End If
59090   End If
59100   tstr = hOpt.Retrieve("PDFGeneralAutorotate")
59110   If IsNumeric(tstr) Then
59120     If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
59130       .PDFGeneralAutorotate = CLng(tstr)
59140      Else
59150       If UseStandard Then
59160        .PDFGeneralAutorotate = 2
59170       End If
59180     End If
59190    Else
59200     If UseStandard Then
59210      .PDFGeneralAutorotate = 2
59220     End If
59230   End If
59240   tstr = hOpt.Retrieve("PDFGeneralCompatibility")
59250   If IsNumeric(tstr) Then
59260     If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
59270       .PDFGeneralCompatibility = CLng(tstr)
59280      Else
59290       If UseStandard Then
59300        .PDFGeneralCompatibility = 1
59310       End If
59320     End If
59330    Else
59340     If UseStandard Then
59350      .PDFGeneralCompatibility = 1
59360     End If
59370   End If
59380   tstr = hOpt.Retrieve("PDFGeneralOverprint")
59390   If IsNumeric(tstr) Then
59400     If CLng(tstr) >= 0 And CLng(tstr) <= 1 Then
59410       .PDFGeneralOverprint = CLng(tstr)
59420      Else
59430       If UseStandard Then
59440        .PDFGeneralOverprint = 0
59450       End If
59460     End If
59470    Else
59480     If UseStandard Then
59490      .PDFGeneralOverprint = 0
59500     End If
59510   End If
59520   tstr = hOpt.Retrieve("PDFGeneralResolution")
59530   If IsNumeric(tstr) Then
59540     If CLng(tstr) >= 0 Then
59550       .PDFGeneralResolution = CLng(tstr)
59560      Else
59570       If UseStandard Then
59580        .PDFGeneralResolution = 600
59590       End If
59600     End If
59610    Else
59620     If UseStandard Then
59630      .PDFGeneralResolution = 600
59640     End If
59650   End If
59660   tstr = hOpt.Retrieve("PDFHighEncryption")
59670   If IsNumeric(tstr) Then
59680     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
59690       .PDFHighEncryption = CLng(tstr)
59700      Else
59710       If UseStandard Then
59720        .PDFHighEncryption = 0
59730       End If
59740     End If
59750    Else
59760     If UseStandard Then
59770      .PDFHighEncryption = 0
59780     End If
59790   End If
59800   tstr = hOpt.Retrieve("PDFLowEncryption")
59810   If IsNumeric(tstr) Then
59820     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
59830       .PDFLowEncryption = CLng(tstr)
59840      Else
59850       If UseStandard Then
59860        .PDFLowEncryption = 1
59870       End If
59880     End If
59890    Else
59900     If UseStandard Then
59910      .PDFLowEncryption = 1
59920     End If
59930   End If
59940   tstr = hOpt.Retrieve("PDFOptimize")
59950   If IsNumeric(tstr) Then
59960     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
59970       .PDFOptimize = CLng(tstr)
59980      Else
59990       If UseStandard Then
60000        .PDFOptimize = 0
60010       End If
60020     End If
60030    Else
60040     If UseStandard Then
60050      .PDFOptimize = 0
60060     End If
60070   End If
60080   tstr = hOpt.Retrieve("PDFOwnerPass")
60090   If IsNumeric(tstr) Then
60100     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
60110       .PDFOwnerPass = CLng(tstr)
60120      Else
60130       If UseStandard Then
60140        .PDFOwnerPass = 0
60150       End If
60160     End If
60170    Else
60180     If UseStandard Then
60190      .PDFOwnerPass = 0
60200     End If
60210   End If
60220   tstr = hOpt.Retrieve("PDFOwnerPasswordString")
60230   If LenB(tstr) = 0 And LenB("") > 0 Then
60240     If UseStandard Then
60250      .PDFOwnerPasswordString = " "
60260     End If
60270    Else
60280     If LenB(tstr) > 0 Then
60290      .PDFOwnerPasswordString = tstr
60300     End If
60310   End If
60320   tstr = hOpt.Retrieve("PDFUserPass")
60330   If IsNumeric(tstr) Then
60340     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
60350       .PDFUserPass = CLng(tstr)
60360      Else
60370       If UseStandard Then
60380        .PDFUserPass = 0
60390       End If
60400     End If
60410    Else
60420     If UseStandard Then
60430      .PDFUserPass = 0
60440     End If
60450   End If
60460   tstr = hOpt.Retrieve("PDFUserPasswordString")
60470   If LenB(tstr) = 0 And LenB("") > 0 Then
60480     If UseStandard Then
60490      .PDFUserPasswordString = " "
60500     End If
60510    Else
60520     If LenB(tstr) > 0 Then
60530      .PDFUserPasswordString = tstr
60540     End If
60550   End If
60560   tstr = hOpt.Retrieve("PDFUseSecurity")
60570   If IsNumeric(tstr) Then
60580     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
60590       .PDFUseSecurity = CLng(tstr)
60600      Else
60610       If UseStandard Then
60620        .PDFUseSecurity = 0
60630       End If
60640     End If
60650    Else
60660     If UseStandard Then
60670      .PDFUseSecurity = 0
60680     End If
60690   End If
60700   tstr = hOpt.Retrieve("PNGColorscount")
60710   If IsNumeric(tstr) Then
60720     If CLng(tstr) >= 0 And CLng(tstr) <= 4 Then
60730       .PNGColorscount = CLng(tstr)
60740      Else
60750       If UseStandard Then
60760        .PNGColorscount = 0
60770       End If
60780     End If
60790    Else
60800     If UseStandard Then
60810      .PNGColorscount = 0
60820     End If
60830   End If
60840   tstr = hOpt.Retrieve("PrintAfterSaving")
60850   If IsNumeric(tstr) Then
60860     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
60870       .PrintAfterSaving = CLng(tstr)
60880      Else
60890       If UseStandard Then
60900        .PrintAfterSaving = 0
60910       End If
60920     End If
60930    Else
60940     If UseStandard Then
60950      .PrintAfterSaving = 0
60960     End If
60970   End If
60980   tstr = hOpt.Retrieve("PrintAfterSavingDuplex")
60990   If IsNumeric(tstr) Then
61000     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
61010       .PrintAfterSavingDuplex = CLng(tstr)
61020      Else
61030       If UseStandard Then
61040        .PrintAfterSavingDuplex = 0
61050       End If
61060     End If
61070    Else
61080     If UseStandard Then
61090      .PrintAfterSavingDuplex = 0
61100     End If
61110   End If
61120   tstr = hOpt.Retrieve("PrintAfterSavingNoCancel")
61130   If IsNumeric(tstr) Then
61140     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
61150       .PrintAfterSavingNoCancel = CLng(tstr)
61160      Else
61170       If UseStandard Then
61180        .PrintAfterSavingNoCancel = 0
61190       End If
61200     End If
61210    Else
61220     If UseStandard Then
61230      .PrintAfterSavingNoCancel = 0
61240     End If
61250   End If
61260   tstr = hOpt.Retrieve("PrintAfterSavingPrinter")
61270   If LenB(tstr) = 0 And LenB("") > 0 Then
61280     If UseStandard Then
61290      .PrintAfterSavingPrinter = " "
61300     End If
61310    Else
61320     If LenB(tstr) > 0 Then
61330      .PrintAfterSavingPrinter = tstr
61340     End If
61350   End If
61360   tstr = hOpt.Retrieve("PrintAfterSavingQueryUser")
61370   If IsNumeric(tstr) Then
61380     If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
61390       .PrintAfterSavingQueryUser = CLng(tstr)
61400      Else
61410       If UseStandard Then
61420        .PrintAfterSavingQueryUser = 0
61430       End If
61440     End If
61450    Else
61460     If UseStandard Then
61470      .PrintAfterSavingQueryUser = 0
61480     End If
61490   End If
61500   tstr = hOpt.Retrieve("PrintAfterSavingTumble")
61510   If IsNumeric(tstr) Then
61520     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
61530       .PrintAfterSavingTumble = CLng(tstr)
61540      Else
61550       If UseStandard Then
61560        .PrintAfterSavingTumble = 0
61570       End If
61580     End If
61590    Else
61600     If UseStandard Then
61610      .PrintAfterSavingTumble = 0
61620     End If
61630   End If
61640   tstr = hOpt.Retrieve("PrinterStop")
61650   If IsNumeric(tstr) Then
61660     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
61670       .PrinterStop = CLng(tstr)
61680      Else
61690       If UseStandard Then
61700        .PrinterStop = 0
61710       End If
61720     End If
61730    Else
61740     If UseStandard Then
61750      .PrinterStop = 0
61760     End If
61770   End If
61780   tstr = hOpt.Retrieve("PrinterTemppath")
61790   If LenB(Trim$(tstr)) > 0 Then
61800    If DirExists(GetSubstFilename2(tstr, False)) = True Then
61810      .PrinterTemppath = tstr
61820     Else
61830      MakePath ResolveEnvironment(GetSubstFilename2(tstr, False))
61840      If DirExists(ResolveEnvironment(GetSubstFilename2(tstr, False))) = False Then
61850        If UseStandard Then
61860          .PrinterTemppath = GetTempPath
61870         Else
61880          .PrinterTemppath = ""
61890          If NoMsg = False Then
61900           MsgBox "PrinterTemppath: '" & tstr & "' = '" & ResolveEnvironment(GetSubstFilename2(tstr, False)) & "'" & _
           vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
61920          End If
61930        End If
61940       Else
61950        .PrinterTemppath = tstr
61960      End If
61970    End If
61980   End If
61990   tstr = hOpt.Retrieve("ProcessPriority")
62000   If IsNumeric(tstr) Then
62010     If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
62020       .ProcessPriority = CLng(tstr)
62030      Else
62040       If UseStandard Then
62050        .ProcessPriority = 1
62060       End If
62070     End If
62080    Else
62090     If UseStandard Then
62100      .ProcessPriority = 1
62110     End If
62120   End If
62130   tstr = hOpt.Retrieve("ProgramFont")
62140   If LenB(tstr) = 0 And LenB("MS Sans Serif") > 0 Then
62150     If UseStandard Then
62160      .ProgramFont = "MS Sans Serif"
62170     End If
62180    Else
62190     If LenB(tstr) > 0 Then
62200      .ProgramFont = tstr
62210     End If
62220   End If
62230   tstr = hOpt.Retrieve("ProgramFontCharset")
62240   If IsNumeric(tstr) Then
62250     If CLng(tstr) >= 0 Then
62260       .ProgramFontCharset = CLng(tstr)
62270      Else
62280       If UseStandard Then
62290        .ProgramFontCharset = 0
62300       End If
62310     End If
62320    Else
62330     If UseStandard Then
62340      .ProgramFontCharset = 0
62350     End If
62360   End If
62370   tstr = hOpt.Retrieve("ProgramFontSize")
62380   If IsNumeric(tstr) Then
62390     If CLng(tstr) >= 1 And CLng(tstr) <= 72 Then
62400       .ProgramFontSize = CLng(tstr)
62410      Else
62420       If UseStandard Then
62430        .ProgramFontSize = 8
62440       End If
62450     End If
62460    Else
62470     If UseStandard Then
62480      .ProgramFontSize = 8
62490     End If
62500   End If
62510   tstr = hOpt.Retrieve("PSLanguageLevel")
62520   If IsNumeric(tstr) Then
62530     If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
62540       .PSLanguageLevel = CLng(tstr)
62550      Else
62560       If UseStandard Then
62570        .PSLanguageLevel = 2
62580       End If
62590     End If
62600    Else
62610     If UseStandard Then
62620      .PSLanguageLevel = 2
62630     End If
62640   End If
62650   tstr = hOpt.Retrieve("RemoveAllKnownFileExtensions")
62660   If IsNumeric(tstr) Then
62670     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
62680       .RemoveAllKnownFileExtensions = CLng(tstr)
62690      Else
62700       If UseStandard Then
62710        .RemoveAllKnownFileExtensions = 1
62720       End If
62730     End If
62740    Else
62750     If UseStandard Then
62760      .RemoveAllKnownFileExtensions = 1
62770     End If
62780   End If
62790   tstr = hOpt.Retrieve("RemoveSpaces")
62800   If IsNumeric(tstr) Then
62810     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
62820       .RemoveSpaces = CLng(tstr)
62830      Else
62840       If UseStandard Then
62850        .RemoveSpaces = 1
62860       End If
62870     End If
62880    Else
62890     If UseStandard Then
62900      .RemoveSpaces = 1
62910     End If
62920   End If
62930   tstr = hOpt.Retrieve("RunProgramAfterSaving")
62940   If IsNumeric(tstr) Then
62950     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
62960       .RunProgramAfterSaving = CLng(tstr)
62970      Else
62980       If UseStandard Then
62990        .RunProgramAfterSaving = 0
63000       End If
63010     End If
63020    Else
63030     If UseStandard Then
63040      .RunProgramAfterSaving = 0
63050     End If
63060   End If
63070   tstr = hOpt.Retrieve("RunProgramAfterSavingProgramname")
63080   If LenB(tstr) = 0 And LenB("") > 0 Then
63090     If UseStandard Then
63100      .RunProgramAfterSavingProgramname = " "
63110     End If
63120    Else
63130     If LenB(tstr) > 0 Then
63140      .RunProgramAfterSavingProgramname = tstr
63150     End If
63160   End If
63170   tstr = hOpt.Retrieve("RunProgramAfterSavingProgramParameters")
63180   If LenB(tstr) = 0 And LenB("") > 0 Then
63190     If UseStandard Then
63200      .RunProgramAfterSavingProgramParameters = " "
63210     End If
63220    Else
63230     If LenB(tstr) > 0 Then
63240      .RunProgramAfterSavingProgramParameters = tstr
63250     End If
63260   End If
63270   tstr = hOpt.Retrieve("RunProgramAfterSavingWaitUntilReady")
63280   If IsNumeric(tstr) Then
63290     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
63300       .RunProgramAfterSavingWaitUntilReady = CLng(tstr)
63310      Else
63320       If UseStandard Then
63330        .RunProgramAfterSavingWaitUntilReady = 1
63340       End If
63350     End If
63360    Else
63370     If UseStandard Then
63380      .RunProgramAfterSavingWaitUntilReady = 1
63390     End If
63400   End If
63410   tstr = hOpt.Retrieve("RunProgramAfterSavingWindowstyle")
63420   If IsNumeric(tstr) Then
63430     If CLng(tstr) >= 0 And CLng(tstr) <= 6 Then
63440       .RunProgramAfterSavingWindowstyle = CLng(tstr)
63450      Else
63460       If UseStandard Then
63470        .RunProgramAfterSavingWindowstyle = 1
63480       End If
63490     End If
63500    Else
63510     If UseStandard Then
63520      .RunProgramAfterSavingWindowstyle = 1
63530     End If
63540   End If
63550   tstr = hOpt.Retrieve("RunProgramBeforeSaving")
63560   If IsNumeric(tstr) Then
63570     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
63580       .RunProgramBeforeSaving = CLng(tstr)
63590      Else
63600       If UseStandard Then
63610        .RunProgramBeforeSaving = 0
63620       End If
63630     End If
63640    Else
63650     If UseStandard Then
63660      .RunProgramBeforeSaving = 0
63670     End If
63680   End If
63690   tstr = hOpt.Retrieve("RunProgramBeforeSavingProgramname")
63700   If LenB(tstr) = 0 And LenB("") > 0 Then
63710     If UseStandard Then
63720      .RunProgramBeforeSavingProgramname = " "
63730     End If
63740    Else
63750     If LenB(tstr) > 0 Then
63760      .RunProgramBeforeSavingProgramname = tstr
63770     End If
63780   End If
63790   tstr = hOpt.Retrieve("RunProgramBeforeSavingProgramParameters")
63800   If LenB(tstr) = 0 And LenB("") > 0 Then
63810     If UseStandard Then
63820      .RunProgramBeforeSavingProgramParameters = " "
63830     End If
63840    Else
63850     If LenB(tstr) > 0 Then
63860      .RunProgramBeforeSavingProgramParameters = tstr
63870     End If
63880   End If
63890   tstr = hOpt.Retrieve("RunProgramBeforeSavingWindowstyle")
63900   If IsNumeric(tstr) Then
63910     If CLng(tstr) >= 0 And CLng(tstr) <= 6 Then
63920       .RunProgramBeforeSavingWindowstyle = CLng(tstr)
63930      Else
63940       If UseStandard Then
63950        .RunProgramBeforeSavingWindowstyle = 1
63960       End If
63970     End If
63980    Else
63990     If UseStandard Then
64000      .RunProgramBeforeSavingWindowstyle = 1
64010     End If
64020   End If
64030   tstr = hOpt.Retrieve("SaveFilename")
64040   If LenB(tstr) = 0 And LenB("<Title>") > 0 Then
64050     If UseStandard Then
64060      .SaveFilename = "<Title>"
64070     End If
64080    Else
64090     If LenB(tstr) > 0 Then
64100      .SaveFilename = tstr
64110     End If
64120   End If
64130   tstr = hOpt.Retrieve("SendMailMethod")
64140   If IsNumeric(tstr) Then
64150     If CLng(tstr) >= 0 Then
64160       .SendMailMethod = CLng(tstr)
64170      Else
64180       If UseStandard Then
64190        .SendMailMethod = 0
64200       End If
64210     End If
64220    Else
64230     If UseStandard Then
64240      .SendMailMethod = 0
64250     End If
64260   End If
64270   tstr = hOpt.Retrieve("ShowAnimation")
64280   If IsNumeric(tstr) Then
64290     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
64300       .ShowAnimation = CLng(tstr)
64310      Else
64320       If UseStandard Then
64330        .ShowAnimation = 1
64340       End If
64350     End If
64360    Else
64370     If UseStandard Then
64380      .ShowAnimation = 1
64390     End If
64400   End If
64410   tstr = hOpt.Retrieve("StampFontColor")
64420   If LenB(tstr) = 0 And LenB("#FF0000") > 0 Then
64430     If UseStandard Then
64440      .StampFontColor = "#FF0000"
64450     End If
64460    Else
64470     If LenB(tstr) > 0 Then
64480      .StampFontColor = tstr
64490     End If
64500   End If
64510   tstr = hOpt.Retrieve("StampFontname")
64520   If LenB(tstr) = 0 And LenB("Arial") > 0 Then
64530     If UseStandard Then
64540      .StampFontname = "Arial"
64550     End If
64560    Else
64570     If LenB(tstr) > 0 Then
64580      .StampFontname = tstr
64590     End If
64600   End If
64610   tstr = hOpt.Retrieve("StampFontsize")
64620   If IsNumeric(tstr) Then
64630     If CLng(tstr) >= 1 Then
64640       .StampFontsize = CLng(tstr)
64650      Else
64660       If UseStandard Then
64670        .StampFontsize = 48
64680       End If
64690     End If
64700    Else
64710     If UseStandard Then
64720      .StampFontsize = 48
64730     End If
64740   End If
64750   tstr = hOpt.Retrieve("StampOutlineFontthickness")
64760   If IsNumeric(tstr) Then
64770     If CLng(tstr) >= 0 Then
64780       .StampOutlineFontthickness = CLng(tstr)
64790      Else
64800       If UseStandard Then
64810        .StampOutlineFontthickness = 0
64820       End If
64830     End If
64840    Else
64850     If UseStandard Then
64860      .StampOutlineFontthickness = 0
64870     End If
64880   End If
64890   tstr = hOpt.Retrieve("StampString")
64900   If LenB(tstr) = 0 And LenB("") > 0 Then
64910     If UseStandard Then
64920      .StampString = " "
64930     End If
64940    Else
64950     If LenB(tstr) > 0 Then
64960      .StampString = tstr
64970     End If
64980   End If
64990   tstr = hOpt.Retrieve("StampUseOutlineFont")
65000   If IsNumeric(tstr) Then
65010     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
65020       .StampUseOutlineFont = CLng(tstr)
65030      Else
65040       If UseStandard Then
65050        .StampUseOutlineFont = 1
65060       End If
65070     End If
65080    Else
65090     If UseStandard Then
65100      .StampUseOutlineFont = 1
65110     End If
65120   End If
65130   tstr = hOpt.Retrieve("StandardAuthor")
65140   If LenB(tstr) = 0 And LenB("") > 0 Then
65150     If UseStandard Then
65160      .StandardAuthor = " "
65170     End If
65180    Else
65190     If LenB(tstr) > 0 Then
65200      .StandardAuthor = tstr
65210     End If
65220   End If
65230   tstr = hOpt.Retrieve("StandardCreationdate")
65240   If LenB(tstr) = 0 And LenB("") > 0 Then
65250     If UseStandard Then
65260      .StandardCreationdate = " "
65270     End If
65280    Else
65290     If LenB(tstr) > 0 Then
65300      .StandardCreationdate = tstr
65310     End If
65320   End If
65330   tstr = hOpt.Retrieve("StandardDateformat")
65340   If LenB(tstr) = 0 And LenB("YYYYMMDDHHNNSS") > 0 Then
65350     If UseStandard Then
65360      .StandardDateformat = "YYYYMMDDHHNNSS"
65370     End If
65380    Else
65390     If LenB(tstr) > 0 Then
65400      .StandardDateformat = tstr
65410     End If
65420   End If
65430   tstr = hOpt.Retrieve("StandardKeywords")
65440   If LenB(tstr) = 0 And LenB("") > 0 Then
65450     If UseStandard Then
65460      .StandardKeywords = " "
65470     End If
65480    Else
65490     If LenB(tstr) > 0 Then
65500      .StandardKeywords = tstr
65510     End If
65520   End If
65530   tstr = hOpt.Retrieve("StandardMailDomain")
65540   If LenB(tstr) = 0 And LenB("") > 0 Then
65550     If UseStandard Then
65560      .StandardMailDomain = " "
65570     End If
65580    Else
65590     If LenB(tstr) > 0 Then
65600      .StandardMailDomain = tstr
65610     End If
65620   End If
65630   tstr = hOpt.Retrieve("StandardModifydate")
65640   If LenB(tstr) = 0 And LenB("") > 0 Then
65650     If UseStandard Then
65660      .StandardModifydate = " "
65670     End If
65680    Else
65690     If LenB(tstr) > 0 Then
65700      .StandardModifydate = tstr
65710     End If
65720   End If
65730   tstr = hOpt.Retrieve("StandardSaveformat")
65740   If LenB(tstr) = 0 And LenB("pdf") > 0 Then
65750     If UseStandard Then
65760      .StandardSaveformat = "pdf"
65770     End If
65780    Else
65790     If LenB(tstr) > 0 Then
65800      .StandardSaveformat = tstr
65810     End If
65820   End If
65830   tstr = hOpt.Retrieve("StandardSubject")
65840   If LenB(tstr) = 0 And LenB("") > 0 Then
65850     If UseStandard Then
65860      .StandardSubject = " "
65870     End If
65880    Else
65890     If LenB(tstr) > 0 Then
65900      .StandardSubject = tstr
65910     End If
65920   End If
65930   tstr = hOpt.Retrieve("StandardTitle")
65940   If LenB(tstr) = 0 And LenB("") > 0 Then
65950     If UseStandard Then
65960      .StandardTitle = " "
65970     End If
65980    Else
65990     If LenB(tstr) > 0 Then
66000      .StandardTitle = tstr
66010     End If
66020   End If
66030   tstr = hOpt.Retrieve("StartStandardProgram")
66040   If IsNumeric(tstr) Then
66050     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
66060       .StartStandardProgram = CLng(tstr)
66070      Else
66080       If UseStandard Then
66090        .StartStandardProgram = 1
66100       End If
66110     End If
66120    Else
66130     If UseStandard Then
66140      .StartStandardProgram = 1
66150     End If
66160   End If
66170   tstr = hOpt.Retrieve("TIFFColorscount")
66180   If IsNumeric(tstr) Then
66190     If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
66200       .TIFFColorscount = CLng(tstr)
66210      Else
66220       If UseStandard Then
66230        .TIFFColorscount = 0
66240       End If
66250     End If
66260    Else
66270     If UseStandard Then
66280      .TIFFColorscount = 0
66290     End If
66300   End If
66310   tstr = hOpt.Retrieve("Toolbars")
66320   If IsNumeric(tstr) Then
66330     If CLng(tstr) >= 0 Then
66340       .Toolbars = CLng(tstr)
66350      Else
66360       If UseStandard Then
66370        .Toolbars = 1
66380       End If
66390     End If
66400    Else
66410     If UseStandard Then
66420      .Toolbars = 1
66430     End If
66440   End If
66450   tstr = hOpt.Retrieve("UseAutosave")
66460   If IsNumeric(tstr) Then
66470     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
66480       .UseAutosave = CLng(tstr)
66490      Else
66500       If UseStandard Then
66510        .UseAutosave = 0
66520       End If
66530     End If
66540    Else
66550     If UseStandard Then
66560      .UseAutosave = 0
66570     End If
66580   End If
66590   tstr = hOpt.Retrieve("UseAutosaveDirectory")
66600   If IsNumeric(tstr) Then
66610     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
66620       .UseAutosaveDirectory = CLng(tstr)
66630      Else
66640       If UseStandard Then
66650        .UseAutosaveDirectory = 1
66660       End If
66670     End If
66680    Else
66690     If UseStandard Then
66700      .UseAutosaveDirectory = 1
66710     End If
66720   End If
66730   tstr = hOpt.Retrieve("UseCreationDateNow")
66740   If IsNumeric(tstr) Then
66750     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
66760       .UseCreationDateNow = CLng(tstr)
66770      Else
66780       If UseStandard Then
66790        .UseCreationDateNow = 0
66800       End If
66810     End If
66820    Else
66830     If UseStandard Then
66840      .UseCreationDateNow = 0
66850     End If
66860   End If
66870   tstr = hOpt.Retrieve("UseStandardAuthor")
66880   If IsNumeric(tstr) Then
66890     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
66900       .UseStandardAuthor = CLng(tstr)
66910      Else
66920       If UseStandard Then
66930        .UseStandardAuthor = 0
66940       End If
66950     End If
66960    Else
66970     If UseStandard Then
66980      .UseStandardAuthor = 0
66990     End If
67000   End If
67010  End With
67020  Set ini = Nothing
67030  ReadOptionsINI = myOptions
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

Public Sub SaveOptions(sOptions As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If InstalledAsServer Then
50020    If UseINI Then
50030      SaveOptionsINI sOptions, CompletePath(GetCommonAppData) & "PDFCreator.ini"
50040     Else
50050      SaveOptionsREG sOptions, HKEY_LOCAL_MACHINE
50060    End If
50070   Else
50080    If UseINI Then
50090      SaveOptionsINI sOptions, PDFCreatorINIFile
50100     Else
50110      SaveOptionsREG sOptions
50120    End If
50130  End If
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
50180   ini.SaveKey CStr(.DeviceHeightPoints), "DeviceHeightPoints"
50190   ini.SaveKey CStr(.DeviceWidthPoints), "DeviceWidthPoints"
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
50370   ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50380   ini.SaveKey CStr(.OptionsDesign), "OptionsDesign"
50390   ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50400   ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50410   ini.SaveKey CStr(.Papersize), "Papersize"
50420   ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50430   ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50440   ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50450   ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50460   ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50470   ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50480   ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50490   ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50500   ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50510   ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50520   ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50530   ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50540   ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50550   ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50560   ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50570   ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50580   ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50590   ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50600   ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50610   ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50620   ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50630   ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50640   ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50650   ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50660   ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50670   ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50680   ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50690   ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50700   ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50710   ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50720   ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50730   ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50740   ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50750   ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50760   ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50770   ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50780   ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50790   ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50800   ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
50810   ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
50820   ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
50830   ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
50840   ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
50850   ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
50860   ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
50870   ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
50880   ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
50890   ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
50900   ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
50910   ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
50920   ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
50930   ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
50940   ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
50950   ini.SaveKey CStr(Abs(.PrintAfterSavingTumble)), "PrintAfterSavingTumble"
50960   ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
50970   ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
50980   ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
50990   ini.SaveKey CStr(.ProgramFont), "ProgramFont"
51000   ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
51010   ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
51020   ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
51030   ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
51040   ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
51050   ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
51060   ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
51070   ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
51080   ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51090   ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51100   ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51110   ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51120   ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51130   ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51140   ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51150   ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51160   ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51170   ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51180   ini.SaveKey CStr(.StampFontname), "StampFontname"
51190   ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51200   ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51210   ini.SaveKey CStr(.StampString), "StampString"
51220   ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51230   ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51240   ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51250   ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51260   ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51270   ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51280   ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51290   ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51300   ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51310   ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51320   ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51330   ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51340   ini.SaveKey CStr(.Toolbars), "Toolbars"
51350   ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51360   ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51370   ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51380   ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51390  End With
51400  Set ini = Nothing
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
50010  Dim reg As clsRegistry, tstr As String
50020  Set reg = New clsRegistry
50030  reg.hkey = hkey1
50040  reg.KeyRoot = KeyRoot
50050  With myOptions
50060   reg.Subkey = "Ghostscript"
50070   tstr = reg.GetRegistryValue("DirectoryGhostscriptBinaries")
50080   If LenB(Trim$(tstr)) > 0 Then
50090     .DirectoryGhostscriptBinaries = CompletePath(tstr)
50100    Else
50110     If UseStandard Then
50120      tstr = App.Path
50130      .DirectoryGhostscriptBinaries = CompletePath(tstr)
50140     End If
50150   End If
50160   tstr = reg.GetRegistryValue("DirectoryGhostscriptFonts")
50170   If LenB(Trim$(tstr)) > 0 Then
50180     .DirectoryGhostscriptFonts = CompletePath(tstr)
50190    Else
50200     If UseStandard Then
50210      tstr = App.Path & "\fonts"
50220      .DirectoryGhostscriptFonts = CompletePath(tstr)
50230     End If
50240   End If
50250   tstr = reg.GetRegistryValue("DirectoryGhostscriptLibraries")
50260   If LenB(Trim$(tstr)) > 0 Then
50270     .DirectoryGhostscriptLibraries = CompletePath(tstr)
50280    Else
50290     If UseStandard Then
50300      tstr = App.Path & "\lib"
50310      .DirectoryGhostscriptLibraries = CompletePath(tstr)
50320     End If
50330   End If
50340   tstr = reg.GetRegistryValue("DirectoryGhostscriptResource")
50350   If LenB(tstr) = 0 And LenB("") > 0 Then
50360     If UseStandard Then
50370      .DirectoryGhostscriptResource = " "
50380     End If
50390    Else
50400     If LenB(tstr) > 0 Then
50410      .DirectoryGhostscriptResource = tstr
50420     End If
50430   End If
50440   reg.Subkey = "Printing"
50450   tstr = reg.GetRegistryValue("DeviceHeightPoints")
50460   If IsNumeric(tstr) Then
50470     If CDbl(tstr) >= -1 Then
50480       .DeviceHeightPoints = CDbl(tstr)
50490      Else
50500       If UseStandard Then
50510        .DeviceHeightPoints = -1
50520       End If
50530     End If
50540    Else
50550     If UseStandard Then
50560      .DeviceHeightPoints = -1
50570     End If
50580   End If
50590   tstr = reg.GetRegistryValue("DeviceWidthPoints")
50600   If IsNumeric(tstr) Then
50610     If CDbl(tstr) >= -1 Then
50620       .DeviceWidthPoints = CDbl(tstr)
50630      Else
50640       If UseStandard Then
50650        .DeviceWidthPoints = -1
50660       End If
50670     End If
50680    Else
50690     If UseStandard Then
50700      .DeviceWidthPoints = -1
50710     End If
50720   End If
50730   tstr = reg.GetRegistryValue("OnePagePerFile")
50740   If IsNumeric(tstr) Then
50750     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
50760       .OnePagePerFile = CLng(tstr)
50770      Else
50780       If UseStandard Then
50790        .OnePagePerFile = 0
50800       End If
50810     End If
50820    Else
50830     If UseStandard Then
50840      .OnePagePerFile = 0
50850     End If
50860   End If
50870   tstr = reg.GetRegistryValue("Papersize")
50880   If LenB(tstr) = 0 And LenB("") > 0 Then
50890     If UseStandard Then
50900      .Papersize = " "
50910     End If
50920    Else
50930     If LenB(tstr) > 0 Then
50940      .Papersize = tstr
50950     End If
50960   End If
50970   tstr = reg.GetRegistryValue("StampFontColor")
50980   If LenB(tstr) = 0 And LenB("#FF0000") > 0 Then
50990     If UseStandard Then
51000      .StampFontColor = "#FF0000"
51010     End If
51020    Else
51030     If LenB(tstr) > 0 Then
51040      .StampFontColor = tstr
51050     End If
51060   End If
51070   tstr = reg.GetRegistryValue("StampFontname")
51080   If LenB(tstr) = 0 And LenB("Arial") > 0 Then
51090     If UseStandard Then
51100      .StampFontname = "Arial"
51110     End If
51120    Else
51130     If LenB(tstr) > 0 Then
51140      .StampFontname = tstr
51150     End If
51160   End If
51170   tstr = reg.GetRegistryValue("StampFontsize")
51180   If IsNumeric(tstr) Then
51190     If CLng(tstr) >= 1 Then
51200       .StampFontsize = CLng(tstr)
51210      Else
51220       If UseStandard Then
51230        .StampFontsize = 48
51240       End If
51250     End If
51260    Else
51270     If UseStandard Then
51280      .StampFontsize = 48
51290     End If
51300   End If
51310   tstr = reg.GetRegistryValue("StampOutlineFontthickness")
51320   If IsNumeric(tstr) Then
51330     If CLng(tstr) >= 0 Then
51340       .StampOutlineFontthickness = CLng(tstr)
51350      Else
51360       If UseStandard Then
51370        .StampOutlineFontthickness = 0
51380       End If
51390     End If
51400    Else
51410     If UseStandard Then
51420      .StampOutlineFontthickness = 0
51430     End If
51440   End If
51450   tstr = reg.GetRegistryValue("StampString")
51460   If LenB(tstr) = 0 And LenB("") > 0 Then
51470     If UseStandard Then
51480      .StampString = " "
51490     End If
51500    Else
51510     If LenB(tstr) > 0 Then
51520      .StampString = tstr
51530     End If
51540   End If
51550   tstr = reg.GetRegistryValue("StampUseOutlineFont")
51560   If IsNumeric(tstr) Then
51570     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51580       .StampUseOutlineFont = CLng(tstr)
51590      Else
51600       If UseStandard Then
51610        .StampUseOutlineFont = 1
51620       End If
51630     End If
51640    Else
51650     If UseStandard Then
51660      .StampUseOutlineFont = 1
51670     End If
51680   End If
51690   tstr = reg.GetRegistryValue("StandardAuthor")
51700   If LenB(tstr) = 0 And LenB("") > 0 Then
51710     If UseStandard Then
51720      .StandardAuthor = " "
51730     End If
51740    Else
51750     If LenB(tstr) > 0 Then
51760      .StandardAuthor = tstr
51770     End If
51780   End If
51790   tstr = reg.GetRegistryValue("StandardCreationdate")
51800   If LenB(tstr) = 0 And LenB("") > 0 Then
51810     If UseStandard Then
51820      .StandardCreationdate = " "
51830     End If
51840    Else
51850     If LenB(tstr) > 0 Then
51860      .StandardCreationdate = tstr
51870     End If
51880   End If
51890   tstr = reg.GetRegistryValue("StandardDateformat")
51900   If LenB(tstr) = 0 And LenB("YYYYMMDDHHNNSS") > 0 Then
51910     If UseStandard Then
51920      .StandardDateformat = "YYYYMMDDHHNNSS"
51930     End If
51940    Else
51950     If LenB(tstr) > 0 Then
51960      .StandardDateformat = tstr
51970     End If
51980   End If
51990   tstr = reg.GetRegistryValue("StandardKeywords")
52000   If LenB(tstr) = 0 And LenB("") > 0 Then
52010     If UseStandard Then
52020      .StandardKeywords = " "
52030     End If
52040    Else
52050     If LenB(tstr) > 0 Then
52060      .StandardKeywords = tstr
52070     End If
52080   End If
52090   tstr = reg.GetRegistryValue("StandardMailDomain")
52100   If LenB(tstr) = 0 And LenB("") > 0 Then
52110     If UseStandard Then
52120      .StandardMailDomain = " "
52130     End If
52140    Else
52150     If LenB(tstr) > 0 Then
52160      .StandardMailDomain = tstr
52170     End If
52180   End If
52190   tstr = reg.GetRegistryValue("StandardModifydate")
52200   If LenB(tstr) = 0 And LenB("") > 0 Then
52210     If UseStandard Then
52220      .StandardModifydate = " "
52230     End If
52240    Else
52250     If LenB(tstr) > 0 Then
52260      .StandardModifydate = tstr
52270     End If
52280   End If
52290   tstr = reg.GetRegistryValue("StandardSaveformat")
52300   If LenB(tstr) = 0 And LenB("pdf") > 0 Then
52310     If UseStandard Then
52320      .StandardSaveformat = "pdf"
52330     End If
52340    Else
52350     If LenB(tstr) > 0 Then
52360      .StandardSaveformat = tstr
52370     End If
52380   End If
52390   tstr = reg.GetRegistryValue("StandardSubject")
52400   If LenB(tstr) = 0 And LenB("") > 0 Then
52410     If UseStandard Then
52420      .StandardSubject = " "
52430     End If
52440    Else
52450     If LenB(tstr) > 0 Then
52460      .StandardSubject = tstr
52470     End If
52480   End If
52490   tstr = reg.GetRegistryValue("StandardTitle")
52500   If LenB(tstr) = 0 And LenB("") > 0 Then
52510     If UseStandard Then
52520      .StandardTitle = " "
52530     End If
52540    Else
52550     If LenB(tstr) > 0 Then
52560      .StandardTitle = tstr
52570     End If
52580   End If
52590   tstr = reg.GetRegistryValue("UseCreationDateNow")
52600   If IsNumeric(tstr) Then
52610     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52620       .UseCreationDateNow = CLng(tstr)
52630      Else
52640       If UseStandard Then
52650        .UseCreationDateNow = 0
52660       End If
52670     End If
52680    Else
52690     If UseStandard Then
52700      .UseCreationDateNow = 0
52710     End If
52720   End If
52730   tstr = reg.GetRegistryValue("UseStandardAuthor")
52740   If IsNumeric(tstr) Then
52750     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52760       .UseStandardAuthor = CLng(tstr)
52770      Else
52780       If UseStandard Then
52790        .UseStandardAuthor = 0
52800       End If
52810     End If
52820    Else
52830     If UseStandard Then
52840      .UseStandardAuthor = 0
52850     End If
52860   End If
52870   reg.Subkey = "Printing\Formats\Bitmap\Colors"
52880   tstr = reg.GetRegistryValue("BitmapResolution")
52890   If IsNumeric(tstr) Then
52900     If CLng(tstr) >= 1 Then
52910       .BitmapResolution = CLng(tstr)
52920      Else
52930       If UseStandard Then
52940        .BitmapResolution = 150
52950       End If
52960     End If
52970    Else
52980     If UseStandard Then
52990      .BitmapResolution = 150
53000     End If
53010   End If
53020   tstr = reg.GetRegistryValue("BMPColorscount")
53030   If IsNumeric(tstr) Then
53040     If CLng(tstr) >= 0 And CLng(tstr) <= 6 Then
53050       .BMPColorscount = CLng(tstr)
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
53160   tstr = reg.GetRegistryValue("JPEGColorscount")
53170   If IsNumeric(tstr) Then
53180     If CLng(tstr) >= 0 And CLng(tstr) <= 1 Then
53190       .JPEGColorscount = CLng(tstr)
53200      Else
53210       If UseStandard Then
53220        .JPEGColorscount = 0
53230       End If
53240     End If
53250    Else
53260     If UseStandard Then
53270      .JPEGColorscount = 0
53280     End If
53290   End If
53300   tstr = reg.GetRegistryValue("JPEGQuality")
53310   If IsNumeric(tstr) Then
53320     If CLng(tstr) >= 0 And CLng(tstr) <= 100 Then
53330       .JPEGQuality = CLng(tstr)
53340      Else
53350       If UseStandard Then
53360        .JPEGQuality = 75
53370       End If
53380     End If
53390    Else
53400     If UseStandard Then
53410      .JPEGQuality = 75
53420     End If
53430   End If
53440   tstr = reg.GetRegistryValue("PCXColorscount")
53450   If IsNumeric(tstr) Then
53460     If CLng(tstr) >= 0 And CLng(tstr) <= 5 Then
53470       .PCXColorscount = CLng(tstr)
53480      Else
53490       If UseStandard Then
53500        .PCXColorscount = 0
53510       End If
53520     End If
53530    Else
53540     If UseStandard Then
53550      .PCXColorscount = 0
53560     End If
53570   End If
53580   tstr = reg.GetRegistryValue("PNGColorscount")
53590   If IsNumeric(tstr) Then
53600     If CLng(tstr) >= 0 And CLng(tstr) <= 4 Then
53610       .PNGColorscount = CLng(tstr)
53620      Else
53630       If UseStandard Then
53640        .PNGColorscount = 0
53650       End If
53660     End If
53670    Else
53680     If UseStandard Then
53690      .PNGColorscount = 0
53700     End If
53710   End If
53720   tstr = reg.GetRegistryValue("TIFFColorscount")
53730   If IsNumeric(tstr) Then
53740     If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
53750       .TIFFColorscount = CLng(tstr)
53760      Else
53770       If UseStandard Then
53780        .TIFFColorscount = 0
53790       End If
53800     End If
53810    Else
53820     If UseStandard Then
53830      .TIFFColorscount = 0
53840     End If
53850   End If
53860   reg.Subkey = "Printing\Formats\PDF\Colors"
53870   tstr = reg.GetRegistryValue("PDFColorsCMYKToRGB")
53880   If IsNumeric(tstr) Then
53890     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53900       .PDFColorsCMYKToRGB = CLng(tstr)
53910      Else
53920       If UseStandard Then
53930        .PDFColorsCMYKToRGB = 1
53940       End If
53950     End If
53960    Else
53970     If UseStandard Then
53980      .PDFColorsCMYKToRGB = 1
53990     End If
54000   End If
54010   tstr = reg.GetRegistryValue("PDFColorsColorModel")
54020   If IsNumeric(tstr) Then
54030     If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
54040       .PDFColorsColorModel = CLng(tstr)
54050      Else
54060       If UseStandard Then
54070        .PDFColorsColorModel = 1
54080       End If
54090     End If
54100    Else
54110     If UseStandard Then
54120      .PDFColorsColorModel = 1
54130     End If
54140   End If
54150   tstr = reg.GetRegistryValue("PDFColorsPreserveHalftone")
54160   If IsNumeric(tstr) Then
54170     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54180       .PDFColorsPreserveHalftone = CLng(tstr)
54190      Else
54200       If UseStandard Then
54210        .PDFColorsPreserveHalftone = 0
54220       End If
54230     End If
54240    Else
54250     If UseStandard Then
54260      .PDFColorsPreserveHalftone = 0
54270     End If
54280   End If
54290   tstr = reg.GetRegistryValue("PDFColorsPreserveOverprint")
54300   If IsNumeric(tstr) Then
54310     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54320       .PDFColorsPreserveOverprint = CLng(tstr)
54330      Else
54340       If UseStandard Then
54350        .PDFColorsPreserveOverprint = 1
54360       End If
54370     End If
54380    Else
54390     If UseStandard Then
54400      .PDFColorsPreserveOverprint = 1
54410     End If
54420   End If
54430   tstr = reg.GetRegistryValue("PDFColorsPreserveTransfer")
54440   If IsNumeric(tstr) Then
54450     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54460       .PDFColorsPreserveTransfer = CLng(tstr)
54470      Else
54480       If UseStandard Then
54490        .PDFColorsPreserveTransfer = 1
54500       End If
54510     End If
54520    Else
54530     If UseStandard Then
54540      .PDFColorsPreserveTransfer = 1
54550     End If
54560   End If
54570   reg.Subkey = "Printing\Formats\PDF\Compression"
54580   tstr = reg.GetRegistryValue("PDFCompressionColorCompression")
54590   If IsNumeric(tstr) Then
54600     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54610       .PDFCompressionColorCompression = CLng(tstr)
54620      Else
54630       If UseStandard Then
54640        .PDFCompressionColorCompression = 1
54650       End If
54660     End If
54670    Else
54680     If UseStandard Then
54690      .PDFCompressionColorCompression = 1
54700     End If
54710   End If
54720   tstr = reg.GetRegistryValue("PDFCompressionColorCompressionChoice")
54730   If IsNumeric(tstr) Then
54740     If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
54750       .PDFCompressionColorCompressionChoice = CLng(tstr)
54760      Else
54770       If UseStandard Then
54780        .PDFCompressionColorCompressionChoice = 0
54790       End If
54800     End If
54810    Else
54820     If UseStandard Then
54830      .PDFCompressionColorCompressionChoice = 0
54840     End If
54850   End If
54860   tstr = reg.GetRegistryValue("PDFCompressionColorResample")
54870   If IsNumeric(tstr) Then
54880     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54890       .PDFCompressionColorResample = CLng(tstr)
54900      Else
54910       If UseStandard Then
54920        .PDFCompressionColorResample = 0
54930       End If
54940     End If
54950    Else
54960     If UseStandard Then
54970      .PDFCompressionColorResample = 0
54980     End If
54990   End If
55000   tstr = reg.GetRegistryValue("PDFCompressionColorResampleChoice")
55010   If IsNumeric(tstr) Then
55020     If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
55030       .PDFCompressionColorResampleChoice = CLng(tstr)
55040      Else
55050       If UseStandard Then
55060        .PDFCompressionColorResampleChoice = 0
55070       End If
55080     End If
55090    Else
55100     If UseStandard Then
55110      .PDFCompressionColorResampleChoice = 0
55120     End If
55130   End If
55140   tstr = reg.GetRegistryValue("PDFCompressionColorResolution")
55150   If IsNumeric(tstr) Then
55160     If CLng(tstr) >= 0 Then
55170       .PDFCompressionColorResolution = CLng(tstr)
55180      Else
55190       If UseStandard Then
55200        .PDFCompressionColorResolution = 300
55210       End If
55220     End If
55230    Else
55240     If UseStandard Then
55250      .PDFCompressionColorResolution = 300
55260     End If
55270   End If
55280   tstr = reg.GetRegistryValue("PDFCompressionGreyCompression")
55290   If IsNumeric(tstr) Then
55300     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55310       .PDFCompressionGreyCompression = CLng(tstr)
55320      Else
55330       If UseStandard Then
55340        .PDFCompressionGreyCompression = 1
55350       End If
55360     End If
55370    Else
55380     If UseStandard Then
55390      .PDFCompressionGreyCompression = 1
55400     End If
55410   End If
55420   tstr = reg.GetRegistryValue("PDFCompressionGreyCompressionChoice")
55430   If IsNumeric(tstr) Then
55440     If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
55450       .PDFCompressionGreyCompressionChoice = CLng(tstr)
55460      Else
55470       If UseStandard Then
55480        .PDFCompressionGreyCompressionChoice = 0
55490       End If
55500     End If
55510    Else
55520     If UseStandard Then
55530      .PDFCompressionGreyCompressionChoice = 0
55540     End If
55550   End If
55560   tstr = reg.GetRegistryValue("PDFCompressionGreyResample")
55570   If IsNumeric(tstr) Then
55580     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55590       .PDFCompressionGreyResample = CLng(tstr)
55600      Else
55610       If UseStandard Then
55620        .PDFCompressionGreyResample = 0
55630       End If
55640     End If
55650    Else
55660     If UseStandard Then
55670      .PDFCompressionGreyResample = 0
55680     End If
55690   End If
55700   tstr = reg.GetRegistryValue("PDFCompressionGreyResampleChoice")
55710   If IsNumeric(tstr) Then
55720     If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
55730       .PDFCompressionGreyResampleChoice = CLng(tstr)
55740      Else
55750       If UseStandard Then
55760        .PDFCompressionGreyResampleChoice = 0
55770       End If
55780     End If
55790    Else
55800     If UseStandard Then
55810      .PDFCompressionGreyResampleChoice = 0
55820     End If
55830   End If
55840   tstr = reg.GetRegistryValue("PDFCompressionGreyResolution")
55850   If IsNumeric(tstr) Then
55860     If CLng(tstr) >= 0 Then
55870       .PDFCompressionGreyResolution = CLng(tstr)
55880      Else
55890       If UseStandard Then
55900        .PDFCompressionGreyResolution = 300
55910       End If
55920     End If
55930    Else
55940     If UseStandard Then
55950      .PDFCompressionGreyResolution = 300
55960     End If
55970   End If
55980   tstr = reg.GetRegistryValue("PDFCompressionMonoCompression")
55990   If IsNumeric(tstr) Then
56000     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
56010       .PDFCompressionMonoCompression = CLng(tstr)
56020      Else
56030       If UseStandard Then
56040        .PDFCompressionMonoCompression = 1
56050       End If
56060     End If
56070    Else
56080     If UseStandard Then
56090      .PDFCompressionMonoCompression = 1
56100     End If
56110   End If
56120   tstr = reg.GetRegistryValue("PDFCompressionMonoCompressionChoice")
56130   If IsNumeric(tstr) Then
56140     If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
56150       .PDFCompressionMonoCompressionChoice = CLng(tstr)
56160      Else
56170       If UseStandard Then
56180        .PDFCompressionMonoCompressionChoice = 0
56190       End If
56200     End If
56210    Else
56220     If UseStandard Then
56230      .PDFCompressionMonoCompressionChoice = 0
56240     End If
56250   End If
56260   tstr = reg.GetRegistryValue("PDFCompressionMonoResample")
56270   If IsNumeric(tstr) Then
56280     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
56290       .PDFCompressionMonoResample = CLng(tstr)
56300      Else
56310       If UseStandard Then
56320        .PDFCompressionMonoResample = 0
56330       End If
56340     End If
56350    Else
56360     If UseStandard Then
56370      .PDFCompressionMonoResample = 0
56380     End If
56390   End If
56400   tstr = reg.GetRegistryValue("PDFCompressionMonoResampleChoice")
56410   If IsNumeric(tstr) Then
56420     If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
56430       .PDFCompressionMonoResampleChoice = CLng(tstr)
56440      Else
56450       If UseStandard Then
56460        .PDFCompressionMonoResampleChoice = 0
56470       End If
56480     End If
56490    Else
56500     If UseStandard Then
56510      .PDFCompressionMonoResampleChoice = 0
56520     End If
56530   End If
56540   tstr = reg.GetRegistryValue("PDFCompressionMonoResolution")
56550   If IsNumeric(tstr) Then
56560     If CLng(tstr) >= 0 Then
56570       .PDFCompressionMonoResolution = CLng(tstr)
56580      Else
56590       If UseStandard Then
56600        .PDFCompressionMonoResolution = 1200
56610       End If
56620     End If
56630    Else
56640     If UseStandard Then
56650      .PDFCompressionMonoResolution = 1200
56660     End If
56670   End If
56680   tstr = reg.GetRegistryValue("PDFCompressionTextCompression")
56690   If IsNumeric(tstr) Then
56700     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
56710       .PDFCompressionTextCompression = CLng(tstr)
56720      Else
56730       If UseStandard Then
56740        .PDFCompressionTextCompression = 1
56750       End If
56760     End If
56770    Else
56780     If UseStandard Then
56790      .PDFCompressionTextCompression = 1
56800     End If
56810   End If
56820   reg.Subkey = "Printing\Formats\PDF\Fonts"
56830   tstr = reg.GetRegistryValue("PDFFontsEmbedAll")
56840   If IsNumeric(tstr) Then
56850     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
56860       .PDFFontsEmbedAll = CLng(tstr)
56870      Else
56880       If UseStandard Then
56890        .PDFFontsEmbedAll = 1
56900       End If
56910     End If
56920    Else
56930     If UseStandard Then
56940      .PDFFontsEmbedAll = 1
56950     End If
56960   End If
56970   tstr = reg.GetRegistryValue("PDFFontsSubSetFonts")
56980   If IsNumeric(tstr) Then
56990     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
57000       .PDFFontsSubSetFonts = CLng(tstr)
57010      Else
57020       If UseStandard Then
57030        .PDFFontsSubSetFonts = 1
57040       End If
57050     End If
57060    Else
57070     If UseStandard Then
57080      .PDFFontsSubSetFonts = 1
57090     End If
57100   End If
57110   tstr = reg.GetRegistryValue("PDFFontsSubSetFontsPercent")
57120   If IsNumeric(tstr) Then
57130     If CLng(tstr) >= 0 Then
57140       .PDFFontsSubSetFontsPercent = CLng(tstr)
57150      Else
57160       If UseStandard Then
57170        .PDFFontsSubSetFontsPercent = 100
57180       End If
57190     End If
57200    Else
57210     If UseStandard Then
57220      .PDFFontsSubSetFontsPercent = 100
57230     End If
57240   End If
57250   reg.Subkey = "Printing\Formats\PDF\General"
57260   tstr = reg.GetRegistryValue("PDFGeneralASCII85")
57270   If IsNumeric(tstr) Then
57280     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
57290       .PDFGeneralASCII85 = CLng(tstr)
57300      Else
57310       If UseStandard Then
57320        .PDFGeneralASCII85 = 0
57330       End If
57340     End If
57350    Else
57360     If UseStandard Then
57370      .PDFGeneralASCII85 = 0
57380     End If
57390   End If
57400   tstr = reg.GetRegistryValue("PDFGeneralAutorotate")
57410   If IsNumeric(tstr) Then
57420     If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
57430       .PDFGeneralAutorotate = CLng(tstr)
57440      Else
57450       If UseStandard Then
57460        .PDFGeneralAutorotate = 2
57470       End If
57480     End If
57490    Else
57500     If UseStandard Then
57510      .PDFGeneralAutorotate = 2
57520     End If
57530   End If
57540   tstr = reg.GetRegistryValue("PDFGeneralCompatibility")
57550   If IsNumeric(tstr) Then
57560     If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
57570       .PDFGeneralCompatibility = CLng(tstr)
57580      Else
57590       If UseStandard Then
57600        .PDFGeneralCompatibility = 1
57610       End If
57620     End If
57630    Else
57640     If UseStandard Then
57650      .PDFGeneralCompatibility = 1
57660     End If
57670   End If
57680   tstr = reg.GetRegistryValue("PDFGeneralOverprint")
57690   If IsNumeric(tstr) Then
57700     If CLng(tstr) >= 0 And CLng(tstr) <= 1 Then
57710       .PDFGeneralOverprint = CLng(tstr)
57720      Else
57730       If UseStandard Then
57740        .PDFGeneralOverprint = 0
57750       End If
57760     End If
57770    Else
57780     If UseStandard Then
57790      .PDFGeneralOverprint = 0
57800     End If
57810   End If
57820   tstr = reg.GetRegistryValue("PDFGeneralResolution")
57830   If IsNumeric(tstr) Then
57840     If CLng(tstr) >= 0 Then
57850       .PDFGeneralResolution = CLng(tstr)
57860      Else
57870       If UseStandard Then
57880        .PDFGeneralResolution = 600
57890       End If
57900     End If
57910    Else
57920     If UseStandard Then
57930      .PDFGeneralResolution = 600
57940     End If
57950   End If
57960   tstr = reg.GetRegistryValue("PDFOptimize")
57970   If IsNumeric(tstr) Then
57980     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
57990       .PDFOptimize = CLng(tstr)
58000      Else
58010       If UseStandard Then
58020        .PDFOptimize = 0
58030       End If
58040     End If
58050    Else
58060     If UseStandard Then
58070      .PDFOptimize = 0
58080     End If
58090   End If
58100   reg.Subkey = "Printing\Formats\PDF\Security"
58110   tstr = reg.GetRegistryValue("PDFAllowAssembly")
58120   If IsNumeric(tstr) Then
58130     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
58140       .PDFAllowAssembly = CLng(tstr)
58150      Else
58160       If UseStandard Then
58170        .PDFAllowAssembly = 0
58180       End If
58190     End If
58200    Else
58210     If UseStandard Then
58220      .PDFAllowAssembly = 0
58230     End If
58240   End If
58250   tstr = reg.GetRegistryValue("PDFAllowDegradedPrinting")
58260   If IsNumeric(tstr) Then
58270     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
58280       .PDFAllowDegradedPrinting = CLng(tstr)
58290      Else
58300       If UseStandard Then
58310        .PDFAllowDegradedPrinting = 0
58320       End If
58330     End If
58340    Else
58350     If UseStandard Then
58360      .PDFAllowDegradedPrinting = 0
58370     End If
58380   End If
58390   tstr = reg.GetRegistryValue("PDFAllowFillIn")
58400   If IsNumeric(tstr) Then
58410     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
58420       .PDFAllowFillIn = CLng(tstr)
58430      Else
58440       If UseStandard Then
58450        .PDFAllowFillIn = 0
58460       End If
58470     End If
58480    Else
58490     If UseStandard Then
58500      .PDFAllowFillIn = 0
58510     End If
58520   End If
58530   tstr = reg.GetRegistryValue("PDFAllowScreenReaders")
58540   If IsNumeric(tstr) Then
58550     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
58560       .PDFAllowScreenReaders = CLng(tstr)
58570      Else
58580       If UseStandard Then
58590        .PDFAllowScreenReaders = 0
58600       End If
58610     End If
58620    Else
58630     If UseStandard Then
58640      .PDFAllowScreenReaders = 0
58650     End If
58660   End If
58670   tstr = reg.GetRegistryValue("PDFDisallowCopy")
58680   If IsNumeric(tstr) Then
58690     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
58700       .PDFDisallowCopy = CLng(tstr)
58710      Else
58720       If UseStandard Then
58730        .PDFDisallowCopy = 1
58740       End If
58750     End If
58760    Else
58770     If UseStandard Then
58780      .PDFDisallowCopy = 1
58790     End If
58800   End If
58810   tstr = reg.GetRegistryValue("PDFDisallowModifyAnnotations")
58820   If IsNumeric(tstr) Then
58830     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
58840       .PDFDisallowModifyAnnotations = CLng(tstr)
58850      Else
58860       If UseStandard Then
58870        .PDFDisallowModifyAnnotations = 0
58880       End If
58890     End If
58900    Else
58910     If UseStandard Then
58920      .PDFDisallowModifyAnnotations = 0
58930     End If
58940   End If
58950   tstr = reg.GetRegistryValue("PDFDisallowModifyContents")
58960   If IsNumeric(tstr) Then
58970     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
58980       .PDFDisallowModifyContents = CLng(tstr)
58990      Else
59000       If UseStandard Then
59010        .PDFDisallowModifyContents = 0
59020       End If
59030     End If
59040    Else
59050     If UseStandard Then
59060      .PDFDisallowModifyContents = 0
59070     End If
59080   End If
59090   tstr = reg.GetRegistryValue("PDFDisallowPrinting")
59100   If IsNumeric(tstr) Then
59110     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
59120       .PDFDisallowPrinting = CLng(tstr)
59130      Else
59140       If UseStandard Then
59150        .PDFDisallowPrinting = 0
59160       End If
59170     End If
59180    Else
59190     If UseStandard Then
59200      .PDFDisallowPrinting = 0
59210     End If
59220   End If
59230   tstr = reg.GetRegistryValue("PDFEncryptor")
59240   If IsNumeric(tstr) Then
59250     If CLng(tstr) >= 0 And CLng(tstr) <= 1 Then
59260       .PDFEncryptor = CLng(tstr)
59270      Else
59280       If UseStandard Then
59290        .PDFEncryptor = 0
59300       End If
59310     End If
59320    Else
59330     If UseStandard Then
59340      .PDFEncryptor = 0
59350     End If
59360   End If
59370   tstr = reg.GetRegistryValue("PDFHighEncryption")
59380   If IsNumeric(tstr) Then
59390     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
59400       .PDFHighEncryption = CLng(tstr)
59410      Else
59420       If UseStandard Then
59430        .PDFHighEncryption = 0
59440       End If
59450     End If
59460    Else
59470     If UseStandard Then
59480      .PDFHighEncryption = 0
59490     End If
59500   End If
59510   tstr = reg.GetRegistryValue("PDFLowEncryption")
59520   If IsNumeric(tstr) Then
59530     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
59540       .PDFLowEncryption = CLng(tstr)
59550      Else
59560       If UseStandard Then
59570        .PDFLowEncryption = 1
59580       End If
59590     End If
59600    Else
59610     If UseStandard Then
59620      .PDFLowEncryption = 1
59630     End If
59640   End If
59650   tstr = reg.GetRegistryValue("PDFOwnerPass")
59660   If IsNumeric(tstr) Then
59670     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
59680       .PDFOwnerPass = CLng(tstr)
59690      Else
59700       If UseStandard Then
59710        .PDFOwnerPass = 0
59720       End If
59730     End If
59740    Else
59750     If UseStandard Then
59760      .PDFOwnerPass = 0
59770     End If
59780   End If
59790   tstr = reg.GetRegistryValue("PDFOwnerPasswordString")
59800   If LenB(tstr) = 0 And LenB("") > 0 Then
59810     If UseStandard Then
59820      .PDFOwnerPasswordString = " "
59830     End If
59840    Else
59850     If LenB(tstr) > 0 Then
59860      .PDFOwnerPasswordString = tstr
59870     End If
59880   End If
59890   tstr = reg.GetRegistryValue("PDFUserPass")
59900   If IsNumeric(tstr) Then
59910     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
59920       .PDFUserPass = CLng(tstr)
59930      Else
59940       If UseStandard Then
59950        .PDFUserPass = 0
59960       End If
59970     End If
59980    Else
59990     If UseStandard Then
60000      .PDFUserPass = 0
60010     End If
60020   End If
60030   tstr = reg.GetRegistryValue("PDFUserPasswordString")
60040   If LenB(tstr) = 0 And LenB("") > 0 Then
60050     If UseStandard Then
60060      .PDFUserPasswordString = " "
60070     End If
60080    Else
60090     If LenB(tstr) > 0 Then
60100      .PDFUserPasswordString = tstr
60110     End If
60120   End If
60130   tstr = reg.GetRegistryValue("PDFUseSecurity")
60140   If IsNumeric(tstr) Then
60150     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
60160       .PDFUseSecurity = CLng(tstr)
60170      Else
60180       If UseStandard Then
60190        .PDFUseSecurity = 0
60200       End If
60210     End If
60220    Else
60230     If UseStandard Then
60240      .PDFUseSecurity = 0
60250     End If
60260   End If
60270   reg.Subkey = "Printing\Formats\PS\LanguageLevel"
60280   tstr = reg.GetRegistryValue("EPSLanguageLevel")
60290   If IsNumeric(tstr) Then
60300     If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
60310       .EPSLanguageLevel = CLng(tstr)
60320      Else
60330       If UseStandard Then
60340        .EPSLanguageLevel = 2
60350       End If
60360     End If
60370    Else
60380     If UseStandard Then
60390      .EPSLanguageLevel = 2
60400     End If
60410   End If
60420   tstr = reg.GetRegistryValue("PSLanguageLevel")
60430   If IsNumeric(tstr) Then
60440     If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
60450       .PSLanguageLevel = CLng(tstr)
60460      Else
60470       If UseStandard Then
60480        .PSLanguageLevel = 2
60490       End If
60500     End If
60510    Else
60520     If UseStandard Then
60530      .PSLanguageLevel = 2
60540     End If
60550   End If
60560   reg.Subkey = "Program"
60570   tstr = reg.GetRegistryValue("AdditionalGhostscriptParameters")
60580   If LenB(tstr) = 0 And LenB("") > 0 Then
60590     If UseStandard Then
60600      .AdditionalGhostscriptParameters = " "
60610     End If
60620    Else
60630     If LenB(tstr) > 0 Then
60640      .AdditionalGhostscriptParameters = tstr
60650     End If
60660   End If
60670   tstr = reg.GetRegistryValue("AdditionalGhostscriptSearchpath")
60680   If LenB(tstr) = 0 And LenB("") > 0 Then
60690     If UseStandard Then
60700      .AdditionalGhostscriptSearchpath = " "
60710     End If
60720    Else
60730     If LenB(tstr) > 0 Then
60740      .AdditionalGhostscriptSearchpath = tstr
60750     End If
60760   End If
60770   tstr = reg.GetRegistryValue("AddWindowsFontpath")
60780   If IsNumeric(tstr) Then
60790     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
60800       .AddWindowsFontpath = CLng(tstr)
60810      Else
60820       If UseStandard Then
60830        .AddWindowsFontpath = 1
60840       End If
60850     End If
60860    Else
60870     If UseStandard Then
60880      .AddWindowsFontpath = 1
60890     End If
60900   End If
60910   tstr = reg.GetRegistryValue("AutosaveDirectory")
60920   If LenB(Trim$(tstr)) > 0 Then
60930     .AutosaveDirectory = CompletePath(tstr)
60940    Else
60950     If UseStandard Then
60960      tstr = GetMyFiles
60970      .AutosaveDirectory = CompletePath(tstr)
60980     End If
60990   End If
61000   tstr = reg.GetRegistryValue("AutosaveFilename")
61010   If LenB(tstr) = 0 And LenB("<DateTime>") > 0 Then
61020     If UseStandard Then
61030      .AutosaveFilename = "<DateTime>"
61040     End If
61050    Else
61060     If LenB(tstr) > 0 Then
61070      .AutosaveFilename = tstr
61080     End If
61090   End If
61100   tstr = reg.GetRegistryValue("AutosaveFormat")
61110   If IsNumeric(tstr) Then
61120     If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
61130       .AutosaveFormat = CLng(tstr)
61140      Else
61150       If UseStandard Then
61160        .AutosaveFormat = 0
61170       End If
61180     End If
61190    Else
61200     If UseStandard Then
61210      .AutosaveFormat = 0
61220     End If
61230   End If
61240   tstr = reg.GetRegistryValue("ClientComputerResolveIPAddress")
61250   If IsNumeric(tstr) Then
61260     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
61270       .ClientComputerResolveIPAddress = CLng(tstr)
61280      Else
61290       If UseStandard Then
61300        .ClientComputerResolveIPAddress = 0
61310       End If
61320     End If
61330    Else
61340     If UseStandard Then
61350      .ClientComputerResolveIPAddress = 0
61360     End If
61370   End If
61380   tstr = reg.GetRegistryValue("DisableEmail")
61390   If IsNumeric(tstr) Then
61400     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
61410       .DisableEmail = CLng(tstr)
61420      Else
61430       If UseStandard Then
61440        .DisableEmail = 0
61450       End If
61460     End If
61470    Else
61480     If UseStandard Then
61490      .DisableEmail = 0
61500     End If
61510   End If
61520   tstr = reg.GetRegistryValue("DontUseDocumentSettings")
61530   If IsNumeric(tstr) Then
61540     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
61550       .DontUseDocumentSettings = CLng(tstr)
61560      Else
61570       If UseStandard Then
61580        .DontUseDocumentSettings = 0
61590       End If
61600     End If
61610    Else
61620     If UseStandard Then
61630      .DontUseDocumentSettings = 0
61640     End If
61650   End If
61660   tstr = reg.GetRegistryValue("FilenameSubstitutions")
61670   If LenB(tstr) = 0 And LenB("Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt") > 0 Then
61680     If UseStandard Then
61690      .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
61700     End If
61710    Else
61720     If LenB(tstr) > 0 Then
61730      .FilenameSubstitutions = tstr
61740     End If
61750   End If
61760   tstr = reg.GetRegistryValue("FilenameSubstitutionsOnlyInTitle")
61770   If IsNumeric(tstr) Then
61780     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
61790       .FilenameSubstitutionsOnlyInTitle = CLng(tstr)
61800      Else
61810       If UseStandard Then
61820        .FilenameSubstitutionsOnlyInTitle = 1
61830       End If
61840     End If
61850    Else
61860     If UseStandard Then
61870      .FilenameSubstitutionsOnlyInTitle = 1
61880     End If
61890   End If
61900   tstr = reg.GetRegistryValue("Language")
61910   If LenB(tstr) = 0 And LenB("english") > 0 Then
61920     If UseStandard Then
61930      .Language = "english"
61940     End If
61950    Else
61960     If LenB(tstr) > 0 Then
61970      .Language = tstr
61980     End If
61990   End If
62000   tstr = reg.GetRegistryValue("LastSaveDirectory")
62010   If LenB(Trim$(tstr)) > 0 Then
62020     .LastSaveDirectory = CompletePath(tstr)
62030    Else
62040     If UseStandard Then
62050      tstr = GetMyFiles
62060      .LastSaveDirectory = CompletePath(tstr)
62070     End If
62080   End If
62090   tstr = reg.GetRegistryValue("Logging")
62100   If IsNumeric(tstr) Then
62110     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
62120       .Logging = CLng(tstr)
62130      Else
62140       If UseStandard Then
62150        .Logging = 0
62160       End If
62170     End If
62180    Else
62190     If UseStandard Then
62200      .Logging = 0
62210     End If
62220   End If
62230   tstr = reg.GetRegistryValue("LogLines")
62240   If IsNumeric(tstr) Then
62250     If CLng(tstr) >= 100 And CLng(tstr) <= 1000 Then
62260       .LogLines = CLng(tstr)
62270      Else
62280       If UseStandard Then
62290        .LogLines = 100
62300       End If
62310     End If
62320    Else
62330     If UseStandard Then
62340      .LogLines = 100
62350     End If
62360   End If
62370   tstr = reg.GetRegistryValue("NoConfirmMessageSwitchingDefaultprinter")
62380   If IsNumeric(tstr) Then
62390     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
62400       .NoConfirmMessageSwitchingDefaultprinter = CLng(tstr)
62410      Else
62420       If UseStandard Then
62430        .NoConfirmMessageSwitchingDefaultprinter = 0
62440       End If
62450     End If
62460    Else
62470     If UseStandard Then
62480      .NoConfirmMessageSwitchingDefaultprinter = 0
62490     End If
62500   End If
62510   tstr = reg.GetRegistryValue("NoProcessingAtStartup")
62520   If IsNumeric(tstr) Then
62530     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
62540       .NoProcessingAtStartup = CLng(tstr)
62550      Else
62560       If UseStandard Then
62570        .NoProcessingAtStartup = 0
62580       End If
62590     End If
62600    Else
62610     If UseStandard Then
62620      .NoProcessingAtStartup = 0
62630     End If
62640   End If
62650   tstr = reg.GetRegistryValue("OptionsDesign")
62660   If IsNumeric(tstr) Then
62670     If CLng(tstr) >= 1 And CLng(tstr) <= 2 Then
62680       .OptionsDesign = CLng(tstr)
62690      Else
62700       If UseStandard Then
62710        .OptionsDesign = 1
62720       End If
62730     End If
62740    Else
62750     If UseStandard Then
62760      .OptionsDesign = 1
62770     End If
62780   End If
62790   tstr = reg.GetRegistryValue("OptionsEnabled")
62800   If IsNumeric(tstr) Then
62810     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
62820       .OptionsEnabled = CLng(tstr)
62830      Else
62840       If UseStandard Then
62850        .OptionsEnabled = 1
62860       End If
62870     End If
62880    Else
62890     If UseStandard Then
62900      .OptionsEnabled = 1
62910     End If
62920   End If
62930   tstr = reg.GetRegistryValue("OptionsVisible")
62940   If IsNumeric(tstr) Then
62950     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
62960       .OptionsVisible = CLng(tstr)
62970      Else
62980       If UseStandard Then
62990        .OptionsVisible = 1
63000       End If
63010     End If
63020    Else
63030     If UseStandard Then
63040      .OptionsVisible = 1
63050     End If
63060   End If
63070   tstr = reg.GetRegistryValue("PrintAfterSaving")
63080   If IsNumeric(tstr) Then
63090     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
63100       .PrintAfterSaving = CLng(tstr)
63110      Else
63120       If UseStandard Then
63130        .PrintAfterSaving = 0
63140       End If
63150     End If
63160    Else
63170     If UseStandard Then
63180      .PrintAfterSaving = 0
63190     End If
63200   End If
63210   tstr = reg.GetRegistryValue("PrintAfterSavingDuplex")
63220   If IsNumeric(tstr) Then
63230     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
63240       .PrintAfterSavingDuplex = CLng(tstr)
63250      Else
63260       If UseStandard Then
63270        .PrintAfterSavingDuplex = 0
63280       End If
63290     End If
63300    Else
63310     If UseStandard Then
63320      .PrintAfterSavingDuplex = 0
63330     End If
63340   End If
63350   tstr = reg.GetRegistryValue("PrintAfterSavingNoCancel")
63360   If IsNumeric(tstr) Then
63370     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
63380       .PrintAfterSavingNoCancel = CLng(tstr)
63390      Else
63400       If UseStandard Then
63410        .PrintAfterSavingNoCancel = 0
63420       End If
63430     End If
63440    Else
63450     If UseStandard Then
63460      .PrintAfterSavingNoCancel = 0
63470     End If
63480   End If
63490   tstr = reg.GetRegistryValue("PrintAfterSavingPrinter")
63500   If LenB(tstr) = 0 And LenB("") > 0 Then
63510     If UseStandard Then
63520      .PrintAfterSavingPrinter = " "
63530     End If
63540    Else
63550     If LenB(tstr) > 0 Then
63560      .PrintAfterSavingPrinter = tstr
63570     End If
63580   End If
63590   tstr = reg.GetRegistryValue("PrintAfterSavingQueryUser")
63600   If IsNumeric(tstr) Then
63610     If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
63620       .PrintAfterSavingQueryUser = CLng(tstr)
63630      Else
63640       If UseStandard Then
63650        .PrintAfterSavingQueryUser = 0
63660       End If
63670     End If
63680    Else
63690     If UseStandard Then
63700      .PrintAfterSavingQueryUser = 0
63710     End If
63720   End If
63730   tstr = reg.GetRegistryValue("PrintAfterSavingTumble")
63740   If IsNumeric(tstr) Then
63750     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
63760       .PrintAfterSavingTumble = CLng(tstr)
63770      Else
63780       If UseStandard Then
63790        .PrintAfterSavingTumble = 0
63800       End If
63810     End If
63820    Else
63830     If UseStandard Then
63840      .PrintAfterSavingTumble = 0
63850     End If
63860   End If
63870   tstr = reg.GetRegistryValue("PrinterStop")
63880   If IsNumeric(tstr) Then
63890     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
63900       .PrinterStop = CLng(tstr)
63910      Else
63920       If UseStandard Then
63930        .PrinterStop = 0
63940       End If
63950     End If
63960    Else
63970     If UseStandard Then
63980      .PrinterStop = 0
63990     End If
64000   End If
64010   tstr = reg.GetRegistryValue("PrinterTemppath")
64020   If LenB(Trim$(tstr)) > 0 Then
64030    If DirExists(GetSubstFilename2(tstr, False)) = True Then
64040      .PrinterTemppath = tstr
64050     Else
64060      MakePath ResolveEnvironment(GetSubstFilename2(tstr, False))
64070      If DirExists(ResolveEnvironment(GetSubstFilename2(tstr, False))) = False Then
64080        If UseStandard Then
64090          .PrinterTemppath = GetTempPath
64100         Else
64110          .PrinterTemppath = ""
64120          If NoMsg = False Then
64130           MsgBox "PrinterTemppath: '" & tstr & "' = '" & ResolveEnvironment(GetSubstFilename2(tstr, False)) & "'" & _
           vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
64150          End If
64160        End If
64170       Else
64180        .PrinterTemppath = tstr
64190      End If
64200    End If
64210   End If
64220   tstr = reg.GetRegistryValue("ProcessPriority")
64230   If IsNumeric(tstr) Then
64240     If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
64250       .ProcessPriority = CLng(tstr)
64260      Else
64270       If UseStandard Then
64280        .ProcessPriority = 1
64290       End If
64300     End If
64310    Else
64320     If UseStandard Then
64330      .ProcessPriority = 1
64340     End If
64350   End If
64360   tstr = reg.GetRegistryValue("ProgramFont")
64370   If LenB(tstr) = 0 And LenB("MS Sans Serif") > 0 Then
64380     If UseStandard Then
64390      .ProgramFont = "MS Sans Serif"
64400     End If
64410    Else
64420     If LenB(tstr) > 0 Then
64430      .ProgramFont = tstr
64440     End If
64450   End If
64460   tstr = reg.GetRegistryValue("ProgramFontCharset")
64470   If IsNumeric(tstr) Then
64480     If CLng(tstr) >= 0 Then
64490       .ProgramFontCharset = CLng(tstr)
64500      Else
64510       If UseStandard Then
64520        .ProgramFontCharset = 0
64530       End If
64540     End If
64550    Else
64560     If UseStandard Then
64570      .ProgramFontCharset = 0
64580     End If
64590   End If
64600   tstr = reg.GetRegistryValue("ProgramFontSize")
64610   If IsNumeric(tstr) Then
64620     If CLng(tstr) >= 1 And CLng(tstr) <= 72 Then
64630       .ProgramFontSize = CLng(tstr)
64640      Else
64650       If UseStandard Then
64660        .ProgramFontSize = 8
64670       End If
64680     End If
64690    Else
64700     If UseStandard Then
64710      .ProgramFontSize = 8
64720     End If
64730   End If
64740   tstr = reg.GetRegistryValue("RemoveAllKnownFileExtensions")
64750   If IsNumeric(tstr) Then
64760     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
64770       .RemoveAllKnownFileExtensions = CLng(tstr)
64780      Else
64790       If UseStandard Then
64800        .RemoveAllKnownFileExtensions = 1
64810       End If
64820     End If
64830    Else
64840     If UseStandard Then
64850      .RemoveAllKnownFileExtensions = 1
64860     End If
64870   End If
64880   tstr = reg.GetRegistryValue("RemoveSpaces")
64890   If IsNumeric(tstr) Then
64900     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
64910       .RemoveSpaces = CLng(tstr)
64920      Else
64930       If UseStandard Then
64940        .RemoveSpaces = 1
64950       End If
64960     End If
64970    Else
64980     If UseStandard Then
64990      .RemoveSpaces = 1
65000     End If
65010   End If
65020   tstr = reg.GetRegistryValue("RunProgramAfterSaving")
65030   If IsNumeric(tstr) Then
65040     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
65050       .RunProgramAfterSaving = CLng(tstr)
65060      Else
65070       If UseStandard Then
65080        .RunProgramAfterSaving = 0
65090       End If
65100     End If
65110    Else
65120     If UseStandard Then
65130      .RunProgramAfterSaving = 0
65140     End If
65150   End If
65160   tstr = reg.GetRegistryValue("RunProgramAfterSavingProgramname")
65170   If LenB(tstr) = 0 And LenB("") > 0 Then
65180     If UseStandard Then
65190      .RunProgramAfterSavingProgramname = " "
65200     End If
65210    Else
65220     If LenB(tstr) > 0 Then
65230      .RunProgramAfterSavingProgramname = tstr
65240     End If
65250   End If
65260   tstr = reg.GetRegistryValue("RunProgramAfterSavingProgramParameters")
65270   If LenB(tstr) = 0 And LenB("") > 0 Then
65280     If UseStandard Then
65290      .RunProgramAfterSavingProgramParameters = " "
65300     End If
65310    Else
65320     If LenB(tstr) > 0 Then
65330      .RunProgramAfterSavingProgramParameters = tstr
65340     End If
65350   End If
65360   tstr = reg.GetRegistryValue("RunProgramAfterSavingWaitUntilReady")
65370   If IsNumeric(tstr) Then
65380     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
65390       .RunProgramAfterSavingWaitUntilReady = CLng(tstr)
65400      Else
65410       If UseStandard Then
65420        .RunProgramAfterSavingWaitUntilReady = 1
65430       End If
65440     End If
65450    Else
65460     If UseStandard Then
65470      .RunProgramAfterSavingWaitUntilReady = 1
65480     End If
65490   End If
65500   tstr = reg.GetRegistryValue("RunProgramAfterSavingWindowstyle")
65510   If IsNumeric(tstr) Then
65520     If CLng(tstr) >= 0 And CLng(tstr) <= 6 Then
65530       .RunProgramAfterSavingWindowstyle = CLng(tstr)
65540      Else
65550       If UseStandard Then
65560        .RunProgramAfterSavingWindowstyle = 1
65570       End If
65580     End If
65590    Else
65600     If UseStandard Then
65610      .RunProgramAfterSavingWindowstyle = 1
65620     End If
65630   End If
65640   tstr = reg.GetRegistryValue("RunProgramBeforeSaving")
65650   If IsNumeric(tstr) Then
65660     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
65670       .RunProgramBeforeSaving = CLng(tstr)
65680      Else
65690       If UseStandard Then
65700        .RunProgramBeforeSaving = 0
65710       End If
65720     End If
65730    Else
65740     If UseStandard Then
65750      .RunProgramBeforeSaving = 0
65760     End If
65770   End If
65780   tstr = reg.GetRegistryValue("RunProgramBeforeSavingProgramname")
65790   If LenB(tstr) = 0 And LenB("") > 0 Then
65800     If UseStandard Then
65810      .RunProgramBeforeSavingProgramname = " "
65820     End If
65830    Else
65840     If LenB(tstr) > 0 Then
65850      .RunProgramBeforeSavingProgramname = tstr
65860     End If
65870   End If
65880   tstr = reg.GetRegistryValue("RunProgramBeforeSavingProgramParameters")
65890   If LenB(tstr) = 0 And LenB("") > 0 Then
65900     If UseStandard Then
65910      .RunProgramBeforeSavingProgramParameters = " "
65920     End If
65930    Else
65940     If LenB(tstr) > 0 Then
65950      .RunProgramBeforeSavingProgramParameters = tstr
65960     End If
65970   End If
65980   tstr = reg.GetRegistryValue("RunProgramBeforeSavingWindowstyle")
65990   If IsNumeric(tstr) Then
66000     If CLng(tstr) >= 0 And CLng(tstr) <= 6 Then
66010       .RunProgramBeforeSavingWindowstyle = CLng(tstr)
66020      Else
66030       If UseStandard Then
66040        .RunProgramBeforeSavingWindowstyle = 1
66050       End If
66060     End If
66070    Else
66080     If UseStandard Then
66090      .RunProgramBeforeSavingWindowstyle = 1
66100     End If
66110   End If
66120   tstr = reg.GetRegistryValue("SaveFilename")
66130   If LenB(tstr) = 0 And LenB("<Title>") > 0 Then
66140     If UseStandard Then
66150      .SaveFilename = "<Title>"
66160     End If
66170    Else
66180     If LenB(tstr) > 0 Then
66190      .SaveFilename = tstr
66200     End If
66210   End If
66220   tstr = reg.GetRegistryValue("SendMailMethod")
66230   If IsNumeric(tstr) Then
66240     If CLng(tstr) >= 0 Then
66250       .SendMailMethod = CLng(tstr)
66260      Else
66270       If UseStandard Then
66280        .SendMailMethod = 0
66290       End If
66300     End If
66310    Else
66320     If UseStandard Then
66330      .SendMailMethod = 0
66340     End If
66350   End If
66360   tstr = reg.GetRegistryValue("ShowAnimation")
66370   If IsNumeric(tstr) Then
66380     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
66390       .ShowAnimation = CLng(tstr)
66400      Else
66410       If UseStandard Then
66420        .ShowAnimation = 1
66430       End If
66440     End If
66450    Else
66460     If UseStandard Then
66470      .ShowAnimation = 1
66480     End If
66490   End If
66500   tstr = reg.GetRegistryValue("StartStandardProgram")
66510   If IsNumeric(tstr) Then
66520     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
66530       .StartStandardProgram = CLng(tstr)
66540      Else
66550       If UseStandard Then
66560        .StartStandardProgram = 1
66570       End If
66580     End If
66590    Else
66600     If UseStandard Then
66610      .StartStandardProgram = 1
66620     End If
66630   End If
66640   tstr = reg.GetRegistryValue("Toolbars")
66650   If IsNumeric(tstr) Then
66660     If CLng(tstr) >= 0 Then
66670       .Toolbars = CLng(tstr)
66680      Else
66690       If UseStandard Then
66700        .Toolbars = 1
66710       End If
66720     End If
66730    Else
66740     If UseStandard Then
66750      .Toolbars = 1
66760     End If
66770   End If
66780   tstr = reg.GetRegistryValue("UseAutosave")
66790   If IsNumeric(tstr) Then
66800     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
66810       .UseAutosave = CLng(tstr)
66820      Else
66830       If UseStandard Then
66840        .UseAutosave = 0
66850       End If
66860     End If
66870    Else
66880     If UseStandard Then
66890      .UseAutosave = 0
66900     End If
66910   End If
66920   tstr = reg.GetRegistryValue("UseAutosaveDirectory")
66930   If IsNumeric(tstr) Then
66940     If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
66950       .UseAutosaveDirectory = CLng(tstr)
66960      Else
66970       If UseStandard Then
66980        .UseAutosaveDirectory = 1
66990       End If
67000     End If
67010    Else
67020     If UseStandard Then
67030      .UseAutosaveDirectory = 1
67040     End If
67050   End If
67060  End With
67070  Set reg = Nothing
67080  ReadOptionsReg = myOptions
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
50210   reg.SetRegistryValue "DeviceHeightPoints", CStr(.DeviceHeightPoints), REG_SZ
50220   reg.SetRegistryValue "DeviceWidthPoints", CStr(.DeviceWidthPoints), REG_SZ
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
50680   reg.SetRegistryValue "PDFCompressionColorResample", CStr(Abs(.PDFCompressionColorResample)), REG_SZ
50690   reg.SetRegistryValue "PDFCompressionColorResampleChoice", CStr(.PDFCompressionColorResampleChoice), REG_SZ
50700   reg.SetRegistryValue "PDFCompressionColorResolution", CStr(.PDFCompressionColorResolution), REG_SZ
50710   reg.SetRegistryValue "PDFCompressionGreyCompression", CStr(Abs(.PDFCompressionGreyCompression)), REG_SZ
50720   reg.SetRegistryValue "PDFCompressionGreyCompressionChoice", CStr(.PDFCompressionGreyCompressionChoice), REG_SZ
50730   reg.SetRegistryValue "PDFCompressionGreyResample", CStr(Abs(.PDFCompressionGreyResample)), REG_SZ
50740   reg.SetRegistryValue "PDFCompressionGreyResampleChoice", CStr(.PDFCompressionGreyResampleChoice), REG_SZ
50750   reg.SetRegistryValue "PDFCompressionGreyResolution", CStr(.PDFCompressionGreyResolution), REG_SZ
50760   reg.SetRegistryValue "PDFCompressionMonoCompression", CStr(Abs(.PDFCompressionMonoCompression)), REG_SZ
50770   reg.SetRegistryValue "PDFCompressionMonoCompressionChoice", CStr(.PDFCompressionMonoCompressionChoice), REG_SZ
50780   reg.SetRegistryValue "PDFCompressionMonoResample", CStr(Abs(.PDFCompressionMonoResample)), REG_SZ
50790   reg.SetRegistryValue "PDFCompressionMonoResampleChoice", CStr(.PDFCompressionMonoResampleChoice), REG_SZ
50800   reg.SetRegistryValue "PDFCompressionMonoResolution", CStr(.PDFCompressionMonoResolution), REG_SZ
50810   reg.SetRegistryValue "PDFCompressionTextCompression", CStr(Abs(.PDFCompressionTextCompression)), REG_SZ
50820   reg.Subkey = "Printing\Formats\PDF\Fonts"
50830   If Not reg.KeyExists Then
50840    reg.CreateKey
50850   End If
50860   reg.SetRegistryValue "PDFFontsEmbedAll", CStr(Abs(.PDFFontsEmbedAll)), REG_SZ
50870   reg.SetRegistryValue "PDFFontsSubSetFonts", CStr(Abs(.PDFFontsSubSetFonts)), REG_SZ
50880   reg.SetRegistryValue "PDFFontsSubSetFontsPercent", CStr(.PDFFontsSubSetFontsPercent), REG_SZ
50890   reg.Subkey = "Printing\Formats\PDF\General"
50900   If Not reg.KeyExists Then
50910    reg.CreateKey
50920   End If
50930   reg.SetRegistryValue "PDFGeneralASCII85", CStr(Abs(.PDFGeneralASCII85)), REG_SZ
50940   reg.SetRegistryValue "PDFGeneralAutorotate", CStr(.PDFGeneralAutorotate), REG_SZ
50950   reg.SetRegistryValue "PDFGeneralCompatibility", CStr(.PDFGeneralCompatibility), REG_SZ
50960   reg.SetRegistryValue "PDFGeneralOverprint", CStr(.PDFGeneralOverprint), REG_SZ
50970   reg.SetRegistryValue "PDFGeneralResolution", CStr(.PDFGeneralResolution), REG_SZ
50980   reg.SetRegistryValue "PDFOptimize", CStr(Abs(.PDFOptimize)), REG_SZ
50990   reg.Subkey = "Printing\Formats\PDF\Security"
51000   If Not reg.KeyExists Then
51010    reg.CreateKey
51020   End If
51030   reg.SetRegistryValue "PDFAllowAssembly", CStr(Abs(.PDFAllowAssembly)), REG_SZ
51040   reg.SetRegistryValue "PDFAllowDegradedPrinting", CStr(Abs(.PDFAllowDegradedPrinting)), REG_SZ
51050   reg.SetRegistryValue "PDFAllowFillIn", CStr(Abs(.PDFAllowFillIn)), REG_SZ
51060   reg.SetRegistryValue "PDFAllowScreenReaders", CStr(Abs(.PDFAllowScreenReaders)), REG_SZ
51070   reg.SetRegistryValue "PDFDisallowCopy", CStr(Abs(.PDFDisallowCopy)), REG_SZ
51080   reg.SetRegistryValue "PDFDisallowModifyAnnotations", CStr(Abs(.PDFDisallowModifyAnnotations)), REG_SZ
51090   reg.SetRegistryValue "PDFDisallowModifyContents", CStr(Abs(.PDFDisallowModifyContents)), REG_SZ
51100   reg.SetRegistryValue "PDFDisallowPrinting", CStr(Abs(.PDFDisallowPrinting)), REG_SZ
51110   reg.SetRegistryValue "PDFEncryptor", CStr(.PDFEncryptor), REG_SZ
51120   reg.SetRegistryValue "PDFHighEncryption", CStr(Abs(.PDFHighEncryption)), REG_SZ
51130   reg.SetRegistryValue "PDFLowEncryption", CStr(Abs(.PDFLowEncryption)), REG_SZ
51140   reg.SetRegistryValue "PDFOwnerPass", CStr(Abs(.PDFOwnerPass)), REG_SZ
51150   reg.SetRegistryValue "PDFOwnerPasswordString", CStr(.PDFOwnerPasswordString), REG_SZ
51160   reg.SetRegistryValue "PDFUserPass", CStr(Abs(.PDFUserPass)), REG_SZ
51170   reg.SetRegistryValue "PDFUserPasswordString", CStr(.PDFUserPasswordString), REG_SZ
51180   reg.SetRegistryValue "PDFUseSecurity", CStr(Abs(.PDFUseSecurity)), REG_SZ
51190   reg.Subkey = "Printing\Formats\PS\LanguageLevel"
51200   If Not reg.KeyExists Then
51210    reg.CreateKey
51220   End If
51230   reg.SetRegistryValue "EPSLanguageLevel", CStr(.EPSLanguageLevel), REG_SZ
51240   reg.SetRegistryValue "PSLanguageLevel", CStr(.PSLanguageLevel), REG_SZ
51250   reg.Subkey = "Program"
51260   If Not reg.KeyExists Then
51270    reg.CreateKey
51280   End If
51290   reg.SetRegistryValue "AdditionalGhostscriptParameters", CStr(.AdditionalGhostscriptParameters), REG_SZ
51300   reg.SetRegistryValue "AdditionalGhostscriptSearchpath", CStr(.AdditionalGhostscriptSearchpath), REG_SZ
51310   reg.SetRegistryValue "AddWindowsFontpath", CStr(Abs(.AddWindowsFontpath)), REG_SZ
51320   reg.SetRegistryValue "AutosaveDirectory", CStr(.AutosaveDirectory), REG_SZ
51330   reg.SetRegistryValue "AutosaveFilename", CStr(.AutosaveFilename), REG_SZ
51340   reg.SetRegistryValue "AutosaveFormat", CStr(.AutosaveFormat), REG_SZ
51350   reg.SetRegistryValue "ClientComputerResolveIPAddress", CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
51360   reg.SetRegistryValue "DisableEmail", CStr(Abs(.DisableEmail)), REG_SZ
51370   reg.SetRegistryValue "DontUseDocumentSettings", CStr(Abs(.DontUseDocumentSettings)), REG_SZ
51380   reg.SetRegistryValue "FilenameSubstitutions", CStr(.FilenameSubstitutions), REG_SZ
51390   reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle", CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
51400   reg.SetRegistryValue "Language", CStr(.Language), REG_SZ
51410   reg.SetRegistryValue "LastSaveDirectory", CStr(.LastSaveDirectory), REG_SZ
51420   reg.SetRegistryValue "Logging", CStr(Abs(.Logging)), REG_SZ
51430   reg.SetRegistryValue "LogLines", CStr(.LogLines), REG_SZ
51440   reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter", CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
51450   reg.SetRegistryValue "NoProcessingAtStartup", CStr(Abs(.NoProcessingAtStartup)), REG_SZ
51460   reg.SetRegistryValue "OptionsDesign", CStr(.OptionsDesign), REG_SZ
51470   reg.SetRegistryValue "OptionsEnabled", CStr(Abs(.OptionsEnabled)), REG_SZ
51480   reg.SetRegistryValue "OptionsVisible", CStr(Abs(.OptionsVisible)), REG_SZ
51490   reg.SetRegistryValue "PrintAfterSaving", CStr(Abs(.PrintAfterSaving)), REG_SZ
51500   reg.SetRegistryValue "PrintAfterSavingDuplex", CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
51510   reg.SetRegistryValue "PrintAfterSavingNoCancel", CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
51520   reg.SetRegistryValue "PrintAfterSavingPrinter", CStr(.PrintAfterSavingPrinter), REG_SZ
51530   reg.SetRegistryValue "PrintAfterSavingQueryUser", CStr(.PrintAfterSavingQueryUser), REG_SZ
51540   reg.SetRegistryValue "PrintAfterSavingTumble", CStr(Abs(.PrintAfterSavingTumble)), REG_SZ
51550   reg.SetRegistryValue "PrinterStop", CStr(Abs(.PrinterStop)), REG_SZ
51560   reg.SetRegistryValue "PrinterTemppath", CStr(.PrinterTemppath), REG_SZ
51570   reg.SetRegistryValue "ProcessPriority", CStr(.ProcessPriority), REG_SZ
51580   reg.SetRegistryValue "ProgramFont", CStr(.ProgramFont), REG_SZ
51590   reg.SetRegistryValue "ProgramFontCharset", CStr(.ProgramFontCharset), REG_SZ
51600   reg.SetRegistryValue "ProgramFontSize", CStr(.ProgramFontSize), REG_SZ
51610   reg.SetRegistryValue "RemoveAllKnownFileExtensions", CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
51620   reg.SetRegistryValue "RemoveSpaces", CStr(Abs(.RemoveSpaces)), REG_SZ
51630   reg.SetRegistryValue "RunProgramAfterSaving", CStr(Abs(.RunProgramAfterSaving)), REG_SZ
51640   reg.SetRegistryValue "RunProgramAfterSavingProgramname", CStr(.RunProgramAfterSavingProgramname), REG_SZ
51650   reg.SetRegistryValue "RunProgramAfterSavingProgramParameters", CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
51660   reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady", CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
51670   reg.SetRegistryValue "RunProgramAfterSavingWindowstyle", CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
51680   reg.SetRegistryValue "RunProgramBeforeSaving", CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
51690   reg.SetRegistryValue "RunProgramBeforeSavingProgramname", CStr(.RunProgramBeforeSavingProgramname), REG_SZ
51700   reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters", CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
51710   reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle", CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
51720   reg.SetRegistryValue "SaveFilename", CStr(.SaveFilename), REG_SZ
51730   reg.SetRegistryValue "SendMailMethod", CStr(.SendMailMethod), REG_SZ
51740   reg.SetRegistryValue "ShowAnimation", CStr(Abs(.ShowAnimation)), REG_SZ
51750   reg.SetRegistryValue "StartStandardProgram", CStr(Abs(.StartStandardProgram)), REG_SZ
51760   reg.SetRegistryValue "Toolbars", CStr(.Toolbars), REG_SZ
51770   reg.SetRegistryValue "UseAutosave", CStr(Abs(.UseAutosave)), REG_SZ
51780   reg.SetRegistryValue "UseAutosaveDirectory", CStr(Abs(.UseAutosaveDirectory)), REG_SZ
51790  End With
51800  Set reg = Nothing
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
  Frm.cmbProgramFontsize.Text = .ProgramFontSize
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
50010  Dim i As Long, tstr As String, lsv As ListView
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
50130  tstr = ""
50140  Set lsv = Frm.lsvFilenameSubst
50150  For i = 1 To lsv.ListItems.Count
50160   If i < lsv.ListItems.Count Then
50170     tstr = tstr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1) & "\"
50180    Else
50190     tstr = tstr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1)
50200   End If
50210  Next i
50220  .FilenameSubstitutions = tstr
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
50800  .ProgramFontSize = Frm.cmbProgramFontsize.Text
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
50090  SaveOptions Options
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
50060  SaveOptions Options
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
50010  Dim hOpt As clsHash, tstr
50020  Set hOpt = New clsHash
50030  ReadINISection PDFCreatorINIFile, "Options", hOpt
50040  tstr = Trim$(hOpt.Retrieve("Language"))
50050  If LenB(tstr) > 0 Then
50060    ReadLanguageFromOptionsINI = tstr
50070   Else
50080    If UseStandard Then
50090      ReadLanguageFromOptionsINI = "english"
50100     Else
50110      ReadLanguageFromOptionsINI = Language
50120    End If
50130  End If
50140  Set hOpt = Nothing
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
50010  Dim reg As clsRegistry, tstr As String
50020  Set reg = New clsRegistry
50030  With reg
50040   .KeyRoot = KeyRoot
50050   .Subkey = "Program"
50060   .hkey = hProfile
50070   tstr = Trim$(reg.GetRegistryValue("Language"))
50080  End With
50090  If LenB(tstr) > 0 Then
50100    ReadLanguageFromOptionsReg = tstr
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
50010  Dim reg As clsRegistry, tstr As String
50020  Set reg = New clsRegistry
50030  UseINI = False
50040  With reg
50050   .hkey = HKEY_LOCAL_MACHINE
50060   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50070   tstr = Trim$(.GetRegistryValue("UseINI"))
50080   If tstr = "1" Then
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

