Attribute VB_Name = "modOptions"
Option Explicit

' Module automatically generated with LanguagesTool from Frank Heindörfer
' 2003
' Email: thesmilyface@users.sourceforge.net

Public Type tOptions
 AdditionalGhostscriptSearchpath As String
 AddWindowsFontpath As Long
 AutosaveDirectory As String
 AutosaveFilename As String
 AutosaveFormat As Long
 BitmapResolution As Long
 BMPColorscount As Long
 DeviceHeightPoints As Double
 DeviceWidthPoints As Double
 DirectoryGhostscriptBinaries As String
 DirectoryGhostscriptFonts As String
 DirectoryGhostscriptLibraries As String
 DirectoryGhostscriptResource As String
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
50030   .AdditionalGhostscriptSearchpath = " "
50040   .AddWindowsFontpath = "1"
50050   .AutosaveDirectory = " "
50060   .AutosaveFilename = "<DateTime>"
50070   .AutosaveFormat = "0"
50080   .BitmapResolution = "150"
50090   .BMPColorscount = "1"
50100   .DeviceHeightPoints = "-1"
50110   .DeviceWidthPoints = "-1"
50120   Set reg = New clsRegistry
50130   reg.hkey = HKEY_LOCAL_MACHINE
50140   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50150   .DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50160   Set reg = Nothing
50170   Set reg = New clsRegistry
50180   reg.hkey = HKEY_LOCAL_MACHINE
50190   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50200   .DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50210   Set reg = Nothing
50220   Set reg = New clsRegistry
50230   reg.hkey = HKEY_LOCAL_MACHINE
50240   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50250   .DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50260   Set reg = Nothing
50270   Set reg = New clsRegistry
50280   reg.hkey = HKEY_LOCAL_MACHINE
50290   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50300   .DirectoryGhostscriptResource = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryResource"))
50310   Set reg = Nothing
50320   .DontUseDocumentSettings = "0"
50330   .EPSLanguageLevel = "2"
50340   .FilenameSubstitutions = "Microsoft Word - \.doc"
50350   .FilenameSubstitutionsOnlyInTitle = "1"
50360   .JPEGColorscount = "0"
50370   .JPEGQuality = "75"
50380   .Language = "english"
50390   .LastSaveDirectory = GetMyFiles
50400   .Logging = "0"
50410   .LogLines = "100"
50420   .NoConfirmMessageSwitchingDefaultprinter = "0"
50430   .NoProcessingAtStartup = "0"
50440   .OnePagePerFile = "0"
50450   .OptionsDesign = "1"
50460   .OptionsEnabled = "1"
50470   .OptionsVisible = "1"
50480   .Papersize = " "
50490   .PCXColorscount = "0"
50500   .PDFAllowAssembly = "0"
50510   .PDFAllowDegradedPrinting = "0"
50520   .PDFAllowFillIn = "0"
50530   .PDFAllowScreenReaders = "0"
50540   .PDFColorsCMYKToRGB = "1"
50550   .PDFColorsColorModel = "1"
50560   .PDFColorsPreserveHalftone = "0"
50570   .PDFColorsPreserveOverprint = "1"
50580   .PDFColorsPreserveTransfer = "1"
50590   .PDFCompressionColorCompression = "1"
50600   .PDFCompressionColorCompressionChoice = "0"
50610   .PDFCompressionColorResample = "0"
50620   .PDFCompressionColorResampleChoice = "0"
50630   .PDFCompressionColorResolution = "300"
50640   .PDFCompressionGreyCompression = "1"
50650   .PDFCompressionGreyCompressionChoice = "0"
50660   .PDFCompressionGreyResample = "0"
50670   .PDFCompressionGreyResampleChoice = "0"
50680   .PDFCompressionGreyResolution = "300"
50690   .PDFCompressionMonoCompression = "1"
50700   .PDFCompressionMonoCompressionChoice = "0"
50710   .PDFCompressionMonoResample = "0"
50720   .PDFCompressionMonoResampleChoice = "0"
50730   .PDFCompressionMonoResolution = "1200"
50740   .PDFCompressionTextCompression = "1"
50750   .PDFDisallowCopy = "1"
50760   .PDFDisallowModifyAnnotations = "0"
50770   .PDFDisallowModifyContents = "0"
50780   .PDFDisallowPrinting = "0"
50790   .PDFEncryptor = "0"
50800   .PDFFontsEmbedAll = "1"
50810   .PDFFontsSubSetFonts = "1"
50820   .PDFFontsSubSetFontsPercent = "100"
50830   .PDFGeneralASCII85 = "0"
50840   .PDFGeneralAutorotate = "2"
50850   .PDFGeneralCompatibility = "1"
50860   .PDFGeneralOverprint = "0"
50870   .PDFGeneralResolution = "600"
50880   .PDFHighEncryption = "0"
50890   .PDFLowEncryption = "1"
50900   .PDFOptimize = "0"
50910   .PDFOwnerPass = "0"
50920   .PDFOwnerPasswordString = " "
50930   .PDFUserPass = "0"
50940   .PDFUserPasswordString = " "
50950   .PDFUseSecurity = "0"
50960   .PNGColorscount = "0"
50970   .PrinterStop = "0"
50980   .PrinterTemppath = GetTempPath
50990   .ProcessPriority = "1"
51000   .ProgramFont = "MS Sans Serif"
51010   .ProgramFontCharset = "0"
51020   .ProgramFontSize = "8"
51030   .PSLanguageLevel = "2"
51040   .RemoveAllKnownFileExtensions = "1"
51050   .RemoveSpaces = "1"
51060   .RunProgramAfterSaving = "0"
51070   .RunProgramAfterSavingProgramname = " "
51080   .RunProgramAfterSavingProgramParameters = " "
51090   .RunProgramAfterSavingWaitUntilReady = "1"
51100   .RunProgramAfterSavingWindowstyle = "1"
51110   .RunProgramBeforeSaving = "0"
51120   .RunProgramBeforeSavingProgramname = " "
51130   .RunProgramBeforeSavingProgramParameters = " "
51140   .RunProgramBeforeSavingWindowstyle = "1"
51150   .SaveFilename = "<Title>"
51160   .SendMailMethod = "0"
51170   .ShowAnimation = "1"
51180   .StampFontColor = "#FF0000"
51190   .StampFontname = "Arial"
51200   .StampFontsize = "48"
51210   .StampOutlineFontthickness = "0"
51220   .StampString = " "
51230   .StampUseOutlineFont = "1"
51240   .StandardAuthor = " "
51250   .StandardCreationdate = " "
51260   .StandardDateformat = "YYYYMMDDHHNNSS"
51270   .StandardKeywords = " "
51280   .StandardMailDomain = " "
51290   .StandardModifydate = " "
51300   .StandardSaveformat = "pdf"
51310   .StandardSubject = " "
51320   .StandardTitle = " "
51330   .StartStandardProgram = "1"
51340   .TIFFColorscount = "0"
51350   .Toolbars = "1"
51360   .UseAutosave = "0"
51370   .UseAutosaveDirectory = "1"
51380   .UseCreationDateNow = "0"
51390   .UseStandardAuthor = "0"
51400  End With
51410  StandardOptions = myOptions
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

Public Function ReadOptions() As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI, myOptions As tOptions, tstr As String, hOpt As New clsHash
50020  Set ini = New clsINI
50030  ini.Filename = PDFCreatorINIFile
50040  ini.Section = "Options"
50050  If ini.CheckIniFile = False Then
50060   ReadOptions = StandardOptions
50070   Exit Function
50080  End If
50090  ReadINISection PDFCreatorINIFile, "Options", hOpt
50100  With myOptions
50110   tstr = hOpt.Retrieve("AdditionalGhostscriptSearchpath", " ")
50120   .AdditionalGhostscriptSearchpath = tstr
50130   tstr = hOpt.Retrieve("AddWindowsFontpath")
50140   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
50150     .AddWindowsFontpath = CLng(tstr)
50160    Else
50170     .AddWindowsFontpath = 1
50180   End If
50190   tstr = hOpt.Retrieve("AutosaveDirectory", GetMyFiles)
50200   .AutosaveDirectory = CompletePath(tstr)
50210   tstr = hOpt.Retrieve("AutosaveFilename", "<DateTime>")
50220   .AutosaveFilename = tstr
50230   tstr = hOpt.Retrieve("AutosaveFormat")
50240   If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
50250     .AutosaveFormat = CLng(tstr)
50260    Else
50270     .AutosaveFormat = 0
50280   End If
50290   tstr = hOpt.Retrieve("BitmapResolution")
50300   If CLng(tstr) >= 1 Then
50310     .BitmapResolution = CLng(tstr)
50320    Else
50330     .BitmapResolution = 150
50340   End If
50350   tstr = hOpt.Retrieve("BMPColorscount")
50360   If CLng(tstr) >= 0 And CLng(tstr) <= 6 Then
50370     .BMPColorscount = CLng(tstr)
50380    Else
50390     .BMPColorscount = 1
50400   End If
50410   tstr = hOpt.Retrieve("DeviceHeightPoints")
50420   If CDbl(tstr) >= -1 Then
50430     .DeviceHeightPoints = CDbl(tstr)
50440    Else
50450     .DeviceHeightPoints = -1
50460   End If
50470   tstr = hOpt.Retrieve("DeviceWidthPoints")
50480   If CDbl(tstr) >= -1 Then
50490     .DeviceWidthPoints = CDbl(tstr)
50500    Else
50510     .DeviceWidthPoints = -1
50520   End If
50530   tstr = hOpt.Retrieve("DirectoryGhostscriptBinaries", App.Path)
50540   If DirExists(tstr) = True Then
50550     .DirectoryGhostscriptBinaries = CompletePath(tstr)
50560    Else
50570     .DirectoryGhostscriptBinaries = ""
50580   End If
50590   tstr = hOpt.Retrieve("DirectoryGhostscriptFonts", App.Path & "\fonts")
50600   If DirExists(tstr) = True Then
50610     .DirectoryGhostscriptFonts = CompletePath(tstr)
50620    Else
50630     .DirectoryGhostscriptFonts = ""
50640   End If
50650   tstr = hOpt.Retrieve("DirectoryGhostscriptLibraries", App.Path & "\lib")
50660   If DirExists(tstr) = True Then
50670     .DirectoryGhostscriptLibraries = CompletePath(tstr)
50680    Else
50690     .DirectoryGhostscriptLibraries = ""
50700   End If
50710   tstr = hOpt.Retrieve("DirectoryGhostscriptResource", " ")
50720   .DirectoryGhostscriptResource = tstr
50730   tstr = hOpt.Retrieve("DontUseDocumentSettings")
50740   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
50750     .DontUseDocumentSettings = CLng(tstr)
50760    Else
50770     .DontUseDocumentSettings = 0
50780   End If
50790   tstr = hOpt.Retrieve("EPSLanguageLevel")
50800   If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
50810     .EPSLanguageLevel = CLng(tstr)
50820    Else
50830     .EPSLanguageLevel = 2
50840   End If
50850   tstr = hOpt.Retrieve("FilenameSubstitutions", "Microsoft Word - \.doc")
50860   .FilenameSubstitutions = tstr
50870   tstr = hOpt.Retrieve("FilenameSubstitutionsOnlyInTitle")
50880   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
50890     .FilenameSubstitutionsOnlyInTitle = CLng(tstr)
50900    Else
50910     .FilenameSubstitutionsOnlyInTitle = 1
50920   End If
50930   tstr = hOpt.Retrieve("JPEGColorscount")
50940   If CLng(tstr) >= 0 And CLng(tstr) <= 1 Then
50950     .JPEGColorscount = CLng(tstr)
50960    Else
50970     .JPEGColorscount = 0
50980   End If
50990   tstr = hOpt.Retrieve("JPEGQuality")
51000   If CLng(tstr) >= 0 And CLng(tstr) <= 100 Then
51010     .JPEGQuality = CLng(tstr)
51020    Else
51030     .JPEGQuality = 75
51040   End If
51050   tstr = hOpt.Retrieve("Language", "english")
51060   .Language = tstr
51070   tstr = hOpt.Retrieve("LastSaveDirectory", GetMyFiles)
51080   If DirExists(tstr) = True Then
51090     .LastSaveDirectory = CompletePath(tstr)
51100    Else
51110     .LastSaveDirectory = GetMyFiles
51120   End If
51130   tstr = hOpt.Retrieve("Logging")
51140   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51150     .Logging = CLng(tstr)
51160    Else
51170     .Logging = 0
51180   End If
51190   tstr = hOpt.Retrieve("LogLines")
51200   If CLng(tstr) >= 100 And CLng(tstr) <= 1000 Then
51210     .LogLines = CLng(tstr)
51220    Else
51230     .LogLines = 100
51240   End If
51250   tstr = hOpt.Retrieve("NoConfirmMessageSwitchingDefaultprinter")
51260   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51270     .NoConfirmMessageSwitchingDefaultprinter = CLng(tstr)
51280    Else
51290     .NoConfirmMessageSwitchingDefaultprinter = 0
51300   End If
51310   tstr = hOpt.Retrieve("NoProcessingAtStartup")
51320   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51330     .NoProcessingAtStartup = CLng(tstr)
51340    Else
51350     .NoProcessingAtStartup = 0
51360   End If
51370   tstr = hOpt.Retrieve("OnePagePerFile")
51380   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51390     .OnePagePerFile = CLng(tstr)
51400    Else
51410     .OnePagePerFile = 0
51420   End If
51430   tstr = hOpt.Retrieve("OptionsDesign")
51440   If CLng(tstr) >= 1 And CLng(tstr) <= 2 Then
51450     .OptionsDesign = CLng(tstr)
51460    Else
51470     .OptionsDesign = 1
51480   End If
51490   tstr = hOpt.Retrieve("OptionsEnabled")
51500   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51510     .OptionsEnabled = CLng(tstr)
51520    Else
51530     .OptionsEnabled = 1
51540   End If
51550   tstr = hOpt.Retrieve("OptionsVisible")
51560   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51570     .OptionsVisible = CLng(tstr)
51580    Else
51590     .OptionsVisible = 1
51600   End If
51610   tstr = hOpt.Retrieve("Papersize", " ")
51620   .Papersize = tstr
51630   tstr = hOpt.Retrieve("PCXColorscount")
51640   If CLng(tstr) >= 0 And CLng(tstr) <= 5 Then
51650     .PCXColorscount = CLng(tstr)
51660    Else
51670     .PCXColorscount = 0
51680   End If
51690   tstr = hOpt.Retrieve("PDFAllowAssembly")
51700   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51710     .PDFAllowAssembly = CLng(tstr)
51720    Else
51730     .PDFAllowAssembly = 0
51740   End If
51750   tstr = hOpt.Retrieve("PDFAllowDegradedPrinting")
51760   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51770     .PDFAllowDegradedPrinting = CLng(tstr)
51780    Else
51790     .PDFAllowDegradedPrinting = 0
51800   End If
51810   tstr = hOpt.Retrieve("PDFAllowFillIn")
51820   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51830     .PDFAllowFillIn = CLng(tstr)
51840    Else
51850     .PDFAllowFillIn = 0
51860   End If
51870   tstr = hOpt.Retrieve("PDFAllowScreenReaders")
51880   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51890     .PDFAllowScreenReaders = CLng(tstr)
51900    Else
51910     .PDFAllowScreenReaders = 0
51920   End If
51930   tstr = hOpt.Retrieve("PDFColorsCMYKToRGB")
51940   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51950     .PDFColorsCMYKToRGB = CLng(tstr)
51960    Else
51970     .PDFColorsCMYKToRGB = 1
51980   End If
51990   tstr = hOpt.Retrieve("PDFColorsColorModel")
52000   If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
52010     .PDFColorsColorModel = CLng(tstr)
52020    Else
52030     .PDFColorsColorModel = 1
52040   End If
52050   tstr = hOpt.Retrieve("PDFColorsPreserveHalftone")
52060   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52070     .PDFColorsPreserveHalftone = CLng(tstr)
52080    Else
52090     .PDFColorsPreserveHalftone = 0
52100   End If
52110   tstr = hOpt.Retrieve("PDFColorsPreserveOverprint")
52120   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52130     .PDFColorsPreserveOverprint = CLng(tstr)
52140    Else
52150     .PDFColorsPreserveOverprint = 1
52160   End If
52170   tstr = hOpt.Retrieve("PDFColorsPreserveTransfer")
52180   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52190     .PDFColorsPreserveTransfer = CLng(tstr)
52200    Else
52210     .PDFColorsPreserveTransfer = 1
52220   End If
52230   tstr = hOpt.Retrieve("PDFCompressionColorCompression")
52240   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52250     .PDFCompressionColorCompression = CLng(tstr)
52260    Else
52270     .PDFCompressionColorCompression = 1
52280   End If
52290   tstr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
52300   If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
52310     .PDFCompressionColorCompressionChoice = CLng(tstr)
52320    Else
52330     .PDFCompressionColorCompressionChoice = 0
52340   End If
52350   tstr = hOpt.Retrieve("PDFCompressionColorResample")
52360   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52370     .PDFCompressionColorResample = CLng(tstr)
52380    Else
52390     .PDFCompressionColorResample = 0
52400   End If
52410   tstr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
52420   If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
52430     .PDFCompressionColorResampleChoice = CLng(tstr)
52440    Else
52450     .PDFCompressionColorResampleChoice = 0
52460   End If
52470   tstr = hOpt.Retrieve("PDFCompressionColorResolution")
52480   If CLng(tstr) >= 0 Then
52490     .PDFCompressionColorResolution = CLng(tstr)
52500    Else
52510     .PDFCompressionColorResolution = 300
52520   End If
52530   tstr = hOpt.Retrieve("PDFCompressionGreyCompression")
52540   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52550     .PDFCompressionGreyCompression = CLng(tstr)
52560    Else
52570     .PDFCompressionGreyCompression = 1
52580   End If
52590   tstr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
52600   If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
52610     .PDFCompressionGreyCompressionChoice = CLng(tstr)
52620    Else
52630     .PDFCompressionGreyCompressionChoice = 0
52640   End If
52650   tstr = hOpt.Retrieve("PDFCompressionGreyResample")
52660   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52670     .PDFCompressionGreyResample = CLng(tstr)
52680    Else
52690     .PDFCompressionGreyResample = 0
52700   End If
52710   tstr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
52720   If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
52730     .PDFCompressionGreyResampleChoice = CLng(tstr)
52740    Else
52750     .PDFCompressionGreyResampleChoice = 0
52760   End If
52770   tstr = hOpt.Retrieve("PDFCompressionGreyResolution")
52780   If CLng(tstr) >= 0 Then
52790     .PDFCompressionGreyResolution = CLng(tstr)
52800    Else
52810     .PDFCompressionGreyResolution = 300
52820   End If
52830   tstr = hOpt.Retrieve("PDFCompressionMonoCompression")
52840   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52850     .PDFCompressionMonoCompression = CLng(tstr)
52860    Else
52870     .PDFCompressionMonoCompression = 1
52880   End If
52890   tstr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
52900   If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
52910     .PDFCompressionMonoCompressionChoice = CLng(tstr)
52920    Else
52930     .PDFCompressionMonoCompressionChoice = 0
52940   End If
52950   tstr = hOpt.Retrieve("PDFCompressionMonoResample")
52960   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52970     .PDFCompressionMonoResample = CLng(tstr)
52980    Else
52990     .PDFCompressionMonoResample = 0
53000   End If
53010   tstr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
53020   If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
53030     .PDFCompressionMonoResampleChoice = CLng(tstr)
53040    Else
53050     .PDFCompressionMonoResampleChoice = 0
53060   End If
53070   tstr = hOpt.Retrieve("PDFCompressionMonoResolution")
53080   If CLng(tstr) >= 0 Then
53090     .PDFCompressionMonoResolution = CLng(tstr)
53100    Else
53110     .PDFCompressionMonoResolution = 1200
53120   End If
53130   tstr = hOpt.Retrieve("PDFCompressionTextCompression")
53140   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53150     .PDFCompressionTextCompression = CLng(tstr)
53160    Else
53170     .PDFCompressionTextCompression = 1
53180   End If
53190   tstr = hOpt.Retrieve("PDFDisallowCopy")
53200   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53210     .PDFDisallowCopy = CLng(tstr)
53220    Else
53230     .PDFDisallowCopy = 1
53240   End If
53250   tstr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
53260   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53270     .PDFDisallowModifyAnnotations = CLng(tstr)
53280    Else
53290     .PDFDisallowModifyAnnotations = 0
53300   End If
53310   tstr = hOpt.Retrieve("PDFDisallowModifyContents")
53320   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53330     .PDFDisallowModifyContents = CLng(tstr)
53340    Else
53350     .PDFDisallowModifyContents = 0
53360   End If
53370   tstr = hOpt.Retrieve("PDFDisallowPrinting")
53380   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53390     .PDFDisallowPrinting = CLng(tstr)
53400    Else
53410     .PDFDisallowPrinting = 0
53420   End If
53430   tstr = hOpt.Retrieve("PDFEncryptor")
53440   If CLng(tstr) >= 0 And CLng(tstr) <= 1 Then
53450     .PDFEncryptor = CLng(tstr)
53460    Else
53470     .PDFEncryptor = 0
53480   End If
53490   tstr = hOpt.Retrieve("PDFFontsEmbedAll")
53500   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53510     .PDFFontsEmbedAll = CLng(tstr)
53520    Else
53530     .PDFFontsEmbedAll = 1
53540   End If
53550   tstr = hOpt.Retrieve("PDFFontsSubSetFonts")
53560   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53570     .PDFFontsSubSetFonts = CLng(tstr)
53580    Else
53590     .PDFFontsSubSetFonts = 1
53600   End If
53610   tstr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
53620   If CLng(tstr) >= 0 Then
53630     .PDFFontsSubSetFontsPercent = CLng(tstr)
53640    Else
53650     .PDFFontsSubSetFontsPercent = 100
53660   End If
53670   tstr = hOpt.Retrieve("PDFGeneralASCII85")
53680   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53690     .PDFGeneralASCII85 = CLng(tstr)
53700    Else
53710     .PDFGeneralASCII85 = 0
53720   End If
53730   tstr = hOpt.Retrieve("PDFGeneralAutorotate")
53740   If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
53750     .PDFGeneralAutorotate = CLng(tstr)
53760    Else
53770     .PDFGeneralAutorotate = 2
53780   End If
53790   tstr = hOpt.Retrieve("PDFGeneralCompatibility")
53800   If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
53810     .PDFGeneralCompatibility = CLng(tstr)
53820    Else
53830     .PDFGeneralCompatibility = 1
53840   End If
53850   tstr = hOpt.Retrieve("PDFGeneralOverprint")
53860   If CLng(tstr) >= 0 And CLng(tstr) <= 1 Then
53870     .PDFGeneralOverprint = CLng(tstr)
53880    Else
53890     .PDFGeneralOverprint = 0
53900   End If
53910   tstr = hOpt.Retrieve("PDFGeneralResolution")
53920   If CLng(tstr) >= 0 Then
53930     .PDFGeneralResolution = CLng(tstr)
53940    Else
53950     .PDFGeneralResolution = 600
53960   End If
53970   tstr = hOpt.Retrieve("PDFHighEncryption")
53980   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53990     .PDFHighEncryption = CLng(tstr)
54000    Else
54010     .PDFHighEncryption = 0
54020   End If
54030   tstr = hOpt.Retrieve("PDFLowEncryption")
54040   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54050     .PDFLowEncryption = CLng(tstr)
54060    Else
54070     .PDFLowEncryption = 1
54080   End If
54090   tstr = hOpt.Retrieve("PDFOptimize")
54100   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54110     .PDFOptimize = CLng(tstr)
54120    Else
54130     .PDFOptimize = 0
54140   End If
54150   tstr = hOpt.Retrieve("PDFOwnerPass")
54160   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54170     .PDFOwnerPass = CLng(tstr)
54180    Else
54190     .PDFOwnerPass = 0
54200   End If
54210   tstr = hOpt.Retrieve("PDFOwnerPasswordString", " ")
54220   .PDFOwnerPasswordString = tstr
54230   tstr = hOpt.Retrieve("PDFUserPass")
54240   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54250     .PDFUserPass = CLng(tstr)
54260    Else
54270     .PDFUserPass = 0
54280   End If
54290   tstr = hOpt.Retrieve("PDFUserPasswordString", " ")
54300   .PDFUserPasswordString = tstr
54310   tstr = hOpt.Retrieve("PDFUseSecurity")
54320   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54330     .PDFUseSecurity = CLng(tstr)
54340    Else
54350     .PDFUseSecurity = 0
54360   End If
54370   tstr = hOpt.Retrieve("PNGColorscount")
54380   If CLng(tstr) >= 0 And CLng(tstr) <= 4 Then
54390     .PNGColorscount = CLng(tstr)
54400    Else
54410     .PNGColorscount = 0
54420   End If
54430   tstr = hOpt.Retrieve("PrinterStop")
54440   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54450     .PrinterStop = CLng(tstr)
54460    Else
54470     .PrinterStop = 0
54480   End If
54490   tstr = hOpt.Retrieve("PrinterTemppath", GetTempPath)
54500   If DirExists(tstr) = True Then
54510     .PrinterTemppath = CompletePath(tstr)
54520    Else
54530     .PrinterTemppath = GetTempPath
54540   End If
54550   tstr = hOpt.Retrieve("ProcessPriority")
54560   If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
54570     .ProcessPriority = CLng(tstr)
54580    Else
54590     .ProcessPriority = 1
54600   End If
54610   tstr = hOpt.Retrieve("ProgramFont", "MS Sans Serif")
54620   .ProgramFont = tstr
54630   tstr = hOpt.Retrieve("ProgramFontCharset")
54640   If CLng(tstr) >= 0 Then
54650     .ProgramFontCharset = CLng(tstr)
54660    Else
54670     .ProgramFontCharset = 0
54680   End If
54690   tstr = hOpt.Retrieve("ProgramFontSize")
54700   If CLng(tstr) >= 1 And CLng(tstr) <= 72 Then
54710     .ProgramFontSize = CLng(tstr)
54720    Else
54730     .ProgramFontSize = 8
54740   End If
54750   tstr = hOpt.Retrieve("PSLanguageLevel")
54760   If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
54770     .PSLanguageLevel = CLng(tstr)
54780    Else
54790     .PSLanguageLevel = 2
54800   End If
54810   tstr = hOpt.Retrieve("RemoveAllKnownFileExtensions")
54820   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54830     .RemoveAllKnownFileExtensions = CLng(tstr)
54840    Else
54850     .RemoveAllKnownFileExtensions = 1
54860   End If
54870   tstr = hOpt.Retrieve("RemoveSpaces")
54880   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54890     .RemoveSpaces = CLng(tstr)
54900    Else
54910     .RemoveSpaces = 1
54920   End If
54930   tstr = hOpt.Retrieve("RunProgramAfterSaving")
54940   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54950     .RunProgramAfterSaving = CLng(tstr)
54960    Else
54970     .RunProgramAfterSaving = 0
54980   End If
54990   tstr = hOpt.Retrieve("RunProgramAfterSavingProgramname", " ")
55000   .RunProgramAfterSavingProgramname = tstr
55010   tstr = hOpt.Retrieve("RunProgramAfterSavingProgramParameters", " ")
55020   .RunProgramAfterSavingProgramParameters = tstr
55030   tstr = hOpt.Retrieve("RunProgramAfterSavingWaitUntilReady")
55040   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55050     .RunProgramAfterSavingWaitUntilReady = CLng(tstr)
55060    Else
55070     .RunProgramAfterSavingWaitUntilReady = 1
55080   End If
55090   tstr = hOpt.Retrieve("RunProgramAfterSavingWindowstyle")
55100   If CLng(tstr) >= 0 And CLng(tstr) <= 6 Then
55110     .RunProgramAfterSavingWindowstyle = CLng(tstr)
55120    Else
55130     .RunProgramAfterSavingWindowstyle = 1
55140   End If
55150   tstr = hOpt.Retrieve("RunProgramBeforeSaving")
55160   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55170     .RunProgramBeforeSaving = CLng(tstr)
55180    Else
55190     .RunProgramBeforeSaving = 0
55200   End If
55210   tstr = hOpt.Retrieve("RunProgramBeforeSavingProgramname", " ")
55220   .RunProgramBeforeSavingProgramname = tstr
55230   tstr = hOpt.Retrieve("RunProgramBeforeSavingProgramParameters", " ")
55240   .RunProgramBeforeSavingProgramParameters = tstr
55250   tstr = hOpt.Retrieve("RunProgramBeforeSavingWindowstyle")
55260   If CLng(tstr) >= 0 And CLng(tstr) <= 6 Then
55270     .RunProgramBeforeSavingWindowstyle = CLng(tstr)
55280    Else
55290     .RunProgramBeforeSavingWindowstyle = 1
55300   End If
55310   tstr = hOpt.Retrieve("SaveFilename", "<Title>")
55320   .SaveFilename = tstr
55330   tstr = hOpt.Retrieve("SendMailMethod")
55340   If CLng(tstr) >= 0 Then
55350     .SendMailMethod = CLng(tstr)
55360    Else
55370     .SendMailMethod = 0
55380   End If
55390   tstr = hOpt.Retrieve("ShowAnimation")
55400   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55410     .ShowAnimation = CLng(tstr)
55420    Else
55430     .ShowAnimation = 1
55440   End If
55450   tstr = hOpt.Retrieve("StampFontColor", "#FF0000")
55460   .StampFontColor = tstr
55470   tstr = hOpt.Retrieve("StampFontname", "Arial")
55480   .StampFontname = tstr
55490   tstr = hOpt.Retrieve("StampFontsize")
55500   If CLng(tstr) >= 1 Then
55510     .StampFontsize = CLng(tstr)
55520    Else
55530     .StampFontsize = 48
55540   End If
55550   tstr = hOpt.Retrieve("StampOutlineFontthickness")
55560   If CLng(tstr) >= 0 Then
55570     .StampOutlineFontthickness = CLng(tstr)
55580    Else
55590     .StampOutlineFontthickness = 0
55600   End If
55610   tstr = hOpt.Retrieve("StampString", " ")
55620   .StampString = tstr
55630   tstr = hOpt.Retrieve("StampUseOutlineFont")
55640   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55650     .StampUseOutlineFont = CLng(tstr)
55660    Else
55670     .StampUseOutlineFont = 1
55680   End If
55690   tstr = hOpt.Retrieve("StandardAuthor", " ")
55700   .StandardAuthor = tstr
55710   tstr = hOpt.Retrieve("StandardCreationdate", " ")
55720   .StandardCreationdate = tstr
55730   tstr = hOpt.Retrieve("StandardDateformat", "YYYYMMDDHHNNSS")
55740   .StandardDateformat = tstr
55750   tstr = hOpt.Retrieve("StandardKeywords", " ")
55760   .StandardKeywords = tstr
55770   tstr = hOpt.Retrieve("StandardMailDomain", " ")
55780   .StandardMailDomain = tstr
55790   tstr = hOpt.Retrieve("StandardModifydate", " ")
55800   .StandardModifydate = tstr
55810   tstr = hOpt.Retrieve("StandardSaveformat", "pdf")
55820   .StandardSaveformat = tstr
55830   tstr = hOpt.Retrieve("StandardSubject", " ")
55840   .StandardSubject = tstr
55850   tstr = hOpt.Retrieve("StandardTitle", " ")
55860   .StandardTitle = tstr
55870   tstr = hOpt.Retrieve("StartStandardProgram")
55880   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55890     .StartStandardProgram = CLng(tstr)
55900    Else
55910     .StartStandardProgram = 1
55920   End If
55930   tstr = hOpt.Retrieve("TIFFColorscount")
55940   If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
55950     .TIFFColorscount = CLng(tstr)
55960    Else
55970     .TIFFColorscount = 0
55980   End If
55990   tstr = hOpt.Retrieve("Toolbars")
56000   If CLng(tstr) >= 0 Then
56010     .Toolbars = CLng(tstr)
56020    Else
56030     .Toolbars = 1
56040   End If
56050   tstr = hOpt.Retrieve("UseAutosave")
56060   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
56070     .UseAutosave = CLng(tstr)
56080    Else
56090     .UseAutosave = 0
56100   End If
56110   tstr = hOpt.Retrieve("UseAutosaveDirectory")
56120   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
56130     .UseAutosaveDirectory = CLng(tstr)
56140    Else
56150     .UseAutosaveDirectory = 1
56160   End If
56170   tstr = hOpt.Retrieve("UseCreationDateNow")
56180   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
56190     .UseCreationDateNow = CLng(tstr)
56200    Else
56210     .UseCreationDateNow = 0
56220   End If
56230   tstr = hOpt.Retrieve("UseStandardAuthor")
56240   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
56250     .UseStandardAuthor = CLng(tstr)
56260    Else
56270     .UseStandardAuthor = 0
56280   End If
56290  End With
56300  Set ini = Nothing
56310  ReadOptions = myOptions
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

Public Sub SaveOptions(sOptions As tOptions)
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
50090   ini.SaveKey CStr(.AdditionalGhostscriptSearchpath), "AdditionalGhostscriptSearchpath"
50100   ini.SaveKey CStr(Abs(.AddWindowsFontpath)), "AddWindowsFontpath"
50110   ini.SaveKey CStr(.AutosaveDirectory), "AutosaveDirectory"
50120   ini.SaveKey CStr(.AutosaveFilename), "AutosaveFilename"
50130   ini.SaveKey CStr(.AutosaveFormat), "AutosaveFormat"
50140   ini.SaveKey CStr(.BitmapResolution), "BitmapResolution"
50150   ini.SaveKey CStr(.BMPColorscount), "BMPColorscount"
50160   ini.SaveKey CStr(.DeviceHeightPoints), "DeviceHeightPoints"
50170   ini.SaveKey CStr(.DeviceWidthPoints), "DeviceWidthPoints"
50180   ini.SaveKey CStr(.DirectoryGhostscriptBinaries), "DirectoryGhostscriptBinaries"
50190   ini.SaveKey CStr(.DirectoryGhostscriptFonts), "DirectoryGhostscriptFonts"
50200   ini.SaveKey CStr(.DirectoryGhostscriptLibraries), "DirectoryGhostscriptLibraries"
50210   ini.SaveKey CStr(.DirectoryGhostscriptResource), "DirectoryGhostscriptResource"
50220   ini.SaveKey CStr(Abs(.DontUseDocumentSettings)), "DontUseDocumentSettings"
50230   ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
50240   ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
50250   ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
50260   ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50270   ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50280   ini.SaveKey CStr(.Language), "Language"
50290   ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50300   ini.SaveKey CStr(Abs(.Logging)), "Logging"
50310   ini.SaveKey CStr(.LogLines), "LogLines"
50320   ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50330   ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
50340   ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50350   ini.SaveKey CStr(.OptionsDesign), "OptionsDesign"
50360   ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50370   ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50380   ini.SaveKey CStr(.Papersize), "Papersize"
50390   ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50400   ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50410   ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50420   ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50430   ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50440   ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50450   ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50460   ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50470   ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50480   ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50490   ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50500   ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50510   ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50520   ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50530   ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50540   ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50550   ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50560   ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50570   ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50580   ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50590   ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50600   ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50610   ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50620   ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50630   ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50640   ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50650   ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50660   ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50670   ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50680   ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50690   ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50700   ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50710   ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50720   ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50730   ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50740   ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50750   ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50760   ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50770   ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
50780   ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
50790   ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
50800   ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
50810   ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
50820   ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
50830   ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
50840   ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
50850   ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
50860   ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
50870   ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
50880   ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
50890   ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
50900   ini.SaveKey CStr(.ProgramFont), "ProgramFont"
50910   ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
50920   ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
50930   ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
50940   ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
50950   ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
50960   ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
50970   ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
50980   ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
50990   ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
51000   ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
51010   ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
51020   ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
51030   ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51040   ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51050   ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51060   ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51070   ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51080   ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51090   ini.SaveKey CStr(.StampFontname), "StampFontname"
51100   ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51110   ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51120   ini.SaveKey CStr(.StampString), "StampString"
51130   ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51140   ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51150   ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51160   ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51170   ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51180   ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51190   ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51200   ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51210   ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51220   ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51230   ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51240   ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51250   ini.SaveKey CStr(.Toolbars), "Toolbars"
51260   ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51270   ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51280   ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51290   ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51300  End With
51310  Set ini = Nothing
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

