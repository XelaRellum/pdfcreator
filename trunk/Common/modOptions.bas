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
 DirectoryGhostscriptBinaries As String
 DirectoryGhostscriptFonts As String
 DirectoryGhostscriptLibraries As String
 DirectoryGhostscriptResource As String
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
 OptionsEnabled As Long
 OptionsVisible As Long
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
 StandardModifydate As String
 StandardSaveformat As String
 StandardSubject As String
 StandardTitle As String
 StartStandardProgram As Long
 TIFFColorscount As Long
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
50100   Set reg = New clsRegistry
50110   reg.hkey = HKEY_LOCAL_MACHINE
50120   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50130   .DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50140   Set reg = Nothing
50150   Set reg = New clsRegistry
50160   reg.hkey = HKEY_LOCAL_MACHINE
50170   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50180   .DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50190   Set reg = Nothing
50200   Set reg = New clsRegistry
50210   reg.hkey = HKEY_LOCAL_MACHINE
50220   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50230   .DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50240   Set reg = Nothing
50250   Set reg = New clsRegistry
50260   reg.hkey = HKEY_LOCAL_MACHINE
50270   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50280   .DirectoryGhostscriptResource = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryResource"))
50290   Set reg = Nothing
50300   .EPSLanguageLevel = "2"
50310   .FilenameSubstitutions = "Microsoft Word - \.doc"
50320   .FilenameSubstitutionsOnlyInTitle = "1"
50330   .JPEGColorscount = "0"
50340   .JPEGQuality = "75"
50350   .Language = "english"
50360   .LastSaveDirectory = GetMyFiles
50370   .Logging = "0"
50380   .LogLines = "100"
50390   .NoConfirmMessageSwitchingDefaultprinter = "0"
50400   .NoProcessingAtStartup = "0"
50410   .OnePagePerFile = "0"
50420   .OptionsEnabled = "1"
50430   .OptionsVisible = "1"
50440   .PCXColorscount = "0"
50450   .PDFAllowAssembly = "0"
50460   .PDFAllowDegradedPrinting = "0"
50470   .PDFAllowFillIn = "0"
50480   .PDFAllowScreenReaders = "0"
50490   .PDFColorsCMYKToRGB = "1"
50500   .PDFColorsColorModel = "1"
50510   .PDFColorsPreserveHalftone = "0"
50520   .PDFColorsPreserveOverprint = "1"
50530   .PDFColorsPreserveTransfer = "1"
50540   .PDFCompressionColorCompression = "1"
50550   .PDFCompressionColorCompressionChoice = "0"
50560   .PDFCompressionColorResample = "0"
50570   .PDFCompressionColorResampleChoice = "0"
50580   .PDFCompressionColorResolution = "300"
50590   .PDFCompressionGreyCompression = "1"
50600   .PDFCompressionGreyCompressionChoice = "0"
50610   .PDFCompressionGreyResample = "0"
50620   .PDFCompressionGreyResampleChoice = "0"
50630   .PDFCompressionGreyResolution = "300"
50640   .PDFCompressionMonoCompression = "1"
50650   .PDFCompressionMonoCompressionChoice = "0"
50660   .PDFCompressionMonoResample = "0"
50670   .PDFCompressionMonoResampleChoice = "0"
50680   .PDFCompressionMonoResolution = "1200"
50690   .PDFCompressionTextCompression = "1"
50700   .PDFDisallowCopy = "1"
50710   .PDFDisallowModifyAnnotations = "0"
50720   .PDFDisallowModifyContents = "0"
50730   .PDFDisallowPrinting = "0"
50740   .PDFEncryptor = "0"
50750   .PDFFontsEmbedAll = "1"
50760   .PDFFontsSubSetFonts = "1"
50770   .PDFFontsSubSetFontsPercent = "100"
50780   .PDFGeneralASCII85 = "0"
50790   .PDFGeneralAutorotate = "2"
50800   .PDFGeneralCompatibility = "1"
50810   .PDFGeneralOverprint = "0"
50820   .PDFGeneralResolution = "600"
50830   .PDFHighEncryption = "0"
50840   .PDFLowEncryption = "1"
50850   .PDFOptimize = "0"
50860   .PDFOwnerPass = "0"
50870   .PDFOwnerPasswordString = " "
50880   .PDFUserPass = "0"
50890   .PDFUserPasswordString = " "
50900   .PDFUseSecurity = "0"
50910   .PNGColorscount = "0"
50920   .PrinterStop = "0"
50930   .PrinterTemppath = GetTempPath
50940   .ProcessPriority = "1"
50950   .ProgramFont = "MS Sans Serif"
50960   .ProgramFontCharset = "0"
50970   .ProgramFontSize = "8"
50980   .PSLanguageLevel = "2"
50990   .RemoveAllKnownFileExtensions = "1"
51000   .RemoveSpaces = "1"
51010   .RunProgramAfterSaving = "0"
51020   .RunProgramAfterSavingProgramname = " "
51030   .RunProgramAfterSavingProgramParameters = " "
51040   .RunProgramAfterSavingWaitUntilReady = "1"
51050   .RunProgramAfterSavingWindowstyle = "1"
51060   .SaveFilename = "<Title>"
51070   .SendMailMethod = "0"
51080   .ShowAnimation = "1"
51090   .StampFontColor = "#FF0000"
51100   .StampFontname = "Arial"
51110   .StampFontsize = "48"
51120   .StampOutlineFontthickness = "0"
51130   .StampString = " "
51140   .StampUseOutlineFont = "1"
51150   .StandardAuthor = " "
51160   .StandardCreationdate = " "
51170   .StandardDateformat = "YYYYMMDDHHNNSS"
51180   .StandardKeywords = " "
51190   .StandardModifydate = " "
51200   .StandardSaveformat = "pdf"
51210   .StandardSubject = " "
51220   .StandardTitle = " "
51230   .StartStandardProgram = "1"
51240   .TIFFColorscount = "0"
51250   .UseAutosave = "0"
51260   .UseAutosaveDirectory = "1"
51270   .UseCreationDateNow = "0"
51280   .UseStandardAuthor = "0"
51290  End With
51300  StandardOptions = myOptions
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
50010  Dim ini As clsINI, myOptions As tOptions, tStr As String, hOpt As New clsHash
50020  Set ini = New clsINI
50030  ini.Filename = PDFCreatorINIFile
50040  ini.Section = "Options"
50050  If ini.CheckIniFile = False Then
50060   ReadOptions = StandardOptions
50070   Exit Function
50080  End If
50090  ReadINISection PDFCreatorINIFile, "Options", hOpt
50100  With myOptions
50110   tStr = hOpt.Retrieve("AdditionalGhostscriptSearchpath", " ")
50120   .AdditionalGhostscriptSearchpath = tStr
50130   tStr = hOpt.Retrieve("AddWindowsFontpath")
50140   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
50150     .AddWindowsFontpath = CLng(tStr)
50160    Else
50170     .AddWindowsFontpath = 1
50180   End If
50190   tStr = hOpt.Retrieve("AutosaveDirectory", GetMyFiles)
50200   .AutosaveDirectory = CompletePath(tStr)
50210   tStr = hOpt.Retrieve("AutosaveFilename", "<DateTime>")
50220   .AutosaveFilename = tStr
50230   tStr = hOpt.Retrieve("AutosaveFormat")
50240   If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
50250     .AutosaveFormat = CLng(tStr)
50260    Else
50270     .AutosaveFormat = 0
50280   End If
50290   tStr = hOpt.Retrieve("BitmapResolution")
50300   If CLng(tStr) >= 1 Then
50310     .BitmapResolution = CLng(tStr)
50320    Else
50330     .BitmapResolution = 150
50340   End If
50350   tStr = hOpt.Retrieve("BMPColorscount")
50360   If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
50370     .BMPColorscount = CLng(tStr)
50380    Else
50390     .BMPColorscount = 1
50400   End If
50410   tStr = hOpt.Retrieve("DirectoryGhostscriptBinaries", App.Path)
50420   If DirExists(tStr) = True Then
50430     .DirectoryGhostscriptBinaries = CompletePath(tStr)
50440    Else
50450     .DirectoryGhostscriptBinaries = ""
50460   End If
50470   tStr = hOpt.Retrieve("DirectoryGhostscriptFonts", App.Path & "\fonts")
50480   If DirExists(tStr) = True Then
50490     .DirectoryGhostscriptFonts = CompletePath(tStr)
50500    Else
50510     .DirectoryGhostscriptFonts = ""
50520   End If
50530   tStr = hOpt.Retrieve("DirectoryGhostscriptLibraries", App.Path & "\lib")
50540   If DirExists(tStr) = True Then
50550     .DirectoryGhostscriptLibraries = CompletePath(tStr)
50560    Else
50570     .DirectoryGhostscriptLibraries = ""
50580   End If
50590   tStr = hOpt.Retrieve("DirectoryGhostscriptResource", " ")
50600   .DirectoryGhostscriptResource = tStr
50610   tStr = hOpt.Retrieve("EPSLanguageLevel")
50620   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
50630     .EPSLanguageLevel = CLng(tStr)
50640    Else
50650     .EPSLanguageLevel = 2
50660   End If
50670   tStr = hOpt.Retrieve("FilenameSubstitutions", "Microsoft Word - \.doc")
50680   .FilenameSubstitutions = tStr
50690   tStr = hOpt.Retrieve("FilenameSubstitutionsOnlyInTitle")
50700   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
50710     .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
50720    Else
50730     .FilenameSubstitutionsOnlyInTitle = 1
50740   End If
50750   tStr = hOpt.Retrieve("JPEGColorscount")
50760   If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
50770     .JPEGColorscount = CLng(tStr)
50780    Else
50790     .JPEGColorscount = 0
50800   End If
50810   tStr = hOpt.Retrieve("JPEGQuality")
50820   If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
50830     .JPEGQuality = CLng(tStr)
50840    Else
50850     .JPEGQuality = 75
50860   End If
50870   tStr = hOpt.Retrieve("Language", "english")
50880   .Language = tStr
50890   tStr = hOpt.Retrieve("LastSaveDirectory", GetMyFiles)
50900   If DirExists(tStr) = True Then
50910     .LastSaveDirectory = CompletePath(tStr)
50920    Else
50930     .LastSaveDirectory = GetMyFiles
50940   End If
50950   tStr = hOpt.Retrieve("Logging")
50960   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
50970     .Logging = CLng(tStr)
50980    Else
50990     .Logging = 0
51000   End If
51010   tStr = hOpt.Retrieve("LogLines")
51020   If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
51030     .LogLines = CLng(tStr)
51040    Else
51050     .LogLines = 100
51060   End If
51070   tStr = hOpt.Retrieve("NoConfirmMessageSwitchingDefaultprinter")
51080   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51090     .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
51100    Else
51110     .NoConfirmMessageSwitchingDefaultprinter = 0
51120   End If
51130   tStr = hOpt.Retrieve("NoProcessingAtStartup")
51140   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51150     .NoProcessingAtStartup = CLng(tStr)
51160    Else
51170     .NoProcessingAtStartup = 0
51180   End If
51190   tStr = hOpt.Retrieve("OnePagePerFile")
51200   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51210     .OnePagePerFile = CLng(tStr)
51220    Else
51230     .OnePagePerFile = 0
51240   End If
51250   tStr = hOpt.Retrieve("OptionsEnabled")
51260   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51270     .OptionsEnabled = CLng(tStr)
51280    Else
51290     .OptionsEnabled = 1
51300   End If
51310   tStr = hOpt.Retrieve("OptionsVisible")
51320   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51330     .OptionsVisible = CLng(tStr)
51340    Else
51350     .OptionsVisible = 1
51360   End If
51370   tStr = hOpt.Retrieve("PCXColorscount")
51380   If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
51390     .PCXColorscount = CLng(tStr)
51400    Else
51410     .PCXColorscount = 0
51420   End If
51430   tStr = hOpt.Retrieve("PDFAllowAssembly")
51440   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51450     .PDFAllowAssembly = CLng(tStr)
51460    Else
51470     .PDFAllowAssembly = 0
51480   End If
51490   tStr = hOpt.Retrieve("PDFAllowDegradedPrinting")
51500   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51510     .PDFAllowDegradedPrinting = CLng(tStr)
51520    Else
51530     .PDFAllowDegradedPrinting = 0
51540   End If
51550   tStr = hOpt.Retrieve("PDFAllowFillIn")
51560   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51570     .PDFAllowFillIn = CLng(tStr)
51580    Else
51590     .PDFAllowFillIn = 0
51600   End If
51610   tStr = hOpt.Retrieve("PDFAllowScreenReaders")
51620   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51630     .PDFAllowScreenReaders = CLng(tStr)
51640    Else
51650     .PDFAllowScreenReaders = 0
51660   End If
51670   tStr = hOpt.Retrieve("PDFColorsCMYKToRGB")
51680   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51690     .PDFColorsCMYKToRGB = CLng(tStr)
51700    Else
51710     .PDFColorsCMYKToRGB = 1
51720   End If
51730   tStr = hOpt.Retrieve("PDFColorsColorModel")
51740   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
51750     .PDFColorsColorModel = CLng(tStr)
51760    Else
51770     .PDFColorsColorModel = 1
51780   End If
51790   tStr = hOpt.Retrieve("PDFColorsPreserveHalftone")
51800   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51810     .PDFColorsPreserveHalftone = CLng(tStr)
51820    Else
51830     .PDFColorsPreserveHalftone = 0
51840   End If
51850   tStr = hOpt.Retrieve("PDFColorsPreserveOverprint")
51860   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51870     .PDFColorsPreserveOverprint = CLng(tStr)
51880    Else
51890     .PDFColorsPreserveOverprint = 1
51900   End If
51910   tStr = hOpt.Retrieve("PDFColorsPreserveTransfer")
51920   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51930     .PDFColorsPreserveTransfer = CLng(tStr)
51940    Else
51950     .PDFColorsPreserveTransfer = 1
51960   End If
51970   tStr = hOpt.Retrieve("PDFCompressionColorCompression")
51980   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51990     .PDFCompressionColorCompression = CLng(tStr)
52000    Else
52010     .PDFCompressionColorCompression = 1
52020   End If
52030   tStr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
52040   If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
52050     .PDFCompressionColorCompressionChoice = CLng(tStr)
52060    Else
52070     .PDFCompressionColorCompressionChoice = 0
52080   End If
52090   tStr = hOpt.Retrieve("PDFCompressionColorResample")
52100   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52110     .PDFCompressionColorResample = CLng(tStr)
52120    Else
52130     .PDFCompressionColorResample = 0
52140   End If
52150   tStr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
52160   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
52170     .PDFCompressionColorResampleChoice = CLng(tStr)
52180    Else
52190     .PDFCompressionColorResampleChoice = 0
52200   End If
52210   tStr = hOpt.Retrieve("PDFCompressionColorResolution")
52220   If CLng(tStr) >= 0 Then
52230     .PDFCompressionColorResolution = CLng(tStr)
52240    Else
52250     .PDFCompressionColorResolution = 300
52260   End If
52270   tStr = hOpt.Retrieve("PDFCompressionGreyCompression")
52280   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52290     .PDFCompressionGreyCompression = CLng(tStr)
52300    Else
52310     .PDFCompressionGreyCompression = 1
52320   End If
52330   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
52340   If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
52350     .PDFCompressionGreyCompressionChoice = CLng(tStr)
52360    Else
52370     .PDFCompressionGreyCompressionChoice = 0
52380   End If
52390   tStr = hOpt.Retrieve("PDFCompressionGreyResample")
52400   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52410     .PDFCompressionGreyResample = CLng(tStr)
52420    Else
52430     .PDFCompressionGreyResample = 0
52440   End If
52450   tStr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
52460   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
52470     .PDFCompressionGreyResampleChoice = CLng(tStr)
52480    Else
52490     .PDFCompressionGreyResampleChoice = 0
52500   End If
52510   tStr = hOpt.Retrieve("PDFCompressionGreyResolution")
52520   If CLng(tStr) >= 0 Then
52530     .PDFCompressionGreyResolution = CLng(tStr)
52540    Else
52550     .PDFCompressionGreyResolution = 300
52560   End If
52570   tStr = hOpt.Retrieve("PDFCompressionMonoCompression")
52580   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52590     .PDFCompressionMonoCompression = CLng(tStr)
52600    Else
52610     .PDFCompressionMonoCompression = 1
52620   End If
52630   tStr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
52640   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
52650     .PDFCompressionMonoCompressionChoice = CLng(tStr)
52660    Else
52670     .PDFCompressionMonoCompressionChoice = 0
52680   End If
52690   tStr = hOpt.Retrieve("PDFCompressionMonoResample")
52700   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52710     .PDFCompressionMonoResample = CLng(tStr)
52720    Else
52730     .PDFCompressionMonoResample = 0
52740   End If
52750   tStr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
52760   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
52770     .PDFCompressionMonoResampleChoice = CLng(tStr)
52780    Else
52790     .PDFCompressionMonoResampleChoice = 0
52800   End If
52810   tStr = hOpt.Retrieve("PDFCompressionMonoResolution")
52820   If CLng(tStr) >= 0 Then
52830     .PDFCompressionMonoResolution = CLng(tStr)
52840    Else
52850     .PDFCompressionMonoResolution = 1200
52860   End If
52870   tStr = hOpt.Retrieve("PDFCompressionTextCompression")
52880   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52890     .PDFCompressionTextCompression = CLng(tStr)
52900    Else
52910     .PDFCompressionTextCompression = 1
52920   End If
52930   tStr = hOpt.Retrieve("PDFDisallowCopy")
52940   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52950     .PDFDisallowCopy = CLng(tStr)
52960    Else
52970     .PDFDisallowCopy = 1
52980   End If
52990   tStr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
53000   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53010     .PDFDisallowModifyAnnotations = CLng(tStr)
53020    Else
53030     .PDFDisallowModifyAnnotations = 0
53040   End If
53050   tStr = hOpt.Retrieve("PDFDisallowModifyContents")
53060   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53070     .PDFDisallowModifyContents = CLng(tStr)
53080    Else
53090     .PDFDisallowModifyContents = 0
53100   End If
53110   tStr = hOpt.Retrieve("PDFDisallowPrinting")
53120   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53130     .PDFDisallowPrinting = CLng(tStr)
53140    Else
53150     .PDFDisallowPrinting = 0
53160   End If
53170   tStr = hOpt.Retrieve("PDFEncryptor")
53180   If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
53190     .PDFEncryptor = CLng(tStr)
53200    Else
53210     .PDFEncryptor = 0
53220   End If
53230   tStr = hOpt.Retrieve("PDFFontsEmbedAll")
53240   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53250     .PDFFontsEmbedAll = CLng(tStr)
53260    Else
53270     .PDFFontsEmbedAll = 1
53280   End If
53290   tStr = hOpt.Retrieve("PDFFontsSubSetFonts")
53300   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53310     .PDFFontsSubSetFonts = CLng(tStr)
53320    Else
53330     .PDFFontsSubSetFonts = 1
53340   End If
53350   tStr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
53360   If CLng(tStr) >= 0 Then
53370     .PDFFontsSubSetFontsPercent = CLng(tStr)
53380    Else
53390     .PDFFontsSubSetFontsPercent = 100
53400   End If
53410   tStr = hOpt.Retrieve("PDFGeneralASCII85")
53420   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53430     .PDFGeneralASCII85 = CLng(tStr)
53440    Else
53450     .PDFGeneralASCII85 = 0
53460   End If
53470   tStr = hOpt.Retrieve("PDFGeneralAutorotate")
53480   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
53490     .PDFGeneralAutorotate = CLng(tStr)
53500    Else
53510     .PDFGeneralAutorotate = 2
53520   End If
53530   tStr = hOpt.Retrieve("PDFGeneralCompatibility")
53540   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
53550     .PDFGeneralCompatibility = CLng(tStr)
53560    Else
53570     .PDFGeneralCompatibility = 1
53580   End If
53590   tStr = hOpt.Retrieve("PDFGeneralOverprint")
53600   If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
53610     .PDFGeneralOverprint = CLng(tStr)
53620    Else
53630     .PDFGeneralOverprint = 0
53640   End If
53650   tStr = hOpt.Retrieve("PDFGeneralResolution")
53660   If CLng(tStr) >= 0 Then
53670     .PDFGeneralResolution = CLng(tStr)
53680    Else
53690     .PDFGeneralResolution = 600
53700   End If
53710   tStr = hOpt.Retrieve("PDFHighEncryption")
53720   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53730     .PDFHighEncryption = CLng(tStr)
53740    Else
53750     .PDFHighEncryption = 0
53760   End If
53770   tStr = hOpt.Retrieve("PDFLowEncryption")
53780   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53790     .PDFLowEncryption = CLng(tStr)
53800    Else
53810     .PDFLowEncryption = 1
53820   End If
53830   tStr = hOpt.Retrieve("PDFOptimize")
53840   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53850     .PDFOptimize = CLng(tStr)
53860    Else
53870     .PDFOptimize = 0
53880   End If
53890   tStr = hOpt.Retrieve("PDFOwnerPass")
53900   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53910     .PDFOwnerPass = CLng(tStr)
53920    Else
53930     .PDFOwnerPass = 0
53940   End If
53950   tStr = hOpt.Retrieve("PDFOwnerPasswordString", " ")
53960   .PDFOwnerPasswordString = tStr
53970   tStr = hOpt.Retrieve("PDFUserPass")
53980   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53990     .PDFUserPass = CLng(tStr)
54000    Else
54010     .PDFUserPass = 0
54020   End If
54030   tStr = hOpt.Retrieve("PDFUserPasswordString", " ")
54040   .PDFUserPasswordString = tStr
54050   tStr = hOpt.Retrieve("PDFUseSecurity")
54060   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54070     .PDFUseSecurity = CLng(tStr)
54080    Else
54090     .PDFUseSecurity = 0
54100   End If
54110   tStr = hOpt.Retrieve("PNGColorscount")
54120   If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
54130     .PNGColorscount = CLng(tStr)
54140    Else
54150     .PNGColorscount = 0
54160   End If
54170   tStr = hOpt.Retrieve("PrinterStop")
54180   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54190     .PrinterStop = CLng(tStr)
54200    Else
54210     .PrinterStop = 0
54220   End If
54230   tStr = hOpt.Retrieve("PrinterTemppath", GetTempPath)
54240   If DirExists(tStr) = True Then
54250     .PrinterTemppath = CompletePath(tStr)
54260    Else
54270     .PrinterTemppath = GetTempPath
54280   End If
54290   tStr = hOpt.Retrieve("ProcessPriority")
54300   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
54310     .ProcessPriority = CLng(tStr)
54320    Else
54330     .ProcessPriority = 1
54340   End If
54350   tStr = hOpt.Retrieve("ProgramFont", "MS Sans Serif")
54360   .ProgramFont = tStr
54370   tStr = hOpt.Retrieve("ProgramFontCharset")
54380   If CLng(tStr) >= 0 Then
54390     .ProgramFontCharset = CLng(tStr)
54400    Else
54410     .ProgramFontCharset = 0
54420   End If
54430   tStr = hOpt.Retrieve("ProgramFontSize")
54440   If CLng(tStr) >= 1 And CLng(tStr) <= 72 Then
54450     .ProgramFontSize = CLng(tStr)
54460    Else
54470     .ProgramFontSize = 8
54480   End If
54490   tStr = hOpt.Retrieve("PSLanguageLevel")
54500   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
54510     .PSLanguageLevel = CLng(tStr)
54520    Else
54530     .PSLanguageLevel = 2
54540   End If
54550   tStr = hOpt.Retrieve("RemoveAllKnownFileExtensions")
54560   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54570     .RemoveAllKnownFileExtensions = CLng(tStr)
54580    Else
54590     .RemoveAllKnownFileExtensions = 1
54600   End If
54610   tStr = hOpt.Retrieve("RemoveSpaces")
54620   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54630     .RemoveSpaces = CLng(tStr)
54640    Else
54650     .RemoveSpaces = 1
54660   End If
54670   tStr = hOpt.Retrieve("RunProgramAfterSaving")
54680   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54690     .RunProgramAfterSaving = CLng(tStr)
54700    Else
54710     .RunProgramAfterSaving = 0
54720   End If
54730   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramname", " ")
54740   .RunProgramAfterSavingProgramname = tStr
54750   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramParameters", " ")
54760   .RunProgramAfterSavingProgramParameters = tStr
54770   tStr = hOpt.Retrieve("RunProgramAfterSavingWaitUntilReady")
54780   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54790     .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
54800    Else
54810     .RunProgramAfterSavingWaitUntilReady = 1
54820   End If
54830   tStr = hOpt.Retrieve("RunProgramAfterSavingWindowstyle")
54840   If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
54850     .RunProgramAfterSavingWindowstyle = CLng(tStr)
54860    Else
54870     .RunProgramAfterSavingWindowstyle = 1
54880   End If
54890   tStr = hOpt.Retrieve("SaveFilename", "<Title>")
54900   .SaveFilename = tStr
54910   tStr = hOpt.Retrieve("SendMailMethod")
54920   If CLng(tStr) >= 0 Then
54930     .SendMailMethod = CLng(tStr)
54940    Else
54950     .SendMailMethod = 0
54960   End If
54970   tStr = hOpt.Retrieve("ShowAnimation")
54980   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54990     .ShowAnimation = CLng(tStr)
55000    Else
55010     .ShowAnimation = 1
55020   End If
55030   tStr = hOpt.Retrieve("StampFontColor", "#FF0000")
55040   .StampFontColor = tStr
55050   tStr = hOpt.Retrieve("StampFontname", "Arial")
55060   .StampFontname = tStr
55070   tStr = hOpt.Retrieve("StampFontsize")
55080   If CLng(tStr) >= 1 Then
55090     .StampFontsize = CLng(tStr)
55100    Else
55110     .StampFontsize = 48
55120   End If
55130   tStr = hOpt.Retrieve("StampOutlineFontthickness")
55140   If CLng(tStr) >= 0 Then
55150     .StampOutlineFontthickness = CLng(tStr)
55160    Else
55170     .StampOutlineFontthickness = 0
55180   End If
55190   tStr = hOpt.Retrieve("StampString", " ")
55200   .StampString = tStr
55210   tStr = hOpt.Retrieve("StampUseOutlineFont")
55220   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55230     .StampUseOutlineFont = CLng(tStr)
55240    Else
55250     .StampUseOutlineFont = 1
55260   End If
55270   tStr = hOpt.Retrieve("StandardAuthor", " ")
55280   .StandardAuthor = tStr
55290   tStr = hOpt.Retrieve("StandardCreationdate", " ")
55300   .StandardCreationdate = tStr
55310   tStr = hOpt.Retrieve("StandardDateformat", "YYYYMMDDHHNNSS")
55320   .StandardDateformat = tStr
55330   tStr = hOpt.Retrieve("StandardKeywords", " ")
55340   .StandardKeywords = tStr
55350   tStr = hOpt.Retrieve("StandardModifydate", " ")
55360   .StandardModifydate = tStr
55370   tStr = hOpt.Retrieve("StandardSaveformat", "pdf")
55380   .StandardSaveformat = tStr
55390   tStr = hOpt.Retrieve("StandardSubject", " ")
55400   .StandardSubject = tStr
55410   tStr = hOpt.Retrieve("StandardTitle", " ")
55420   .StandardTitle = tStr
55430   tStr = hOpt.Retrieve("StartStandardProgram")
55440   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55450     .StartStandardProgram = CLng(tStr)
55460    Else
55470     .StartStandardProgram = 1
55480   End If
55490   tStr = hOpt.Retrieve("TIFFColorscount")
55500   If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
55510     .TIFFColorscount = CLng(tStr)
55520    Else
55530     .TIFFColorscount = 0
55540   End If
55550   tStr = hOpt.Retrieve("UseAutosave")
55560   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55570     .UseAutosave = CLng(tStr)
55580    Else
55590     .UseAutosave = 0
55600   End If
55610   tStr = hOpt.Retrieve("UseAutosaveDirectory")
55620   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55630     .UseAutosaveDirectory = CLng(tStr)
55640    Else
55650     .UseAutosaveDirectory = 1
55660   End If
55670   tStr = hOpt.Retrieve("UseCreationDateNow")
55680   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55690     .UseCreationDateNow = CLng(tStr)
55700    Else
55710     .UseCreationDateNow = 0
55720   End If
55730   tStr = hOpt.Retrieve("UseStandardAuthor")
55740   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55750     .UseStandardAuthor = CLng(tStr)
55760    Else
55770     .UseStandardAuthor = 0
55780   End If
55790  End With
55800  Set ini = Nothing
55810  ReadOptions = myOptions
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
50160   ini.SaveKey CStr(.DirectoryGhostscriptBinaries), "DirectoryGhostscriptBinaries"
50170   ini.SaveKey CStr(.DirectoryGhostscriptFonts), "DirectoryGhostscriptFonts"
50180   ini.SaveKey CStr(.DirectoryGhostscriptLibraries), "DirectoryGhostscriptLibraries"
50190   ini.SaveKey CStr(.DirectoryGhostscriptResource), "DirectoryGhostscriptResource"
50200   ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
50210   ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
50220   ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
50230   ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50240   ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50250   ini.SaveKey CStr(.Language), "Language"
50260   ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50270   ini.SaveKey CStr(Abs(.Logging)), "Logging"
50280   ini.SaveKey CStr(.LogLines), "LogLines"
50290   ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50300   ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
50310   ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50320   ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50330   ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50340   ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50350   ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50360   ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50370   ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50380   ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50390   ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50400   ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50410   ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50420   ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50430   ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50440   ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50450   ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50460   ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50470   ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50480   ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50490   ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50500   ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50510   ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50520   ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50530   ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50540   ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50550   ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50560   ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50570   ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50580   ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50590   ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50600   ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50610   ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50620   ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50630   ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50640   ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50650   ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50660   ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50670   ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50680   ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50690   ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50700   ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50710   ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50720   ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
50730   ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
50740   ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
50750   ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
50760   ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
50770   ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
50780   ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
50790   ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
50800   ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
50810   ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
50820   ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
50830   ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
50840   ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
50850   ini.SaveKey CStr(.ProgramFont), "ProgramFont"
50860   ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
50870   ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
50880   ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
50890   ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
50900   ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
50910   ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
50920   ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
50930   ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
50940   ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
50950   ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
50960   ini.SaveKey CStr(.SaveFilename), "SaveFilename"
50970   ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
50980   ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
50990   ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51000   ini.SaveKey CStr(.StampFontname), "StampFontname"
51010   ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51020   ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51030   ini.SaveKey CStr(.StampString), "StampString"
51040   ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51050   ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51060   ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51070   ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51080   ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51090   ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51100   ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51110   ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51120   ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51130   ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51140   ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51150   ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51160   ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51170   ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51180   ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51190  End With
51200  Set ini = Nothing
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
  Frm.txtGSresource.Text = .DirectoryGhostscriptResource
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
50110  .DirectoryGhostscriptResource = Frm.txtGSresource.Text
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

