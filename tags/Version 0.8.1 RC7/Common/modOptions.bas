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
50300   .DontUseDocumentSettings = "0"
50310   .EPSLanguageLevel = "2"
50320   .FilenameSubstitutions = "Microsoft Word - \.doc"
50330   .FilenameSubstitutionsOnlyInTitle = "1"
50340   .JPEGColorscount = "0"
50350   .JPEGQuality = "75"
50360   .Language = "english"
50370   .LastSaveDirectory = GetMyFiles
50380   .Logging = "0"
50390   .LogLines = "100"
50400   .NoConfirmMessageSwitchingDefaultprinter = "0"
50410   .NoProcessingAtStartup = "0"
50420   .OnePagePerFile = "0"
50430   .OptionsEnabled = "1"
50440   .OptionsVisible = "1"
50450   .PCXColorscount = "0"
50460   .PDFAllowAssembly = "0"
50470   .PDFAllowDegradedPrinting = "0"
50480   .PDFAllowFillIn = "0"
50490   .PDFAllowScreenReaders = "0"
50500   .PDFColorsCMYKToRGB = "1"
50510   .PDFColorsColorModel = "1"
50520   .PDFColorsPreserveHalftone = "0"
50530   .PDFColorsPreserveOverprint = "1"
50540   .PDFColorsPreserveTransfer = "1"
50550   .PDFCompressionColorCompression = "1"
50560   .PDFCompressionColorCompressionChoice = "0"
50570   .PDFCompressionColorResample = "0"
50580   .PDFCompressionColorResampleChoice = "0"
50590   .PDFCompressionColorResolution = "300"
50600   .PDFCompressionGreyCompression = "1"
50610   .PDFCompressionGreyCompressionChoice = "0"
50620   .PDFCompressionGreyResample = "0"
50630   .PDFCompressionGreyResampleChoice = "0"
50640   .PDFCompressionGreyResolution = "300"
50650   .PDFCompressionMonoCompression = "1"
50660   .PDFCompressionMonoCompressionChoice = "0"
50670   .PDFCompressionMonoResample = "0"
50680   .PDFCompressionMonoResampleChoice = "0"
50690   .PDFCompressionMonoResolution = "1200"
50700   .PDFCompressionTextCompression = "1"
50710   .PDFDisallowCopy = "1"
50720   .PDFDisallowModifyAnnotations = "0"
50730   .PDFDisallowModifyContents = "0"
50740   .PDFDisallowPrinting = "0"
50750   .PDFEncryptor = "0"
50760   .PDFFontsEmbedAll = "1"
50770   .PDFFontsSubSetFonts = "1"
50780   .PDFFontsSubSetFontsPercent = "100"
50790   .PDFGeneralASCII85 = "0"
50800   .PDFGeneralAutorotate = "2"
50810   .PDFGeneralCompatibility = "1"
50820   .PDFGeneralOverprint = "0"
50830   .PDFGeneralResolution = "600"
50840   .PDFHighEncryption = "0"
50850   .PDFLowEncryption = "1"
50860   .PDFOptimize = "0"
50870   .PDFOwnerPass = "0"
50880   .PDFOwnerPasswordString = " "
50890   .PDFUserPass = "0"
50900   .PDFUserPasswordString = " "
50910   .PDFUseSecurity = "0"
50920   .PNGColorscount = "0"
50930   .PrinterStop = "0"
50940   .PrinterTemppath = GetTempPath
50950   .ProcessPriority = "1"
50960   .ProgramFont = "MS Sans Serif"
50970   .ProgramFontCharset = "0"
50980   .ProgramFontSize = "8"
50990   .PSLanguageLevel = "2"
51000   .RemoveAllKnownFileExtensions = "1"
51010   .RemoveSpaces = "1"
51020   .RunProgramAfterSaving = "0"
51030   .RunProgramAfterSavingProgramname = " "
51040   .RunProgramAfterSavingProgramParameters = " "
51050   .RunProgramAfterSavingWaitUntilReady = "1"
51060   .RunProgramAfterSavingWindowstyle = "1"
51070   .RunProgramBeforeSaving = "0"
51080   .RunProgramBeforeSavingProgramname = " "
51090   .RunProgramBeforeSavingProgramParameters = " "
51100   .RunProgramBeforeSavingWindowstyle = "1"
51110   .SaveFilename = "<Title>"
51120   .SendMailMethod = "0"
51130   .ShowAnimation = "1"
51140   .StampFontColor = "#FF0000"
51150   .StampFontname = "Arial"
51160   .StampFontsize = "48"
51170   .StampOutlineFontthickness = "0"
51180   .StampString = " "
51190   .StampUseOutlineFont = "1"
51200   .StandardAuthor = " "
51210   .StandardCreationdate = " "
51220   .StandardDateformat = "YYYYMMDDHHNNSS"
51230   .StandardKeywords = " "
51240   .StandardMailDomain = " "
51250   .StandardModifydate = " "
51260   .StandardSaveformat = "pdf"
51270   .StandardSubject = " "
51280   .StandardTitle = " "
51290   .StartStandardProgram = "1"
51300   .TIFFColorscount = "0"
51310   .Toolbars = "1"
51320   .UseAutosave = "0"
51330   .UseAutosaveDirectory = "1"
51340   .UseCreationDateNow = "0"
51350   .UseStandardAuthor = "0"
51360  End With
51370  StandardOptions = myOptions
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
50410   tstr = hOpt.Retrieve("DirectoryGhostscriptBinaries", App.Path)
50420   If DirExists(tstr) = True Then
50430     .DirectoryGhostscriptBinaries = CompletePath(tstr)
50440    Else
50450     .DirectoryGhostscriptBinaries = ""
50460   End If
50470   tstr = hOpt.Retrieve("DirectoryGhostscriptFonts", App.Path & "\fonts")
50480   If DirExists(tstr) = True Then
50490     .DirectoryGhostscriptFonts = CompletePath(tstr)
50500    Else
50510     .DirectoryGhostscriptFonts = ""
50520   End If
50530   tstr = hOpt.Retrieve("DirectoryGhostscriptLibraries", App.Path & "\lib")
50540   If DirExists(tstr) = True Then
50550     .DirectoryGhostscriptLibraries = CompletePath(tstr)
50560    Else
50570     .DirectoryGhostscriptLibraries = ""
50580   End If
50590   tstr = hOpt.Retrieve("DirectoryGhostscriptResource", " ")
50600   .DirectoryGhostscriptResource = tstr
50610   tstr = hOpt.Retrieve("DontUseDocumentSettings")
50620   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
50630     .DontUseDocumentSettings = CLng(tstr)
50640    Else
50650     .DontUseDocumentSettings = 0
50660   End If
50670   tstr = hOpt.Retrieve("EPSLanguageLevel")
50680   If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
50690     .EPSLanguageLevel = CLng(tstr)
50700    Else
50710     .EPSLanguageLevel = 2
50720   End If
50730   tstr = hOpt.Retrieve("FilenameSubstitutions", "Microsoft Word - \.doc")
50740   .FilenameSubstitutions = tstr
50750   tstr = hOpt.Retrieve("FilenameSubstitutionsOnlyInTitle")
50760   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
50770     .FilenameSubstitutionsOnlyInTitle = CLng(tstr)
50780    Else
50790     .FilenameSubstitutionsOnlyInTitle = 1
50800   End If
50810   tstr = hOpt.Retrieve("JPEGColorscount")
50820   If CLng(tstr) >= 0 And CLng(tstr) <= 1 Then
50830     .JPEGColorscount = CLng(tstr)
50840    Else
50850     .JPEGColorscount = 0
50860   End If
50870   tstr = hOpt.Retrieve("JPEGQuality")
50880   If CLng(tstr) >= 0 And CLng(tstr) <= 100 Then
50890     .JPEGQuality = CLng(tstr)
50900    Else
50910     .JPEGQuality = 75
50920   End If
50930   tstr = hOpt.Retrieve("Language", "english")
50940   .Language = tstr
50950   tstr = hOpt.Retrieve("LastSaveDirectory", GetMyFiles)
50960   If DirExists(tstr) = True Then
50970     .LastSaveDirectory = CompletePath(tstr)
50980    Else
50990     .LastSaveDirectory = GetMyFiles
51000   End If
51010   tstr = hOpt.Retrieve("Logging")
51020   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51030     .Logging = CLng(tstr)
51040    Else
51050     .Logging = 0
51060   End If
51070   tstr = hOpt.Retrieve("LogLines")
51080   If CLng(tstr) >= 100 And CLng(tstr) <= 1000 Then
51090     .LogLines = CLng(tstr)
51100    Else
51110     .LogLines = 100
51120   End If
51130   tstr = hOpt.Retrieve("NoConfirmMessageSwitchingDefaultprinter")
51140   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51150     .NoConfirmMessageSwitchingDefaultprinter = CLng(tstr)
51160    Else
51170     .NoConfirmMessageSwitchingDefaultprinter = 0
51180   End If
51190   tstr = hOpt.Retrieve("NoProcessingAtStartup")
51200   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51210     .NoProcessingAtStartup = CLng(tstr)
51220    Else
51230     .NoProcessingAtStartup = 0
51240   End If
51250   tstr = hOpt.Retrieve("OnePagePerFile")
51260   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51270     .OnePagePerFile = CLng(tstr)
51280    Else
51290     .OnePagePerFile = 0
51300   End If
51310   tstr = hOpt.Retrieve("OptionsEnabled")
51320   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51330     .OptionsEnabled = CLng(tstr)
51340    Else
51350     .OptionsEnabled = 1
51360   End If
51370   tstr = hOpt.Retrieve("OptionsVisible")
51380   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51390     .OptionsVisible = CLng(tstr)
51400    Else
51410     .OptionsVisible = 1
51420   End If
51430   tstr = hOpt.Retrieve("PCXColorscount")
51440   If CLng(tstr) >= 0 And CLng(tstr) <= 5 Then
51450     .PCXColorscount = CLng(tstr)
51460    Else
51470     .PCXColorscount = 0
51480   End If
51490   tstr = hOpt.Retrieve("PDFAllowAssembly")
51500   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51510     .PDFAllowAssembly = CLng(tstr)
51520    Else
51530     .PDFAllowAssembly = 0
51540   End If
51550   tstr = hOpt.Retrieve("PDFAllowDegradedPrinting")
51560   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51570     .PDFAllowDegradedPrinting = CLng(tstr)
51580    Else
51590     .PDFAllowDegradedPrinting = 0
51600   End If
51610   tstr = hOpt.Retrieve("PDFAllowFillIn")
51620   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51630     .PDFAllowFillIn = CLng(tstr)
51640    Else
51650     .PDFAllowFillIn = 0
51660   End If
51670   tstr = hOpt.Retrieve("PDFAllowScreenReaders")
51680   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51690     .PDFAllowScreenReaders = CLng(tstr)
51700    Else
51710     .PDFAllowScreenReaders = 0
51720   End If
51730   tstr = hOpt.Retrieve("PDFColorsCMYKToRGB")
51740   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51750     .PDFColorsCMYKToRGB = CLng(tstr)
51760    Else
51770     .PDFColorsCMYKToRGB = 1
51780   End If
51790   tstr = hOpt.Retrieve("PDFColorsColorModel")
51800   If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
51810     .PDFColorsColorModel = CLng(tstr)
51820    Else
51830     .PDFColorsColorModel = 1
51840   End If
51850   tstr = hOpt.Retrieve("PDFColorsPreserveHalftone")
51860   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51870     .PDFColorsPreserveHalftone = CLng(tstr)
51880    Else
51890     .PDFColorsPreserveHalftone = 0
51900   End If
51910   tstr = hOpt.Retrieve("PDFColorsPreserveOverprint")
51920   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51930     .PDFColorsPreserveOverprint = CLng(tstr)
51940    Else
51950     .PDFColorsPreserveOverprint = 1
51960   End If
51970   tstr = hOpt.Retrieve("PDFColorsPreserveTransfer")
51980   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
51990     .PDFColorsPreserveTransfer = CLng(tstr)
52000    Else
52010     .PDFColorsPreserveTransfer = 1
52020   End If
52030   tstr = hOpt.Retrieve("PDFCompressionColorCompression")
52040   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52050     .PDFCompressionColorCompression = CLng(tstr)
52060    Else
52070     .PDFCompressionColorCompression = 1
52080   End If
52090   tstr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
52100   If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
52110     .PDFCompressionColorCompressionChoice = CLng(tstr)
52120    Else
52130     .PDFCompressionColorCompressionChoice = 0
52140   End If
52150   tstr = hOpt.Retrieve("PDFCompressionColorResample")
52160   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52170     .PDFCompressionColorResample = CLng(tstr)
52180    Else
52190     .PDFCompressionColorResample = 0
52200   End If
52210   tstr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
52220   If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
52230     .PDFCompressionColorResampleChoice = CLng(tstr)
52240    Else
52250     .PDFCompressionColorResampleChoice = 0
52260   End If
52270   tstr = hOpt.Retrieve("PDFCompressionColorResolution")
52280   If CLng(tstr) >= 0 Then
52290     .PDFCompressionColorResolution = CLng(tstr)
52300    Else
52310     .PDFCompressionColorResolution = 300
52320   End If
52330   tstr = hOpt.Retrieve("PDFCompressionGreyCompression")
52340   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52350     .PDFCompressionGreyCompression = CLng(tstr)
52360    Else
52370     .PDFCompressionGreyCompression = 1
52380   End If
52390   tstr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
52400   If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
52410     .PDFCompressionGreyCompressionChoice = CLng(tstr)
52420    Else
52430     .PDFCompressionGreyCompressionChoice = 0
52440   End If
52450   tstr = hOpt.Retrieve("PDFCompressionGreyResample")
52460   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52470     .PDFCompressionGreyResample = CLng(tstr)
52480    Else
52490     .PDFCompressionGreyResample = 0
52500   End If
52510   tstr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
52520   If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
52530     .PDFCompressionGreyResampleChoice = CLng(tstr)
52540    Else
52550     .PDFCompressionGreyResampleChoice = 0
52560   End If
52570   tstr = hOpt.Retrieve("PDFCompressionGreyResolution")
52580   If CLng(tstr) >= 0 Then
52590     .PDFCompressionGreyResolution = CLng(tstr)
52600    Else
52610     .PDFCompressionGreyResolution = 300
52620   End If
52630   tstr = hOpt.Retrieve("PDFCompressionMonoCompression")
52640   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52650     .PDFCompressionMonoCompression = CLng(tstr)
52660    Else
52670     .PDFCompressionMonoCompression = 1
52680   End If
52690   tstr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
52700   If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
52710     .PDFCompressionMonoCompressionChoice = CLng(tstr)
52720    Else
52730     .PDFCompressionMonoCompressionChoice = 0
52740   End If
52750   tstr = hOpt.Retrieve("PDFCompressionMonoResample")
52760   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52770     .PDFCompressionMonoResample = CLng(tstr)
52780    Else
52790     .PDFCompressionMonoResample = 0
52800   End If
52810   tstr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
52820   If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
52830     .PDFCompressionMonoResampleChoice = CLng(tstr)
52840    Else
52850     .PDFCompressionMonoResampleChoice = 0
52860   End If
52870   tstr = hOpt.Retrieve("PDFCompressionMonoResolution")
52880   If CLng(tstr) >= 0 Then
52890     .PDFCompressionMonoResolution = CLng(tstr)
52900    Else
52910     .PDFCompressionMonoResolution = 1200
52920   End If
52930   tstr = hOpt.Retrieve("PDFCompressionTextCompression")
52940   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
52950     .PDFCompressionTextCompression = CLng(tstr)
52960    Else
52970     .PDFCompressionTextCompression = 1
52980   End If
52990   tstr = hOpt.Retrieve("PDFDisallowCopy")
53000   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53010     .PDFDisallowCopy = CLng(tstr)
53020    Else
53030     .PDFDisallowCopy = 1
53040   End If
53050   tstr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
53060   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53070     .PDFDisallowModifyAnnotations = CLng(tstr)
53080    Else
53090     .PDFDisallowModifyAnnotations = 0
53100   End If
53110   tstr = hOpt.Retrieve("PDFDisallowModifyContents")
53120   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53130     .PDFDisallowModifyContents = CLng(tstr)
53140    Else
53150     .PDFDisallowModifyContents = 0
53160   End If
53170   tstr = hOpt.Retrieve("PDFDisallowPrinting")
53180   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53190     .PDFDisallowPrinting = CLng(tstr)
53200    Else
53210     .PDFDisallowPrinting = 0
53220   End If
53230   tstr = hOpt.Retrieve("PDFEncryptor")
53240   If CLng(tstr) >= 0 And CLng(tstr) <= 1 Then
53250     .PDFEncryptor = CLng(tstr)
53260    Else
53270     .PDFEncryptor = 0
53280   End If
53290   tstr = hOpt.Retrieve("PDFFontsEmbedAll")
53300   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53310     .PDFFontsEmbedAll = CLng(tstr)
53320    Else
53330     .PDFFontsEmbedAll = 1
53340   End If
53350   tstr = hOpt.Retrieve("PDFFontsSubSetFonts")
53360   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53370     .PDFFontsSubSetFonts = CLng(tstr)
53380    Else
53390     .PDFFontsSubSetFonts = 1
53400   End If
53410   tstr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
53420   If CLng(tstr) >= 0 Then
53430     .PDFFontsSubSetFontsPercent = CLng(tstr)
53440    Else
53450     .PDFFontsSubSetFontsPercent = 100
53460   End If
53470   tstr = hOpt.Retrieve("PDFGeneralASCII85")
53480   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53490     .PDFGeneralASCII85 = CLng(tstr)
53500    Else
53510     .PDFGeneralASCII85 = 0
53520   End If
53530   tstr = hOpt.Retrieve("PDFGeneralAutorotate")
53540   If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
53550     .PDFGeneralAutorotate = CLng(tstr)
53560    Else
53570     .PDFGeneralAutorotate = 2
53580   End If
53590   tstr = hOpt.Retrieve("PDFGeneralCompatibility")
53600   If CLng(tstr) >= 0 And CLng(tstr) <= 2 Then
53610     .PDFGeneralCompatibility = CLng(tstr)
53620    Else
53630     .PDFGeneralCompatibility = 1
53640   End If
53650   tstr = hOpt.Retrieve("PDFGeneralOverprint")
53660   If CLng(tstr) >= 0 And CLng(tstr) <= 1 Then
53670     .PDFGeneralOverprint = CLng(tstr)
53680    Else
53690     .PDFGeneralOverprint = 0
53700   End If
53710   tstr = hOpt.Retrieve("PDFGeneralResolution")
53720   If CLng(tstr) >= 0 Then
53730     .PDFGeneralResolution = CLng(tstr)
53740    Else
53750     .PDFGeneralResolution = 600
53760   End If
53770   tstr = hOpt.Retrieve("PDFHighEncryption")
53780   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53790     .PDFHighEncryption = CLng(tstr)
53800    Else
53810     .PDFHighEncryption = 0
53820   End If
53830   tstr = hOpt.Retrieve("PDFLowEncryption")
53840   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53850     .PDFLowEncryption = CLng(tstr)
53860    Else
53870     .PDFLowEncryption = 1
53880   End If
53890   tstr = hOpt.Retrieve("PDFOptimize")
53900   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53910     .PDFOptimize = CLng(tstr)
53920    Else
53930     .PDFOptimize = 0
53940   End If
53950   tstr = hOpt.Retrieve("PDFOwnerPass")
53960   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
53970     .PDFOwnerPass = CLng(tstr)
53980    Else
53990     .PDFOwnerPass = 0
54000   End If
54010   tstr = hOpt.Retrieve("PDFOwnerPasswordString", " ")
54020   .PDFOwnerPasswordString = tstr
54030   tstr = hOpt.Retrieve("PDFUserPass")
54040   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54050     .PDFUserPass = CLng(tstr)
54060    Else
54070     .PDFUserPass = 0
54080   End If
54090   tstr = hOpt.Retrieve("PDFUserPasswordString", " ")
54100   .PDFUserPasswordString = tstr
54110   tstr = hOpt.Retrieve("PDFUseSecurity")
54120   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54130     .PDFUseSecurity = CLng(tstr)
54140    Else
54150     .PDFUseSecurity = 0
54160   End If
54170   tstr = hOpt.Retrieve("PNGColorscount")
54180   If CLng(tstr) >= 0 And CLng(tstr) <= 4 Then
54190     .PNGColorscount = CLng(tstr)
54200    Else
54210     .PNGColorscount = 0
54220   End If
54230   tstr = hOpt.Retrieve("PrinterStop")
54240   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54250     .PrinterStop = CLng(tstr)
54260    Else
54270     .PrinterStop = 0
54280   End If
54290   tstr = hOpt.Retrieve("PrinterTemppath", GetTempPath)
54300   If DirExists(tstr) = True Then
54310     .PrinterTemppath = CompletePath(tstr)
54320    Else
54330     .PrinterTemppath = GetTempPath
54340   End If
54350   tstr = hOpt.Retrieve("ProcessPriority")
54360   If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
54370     .ProcessPriority = CLng(tstr)
54380    Else
54390     .ProcessPriority = 1
54400   End If
54410   tstr = hOpt.Retrieve("ProgramFont", "MS Sans Serif")
54420   .ProgramFont = tstr
54430   tstr = hOpt.Retrieve("ProgramFontCharset")
54440   If CLng(tstr) >= 0 Then
54450     .ProgramFontCharset = CLng(tstr)
54460    Else
54470     .ProgramFontCharset = 0
54480   End If
54490   tstr = hOpt.Retrieve("ProgramFontSize")
54500   If CLng(tstr) >= 1 And CLng(tstr) <= 72 Then
54510     .ProgramFontSize = CLng(tstr)
54520    Else
54530     .ProgramFontSize = 8
54540   End If
54550   tstr = hOpt.Retrieve("PSLanguageLevel")
54560   If CLng(tstr) >= 0 And CLng(tstr) <= 3 Then
54570     .PSLanguageLevel = CLng(tstr)
54580    Else
54590     .PSLanguageLevel = 2
54600   End If
54610   tstr = hOpt.Retrieve("RemoveAllKnownFileExtensions")
54620   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54630     .RemoveAllKnownFileExtensions = CLng(tstr)
54640    Else
54650     .RemoveAllKnownFileExtensions = 1
54660   End If
54670   tstr = hOpt.Retrieve("RemoveSpaces")
54680   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54690     .RemoveSpaces = CLng(tstr)
54700    Else
54710     .RemoveSpaces = 1
54720   End If
54730   tstr = hOpt.Retrieve("RunProgramAfterSaving")
54740   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54750     .RunProgramAfterSaving = CLng(tstr)
54760    Else
54770     .RunProgramAfterSaving = 0
54780   End If
54790   tstr = hOpt.Retrieve("RunProgramAfterSavingProgramname", " ")
54800   .RunProgramAfterSavingProgramname = tstr
54810   tstr = hOpt.Retrieve("RunProgramAfterSavingProgramParameters", " ")
54820   .RunProgramAfterSavingProgramParameters = tstr
54830   tstr = hOpt.Retrieve("RunProgramAfterSavingWaitUntilReady")
54840   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54850     .RunProgramAfterSavingWaitUntilReady = CLng(tstr)
54860    Else
54870     .RunProgramAfterSavingWaitUntilReady = 1
54880   End If
54890   tstr = hOpt.Retrieve("RunProgramAfterSavingWindowstyle")
54900   If CLng(tstr) >= 0 And CLng(tstr) <= 6 Then
54910     .RunProgramAfterSavingWindowstyle = CLng(tstr)
54920    Else
54930     .RunProgramAfterSavingWindowstyle = 1
54940   End If
54950   tstr = hOpt.Retrieve("RunProgramBeforeSaving")
54960   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
54970     .RunProgramBeforeSaving = CLng(tstr)
54980    Else
54990     .RunProgramBeforeSaving = 0
55000   End If
55010   tstr = hOpt.Retrieve("RunProgramBeforeSavingProgramname", " ")
55020   .RunProgramBeforeSavingProgramname = tstr
55030   tstr = hOpt.Retrieve("RunProgramBeforeSavingProgramParameters", " ")
55040   .RunProgramBeforeSavingProgramParameters = tstr
55050   tstr = hOpt.Retrieve("RunProgramBeforeSavingWindowstyle")
55060   If CLng(tstr) >= 0 And CLng(tstr) <= 6 Then
55070     .RunProgramBeforeSavingWindowstyle = CLng(tstr)
55080    Else
55090     .RunProgramBeforeSavingWindowstyle = 1
55100   End If
55110   tstr = hOpt.Retrieve("SaveFilename", "<Title>")
55120   .SaveFilename = tstr
55130   tstr = hOpt.Retrieve("SendMailMethod")
55140   If CLng(tstr) >= 0 Then
55150     .SendMailMethod = CLng(tstr)
55160    Else
55170     .SendMailMethod = 0
55180   End If
55190   tstr = hOpt.Retrieve("ShowAnimation")
55200   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55210     .ShowAnimation = CLng(tstr)
55220    Else
55230     .ShowAnimation = 1
55240   End If
55250   tstr = hOpt.Retrieve("StampFontColor", "#FF0000")
55260   .StampFontColor = tstr
55270   tstr = hOpt.Retrieve("StampFontname", "Arial")
55280   .StampFontname = tstr
55290   tstr = hOpt.Retrieve("StampFontsize")
55300   If CLng(tstr) >= 1 Then
55310     .StampFontsize = CLng(tstr)
55320    Else
55330     .StampFontsize = 48
55340   End If
55350   tstr = hOpt.Retrieve("StampOutlineFontthickness")
55360   If CLng(tstr) >= 0 Then
55370     .StampOutlineFontthickness = CLng(tstr)
55380    Else
55390     .StampOutlineFontthickness = 0
55400   End If
55410   tstr = hOpt.Retrieve("StampString", " ")
55420   .StampString = tstr
55430   tstr = hOpt.Retrieve("StampUseOutlineFont")
55440   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55450     .StampUseOutlineFont = CLng(tstr)
55460    Else
55470     .StampUseOutlineFont = 1
55480   End If
55490   tstr = hOpt.Retrieve("StandardAuthor", " ")
55500   .StandardAuthor = tstr
55510   tstr = hOpt.Retrieve("StandardCreationdate", " ")
55520   .StandardCreationdate = tstr
55530   tstr = hOpt.Retrieve("StandardDateformat", "YYYYMMDDHHNNSS")
55540   .StandardDateformat = tstr
55550   tstr = hOpt.Retrieve("StandardKeywords", " ")
55560   .StandardKeywords = tstr
55570   tstr = hOpt.Retrieve("StandardMailDomain", " ")
55580   .StandardMailDomain = tstr
55590   tstr = hOpt.Retrieve("StandardModifydate", " ")
55600   .StandardModifydate = tstr
55610   tstr = hOpt.Retrieve("StandardSaveformat", "pdf")
55620   .StandardSaveformat = tstr
55630   tstr = hOpt.Retrieve("StandardSubject", " ")
55640   .StandardSubject = tstr
55650   tstr = hOpt.Retrieve("StandardTitle", " ")
55660   .StandardTitle = tstr
55670   tstr = hOpt.Retrieve("StartStandardProgram")
55680   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55690     .StartStandardProgram = CLng(tstr)
55700    Else
55710     .StartStandardProgram = 1
55720   End If
55730   tstr = hOpt.Retrieve("TIFFColorscount")
55740   If CLng(tstr) >= 0 And CLng(tstr) <= 7 Then
55750     .TIFFColorscount = CLng(tstr)
55760    Else
55770     .TIFFColorscount = 0
55780   End If
55790   tstr = hOpt.Retrieve("Toolbars")
55800   If CLng(tstr) >= 0 Then
55810     .Toolbars = CLng(tstr)
55820    Else
55830     .Toolbars = 1
55840   End If
55850   tstr = hOpt.Retrieve("UseAutosave")
55860   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55870     .UseAutosave = CLng(tstr)
55880    Else
55890     .UseAutosave = 0
55900   End If
55910   tstr = hOpt.Retrieve("UseAutosaveDirectory")
55920   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55930     .UseAutosaveDirectory = CLng(tstr)
55940    Else
55950     .UseAutosaveDirectory = 1
55960   End If
55970   tstr = hOpt.Retrieve("UseCreationDateNow")
55980   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
55990     .UseCreationDateNow = CLng(tstr)
56000    Else
56010     .UseCreationDateNow = 0
56020   End If
56030   tstr = hOpt.Retrieve("UseStandardAuthor")
56040   If CLng(tstr) = 0 Or CLng(tstr) = 1 Then
56050     .UseStandardAuthor = CLng(tstr)
56060    Else
56070     .UseStandardAuthor = 0
56080   End If
56090  End With
56100  Set ini = Nothing
56110  ReadOptions = myOptions
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
50200   ini.SaveKey CStr(Abs(.DontUseDocumentSettings)), "DontUseDocumentSettings"
50210   ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
50220   ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
50230   ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
50240   ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50250   ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50260   ini.SaveKey CStr(.Language), "Language"
50270   ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50280   ini.SaveKey CStr(Abs(.Logging)), "Logging"
50290   ini.SaveKey CStr(.LogLines), "LogLines"
50300   ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50310   ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
50320   ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50330   ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50340   ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50350   ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50360   ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50370   ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50380   ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50390   ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50400   ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50410   ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50420   ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50430   ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50440   ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50450   ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50460   ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50470   ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50480   ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50490   ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50500   ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50510   ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50520   ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50530   ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50540   ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50550   ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50560   ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50570   ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50580   ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50590   ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50600   ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50610   ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50620   ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50630   ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50640   ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50650   ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50660   ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50670   ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50680   ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50690   ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50700   ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50710   ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50720   ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50730   ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
50740   ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
50750   ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
50760   ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
50770   ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
50780   ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
50790   ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
50800   ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
50810   ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
50820   ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
50830   ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
50840   ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
50850   ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
50860   ini.SaveKey CStr(.ProgramFont), "ProgramFont"
50870   ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
50880   ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
50890   ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
50900   ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
50910   ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
50920   ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
50930   ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
50940   ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
50950   ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
50960   ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
50970   ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
50980   ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
50990   ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
51000   ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
51010   ini.SaveKey CStr(.SaveFilename), "SaveFilename"
51020   ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
51030   ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
51040   ini.SaveKey CStr(.StampFontColor), "StampFontColor"
51050   ini.SaveKey CStr(.StampFontname), "StampFontname"
51060   ini.SaveKey CStr(.StampFontsize), "StampFontsize"
51070   ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
51080   ini.SaveKey CStr(.StampString), "StampString"
51090   ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
51100   ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51110   ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51120   ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51130   ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51140   ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
51150   ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51160   ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
51170   ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51180   ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51190   ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51200   ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51210   ini.SaveKey CStr(.Toolbars), "Toolbars"
51220   ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51230   ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51240   ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51250   ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51260  End With
51270  Set ini = Nothing
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

