Attribute VB_Name = "modOptions"
Option Explicit

' Module automatically generated with LanguagesTool from Frank Heindörfer
' 2003
' Email: thesmilyface@users.sourceforge.net

Public Type tOptions
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
 RemoveSpaces As Long
 RunProgramAfterSaving As Long
 RunProgramAfterSavingProgramname As String
 RunProgramAfterSavingProgramParameters As String
 RunProgramAfterSavingWaitUntilReady As Long
 RunProgramAfterSavingWindowstyle As Long
 SaveFilename As String
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
50030   .AutosaveDirectory = " "
50040   .AutosaveFilename = "<DateTime>"
50050   .AutosaveFormat = "0"
50060   .BitmapResolution = "150"
50070   .BMPColorscount = "1"
50080   Set reg = New clsRegistry
50090   reg.hkey = HKEY_LOCAL_MACHINE
50100   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50110   .DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50120   Set reg = Nothing
50130   Set reg = New clsRegistry
50140   reg.hkey = HKEY_LOCAL_MACHINE
50150   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50160   .DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50170   Set reg = Nothing
50180   Set reg = New clsRegistry
50190   reg.hkey = HKEY_LOCAL_MACHINE
50200   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50210   .DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50220   Set reg = Nothing
50230   Set reg = New clsRegistry
50240   reg.hkey = HKEY_LOCAL_MACHINE
50250   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50260   .DirectoryGhostscriptResource = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryResource"))
50270   Set reg = Nothing
50280   .EPSLanguageLevel = "2"
50290   .FilenameSubstitutions = "Microsoft Word - \.doc"
50300   .FilenameSubstitutionsOnlyInTitle = "1"
50310   .JPEGColorscount = "0"
50320   .JPEGQuality = "75"
50330   .Language = "english"
50340   .LastSaveDirectory = GetMyFiles
50350   .Logging = "0"
50360   .LogLines = "100"
50370   .NoConfirmMessageSwitchingDefaultprinter = "0"
50380   .OnePagePerFile = "0"
50390   .OptionsEnabled = "1"
50400   .OptionsVisible = "1"
50410   .PCXColorscount = "0"
50420   .PDFAllowAssembly = "0"
50430   .PDFAllowDegradedPrinting = "0"
50440   .PDFAllowFillIn = "0"
50450   .PDFAllowScreenReaders = "0"
50460   .PDFColorsCMYKToRGB = "1"
50470   .PDFColorsColorModel = "1"
50480   .PDFColorsPreserveHalftone = "0"
50490   .PDFColorsPreserveOverprint = "1"
50500   .PDFColorsPreserveTransfer = "1"
50510   .PDFCompressionColorCompression = "1"
50520   .PDFCompressionColorCompressionChoice = "0"
50530   .PDFCompressionColorResample = "0"
50540   .PDFCompressionColorResampleChoice = "0"
50550   .PDFCompressionColorResolution = "300"
50560   .PDFCompressionGreyCompression = "1"
50570   .PDFCompressionGreyCompressionChoice = "0"
50580   .PDFCompressionGreyResample = "0"
50590   .PDFCompressionGreyResampleChoice = "0"
50600   .PDFCompressionGreyResolution = "300"
50610   .PDFCompressionMonoCompression = "1"
50620   .PDFCompressionMonoCompressionChoice = "0"
50630   .PDFCompressionMonoResample = "0"
50640   .PDFCompressionMonoResampleChoice = "0"
50650   .PDFCompressionMonoResolution = "1200"
50660   .PDFCompressionTextCompression = "1"
50670   .PDFDisallowCopy = "1"
50680   .PDFDisallowModifyAnnotations = "0"
50690   .PDFDisallowModifyContents = "0"
50700   .PDFDisallowPrinting = "0"
50710   .PDFEncryptor = "0"
50720   .PDFFontsEmbedAll = "1"
50730   .PDFFontsSubSetFonts = "1"
50740   .PDFFontsSubSetFontsPercent = "100"
50750   .PDFGeneralASCII85 = "0"
50760   .PDFGeneralAutorotate = "2"
50770   .PDFGeneralCompatibility = "1"
50780   .PDFGeneralOverprint = "0"
50790   .PDFGeneralResolution = "600"
50800   .PDFHighEncryption = "0"
50810   .PDFLowEncryption = "1"
50820   .PDFOwnerPass = "0"
50830   .PDFOwnerPasswordString = " "
50840   .PDFUserPass = "0"
50850   .PDFUserPasswordString = " "
50860   .PDFUseSecurity = "0"
50870   .PNGColorscount = "0"
50880   .PrinterStop = "0"
50890   .PrinterTemppath = GetTempPath
50900   .ProcessPriority = "1"
50910   .ProgramFont = "MS Sans Serif"
50920   .ProgramFontCharset = "0"
50930   .ProgramFontSize = "8"
50940   .PSLanguageLevel = "2"
50950   .RemoveSpaces = "1"
50960   .RunProgramAfterSaving = "0"
50970   .RunProgramAfterSavingProgramname = " "
50980   .RunProgramAfterSavingProgramParameters = " "
50990   .RunProgramAfterSavingWaitUntilReady = "1"
51000   .RunProgramAfterSavingWindowstyle = "1"
51010   .SaveFilename = "<Title>"
51020   .ShowAnimation = "1"
51030   .StampFontColor = "#FF0000"
51040   .StampFontname = "Arial"
51050   .StampFontsize = "48"
51060   .StampOutlineFontthickness = "0"
51070   .StampString = " "
51080   .StampUseOutlineFont = "1"
51090   .StandardAuthor = " "
51100   .StandardCreationdate = " "
51110   .StandardDateformat = "YYYYMMDDHHNNSS"
51120   .StandardKeywords = " "
51130   .StandardModifydate = " "
51140   .StandardSubject = " "
51150   .StandardTitle = " "
51160   .StartStandardProgram = "1"
51170   .TIFFColorscount = "0"
51180   .UseAutosave = "0"
51190   .UseAutosaveDirectory = "1"
51200   .UseCreationDateNow = "0"
51210   .UseStandardAuthor = "0"
51220  End With
51230  StandardOptions = myOptions
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
50110   tStr = hOpt.Retrieve("AutosaveDirectory", GetMyFiles)
50120   .AutosaveDirectory = CompletePath(tStr)
50130   tStr = hOpt.Retrieve("AutosaveFilename", "<DateTime>")
50140   .AutosaveFilename = tStr
50150   tStr = hOpt.Retrieve("AutosaveFormat")
50160   If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
50170     .AutosaveFormat = CLng(tStr)
50180    Else
50190     .AutosaveFormat = 0
50200   End If
50210   tStr = hOpt.Retrieve("BitmapResolution")
50220   If CLng(tStr) >= 1 Then
50230     .BitmapResolution = CLng(tStr)
50240    Else
50250     .BitmapResolution = 150
50260   End If
50270   tStr = hOpt.Retrieve("BMPColorscount")
50280   If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
50290     .BMPColorscount = CLng(tStr)
50300    Else
50310     .BMPColorscount = 1
50320   End If
50330   tStr = hOpt.Retrieve("DirectoryGhostscriptBinaries", App.Path)
50340   If DirExists(tStr) = True Then
50350     .DirectoryGhostscriptBinaries = CompletePath(tStr)
50360    Else
50370     .DirectoryGhostscriptBinaries = ""
50380   End If
50390   tStr = hOpt.Retrieve("DirectoryGhostscriptFonts", App.Path & "\fonts")
50400   If DirExists(tStr) = True Then
50410     .DirectoryGhostscriptFonts = CompletePath(tStr)
50420    Else
50430     .DirectoryGhostscriptFonts = ""
50440   End If
50450   tStr = hOpt.Retrieve("DirectoryGhostscriptLibraries", App.Path & "\lib")
50460   If DirExists(tStr) = True Then
50470     .DirectoryGhostscriptLibraries = CompletePath(tStr)
50480    Else
50490     .DirectoryGhostscriptLibraries = ""
50500   End If
50510   tStr = hOpt.Retrieve("DirectoryGhostscriptResource", " ")
50520   .DirectoryGhostscriptResource = tStr
50530   tStr = hOpt.Retrieve("EPSLanguageLevel")
50540   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
50550     .EPSLanguageLevel = CLng(tStr)
50560    Else
50570     .EPSLanguageLevel = 2
50580   End If
50590   tStr = hOpt.Retrieve("FilenameSubstitutions", "Microsoft Word - \.doc")
50600   .FilenameSubstitutions = tStr
50610   tStr = hOpt.Retrieve("FilenameSubstitutionsOnlyInTitle")
50620   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
50630     .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
50640    Else
50650     .FilenameSubstitutionsOnlyInTitle = 1
50660   End If
50670   tStr = hOpt.Retrieve("JPEGColorscount")
50680   If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
50690     .JPEGColorscount = CLng(tStr)
50700    Else
50710     .JPEGColorscount = 0
50720   End If
50730   tStr = hOpt.Retrieve("JPEGQuality")
50740   If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
50750     .JPEGQuality = CLng(tStr)
50760    Else
50770     .JPEGQuality = 75
50780   End If
50790   tStr = hOpt.Retrieve("Language", "english")
50800   .Language = tStr
50810   tStr = hOpt.Retrieve("LastSaveDirectory", GetMyFiles)
50820   If DirExists(tStr) = True Then
50830     .LastSaveDirectory = CompletePath(tStr)
50840    Else
50850     .LastSaveDirectory = GetMyFiles
50860   End If
50870   tStr = hOpt.Retrieve("Logging")
50880   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
50890     .Logging = CLng(tStr)
50900    Else
50910     .Logging = 0
50920   End If
50930   tStr = hOpt.Retrieve("LogLines")
50940   If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
50950     .LogLines = CLng(tStr)
50960    Else
50970     .LogLines = 100
50980   End If
50990   tStr = hOpt.Retrieve("NoConfirmMessageSwitchingDefaultprinter")
51000   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51010     .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
51020    Else
51030     .NoConfirmMessageSwitchingDefaultprinter = 0
51040   End If
51050   tStr = hOpt.Retrieve("OnePagePerFile")
51060   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51070     .OnePagePerFile = CLng(tStr)
51080    Else
51090     .OnePagePerFile = 0
51100   End If
51110   tStr = hOpt.Retrieve("OptionsEnabled")
51120   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51130     .OptionsEnabled = CLng(tStr)
51140    Else
51150     .OptionsEnabled = 1
51160   End If
51170   tStr = hOpt.Retrieve("OptionsVisible")
51180   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51190     .OptionsVisible = CLng(tStr)
51200    Else
51210     .OptionsVisible = 1
51220   End If
51230   tStr = hOpt.Retrieve("PCXColorscount")
51240   If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
51250     .PCXColorscount = CLng(tStr)
51260    Else
51270     .PCXColorscount = 0
51280   End If
51290   tStr = hOpt.Retrieve("PDFAllowAssembly")
51300   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51310     .PDFAllowAssembly = CLng(tStr)
51320    Else
51330     .PDFAllowAssembly = 0
51340   End If
51350   tStr = hOpt.Retrieve("PDFAllowDegradedPrinting")
51360   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51370     .PDFAllowDegradedPrinting = CLng(tStr)
51380    Else
51390     .PDFAllowDegradedPrinting = 0
51400   End If
51410   tStr = hOpt.Retrieve("PDFAllowFillIn")
51420   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51430     .PDFAllowFillIn = CLng(tStr)
51440    Else
51450     .PDFAllowFillIn = 0
51460   End If
51470   tStr = hOpt.Retrieve("PDFAllowScreenReaders")
51480   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51490     .PDFAllowScreenReaders = CLng(tStr)
51500    Else
51510     .PDFAllowScreenReaders = 0
51520   End If
51530   tStr = hOpt.Retrieve("PDFColorsCMYKToRGB")
51540   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51550     .PDFColorsCMYKToRGB = CLng(tStr)
51560    Else
51570     .PDFColorsCMYKToRGB = 1
51580   End If
51590   tStr = hOpt.Retrieve("PDFColorsColorModel")
51600   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
51610     .PDFColorsColorModel = CLng(tStr)
51620    Else
51630     .PDFColorsColorModel = 1
51640   End If
51650   tStr = hOpt.Retrieve("PDFColorsPreserveHalftone")
51660   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51670     .PDFColorsPreserveHalftone = CLng(tStr)
51680    Else
51690     .PDFColorsPreserveHalftone = 0
51700   End If
51710   tStr = hOpt.Retrieve("PDFColorsPreserveOverprint")
51720   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51730     .PDFColorsPreserveOverprint = CLng(tStr)
51740    Else
51750     .PDFColorsPreserveOverprint = 1
51760   End If
51770   tStr = hOpt.Retrieve("PDFColorsPreserveTransfer")
51780   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51790     .PDFColorsPreserveTransfer = CLng(tStr)
51800    Else
51810     .PDFColorsPreserveTransfer = 1
51820   End If
51830   tStr = hOpt.Retrieve("PDFCompressionColorCompression")
51840   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51850     .PDFCompressionColorCompression = CLng(tStr)
51860    Else
51870     .PDFCompressionColorCompression = 1
51880   End If
51890   tStr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
51900   If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
51910     .PDFCompressionColorCompressionChoice = CLng(tStr)
51920    Else
51930     .PDFCompressionColorCompressionChoice = 0
51940   End If
51950   tStr = hOpt.Retrieve("PDFCompressionColorResample")
51960   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51970     .PDFCompressionColorResample = CLng(tStr)
51980    Else
51990     .PDFCompressionColorResample = 0
52000   End If
52010   tStr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
52020   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
52030     .PDFCompressionColorResampleChoice = CLng(tStr)
52040    Else
52050     .PDFCompressionColorResampleChoice = 0
52060   End If
52070   tStr = hOpt.Retrieve("PDFCompressionColorResolution")
52080   If CLng(tStr) >= 0 Then
52090     .PDFCompressionColorResolution = CLng(tStr)
52100    Else
52110     .PDFCompressionColorResolution = 300
52120   End If
52130   tStr = hOpt.Retrieve("PDFCompressionGreyCompression")
52140   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52150     .PDFCompressionGreyCompression = CLng(tStr)
52160    Else
52170     .PDFCompressionGreyCompression = 1
52180   End If
52190   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
52200   If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
52210     .PDFCompressionGreyCompressionChoice = CLng(tStr)
52220    Else
52230     .PDFCompressionGreyCompressionChoice = 0
52240   End If
52250   tStr = hOpt.Retrieve("PDFCompressionGreyResample")
52260   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52270     .PDFCompressionGreyResample = CLng(tStr)
52280    Else
52290     .PDFCompressionGreyResample = 0
52300   End If
52310   tStr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
52320   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
52330     .PDFCompressionGreyResampleChoice = CLng(tStr)
52340    Else
52350     .PDFCompressionGreyResampleChoice = 0
52360   End If
52370   tStr = hOpt.Retrieve("PDFCompressionGreyResolution")
52380   If CLng(tStr) >= 0 Then
52390     .PDFCompressionGreyResolution = CLng(tStr)
52400    Else
52410     .PDFCompressionGreyResolution = 300
52420   End If
52430   tStr = hOpt.Retrieve("PDFCompressionMonoCompression")
52440   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52450     .PDFCompressionMonoCompression = CLng(tStr)
52460    Else
52470     .PDFCompressionMonoCompression = 1
52480   End If
52490   tStr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
52500   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
52510     .PDFCompressionMonoCompressionChoice = CLng(tStr)
52520    Else
52530     .PDFCompressionMonoCompressionChoice = 0
52540   End If
52550   tStr = hOpt.Retrieve("PDFCompressionMonoResample")
52560   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52570     .PDFCompressionMonoResample = CLng(tStr)
52580    Else
52590     .PDFCompressionMonoResample = 0
52600   End If
52610   tStr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
52620   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
52630     .PDFCompressionMonoResampleChoice = CLng(tStr)
52640    Else
52650     .PDFCompressionMonoResampleChoice = 0
52660   End If
52670   tStr = hOpt.Retrieve("PDFCompressionMonoResolution")
52680   If CLng(tStr) >= 0 Then
52690     .PDFCompressionMonoResolution = CLng(tStr)
52700    Else
52710     .PDFCompressionMonoResolution = 1200
52720   End If
52730   tStr = hOpt.Retrieve("PDFCompressionTextCompression")
52740   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52750     .PDFCompressionTextCompression = CLng(tStr)
52760    Else
52770     .PDFCompressionTextCompression = 1
52780   End If
52790   tStr = hOpt.Retrieve("PDFDisallowCopy")
52800   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52810     .PDFDisallowCopy = CLng(tStr)
52820    Else
52830     .PDFDisallowCopy = 1
52840   End If
52850   tStr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
52860   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52870     .PDFDisallowModifyAnnotations = CLng(tStr)
52880    Else
52890     .PDFDisallowModifyAnnotations = 0
52900   End If
52910   tStr = hOpt.Retrieve("PDFDisallowModifyContents")
52920   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52930     .PDFDisallowModifyContents = CLng(tStr)
52940    Else
52950     .PDFDisallowModifyContents = 0
52960   End If
52970   tStr = hOpt.Retrieve("PDFDisallowPrinting")
52980   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52990     .PDFDisallowPrinting = CLng(tStr)
53000    Else
53010     .PDFDisallowPrinting = 0
53020   End If
53030   tStr = hOpt.Retrieve("PDFEncryptor")
53040   If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
53050     .PDFEncryptor = CLng(tStr)
53060    Else
53070     .PDFEncryptor = 0
53080   End If
53090   tStr = hOpt.Retrieve("PDFFontsEmbedAll")
53100   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53110     .PDFFontsEmbedAll = CLng(tStr)
53120    Else
53130     .PDFFontsEmbedAll = 1
53140   End If
53150   tStr = hOpt.Retrieve("PDFFontsSubSetFonts")
53160   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53170     .PDFFontsSubSetFonts = CLng(tStr)
53180    Else
53190     .PDFFontsSubSetFonts = 1
53200   End If
53210   tStr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
53220   If CLng(tStr) >= 0 Then
53230     .PDFFontsSubSetFontsPercent = CLng(tStr)
53240    Else
53250     .PDFFontsSubSetFontsPercent = 100
53260   End If
53270   tStr = hOpt.Retrieve("PDFGeneralASCII85")
53280   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53290     .PDFGeneralASCII85 = CLng(tStr)
53300    Else
53310     .PDFGeneralASCII85 = 0
53320   End If
53330   tStr = hOpt.Retrieve("PDFGeneralAutorotate")
53340   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
53350     .PDFGeneralAutorotate = CLng(tStr)
53360    Else
53370     .PDFGeneralAutorotate = 2
53380   End If
53390   tStr = hOpt.Retrieve("PDFGeneralCompatibility")
53400   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
53410     .PDFGeneralCompatibility = CLng(tStr)
53420    Else
53430     .PDFGeneralCompatibility = 1
53440   End If
53450   tStr = hOpt.Retrieve("PDFGeneralOverprint")
53460   If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
53470     .PDFGeneralOverprint = CLng(tStr)
53480    Else
53490     .PDFGeneralOverprint = 0
53500   End If
53510   tStr = hOpt.Retrieve("PDFGeneralResolution")
53520   If CLng(tStr) >= 0 Then
53530     .PDFGeneralResolution = CLng(tStr)
53540    Else
53550     .PDFGeneralResolution = 600
53560   End If
53570   tStr = hOpt.Retrieve("PDFHighEncryption")
53580   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53590     .PDFHighEncryption = CLng(tStr)
53600    Else
53610     .PDFHighEncryption = 0
53620   End If
53630   tStr = hOpt.Retrieve("PDFLowEncryption")
53640   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53650     .PDFLowEncryption = CLng(tStr)
53660    Else
53670     .PDFLowEncryption = 1
53680   End If
53690   tStr = hOpt.Retrieve("PDFOwnerPass")
53700   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53710     .PDFOwnerPass = CLng(tStr)
53720    Else
53730     .PDFOwnerPass = 0
53740   End If
53750   tStr = hOpt.Retrieve("PDFOwnerPasswordString", " ")
53760   .PDFOwnerPasswordString = tStr
53770   tStr = hOpt.Retrieve("PDFUserPass")
53780   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53790     .PDFUserPass = CLng(tStr)
53800    Else
53810     .PDFUserPass = 0
53820   End If
53830   tStr = hOpt.Retrieve("PDFUserPasswordString", " ")
53840   .PDFUserPasswordString = tStr
53850   tStr = hOpt.Retrieve("PDFUseSecurity")
53860   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53870     .PDFUseSecurity = CLng(tStr)
53880    Else
53890     .PDFUseSecurity = 0
53900   End If
53910   tStr = hOpt.Retrieve("PNGColorscount")
53920   If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
53930     .PNGColorscount = CLng(tStr)
53940    Else
53950     .PNGColorscount = 0
53960   End If
53970   tStr = hOpt.Retrieve("PrinterStop")
53980   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53990     .PrinterStop = CLng(tStr)
54000    Else
54010     .PrinterStop = 0
54020   End If
54030   tStr = hOpt.Retrieve("PrinterTemppath", GetTempPath)
54040   If DirExists(tStr) = True Then
54050     .PrinterTemppath = CompletePath(tStr)
54060    Else
54070     .PrinterTemppath = GetTempPath
54080   End If
54090   tStr = hOpt.Retrieve("ProcessPriority")
54100   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
54110     .ProcessPriority = CLng(tStr)
54120    Else
54130     .ProcessPriority = 1
54140   End If
54150   tStr = hOpt.Retrieve("ProgramFont", "MS Sans Serif")
54160   .ProgramFont = tStr
54170   tStr = hOpt.Retrieve("ProgramFontCharset")
54180   If CLng(tStr) >= 0 Then
54190     .ProgramFontCharset = CLng(tStr)
54200    Else
54210     .ProgramFontCharset = 0
54220   End If
54230   tStr = hOpt.Retrieve("ProgramFontSize")
54240   If CLng(tStr) >= 1 And CLng(tStr) <= 72 Then
54250     .ProgramFontSize = CLng(tStr)
54260    Else
54270     .ProgramFontSize = 8
54280   End If
54290   tStr = hOpt.Retrieve("PSLanguageLevel")
54300   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
54310     .PSLanguageLevel = CLng(tStr)
54320    Else
54330     .PSLanguageLevel = 2
54340   End If
54350   tStr = hOpt.Retrieve("RemoveSpaces")
54360   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54370     .RemoveSpaces = CLng(tStr)
54380    Else
54390     .RemoveSpaces = 1
54400   End If
54410   tStr = hOpt.Retrieve("RunProgramAfterSaving")
54420   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54430     .RunProgramAfterSaving = CLng(tStr)
54440    Else
54450     .RunProgramAfterSaving = 0
54460   End If
54470   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramname", " ")
54480   .RunProgramAfterSavingProgramname = tStr
54490   tStr = hOpt.Retrieve("RunProgramAfterSavingProgramParameters", " ")
54500   .RunProgramAfterSavingProgramParameters = tStr
54510   tStr = hOpt.Retrieve("RunProgramAfterSavingWaitUntilReady")
54520   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54530     .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
54540    Else
54550     .RunProgramAfterSavingWaitUntilReady = 1
54560   End If
54570   tStr = hOpt.Retrieve("RunProgramAfterSavingWindowstyle")
54580   If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
54590     .RunProgramAfterSavingWindowstyle = CLng(tStr)
54600    Else
54610     .RunProgramAfterSavingWindowstyle = 1
54620   End If
54630   tStr = hOpt.Retrieve("SaveFilename", "<Title>")
54640   .SaveFilename = tStr
54650   tStr = hOpt.Retrieve("ShowAnimation")
54660   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54670     .ShowAnimation = CLng(tStr)
54680    Else
54690     .ShowAnimation = 1
54700   End If
54710   tStr = hOpt.Retrieve("StampFontColor", "#FF0000")
54720   .StampFontColor = tStr
54730   tStr = hOpt.Retrieve("StampFontname", "Arial")
54740   .StampFontname = tStr
54750   tStr = hOpt.Retrieve("StampFontsize")
54760   If CLng(tStr) >= 1 Then
54770     .StampFontsize = CLng(tStr)
54780    Else
54790     .StampFontsize = 48
54800   End If
54810   tStr = hOpt.Retrieve("StampOutlineFontthickness")
54820   If CLng(tStr) >= 0 Then
54830     .StampOutlineFontthickness = CLng(tStr)
54840    Else
54850     .StampOutlineFontthickness = 0
54860   End If
54870   tStr = hOpt.Retrieve("StampString", " ")
54880   .StampString = tStr
54890   tStr = hOpt.Retrieve("StampUseOutlineFont")
54900   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54910     .StampUseOutlineFont = CLng(tStr)
54920    Else
54930     .StampUseOutlineFont = 1
54940   End If
54950   tStr = hOpt.Retrieve("StandardAuthor", " ")
54960   .StandardAuthor = tStr
54970   tStr = hOpt.Retrieve("StandardCreationdate", " ")
54980   .StandardCreationdate = tStr
54990   tStr = hOpt.Retrieve("StandardDateformat", "YYYYMMDDHHNNSS")
55000   .StandardDateformat = tStr
55010   tStr = hOpt.Retrieve("StandardKeywords", " ")
55020   .StandardKeywords = tStr
55030   tStr = hOpt.Retrieve("StandardModifydate", " ")
55040   .StandardModifydate = tStr
55050   tStr = hOpt.Retrieve("StandardSubject", " ")
55060   .StandardSubject = tStr
55070   tStr = hOpt.Retrieve("StandardTitle", " ")
55080   .StandardTitle = tStr
55090   tStr = hOpt.Retrieve("StartStandardProgram")
55100   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55110     .StartStandardProgram = CLng(tStr)
55120    Else
55130     .StartStandardProgram = 1
55140   End If
55150   tStr = hOpt.Retrieve("TIFFColorscount")
55160   If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
55170     .TIFFColorscount = CLng(tStr)
55180    Else
55190     .TIFFColorscount = 0
55200   End If
55210   tStr = hOpt.Retrieve("UseAutosave")
55220   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55230     .UseAutosave = CLng(tStr)
55240    Else
55250     .UseAutosave = 0
55260   End If
55270   tStr = hOpt.Retrieve("UseAutosaveDirectory")
55280   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55290     .UseAutosaveDirectory = CLng(tStr)
55300    Else
55310     .UseAutosaveDirectory = 1
55320   End If
55330   tStr = hOpt.Retrieve("UseCreationDateNow")
55340   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55350     .UseCreationDateNow = CLng(tStr)
55360    Else
55370     .UseCreationDateNow = 0
55380   End If
55390   tStr = hOpt.Retrieve("UseStandardAuthor")
55400   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
55410     .UseStandardAuthor = CLng(tStr)
55420    Else
55430     .UseStandardAuthor = 0
55440   End If
55450  End With
55460  Set ini = Nothing
55470  ReadOptions = myOptions
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
50090   ini.SaveKey CStr(.AutosaveDirectory), "AutosaveDirectory"
50100   ini.SaveKey CStr(.AutosaveFilename), "AutosaveFilename"
50110   ini.SaveKey CStr(.AutosaveFormat), "AutosaveFormat"
50120   ini.SaveKey CStr(.BitmapResolution), "BitmapResolution"
50130   ini.SaveKey CStr(.BMPColorscount), "BMPColorscount"
50140   ini.SaveKey CStr(.DirectoryGhostscriptBinaries), "DirectoryGhostscriptBinaries"
50150   ini.SaveKey CStr(.DirectoryGhostscriptFonts), "DirectoryGhostscriptFonts"
50160   ini.SaveKey CStr(.DirectoryGhostscriptLibraries), "DirectoryGhostscriptLibraries"
50170   ini.SaveKey CStr(.DirectoryGhostscriptResource), "DirectoryGhostscriptResource"
50180   ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
50190   ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
50200   ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
50210   ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50220   ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50230   ini.SaveKey CStr(.Language), "Language"
50240   ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50250   ini.SaveKey CStr(Abs(.Logging)), "Logging"
50260   ini.SaveKey CStr(.LogLines), "LogLines"
50270   ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50280   ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
50290   ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50300   ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50310   ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50320   ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50330   ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50340   ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50350   ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50360   ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50370   ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50380   ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50390   ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50400   ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50410   ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50420   ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50430   ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50440   ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50450   ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50460   ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50470   ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50480   ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50490   ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50500   ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50510   ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50520   ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50530   ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50540   ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50550   ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50560   ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50570   ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50580   ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50590   ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50600   ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50610   ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50620   ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50630   ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50640   ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50650   ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50660   ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50670   ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50680   ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50690   ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
50700   ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
50710   ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
50720   ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
50730   ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
50740   ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
50750   ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
50760   ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
50770   ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
50780   ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
50790   ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
50800   ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
50810   ini.SaveKey CStr(.ProgramFont), "ProgramFont"
50820   ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
50830   ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
50840   ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
50850   ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
50860   ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
50870   ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
50880   ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
50890   ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
50900   ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
50910   ini.SaveKey CStr(.SaveFilename), "SaveFilename"
50920   ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
50930   ini.SaveKey CStr(.StampFontColor), "StampFontColor"
50940   ini.SaveKey CStr(.StampFontname), "StampFontname"
50950   ini.SaveKey CStr(.StampFontsize), "StampFontsize"
50960   ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
50970   ini.SaveKey CStr(.StampString), "StampString"
50980   ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
50990   ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
51000   ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
51010   ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
51020   ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
51030   ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
51040   ini.SaveKey CStr(.StandardSubject), "StandardSubject"
51050   ini.SaveKey CStr(.StandardTitle), "StandardTitle"
51060   ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
51070   ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
51080   ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
51090   ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
51100   ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
51110   ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
51120  End With
51130  Set ini = Nothing
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
50040   Else
50050    Options.PrinterStop = 0
50060    PrinterStop = False
50070  End If
50080  SaveOptions Options
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

