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
 PDFUserPass As Long
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
 SaveFilename As String
 StandardAuthor As String
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
50230   .EPSLanguageLevel = "2"
50240   .FilenameSubstitutions = "Microsoft Word - \.doc"
50250   .FilenameSubstitutionsOnlyInTitle = "1"
50260   .JPEGColorscount = "0"
50270   .JPEGQuality = "75"
50280   .Language = "english"
50290   .LastSaveDirectory = GetMyFiles
50300   .Logging = "0"
50310   .LogLines = "100"
50320   .NoConfirmMessageSwitchingDefaultprinter = "0"
50330   .OptionsEnabled = "1"
50340   .OptionsVisible = "1"
50350   .PCXColorscount = "0"
50360   .PDFAllowAssembly = "0"
50370   .PDFAllowDegradedPrinting = "0"
50380   .PDFAllowFillIn = "0"
50390   .PDFAllowScreenReaders = "0"
50400   .PDFColorsCMYKToRGB = "1"
50410   .PDFColorsColorModel = "1"
50420   .PDFColorsPreserveHalftone = "0"
50430   .PDFColorsPreserveOverprint = "1"
50440   .PDFColorsPreserveTransfer = "1"
50450   .PDFCompressionColorCompression = "1"
50460   .PDFCompressionColorCompressionChoice = "0"
50470   .PDFCompressionColorResample = "0"
50480   .PDFCompressionColorResampleChoice = "0"
50490   .PDFCompressionColorResolution = "300"
50500   .PDFCompressionGreyCompression = "1"
50510   .PDFCompressionGreyCompressionChoice = "0"
50520   .PDFCompressionGreyResample = "0"
50530   .PDFCompressionGreyResampleChoice = "0"
50540   .PDFCompressionGreyResolution = "300"
50550   .PDFCompressionMonoCompression = "1"
50560   .PDFCompressionMonoCompressionChoice = "0"
50570   .PDFCompressionMonoResample = "0"
50580   .PDFCompressionMonoResampleChoice = "0"
50590   .PDFCompressionMonoResolution = "1200"
50600   .PDFCompressionTextCompression = "1"
50610   .PDFDisallowCopy = "1"
50620   .PDFDisallowModifyAnnotations = "0"
50630   .PDFDisallowModifyContents = "0"
50640   .PDFDisallowPrinting = "0"
50650   .PDFEncryptor = "0"
50660   .PDFFontsEmbedAll = "1"
50670   .PDFFontsSubSetFonts = "1"
50680   .PDFFontsSubSetFontsPercent = "100"
50690   .PDFGeneralASCII85 = "0"
50700   .PDFGeneralAutorotate = "2"
50710   .PDFGeneralCompatibility = "1"
50720   .PDFGeneralOverprint = "0"
50730   .PDFGeneralResolution = "600"
50740   .PDFHighEncryption = "0"
50750   .PDFLowEncryption = "1"
50760   .PDFOwnerPass = "0"
50770   .PDFUserPass = "0"
50780   .PDFUseSecurity = "0"
50790   .PNGColorscount = "0"
50800   .PrinterStop = "0"
50810   .PrinterTemppath = GetTempPath
50820   .ProcessPriority = "1"
50830   .ProgramFont = "MS Sans Serif"
50840   .ProgramFontCharset = "0"
50850   .ProgramFontSize = "8"
50860   .PSLanguageLevel = "2"
50870   .RemoveSpaces = "1"
50880   .SaveFilename = "<Title>"
50890   .StandardAuthor = " "
50900   .StartStandardProgram = "1"
50910   .TIFFColorscount = "0"
50920   .UseAutosave = "0"
50930   .UseAutosaveDirectory = "1"
50940   .UseCreationDateNow = "0"
50950   .UseStandardAuthor = "0"
50960  End With
50970  StandardOptions = myOptions
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
50120   If DirExists(tStr) = True Then
50130     .AutosaveDirectory = CompletePath(tStr)
50140    Else
50150     .AutosaveDirectory = GetMyFiles
50160   End If
50170   tStr = hOpt.Retrieve("AutosaveFilename", "<DateTime>")
50180   .AutosaveFilename = tStr
50190   tStr = hOpt.Retrieve("AutosaveFormat")
50200   If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
50210     .AutosaveFormat = CLng(tStr)
50220    Else
50230     .AutosaveFormat = 0
50240   End If
50250   tStr = hOpt.Retrieve("BitmapResolution")
50260   If CLng(tStr) >= 1 Then
50270     .BitmapResolution = CLng(tStr)
50280    Else
50290     .BitmapResolution = 150
50300   End If
50310   tStr = hOpt.Retrieve("BMPColorscount")
50320   If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
50330     .BMPColorscount = CLng(tStr)
50340    Else
50350     .BMPColorscount = 1
50360   End If
50370   tStr = hOpt.Retrieve("DirectoryGhostscriptBinaries", App.Path)
50380   If DirExists(tStr) = True Then
50390     .DirectoryGhostscriptBinaries = CompletePath(tStr)
50400    Else
50410     .DirectoryGhostscriptBinaries = ""
50420   End If
50430   tStr = hOpt.Retrieve("DirectoryGhostscriptFonts", App.Path & "\fonts")
50440   If DirExists(tStr) = True Then
50450     .DirectoryGhostscriptFonts = CompletePath(tStr)
50460    Else
50470     .DirectoryGhostscriptFonts = ""
50480   End If
50490   tStr = hOpt.Retrieve("DirectoryGhostscriptLibraries", App.Path & "\lib")
50500   If DirExists(tStr) = True Then
50510     .DirectoryGhostscriptLibraries = CompletePath(tStr)
50520    Else
50530     .DirectoryGhostscriptLibraries = ""
50540   End If
50550   tStr = hOpt.Retrieve("EPSLanguageLevel")
50560   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
50570     .EPSLanguageLevel = CLng(tStr)
50580    Else
50590     .EPSLanguageLevel = 2
50600   End If
50610   tStr = hOpt.Retrieve("FilenameSubstitutions", "Microsoft Word - \.doc")
50620   .FilenameSubstitutions = tStr
50630   tStr = hOpt.Retrieve("FilenameSubstitutionsOnlyInTitle")
50640   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
50650     .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
50660    Else
50670     .FilenameSubstitutionsOnlyInTitle = 1
50680   End If
50690   tStr = hOpt.Retrieve("JPEGColorscount")
50700   If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
50710     .JPEGColorscount = CLng(tStr)
50720    Else
50730     .JPEGColorscount = 0
50740   End If
50750   tStr = hOpt.Retrieve("JPEGQuality")
50760   If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
50770     .JPEGQuality = CLng(tStr)
50780    Else
50790     .JPEGQuality = 75
50800   End If
50810   tStr = hOpt.Retrieve("Language", "english")
50820   .Language = tStr
50830   tStr = hOpt.Retrieve("LastSaveDirectory", GetMyFiles)
50840   If DirExists(tStr) = True Then
50850     .LastSaveDirectory = CompletePath(tStr)
50860    Else
50870     .LastSaveDirectory = GetMyFiles
50880   End If
50890   tStr = hOpt.Retrieve("Logging")
50900   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
50910     .Logging = CLng(tStr)
50920    Else
50930     .Logging = 0
50940   End If
50950   tStr = hOpt.Retrieve("LogLines")
50960   If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
50970     .LogLines = CLng(tStr)
50980    Else
50990     .LogLines = 100
51000   End If
51010   tStr = hOpt.Retrieve("NoConfirmMessageSwitchingDefaultprinter")
51020   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51030     .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
51040    Else
51050     .NoConfirmMessageSwitchingDefaultprinter = 0
51060   End If
51070   tStr = hOpt.Retrieve("OptionsEnabled")
51080   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51090     .OptionsEnabled = CLng(tStr)
51100    Else
51110     .OptionsEnabled = 1
51120   End If
51130   tStr = hOpt.Retrieve("OptionsVisible")
51140   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51150     .OptionsVisible = CLng(tStr)
51160    Else
51170     .OptionsVisible = 1
51180   End If
51190   tStr = hOpt.Retrieve("PCXColorscount")
51200   If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
51210     .PCXColorscount = CLng(tStr)
51220    Else
51230     .PCXColorscount = 0
51240   End If
51250   tStr = hOpt.Retrieve("PDFAllowAssembly")
51260   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51270     .PDFAllowAssembly = CLng(tStr)
51280    Else
51290     .PDFAllowAssembly = 0
51300   End If
51310   tStr = hOpt.Retrieve("PDFAllowDegradedPrinting")
51320   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51330     .PDFAllowDegradedPrinting = CLng(tStr)
51340    Else
51350     .PDFAllowDegradedPrinting = 0
51360   End If
51370   tStr = hOpt.Retrieve("PDFAllowFillIn")
51380   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51390     .PDFAllowFillIn = CLng(tStr)
51400    Else
51410     .PDFAllowFillIn = 0
51420   End If
51430   tStr = hOpt.Retrieve("PDFAllowScreenReaders")
51440   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51450     .PDFAllowScreenReaders = CLng(tStr)
51460    Else
51470     .PDFAllowScreenReaders = 0
51480   End If
51490   tStr = hOpt.Retrieve("PDFColorsCMYKToRGB")
51500   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51510     .PDFColorsCMYKToRGB = CLng(tStr)
51520    Else
51530     .PDFColorsCMYKToRGB = 1
51540   End If
51550   tStr = hOpt.Retrieve("PDFColorsColorModel")
51560   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
51570     .PDFColorsColorModel = CLng(tStr)
51580    Else
51590     .PDFColorsColorModel = 1
51600   End If
51610   tStr = hOpt.Retrieve("PDFColorsPreserveHalftone")
51620   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51630     .PDFColorsPreserveHalftone = CLng(tStr)
51640    Else
51650     .PDFColorsPreserveHalftone = 0
51660   End If
51670   tStr = hOpt.Retrieve("PDFColorsPreserveOverprint")
51680   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51690     .PDFColorsPreserveOverprint = CLng(tStr)
51700    Else
51710     .PDFColorsPreserveOverprint = 1
51720   End If
51730   tStr = hOpt.Retrieve("PDFColorsPreserveTransfer")
51740   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51750     .PDFColorsPreserveTransfer = CLng(tStr)
51760    Else
51770     .PDFColorsPreserveTransfer = 1
51780   End If
51790   tStr = hOpt.Retrieve("PDFCompressionColorCompression")
51800   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51810     .PDFCompressionColorCompression = CLng(tStr)
51820    Else
51830     .PDFCompressionColorCompression = 1
51840   End If
51850   tStr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
51860   If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
51870     .PDFCompressionColorCompressionChoice = CLng(tStr)
51880    Else
51890     .PDFCompressionColorCompressionChoice = 0
51900   End If
51910   tStr = hOpt.Retrieve("PDFCompressionColorResample")
51920   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51930     .PDFCompressionColorResample = CLng(tStr)
51940    Else
51950     .PDFCompressionColorResample = 0
51960   End If
51970   tStr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
51980   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
51990     .PDFCompressionColorResampleChoice = CLng(tStr)
52000    Else
52010     .PDFCompressionColorResampleChoice = 0
52020   End If
52030   tStr = hOpt.Retrieve("PDFCompressionColorResolution")
52040   If CLng(tStr) >= 0 Then
52050     .PDFCompressionColorResolution = CLng(tStr)
52060    Else
52070     .PDFCompressionColorResolution = 300
52080   End If
52090   tStr = hOpt.Retrieve("PDFCompressionGreyCompression")
52100   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52110     .PDFCompressionGreyCompression = CLng(tStr)
52120    Else
52130     .PDFCompressionGreyCompression = 1
52140   End If
52150   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
52160   If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
52170     .PDFCompressionGreyCompressionChoice = CLng(tStr)
52180    Else
52190     .PDFCompressionGreyCompressionChoice = 0
52200   End If
52210   tStr = hOpt.Retrieve("PDFCompressionGreyResample")
52220   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52230     .PDFCompressionGreyResample = CLng(tStr)
52240    Else
52250     .PDFCompressionGreyResample = 0
52260   End If
52270   tStr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
52280   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
52290     .PDFCompressionGreyResampleChoice = CLng(tStr)
52300    Else
52310     .PDFCompressionGreyResampleChoice = 0
52320   End If
52330   tStr = hOpt.Retrieve("PDFCompressionGreyResolution")
52340   If CLng(tStr) >= 0 Then
52350     .PDFCompressionGreyResolution = CLng(tStr)
52360    Else
52370     .PDFCompressionGreyResolution = 300
52380   End If
52390   tStr = hOpt.Retrieve("PDFCompressionMonoCompression")
52400   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52410     .PDFCompressionMonoCompression = CLng(tStr)
52420    Else
52430     .PDFCompressionMonoCompression = 1
52440   End If
52450   tStr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
52460   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
52470     .PDFCompressionMonoCompressionChoice = CLng(tStr)
52480    Else
52490     .PDFCompressionMonoCompressionChoice = 0
52500   End If
52510   tStr = hOpt.Retrieve("PDFCompressionMonoResample")
52520   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52530     .PDFCompressionMonoResample = CLng(tStr)
52540    Else
52550     .PDFCompressionMonoResample = 0
52560   End If
52570   tStr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
52580   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
52590     .PDFCompressionMonoResampleChoice = CLng(tStr)
52600    Else
52610     .PDFCompressionMonoResampleChoice = 0
52620   End If
52630   tStr = hOpt.Retrieve("PDFCompressionMonoResolution")
52640   If CLng(tStr) >= 0 Then
52650     .PDFCompressionMonoResolution = CLng(tStr)
52660    Else
52670     .PDFCompressionMonoResolution = 1200
52680   End If
52690   tStr = hOpt.Retrieve("PDFCompressionTextCompression")
52700   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52710     .PDFCompressionTextCompression = CLng(tStr)
52720    Else
52730     .PDFCompressionTextCompression = 1
52740   End If
52750   tStr = hOpt.Retrieve("PDFDisallowCopy")
52760   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52770     .PDFDisallowCopy = CLng(tStr)
52780    Else
52790     .PDFDisallowCopy = 1
52800   End If
52810   tStr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
52820   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52830     .PDFDisallowModifyAnnotations = CLng(tStr)
52840    Else
52850     .PDFDisallowModifyAnnotations = 0
52860   End If
52870   tStr = hOpt.Retrieve("PDFDisallowModifyContents")
52880   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52890     .PDFDisallowModifyContents = CLng(tStr)
52900    Else
52910     .PDFDisallowModifyContents = 0
52920   End If
52930   tStr = hOpt.Retrieve("PDFDisallowPrinting")
52940   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52950     .PDFDisallowPrinting = CLng(tStr)
52960    Else
52970     .PDFDisallowPrinting = 0
52980   End If
52990   tStr = hOpt.Retrieve("PDFEncryptor")
53000   If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
53010     .PDFEncryptor = CLng(tStr)
53020    Else
53030     .PDFEncryptor = 0
53040   End If
53050   tStr = hOpt.Retrieve("PDFFontsEmbedAll")
53060   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53070     .PDFFontsEmbedAll = CLng(tStr)
53080    Else
53090     .PDFFontsEmbedAll = 1
53100   End If
53110   tStr = hOpt.Retrieve("PDFFontsSubSetFonts")
53120   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53130     .PDFFontsSubSetFonts = CLng(tStr)
53140    Else
53150     .PDFFontsSubSetFonts = 1
53160   End If
53170   tStr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
53180   If CLng(tStr) >= 0 Then
53190     .PDFFontsSubSetFontsPercent = CLng(tStr)
53200    Else
53210     .PDFFontsSubSetFontsPercent = 100
53220   End If
53230   tStr = hOpt.Retrieve("PDFGeneralASCII85")
53240   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53250     .PDFGeneralASCII85 = CLng(tStr)
53260    Else
53270     .PDFGeneralASCII85 = 0
53280   End If
53290   tStr = hOpt.Retrieve("PDFGeneralAutorotate")
53300   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
53310     .PDFGeneralAutorotate = CLng(tStr)
53320    Else
53330     .PDFGeneralAutorotate = 2
53340   End If
53350   tStr = hOpt.Retrieve("PDFGeneralCompatibility")
53360   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
53370     .PDFGeneralCompatibility = CLng(tStr)
53380    Else
53390     .PDFGeneralCompatibility = 1
53400   End If
53410   tStr = hOpt.Retrieve("PDFGeneralOverprint")
53420   If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
53430     .PDFGeneralOverprint = CLng(tStr)
53440    Else
53450     .PDFGeneralOverprint = 0
53460   End If
53470   tStr = hOpt.Retrieve("PDFGeneralResolution")
53480   If CLng(tStr) >= 0 Then
53490     .PDFGeneralResolution = CLng(tStr)
53500    Else
53510     .PDFGeneralResolution = 600
53520   End If
53530   tStr = hOpt.Retrieve("PDFHighEncryption")
53540   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53550     .PDFHighEncryption = CLng(tStr)
53560    Else
53570     .PDFHighEncryption = 0
53580   End If
53590   tStr = hOpt.Retrieve("PDFLowEncryption")
53600   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53610     .PDFLowEncryption = CLng(tStr)
53620    Else
53630     .PDFLowEncryption = 1
53640   End If
53650   tStr = hOpt.Retrieve("PDFOwnerPass")
53660   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53670     .PDFOwnerPass = CLng(tStr)
53680    Else
53690     .PDFOwnerPass = 0
53700   End If
53710   tStr = hOpt.Retrieve("PDFUserPass")
53720   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53730     .PDFUserPass = CLng(tStr)
53740    Else
53750     .PDFUserPass = 0
53760   End If
53770   tStr = hOpt.Retrieve("PDFUseSecurity")
53780   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53790     .PDFUseSecurity = CLng(tStr)
53800    Else
53810     .PDFUseSecurity = 0
53820   End If
53830   tStr = hOpt.Retrieve("PNGColorscount")
53840   If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
53850     .PNGColorscount = CLng(tStr)
53860    Else
53870     .PNGColorscount = 0
53880   End If
53890   tStr = hOpt.Retrieve("PrinterStop")
53900   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53910     .PrinterStop = CLng(tStr)
53920    Else
53930     .PrinterStop = 0
53940   End If
53950   tStr = hOpt.Retrieve("PrinterTemppath", GetTempPath)
53960   If DirExists(tStr) = True Then
53970     .PrinterTemppath = CompletePath(tStr)
53980    Else
53990     .PrinterTemppath = GetTempPath
54000   End If
54010   tStr = hOpt.Retrieve("ProcessPriority")
54020   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
54030     .ProcessPriority = CLng(tStr)
54040    Else
54050     .ProcessPriority = 1
54060   End If
54070   tStr = hOpt.Retrieve("ProgramFont", "MS Sans Serif")
54080   .ProgramFont = tStr
54090   tStr = hOpt.Retrieve("ProgramFontCharset")
54100   If CLng(tStr) >= 0 Then
54110     .ProgramFontCharset = CLng(tStr)
54120    Else
54130     .ProgramFontCharset = 0
54140   End If
54150   tStr = hOpt.Retrieve("ProgramFontSize")
54160   If CLng(tStr) >= 1 And CLng(tStr) <= 72 Then
54170     .ProgramFontSize = CLng(tStr)
54180    Else
54190     .ProgramFontSize = 8
54200   End If
54210   tStr = hOpt.Retrieve("PSLanguageLevel")
54220   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
54230     .PSLanguageLevel = CLng(tStr)
54240    Else
54250     .PSLanguageLevel = 2
54260   End If
54270   tStr = hOpt.Retrieve("RemoveSpaces")
54280   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54290     .RemoveSpaces = CLng(tStr)
54300    Else
54310     .RemoveSpaces = 1
54320   End If
54330   tStr = hOpt.Retrieve("SaveFilename", "<Title>")
54340   .SaveFilename = tStr
54350   tStr = hOpt.Retrieve("StandardAuthor", " ")
54360   .StandardAuthor = tStr
54370   tStr = hOpt.Retrieve("StartStandardProgram")
54380   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54390     .StartStandardProgram = CLng(tStr)
54400    Else
54410     .StartStandardProgram = 1
54420   End If
54430   tStr = hOpt.Retrieve("TIFFColorscount")
54440   If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
54450     .TIFFColorscount = CLng(tStr)
54460    Else
54470     .TIFFColorscount = 0
54480   End If
54490   tStr = hOpt.Retrieve("UseAutosave")
54500   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54510     .UseAutosave = CLng(tStr)
54520    Else
54530     .UseAutosave = 0
54540   End If
54550   tStr = hOpt.Retrieve("UseAutosaveDirectory")
54560   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54570     .UseAutosaveDirectory = CLng(tStr)
54580    Else
54590     .UseAutosaveDirectory = 1
54600   End If
54610   tStr = hOpt.Retrieve("UseCreationDateNow")
54620   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54630     .UseCreationDateNow = CLng(tStr)
54640    Else
54650     .UseCreationDateNow = 0
54660   End If
54670   tStr = hOpt.Retrieve("UseStandardAuthor")
54680   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54690     .UseStandardAuthor = CLng(tStr)
54700    Else
54710     .UseStandardAuthor = 0
54720   End If
54730  End With
54740  Set ini = Nothing
54750  ReadOptions = myOptions
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
50170   ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
50180   ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
50190   ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
50200   ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50210   ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50220   ini.SaveKey CStr(.Language), "Language"
50230   ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50240   ini.SaveKey CStr(Abs(.Logging)), "Logging"
50250   ini.SaveKey CStr(.LogLines), "LogLines"
50260   ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
50270   ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
50280   ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
50290   ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50300   ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
50310   ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
50320   ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
50330   ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
50340   ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
50350   ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50360   ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
50370   ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
50380   ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
50390   ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
50400   ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50410   ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
50420   ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50430   ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50440   ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
50450   ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50460   ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
50470   ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50480   ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50490   ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
50500   ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50510   ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
50520   ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50530   ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50540   ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
50550   ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
50560   ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
50570   ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
50580   ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
50590   ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
50600   ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
50610   ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
50620   ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50630   ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
50640   ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50650   ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50660   ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50670   ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
50680   ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
50690   ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
50700   ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
50710   ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
50720   ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
50730   ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
50740   ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
50750   ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
50760   ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
50770   ini.SaveKey CStr(.ProgramFont), "ProgramFont"
50780   ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
50790   ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
50800   ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
50810   ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
50820   ini.SaveKey CStr(.SaveFilename), "SaveFilename"
50830   ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
50840   ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
50850   ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
50860   ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
50870   ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
50880   ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
50890   ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
50900  End With
50910  Set ini = Nothing
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
  Frm.txtProgramFontsize.Text = .ProgramFontSize
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
50110  .EPSLanguageLevel = Frm.cmbEPSLanguageLevel.ListIndex
50120  tStr = ""
50130  Set lsv = Frm.lsvFilenameSubst
50140  For i = 1 To lsv.ListItems.Count
50150   If i < lsv.ListItems.Count Then
50160     tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1) & "\"
50170    Else
50180     tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1)
50190   End If
50200  Next i
50210  .FilenameSubstitutions = tStr
50220  .FilenameSubstitutionsOnlyInTitle = Abs(Frm.chkFilenameSubst.Value)
50230  .JPEGColorscount = Frm.cmbJPEGColors.ListIndex
50240  .JPEGQuality = Frm.txtJPEGQuality.Text
50250  .NoConfirmMessageSwitchingDefaultprinter = Abs(Frm.chkNoConfirmMessageSwitchingDefaultprinter)
50260  .PCXColorscount = Frm.cmbPCXColors.ListIndex
50270  .PDFAllowAssembly = Abs(Frm.chkAllowAssembly.Value)
50280  .PDFAllowDegradedPrinting = Abs(Frm.chkAllowDegradedPrinting.Value)
50290  .PDFAllowFillIn = Abs(Frm.chkAllowFillIn.Value)
50300  .PDFAllowScreenReaders = Abs(Frm.chkAllowScreenReaders.Value)
50310  .PDFColorsCMYKToRGB = Abs(Frm.chkPDFCMYKtoRGB.Value)
50320  .PDFColorsColorModel = Frm.cmbPDFColorModel.ListIndex
50330  .PDFColorsPreserveHalftone = Abs(Frm.chkPDFPreserveHalftone.Value)
50340  .PDFColorsPreserveOverprint = Abs(Frm.chkPDFPreserveOverprint.Value)
50350  .PDFColorsPreserveTransfer = Abs(Frm.chkPDFPreserveTransfer.Value)
50360  .PDFCompressionColorCompression = Abs(Frm.chkPDFColorComp.Value)
50370  .PDFCompressionColorCompressionChoice = Frm.cmbPDFColorComp.ListIndex
50380  .PDFCompressionColorResample = Abs(Frm.chkPDFColorResample.Value)
50390  .PDFCompressionColorResampleChoice = Frm.cmbPDFColorResample.ListIndex
50400  .PDFCompressionColorResolution = Frm.txtPDFColorRes.Text
50410  .PDFCompressionGreyCompression = Abs(Frm.chkPDFGreyComp.Value)
50420  .PDFCompressionGreyCompressionChoice = Frm.cmbPDFGreyComp.ListIndex
50430  .PDFCompressionGreyResample = Abs(Frm.chkPDFGreyResample.Value)
50440  .PDFCompressionGreyResampleChoice = Frm.cmbPDFGreyResample.ListIndex
50450  .PDFCompressionGreyResolution = Frm.txtPDFGreyRes.Text
50460  .PDFCompressionMonoCompression = Abs(Frm.chkPDFMonoComp.Value)
50470  .PDFCompressionMonoCompressionChoice = Frm.cmbPDFMonoComp.ListIndex
50480  .PDFCompressionMonoResample = Abs(Frm.chkPDFMonoResample.Value)
50490  .PDFCompressionMonoResampleChoice = Frm.cmbPDFMonoResample.ListIndex
50500  .PDFCompressionMonoResolution = Frm.txtPDFMonoRes.Text
50510  .PDFCompressionTextCompression = Abs(Frm.chkPDFTextComp.Value)
50520  .PDFDisallowCopy = Abs(Frm.chkAllowCopy.Value)
50530  .PDFDisallowModifyAnnotations = Abs(Frm.chkAllowModifyAnnotations.Value)
50540  .PDFDisallowModifyContents = Abs(Frm.chkAllowModifyContents.Value)
50550  .PDFDisallowPrinting = Abs(Frm.chkAllowPrinting.Value)
50560  If Frm.cmbPDFEncryptor.ListIndex < 0 Then
50570    .PDFEncryptor = 0
50580   Else
50590    .PDFEncryptor = Frm.cmbPDFEncryptor.ItemData(Frm.cmbPDFEncryptor.ListIndex)
50600  End If
50610  .PDFFontsEmbedAll = Abs(Frm.chkPDFEmbedAll.Value)
50620  .PDFFontsSubSetFonts = Abs(Frm.chkPDFSubSetFonts.Value)
50630  .PDFFontsSubSetFontsPercent = Frm.txtPDFSubSetPerc.Text
50640  .PDFGeneralASCII85 = Abs(Frm.chkPDFASCII85.Value)
50650  .PDFGeneralAutorotate = Frm.cmbPDFRotate.ListIndex
50660  .PDFGeneralCompatibility = Frm.cmbPDFCompat.ListIndex
50670  .PDFGeneralOverprint = Frm.cmbPDFOverprint.ListIndex
50680  .PDFGeneralResolution = Frm.txtPDFRes.Text
50690  .PDFHighEncryption = Abs(Frm.optEncHigh.Value)
50700  .PDFLowEncryption = Abs(Frm.optEncLow.Value)
50710  .PDFOwnerPass = Abs(Frm.chkOwnerPass.Value)
50720  .PDFUserPass = Abs(Frm.chkUserPass.Value)
50730  .PDFUseSecurity = Abs(Frm.chkUseSecurity.Value)
50740  .PNGColorscount = Frm.cmbPNGColors.ListIndex
50750  .PrinterTemppath = Frm.txtTemppath.Text
50760  .ProcessPriority = Frm.sldProcessPriority.Value
50770  .ProgramFont = Frm.cmbFonts.List(Frm.cmbFonts.ListIndex)
50780  .ProgramFontCharset = Frm.cmbCharset.Text
50790  .ProgramFontSize = Frm.txtProgramFontsize.Text
50800  .PSLanguageLevel = Frm.cmbPSLanguageLevel.ListIndex
50810  .RemoveSpaces = Abs(Frm.chkSpaces.Value)
50820  .SaveFilename = Frm.txtSaveFilename.Text
50830  .StandardAuthor = Frm.txtStandardAuthor.Text
50840  .TIFFColorscount = Frm.cmbTIFFColors.ListIndex
50850  .UseAutosave = Abs(Frm.chkUseAutosave.Value)
50860  .UseAutosaveDirectory = Abs(Frm.chkUseAutosaveDirectory.Value)
50870  .UseCreationDateNow = Abs(Frm.chkUseCreationDateNow.Value)
50880  .UseStandardAuthor = Abs(Frm.chkUseStandardAuthor.Value)
50890  End With
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

