Attribute VB_Name = "modOptions"
Option Explicit

' Module automatically generated with LanguagesTool from Frank Heindörfer
' 2003
' Email: thesmilyface@users.sourceforge.net

Global Options As tOptions

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
 PCXColorscount As Long
 PDFAllowAssembly As Long
 PDFAllowCopy As Long
 PDFAllowDegradedPrinting As Long
 PDFAllowFillIn As Long
 PDFAllowModifyAnnotations As Long
 PDFAllowModifyContents As Long
 PDFAllowPrinting As Long
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

Public Function StandardOptions() As tOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim myOptions As tOptions
50020  With myOptions
50030   .AutosaveDirectory = vbNullString
50040   .AutosaveFilename = "<DateTime>"
50050   .AutosaveFormat = "0"
50060   .BitmapResolution = "150"
50070   .BMPColorscount = "1"
50080   .DirectoryGhostscriptBinaries = App.Path & "\"
50090   .DirectoryGhostscriptFonts = App.Path & "\fonts\"
50100   .DirectoryGhostscriptLibraries = App.Path & "\lib\"
50110   .EPSLanguageLevel = "2"
50120   .FilenameSubstitutions = " Microsoft Word - \.doc"
50130   .FilenameSubstitutionsOnlyInTitle = "1"
50140   .JPEGColorscount = "0"
50150   .JPEGQuality = "75"
50160   .Language = "english"
50170   .LastSaveDirectory = GetMyFiles
50180   .Logging = "0"
50190   .LogLines = "100"
50200   .PCXColorscount = "0"
50210   .PDFAllowAssembly = "0"
50220   .PDFAllowCopy = "0"
50230   .PDFAllowDegradedPrinting = "0"
50240   .PDFAllowFillIn = "0"
50250   .PDFAllowModifyAnnotations = "0"
50260   .PDFAllowModifyContents = "0"
50270   .PDFAllowPrinting = "0"
50280   .PDFAllowScreenReaders = "0"
50290   .PDFColorsCMYKToRGB = "1"
50300   .PDFColorsColorModel = "1"
50310   .PDFColorsPreserveHalftone = "0"
50320   .PDFColorsPreserveOverprint = "1"
50330   .PDFColorsPreserveTransfer = "1"
50340   .PDFCompressionColorCompression = "1"
50350   .PDFCompressionColorCompressionChoice = "0"
50360   .PDFCompressionColorResample = "0"
50370   .PDFCompressionColorResampleChoice = "0"
50380   .PDFCompressionColorResolution = "300"
50390   .PDFCompressionGreyCompression = "1"
50400   .PDFCompressionGreyCompressionChoice = "0"
50410   .PDFCompressionGreyResample = "0"
50420   .PDFCompressionGreyResampleChoice = "0"
50430   .PDFCompressionGreyResolution = "300"
50440   .PDFCompressionMonoCompression = "1"
50450   .PDFCompressionMonoCompressionChoice = "0"
50460   .PDFCompressionMonoResample = "0"
50470   .PDFCompressionMonoResampleChoice = "0"
50480   .PDFCompressionMonoResolution = "1200"
50490   .PDFCompressionTextCompression = "1"
50500   .PDFFontsEmbedAll = "1"
50510   .PDFFontsSubSetFonts = "1"
50520   .PDFFontsSubSetFontsPercent = "100"
50530   .PDFGeneralASCII85 = "0"
50540   .PDFGeneralAutorotate = "0"
50550   .PDFGeneralCompatibility = "1"
50560   .PDFGeneralOverprint = "0"
50570   .PDFGeneralResolution = "600"
50580   .PDFHighEncryption = "0"
50590   .PDFLowEncryption = "1"
50600   .PDFOwnerPass = "0"
50610   .PDFUserPass = "0"
50620   .PDFUseSecurity = "0"
50630   .PNGColorscount = "0"
50640   .PrinterStop = "0"
50650   .PrinterTemppath = GetPDFCreatorTempfolder
50660   .ProcessPriority = "1"
50670   .ProgramFont = "MS Sans Serif"
50680   .ProgramFontCharset = "0"
50690   .ProgramFontSize = "8"
50700   .PSLanguageLevel = "2"
50710   .RemoveSpaces = "1"
50720   .SaveFilename = "<Title>"
50730   .StandardAuthor = vbNullString
50740   .StartStandardProgram = "1"
50750   .TIFFColorscount = "0"
50760   .UseAutosave = "0"
50770   .UseAutosaveDirectory = "1"
50780   .UseCreationDateNow = "0"
50790   .UseStandardAuthor = "0"
50800  End With
50810  StandardOptions = myOptions
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
50030  ini.FileName = PDFCreatorINIFile
50040  ini.Section = "Options"
50050  If ini.CheckIniFile = False Then
50060   ReadOptions = StandardOptions
50070   Exit Function
50080  End If
50090  ReadINISection PDFCreatorINIFile, "Options", hOpt
50100  With myOptions
50110   tStr = hOpt.Retrieve("AutosaveDirectory", GetMyFiles)
50120   If Len(Dir(tStr, vbDirectory)) > 0 And Len(tStr) > 0 Then
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
50380   If Len(Dir(tStr, vbDirectory)) > 0 And Len(tStr) > 0 Then
50390     .DirectoryGhostscriptBinaries = CompletePath(tStr)
50400    Else
50410     .DirectoryGhostscriptBinaries = App.Path & "\"
50420   End If
50430   tStr = hOpt.Retrieve("DirectoryGhostscriptFonts", App.Path & "\fonts")
50440   If Len(Dir(tStr, vbDirectory)) > 0 And Len(tStr) > 0 Then
50450     .DirectoryGhostscriptFonts = CompletePath(tStr)
50460    Else
50470     .DirectoryGhostscriptFonts = App.Path & "\fonts\"
50480   End If
50490   tStr = hOpt.Retrieve("DirectoryGhostscriptLibraries", App.Path & "\lib")
50500   If Len(Dir(tStr, vbDirectory)) > 0 And Len(tStr) > 0 Then
50510     .DirectoryGhostscriptLibraries = CompletePath(tStr)
50520    Else
50530     .DirectoryGhostscriptLibraries = App.Path & "\lib\"
50540   End If
50550   tStr = hOpt.Retrieve("EPSLanguageLevel")
50560   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
50570     .EPSLanguageLevel = CLng(tStr)
50580    Else
50590     .EPSLanguageLevel = 2
50600   End If
50610   tStr = hOpt.Retrieve("FilenameSubstitutions", " Microsoft Word - \.doc")
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
50840   If Len(Dir(tStr, vbDirectory)) > 0 And Len(tStr) > 0 Then
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
51010   tStr = hOpt.Retrieve("PCXColorscount")
51020   If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
51030     .PCXColorscount = CLng(tStr)
51040    Else
51050     .PCXColorscount = 0
51060   End If
51070   tStr = hOpt.Retrieve("PDFAllowAssembly")
51080   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51090     .PDFAllowAssembly = CLng(tStr)
51100    Else
51110     .PDFAllowAssembly = 0
51120   End If
51130   tStr = hOpt.Retrieve("PDFAllowCopy")
51140   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51150     .PDFAllowCopy = CLng(tStr)
51160    Else
51170     .PDFAllowCopy = 0
51180   End If
51190   tStr = hOpt.Retrieve("PDFAllowDegradedPrinting")
51200   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51210     .PDFAllowDegradedPrinting = CLng(tStr)
51220    Else
51230     .PDFAllowDegradedPrinting = 0
51240   End If
51250   tStr = hOpt.Retrieve("PDFAllowFillIn")
51260   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51270     .PDFAllowFillIn = CLng(tStr)
51280    Else
51290     .PDFAllowFillIn = 0
51300   End If
51310   tStr = hOpt.Retrieve("PDFAllowModifyAnnotations")
51320   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51330     .PDFAllowModifyAnnotations = CLng(tStr)
51340    Else
51350     .PDFAllowModifyAnnotations = 0
51360   End If
51370   tStr = hOpt.Retrieve("PDFAllowModifyContents")
51380   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51390     .PDFAllowModifyContents = CLng(tStr)
51400    Else
51410     .PDFAllowModifyContents = 0
51420   End If
51430   tStr = hOpt.Retrieve("PDFAllowPrinting")
51440   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51450     .PDFAllowPrinting = CLng(tStr)
51460    Else
51470     .PDFAllowPrinting = 0
51480   End If
51490   tStr = hOpt.Retrieve("PDFAllowScreenReaders")
51500   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51510     .PDFAllowScreenReaders = CLng(tStr)
51520    Else
51530     .PDFAllowScreenReaders = 0
51540   End If
51550   tStr = hOpt.Retrieve("PDFColorsCMYKToRGB")
51560   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51570     .PDFColorsCMYKToRGB = CLng(tStr)
51580    Else
51590     .PDFColorsCMYKToRGB = 1
51600   End If
51610   tStr = hOpt.Retrieve("PDFColorsColorModel")
51620   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
51630     .PDFColorsColorModel = CLng(tStr)
51640    Else
51650     .PDFColorsColorModel = 1
51660   End If
51670   tStr = hOpt.Retrieve("PDFColorsPreserveHalftone")
51680   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51690     .PDFColorsPreserveHalftone = CLng(tStr)
51700    Else
51710     .PDFColorsPreserveHalftone = 0
51720   End If
51730   tStr = hOpt.Retrieve("PDFColorsPreserveOverprint")
51740   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51750     .PDFColorsPreserveOverprint = CLng(tStr)
51760    Else
51770     .PDFColorsPreserveOverprint = 1
51780   End If
51790   tStr = hOpt.Retrieve("PDFColorsPreserveTransfer")
51800   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51810     .PDFColorsPreserveTransfer = CLng(tStr)
51820    Else
51830     .PDFColorsPreserveTransfer = 1
51840   End If
51850   tStr = hOpt.Retrieve("PDFCompressionColorCompression")
51860   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51870     .PDFCompressionColorCompression = CLng(tStr)
51880    Else
51890     .PDFCompressionColorCompression = 1
51900   End If
51910   tStr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
51920   If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
51930     .PDFCompressionColorCompressionChoice = CLng(tStr)
51940    Else
51950     .PDFCompressionColorCompressionChoice = 0
51960   End If
51970   tStr = hOpt.Retrieve("PDFCompressionColorResample")
51980   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
51990     .PDFCompressionColorResample = CLng(tStr)
52000    Else
52010     .PDFCompressionColorResample = 0
52020   End If
52030   tStr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
52040   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
52050     .PDFCompressionColorResampleChoice = CLng(tStr)
52060    Else
52070     .PDFCompressionColorResampleChoice = 0
52080   End If
52090   tStr = hOpt.Retrieve("PDFCompressionColorResolution")
52100   If CLng(tStr) >= 0 Then
52110     .PDFCompressionColorResolution = CLng(tStr)
52120    Else
52130     .PDFCompressionColorResolution = 300
52140   End If
52150   tStr = hOpt.Retrieve("PDFCompressionGreyCompression")
52160   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52170     .PDFCompressionGreyCompression = CLng(tStr)
52180    Else
52190     .PDFCompressionGreyCompression = 1
52200   End If
52210   tStr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
52220   If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
52230     .PDFCompressionGreyCompressionChoice = CLng(tStr)
52240    Else
52250     .PDFCompressionGreyCompressionChoice = 0
52260   End If
52270   tStr = hOpt.Retrieve("PDFCompressionGreyResample")
52280   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52290     .PDFCompressionGreyResample = CLng(tStr)
52300    Else
52310     .PDFCompressionGreyResample = 0
52320   End If
52330   tStr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
52340   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
52350     .PDFCompressionGreyResampleChoice = CLng(tStr)
52360    Else
52370     .PDFCompressionGreyResampleChoice = 0
52380   End If
52390   tStr = hOpt.Retrieve("PDFCompressionGreyResolution")
52400   If CLng(tStr) >= 0 Then
52410     .PDFCompressionGreyResolution = CLng(tStr)
52420    Else
52430     .PDFCompressionGreyResolution = 300
52440   End If
52450   tStr = hOpt.Retrieve("PDFCompressionMonoCompression")
52460   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52470     .PDFCompressionMonoCompression = CLng(tStr)
52480    Else
52490     .PDFCompressionMonoCompression = 1
52500   End If
52510   tStr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
52520   If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
52530     .PDFCompressionMonoCompressionChoice = CLng(tStr)
52540    Else
52550     .PDFCompressionMonoCompressionChoice = 0
52560   End If
52570   tStr = hOpt.Retrieve("PDFCompressionMonoResample")
52580   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52590     .PDFCompressionMonoResample = CLng(tStr)
52600    Else
52610     .PDFCompressionMonoResample = 0
52620   End If
52630   tStr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
52640   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
52650     .PDFCompressionMonoResampleChoice = CLng(tStr)
52660    Else
52670     .PDFCompressionMonoResampleChoice = 0
52680   End If
52690   tStr = hOpt.Retrieve("PDFCompressionMonoResolution")
52700   If CLng(tStr) >= 0 Then
52710     .PDFCompressionMonoResolution = CLng(tStr)
52720    Else
52730     .PDFCompressionMonoResolution = 1200
52740   End If
52750   tStr = hOpt.Retrieve("PDFCompressionTextCompression")
52760   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52770     .PDFCompressionTextCompression = CLng(tStr)
52780    Else
52790     .PDFCompressionTextCompression = 1
52800   End If
52810   tStr = hOpt.Retrieve("PDFFontsEmbedAll")
52820   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52830     .PDFFontsEmbedAll = CLng(tStr)
52840    Else
52850     .PDFFontsEmbedAll = 1
52860   End If
52870   tStr = hOpt.Retrieve("PDFFontsSubSetFonts")
52880   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
52890     .PDFFontsSubSetFonts = CLng(tStr)
52900    Else
52910     .PDFFontsSubSetFonts = 1
52920   End If
52930   tStr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
52940   If CLng(tStr) >= 0 Then
52950     .PDFFontsSubSetFontsPercent = CLng(tStr)
52960    Else
52970     .PDFFontsSubSetFontsPercent = 100
52980   End If
52990   tStr = hOpt.Retrieve("PDFGeneralASCII85")
53000   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53010     .PDFGeneralASCII85 = CLng(tStr)
53020    Else
53030     .PDFGeneralASCII85 = 0
53040   End If
53050   tStr = hOpt.Retrieve("PDFGeneralAutorotate")
53060   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
53070     .PDFGeneralAutorotate = CLng(tStr)
53080    Else
53090     .PDFGeneralAutorotate = 0
53100   End If
53110   tStr = hOpt.Retrieve("PDFGeneralCompatibility")
53120   If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
53130     .PDFGeneralCompatibility = CLng(tStr)
53140    Else
53150     .PDFGeneralCompatibility = 1
53160   End If
53170   tStr = hOpt.Retrieve("PDFGeneralOverprint")
53180   If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
53190     .PDFGeneralOverprint = CLng(tStr)
53200    Else
53210     .PDFGeneralOverprint = 0
53220   End If
53230   tStr = hOpt.Retrieve("PDFGeneralResolution")
53240   If CLng(tStr) >= 0 Then
53250     .PDFGeneralResolution = CLng(tStr)
53260    Else
53270     .PDFGeneralResolution = 600
53280   End If
53290   tStr = hOpt.Retrieve("PDFHighEncryption")
53300   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53310     .PDFHighEncryption = CLng(tStr)
53320    Else
53330     .PDFHighEncryption = 0
53340   End If
53350   tStr = hOpt.Retrieve("PDFLowEncryption")
53360   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53370     .PDFLowEncryption = CLng(tStr)
53380    Else
53390     .PDFLowEncryption = 1
53400   End If
53410   tStr = hOpt.Retrieve("PDFOwnerPass")
53420   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53430     .PDFOwnerPass = CLng(tStr)
53440    Else
53450     .PDFOwnerPass = 0
53460   End If
53470   tStr = hOpt.Retrieve("PDFUserPass")
53480   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53490     .PDFUserPass = CLng(tStr)
53500    Else
53510     .PDFUserPass = 0
53520   End If
53530   tStr = hOpt.Retrieve("PDFUseSecurity")
53540   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53550     .PDFUseSecurity = CLng(tStr)
53560    Else
53570     .PDFUseSecurity = 0
53580   End If
53590   tStr = hOpt.Retrieve("PNGColorscount")
53600   If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
53610     .PNGColorscount = CLng(tStr)
53620    Else
53630     .PNGColorscount = 0
53640   End If
53650   tStr = hOpt.Retrieve("PrinterStop")
53660   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
53670     .PrinterStop = CLng(tStr)
53680    Else
53690     .PrinterStop = 0
53700   End If
53710   tStr = hOpt.Retrieve("PrinterTemppath", GetPDFCreatorTempfolder)
53720   If Len(Dir(tStr, vbDirectory)) > 0 And Len(tStr) > 0 Then
53730     .PrinterTemppath = CompletePath(tStr)
53740    Else
53750     .PrinterTemppath = GetPDFCreatorTempfolder
53760   End If
53770   tStr = hOpt.Retrieve("ProcessPriority")
53780   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
53790     .ProcessPriority = CLng(tStr)
53800    Else
53810     .ProcessPriority = 1
53820   End If
53830   tStr = hOpt.Retrieve("ProgramFont", "MS Sans Serif")
53840   .ProgramFont = tStr
53850   tStr = hOpt.Retrieve("ProgramFontCharset")
53860   If CLng(tStr) >= 0 Then
53870     .ProgramFontCharset = CLng(tStr)
53880    Else
53890     .ProgramFontCharset = 0
53900   End If
53910   tStr = hOpt.Retrieve("ProgramFontSize")
53920   If CLng(tStr) >= 1 And CLng(tStr) <= 72 Then
53930     .ProgramFontSize = CLng(tStr)
53940    Else
53950     .ProgramFontSize = 8
53960   End If
53970   tStr = hOpt.Retrieve("PSLanguageLevel")
53980   If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
53990     .PSLanguageLevel = CLng(tStr)
54000    Else
54010     .PSLanguageLevel = 2
54020   End If
54030   tStr = hOpt.Retrieve("RemoveSpaces")
54040   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54050     .RemoveSpaces = CLng(tStr)
54060    Else
54070     .RemoveSpaces = 1
54080   End If
54090   tStr = hOpt.Retrieve("SaveFilename", "<Title>")
54100   .SaveFilename = tStr
54110   tStr = hOpt.Retrieve("StandardAuthor", "")
54120   .StandardAuthor = tStr
54130   tStr = hOpt.Retrieve("StartStandardProgram")
54140   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54150     .StartStandardProgram = CLng(tStr)
54160    Else
54170     .StartStandardProgram = 1
54180   End If
54190   tStr = hOpt.Retrieve("TIFFColorscount")
54200   If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
54210     .TIFFColorscount = CLng(tStr)
54220    Else
54230     .TIFFColorscount = 0
54240   End If
54250   tStr = hOpt.Retrieve("UseAutosave")
54260   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54270     .UseAutosave = CLng(tStr)
54280    Else
54290     .UseAutosave = 0
54300   End If
54310   tStr = hOpt.Retrieve("UseAutosaveDirectory")
54320   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54330     .UseAutosaveDirectory = CLng(tStr)
54340    Else
54350     .UseAutosaveDirectory = 1
54360   End If
54370   tStr = hOpt.Retrieve("UseCreationDateNow")
54380   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54390     .UseCreationDateNow = CLng(tStr)
54400    Else
54410     .UseCreationDateNow = 0
54420   End If
54430   tStr = hOpt.Retrieve("UseStandardAuthor")
54440   If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
54450     .UseStandardAuthor = CLng(tStr)
54460    Else
54470     .UseStandardAuthor = 0
54480   End If
54490  End With
54500  Set ini = Nothing
54510  ReadOptions = myOptions
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
50030  ini.FileName = PDFCreatorINIFile
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
50190   ini.SaveKey CStr(.FilenameSubstitutionsOnlyInTitle), "FilenameSubstitutionsOnlyInTitle"
50200   ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
50210   ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
50220   ini.SaveKey CStr(.Language), "Language"
50230   ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
50240   ini.SaveKey CStr(.Logging), "Logging"
50250   ini.SaveKey CStr(.LogLines), "LogLines"
50260   ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
50270   ini.SaveKey CStr(.PDFAllowAssembly), "PDFAllowAssembly"
50280   ini.SaveKey CStr(.PDFAllowCopy), "PDFAllowCopy"
50290   ini.SaveKey CStr(.PDFAllowDegradedPrinting), "PDFAllowDegradedPrinting"
50300   ini.SaveKey CStr(.PDFAllowFillIn), "PDFAllowFillIn"
50310   ini.SaveKey CStr(.PDFAllowModifyAnnotations), "PDFAllowModifyAnnotations"
50320   ini.SaveKey CStr(.PDFAllowModifyContents), "PDFAllowModifyContents"
50330   ini.SaveKey CStr(.PDFAllowPrinting), "PDFAllowPrinting"
50340   ini.SaveKey CStr(.PDFAllowScreenReaders), "PDFAllowScreenReaders"
50350   ini.SaveKey CStr(.PDFColorsCMYKToRGB), "PDFColorsCMYKToRGB"
50360   ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
50370   ini.SaveKey CStr(.PDFColorsPreserveHalftone), "PDFColorsPreserveHalftone"
50380   ini.SaveKey CStr(.PDFColorsPreserveOverprint), "PDFColorsPreserveOverprint"
50390   ini.SaveKey CStr(.PDFColorsPreserveTransfer), "PDFColorsPreserveTransfer"
50400   ini.SaveKey CStr(.PDFCompressionColorCompression), "PDFCompressionColorCompression"
50410   ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
50420   ini.SaveKey CStr(.PDFCompressionColorResample), "PDFCompressionColorResample"
50430   ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
50440   ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
50450   ini.SaveKey CStr(.PDFCompressionGreyCompression), "PDFCompressionGreyCompression"
50460   ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
50470   ini.SaveKey CStr(.PDFCompressionGreyResample), "PDFCompressionGreyResample"
50480   ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
50490   ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
50500   ini.SaveKey CStr(.PDFCompressionMonoCompression), "PDFCompressionMonoCompression"
50510   ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
50520   ini.SaveKey CStr(.PDFCompressionMonoResample), "PDFCompressionMonoResample"
50530   ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
50540   ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
50550   ini.SaveKey CStr(.PDFCompressionTextCompression), "PDFCompressionTextCompression"
50560   ini.SaveKey CStr(.PDFFontsEmbedAll), "PDFFontsEmbedAll"
50570   ini.SaveKey CStr(.PDFFontsSubSetFonts), "PDFFontsSubSetFonts"
50580   ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
50590   ini.SaveKey CStr(.PDFGeneralASCII85), "PDFGeneralASCII85"
50600   ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
50610   ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
50620   ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
50630   ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
50640   ini.SaveKey CStr(.PDFHighEncryption), "PDFHighEncryption"
50650   ini.SaveKey CStr(.PDFLowEncryption), "PDFLowEncryption"
50660   ini.SaveKey CStr(.PDFOwnerPass), "PDFOwnerPass"
50670   ini.SaveKey CStr(.PDFUserPass), "PDFUserPass"
50680   ini.SaveKey CStr(.PDFUseSecurity), "PDFUseSecurity"
50690   ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
50700   ini.SaveKey CStr(.PrinterStop), "PrinterStop"
50710   ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
50720   ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
50730   ini.SaveKey CStr(.ProgramFont), "ProgramFont"
50740   ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
50750   ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
50760   ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
50770   ini.SaveKey CStr(.RemoveSpaces), "RemoveSpaces"
50780   ini.SaveKey CStr(.SaveFilename), "SaveFilename"
50790   ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
50800   ini.SaveKey CStr(.StartStandardProgram), "StartStandardProgram"
50810   ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
50820   ini.SaveKey CStr(.UseAutosave), "UseAutosave"
50830   ini.SaveKey CStr(.UseAutosaveDirectory), "UseAutosaveDirectory"
50840   ini.SaveKey CStr(.UseCreationDateNow), "UseCreationDateNow"
50850   ini.SaveKey CStr(.UseStandardAuthor), "UseStandardAuthor"
50860  End With
50870  Set ini = Nothing
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50020  Dim i As Long, tList() As String, tStrA() As String, lsv As ListView
50030  With sOptions
50040   Frm.txtAutosaveDirectory.Text = .AutosaveDirectory
50050   Frm.txtAutosaveFilename.Text = .AutosaveFilename
50060   Frm.cmbAutosaveFormat.ListIndex = .AutosaveFormat
50070   Frm.txtBitmapResolution.Text = .BitmapResolution
50080   Frm.cmbBMPColors.ListIndex = .BMPColorscount
50090   Frm.txtGSbin.Text = .DirectoryGhostscriptBinaries
50100   Frm.txtGSfonts.Text = .DirectoryGhostscriptFonts
50110   Frm.txtGSlib.Text = .DirectoryGhostscriptLibraries
50120   Frm.cmbEPSLanguageLevel.ListIndex = .EPSLanguageLevel
50130   Set lsv = Frm.lsvFilenameSubst
50140   tList = Split(.FilenameSubstitutions, "\")
50150   For i = 0 To UBound(tList)
50160    If InStr(tList(i), "|") <= 0 Then
50170     tList(i) = tList(i) & "|"
50180    End If
50190    If UBound(Split(tList(i), "|")) = 1 Then
50200     tStrA = Split(tList(i), "|")
50210     lsv.ListItems.Add , , tStrA(0)
50220     lsv.ListItems(lsv.ListItems.Count).SubItems(1) = tStrA(1)
50230    End If
50240   Next i
50250   If lsv.ListItems.Count > 0 Then
50260    lsv.ListItems(1).Selected = True
50270    Frm.txtFilenameSubst(0).Text = lsv.ListItems(1).Text
50280    Frm.txtFilenameSubst(0).ToolTipText = Frm.txtFilenameSubst(0).Text
50290    Frm.txtFilenameSubst(1).Text = lsv.ListItems(1).SubItems(1)
50300    Frm.txtFilenameSubst(1).ToolTipText = Frm.txtFilenameSubst(1).Text
50310   End If
50320   Frm.chkFilenameSubst.Value = .FilenameSubstitutionsOnlyInTitle
50330   Frm.cmbJPEGColors.ListIndex = .JPEGColorscount
50340   Frm.txtJPEGQuality.Text = .JPEGQuality
50350   Frm.cmbPCXColors.ListIndex = .PCXColorscount
50360   Frm.chkAllowAssembly.Value = .PDFAllowAssembly
50370   Frm.chkAllowCopy.Value = .PDFAllowCopy
50380   Frm.chkAllowDegradedPrinting.Value = .PDFAllowDegradedPrinting
50390   Frm.chkAllowFillIn.Value = .PDFAllowFillIn
50400   Frm.chkAllowModifyAnnotations.Value = .PDFAllowModifyAnnotations
50410   Frm.chkAllowModifyContents.Value = .PDFAllowModifyContents
50420   Frm.chkAllowPrinting.Value = .PDFAllowPrinting
50430   Frm.chkAllowScreenReaders.Value = .PDFAllowScreenReaders
50440   Frm.chkPDFCMYKtoRGB.Value = .PDFColorsCMYKToRGB
50450   Frm.cmbPDFColorModel.ListIndex = .PDFColorsColorModel
50460   Frm.chkPDFPreserveHalftone.Value = .PDFColorsPreserveHalftone
50470   Frm.chkPDFPreserveOverprint.Value = .PDFColorsPreserveOverprint
50480   Frm.chkPDFPreserveTransfer.Value = .PDFColorsPreserveTransfer
50490   Frm.chkPDFColorComp.Value = .PDFCompressionColorCompression
50500   Frm.cmbPDFColorComp.ListIndex = .PDFCompressionColorCompressionChoice
50510   Frm.chkPDFColorResample.Value = .PDFCompressionColorResample
50520   Frm.cmbPDFColorResample.ListIndex = .PDFCompressionColorResampleChoice
50530   Frm.txtPDFColorRes.Text = .PDFCompressionColorResolution
50540   Frm.chkPDFGreyComp.Value = .PDFCompressionGreyCompression
50550   Frm.cmbPDFGreyComp.ListIndex = .PDFCompressionGreyCompressionChoice
50560   Frm.chkPDFGreyResample.Value = .PDFCompressionGreyResample
50570   Frm.cmbPDFGreyResample.ListIndex = .PDFCompressionGreyResampleChoice
50580   Frm.txtPDFGreyRes.Text = .PDFCompressionGreyResolution
50590   Frm.chkPDFMonoComp.Value = .PDFCompressionMonoCompression
50600   Frm.cmbPDFMonoComp.ListIndex = .PDFCompressionMonoCompressionChoice
50610   Frm.chkPDFMonoResample.Value = .PDFCompressionMonoResample
50620   Frm.cmbPDFMonoResample.ListIndex = .PDFCompressionMonoResampleChoice
50630   Frm.txtPDFMonoRes.Text = .PDFCompressionMonoResolution
50640   Frm.chkPDFTextComp.Value = .PDFCompressionTextCompression
50650   Frm.chkPDFEmbedAll.Value = .PDFFontsEmbedAll
50660   Frm.chkPDFSubSetFonts.Value = .PDFFontsSubSetFonts
50670   Frm.txtPDFSubSetPerc.Text = .PDFFontsSubSetFontsPercent
50680   Frm.chkPDFASCII85.Value = .PDFGeneralASCII85
50690   Frm.cmbPDFRotate.ListIndex = .PDFGeneralAutorotate
50700   Frm.cmbPDFCompat.ListIndex = .PDFGeneralCompatibility
50710   Frm.cmbPDFOverprint.ListIndex = .PDFGeneralOverprint
50720   Frm.txtPDFRes.Text = .PDFGeneralResolution
50730   Frm.optEncHigh.Value = .PDFHighEncryption
50740   Frm.optEncLow.Value = .PDFLowEncryption
50750   Frm.chkOwnerPass.Value = .PDFOwnerPass
50760   Frm.chkUserPass.Value = .PDFUserPass
50770   Frm.chkUseSecurity.Value = .PDFUseSecurity
50780   Frm.cmbPNGColors.ListIndex = .PNGColorscount
50790   Frm.txtTemppath.Text = .PrinterTemppath
50800   Frm.sldProcessPriority.Value = .ProcessPriority
50810   For i = 0 To Frm.cmbFonts.ListCount - 1
50820     If UCase$(Frm.cmbFonts.List(i)) = UCase$(.ProgramFont) Then
50830      Frm.cmbFonts.ListIndex = i
50840      Exit For
50850     End If
50860   Next i
50870   Frm.cmbCharset.Text = .ProgramFontCharset
50880   Frm.txtProgramFontsize.Text = .ProgramFontSize
50890   Frm.cmbPSLanguageLevel.ListIndex = .PSLanguageLevel
50900   Frm.chkSpaces.Value = .RemoveSpaces
50910   Frm.txtSaveFilename.Text = .SaveFilename
50920   Frm.txtStandardAuthor.Text = .StandardAuthor
50930   Frm.cmbTIFFColors.ListIndex = .TIFFColorscount
50940   Frm.chkUseAutosave.Value = .UseAutosave
50950   Frm.chkUseAutosaveDirectory.Value = .UseAutosaveDirectory
50960   Frm.chkUseCreationDateNow.Value = .UseCreationDateNow
50970   Frm.chkUseStandardAuthor.Value = .UseStandardAuthor
50980  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions", "ShowOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub GetOptions(Frm As Form, sOptions As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, tStr As String, lsv As ListView
50020  With sOptions
50030   .AutosaveDirectory = Frm.txtAutosaveDirectory.Text
50040   .AutosaveFilename = Frm.txtAutosaveFilename.Text
50050   .AutosaveFormat = Frm.cmbAutosaveFormat.ListIndex
50060   .BitmapResolution = Frm.txtBitmapResolution.Text
50070   .BMPColorscount = Frm.cmbBMPColors.ListIndex
50080   .DirectoryGhostscriptBinaries = Frm.txtGSbin.Text
50090   .DirectoryGhostscriptFonts = Frm.txtGSfonts.Text
50100   .DirectoryGhostscriptLibraries = Frm.txtGSlib.Text
50110   .EPSLanguageLevel = Frm.cmbEPSLanguageLevel.ListIndex
50120   tStr = ""
50130   Set lsv = Frm.lsvFilenameSubst
50140   For i = 1 To lsv.ListItems.Count
50150    If i < lsv.ListItems.Count Then
50160      tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1) & "\"
50170     Else
50180      tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1)
50190    End If
50200   Next i
50210   .FilenameSubstitutions = tStr
50220   .FilenameSubstitutionsOnlyInTitle = Frm.chkFilenameSubst.Value
50230   .JPEGColorscount = Frm.cmbJPEGColors.ListIndex
50240   .JPEGQuality = Frm.txtJPEGQuality.Text
50250   .PCXColorscount = Frm.cmbPCXColors.ListIndex
50260   .PDFAllowAssembly = Frm.chkAllowAssembly.Value
50270   .PDFAllowCopy = Frm.chkAllowCopy.Value
50280   .PDFAllowDegradedPrinting = Frm.chkAllowDegradedPrinting.Value
50290   .PDFAllowFillIn = Frm.chkAllowFillIn.Value
50300   .PDFAllowModifyAnnotations = Frm.chkAllowModifyAnnotations.Value
50310   .PDFAllowModifyContents = Frm.chkAllowModifyContents.Value
50320   .PDFAllowPrinting = Frm.chkAllowPrinting.Value
50330   .PDFAllowScreenReaders = Frm.chkAllowScreenReaders.Value
50340   .PDFColorsCMYKToRGB = Frm.chkPDFCMYKtoRGB.Value
50350   .PDFColorsColorModel = Frm.cmbPDFColorModel.ListIndex
50360   .PDFColorsPreserveHalftone = Frm.chkPDFPreserveHalftone.Value
50370   .PDFColorsPreserveOverprint = Frm.chkPDFPreserveOverprint.Value
50380   .PDFColorsPreserveTransfer = Frm.chkPDFPreserveTransfer.Value
50390   .PDFCompressionColorCompression = Frm.chkPDFColorComp.Value
50400   .PDFCompressionColorCompressionChoice = Frm.cmbPDFColorComp.ListIndex
50410   .PDFCompressionColorResample = Frm.chkPDFColorResample.Value
50420   .PDFCompressionColorResampleChoice = Frm.cmbPDFColorResample.ListIndex
50430   .PDFCompressionColorResolution = Frm.txtPDFColorRes.Text
50440   .PDFCompressionGreyCompression = Frm.chkPDFGreyComp.Value
50450   .PDFCompressionGreyCompressionChoice = Frm.cmbPDFGreyComp.ListIndex
50460   .PDFCompressionGreyResample = Frm.chkPDFGreyResample.Value
50470   .PDFCompressionGreyResampleChoice = Frm.cmbPDFGreyResample.ListIndex
50480   .PDFCompressionGreyResolution = Frm.txtPDFGreyRes.Text
50490   .PDFCompressionMonoCompression = Frm.chkPDFMonoComp.Value
50500   .PDFCompressionMonoCompressionChoice = Frm.cmbPDFMonoComp.ListIndex
50510   .PDFCompressionMonoResample = Frm.chkPDFMonoResample.Value
50520   .PDFCompressionMonoResampleChoice = Frm.cmbPDFMonoResample.ListIndex
50530   .PDFCompressionMonoResolution = Frm.txtPDFMonoRes.Text
50540   .PDFCompressionTextCompression = Frm.chkPDFTextComp.Value
50550   .PDFFontsEmbedAll = Frm.chkPDFEmbedAll.Value
50560   .PDFFontsSubSetFonts = Frm.chkPDFSubSetFonts.Value
50570   .PDFFontsSubSetFontsPercent = Frm.txtPDFSubSetPerc.Text
50580   .PDFGeneralASCII85 = Frm.chkPDFASCII85.Value
50590   .PDFGeneralAutorotate = Frm.cmbPDFRotate.ListIndex
50600   .PDFGeneralCompatibility = Frm.cmbPDFCompat.ListIndex
50610   .PDFGeneralOverprint = Frm.cmbPDFOverprint.ListIndex
50620   .PDFGeneralResolution = Frm.txtPDFRes.Text
50630   .PDFHighEncryption = Frm.optEncHigh.Value
50640   .PDFLowEncryption = Frm.optEncLow.Value
50650   .PDFOwnerPass = Frm.chkOwnerPass.Value
50660   .PDFUserPass = Frm.chkUserPass.Value
50670   .PDFUseSecurity = Frm.chkUseSecurity.Value
50680   .PNGColorscount = Frm.cmbPNGColors.ListIndex
50690   .PrinterTemppath = Frm.txtTemppath.Text
50700   .ProcessPriority = Frm.sldProcessPriority.Value
50710   .ProgramFont = Frm.cmbFonts.List(Frm.cmbFonts.ListIndex)
50720   .ProgramFontCharset = Frm.cmbCharset.Text
50730   .ProgramFontSize = Frm.txtProgramFontsize.Text
50740   .PSLanguageLevel = Frm.cmbPSLanguageLevel.ListIndex
50750   .RemoveSpaces = Frm.chkSpaces.Value
50760   .SaveFilename = Frm.txtSaveFilename.Text
50770   .StandardAuthor = Frm.txtStandardAuthor.Text
50780   .TIFFColorscount = Frm.cmbTIFFColors.ListIndex
50790   .UseAutosave = Frm.chkUseAutosave.Value
50800   .UseAutosaveDirectory = Frm.chkUseAutosaveDirectory.Value
50810   .UseCreationDateNow = Frm.chkUseCreationDateNow.Value
50820   .UseStandardAuthor = Frm.chkUseStandardAuthor.Value
50830  End With
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

