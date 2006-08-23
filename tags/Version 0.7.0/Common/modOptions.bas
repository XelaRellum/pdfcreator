Attribute VB_Name = "modOptions"
Option Explicit

' Module automatically generated with LanguagesTool from Frank Heind�rfer
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
 Dim myOptions As tOptions
 With myOptions
  .AutosaveDirectory = vbNullString
  .AutosaveFilename = "<DateTime>"
  .AutosaveFormat = "0"
  .BitmapResolution = "150"
  .BMPColorscount = "1"
  .DirectoryGhostscriptBinaries = App.Path & "\"
  .DirectoryGhostscriptFonts = App.Path & "\fonts\"
  .DirectoryGhostscriptLibraries = App.Path & "\lib\"
  .EPSLanguageLevel = "2"
  .FilenameSubstitutions = " Microsoft Word - \.doc"
  .FilenameSubstitutionsOnlyInTitle = "1"
  .JPEGColorscount = "0"
  .JPEGQuality = "75"
  .Language = "english"
  .LastSaveDirectory = GetMyFiles
  .Logging = "0"
  .LogLines = "100"
  .PCXColorscount = "0"
  .PDFAllowAssembly = "0"
  .PDFAllowCopy = "0"
  .PDFAllowDegradedPrinting = "0"
  .PDFAllowFillIn = "0"
  .PDFAllowModifyAnnotations = "0"
  .PDFAllowModifyContents = "0"
  .PDFAllowPrinting = "0"
  .PDFAllowScreenReaders = "0"
  .PDFColorsCMYKToRGB = "1"
  .PDFColorsColorModel = "1"
  .PDFColorsPreserveHalftone = "0"
  .PDFColorsPreserveOverprint = "1"
  .PDFColorsPreserveTransfer = "1"
  .PDFCompressionColorCompression = "1"
  .PDFCompressionColorCompressionChoice = "0"
  .PDFCompressionColorResample = "0"
  .PDFCompressionColorResampleChoice = "0"
  .PDFCompressionColorResolution = "300"
  .PDFCompressionGreyCompression = "1"
  .PDFCompressionGreyCompressionChoice = "0"
  .PDFCompressionGreyResample = "0"
  .PDFCompressionGreyResampleChoice = "0"
  .PDFCompressionGreyResolution = "300"
  .PDFCompressionMonoCompression = "1"
  .PDFCompressionMonoCompressionChoice = "0"
  .PDFCompressionMonoResample = "0"
  .PDFCompressionMonoResampleChoice = "0"
  .PDFCompressionMonoResolution = "1200"
  .PDFCompressionTextCompression = "1"
  .PDFFontsEmbedAll = "1"
  .PDFFontsSubSetFonts = "1"
  .PDFFontsSubSetFontsPercent = "100"
  .PDFGeneralASCII85 = "0"
  .PDFGeneralAutorotate = "0"
  .PDFGeneralCompatibility = "1"
  .PDFGeneralOverprint = "0"
  .PDFGeneralResolution = "600"
  .PDFHighEncryption = "0"
  .PDFLowEncryption = "1"
  .PDFOwnerPass = "0"
  .PDFUserPass = "0"
  .PDFUseSecurity = "0"
  .PNGColorscount = "0"
  .PrinterStop = "0"
  .PrinterTemppath = GetPDFCreatorTempfolder
  .ProcessPriority = "1"
  .ProgramFont = "MS Sans Serif"
  .ProgramFontCharset = "0"
  .ProgramFontSize = "8"
  .PSLanguageLevel = "2"
  .RemoveSpaces = "1"
  .SaveFilename = "<Title>"
  .StandardAuthor = vbNullString
  .StartStandardProgram = "1"
  .TIFFColorscount = "0"
  .UseAutosave = "0"
  .UseAutosaveDirectory = "1"
  .UseCreationDateNow = "0"
  .UseStandardAuthor = "0"
 End With
 StandardOptions = myOptions
End Function

Public Function ReadOptions() As tOptions
 Dim ini As clsINI, myOptions As tOptions, tStr As String, hOpt As New clsHash
 Set ini = New clsINI
 ini.FileName = PDFCreatorINIFile
 ini.Section = "Options"
 If ini.CheckIniFile = False Then
  ReadOptions = StandardOptions
  Exit Function
 End If
 ReadINISection PDFCreatorINIFile, "Options", hOpt
 With myOptions
  tStr = hOpt.Retrieve("AutosaveDirectory", GetMyFiles)
  If Len(Dir(tStr, vbDirectory)) > 0 And Len(tStr) > 0 Then
    .AutosaveDirectory = CompletePath(tStr)
   Else
    .AutosaveDirectory = GetMyFiles
  End If
  tStr = hOpt.Retrieve("AutosaveFilename", "<DateTime>")
  .AutosaveFilename = tStr
  tStr = hOpt.Retrieve("AutosaveFormat")
  If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
    .AutosaveFormat = CLng(tStr)
   Else
    .AutosaveFormat = 0
  End If
  tStr = hOpt.Retrieve("BitmapResolution")
  If CLng(tStr) >= 1 Then
    .BitmapResolution = CLng(tStr)
   Else
    .BitmapResolution = 150
  End If
  tStr = hOpt.Retrieve("BMPColorscount")
  If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
    .BMPColorscount = CLng(tStr)
   Else
    .BMPColorscount = 1
  End If
  tStr = hOpt.Retrieve("DirectoryGhostscriptBinaries", App.Path)
  If Len(Dir(tStr, vbDirectory)) > 0 And Len(tStr) > 0 Then
    .DirectoryGhostscriptBinaries = CompletePath(tStr)
   Else
    .DirectoryGhostscriptBinaries = App.Path & "\"
  End If
  tStr = hOpt.Retrieve("DirectoryGhostscriptFonts", App.Path & "\fonts")
  If Len(Dir(tStr, vbDirectory)) > 0 And Len(tStr) > 0 Then
    .DirectoryGhostscriptFonts = CompletePath(tStr)
   Else
    .DirectoryGhostscriptFonts = App.Path & "\fonts\"
  End If
  tStr = hOpt.Retrieve("DirectoryGhostscriptLibraries", App.Path & "\lib")
  If Len(Dir(tStr, vbDirectory)) > 0 And Len(tStr) > 0 Then
    .DirectoryGhostscriptLibraries = CompletePath(tStr)
   Else
    .DirectoryGhostscriptLibraries = App.Path & "\lib\"
  End If
  tStr = hOpt.Retrieve("EPSLanguageLevel")
  If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
    .EPSLanguageLevel = CLng(tStr)
   Else
    .EPSLanguageLevel = 2
  End If
  tStr = hOpt.Retrieve("FilenameSubstitutions", " Microsoft Word - \.doc")
  .FilenameSubstitutions = tStr
  tStr = hOpt.Retrieve("FilenameSubstitutionsOnlyInTitle")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
   Else
    .FilenameSubstitutionsOnlyInTitle = 1
  End If
  tStr = hOpt.Retrieve("JPEGColorscount")
  If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
    .JPEGColorscount = CLng(tStr)
   Else
    .JPEGColorscount = 0
  End If
  tStr = hOpt.Retrieve("JPEGQuality")
  If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
    .JPEGQuality = CLng(tStr)
   Else
    .JPEGQuality = 75
  End If
  tStr = hOpt.Retrieve("Language", "english")
  .Language = tStr
  tStr = hOpt.Retrieve("LastSaveDirectory", GetMyFiles)
  If Len(Dir(tStr, vbDirectory)) > 0 And Len(tStr) > 0 Then
    .LastSaveDirectory = CompletePath(tStr)
   Else
    .LastSaveDirectory = GetMyFiles
  End If
  tStr = hOpt.Retrieve("Logging")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .Logging = CLng(tStr)
   Else
    .Logging = 0
  End If
  tStr = hOpt.Retrieve("LogLines")
  If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
    .LogLines = CLng(tStr)
   Else
    .LogLines = 100
  End If
  tStr = hOpt.Retrieve("PCXColorscount")
  If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
    .PCXColorscount = CLng(tStr)
   Else
    .PCXColorscount = 0
  End If
  tStr = hOpt.Retrieve("PDFAllowAssembly")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFAllowAssembly = CLng(tStr)
   Else
    .PDFAllowAssembly = 0
  End If
  tStr = hOpt.Retrieve("PDFAllowCopy")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFAllowCopy = CLng(tStr)
   Else
    .PDFAllowCopy = 0
  End If
  tStr = hOpt.Retrieve("PDFAllowDegradedPrinting")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFAllowDegradedPrinting = CLng(tStr)
   Else
    .PDFAllowDegradedPrinting = 0
  End If
  tStr = hOpt.Retrieve("PDFAllowFillIn")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFAllowFillIn = CLng(tStr)
   Else
    .PDFAllowFillIn = 0
  End If
  tStr = hOpt.Retrieve("PDFAllowModifyAnnotations")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFAllowModifyAnnotations = CLng(tStr)
   Else
    .PDFAllowModifyAnnotations = 0
  End If
  tStr = hOpt.Retrieve("PDFAllowModifyContents")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFAllowModifyContents = CLng(tStr)
   Else
    .PDFAllowModifyContents = 0
  End If
  tStr = hOpt.Retrieve("PDFAllowPrinting")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFAllowPrinting = CLng(tStr)
   Else
    .PDFAllowPrinting = 0
  End If
  tStr = hOpt.Retrieve("PDFAllowScreenReaders")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFAllowScreenReaders = CLng(tStr)
   Else
    .PDFAllowScreenReaders = 0
  End If
  tStr = hOpt.Retrieve("PDFColorsCMYKToRGB")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFColorsCMYKToRGB = CLng(tStr)
   Else
    .PDFColorsCMYKToRGB = 1
  End If
  tStr = hOpt.Retrieve("PDFColorsColorModel")
  If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
    .PDFColorsColorModel = CLng(tStr)
   Else
    .PDFColorsColorModel = 1
  End If
  tStr = hOpt.Retrieve("PDFColorsPreserveHalftone")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFColorsPreserveHalftone = CLng(tStr)
   Else
    .PDFColorsPreserveHalftone = 0
  End If
  tStr = hOpt.Retrieve("PDFColorsPreserveOverprint")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFColorsPreserveOverprint = CLng(tStr)
   Else
    .PDFColorsPreserveOverprint = 1
  End If
  tStr = hOpt.Retrieve("PDFColorsPreserveTransfer")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFColorsPreserveTransfer = CLng(tStr)
   Else
    .PDFColorsPreserveTransfer = 1
  End If
  tStr = hOpt.Retrieve("PDFCompressionColorCompression")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionColorCompression = CLng(tStr)
   Else
    .PDFCompressionColorCompression = 1
  End If
  tStr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
  If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
    .PDFCompressionColorCompressionChoice = CLng(tStr)
   Else
    .PDFCompressionColorCompressionChoice = 0
  End If
  tStr = hOpt.Retrieve("PDFCompressionColorResample")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionColorResample = CLng(tStr)
   Else
    .PDFCompressionColorResample = 0
  End If
  tStr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
  If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
    .PDFCompressionColorResampleChoice = CLng(tStr)
   Else
    .PDFCompressionColorResampleChoice = 0
  End If
  tStr = hOpt.Retrieve("PDFCompressionColorResolution")
  If CLng(tStr) >= 0 Then
    .PDFCompressionColorResolution = CLng(tStr)
   Else
    .PDFCompressionColorResolution = 300
  End If
  tStr = hOpt.Retrieve("PDFCompressionGreyCompression")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionGreyCompression = CLng(tStr)
   Else
    .PDFCompressionGreyCompression = 1
  End If
  tStr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
  If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
    .PDFCompressionGreyCompressionChoice = CLng(tStr)
   Else
    .PDFCompressionGreyCompressionChoice = 0
  End If
  tStr = hOpt.Retrieve("PDFCompressionGreyResample")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionGreyResample = CLng(tStr)
   Else
    .PDFCompressionGreyResample = 0
  End If
  tStr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
  If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
    .PDFCompressionGreyResampleChoice = CLng(tStr)
   Else
    .PDFCompressionGreyResampleChoice = 0
  End If
  tStr = hOpt.Retrieve("PDFCompressionGreyResolution")
  If CLng(tStr) >= 0 Then
    .PDFCompressionGreyResolution = CLng(tStr)
   Else
    .PDFCompressionGreyResolution = 300
  End If
  tStr = hOpt.Retrieve("PDFCompressionMonoCompression")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionMonoCompression = CLng(tStr)
   Else
    .PDFCompressionMonoCompression = 1
  End If
  tStr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
  If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
    .PDFCompressionMonoCompressionChoice = CLng(tStr)
   Else
    .PDFCompressionMonoCompressionChoice = 0
  End If
  tStr = hOpt.Retrieve("PDFCompressionMonoResample")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionMonoResample = CLng(tStr)
   Else
    .PDFCompressionMonoResample = 0
  End If
  tStr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
  If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
    .PDFCompressionMonoResampleChoice = CLng(tStr)
   Else
    .PDFCompressionMonoResampleChoice = 0
  End If
  tStr = hOpt.Retrieve("PDFCompressionMonoResolution")
  If CLng(tStr) >= 0 Then
    .PDFCompressionMonoResolution = CLng(tStr)
   Else
    .PDFCompressionMonoResolution = 1200
  End If
  tStr = hOpt.Retrieve("PDFCompressionTextCompression")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionTextCompression = CLng(tStr)
   Else
    .PDFCompressionTextCompression = 1
  End If
  tStr = hOpt.Retrieve("PDFFontsEmbedAll")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFFontsEmbedAll = CLng(tStr)
   Else
    .PDFFontsEmbedAll = 1
  End If
  tStr = hOpt.Retrieve("PDFFontsSubSetFonts")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFFontsSubSetFonts = CLng(tStr)
   Else
    .PDFFontsSubSetFonts = 1
  End If
  tStr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
  If CLng(tStr) >= 0 Then
    .PDFFontsSubSetFontsPercent = CLng(tStr)
   Else
    .PDFFontsSubSetFontsPercent = 100
  End If
  tStr = hOpt.Retrieve("PDFGeneralASCII85")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFGeneralASCII85 = CLng(tStr)
   Else
    .PDFGeneralASCII85 = 0
  End If
  tStr = hOpt.Retrieve("PDFGeneralAutorotate")
  If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
    .PDFGeneralAutorotate = CLng(tStr)
   Else
    .PDFGeneralAutorotate = 0
  End If
  tStr = hOpt.Retrieve("PDFGeneralCompatibility")
  If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
    .PDFGeneralCompatibility = CLng(tStr)
   Else
    .PDFGeneralCompatibility = 1
  End If
  tStr = hOpt.Retrieve("PDFGeneralOverprint")
  If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
    .PDFGeneralOverprint = CLng(tStr)
   Else
    .PDFGeneralOverprint = 0
  End If
  tStr = hOpt.Retrieve("PDFGeneralResolution")
  If CLng(tStr) >= 0 Then
    .PDFGeneralResolution = CLng(tStr)
   Else
    .PDFGeneralResolution = 600
  End If
  tStr = hOpt.Retrieve("PDFHighEncryption")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFHighEncryption = CLng(tStr)
   Else
    .PDFHighEncryption = 0
  End If
  tStr = hOpt.Retrieve("PDFLowEncryption")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFLowEncryption = CLng(tStr)
   Else
    .PDFLowEncryption = 1
  End If
  tStr = hOpt.Retrieve("PDFOwnerPass")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFOwnerPass = CLng(tStr)
   Else
    .PDFOwnerPass = 0
  End If
  tStr = hOpt.Retrieve("PDFUserPass")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFUserPass = CLng(tStr)
   Else
    .PDFUserPass = 0
  End If
  tStr = hOpt.Retrieve("PDFUseSecurity")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFUseSecurity = CLng(tStr)
   Else
    .PDFUseSecurity = 0
  End If
  tStr = hOpt.Retrieve("PNGColorscount")
  If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
    .PNGColorscount = CLng(tStr)
   Else
    .PNGColorscount = 0
  End If
  tStr = hOpt.Retrieve("PrinterStop")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PrinterStop = CLng(tStr)
   Else
    .PrinterStop = 0
  End If
  tStr = hOpt.Retrieve("PrinterTemppath", GetPDFCreatorTempfolder)
  If Len(Dir(tStr, vbDirectory)) > 0 And Len(tStr) > 0 Then
    .PrinterTemppath = CompletePath(tStr)
   Else
    .PrinterTemppath = GetPDFCreatorTempfolder
  End If
  tStr = hOpt.Retrieve("ProcessPriority")
  If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
    .ProcessPriority = CLng(tStr)
   Else
    .ProcessPriority = 1
  End If
  tStr = hOpt.Retrieve("ProgramFont", "MS Sans Serif")
  .ProgramFont = tStr
  tStr = hOpt.Retrieve("ProgramFontCharset")
  If CLng(tStr) >= 0 Then
    .ProgramFontCharset = CLng(tStr)
   Else
    .ProgramFontCharset = 0
  End If
  tStr = hOpt.Retrieve("ProgramFontSize")
  If CLng(tStr) >= 1 And CLng(tStr) <= 72 Then
    .ProgramFontSize = CLng(tStr)
   Else
    .ProgramFontSize = 8
  End If
  tStr = hOpt.Retrieve("PSLanguageLevel")
  If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
    .PSLanguageLevel = CLng(tStr)
   Else
    .PSLanguageLevel = 2
  End If
  tStr = hOpt.Retrieve("RemoveSpaces")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .RemoveSpaces = CLng(tStr)
   Else
    .RemoveSpaces = 1
  End If
  tStr = hOpt.Retrieve("SaveFilename", "<Title>")
  .SaveFilename = tStr
  tStr = hOpt.Retrieve("StandardAuthor", "")
  .StandardAuthor = tStr
  tStr = hOpt.Retrieve("StartStandardProgram")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .StartStandardProgram = CLng(tStr)
   Else
    .StartStandardProgram = 1
  End If
  tStr = hOpt.Retrieve("TIFFColorscount")
  If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
    .TIFFColorscount = CLng(tStr)
   Else
    .TIFFColorscount = 0
  End If
  tStr = hOpt.Retrieve("UseAutosave")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .UseAutosave = CLng(tStr)
   Else
    .UseAutosave = 0
  End If
  tStr = hOpt.Retrieve("UseAutosaveDirectory")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .UseAutosaveDirectory = CLng(tStr)
   Else
    .UseAutosaveDirectory = 1
  End If
  tStr = hOpt.Retrieve("UseCreationDateNow")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .UseCreationDateNow = CLng(tStr)
   Else
    .UseCreationDateNow = 0
  End If
  tStr = hOpt.Retrieve("UseStandardAuthor")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .UseStandardAuthor = CLng(tStr)
   Else
    .UseStandardAuthor = 0
  End If
 End With
 Set ini = Nothing
 ReadOptions = myOptions
End Function

Public Sub SaveOptions(sOptions As tOptions)
 Dim ini As clsINI
 Set ini = New clsINI
 ini.FileName = PDFCreatorINIFile
 ini.Section = "Options"
 If ini.CheckIniFile = False Then
  ini.CreateIniFile
 End If
 With sOptions
  ini.SaveKey CStr(.AutosaveDirectory), "AutosaveDirectory"
  ini.SaveKey CStr(.AutosaveFilename), "AutosaveFilename"
  ini.SaveKey CStr(.AutosaveFormat), "AutosaveFormat"
  ini.SaveKey CStr(.BitmapResolution), "BitmapResolution"
  ini.SaveKey CStr(.BMPColorscount), "BMPColorscount"
  ini.SaveKey CStr(.DirectoryGhostscriptBinaries), "DirectoryGhostscriptBinaries"
  ini.SaveKey CStr(.DirectoryGhostscriptFonts), "DirectoryGhostscriptFonts"
  ini.SaveKey CStr(.DirectoryGhostscriptLibraries), "DirectoryGhostscriptLibraries"
  ini.SaveKey CStr(.EPSLanguageLevel), "EPSLanguageLevel"
  ini.SaveKey CStr(.FilenameSubstitutions), "FilenameSubstitutions"
  ini.SaveKey CStr(.FilenameSubstitutionsOnlyInTitle), "FilenameSubstitutionsOnlyInTitle"
  ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
  ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
  ini.SaveKey CStr(.Language), "Language"
  ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
  ini.SaveKey CStr(.Logging), "Logging"
  ini.SaveKey CStr(.LogLines), "LogLines"
  ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
  ini.SaveKey CStr(.PDFAllowAssembly), "PDFAllowAssembly"
  ini.SaveKey CStr(.PDFAllowCopy), "PDFAllowCopy"
  ini.SaveKey CStr(.PDFAllowDegradedPrinting), "PDFAllowDegradedPrinting"
  ini.SaveKey CStr(.PDFAllowFillIn), "PDFAllowFillIn"
  ini.SaveKey CStr(.PDFAllowModifyAnnotations), "PDFAllowModifyAnnotations"
  ini.SaveKey CStr(.PDFAllowModifyContents), "PDFAllowModifyContents"
  ini.SaveKey CStr(.PDFAllowPrinting), "PDFAllowPrinting"
  ini.SaveKey CStr(.PDFAllowScreenReaders), "PDFAllowScreenReaders"
  ini.SaveKey CStr(.PDFColorsCMYKToRGB), "PDFColorsCMYKToRGB"
  ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
  ini.SaveKey CStr(.PDFColorsPreserveHalftone), "PDFColorsPreserveHalftone"
  ini.SaveKey CStr(.PDFColorsPreserveOverprint), "PDFColorsPreserveOverprint"
  ini.SaveKey CStr(.PDFColorsPreserveTransfer), "PDFColorsPreserveTransfer"
  ini.SaveKey CStr(.PDFCompressionColorCompression), "PDFCompressionColorCompression"
  ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
  ini.SaveKey CStr(.PDFCompressionColorResample), "PDFCompressionColorResample"
  ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
  ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
  ini.SaveKey CStr(.PDFCompressionGreyCompression), "PDFCompressionGreyCompression"
  ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
  ini.SaveKey CStr(.PDFCompressionGreyResample), "PDFCompressionGreyResample"
  ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
  ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
  ini.SaveKey CStr(.PDFCompressionMonoCompression), "PDFCompressionMonoCompression"
  ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
  ini.SaveKey CStr(.PDFCompressionMonoResample), "PDFCompressionMonoResample"
  ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
  ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
  ini.SaveKey CStr(.PDFCompressionTextCompression), "PDFCompressionTextCompression"
  ini.SaveKey CStr(.PDFFontsEmbedAll), "PDFFontsEmbedAll"
  ini.SaveKey CStr(.PDFFontsSubSetFonts), "PDFFontsSubSetFonts"
  ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
  ini.SaveKey CStr(.PDFGeneralASCII85), "PDFGeneralASCII85"
  ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
  ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
  ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
  ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
  ini.SaveKey CStr(.PDFHighEncryption), "PDFHighEncryption"
  ini.SaveKey CStr(.PDFLowEncryption), "PDFLowEncryption"
  ini.SaveKey CStr(.PDFOwnerPass), "PDFOwnerPass"
  ini.SaveKey CStr(.PDFUserPass), "PDFUserPass"
  ini.SaveKey CStr(.PDFUseSecurity), "PDFUseSecurity"
  ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
  ini.SaveKey CStr(.PrinterStop), "PrinterStop"
  ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
  ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
  ini.SaveKey CStr(.ProgramFont), "ProgramFont"
  ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
  ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
  ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
  ini.SaveKey CStr(.RemoveSpaces), "RemoveSpaces"
  ini.SaveKey CStr(.SaveFilename), "SaveFilename"
  ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
  ini.SaveKey CStr(.StartStandardProgram), "StartStandardProgram"
  ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
  ini.SaveKey CStr(.UseAutosave), "UseAutosave"
  ini.SaveKey CStr(.UseAutosaveDirectory), "UseAutosaveDirectory"
  ini.SaveKey CStr(.UseCreationDateNow), "UseCreationDateNow"
  ini.SaveKey CStr(.UseStandardAuthor), "UseStandardAuthor"
 End With
 Set ini = Nothing
End Sub

Public Sub ShowOptions(Frm As Form, sOptions As tOptions)
' On Local Error Resume Next
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
  Frm.cmbPCXColors.ListIndex = .PCXColorscount
  Frm.chkAllowAssembly.Value = .PDFAllowAssembly
  Frm.chkAllowCopy.Value = .PDFAllowCopy
  Frm.chkAllowDegradedPrinting.Value = .PDFAllowDegradedPrinting
  Frm.chkAllowFillIn.Value = .PDFAllowFillIn
  Frm.chkAllowModifyAnnotations.Value = .PDFAllowModifyAnnotations
  Frm.chkAllowModifyContents.Value = .PDFAllowModifyContents
  Frm.chkAllowPrinting.Value = .PDFAllowPrinting
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
 Dim i As Long, tStr As String, lsv As ListView
 With sOptions
  .AutosaveDirectory = Frm.txtAutosaveDirectory.Text
  .AutosaveFilename = Frm.txtAutosaveFilename.Text
  .AutosaveFormat = Frm.cmbAutosaveFormat.ListIndex
  .BitmapResolution = Frm.txtBitmapResolution.Text
  .BMPColorscount = Frm.cmbBMPColors.ListIndex
  .DirectoryGhostscriptBinaries = Frm.txtGSbin.Text
  .DirectoryGhostscriptFonts = Frm.txtGSfonts.Text
  .DirectoryGhostscriptLibraries = Frm.txtGSlib.Text
  .EPSLanguageLevel = Frm.cmbEPSLanguageLevel.ListIndex
  tStr = ""
  Set lsv = Frm.lsvFilenameSubst
  For i = 1 To lsv.ListItems.Count
   If i < lsv.ListItems.Count Then
     tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1) & "\"
    Else
     tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1)
   End If
  Next i
  .FilenameSubstitutions = tStr
  .FilenameSubstitutionsOnlyInTitle = Frm.chkFilenameSubst.Value
  .JPEGColorscount = Frm.cmbJPEGColors.ListIndex
  .JPEGQuality = Frm.txtJPEGQuality.Text
  .PCXColorscount = Frm.cmbPCXColors.ListIndex
  .PDFAllowAssembly = Frm.chkAllowAssembly.Value
  .PDFAllowCopy = Frm.chkAllowCopy.Value
  .PDFAllowDegradedPrinting = Frm.chkAllowDegradedPrinting.Value
  .PDFAllowFillIn = Frm.chkAllowFillIn.Value
  .PDFAllowModifyAnnotations = Frm.chkAllowModifyAnnotations.Value
  .PDFAllowModifyContents = Frm.chkAllowModifyContents.Value
  .PDFAllowPrinting = Frm.chkAllowPrinting.Value
  .PDFAllowScreenReaders = Frm.chkAllowScreenReaders.Value
  .PDFColorsCMYKToRGB = Frm.chkPDFCMYKtoRGB.Value
  .PDFColorsColorModel = Frm.cmbPDFColorModel.ListIndex
  .PDFColorsPreserveHalftone = Frm.chkPDFPreserveHalftone.Value
  .PDFColorsPreserveOverprint = Frm.chkPDFPreserveOverprint.Value
  .PDFColorsPreserveTransfer = Frm.chkPDFPreserveTransfer.Value
  .PDFCompressionColorCompression = Frm.chkPDFColorComp.Value
  .PDFCompressionColorCompressionChoice = Frm.cmbPDFColorComp.ListIndex
  .PDFCompressionColorResample = Frm.chkPDFColorResample.Value
  .PDFCompressionColorResampleChoice = Frm.cmbPDFColorResample.ListIndex
  .PDFCompressionColorResolution = Frm.txtPDFColorRes.Text
  .PDFCompressionGreyCompression = Frm.chkPDFGreyComp.Value
  .PDFCompressionGreyCompressionChoice = Frm.cmbPDFGreyComp.ListIndex
  .PDFCompressionGreyResample = Frm.chkPDFGreyResample.Value
  .PDFCompressionGreyResampleChoice = Frm.cmbPDFGreyResample.ListIndex
  .PDFCompressionGreyResolution = Frm.txtPDFGreyRes.Text
  .PDFCompressionMonoCompression = Frm.chkPDFMonoComp.Value
  .PDFCompressionMonoCompressionChoice = Frm.cmbPDFMonoComp.ListIndex
  .PDFCompressionMonoResample = Frm.chkPDFMonoResample.Value
  .PDFCompressionMonoResampleChoice = Frm.cmbPDFMonoResample.ListIndex
  .PDFCompressionMonoResolution = Frm.txtPDFMonoRes.Text
  .PDFCompressionTextCompression = Frm.chkPDFTextComp.Value
  .PDFFontsEmbedAll = Frm.chkPDFEmbedAll.Value
  .PDFFontsSubSetFonts = Frm.chkPDFSubSetFonts.Value
  .PDFFontsSubSetFontsPercent = Frm.txtPDFSubSetPerc.Text
  .PDFGeneralASCII85 = Frm.chkPDFASCII85.Value
  .PDFGeneralAutorotate = Frm.cmbPDFRotate.ListIndex
  .PDFGeneralCompatibility = Frm.cmbPDFCompat.ListIndex
  .PDFGeneralOverprint = Frm.cmbPDFOverprint.ListIndex
  .PDFGeneralResolution = Frm.txtPDFRes.Text
  .PDFHighEncryption = Frm.optEncHigh.Value
  .PDFLowEncryption = Frm.optEncLow.Value
  .PDFOwnerPass = Frm.chkOwnerPass.Value
  .PDFUserPass = Frm.chkUserPass.Value
  .PDFUseSecurity = Frm.chkUseSecurity.Value
  .PNGColorscount = Frm.cmbPNGColors.ListIndex
  .PrinterTemppath = Frm.txtTemppath.Text
  .ProcessPriority = Frm.sldProcessPriority.Value
  .ProgramFont = Frm.cmbFonts.List(Frm.cmbFonts.ListIndex)
  .ProgramFontCharset = Frm.cmbCharset.Text
  .ProgramFontSize = Frm.txtProgramFontsize.Text
  .PSLanguageLevel = Frm.cmbPSLanguageLevel.ListIndex
  .RemoveSpaces = Frm.chkSpaces.Value
  .SaveFilename = Frm.txtSaveFilename.Text
  .StandardAuthor = Frm.txtStandardAuthor.Text
  .TIFFColorscount = Frm.cmbTIFFColors.ListIndex
  .UseAutosave = Frm.chkUseAutosave.Value
  .UseAutosaveDirectory = Frm.chkUseAutosaveDirectory.Value
  .UseCreationDateNow = Frm.chkUseCreationDateNow.Value
  .UseStandardAuthor = Frm.chkUseStandardAuthor.Value
 End With
End Sub

Public Sub SetPrinterStop(StopPrinter As Boolean)
 If StopPrinter = True Then
   Options.PrinterStop = 1
   PrinterStop = True
  Else
   Options.PrinterStop = 0
   PrinterStop = False
 End If
 SaveOptions Options
End Sub

Public Sub SetLogging(Logging As Boolean)
 If Logging = True Then
   Options.Logging = 1
  Else
   Options.Logging = 0
 End If
 SaveOptions Options
End Sub

Public Sub SetLanguage(Language As String)
 Options.Language = Language
 SaveOptions Options
End Sub
