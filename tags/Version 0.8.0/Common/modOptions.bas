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
 Dim myOptions As tOptions, reg as clsRegistry
 With myOptions
  .AutosaveDirectory = " "
  .AutosaveFilename = "<DateTime>"
  .AutosaveFormat = "0"
  .BitmapResolution = "150"
  .BMPColorscount = "1"
  Set reg = New clsRegistry
  reg.hkey = HKEY_LOCAL_MACHINE
  reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
  .DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
  Set reg = Nothing
  Set reg = New clsRegistry
  reg.hkey = HKEY_LOCAL_MACHINE
  reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
  .DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
  Set reg = Nothing
  Set reg = New clsRegistry
  reg.hkey = HKEY_LOCAL_MACHINE
  reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
  .DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
  Set reg = Nothing
  .EPSLanguageLevel = "2"
  .FilenameSubstitutions = "Microsoft Word - \.doc"
  .FilenameSubstitutionsOnlyInTitle = "1"
  .JPEGColorscount = "0"
  .JPEGQuality = "75"
  .Language = "english"
  .LastSaveDirectory = GetMyFiles
  .Logging = "0"
  .LogLines = "100"
  .NoConfirmMessageSwitchingDefaultprinter = "0"
  .OptionsEnabled = "1"
  .OptionsVisible = "1"
  .PCXColorscount = "0"
  .PDFAllowAssembly = "0"
  .PDFAllowDegradedPrinting = "0"
  .PDFAllowFillIn = "0"
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
  .PDFDisallowCopy = "1"
  .PDFDisallowModifyAnnotations = "0"
  .PDFDisallowModifyContents = "0"
  .PDFDisallowPrinting = "0"
  .PDFEncryptor = "0"
  .PDFFontsEmbedAll = "1"
  .PDFFontsSubSetFonts = "1"
  .PDFFontsSubSetFontsPercent = "100"
  .PDFGeneralASCII85 = "0"
  .PDFGeneralAutorotate = "2"
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
  .PrinterTemppath = GetTempPath & "PDFCreator\"
  .ProcessPriority = "1"
  .ProgramFont = "MS Sans Serif"
  .ProgramFontCharset = "0"
  .ProgramFontSize = "8"
  .PSLanguageLevel = "2"
  .RemoveSpaces = "1"
  .SaveFilename = "<Title>"
  .StandardAuthor = " "
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
 Dim ini As clsINI, myOptions As tOptions, tStr as String, hOpt As New clsHash
 Set ini = New clsINI
 ini.Filename = PDFCreatorINIFile
 ini.Section = "Options"
 If ini.Checkinifile = False Then
  ReadOptions = StandardOptions
  Exit Function
 End If
 ReadINISection PDFCreatorINIFile, "Options", hOpt
 With myOptions
  tStr = hOpt.Retrieve("AutosaveDirectory", GetMyFiles)
  If DirExists(tStr) = True Then
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
  If DirExists(tStr) = True Then
    .DirectoryGhostscriptBinaries = CompletePath(tStr)
   Else
    .DirectoryGhostscriptBinaries = ""
  End If
  tStr = hOpt.Retrieve("DirectoryGhostscriptFonts", App.Path & "\fonts")
  If DirExists(tStr) = True Then
    .DirectoryGhostscriptFonts = CompletePath(tStr)
   Else
    .DirectoryGhostscriptFonts = ""
  End If
  tStr = hOpt.Retrieve("DirectoryGhostscriptLibraries", App.Path & "\lib")
  If DirExists(tStr) = True Then
    .DirectoryGhostscriptLibraries = CompletePath(tStr)
   Else
    .DirectoryGhostscriptLibraries = ""
  End If
  tStr = hOpt.Retrieve("EPSLanguageLevel")
  If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
    .EPSLanguageLevel = CLng(tStr)
   Else
    .EPSLanguageLevel = 2
  End If
  tStr = hOpt.Retrieve("FilenameSubstitutions", "Microsoft Word - \.doc")
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
  If DirExists(tStr) = True Then
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
  tStr = hOpt.Retrieve("NoConfirmMessageSwitchingDefaultprinter")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
   Else
    .NoConfirmMessageSwitchingDefaultprinter = 0
  End If
  tStr = hOpt.Retrieve("OptionsEnabled")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .OptionsEnabled = CLng(tStr)
   Else
    .OptionsEnabled = 1
  End If
  tStr = hOpt.Retrieve("OptionsVisible")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .OptionsVisible = CLng(tStr)
   Else
    .OptionsVisible = 1
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
  If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
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
  If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
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
  If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
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
  tStr = hOpt.Retrieve("PDFDisallowCopy")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFDisallowCopy = CLng(tStr)
   Else
    .PDFDisallowCopy = 1
  End If
  tStr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFDisallowModifyAnnotations = CLng(tStr)
   Else
    .PDFDisallowModifyAnnotations = 0
  End If
  tStr = hOpt.Retrieve("PDFDisallowModifyContents")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFDisallowModifyContents = CLng(tStr)
   Else
    .PDFDisallowModifyContents = 0
  End If
  tStr = hOpt.Retrieve("PDFDisallowPrinting")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFDisallowPrinting = CLng(tStr)
   Else
    .PDFDisallowPrinting = 0
  End If
  tStr = hOpt.Retrieve("PDFEncryptor")
  If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
    .PDFEncryptor = CLng(tStr)
   Else
    .PDFEncryptor = 0
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
    .PDFGeneralAutorotate = 2
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
  tStr = hOpt.Retrieve("PrinterTemppath", GetTempPath & "PDFCreator\")
  If DirExists(tStr) = True Then
    .PrinterTemppath = CompletePath(tStr)
   Else
    .PrinterTemppath = GetTempPath & "PDFCreator\"
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
  tStr = hOpt.Retrieve("StandardAuthor", " ")
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

Public Sub SaveOptions(sOptions as tOptions)
 Dim ini As clsINI
 Set ini = New clsINI
 ini.Filename = PDFCreatorINIFile
 ini.Section = "Options"
 If ini.CheckInifile = False Then
  ini.CreateInifile
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
  ini.SaveKey CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), "FilenameSubstitutionsOnlyInTitle"
  ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
  ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
  ini.SaveKey CStr(.Language), "Language"
  ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
  ini.SaveKey CStr(Abs(.Logging)), "Logging"
  ini.SaveKey CStr(.LogLines), "LogLines"
  ini.SaveKey CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), "NoConfirmMessageSwitchingDefaultprinter"
  ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
  ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
  ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
  ini.SaveKey CStr(Abs(.PDFAllowAssembly)), "PDFAllowAssembly"
  ini.SaveKey CStr(Abs(.PDFAllowDegradedPrinting)), "PDFAllowDegradedPrinting"
  ini.SaveKey CStr(Abs(.PDFAllowFillIn)), "PDFAllowFillIn"
  ini.SaveKey CStr(Abs(.PDFAllowScreenReaders)), "PDFAllowScreenReaders"
  ini.SaveKey CStr(Abs(.PDFColorsCMYKToRGB)), "PDFColorsCMYKToRGB"
  ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
  ini.SaveKey CStr(Abs(.PDFColorsPreserveHalftone)), "PDFColorsPreserveHalftone"
  ini.SaveKey CStr(Abs(.PDFColorsPreserveOverprint)), "PDFColorsPreserveOverprint"
  ini.SaveKey CStr(Abs(.PDFColorsPreserveTransfer)), "PDFColorsPreserveTransfer"
  ini.SaveKey CStr(Abs(.PDFCompressionColorCompression)), "PDFCompressionColorCompression"
  ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
  ini.SaveKey CStr(Abs(.PDFCompressionColorResample)), "PDFCompressionColorResample"
  ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
  ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
  ini.SaveKey CStr(Abs(.PDFCompressionGreyCompression)), "PDFCompressionGreyCompression"
  ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
  ini.SaveKey CStr(Abs(.PDFCompressionGreyResample)), "PDFCompressionGreyResample"
  ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
  ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
  ini.SaveKey CStr(Abs(.PDFCompressionMonoCompression)), "PDFCompressionMonoCompression"
  ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
  ini.SaveKey CStr(Abs(.PDFCompressionMonoResample)), "PDFCompressionMonoResample"
  ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
  ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
  ini.SaveKey CStr(Abs(.PDFCompressionTextCompression)), "PDFCompressionTextCompression"
  ini.SaveKey CStr(Abs(.PDFDisallowCopy)), "PDFDisallowCopy"
  ini.SaveKey CStr(Abs(.PDFDisallowModifyAnnotations)), "PDFDisallowModifyAnnotations"
  ini.SaveKey CStr(Abs(.PDFDisallowModifyContents)), "PDFDisallowModifyContents"
  ini.SaveKey CStr(Abs(.PDFDisallowPrinting)), "PDFDisallowPrinting"
  ini.SaveKey CStr(.PDFEncryptor), "PDFEncryptor"
  ini.SaveKey CStr(Abs(.PDFFontsEmbedAll)), "PDFFontsEmbedAll"
  ini.SaveKey CStr(Abs(.PDFFontsSubSetFonts)), "PDFFontsSubSetFonts"
  ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
  ini.SaveKey CStr(Abs(.PDFGeneralASCII85)), "PDFGeneralASCII85"
  ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
  ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
  ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
  ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
  ini.SaveKey CStr(Abs(.PDFHighEncryption)), "PDFHighEncryption"
  ini.SaveKey CStr(Abs(.PDFLowEncryption)), "PDFLowEncryption"
  ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
  ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
  ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
  ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
  ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
  ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
  ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
  ini.SaveKey CStr(.ProgramFont), "ProgramFont"
  ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
  ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
  ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
  ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
  ini.SaveKey CStr(.SaveFilename), "SaveFilename"
  ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
  ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
  ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
  ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
  ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
  ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
  ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
 End With
 Set ini = Nothing
End Sub

Public Sub ShowOptions(Frm as Form, sOptions as tOptions)
 On Error Resume Next
 Dim i as Long, tList() as String, tStrA() As String, lsv As ListView
 With sOptions
  frm.txtAutosaveDirectory.Text = .AutosaveDirectory
  frm.txtAutosaveFilename.Text = .AutosaveFilename
  frm.cmbAutosaveFormat.Listindex = .AutosaveFormat
  frm.txtBitmapResolution.Text = .BitmapResolution
  frm.cmbBMPColors.Listindex = .BMPColorscount
  frm.txtGSbin.text = .DirectoryGhostscriptBinaries
  frm.txtGSfonts.text = .DirectoryGhostscriptFonts
  frm.txtGSlib.text = .DirectoryGhostscriptLibraries
  frm.cmbEPSLanguageLevel.Listindex = .EPSLanguageLevel
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
  frm.chkFilenameSubst.Value = .FilenameSubstitutionsOnlyInTitle
  frm.cmbJPEGColors.Listindex = .JPEGColorscount
  frm.txtJPEGQuality.Text = .JPEGQuality
  frm.chkNoConfirmMessageSwitchingDefaultprinter = .NoConfirmMessageSwitchingDefaultprinter
  frm.cmbPCXColors.Listindex = .PCXColorscount
  frm.chkAllowAssembly.Value = .PDFAllowAssembly
  frm.chkAllowDegradedPrinting.Value = .PDFAllowDegradedPrinting
  frm.chkAllowFillIn.Value = .PDFAllowFillIn
  frm.chkAllowScreenReaders.Value = .PDFAllowScreenReaders
  frm.chkPDFCMYKtoRGB.Value = .PDFColorsCMYKToRGB
  frm.cmbPDFColorModel.Listindex = .PDFColorsColorModel
  frm.chkPDFPreserveHalftone.Value = .PDFColorsPreserveHalftone
  frm.chkPDFPreserveOverprint.Value = .PDFColorsPreserveOverprint
  frm.chkPDFPreserveTransfer.Value = .PDFColorsPreserveTransfer
  frm.chkPDFColorComp.Value = .PDFCompressionColorCompression
  frm.cmbPDFColorComp.Listindex = .PDFCompressionColorCompressionChoice
  frm.chkPDFColorResample.Value = .PDFCompressionColorResample
  frm.cmbPDFColorResample.Listindex = .PDFCompressionColorResampleChoice
  frm.txtPDFColorRes.Text = .PDFCompressionColorResolution
  frm.chkPDFGreyComp.Value = .PDFCompressionGreyCompression
  frm.cmbPDFGreyComp.Listindex = .PDFCompressionGreyCompressionChoice
  frm.chkPDFGreyResample.Value = .PDFCompressionGreyResample
  frm.cmbPDFGreyResample.Listindex = .PDFCompressionGreyResampleChoice
  frm.txtPDFGreyRes.Text = .PDFCompressionGreyResolution
  frm.chkPDFMonoComp.Value = .PDFCompressionMonoCompression
  frm.cmbPDFMonoComp.Listindex = .PDFCompressionMonoCompressionChoice
  frm.chkPDFMonoResample.Value = .PDFCompressionMonoResample
  frm.cmbPDFMonoResample.Listindex = .PDFCompressionMonoResampleChoice
  frm.txtPDFMonoRes.Text = .PDFCompressionMonoResolution
  frm.chkPDFTextComp.Value = .PDFCompressionTextCompression
  frm.chkAllowCopy.Value = .PDFDisallowCopy
  frm.chkAllowModifyAnnotations.Value = .PDFDisallowModifyAnnotations
  frm.chkAllowModifyContents.Value = .PDFDisallowModifyContents
  frm.chkAllowPrinting.Value = .PDFDisallowPrinting
  frm.cmbPDFEncryptor.Itemdata(Frm.cmbPDFEncryptor.Listindex) = .PDFEncryptor
  frm.chkPDFEmbedAll.Value = .PDFFontsEmbedAll
  frm.chkPDFSubSetFonts.Value = .PDFFontsSubSetFonts
  frm.txtPDFSubSetPerc.Text = .PDFFontsSubSetFontsPercent
  frm.chkPDFASCII85.Value = .PDFGeneralASCII85
  frm.cmbPDFRotate.Listindex = .PDFGeneralAutorotate
  frm.cmbPDFCompat.Listindex = .PDFGeneralCompatibility
  frm.cmbPDFOverprint.Listindex = .PDFGeneralOverprint
  frm.txtPDFRes.Text = .PDFGeneralResolution
  frm.optEncHigh.Value = .PDFHighEncryption
  frm.optEncLow.Value = .PDFLowEncryption
  frm.chkOwnerPass.Value = .PDFOwnerPass
  frm.chkUserPass.Value = .PDFUserPass
  frm.chkUseSecurity.Value = .PDFUseSecurity
  frm.cmbPNGColors.Listindex = .PNGColorscount
  frm.txtTemppath.text = .PrinterTemppath
  frm.sldProcessPriority.value = .ProcessPriority
  For i=0 to frm.cmbFonts.Listcount - 1
    If Ucase$(frm.cmbFonts.List(i)) = Ucase$(.ProgramFont) Then
     frm.cmbFonts.Listindex = i
     Exit For
    End If
  Next i
  frm.cmbCharset.Text = .ProgramFontCharset
  frm.txtProgramFontSize.text = .ProgramFontSize
  frm.cmbPSLanguageLevel.Listindex = .PSLanguageLevel
  frm.chkSpaces.Value = .RemoveSpaces
  frm.txtSaveFilename.Text = .SaveFilename
  frm.txtStandardAuthor.Text = .StandardAuthor
  frm.cmbTIFFColors.Listindex = .TIFFColorscount
  frm.chkUseAutosave.Value = .UseAutosave
  frm.chkUseAutosaveDirectory.Value = .UseAutosaveDirectory
  frm.chkUseCreationDateNow.Value = .UseCreationDateNow
  frm.chkUseStandardAuthor.Value = .UseStandardAuthor
 End With
End Sub

Public Sub GetOptions(Frm as Form, sOptions as tOptions)
 Dim i as Long, tStr as String, lsv As ListView
 With sOptions
 .AutosaveDirectory =  frm.txtAutosaveDirectory.Text
 .AutosaveFilename =  frm.txtAutosaveFilename.Text
 .AutosaveFormat =  frm.cmbAutosaveFormat.Listindex
 .BitmapResolution =  frm.txtBitmapResolution.Text
 .BMPColorscount =  frm.cmbBMPColors.Listindex
 .DirectoryGhostscriptBinaries =  frm.txtGSbin.text
 .DirectoryGhostscriptFonts =  frm.txtGSfonts.text
 .DirectoryGhostscriptLibraries =  frm.txtGSlib.text
 .EPSLanguageLevel =  frm.cmbEPSLanguageLevel.Listindex
 tStr=""
 Set lsv = Frm.lsvFilenameSubst
 For i = 1 To lsv.ListItems.Count
  If i < lsv.ListItems.Count Then
    tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1) & "\"
   Else
    tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1)
  End If
 Next i
 .FilenameSubstitutions = tStr
 .FilenameSubstitutionsOnlyInTitle =  Abs(frm.chkFilenameSubst.Value)
 .JPEGColorscount =  frm.cmbJPEGColors.Listindex
 .JPEGQuality =  frm.txtJPEGQuality.Text
 .NoConfirmMessageSwitchingDefaultprinter =  Abs(frm.chkNoConfirmMessageSwitchingDefaultprinter)
 .PCXColorscount =  frm.cmbPCXColors.Listindex
 .PDFAllowAssembly =  Abs(frm.chkAllowAssembly.Value)
 .PDFAllowDegradedPrinting =  Abs(frm.chkAllowDegradedPrinting.Value)
 .PDFAllowFillIn =  Abs(frm.chkAllowFillIn.Value)
 .PDFAllowScreenReaders =  Abs(frm.chkAllowScreenReaders.Value)
 .PDFColorsCMYKToRGB =  Abs(frm.chkPDFCMYKtoRGB.Value)
 .PDFColorsColorModel =  frm.cmbPDFColorModel.Listindex
 .PDFColorsPreserveHalftone =  Abs(frm.chkPDFPreserveHalftone.Value)
 .PDFColorsPreserveOverprint =  Abs(frm.chkPDFPreserveOverprint.Value)
 .PDFColorsPreserveTransfer =  Abs(frm.chkPDFPreserveTransfer.Value)
 .PDFCompressionColorCompression =  Abs(frm.chkPDFColorComp.Value)
 .PDFCompressionColorCompressionChoice =  frm.cmbPDFColorComp.Listindex
 .PDFCompressionColorResample =  Abs(frm.chkPDFColorResample.Value)
 .PDFCompressionColorResampleChoice =  frm.cmbPDFColorResample.Listindex
 .PDFCompressionColorResolution =  frm.txtPDFColorRes.Text
 .PDFCompressionGreyCompression =  Abs(frm.chkPDFGreyComp.Value)
 .PDFCompressionGreyCompressionChoice =  frm.cmbPDFGreyComp.Listindex
 .PDFCompressionGreyResample =  Abs(frm.chkPDFGreyResample.Value)
 .PDFCompressionGreyResampleChoice =  frm.cmbPDFGreyResample.Listindex
 .PDFCompressionGreyResolution =  frm.txtPDFGreyRes.Text
 .PDFCompressionMonoCompression =  Abs(frm.chkPDFMonoComp.Value)
 .PDFCompressionMonoCompressionChoice =  frm.cmbPDFMonoComp.Listindex
 .PDFCompressionMonoResample =  Abs(frm.chkPDFMonoResample.Value)
 .PDFCompressionMonoResampleChoice =  frm.cmbPDFMonoResample.Listindex
 .PDFCompressionMonoResolution =  frm.txtPDFMonoRes.Text
 .PDFCompressionTextCompression =  Abs(frm.chkPDFTextComp.Value)
 .PDFDisallowCopy =  Abs(frm.chkAllowCopy.Value)
 .PDFDisallowModifyAnnotations =  Abs(frm.chkAllowModifyAnnotations.Value)
 .PDFDisallowModifyContents =  Abs(frm.chkAllowModifyContents.Value)
 .PDFDisallowPrinting =  Abs(frm.chkAllowPrinting.Value)
 If Frm.cmbPDFEncryptor.ListIndex < 0 Then
   .PDFEncryptor = 0
  Else
   .PDFEncryptor =  frm.cmbPDFEncryptor.Itemdata(Frm.cmbPDFEncryptor.Listindex)
 End If
 .PDFFontsEmbedAll =  Abs(frm.chkPDFEmbedAll.Value)
 .PDFFontsSubSetFonts =  Abs(frm.chkPDFSubSetFonts.Value)
 .PDFFontsSubSetFontsPercent =  frm.txtPDFSubSetPerc.Text
 .PDFGeneralASCII85 =  Abs(frm.chkPDFASCII85.Value)
 .PDFGeneralAutorotate =  frm.cmbPDFRotate.Listindex
 .PDFGeneralCompatibility =  frm.cmbPDFCompat.Listindex
 .PDFGeneralOverprint =  frm.cmbPDFOverprint.Listindex
 .PDFGeneralResolution =  frm.txtPDFRes.Text
 .PDFHighEncryption =  Abs(frm.optEncHigh.Value)
 .PDFLowEncryption =  Abs(frm.optEncLow.Value)
 .PDFOwnerPass =  Abs(frm.chkOwnerPass.Value)
 .PDFUserPass =  Abs(frm.chkUserPass.Value)
 .PDFUseSecurity =  Abs(frm.chkUseSecurity.Value)
 .PNGColorscount =  frm.cmbPNGColors.Listindex
 .PrinterTemppath =  frm.txtTemppath.text
 .ProcessPriority =  frm.sldProcessPriority.value
 .ProgramFont =  frm.cmbFonts.List(frm.cmbFonts.Listindex)
 .ProgramFontCharset =  frm.cmbCharset.Text
 .ProgramFontSize =  frm.txtProgramFontSize.text
 .PSLanguageLevel =  frm.cmbPSLanguageLevel.Listindex
 .RemoveSpaces =  Abs(frm.chkSpaces.Value)
 .SaveFilename =  frm.txtSaveFilename.Text
 .StandardAuthor =  frm.txtStandardAuthor.Text
 .TIFFColorscount =  frm.cmbTIFFColors.Listindex
 .UseAutosave =  Abs(frm.chkUseAutosave.Value)
 .UseAutosaveDirectory =  Abs(frm.chkUseAutosaveDirectory.Value)
 .UseCreationDateNow =  Abs(frm.chkUseCreationDateNow.Value)
 .UseStandardAuthor =  Abs(frm.chkUseStandardAuthor.Value)
 End With
End Sub

Public Sub SetPrinterStop(StopPrinter as Boolean)
 If StopPrinter = True Then
   Options.PrinterStop = 1
   PrinterStop = True
  Else
   Options.PrinterStop = 0
   PrinterStop = False
 End If
 SaveOptions Options
End Sub

Public Sub SetLogging(Logging as Boolean)
 If Logging = True Then
   Options.Logging = 1
  Else
   Options.Logging = 0
 End If
 SaveOptions Options
End Sub

Public Sub SetLanguage(Language as String)
 Options.Language = Language
 SaveOptions Options
End Sub

