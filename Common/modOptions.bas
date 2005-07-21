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
 Dim myOptions As tOptions, reg as clsRegistry
 With myOptions
  .AdditionalGhostscriptParameters = " "
  .AdditionalGhostscriptSearchpath = " "
  .AddWindowsFontpath = "1"
  .AutosaveDirectory = " "
  .AutosaveFilename = "<DateTime>"
  .AutosaveFormat = "0"
  .BitmapResolution = "150"
  .BMPColorscount = "1"
  .ClientComputerResolveIPAddress = "0"
  .DeviceHeightPoints = "-1"
  .DeviceWidthPoints = "-1"
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
  Set reg = New clsRegistry
  reg.hkey = HKEY_LOCAL_MACHINE
  reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
  .DirectoryGhostscriptResource = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryResource"))
  Set reg = Nothing
  .DisableEmail = "0"
  .DontUseDocumentSettings = "0"
  .EPSLanguageLevel = "2"
  .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
  .FilenameSubstitutionsOnlyInTitle = "1"
  .JPEGColorscount = "0"
  .JPEGQuality = "75"
  .Language = "english"
  .LastSaveDirectory = GetMyFiles
  .Logging = "0"
  .LogLines = "100"
  .NoConfirmMessageSwitchingDefaultprinter = "0"
  .NoProcessingAtStartup = "0"
  .OnePagePerFile = "0"
  .OptionsDesign = "1"
  .OptionsEnabled = "1"
  .OptionsVisible = "1"
  .Papersize = " "
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
  .PDFOptimize = "0"
  .PDFOwnerPass = "0"
  .PDFOwnerPasswordString = " "
  .PDFUserPass = "0"
  .PDFUserPasswordString = " "
  .PDFUseSecurity = "0"
  .PNGColorscount = "0"
  .PrintAfterSaving = "0"
  .PrintAfterSavingDuplex = "0"
  .PrintAfterSavingNoCancel = "0"
  .PrintAfterSavingPrinter = " "
  .PrintAfterSavingQueryUser = "0"
  .PrintAfterSavingTumble = "0"
  .PrinterStop = "0"
  .PrinterTemppath = GetTempPath 
  .ProcessPriority = "1"
  .ProgramFont = "MS Sans Serif"
  .ProgramFontCharset = "0"
  .ProgramFontSize = "8"
  .PSLanguageLevel = "2"
  .RemoveAllKnownFileExtensions = "1"
  .RemoveSpaces = "1"
  .RunProgramAfterSaving = "0"
  .RunProgramAfterSavingProgramname = " "
  .RunProgramAfterSavingProgramParameters = " "
  .RunProgramAfterSavingWaitUntilReady = "1"
  .RunProgramAfterSavingWindowstyle = "1"
  .RunProgramBeforeSaving = "0"
  .RunProgramBeforeSavingProgramname = " "
  .RunProgramBeforeSavingProgramParameters = " "
  .RunProgramBeforeSavingWindowstyle = "1"
  .SaveFilename = "<Title>"
  .SendMailMethod = "0"
  .ShowAnimation = "1"
  .StampFontColor = "#FF0000"
  .StampFontname = "Arial"
  .StampFontsize = "48"
  .StampOutlineFontthickness = "0"
  .StampString = " "
  .StampUseOutlineFont = "1"
  .StandardAuthor = " "
  .StandardCreationdate = " "
  .StandardDateformat = "YYYYMMDDHHNNSS"
  .StandardKeywords = " "
  .StandardMailDomain = " "
  .StandardModifydate = " "
  .StandardSaveformat = "pdf"
  .StandardSubject = " "
  .StandardTitle = " "
  .StartStandardProgram = "1"
  .TIFFColorscount = "0"
  .Toolbars = "1"
  .UseAutosave = "0"
  .UseAutosaveDirectory = "1"
  .UseCreationDateNow = "0"
  .UseStandardAuthor = "0"
 End With
 StandardOptions = myOptions
End Function

Public Function ReadOptions(Optional NoMsg As Boolean = False, Optional hProfile As hkey = HKEY_CURRENT_USER) As tOptions
 Dim myOptions As tOptions
 If InstalledAsServer Then
   If UseINI Then
     myOptions = ReadOptionsINI(myOptions, Completepath(GetCommonAppData) & "PDFCreator.ini", NoMsg)
    Else
     myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", HKEY_LOCAL_MACHINE, NoMsg)
   End If
  Else
   If UseINI Then
     If Not IsWin9xMe Then
       myOptions = ReadOptionsINI(myOptions, Completepath(GetDefaultAppData) & "PDFCreator.ini", NoMsg)
       myOptions = ReadOptionsINI(myOptions, PDFCreatorINIFile, NoMsg, False)
      Else
       myOptions = ReadOptionsINI(myOptions, PDFCreatorINIFile, NoMsg)
     End If
     myOptions = ReadOptionsINI(myOptions, Completepath(GetCommonAppData) & "PDFCreator.ini", NoMsg, False)
    Else
     If Not IsWin9xMe Then
       myOptions = ReadOptionsReg(myOptions, ".DEFAULT\Software\PDFCreator", HKEY_USERS, NoMsg)
       myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", hProfile, NoMsg, False)
      Else
       myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", hProfile, NoMsg)
     End If
     myOptions = ReadOptionsReg(myOptions, "Software\PDFCreator", HKEY_LOCAL_MACHINE, NoMsg, False)
   End If
 End If
 ReadOptions = myOptions
End Function

Public Function ReadOptionsINI(myOptions As tOptions, PDFCreatorINIFile As String, Optional NoMsg as Boolean = False, Optional UseStandard as Boolean = True) As tOptions
 Dim ini As clsINI, tStr as String, hOpt As New clsHash
 Set ini = New clsINI
 ini.Filename = PDFCreatorINIFile
 ini.Section = "Options"
 If ini.Checkinifile = False Then
  ReadOptionsINI = StandardOptions
  Exit Function
 End If
 ReadINISection PDFCreatorINIFile, "Options", hOpt
 With myOptions
  tStr = hOpt.Retrieve("AdditionalGhostscriptParameters")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .AdditionalGhostscriptParameters = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .AdditionalGhostscriptParameters = tStr
    End If
  End If
  tStr = hOpt.Retrieve("AdditionalGhostscriptSearchpath")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .AdditionalGhostscriptSearchpath = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .AdditionalGhostscriptSearchpath = tStr
    End If
  End If
  tStr = hOpt.Retrieve("AddWindowsFontpath")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .AddWindowsFontpath = CLng(tStr)
     Else
      If UseStandard Then
       .AddWindowsFontpath = 1
      End If
    End If
   Else
    If UseStandard Then
     .AddWindowsFontpath = 1
    End If
  End If
  tStr = hOpt.Retrieve("AutosaveDirectory")
  If LenB(Trim$(tStr)) > 0 Then
    .AutosaveDirectory = CompletePath(tStr)
   Else
    If UseStandard Then
     tStr = GetMyFiles
     .AutosaveDirectory = CompletePath(tStr)
    End If
  End If
  tStr = hOpt.Retrieve("AutosaveFilename")
  If LenB(tStr) = 0 And LenB("<DateTime>")>0 Then
    If UseStandard Then
     .AutosaveFilename = "<DateTime>"
    End If
   Else
    If LenB(tStr) > 0 Then
     .AutosaveFilename = tStr
    End If
  End If
  tStr = hOpt.Retrieve("AutosaveFormat")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
      .AutosaveFormat = CLng(tStr)
     Else
      If UseStandard Then
       .AutosaveFormat = 0
      End If
    End If
   Else
    If UseStandard Then
     .AutosaveFormat = 0
    End If
  End If
  tStr = hOpt.Retrieve("BitmapResolution")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 1 Then
      .BitmapResolution = CLng(tStr)
     Else
      If UseStandard Then
       .BitmapResolution = 150
      End If
    End If
   Else
    If UseStandard Then
     .BitmapResolution = 150
    End If
  End If
  tStr = hOpt.Retrieve("BMPColorscount")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
      .BMPColorscount = CLng(tStr)
     Else
      If UseStandard Then
       .BMPColorscount = 1
      End If
    End If
   Else
    If UseStandard Then
     .BMPColorscount = 1
    End If
  End If
  tStr = hOpt.Retrieve("ClientComputerResolveIPAddress")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .ClientComputerResolveIPAddress = CLng(tStr)
     Else
      If UseStandard Then
       .ClientComputerResolveIPAddress = 0
      End If
    End If
   Else
    If UseStandard Then
     .ClientComputerResolveIPAddress = 0
    End If
  End If
  tStr = hOpt.Retrieve("DeviceHeightPoints")
  If IsNumeric(tStr) Then
    If CDbl(tStr) >= -1 Then
      .DeviceHeightPoints = CDbl(tStr)
     Else
      If UseStandard Then
       .DeviceHeightPoints = -1
      End If
    End If
   Else
    If UseStandard Then
     .DeviceHeightPoints = -1
    End If
  End If
  tStr = hOpt.Retrieve("DeviceWidthPoints")
  If IsNumeric(tStr) Then
    If CDbl(tStr) >= -1 Then
      .DeviceWidthPoints = CDbl(tStr)
     Else
      If UseStandard Then
       .DeviceWidthPoints = -1
      End If
    End If
   Else
    If UseStandard Then
     .DeviceWidthPoints = -1
    End If
  End If
  tStr = hOpt.Retrieve("DirectoryGhostscriptBinaries")
  If LenB(Trim$(tStr)) > 0 Then
    .DirectoryGhostscriptBinaries = CompletePath(tStr)
   Else
    If UseStandard Then
     tStr = App.Path
     .DirectoryGhostscriptBinaries = CompletePath(tStr)
    End If
  End If
  tStr = hOpt.Retrieve("DirectoryGhostscriptFonts")
  If LenB(Trim$(tStr)) > 0 Then
    .DirectoryGhostscriptFonts = CompletePath(tStr)
   Else
    If UseStandard Then
     tStr = App.Path & "\fonts"
     .DirectoryGhostscriptFonts = CompletePath(tStr)
    End If
  End If
  tStr = hOpt.Retrieve("DirectoryGhostscriptLibraries")
  If LenB(Trim$(tStr)) > 0 Then
    .DirectoryGhostscriptLibraries = CompletePath(tStr)
   Else
    If UseStandard Then
     tStr = App.Path & "\lib"
     .DirectoryGhostscriptLibraries = CompletePath(tStr)
    End If
  End If
  tStr = hOpt.Retrieve("DirectoryGhostscriptResource")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .DirectoryGhostscriptResource = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .DirectoryGhostscriptResource = tStr
    End If
  End If
  tStr = hOpt.Retrieve("DisableEmail")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .DisableEmail = CLng(tStr)
     Else
      If UseStandard Then
       .DisableEmail = 0
      End If
    End If
   Else
    If UseStandard Then
     .DisableEmail = 0
    End If
  End If
  tStr = hOpt.Retrieve("DontUseDocumentSettings")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .DontUseDocumentSettings = CLng(tStr)
     Else
      If UseStandard Then
       .DontUseDocumentSettings = 0
      End If
    End If
   Else
    If UseStandard Then
     .DontUseDocumentSettings = 0
    End If
  End If
  tStr = hOpt.Retrieve("EPSLanguageLevel")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
      .EPSLanguageLevel = CLng(tStr)
     Else
      If UseStandard Then
       .EPSLanguageLevel = 2
      End If
    End If
   Else
    If UseStandard Then
     .EPSLanguageLevel = 2
    End If
  End If
  tStr = hOpt.Retrieve("FilenameSubstitutions")
  If LenB(tStr) = 0 And LenB("Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt")>0 Then
    If UseStandard Then
     .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
    End If
   Else
    If LenB(tStr) > 0 Then
     .FilenameSubstitutions = tStr
    End If
  End If
  tStr = hOpt.Retrieve("FilenameSubstitutionsOnlyInTitle")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
     Else
      If UseStandard Then
       .FilenameSubstitutionsOnlyInTitle = 1
      End If
    End If
   Else
    If UseStandard Then
     .FilenameSubstitutionsOnlyInTitle = 1
    End If
  End If
  tStr = hOpt.Retrieve("JPEGColorscount")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
      .JPEGColorscount = CLng(tStr)
     Else
      If UseStandard Then
       .JPEGColorscount = 0
      End If
    End If
   Else
    If UseStandard Then
     .JPEGColorscount = 0
    End If
  End If
  tStr = hOpt.Retrieve("JPEGQuality")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
      .JPEGQuality = CLng(tStr)
     Else
      If UseStandard Then
       .JPEGQuality = 75
      End If
    End If
   Else
    If UseStandard Then
     .JPEGQuality = 75
    End If
  End If
  tStr = hOpt.Retrieve("Language")
  If LenB(tStr) = 0 And LenB("english")>0 Then
    If UseStandard Then
     .Language = "english"
    End If
   Else
    If LenB(tStr) > 0 Then
     .Language = tStr
    End If
  End If
  tStr = hOpt.Retrieve("LastSaveDirectory")
  If LenB(Trim$(tStr)) > 0 Then
    .LastSaveDirectory = CompletePath(tStr)
   Else
    If UseStandard Then
     tStr = GetMyFiles
     .LastSaveDirectory = CompletePath(tStr)
    End If
  End If
  tStr = hOpt.Retrieve("Logging")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .Logging = CLng(tStr)
     Else
      If UseStandard Then
       .Logging = 0
      End If
    End If
   Else
    If UseStandard Then
     .Logging = 0
    End If
  End If
  tStr = hOpt.Retrieve("LogLines")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
      .LogLines = CLng(tStr)
     Else
      If UseStandard Then
       .LogLines = 100
      End If
    End If
   Else
    If UseStandard Then
     .LogLines = 100
    End If
  End If
  tStr = hOpt.Retrieve("NoConfirmMessageSwitchingDefaultprinter")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
     Else
      If UseStandard Then
       .NoConfirmMessageSwitchingDefaultprinter = 0
      End If
    End If
   Else
    If UseStandard Then
     .NoConfirmMessageSwitchingDefaultprinter = 0
    End If
  End If
  tStr = hOpt.Retrieve("NoProcessingAtStartup")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .NoProcessingAtStartup = CLng(tStr)
     Else
      If UseStandard Then
       .NoProcessingAtStartup = 0
      End If
    End If
   Else
    If UseStandard Then
     .NoProcessingAtStartup = 0
    End If
  End If
  tStr = hOpt.Retrieve("OnePagePerFile")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .OnePagePerFile = CLng(tStr)
     Else
      If UseStandard Then
       .OnePagePerFile = 0
      End If
    End If
   Else
    If UseStandard Then
     .OnePagePerFile = 0
    End If
  End If
  tStr = hOpt.Retrieve("OptionsDesign")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 1 And CLng(tStr) <= 2 Then
      .OptionsDesign = CLng(tStr)
     Else
      If UseStandard Then
       .OptionsDesign = 1
      End If
    End If
   Else
    If UseStandard Then
     .OptionsDesign = 1
    End If
  End If
  tStr = hOpt.Retrieve("OptionsEnabled")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .OptionsEnabled = CLng(tStr)
     Else
      If UseStandard Then
       .OptionsEnabled = 1
      End If
    End If
   Else
    If UseStandard Then
     .OptionsEnabled = 1
    End If
  End If
  tStr = hOpt.Retrieve("OptionsVisible")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .OptionsVisible = CLng(tStr)
     Else
      If UseStandard Then
       .OptionsVisible = 1
      End If
    End If
   Else
    If UseStandard Then
     .OptionsVisible = 1
    End If
  End If
  tStr = hOpt.Retrieve("Papersize")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .Papersize = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .Papersize = tStr
    End If
  End If
  tStr = hOpt.Retrieve("PCXColorscount")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
      .PCXColorscount = CLng(tStr)
     Else
      If UseStandard Then
       .PCXColorscount = 0
      End If
    End If
   Else
    If UseStandard Then
     .PCXColorscount = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFAllowAssembly")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFAllowAssembly = CLng(tStr)
     Else
      If UseStandard Then
       .PDFAllowAssembly = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFAllowAssembly = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFAllowDegradedPrinting")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFAllowDegradedPrinting = CLng(tStr)
     Else
      If UseStandard Then
       .PDFAllowDegradedPrinting = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFAllowDegradedPrinting = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFAllowFillIn")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFAllowFillIn = CLng(tStr)
     Else
      If UseStandard Then
       .PDFAllowFillIn = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFAllowFillIn = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFAllowScreenReaders")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFAllowScreenReaders = CLng(tStr)
     Else
      If UseStandard Then
       .PDFAllowScreenReaders = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFAllowScreenReaders = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFColorsCMYKToRGB")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFColorsCMYKToRGB = CLng(tStr)
     Else
      If UseStandard Then
       .PDFColorsCMYKToRGB = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFColorsCMYKToRGB = 1
    End If
  End If
  tStr = hOpt.Retrieve("PDFColorsColorModel")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
      .PDFColorsColorModel = CLng(tStr)
     Else
      If UseStandard Then
       .PDFColorsColorModel = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFColorsColorModel = 1
    End If
  End If
  tStr = hOpt.Retrieve("PDFColorsPreserveHalftone")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFColorsPreserveHalftone = CLng(tStr)
     Else
      If UseStandard Then
       .PDFColorsPreserveHalftone = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFColorsPreserveHalftone = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFColorsPreserveOverprint")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFColorsPreserveOverprint = CLng(tStr)
     Else
      If UseStandard Then
       .PDFColorsPreserveOverprint = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFColorsPreserveOverprint = 1
    End If
  End If
  tStr = hOpt.Retrieve("PDFColorsPreserveTransfer")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFColorsPreserveTransfer = CLng(tStr)
     Else
      If UseStandard Then
       .PDFColorsPreserveTransfer = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFColorsPreserveTransfer = 1
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionColorCompression")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionColorCompression = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionColorCompression = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionColorCompression = 1
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionColorCompressionChoice")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
      .PDFCompressionColorCompressionChoice = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionColorCompressionChoice = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionColorCompressionChoice = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionColorResample")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionColorResample = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionColorResample = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionColorResample = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionColorResampleChoice")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
      .PDFCompressionColorResampleChoice = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionColorResampleChoice = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionColorResampleChoice = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionColorResolution")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .PDFCompressionColorResolution = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionColorResolution = 300
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionColorResolution = 300
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionGreyCompression")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionGreyCompression = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionGreyCompression = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionGreyCompression = 1
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionGreyCompressionChoice")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
      .PDFCompressionGreyCompressionChoice = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionGreyCompressionChoice = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionGreyCompressionChoice = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionGreyResample")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionGreyResample = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionGreyResample = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionGreyResample = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionGreyResampleChoice")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
      .PDFCompressionGreyResampleChoice = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionGreyResampleChoice = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionGreyResampleChoice = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionGreyResolution")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .PDFCompressionGreyResolution = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionGreyResolution = 300
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionGreyResolution = 300
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionMonoCompression")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionMonoCompression = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionMonoCompression = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionMonoCompression = 1
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionMonoCompressionChoice")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
      .PDFCompressionMonoCompressionChoice = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionMonoCompressionChoice = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionMonoCompressionChoice = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionMonoResample")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionMonoResample = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionMonoResample = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionMonoResample = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionMonoResampleChoice")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
      .PDFCompressionMonoResampleChoice = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionMonoResampleChoice = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionMonoResampleChoice = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionMonoResolution")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .PDFCompressionMonoResolution = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionMonoResolution = 1200
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionMonoResolution = 1200
    End If
  End If
  tStr = hOpt.Retrieve("PDFCompressionTextCompression")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionTextCompression = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionTextCompression = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionTextCompression = 1
    End If
  End If
  tStr = hOpt.Retrieve("PDFDisallowCopy")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFDisallowCopy = CLng(tStr)
     Else
      If UseStandard Then
       .PDFDisallowCopy = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFDisallowCopy = 1
    End If
  End If
  tStr = hOpt.Retrieve("PDFDisallowModifyAnnotations")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFDisallowModifyAnnotations = CLng(tStr)
     Else
      If UseStandard Then
       .PDFDisallowModifyAnnotations = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFDisallowModifyAnnotations = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFDisallowModifyContents")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFDisallowModifyContents = CLng(tStr)
     Else
      If UseStandard Then
       .PDFDisallowModifyContents = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFDisallowModifyContents = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFDisallowPrinting")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFDisallowPrinting = CLng(tStr)
     Else
      If UseStandard Then
       .PDFDisallowPrinting = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFDisallowPrinting = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFEncryptor")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
      .PDFEncryptor = CLng(tStr)
     Else
      If UseStandard Then
       .PDFEncryptor = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFEncryptor = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFFontsEmbedAll")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFFontsEmbedAll = CLng(tStr)
     Else
      If UseStandard Then
       .PDFFontsEmbedAll = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFFontsEmbedAll = 1
    End If
  End If
  tStr = hOpt.Retrieve("PDFFontsSubSetFonts")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFFontsSubSetFonts = CLng(tStr)
     Else
      If UseStandard Then
       .PDFFontsSubSetFonts = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFFontsSubSetFonts = 1
    End If
  End If
  tStr = hOpt.Retrieve("PDFFontsSubSetFontsPercent")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .PDFFontsSubSetFontsPercent = CLng(tStr)
     Else
      If UseStandard Then
       .PDFFontsSubSetFontsPercent = 100
      End If
    End If
   Else
    If UseStandard Then
     .PDFFontsSubSetFontsPercent = 100
    End If
  End If
  tStr = hOpt.Retrieve("PDFGeneralASCII85")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFGeneralASCII85 = CLng(tStr)
     Else
      If UseStandard Then
       .PDFGeneralASCII85 = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFGeneralASCII85 = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFGeneralAutorotate")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
      .PDFGeneralAutorotate = CLng(tStr)
     Else
      If UseStandard Then
       .PDFGeneralAutorotate = 2
      End If
    End If
   Else
    If UseStandard Then
     .PDFGeneralAutorotate = 2
    End If
  End If
  tStr = hOpt.Retrieve("PDFGeneralCompatibility")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
      .PDFGeneralCompatibility = CLng(tStr)
     Else
      If UseStandard Then
       .PDFGeneralCompatibility = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFGeneralCompatibility = 1
    End If
  End If
  tStr = hOpt.Retrieve("PDFGeneralOverprint")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
      .PDFGeneralOverprint = CLng(tStr)
     Else
      If UseStandard Then
       .PDFGeneralOverprint = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFGeneralOverprint = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFGeneralResolution")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .PDFGeneralResolution = CLng(tStr)
     Else
      If UseStandard Then
       .PDFGeneralResolution = 600
      End If
    End If
   Else
    If UseStandard Then
     .PDFGeneralResolution = 600
    End If
  End If
  tStr = hOpt.Retrieve("PDFHighEncryption")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFHighEncryption = CLng(tStr)
     Else
      If UseStandard Then
       .PDFHighEncryption = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFHighEncryption = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFLowEncryption")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFLowEncryption = CLng(tStr)
     Else
      If UseStandard Then
       .PDFLowEncryption = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFLowEncryption = 1
    End If
  End If
  tStr = hOpt.Retrieve("PDFOptimize")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFOptimize = CLng(tStr)
     Else
      If UseStandard Then
       .PDFOptimize = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFOptimize = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFOwnerPass")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFOwnerPass = CLng(tStr)
     Else
      If UseStandard Then
       .PDFOwnerPass = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFOwnerPass = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFOwnerPasswordString")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .PDFOwnerPasswordString = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .PDFOwnerPasswordString = tStr
    End If
  End If
  tStr = hOpt.Retrieve("PDFUserPass")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFUserPass = CLng(tStr)
     Else
      If UseStandard Then
       .PDFUserPass = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFUserPass = 0
    End If
  End If
  tStr = hOpt.Retrieve("PDFUserPasswordString")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .PDFUserPasswordString = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .PDFUserPasswordString = tStr
    End If
  End If
  tStr = hOpt.Retrieve("PDFUseSecurity")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFUseSecurity = CLng(tStr)
     Else
      If UseStandard Then
       .PDFUseSecurity = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFUseSecurity = 0
    End If
  End If
  tStr = hOpt.Retrieve("PNGColorscount")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
      .PNGColorscount = CLng(tStr)
     Else
      If UseStandard Then
       .PNGColorscount = 0
      End If
    End If
   Else
    If UseStandard Then
     .PNGColorscount = 0
    End If
  End If
  tStr = hOpt.Retrieve("PrintAfterSaving")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PrintAfterSaving = CLng(tStr)
     Else
      If UseStandard Then
       .PrintAfterSaving = 0
      End If
    End If
   Else
    If UseStandard Then
     .PrintAfterSaving = 0
    End If
  End If
  tStr = hOpt.Retrieve("PrintAfterSavingDuplex")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PrintAfterSavingDuplex = CLng(tStr)
     Else
      If UseStandard Then
       .PrintAfterSavingDuplex = 0
      End If
    End If
   Else
    If UseStandard Then
     .PrintAfterSavingDuplex = 0
    End If
  End If
  tStr = hOpt.Retrieve("PrintAfterSavingNoCancel")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PrintAfterSavingNoCancel = CLng(tStr)
     Else
      If UseStandard Then
       .PrintAfterSavingNoCancel = 0
      End If
    End If
   Else
    If UseStandard Then
     .PrintAfterSavingNoCancel = 0
    End If
  End If
  tStr = hOpt.Retrieve("PrintAfterSavingPrinter")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .PrintAfterSavingPrinter = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .PrintAfterSavingPrinter = tStr
    End If
  End If
  tStr = hOpt.Retrieve("PrintAfterSavingQueryUser")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
      .PrintAfterSavingQueryUser = CLng(tStr)
     Else
      If UseStandard Then
       .PrintAfterSavingQueryUser = 0
      End If
    End If
   Else
    If UseStandard Then
     .PrintAfterSavingQueryUser = 0
    End If
  End If
  tStr = hOpt.Retrieve("PrintAfterSavingTumble")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PrintAfterSavingTumble = CLng(tStr)
     Else
      If UseStandard Then
       .PrintAfterSavingTumble = 0
      End If
    End If
   Else
    If UseStandard Then
     .PrintAfterSavingTumble = 0
    End If
  End If
  tStr = hOpt.Retrieve("PrinterStop")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PrinterStop = CLng(tStr)
     Else
      If UseStandard Then
       .PrinterStop = 0
      End If
    End If
   Else
    If UseStandard Then
     .PrinterStop = 0
    End If
  End If
  tStr = hOpt.Retrieve("PrinterTemppath")
  If LenB(Trim$(tstr)) > 0 Then
   If DirExists(GetSubstFilename2(tstr, False)) = True Then
     .PrinterTemppath = tStr
    Else
     MakePath ResolveEnvironment(GetSubstFilename2(tstr, False))
     If DirExists(ResolveEnvironment(GetSubstFilename2(tstr, False))) = False Then
       If UseStandard Then
         .PrinterTemppath = GetTempPath
        Else
         .PrinterTemppath = ""
         If NoMsg = False Then
          MsgBox "PrinterTemppath: '" & tstr & "' = '" & ResolveEnvironment(GetSubstFilename2(tstr, False)) & "'" & _
           vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
         End If
       End If
      Else
       .PrinterTemppath = tstr
     End If
   End If
  End If
  tStr = hOpt.Retrieve("ProcessPriority")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
      .ProcessPriority = CLng(tStr)
     Else
      If UseStandard Then
       .ProcessPriority = 1
      End If
    End If
   Else
    If UseStandard Then
     .ProcessPriority = 1
    End If
  End If
  tStr = hOpt.Retrieve("ProgramFont")
  If LenB(tStr) = 0 And LenB("MS Sans Serif")>0 Then
    If UseStandard Then
     .ProgramFont = "MS Sans Serif"
    End If
   Else
    If LenB(tStr) > 0 Then
     .ProgramFont = tStr
    End If
  End If
  tStr = hOpt.Retrieve("ProgramFontCharset")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .ProgramFontCharset = CLng(tStr)
     Else
      If UseStandard Then
       .ProgramFontCharset = 0
      End If
    End If
   Else
    If UseStandard Then
     .ProgramFontCharset = 0
    End If
  End If
  tStr = hOpt.Retrieve("ProgramFontSize")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 1 And CLng(tStr) <= 72 Then
      .ProgramFontSize = CLng(tStr)
     Else
      If UseStandard Then
       .ProgramFontSize = 8
      End If
    End If
   Else
    If UseStandard Then
     .ProgramFontSize = 8
    End If
  End If
  tStr = hOpt.Retrieve("PSLanguageLevel")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
      .PSLanguageLevel = CLng(tStr)
     Else
      If UseStandard Then
       .PSLanguageLevel = 2
      End If
    End If
   Else
    If UseStandard Then
     .PSLanguageLevel = 2
    End If
  End If
  tStr = hOpt.Retrieve("RemoveAllKnownFileExtensions")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .RemoveAllKnownFileExtensions = CLng(tStr)
     Else
      If UseStandard Then
       .RemoveAllKnownFileExtensions = 1
      End If
    End If
   Else
    If UseStandard Then
     .RemoveAllKnownFileExtensions = 1
    End If
  End If
  tStr = hOpt.Retrieve("RemoveSpaces")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .RemoveSpaces = CLng(tStr)
     Else
      If UseStandard Then
       .RemoveSpaces = 1
      End If
    End If
   Else
    If UseStandard Then
     .RemoveSpaces = 1
    End If
  End If
  tStr = hOpt.Retrieve("RunProgramAfterSaving")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .RunProgramAfterSaving = CLng(tStr)
     Else
      If UseStandard Then
       .RunProgramAfterSaving = 0
      End If
    End If
   Else
    If UseStandard Then
     .RunProgramAfterSaving = 0
    End If
  End If
  tStr = hOpt.Retrieve("RunProgramAfterSavingProgramname")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .RunProgramAfterSavingProgramname = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .RunProgramAfterSavingProgramname = tStr
    End If
  End If
  tStr = hOpt.Retrieve("RunProgramAfterSavingProgramParameters")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .RunProgramAfterSavingProgramParameters = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .RunProgramAfterSavingProgramParameters = tStr
    End If
  End If
  tStr = hOpt.Retrieve("RunProgramAfterSavingWaitUntilReady")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
     Else
      If UseStandard Then
       .RunProgramAfterSavingWaitUntilReady = 1
      End If
    End If
   Else
    If UseStandard Then
     .RunProgramAfterSavingWaitUntilReady = 1
    End If
  End If
  tStr = hOpt.Retrieve("RunProgramAfterSavingWindowstyle")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
      .RunProgramAfterSavingWindowstyle = CLng(tStr)
     Else
      If UseStandard Then
       .RunProgramAfterSavingWindowstyle = 1
      End If
    End If
   Else
    If UseStandard Then
     .RunProgramAfterSavingWindowstyle = 1
    End If
  End If
  tStr = hOpt.Retrieve("RunProgramBeforeSaving")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .RunProgramBeforeSaving = CLng(tStr)
     Else
      If UseStandard Then
       .RunProgramBeforeSaving = 0
      End If
    End If
   Else
    If UseStandard Then
     .RunProgramBeforeSaving = 0
    End If
  End If
  tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramname")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .RunProgramBeforeSavingProgramname = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .RunProgramBeforeSavingProgramname = tStr
    End If
  End If
  tStr = hOpt.Retrieve("RunProgramBeforeSavingProgramParameters")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .RunProgramBeforeSavingProgramParameters = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .RunProgramBeforeSavingProgramParameters = tStr
    End If
  End If
  tStr = hOpt.Retrieve("RunProgramBeforeSavingWindowstyle")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
      .RunProgramBeforeSavingWindowstyle = CLng(tStr)
     Else
      If UseStandard Then
       .RunProgramBeforeSavingWindowstyle = 1
      End If
    End If
   Else
    If UseStandard Then
     .RunProgramBeforeSavingWindowstyle = 1
    End If
  End If
  tStr = hOpt.Retrieve("SaveFilename")
  If LenB(tStr) = 0 And LenB("<Title>")>0 Then
    If UseStandard Then
     .SaveFilename = "<Title>"
    End If
   Else
    If LenB(tStr) > 0 Then
     .SaveFilename = tStr
    End If
  End If
  tStr = hOpt.Retrieve("SendMailMethod")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .SendMailMethod = CLng(tStr)
     Else
      If UseStandard Then
       .SendMailMethod = 0
      End If
    End If
   Else
    If UseStandard Then
     .SendMailMethod = 0
    End If
  End If
  tStr = hOpt.Retrieve("ShowAnimation")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .ShowAnimation = CLng(tStr)
     Else
      If UseStandard Then
       .ShowAnimation = 1
      End If
    End If
   Else
    If UseStandard Then
     .ShowAnimation = 1
    End If
  End If
  tStr = hOpt.Retrieve("StampFontColor")
  If LenB(tStr) = 0 And LenB("#FF0000")>0 Then
    If UseStandard Then
     .StampFontColor = "#FF0000"
    End If
   Else
    If LenB(tStr) > 0 Then
     .StampFontColor = tStr
    End If
  End If
  tStr = hOpt.Retrieve("StampFontname")
  If LenB(tStr) = 0 And LenB("Arial")>0 Then
    If UseStandard Then
     .StampFontname = "Arial"
    End If
   Else
    If LenB(tStr) > 0 Then
     .StampFontname = tStr
    End If
  End If
  tStr = hOpt.Retrieve("StampFontsize")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 1 Then
      .StampFontsize = CLng(tStr)
     Else
      If UseStandard Then
       .StampFontsize = 48
      End If
    End If
   Else
    If UseStandard Then
     .StampFontsize = 48
    End If
  End If
  tStr = hOpt.Retrieve("StampOutlineFontthickness")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .StampOutlineFontthickness = CLng(tStr)
     Else
      If UseStandard Then
       .StampOutlineFontthickness = 0
      End If
    End If
   Else
    If UseStandard Then
     .StampOutlineFontthickness = 0
    End If
  End If
  tStr = hOpt.Retrieve("StampString")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StampString = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StampString = tStr
    End If
  End If
  tStr = hOpt.Retrieve("StampUseOutlineFont")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .StampUseOutlineFont = CLng(tStr)
     Else
      If UseStandard Then
       .StampUseOutlineFont = 1
      End If
    End If
   Else
    If UseStandard Then
     .StampUseOutlineFont = 1
    End If
  End If
  tStr = hOpt.Retrieve("StandardAuthor")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardAuthor = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardAuthor = tStr
    End If
  End If
  tStr = hOpt.Retrieve("StandardCreationdate")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardCreationdate = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardCreationdate = tStr
    End If
  End If
  tStr = hOpt.Retrieve("StandardDateformat")
  If LenB(tStr) = 0 And LenB("YYYYMMDDHHNNSS")>0 Then
    If UseStandard Then
     .StandardDateformat = "YYYYMMDDHHNNSS"
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardDateformat = tStr
    End If
  End If
  tStr = hOpt.Retrieve("StandardKeywords")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardKeywords = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardKeywords = tStr
    End If
  End If
  tStr = hOpt.Retrieve("StandardMailDomain")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardMailDomain = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardMailDomain = tStr
    End If
  End If
  tStr = hOpt.Retrieve("StandardModifydate")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardModifydate = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardModifydate = tStr
    End If
  End If
  tStr = hOpt.Retrieve("StandardSaveformat")
  If LenB(tStr) = 0 And LenB("pdf")>0 Then
    If UseStandard Then
     .StandardSaveformat = "pdf"
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardSaveformat = tStr
    End If
  End If
  tStr = hOpt.Retrieve("StandardSubject")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardSubject = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardSubject = tStr
    End If
  End If
  tStr = hOpt.Retrieve("StandardTitle")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardTitle = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardTitle = tStr
    End If
  End If
  tStr = hOpt.Retrieve("StartStandardProgram")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .StartStandardProgram = CLng(tStr)
     Else
      If UseStandard Then
       .StartStandardProgram = 1
      End If
    End If
   Else
    If UseStandard Then
     .StartStandardProgram = 1
    End If
  End If
  tStr = hOpt.Retrieve("TIFFColorscount")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
      .TIFFColorscount = CLng(tStr)
     Else
      If UseStandard Then
       .TIFFColorscount = 0
      End If
    End If
   Else
    If UseStandard Then
     .TIFFColorscount = 0
    End If
  End If
  tStr = hOpt.Retrieve("Toolbars")
  If IsNumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .Toolbars = CLng(tStr)
     Else
      If UseStandard Then
       .Toolbars = 1
      End If
    End If
   Else
    If UseStandard Then
     .Toolbars = 1
    End If
  End If
  tStr = hOpt.Retrieve("UseAutosave")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .UseAutosave = CLng(tStr)
     Else
      If UseStandard Then
       .UseAutosave = 0
      End If
    End If
   Else
    If UseStandard Then
     .UseAutosave = 0
    End If
  End If
  tStr = hOpt.Retrieve("UseAutosaveDirectory")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .UseAutosaveDirectory = CLng(tStr)
     Else
      If UseStandard Then
       .UseAutosaveDirectory = 1
      End If
    End If
   Else
    If UseStandard Then
     .UseAutosaveDirectory = 1
    End If
  End If
  tStr = hOpt.Retrieve("UseCreationDateNow")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .UseCreationDateNow = CLng(tStr)
     Else
      If UseStandard Then
       .UseCreationDateNow = 0
      End If
    End If
   Else
    If UseStandard Then
     .UseCreationDateNow = 0
    End If
  End If
  tStr = hOpt.Retrieve("UseStandardAuthor")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .UseStandardAuthor = CLng(tStr)
     Else
      If UseStandard Then
       .UseStandardAuthor = 0
      End If
    End If
   Else
    If UseStandard Then
     .UseStandardAuthor = 0
    End If
  End If
 End With
 Set ini = Nothing
 ReadOptionsINI = myOptions
End Function

Public Sub SaveOptions(sOptions as tOptions)
 If UseINI Then
   SaveOptionsINI sOptions
  Else
   SaveOptionsReg sOptions
 End If
End Sub

Public Sub SaveOptionsINI(sOptions as tOptions)
 Dim ini As clsINI
 Set ini = New clsINI
 ini.Filename = PDFCreatorINIFile
 ini.Section = "Options"
 If ini.CheckInifile = False Then
  ini.CreateInifile
 End If
 With sOptions
  ini.SaveKey CStr(.AdditionalGhostscriptParameters), "AdditionalGhostscriptParameters"
  ini.SaveKey CStr(.AdditionalGhostscriptSearchpath), "AdditionalGhostscriptSearchpath"
  ini.SaveKey CStr(Abs(.AddWindowsFontpath)), "AddWindowsFontpath"
  ini.SaveKey CStr(.AutosaveDirectory), "AutosaveDirectory"
  ini.SaveKey CStr(.AutosaveFilename), "AutosaveFilename"
  ini.SaveKey CStr(.AutosaveFormat), "AutosaveFormat"
  ini.SaveKey CStr(.BitmapResolution), "BitmapResolution"
  ini.SaveKey CStr(.BMPColorscount), "BMPColorscount"
  ini.SaveKey CStr(Abs(.ClientComputerResolveIPAddress)), "ClientComputerResolveIPAddress"
  ini.SaveKey CStr(.DeviceHeightPoints), "DeviceHeightPoints"
  ini.SaveKey CStr(.DeviceWidthPoints), "DeviceWidthPoints"
  ini.SaveKey CStr(.DirectoryGhostscriptBinaries), "DirectoryGhostscriptBinaries"
  ini.SaveKey CStr(.DirectoryGhostscriptFonts), "DirectoryGhostscriptFonts"
  ini.SaveKey CStr(.DirectoryGhostscriptLibraries), "DirectoryGhostscriptLibraries"
  ini.SaveKey CStr(.DirectoryGhostscriptResource), "DirectoryGhostscriptResource"
  ini.SaveKey CStr(Abs(.DisableEmail)), "DisableEmail"
  ini.SaveKey CStr(Abs(.DontUseDocumentSettings)), "DontUseDocumentSettings"
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
  ini.SaveKey CStr(Abs(.NoProcessingAtStartup)), "NoProcessingAtStartup"
  ini.SaveKey CStr(Abs(.OnePagePerFile)), "OnePagePerFile"
  ini.SaveKey CStr(.OptionsDesign), "OptionsDesign"
  ini.SaveKey CStr(Abs(.OptionsEnabled)), "OptionsEnabled"
  ini.SaveKey CStr(Abs(.OptionsVisible)), "OptionsVisible"
  ini.SaveKey CStr(.Papersize), "Papersize"
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
  ini.SaveKey CStr(Abs(.PDFOptimize)), "PDFOptimize"
  ini.SaveKey CStr(Abs(.PDFOwnerPass)), "PDFOwnerPass"
  ini.SaveKey CStr(.PDFOwnerPasswordString), "PDFOwnerPasswordString"
  ini.SaveKey CStr(Abs(.PDFUserPass)), "PDFUserPass"
  ini.SaveKey CStr(.PDFUserPasswordString), "PDFUserPasswordString"
  ini.SaveKey CStr(Abs(.PDFUseSecurity)), "PDFUseSecurity"
  ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
  ini.SaveKey CStr(Abs(.PrintAfterSaving)), "PrintAfterSaving"
  ini.SaveKey CStr(Abs(.PrintAfterSavingDuplex)), "PrintAfterSavingDuplex"
  ini.SaveKey CStr(Abs(.PrintAfterSavingNoCancel)), "PrintAfterSavingNoCancel"
  ini.SaveKey CStr(.PrintAfterSavingPrinter), "PrintAfterSavingPrinter"
  ini.SaveKey CStr(.PrintAfterSavingQueryUser), "PrintAfterSavingQueryUser"
  ini.SaveKey CStr(Abs(.PrintAfterSavingTumble)), "PrintAfterSavingTumble"
  ini.SaveKey CStr(Abs(.PrinterStop)), "PrinterStop"
  ini.SaveKey CStr(.PrinterTemppath), "PrinterTemppath"
  ini.SaveKey CStr(.ProcessPriority), "ProcessPriority"
  ini.SaveKey CStr(.ProgramFont), "ProgramFont"
  ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
  ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
  ini.SaveKey CStr(.PSLanguageLevel), "PSLanguageLevel"
  ini.SaveKey CStr(Abs(.RemoveAllKnownFileExtensions)), "RemoveAllKnownFileExtensions"
  ini.SaveKey CStr(Abs(.RemoveSpaces)), "RemoveSpaces"
  ini.SaveKey CStr(Abs(.RunProgramAfterSaving)), "RunProgramAfterSaving"
  ini.SaveKey CStr(.RunProgramAfterSavingProgramname), "RunProgramAfterSavingProgramname"
  ini.SaveKey CStr(.RunProgramAfterSavingProgramParameters), "RunProgramAfterSavingProgramParameters"
  ini.SaveKey CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), "RunProgramAfterSavingWaitUntilReady"
  ini.SaveKey CStr(.RunProgramAfterSavingWindowstyle), "RunProgramAfterSavingWindowstyle"
  ini.SaveKey CStr(Abs(.RunProgramBeforeSaving)), "RunProgramBeforeSaving"
  ini.SaveKey CStr(.RunProgramBeforeSavingProgramname), "RunProgramBeforeSavingProgramname"
  ini.SaveKey CStr(.RunProgramBeforeSavingProgramParameters), "RunProgramBeforeSavingProgramParameters"
  ini.SaveKey CStr(.RunProgramBeforeSavingWindowstyle), "RunProgramBeforeSavingWindowstyle"
  ini.SaveKey CStr(.SaveFilename), "SaveFilename"
  ini.SaveKey CStr(.SendMailMethod), "SendMailMethod"
  ini.SaveKey CStr(Abs(.ShowAnimation)), "ShowAnimation"
  ini.SaveKey CStr(.StampFontColor), "StampFontColor"
  ini.SaveKey CStr(.StampFontname), "StampFontname"
  ini.SaveKey CStr(.StampFontsize), "StampFontsize"
  ini.SaveKey CStr(.StampOutlineFontthickness), "StampOutlineFontthickness"
  ini.SaveKey CStr(.StampString), "StampString"
  ini.SaveKey CStr(Abs(.StampUseOutlineFont)), "StampUseOutlineFont"
  ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
  ini.SaveKey CStr(.StandardCreationdate), "StandardCreationdate"
  ini.SaveKey CStr(.StandardDateformat), "StandardDateformat"
  ini.SaveKey CStr(.StandardKeywords), "StandardKeywords"
  ini.SaveKey CStr(.StandardMailDomain), "StandardMailDomain"
  ini.SaveKey CStr(.StandardModifydate), "StandardModifydate"
  ini.SaveKey CStr(.StandardSaveformat), "StandardSaveformat"
  ini.SaveKey CStr(.StandardSubject), "StandardSubject"
  ini.SaveKey CStr(.StandardTitle), "StandardTitle"
  ini.SaveKey CStr(Abs(.StartStandardProgram)), "StartStandardProgram"
  ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
  ini.SaveKey CStr(.Toolbars), "Toolbars"
  ini.SaveKey CStr(Abs(.UseAutosave)), "UseAutosave"
  ini.SaveKey CStr(Abs(.UseAutosaveDirectory)), "UseAutosaveDirectory"
  ini.SaveKey CStr(Abs(.UseCreationDateNow)), "UseCreationDateNow"
  ini.SaveKey CStr(Abs(.UseStandardAuthor)), "UseStandardAuthor"
 End With
 Set ini = Nothing
End Sub

Public Function ReadOptionsReg(myOptions As tOptions, KeyRoot as String, Optional hkey1 as hkey = HKEY_CURRENT_USER, Optional NoMsg as Boolean = False, Optional UseStandard as Boolean = True) As tOptions
 Dim reg As clsRegistry, tStr as String
 Set reg = New clsRegistry
 reg.hkey = hkey1
 reg.KeyRoot = KeyRoot
 With myOptions
  reg.Subkey = "Ghostscript"
  tStr = reg.GetRegistryValue("DirectoryGhostscriptBinaries")
  If LenB(Trim$(tStr)) > 0 Then
    .DirectoryGhostscriptBinaries = CompletePath(tStr)
   Else
    If UseStandard Then
     tStr = App.Path
     .DirectoryGhostscriptBinaries = CompletePath(tStr)
    End If
  End If
  tStr = reg.GetRegistryValue("DirectoryGhostscriptFonts")
  If LenB(Trim$(tStr)) > 0 Then
    .DirectoryGhostscriptFonts = CompletePath(tStr)
   Else
    If UseStandard Then
     tStr = App.Path & "\fonts"
     .DirectoryGhostscriptFonts = CompletePath(tStr)
    End If
  End If
  tStr = reg.GetRegistryValue("DirectoryGhostscriptLibraries")
  If LenB(Trim$(tStr)) > 0 Then
    .DirectoryGhostscriptLibraries = CompletePath(tStr)
   Else
    If UseStandard Then
     tStr = App.Path & "\lib"
     .DirectoryGhostscriptLibraries = CompletePath(tStr)
    End If
  End If
  tStr = reg.GetRegistryValue("DirectoryGhostscriptResource")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .DirectoryGhostscriptResource = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .DirectoryGhostscriptResource = tStr
    End If
  End If
  reg.Subkey = "Printing"
  tStr = reg.GetRegistryValue("DeviceHeightPoints")
  If IsNumeric(tStr) Then
    If CDbl(tStr) >= -1 Then
      .DeviceHeightPoints = CDbl(tStr)
     Else
      If UseStandard Then
       .DeviceHeightPoints = -1
      End If
    End If
   Else
    If UseStandard Then
     .DeviceHeightPoints = -1
    End If
  End If
  tStr = reg.GetRegistryValue("DeviceWidthPoints")
  If IsNumeric(tStr) Then
    If CDbl(tStr) >= -1 Then
      .DeviceWidthPoints = CDbl(tStr)
     Else
      If UseStandard Then
       .DeviceWidthPoints = -1
      End If
    End If
   Else
    If UseStandard Then
     .DeviceWidthPoints = -1
    End If
  End If
  tStr = reg.GetRegistryValue("OnePagePerFile")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .OnePagePerFile = CLng(tStr)
     Else
      If UseStandard Then
       .OnePagePerFile = 0
      End If
    End If
   Else
    If UseStandard Then
     .OnePagePerFile = 0
    End If
  End If
  tStr = reg.GetRegistryValue("Papersize")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .Papersize = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .Papersize = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("StampFontColor")
  If LenB(tStr) = 0 And LenB("#FF0000")>0 Then
    If UseStandard Then
     .StampFontColor = "#FF0000"
    End If
   Else
    If LenB(tStr) > 0 Then
     .StampFontColor = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("StampFontname")
  If LenB(tStr) = 0 And LenB("Arial")>0 Then
    If UseStandard Then
     .StampFontname = "Arial"
    End If
   Else
    If LenB(tStr) > 0 Then
     .StampFontname = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("StampFontsize")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 1 Then
      .StampFontsize = CLng(tStr)
     Else
      If UseStandard Then
       .StampFontsize = 48
      End If
    End If
   Else
    If UseStandard Then
     .StampFontsize = 48
    End If
  End If
  tStr = reg.GetRegistryValue("StampOutlineFontthickness")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .StampOutlineFontthickness = CLng(tStr)
     Else
      If UseStandard Then
       .StampOutlineFontthickness = 0
      End If
    End If
   Else
    If UseStandard Then
     .StampOutlineFontthickness = 0
    End If
  End If
  tStr = reg.GetRegistryValue("StampString")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StampString = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StampString = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("StampUseOutlineFont")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .StampUseOutlineFont = CLng(tStr)
     Else
      If UseStandard Then
       .StampUseOutlineFont = 1
      End If
    End If
   Else
    If UseStandard Then
     .StampUseOutlineFont = 1
    End If
  End If
  tStr = reg.GetRegistryValue("StandardAuthor")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardAuthor = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardAuthor = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("StandardCreationdate")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardCreationdate = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardCreationdate = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("StandardDateformat")
  If LenB(tStr) = 0 And LenB("YYYYMMDDHHNNSS")>0 Then
    If UseStandard Then
     .StandardDateformat = "YYYYMMDDHHNNSS"
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardDateformat = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("StandardKeywords")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardKeywords = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardKeywords = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("StandardMailDomain")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardMailDomain = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardMailDomain = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("StandardModifydate")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardModifydate = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardModifydate = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("StandardSaveformat")
  If LenB(tStr) = 0 And LenB("pdf")>0 Then
    If UseStandard Then
     .StandardSaveformat = "pdf"
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardSaveformat = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("StandardSubject")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardSubject = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardSubject = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("StandardTitle")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .StandardTitle = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .StandardTitle = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("UseCreationDateNow")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .UseCreationDateNow = CLng(tStr)
     Else
      If UseStandard Then
       .UseCreationDateNow = 0
      End If
    End If
   Else
    If UseStandard Then
     .UseCreationDateNow = 0
    End If
  End If
  tStr = reg.GetRegistryValue("UseStandardAuthor")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .UseStandardAuthor = CLng(tStr)
     Else
      If UseStandard Then
       .UseStandardAuthor = 0
      End If
    End If
   Else
    If UseStandard Then
     .UseStandardAuthor = 0
    End If
  End If
  reg.Subkey = "Printing\Formats\Bitmap\Colors"
  tStr = reg.GetRegistryValue("BitmapResolution")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 1 Then
      .BitmapResolution = CLng(tStr)
     Else
      If UseStandard Then
       .BitmapResolution = 150
      End If
    End If
   Else
    If UseStandard Then
     .BitmapResolution = 150
    End If
  End If
  tStr = reg.GetRegistryValue("BMPColorscount")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
      .BMPColorscount = CLng(tStr)
     Else
      If UseStandard Then
       .BMPColorscount = 1
      End If
    End If
   Else
    If UseStandard Then
     .BMPColorscount = 1
    End If
  End If
  tStr = reg.GetRegistryValue("JPEGColorscount")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
      .JPEGColorscount = CLng(tStr)
     Else
      If UseStandard Then
       .JPEGColorscount = 0
      End If
    End If
   Else
    If UseStandard Then
     .JPEGColorscount = 0
    End If
  End If
  tStr = reg.GetRegistryValue("JPEGQuality")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
      .JPEGQuality = CLng(tStr)
     Else
      If UseStandard Then
       .JPEGQuality = 75
      End If
    End If
   Else
    If UseStandard Then
     .JPEGQuality = 75
    End If
  End If
  tStr = reg.GetRegistryValue("PCXColorscount")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
      .PCXColorscount = CLng(tStr)
     Else
      If UseStandard Then
       .PCXColorscount = 0
      End If
    End If
   Else
    If UseStandard Then
     .PCXColorscount = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PNGColorscount")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
      .PNGColorscount = CLng(tStr)
     Else
      If UseStandard Then
       .PNGColorscount = 0
      End If
    End If
   Else
    If UseStandard Then
     .PNGColorscount = 0
    End If
  End If
  tStr = reg.GetRegistryValue("TIFFColorscount")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
      .TIFFColorscount = CLng(tStr)
     Else
      If UseStandard Then
       .TIFFColorscount = 0
      End If
    End If
   Else
    If UseStandard Then
     .TIFFColorscount = 0
    End If
  End If
  reg.Subkey = "Printing\Formats\PDF\Colors"
  tStr = reg.GetRegistryValue("PDFColorsCMYKToRGB")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFColorsCMYKToRGB = CLng(tStr)
     Else
      If UseStandard Then
       .PDFColorsCMYKToRGB = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFColorsCMYKToRGB = 1
    End If
  End If
  tStr = reg.GetRegistryValue("PDFColorsColorModel")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
      .PDFColorsColorModel = CLng(tStr)
     Else
      If UseStandard Then
       .PDFColorsColorModel = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFColorsColorModel = 1
    End If
  End If
  tStr = reg.GetRegistryValue("PDFColorsPreserveHalftone")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFColorsPreserveHalftone = CLng(tStr)
     Else
      If UseStandard Then
       .PDFColorsPreserveHalftone = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFColorsPreserveHalftone = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFColorsPreserveOverprint")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFColorsPreserveOverprint = CLng(tStr)
     Else
      If UseStandard Then
       .PDFColorsPreserveOverprint = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFColorsPreserveOverprint = 1
    End If
  End If
  tStr = reg.GetRegistryValue("PDFColorsPreserveTransfer")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFColorsPreserveTransfer = CLng(tStr)
     Else
      If UseStandard Then
       .PDFColorsPreserveTransfer = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFColorsPreserveTransfer = 1
    End If
  End If
  reg.Subkey = "Printing\Formats\PDF\Compression"
  tStr = reg.GetRegistryValue("PDFCompressionColorCompression")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionColorCompression = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionColorCompression = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionColorCompression = 1
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionColorCompressionChoice")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
      .PDFCompressionColorCompressionChoice = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionColorCompressionChoice = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionColorCompressionChoice = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionColorResample")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionColorResample = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionColorResample = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionColorResample = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionColorResampleChoice")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
      .PDFCompressionColorResampleChoice = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionColorResampleChoice = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionColorResampleChoice = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionColorResolution")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .PDFCompressionColorResolution = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionColorResolution = 300
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionColorResolution = 300
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionGreyCompression")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionGreyCompression = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionGreyCompression = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionGreyCompression = 1
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionGreyCompressionChoice")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
      .PDFCompressionGreyCompressionChoice = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionGreyCompressionChoice = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionGreyCompressionChoice = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionGreyResample")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionGreyResample = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionGreyResample = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionGreyResample = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionGreyResampleChoice")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
      .PDFCompressionGreyResampleChoice = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionGreyResampleChoice = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionGreyResampleChoice = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionGreyResolution")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .PDFCompressionGreyResolution = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionGreyResolution = 300
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionGreyResolution = 300
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionMonoCompression")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionMonoCompression = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionMonoCompression = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionMonoCompression = 1
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionMonoCompressionChoice")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
      .PDFCompressionMonoCompressionChoice = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionMonoCompressionChoice = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionMonoCompressionChoice = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionMonoResample")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionMonoResample = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionMonoResample = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionMonoResample = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionMonoResampleChoice")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
      .PDFCompressionMonoResampleChoice = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionMonoResampleChoice = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionMonoResampleChoice = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionMonoResolution")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .PDFCompressionMonoResolution = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionMonoResolution = 1200
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionMonoResolution = 1200
    End If
  End If
  tStr = reg.GetRegistryValue("PDFCompressionTextCompression")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFCompressionTextCompression = CLng(tStr)
     Else
      If UseStandard Then
       .PDFCompressionTextCompression = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFCompressionTextCompression = 1
    End If
  End If
  reg.Subkey = "Printing\Formats\PDF\Fonts"
  tStr = reg.GetRegistryValue("PDFFontsEmbedAll")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFFontsEmbedAll = CLng(tStr)
     Else
      If UseStandard Then
       .PDFFontsEmbedAll = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFFontsEmbedAll = 1
    End If
  End If
  tStr = reg.GetRegistryValue("PDFFontsSubSetFonts")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFFontsSubSetFonts = CLng(tStr)
     Else
      If UseStandard Then
       .PDFFontsSubSetFonts = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFFontsSubSetFonts = 1
    End If
  End If
  tStr = reg.GetRegistryValue("PDFFontsSubSetFontsPercent")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .PDFFontsSubSetFontsPercent = CLng(tStr)
     Else
      If UseStandard Then
       .PDFFontsSubSetFontsPercent = 100
      End If
    End If
   Else
    If UseStandard Then
     .PDFFontsSubSetFontsPercent = 100
    End If
  End If
  reg.Subkey = "Printing\Formats\PDF\General"
  tStr = reg.GetRegistryValue("PDFGeneralASCII85")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFGeneralASCII85 = CLng(tStr)
     Else
      If UseStandard Then
       .PDFGeneralASCII85 = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFGeneralASCII85 = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFGeneralAutorotate")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
      .PDFGeneralAutorotate = CLng(tStr)
     Else
      If UseStandard Then
       .PDFGeneralAutorotate = 2
      End If
    End If
   Else
    If UseStandard Then
     .PDFGeneralAutorotate = 2
    End If
  End If
  tStr = reg.GetRegistryValue("PDFGeneralCompatibility")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
      .PDFGeneralCompatibility = CLng(tStr)
     Else
      If UseStandard Then
       .PDFGeneralCompatibility = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFGeneralCompatibility = 1
    End If
  End If
  tStr = reg.GetRegistryValue("PDFGeneralOverprint")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
      .PDFGeneralOverprint = CLng(tStr)
     Else
      If UseStandard Then
       .PDFGeneralOverprint = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFGeneralOverprint = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFGeneralResolution")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .PDFGeneralResolution = CLng(tStr)
     Else
      If UseStandard Then
       .PDFGeneralResolution = 600
      End If
    End If
   Else
    If UseStandard Then
     .PDFGeneralResolution = 600
    End If
  End If
  tStr = reg.GetRegistryValue("PDFOptimize")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFOptimize = CLng(tStr)
     Else
      If UseStandard Then
       .PDFOptimize = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFOptimize = 0
    End If
  End If
  reg.Subkey = "Printing\Formats\PDF\Security"
  tStr = reg.GetRegistryValue("PDFAllowAssembly")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFAllowAssembly = CLng(tStr)
     Else
      If UseStandard Then
       .PDFAllowAssembly = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFAllowAssembly = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFAllowDegradedPrinting")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFAllowDegradedPrinting = CLng(tStr)
     Else
      If UseStandard Then
       .PDFAllowDegradedPrinting = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFAllowDegradedPrinting = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFAllowFillIn")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFAllowFillIn = CLng(tStr)
     Else
      If UseStandard Then
       .PDFAllowFillIn = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFAllowFillIn = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFAllowScreenReaders")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFAllowScreenReaders = CLng(tStr)
     Else
      If UseStandard Then
       .PDFAllowScreenReaders = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFAllowScreenReaders = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFDisallowCopy")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFDisallowCopy = CLng(tStr)
     Else
      If UseStandard Then
       .PDFDisallowCopy = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFDisallowCopy = 1
    End If
  End If
  tStr = reg.GetRegistryValue("PDFDisallowModifyAnnotations")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFDisallowModifyAnnotations = CLng(tStr)
     Else
      If UseStandard Then
       .PDFDisallowModifyAnnotations = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFDisallowModifyAnnotations = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFDisallowModifyContents")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFDisallowModifyContents = CLng(tStr)
     Else
      If UseStandard Then
       .PDFDisallowModifyContents = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFDisallowModifyContents = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFDisallowPrinting")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFDisallowPrinting = CLng(tStr)
     Else
      If UseStandard Then
       .PDFDisallowPrinting = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFDisallowPrinting = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFEncryptor")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
      .PDFEncryptor = CLng(tStr)
     Else
      If UseStandard Then
       .PDFEncryptor = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFEncryptor = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFHighEncryption")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFHighEncryption = CLng(tStr)
     Else
      If UseStandard Then
       .PDFHighEncryption = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFHighEncryption = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFLowEncryption")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFLowEncryption = CLng(tStr)
     Else
      If UseStandard Then
       .PDFLowEncryption = 1
      End If
    End If
   Else
    If UseStandard Then
     .PDFLowEncryption = 1
    End If
  End If
  tStr = reg.GetRegistryValue("PDFOwnerPass")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFOwnerPass = CLng(tStr)
     Else
      If UseStandard Then
       .PDFOwnerPass = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFOwnerPass = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFOwnerPasswordString")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .PDFOwnerPasswordString = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .PDFOwnerPasswordString = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("PDFUserPass")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFUserPass = CLng(tStr)
     Else
      If UseStandard Then
       .PDFUserPass = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFUserPass = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PDFUserPasswordString")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .PDFUserPasswordString = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .PDFUserPasswordString = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("PDFUseSecurity")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PDFUseSecurity = CLng(tStr)
     Else
      If UseStandard Then
       .PDFUseSecurity = 0
      End If
    End If
   Else
    If UseStandard Then
     .PDFUseSecurity = 0
    End If
  End If
  reg.Subkey = "Printing\Formats\PS\LanguageLevel"
  tStr = reg.GetRegistryValue("EPSLanguageLevel")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
      .EPSLanguageLevel = CLng(tStr)
     Else
      If UseStandard Then
       .EPSLanguageLevel = 2
      End If
    End If
   Else
    If UseStandard Then
     .EPSLanguageLevel = 2
    End If
  End If
  tStr = reg.GetRegistryValue("PSLanguageLevel")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
      .PSLanguageLevel = CLng(tStr)
     Else
      If UseStandard Then
       .PSLanguageLevel = 2
      End If
    End If
   Else
    If UseStandard Then
     .PSLanguageLevel = 2
    End If
  End If
  reg.Subkey = "Program"
  tStr = reg.GetRegistryValue("AdditionalGhostscriptParameters")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .AdditionalGhostscriptParameters = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .AdditionalGhostscriptParameters = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("AdditionalGhostscriptSearchpath")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .AdditionalGhostscriptSearchpath = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .AdditionalGhostscriptSearchpath = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("AddWindowsFontpath")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .AddWindowsFontpath = CLng(tStr)
     Else
      If UseStandard Then
       .AddWindowsFontpath = 1
      End If
    End If
   Else
    If UseStandard Then
     .AddWindowsFontpath = 1
    End If
  End If
  tStr = reg.GetRegistryValue("AutosaveDirectory")
  If LenB(Trim$(tStr)) > 0 Then
    .AutosaveDirectory = CompletePath(tStr)
   Else
    If UseStandard Then
     tStr = GetMyFiles
     .AutosaveDirectory = CompletePath(tStr)
    End If
  End If
  tStr = reg.GetRegistryValue("AutosaveFilename")
  If LenB(tStr) = 0 And LenB("<DateTime>")>0 Then
    If UseStandard Then
     .AutosaveFilename = "<DateTime>"
    End If
   Else
    If LenB(tStr) > 0 Then
     .AutosaveFilename = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("AutosaveFormat")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
      .AutosaveFormat = CLng(tStr)
     Else
      If UseStandard Then
       .AutosaveFormat = 0
      End If
    End If
   Else
    If UseStandard Then
     .AutosaveFormat = 0
    End If
  End If
  tStr = reg.GetRegistryValue("ClientComputerResolveIPAddress")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .ClientComputerResolveIPAddress = CLng(tStr)
     Else
      If UseStandard Then
       .ClientComputerResolveIPAddress = 0
      End If
    End If
   Else
    If UseStandard Then
     .ClientComputerResolveIPAddress = 0
    End If
  End If
  tStr = reg.GetRegistryValue("DisableEmail")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .DisableEmail = CLng(tStr)
     Else
      If UseStandard Then
       .DisableEmail = 0
      End If
    End If
   Else
    If UseStandard Then
     .DisableEmail = 0
    End If
  End If
  tStr = reg.GetRegistryValue("DontUseDocumentSettings")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .DontUseDocumentSettings = CLng(tStr)
     Else
      If UseStandard Then
       .DontUseDocumentSettings = 0
      End If
    End If
   Else
    If UseStandard Then
     .DontUseDocumentSettings = 0
    End If
  End If
  tStr = reg.GetRegistryValue("FilenameSubstitutions")
  If LenB(tStr) = 0 And LenB("Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt")>0 Then
    If UseStandard Then
     .FilenameSubstitutions = "Microsoft Word - \.doc\Microsoft Excel - \.xls\Microsoft PowerPoint - \.ppt"
    End If
   Else
    If LenB(tStr) > 0 Then
     .FilenameSubstitutions = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("FilenameSubstitutionsOnlyInTitle")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .FilenameSubstitutionsOnlyInTitle = CLng(tStr)
     Else
      If UseStandard Then
       .FilenameSubstitutionsOnlyInTitle = 1
      End If
    End If
   Else
    If UseStandard Then
     .FilenameSubstitutionsOnlyInTitle = 1
    End If
  End If
  tStr = reg.GetRegistryValue("Language")
  If LenB(tStr) = 0 And LenB("english")>0 Then
    If UseStandard Then
     .Language = "english"
    End If
   Else
    If LenB(tStr) > 0 Then
     .Language = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("LastSaveDirectory")
  If LenB(Trim$(tStr)) > 0 Then
    .LastSaveDirectory = CompletePath(tStr)
   Else
    If UseStandard Then
     tStr = GetMyFiles
     .LastSaveDirectory = CompletePath(tStr)
    End If
  End If
  tStr = reg.GetRegistryValue("Logging")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .Logging = CLng(tStr)
     Else
      If UseStandard Then
       .Logging = 0
      End If
    End If
   Else
    If UseStandard Then
     .Logging = 0
    End If
  End If
  tStr = reg.GetRegistryValue("LogLines")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
      .LogLines = CLng(tStr)
     Else
      If UseStandard Then
       .LogLines = 100
      End If
    End If
   Else
    If UseStandard Then
     .LogLines = 100
    End If
  End If
  tStr = reg.GetRegistryValue("NoConfirmMessageSwitchingDefaultprinter")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .NoConfirmMessageSwitchingDefaultprinter = CLng(tStr)
     Else
      If UseStandard Then
       .NoConfirmMessageSwitchingDefaultprinter = 0
      End If
    End If
   Else
    If UseStandard Then
     .NoConfirmMessageSwitchingDefaultprinter = 0
    End If
  End If
  tStr = reg.GetRegistryValue("NoProcessingAtStartup")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .NoProcessingAtStartup = CLng(tStr)
     Else
      If UseStandard Then
       .NoProcessingAtStartup = 0
      End If
    End If
   Else
    If UseStandard Then
     .NoProcessingAtStartup = 0
    End If
  End If
  tStr = reg.GetRegistryValue("OptionsDesign")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 1 And CLng(tStr) <= 2 Then
      .OptionsDesign = CLng(tStr)
     Else
      If UseStandard Then
       .OptionsDesign = 1
      End If
    End If
   Else
    If UseStandard Then
     .OptionsDesign = 1
    End If
  End If
  tStr = reg.GetRegistryValue("OptionsEnabled")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .OptionsEnabled = CLng(tStr)
     Else
      If UseStandard Then
       .OptionsEnabled = 1
      End If
    End If
   Else
    If UseStandard Then
     .OptionsEnabled = 1
    End If
  End If
  tStr = reg.GetRegistryValue("OptionsVisible")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .OptionsVisible = CLng(tStr)
     Else
      If UseStandard Then
       .OptionsVisible = 1
      End If
    End If
   Else
    If UseStandard Then
     .OptionsVisible = 1
    End If
  End If
  tStr = reg.GetRegistryValue("PrintAfterSaving")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PrintAfterSaving = CLng(tStr)
     Else
      If UseStandard Then
       .PrintAfterSaving = 0
      End If
    End If
   Else
    If UseStandard Then
     .PrintAfterSaving = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PrintAfterSavingDuplex")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PrintAfterSavingDuplex = CLng(tStr)
     Else
      If UseStandard Then
       .PrintAfterSavingDuplex = 0
      End If
    End If
   Else
    If UseStandard Then
     .PrintAfterSavingDuplex = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PrintAfterSavingNoCancel")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PrintAfterSavingNoCancel = CLng(tStr)
     Else
      If UseStandard Then
       .PrintAfterSavingNoCancel = 0
      End If
    End If
   Else
    If UseStandard Then
     .PrintAfterSavingNoCancel = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PrintAfterSavingPrinter")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .PrintAfterSavingPrinter = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .PrintAfterSavingPrinter = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("PrintAfterSavingQueryUser")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
      .PrintAfterSavingQueryUser = CLng(tStr)
     Else
      If UseStandard Then
       .PrintAfterSavingQueryUser = 0
      End If
    End If
   Else
    If UseStandard Then
     .PrintAfterSavingQueryUser = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PrintAfterSavingTumble")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PrintAfterSavingTumble = CLng(tStr)
     Else
      If UseStandard Then
       .PrintAfterSavingTumble = 0
      End If
    End If
   Else
    If UseStandard Then
     .PrintAfterSavingTumble = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PrinterStop")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .PrinterStop = CLng(tStr)
     Else
      If UseStandard Then
       .PrinterStop = 0
      End If
    End If
   Else
    If UseStandard Then
     .PrinterStop = 0
    End If
  End If
  tStr = reg.GetRegistryValue("PrinterTemppath")
  If LenB(Trim$(tstr)) > 0 Then
   If DirExists(GetSubstFilename2(tstr, False)) = True Then
     .PrinterTemppath = tStr
    Else
     MakePath ResolveEnvironment(GetSubstFilename2(tstr, False))
     If DirExists(ResolveEnvironment(GetSubstFilename2(tstr, False))) = False Then
       If UseStandard Then
         .PrinterTemppath = GetTempPath
        Else
         .PrinterTemppath = ""
         If NoMsg = False Then
          MsgBox "PrinterTemppath: '" & tstr & "' = '" & ResolveEnvironment(GetSubstFilename2(tstr, False)) & "'" & _
           vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07
         End If
       End If
      Else
       .PrinterTemppath = tstr
     End If
   End If
  End If
  tStr = reg.GetRegistryValue("ProcessPriority")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 3 Then
      .ProcessPriority = CLng(tStr)
     Else
      If UseStandard Then
       .ProcessPriority = 1
      End If
    End If
   Else
    If UseStandard Then
     .ProcessPriority = 1
    End If
  End If
  tStr = reg.GetRegistryValue("ProgramFont")
  If LenB(tStr) = 0 And LenB("MS Sans Serif")>0 Then
    If UseStandard Then
     .ProgramFont = "MS Sans Serif"
    End If
   Else
    If LenB(tStr) > 0 Then
     .ProgramFont = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("ProgramFontCharset")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .ProgramFontCharset = CLng(tStr)
     Else
      If UseStandard Then
       .ProgramFontCharset = 0
      End If
    End If
   Else
    If UseStandard Then
     .ProgramFontCharset = 0
    End If
  End If
  tStr = reg.GetRegistryValue("ProgramFontSize")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 1 And CLng(tStr) <= 72 Then
      .ProgramFontSize = CLng(tStr)
     Else
      If UseStandard Then
       .ProgramFontSize = 8
      End If
    End If
   Else
    If UseStandard Then
     .ProgramFontSize = 8
    End If
  End If
  tStr = reg.GetRegistryValue("RemoveAllKnownFileExtensions")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .RemoveAllKnownFileExtensions = CLng(tStr)
     Else
      If UseStandard Then
       .RemoveAllKnownFileExtensions = 1
      End If
    End If
   Else
    If UseStandard Then
     .RemoveAllKnownFileExtensions = 1
    End If
  End If
  tStr = reg.GetRegistryValue("RemoveSpaces")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .RemoveSpaces = CLng(tStr)
     Else
      If UseStandard Then
       .RemoveSpaces = 1
      End If
    End If
   Else
    If UseStandard Then
     .RemoveSpaces = 1
    End If
  End If
  tStr = reg.GetRegistryValue("RunProgramAfterSaving")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .RunProgramAfterSaving = CLng(tStr)
     Else
      If UseStandard Then
       .RunProgramAfterSaving = 0
      End If
    End If
   Else
    If UseStandard Then
     .RunProgramAfterSaving = 0
    End If
  End If
  tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramname")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .RunProgramAfterSavingProgramname = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .RunProgramAfterSavingProgramname = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("RunProgramAfterSavingProgramParameters")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .RunProgramAfterSavingProgramParameters = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .RunProgramAfterSavingProgramParameters = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("RunProgramAfterSavingWaitUntilReady")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .RunProgramAfterSavingWaitUntilReady = CLng(tStr)
     Else
      If UseStandard Then
       .RunProgramAfterSavingWaitUntilReady = 1
      End If
    End If
   Else
    If UseStandard Then
     .RunProgramAfterSavingWaitUntilReady = 1
    End If
  End If
  tStr = reg.GetRegistryValue("RunProgramAfterSavingWindowstyle")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
      .RunProgramAfterSavingWindowstyle = CLng(tStr)
     Else
      If UseStandard Then
       .RunProgramAfterSavingWindowstyle = 1
      End If
    End If
   Else
    If UseStandard Then
     .RunProgramAfterSavingWindowstyle = 1
    End If
  End If
  tStr = reg.GetRegistryValue("RunProgramBeforeSaving")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .RunProgramBeforeSaving = CLng(tStr)
     Else
      If UseStandard Then
       .RunProgramBeforeSaving = 0
      End If
    End If
   Else
    If UseStandard Then
     .RunProgramBeforeSaving = 0
    End If
  End If
  tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramname")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .RunProgramBeforeSavingProgramname = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .RunProgramBeforeSavingProgramname = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("RunProgramBeforeSavingProgramParameters")
  If LenB(tStr) = 0 And LenB("")>0 Then
    If UseStandard Then
     .RunProgramBeforeSavingProgramParameters = " "
    End If
   Else
    If LenB(tStr) > 0 Then
     .RunProgramBeforeSavingProgramParameters = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("RunProgramBeforeSavingWindowstyle")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
      .RunProgramBeforeSavingWindowstyle = CLng(tStr)
     Else
      If UseStandard Then
       .RunProgramBeforeSavingWindowstyle = 1
      End If
    End If
   Else
    If UseStandard Then
     .RunProgramBeforeSavingWindowstyle = 1
    End If
  End If
  tStr = reg.GetRegistryValue("SaveFilename")
  If LenB(tStr) = 0 And LenB("<Title>")>0 Then
    If UseStandard Then
     .SaveFilename = "<Title>"
    End If
   Else
    If LenB(tStr) > 0 Then
     .SaveFilename = tStr
    End If
  End If
  tStr = reg.GetRegistryValue("SendMailMethod")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .SendMailMethod = CLng(tStr)
     Else
      If UseStandard Then
       .SendMailMethod = 0
      End If
    End If
   Else
    If UseStandard Then
     .SendMailMethod = 0
    End If
  End If
  tStr = reg.GetRegistryValue("ShowAnimation")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .ShowAnimation = CLng(tStr)
     Else
      If UseStandard Then
       .ShowAnimation = 1
      End If
    End If
   Else
    If UseStandard Then
     .ShowAnimation = 1
    End If
  End If
  tStr = reg.GetRegistryValue("StartStandardProgram")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .StartStandardProgram = CLng(tStr)
     Else
      If UseStandard Then
       .StartStandardProgram = 1
      End If
    End If
   Else
    If UseStandard Then
     .StartStandardProgram = 1
    End If
  End If
  tStr = reg.GetRegistryValue("Toolbars")
  If Isnumeric(tStr) Then
    If CLng(tStr) >= 0 Then
      .Toolbars = CLng(tStr)
     Else
      If UseStandard Then
       .Toolbars = 1
      End If
    End If
   Else
    If UseStandard Then
     .Toolbars = 1
    End If
  End If
  tStr = reg.GetRegistryValue("UseAutosave")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .UseAutosave = CLng(tStr)
     Else
      If UseStandard Then
       .UseAutosave = 0
      End If
    End If
   Else
    If UseStandard Then
     .UseAutosave = 0
    End If
  End If
  tStr = reg.GetRegistryValue("UseAutosaveDirectory")
  If IsNumeric(tStr) Then
    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
      .UseAutosaveDirectory = CLng(tStr)
     Else
      If UseStandard Then
       .UseAutosaveDirectory = 1
      End If
    End If
   Else
    If UseStandard Then
     .UseAutosaveDirectory = 1
    End If
  End If
 End With
 Set reg = Nothing
 ReadOptionsReg = MyOptions
End Function

Public Sub SaveOptionsREG(sOptions as tOptions, Optional hkey1 as hkey = HKEY_CURRENT_USER)
 Dim reg As clsRegistry
 Set reg = New clsRegistry
 reg.hkey = hkey1
 reg.KeyRoot = "Software\PDFCreator"
 If Not reg.KeyExists Then
  reg.CreateKey
 End If
 With sOptions
  reg.Subkey = "Ghostscript"
  If Not reg.KeyExists Then
   reg.CreateKey
  End If
  reg.SetRegistryValue "DirectoryGhostscriptBinaries",CStr(.DirectoryGhostscriptBinaries), REG_SZ
  reg.SetRegistryValue "DirectoryGhostscriptFonts",CStr(.DirectoryGhostscriptFonts), REG_SZ
  reg.SetRegistryValue "DirectoryGhostscriptLibraries",CStr(.DirectoryGhostscriptLibraries), REG_SZ
  reg.SetRegistryValue "DirectoryGhostscriptResource",CStr(.DirectoryGhostscriptResource), REG_SZ
  reg.Subkey = "Printing"
  If Not reg.KeyExists Then
   reg.CreateKey
  End If
  reg.SetRegistryValue "DeviceHeightPoints",CStr(.DeviceHeightPoints), REG_SZ
  reg.SetRegistryValue "DeviceWidthPoints",CStr(.DeviceWidthPoints), REG_SZ
  reg.SetRegistryValue "OnePagePerFile",CStr(Abs(.OnePagePerFile)), REG_SZ
  reg.SetRegistryValue "Papersize",CStr(.Papersize), REG_SZ
  reg.SetRegistryValue "StampFontColor",CStr(.StampFontColor), REG_SZ
  reg.SetRegistryValue "StampFontname",CStr(.StampFontname), REG_SZ
  reg.SetRegistryValue "StampFontsize",CStr(.StampFontsize), REG_SZ
  reg.SetRegistryValue "StampOutlineFontthickness",CStr(.StampOutlineFontthickness), REG_SZ
  reg.SetRegistryValue "StampString",CStr(.StampString), REG_SZ
  reg.SetRegistryValue "StampUseOutlineFont",CStr(Abs(.StampUseOutlineFont)), REG_SZ
  reg.SetRegistryValue "StandardAuthor",CStr(.StandardAuthor), REG_SZ
  reg.SetRegistryValue "StandardCreationdate",CStr(.StandardCreationdate), REG_SZ
  reg.SetRegistryValue "StandardDateformat",CStr(.StandardDateformat), REG_SZ
  reg.SetRegistryValue "StandardKeywords",CStr(.StandardKeywords), REG_SZ
  reg.SetRegistryValue "StandardMailDomain",CStr(.StandardMailDomain), REG_SZ
  reg.SetRegistryValue "StandardModifydate",CStr(.StandardModifydate), REG_SZ
  reg.SetRegistryValue "StandardSaveformat",CStr(.StandardSaveformat), REG_SZ
  reg.SetRegistryValue "StandardSubject",CStr(.StandardSubject), REG_SZ
  reg.SetRegistryValue "StandardTitle",CStr(.StandardTitle), REG_SZ
  reg.SetRegistryValue "UseCreationDateNow",CStr(Abs(.UseCreationDateNow)), REG_SZ
  reg.SetRegistryValue "UseStandardAuthor",CStr(Abs(.UseStandardAuthor)), REG_SZ
  reg.Subkey = "Printing\Formats\Bitmap\Colors"
  If Not reg.KeyExists Then
   reg.CreateKey
  End If
  reg.SetRegistryValue "BitmapResolution",CStr(.BitmapResolution), REG_SZ
  reg.SetRegistryValue "BMPColorscount",CStr(.BMPColorscount), REG_SZ
  reg.SetRegistryValue "JPEGColorscount",CStr(.JPEGColorscount), REG_SZ
  reg.SetRegistryValue "JPEGQuality",CStr(.JPEGQuality), REG_SZ
  reg.SetRegistryValue "PCXColorscount",CStr(.PCXColorscount), REG_SZ
  reg.SetRegistryValue "PNGColorscount",CStr(.PNGColorscount), REG_SZ
  reg.SetRegistryValue "TIFFColorscount",CStr(.TIFFColorscount), REG_SZ
  reg.Subkey = "Printing\Formats\PDF\Colors"
  If Not reg.KeyExists Then
   reg.CreateKey
  End If
  reg.SetRegistryValue "PDFColorsCMYKToRGB",CStr(Abs(.PDFColorsCMYKToRGB)), REG_SZ
  reg.SetRegistryValue "PDFColorsColorModel",CStr(.PDFColorsColorModel), REG_SZ
  reg.SetRegistryValue "PDFColorsPreserveHalftone",CStr(Abs(.PDFColorsPreserveHalftone)), REG_SZ
  reg.SetRegistryValue "PDFColorsPreserveOverprint",CStr(Abs(.PDFColorsPreserveOverprint)), REG_SZ
  reg.SetRegistryValue "PDFColorsPreserveTransfer",CStr(Abs(.PDFColorsPreserveTransfer)), REG_SZ
  reg.Subkey = "Printing\Formats\PDF\Compression"
  If Not reg.KeyExists Then
   reg.CreateKey
  End If
  reg.SetRegistryValue "PDFCompressionColorCompression",CStr(Abs(.PDFCompressionColorCompression)), REG_SZ
  reg.SetRegistryValue "PDFCompressionColorCompressionChoice",CStr(.PDFCompressionColorCompressionChoice), REG_SZ
  reg.SetRegistryValue "PDFCompressionColorResample",CStr(Abs(.PDFCompressionColorResample)), REG_SZ
  reg.SetRegistryValue "PDFCompressionColorResampleChoice",CStr(.PDFCompressionColorResampleChoice), REG_SZ
  reg.SetRegistryValue "PDFCompressionColorResolution",CStr(.PDFCompressionColorResolution), REG_SZ
  reg.SetRegistryValue "PDFCompressionGreyCompression",CStr(Abs(.PDFCompressionGreyCompression)), REG_SZ
  reg.SetRegistryValue "PDFCompressionGreyCompressionChoice",CStr(.PDFCompressionGreyCompressionChoice), REG_SZ
  reg.SetRegistryValue "PDFCompressionGreyResample",CStr(Abs(.PDFCompressionGreyResample)), REG_SZ
  reg.SetRegistryValue "PDFCompressionGreyResampleChoice",CStr(.PDFCompressionGreyResampleChoice), REG_SZ
  reg.SetRegistryValue "PDFCompressionGreyResolution",CStr(.PDFCompressionGreyResolution), REG_SZ
  reg.SetRegistryValue "PDFCompressionMonoCompression",CStr(Abs(.PDFCompressionMonoCompression)), REG_SZ
  reg.SetRegistryValue "PDFCompressionMonoCompressionChoice",CStr(.PDFCompressionMonoCompressionChoice), REG_SZ
  reg.SetRegistryValue "PDFCompressionMonoResample",CStr(Abs(.PDFCompressionMonoResample)), REG_SZ
  reg.SetRegistryValue "PDFCompressionMonoResampleChoice",CStr(.PDFCompressionMonoResampleChoice), REG_SZ
  reg.SetRegistryValue "PDFCompressionMonoResolution",CStr(.PDFCompressionMonoResolution), REG_SZ
  reg.SetRegistryValue "PDFCompressionTextCompression",CStr(Abs(.PDFCompressionTextCompression)), REG_SZ
  reg.Subkey = "Printing\Formats\PDF\Fonts"
  If Not reg.KeyExists Then
   reg.CreateKey
  End If
  reg.SetRegistryValue "PDFFontsEmbedAll",CStr(Abs(.PDFFontsEmbedAll)), REG_SZ
  reg.SetRegistryValue "PDFFontsSubSetFonts",CStr(Abs(.PDFFontsSubSetFonts)), REG_SZ
  reg.SetRegistryValue "PDFFontsSubSetFontsPercent",CStr(.PDFFontsSubSetFontsPercent), REG_SZ
  reg.Subkey = "Printing\Formats\PDF\General"
  If Not reg.KeyExists Then
   reg.CreateKey
  End If
  reg.SetRegistryValue "PDFGeneralASCII85",CStr(Abs(.PDFGeneralASCII85)), REG_SZ
  reg.SetRegistryValue "PDFGeneralAutorotate",CStr(.PDFGeneralAutorotate), REG_SZ
  reg.SetRegistryValue "PDFGeneralCompatibility",CStr(.PDFGeneralCompatibility), REG_SZ
  reg.SetRegistryValue "PDFGeneralOverprint",CStr(.PDFGeneralOverprint), REG_SZ
  reg.SetRegistryValue "PDFGeneralResolution",CStr(.PDFGeneralResolution), REG_SZ
  reg.SetRegistryValue "PDFOptimize",CStr(Abs(.PDFOptimize)), REG_SZ
  reg.Subkey = "Printing\Formats\PDF\Security"
  If Not reg.KeyExists Then
   reg.CreateKey
  End If
  reg.SetRegistryValue "PDFAllowAssembly",CStr(Abs(.PDFAllowAssembly)), REG_SZ
  reg.SetRegistryValue "PDFAllowDegradedPrinting",CStr(Abs(.PDFAllowDegradedPrinting)), REG_SZ
  reg.SetRegistryValue "PDFAllowFillIn",CStr(Abs(.PDFAllowFillIn)), REG_SZ
  reg.SetRegistryValue "PDFAllowScreenReaders",CStr(Abs(.PDFAllowScreenReaders)), REG_SZ
  reg.SetRegistryValue "PDFDisallowCopy",CStr(Abs(.PDFDisallowCopy)), REG_SZ
  reg.SetRegistryValue "PDFDisallowModifyAnnotations",CStr(Abs(.PDFDisallowModifyAnnotations)), REG_SZ
  reg.SetRegistryValue "PDFDisallowModifyContents",CStr(Abs(.PDFDisallowModifyContents)), REG_SZ
  reg.SetRegistryValue "PDFDisallowPrinting",CStr(Abs(.PDFDisallowPrinting)), REG_SZ
  reg.SetRegistryValue "PDFEncryptor",CStr(.PDFEncryptor), REG_SZ
  reg.SetRegistryValue "PDFHighEncryption",CStr(Abs(.PDFHighEncryption)), REG_SZ
  reg.SetRegistryValue "PDFLowEncryption",CStr(Abs(.PDFLowEncryption)), REG_SZ
  reg.SetRegistryValue "PDFOwnerPass",CStr(Abs(.PDFOwnerPass)), REG_SZ
  reg.SetRegistryValue "PDFOwnerPasswordString",CStr(.PDFOwnerPasswordString), REG_SZ
  reg.SetRegistryValue "PDFUserPass",CStr(Abs(.PDFUserPass)), REG_SZ
  reg.SetRegistryValue "PDFUserPasswordString",CStr(.PDFUserPasswordString), REG_SZ
  reg.SetRegistryValue "PDFUseSecurity",CStr(Abs(.PDFUseSecurity)), REG_SZ
  reg.Subkey = "Printing\Formats\PS\LanguageLevel"
  If Not reg.KeyExists Then
   reg.CreateKey
  End If
  reg.SetRegistryValue "EPSLanguageLevel",CStr(.EPSLanguageLevel), REG_SZ
  reg.SetRegistryValue "PSLanguageLevel",CStr(.PSLanguageLevel), REG_SZ
  reg.Subkey = "Program"
  If Not reg.KeyExists Then
   reg.CreateKey
  End If
  reg.SetRegistryValue "AdditionalGhostscriptParameters",CStr(.AdditionalGhostscriptParameters), REG_SZ
  reg.SetRegistryValue "AdditionalGhostscriptSearchpath",CStr(.AdditionalGhostscriptSearchpath), REG_SZ
  reg.SetRegistryValue "AddWindowsFontpath",CStr(Abs(.AddWindowsFontpath)), REG_SZ
  reg.SetRegistryValue "AutosaveDirectory",CStr(.AutosaveDirectory), REG_SZ
  reg.SetRegistryValue "AutosaveFilename",CStr(.AutosaveFilename), REG_SZ
  reg.SetRegistryValue "AutosaveFormat",CStr(.AutosaveFormat), REG_SZ
  reg.SetRegistryValue "ClientComputerResolveIPAddress",CStr(Abs(.ClientComputerResolveIPAddress)), REG_SZ
  reg.SetRegistryValue "DisableEmail",CStr(Abs(.DisableEmail)), REG_SZ
  reg.SetRegistryValue "DontUseDocumentSettings",CStr(Abs(.DontUseDocumentSettings)), REG_SZ
  reg.SetRegistryValue "FilenameSubstitutions",CStr(.FilenameSubstitutions), REG_SZ
  reg.SetRegistryValue "FilenameSubstitutionsOnlyInTitle",CStr(Abs(.FilenameSubstitutionsOnlyInTitle)), REG_SZ
  reg.SetRegistryValue "Language",CStr(.Language), REG_SZ
  reg.SetRegistryValue "LastSaveDirectory",CStr(.LastSaveDirectory), REG_SZ
  reg.SetRegistryValue "Logging",CStr(Abs(.Logging)), REG_SZ
  reg.SetRegistryValue "LogLines",CStr(.LogLines), REG_SZ
  reg.SetRegistryValue "NoConfirmMessageSwitchingDefaultprinter",CStr(Abs(.NoConfirmMessageSwitchingDefaultprinter)), REG_SZ
  reg.SetRegistryValue "NoProcessingAtStartup",CStr(Abs(.NoProcessingAtStartup)), REG_SZ
  reg.SetRegistryValue "OptionsDesign",CStr(.OptionsDesign), REG_SZ
  reg.SetRegistryValue "OptionsEnabled",CStr(Abs(.OptionsEnabled)), REG_SZ
  reg.SetRegistryValue "OptionsVisible",CStr(Abs(.OptionsVisible)), REG_SZ
  reg.SetRegistryValue "PrintAfterSaving",CStr(Abs(.PrintAfterSaving)), REG_SZ
  reg.SetRegistryValue "PrintAfterSavingDuplex",CStr(Abs(.PrintAfterSavingDuplex)), REG_SZ
  reg.SetRegistryValue "PrintAfterSavingNoCancel",CStr(Abs(.PrintAfterSavingNoCancel)), REG_SZ
  reg.SetRegistryValue "PrintAfterSavingPrinter",CStr(.PrintAfterSavingPrinter), REG_SZ
  reg.SetRegistryValue "PrintAfterSavingQueryUser",CStr(.PrintAfterSavingQueryUser), REG_SZ
  reg.SetRegistryValue "PrintAfterSavingTumble",CStr(Abs(.PrintAfterSavingTumble)), REG_SZ
  reg.SetRegistryValue "PrinterStop",CStr(Abs(.PrinterStop)), REG_SZ
  reg.SetRegistryValue "PrinterTemppath",CStr(.PrinterTemppath), REG_SZ
  reg.SetRegistryValue "ProcessPriority",CStr(.ProcessPriority), REG_SZ
  reg.SetRegistryValue "ProgramFont",CStr(.ProgramFont), REG_SZ
  reg.SetRegistryValue "ProgramFontCharset",CStr(.ProgramFontCharset), REG_SZ
  reg.SetRegistryValue "ProgramFontSize",CStr(.ProgramFontSize), REG_SZ
  reg.SetRegistryValue "RemoveAllKnownFileExtensions",CStr(Abs(.RemoveAllKnownFileExtensions)), REG_SZ
  reg.SetRegistryValue "RemoveSpaces",CStr(Abs(.RemoveSpaces)), REG_SZ
  reg.SetRegistryValue "RunProgramAfterSaving",CStr(Abs(.RunProgramAfterSaving)), REG_SZ
  reg.SetRegistryValue "RunProgramAfterSavingProgramname",CStr(.RunProgramAfterSavingProgramname), REG_SZ
  reg.SetRegistryValue "RunProgramAfterSavingProgramParameters",CStr(.RunProgramAfterSavingProgramParameters), REG_SZ
  reg.SetRegistryValue "RunProgramAfterSavingWaitUntilReady",CStr(Abs(.RunProgramAfterSavingWaitUntilReady)), REG_SZ
  reg.SetRegistryValue "RunProgramAfterSavingWindowstyle",CStr(.RunProgramAfterSavingWindowstyle), REG_SZ
  reg.SetRegistryValue "RunProgramBeforeSaving",CStr(Abs(.RunProgramBeforeSaving)), REG_SZ
  reg.SetRegistryValue "RunProgramBeforeSavingProgramname",CStr(.RunProgramBeforeSavingProgramname), REG_SZ
  reg.SetRegistryValue "RunProgramBeforeSavingProgramParameters",CStr(.RunProgramBeforeSavingProgramParameters), REG_SZ
  reg.SetRegistryValue "RunProgramBeforeSavingWindowstyle",CStr(.RunProgramBeforeSavingWindowstyle), REG_SZ
  reg.SetRegistryValue "SaveFilename",CStr(.SaveFilename), REG_SZ
  reg.SetRegistryValue "SendMailMethod",CStr(.SendMailMethod), REG_SZ
  reg.SetRegistryValue "ShowAnimation",CStr(Abs(.ShowAnimation)), REG_SZ
  reg.SetRegistryValue "StartStandardProgram",CStr(Abs(.StartStandardProgram)), REG_SZ
  reg.SetRegistryValue "Toolbars",CStr(.Toolbars), REG_SZ
  reg.SetRegistryValue "UseAutosave",CStr(Abs(.UseAutosave)), REG_SZ
  reg.SetRegistryValue "UseAutosaveDirectory",CStr(Abs(.UseAutosaveDirectory)), REG_SZ
 End With
 Set reg = Nothing
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
  frm.txtGSResource.text = .DirectoryGhostscriptResource
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
  frm.cmbProgramFontSize.text = .ProgramFontSize
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
 .DirectoryGhostscriptResource =  frm.txtGSResource.text
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
 .ProgramFontSize =  frm.cmbProgramFontSize.text
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
   PrintSelectedJobs = False
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

Public Sub ReadLanguageFromOptions(Optional hProfile As hkey = HKEY_CURRENT_USER)
 Dim sLanguage As String
 If InstalledAsServer Then
   If UseINI Then
     sLanguage = ReadLanguageFromOptionsINI(sLanguage, Completepath(GetCommonAppData) & "PDFCreator.ini")
    Else
     sLanguage = ReadLanguageFromOptionsReg(sLanguage, "Software\PDFCreator", HKEY_LOCAL_MACHINE)
   End If
  Else
   If UseINI Then
     If Not IsWin9xMe Then
       sLanguage = ReadLanguageFromOptionsINI(sLanguage, Completepath(GetDefaultAppData) & "PDFCreator.ini")
       sLanguage = ReadLanguageFromOptionsINI(sLanguage, PDFCreatorINIFile, False)
      Else
       sLanguage = ReadLanguageFromOptionsINI(sLanguage, PDFCreatorINIFile)
     End If
     sLanguage = ReadLanguageFromOptionsINI(sLanguage, Completepath(GetCommonAppData) & "PDFCreator.ini", False)
    Else
     If Not IsWin9xMe Then
       sLanguage = ReadLanguageFromOptionsReg(sLanguage, ".DEFAULT\Software\PDFCreator", HKEY_USERS)
       sLanguage = ReadLanguageFromOptionsReg(sLanguage, "Software\PDFCreator", hProfile, False)
      Else
       sLanguage = ReadLanguageFromOptionsReg(sLanguage, "Software\PDFCreator", hProfile)
     End If
     sLanguage = ReadLanguageFromOptionsReg(sLanguage, "Software\PDFCreator", HKEY_LOCAL_MACHINE, False)
   End If
 End If
 Options.Language = sLanguage
End Sub

Public Function ReadLanguageFromOptionsINI(Language As String, PDFCreatorINIFile As String, Optional UseStandard as Boolean = True) As String
 Dim hOpt As clsHash, tStr
 Set hOpt = New clsHash
 ReadINISection PDFCreatorINIFile, "Options", hOpt
 tStr = Trim$(hOpt.Retrieve("Language"))
 If LenB(tStr) > 0 Then
   ReadLanguageFromOptionsINI = tStr
  Else
   If UseStandard Then
     ReadLanguageFromOptionsINI = "english"
    Else
     ReadLanguageFromOptionsINI = Language
   End If
 End If
 Set hOpt = Nothing
End Function

Public Function ReadLanguageFromOptionsReg(Language As String, KeyRoot as String, Optional hProfile as hkey = HKEY_CURRENT_USER, Optional UseStandard as Boolean = True) As String
 Dim reg As clsRegistry, tStr as String
 Set reg = New clsRegistry
 With reg
  .KeyRoot = KeyRoot
  .Subkey = "Program"
  .hkey = hProfile
  tStr = Trim$(reg.GetRegistryValue("Language"))
 End With
 If LenB(tStr) > 0 Then
   ReadLanguageFromOptionsReg = tStr
  Else
   If UseStandard Then
     ReadLanguageFromOptionsReg = "english"
    Else
     ReadLanguageFromOptionsReg = Language
   End If
 End If
 Set reg = Nothing
End Function

Public Function UseINI() As Boolean
 Dim reg As clsRegistry, tStr as String
 Set reg = New clsRegistry
 UseINI = False
 With reg
  .hkey = HKEY_LOCAL_MACHINE
  .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
  tStr = Trim$(.GetRegistryValue("UseINI"))
  If tStr = "1" Then
   UseINI = True
  End If
 End With
 Set reg = Nothing
End Function

