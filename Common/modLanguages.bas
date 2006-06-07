Attribute VB_Name = "modLanguage"
Option Explicit

' Module automatically generated with LanguagesTool from Frank Heindörfer
' 2004
' Email: thesmilyface@users.sourceforge.net

Public Type tLanguageStrings
 CommonAuthor As String
 CommonLanguagename As String
 CommonTitle As String
 CommonVersion As String

 DialogDocument As String
 DialogDocumentAdd As String
 DialogDocumentBottom As String
 DialogDocumentCombine As String
 DialogDocumentCombineAll As String
 DialogDocumentCombineAllSend As String
 DialogDocumentDelete As String
 DialogDocumentDown As String
 DialogDocumentPrint As String
 DialogDocumentSave As String
 DialogDocumentSend As String
 DialogDocumentTop As String
 DialogDocumentUp As String
 DialogEmailAddress As String
 DialogInfo As String
 DialogInfoCheckUpdates As String
 DialogInfoHomepage As String
 DialogInfoInfo As String
 DialogInfoPaypal As String
 DialogInfoPDFCreatorSourceforge As String
 DialogLanguage As String
 DialogPrinter As String
 DialogPrinterClose As String
 DialogPrinterLogfile As String
 DialogPrinterLogfiles As String
 DialogPrinterLogging As String
 DialogPrinterOptions As String
 DialogPrinterPrinterStop As String
 DialogView As String
 DialogViewStatusbar As String
 DialogViewToolbars As String
 DialogViewToolbarsEmail As String
 DialogViewToolbarsStandard As String

 ListAddFile As String
 ListAllFiles As String
 ListBytes As String
 ListDate As String
 ListDocumenttitle As String
 ListFilename As String
 ListGBytes As String
 ListKBytes As String
 ListMBytes As String
 ListPDFFiles As String
 ListPostscriptFiles As String
 ListPrinting As String
 ListSize As String
 ListStatus As String
 ListWaiting As String

 LoggingClear As String
 LoggingClose As String
 LoggingLogfile As String

 MessagesMsg01 As String
 MessagesMsg02 As String
 MessagesMsg03 As String
 MessagesMsg04 As String
 MessagesMsg05 As String
 MessagesMsg06 As String
 MessagesMsg07 As String
 MessagesMsg08 As String
 MessagesMsg09 As String
 MessagesMsg10 As String
 MessagesMsg11 As String
 MessagesMsg12 As String
 MessagesMsg13 As String
 MessagesMsg14 As String
 MessagesMsg15 As String
 MessagesMsg16 As String
 MessagesMsg17 As String
 MessagesMsg19 As String
 MessagesMsg20 As String
 MessagesMsg21 As String
 MessagesMsg22 As String
 MessagesMsg23 As String
 MessagesMsg24 As String
 MessagesMsg25 As String
 MessagesMsg26 As String
 MessagesMsg27 As String
 MessagesMsg28 As String
 MessagesMsg29 As String
 MessagesMsg30 As String
 MessagesMsg31 As String
 MessagesMsg32 As String
 MessagesMsg33 As String
 MessagesMsg34 As String
 MessagesMsg35 As String
 MessagesMsg36 As String

 OptionsAssociatePSFiles As String
 OptionsAutosaveDirectoryPrompt As String
 OptionsAutosaveFilename As String
 OptionsAutosaveFilenameTokens As String
 OptionsAutosaveFormat As String
 OptionsAutosaveStartStandardProgram As String
 OptionsBitmapResolution As String
 OptionsBMPColorscount01 As String
 OptionsBMPColorscount02 As String
 OptionsBMPColorscount03 As String
 OptionsBMPColorscount04 As String
 OptionsBMPColorscount05 As String
 OptionsBMPColorscount06 As String
 OptionsBMPColorscount07 As String
 OptionsBMPDescription As String
 OptionsBMPSymbol As String
 OptionsCancel As String
 OptionsDirectoriesGSBin As String
 OptionsDirectoriesGSFonts As String
 OptionsDirectoriesGSLibraries As String
 OptionsDirectoriesTempPath As String
 OptionsDocument As String
 OptionsEPSDescription As String
 OptionsEPSFiles As String
 OptionsEPSSymbol As String
 OptionsGhostscriptBinariesDirectoryPrompt As String
 OptionsGhostscriptFontsDirectoryPrompt As String
 OptionsGhostscriptInternal As String
 OptionsGhostscriptLibrariesDirectoryPrompt As String
 OptionsGhostscriptResourceDirectoryPrompt As String
 OptionsGhostscriptversion As String
 OptionsImageSettings As String
 OptionsJavaPath As String
 OptionsJPEGColorscount01 As String
 OptionsJPEGColorscount02 As String
 OptionsJPEGDescription As String
 OptionsJPEGQuality As String
 OptionsJPEGSymbol As String
 OptionsOwnerPass As String
 OptionsPassCancel As String
 OptionsPassOK As String
 OptionsPCXColorscount01 As String
 OptionsPCXColorscount02 As String
 OptionsPCXColorscount03 As String
 OptionsPCXColorscount04 As String
 OptionsPCXColorscount05 As String
 OptionsPCXColorscount06 As String
 OptionsPCXDescription As String
 OptionsPCXSymbol As String
 OptionsPDFAllowAssembly As String
 OptionsPDFAllowDegradedPrinting As String
 OptionsPDFAllowFillIn As String
 OptionsPDFAllowScreenReaders As String
 OptionsPDFColors As String
 OptionsPDFColorsCaption As String
 OptionsPDFColorsCMYKtoRGB As String
 OptionsPDFColorsColorModel01 As String
 OptionsPDFColorsColorModel02 As String
 OptionsPDFColorsColorModel03 As String
 OptionsPDFColorsColorOptions As String
 OptionsPDFColorsPreserveHalftone As String
 OptionsPDFColorsPreserveOverprint As String
 OptionsPDFColorsPreserveTransfer As String
 OptionsPDFCompression As String
 OptionsPDFCompressionCaption As String
 OptionsPDFCompressionColor As String
 OptionsPDFCompressionColorComp As String
 OptionsPDFCompressionColorComp01 As String
 OptionsPDFCompressionColorComp02 As String
 OptionsPDFCompressionColorComp03 As String
 OptionsPDFCompressionColorComp04 As String
 OptionsPDFCompressionColorComp05 As String
 OptionsPDFCompressionColorComp06 As String
 OptionsPDFCompressionColorComp07 As String
 OptionsPDFCompressionColorComp08 As String
 OptionsPDFCompressionColorRes As String
 OptionsPDFCompressionColorResample As String
 OptionsPDFCompressionColorResample01 As String
 OptionsPDFCompressionColorResample02 As String
 OptionsPDFCompressionColorResample03 As String
 OptionsPDFCompressionGrey As String
 OptionsPDFCompressionGreyComp As String
 OptionsPDFCompressionGreyComp01 As String
 OptionsPDFCompressionGreyComp02 As String
 OptionsPDFCompressionGreyComp03 As String
 OptionsPDFCompressionGreyComp04 As String
 OptionsPDFCompressionGreyComp05 As String
 OptionsPDFCompressionGreyComp06 As String
 OptionsPDFCompressionGreyComp07 As String
 OptionsPDFCompressionGreyComp08 As String
 OptionsPDFCompressionGreyRes As String
 OptionsPDFCompressionGreyResample As String
 OptionsPDFCompressionGreyResample01 As String
 OptionsPDFCompressionGreyResample02 As String
 OptionsPDFCompressionGreyResample03 As String
 OptionsPDFCompressionMono As String
 OptionsPDFCompressionMonoComp As String
 OptionsPDFCompressionMonoComp01 As String
 OptionsPDFCompressionMonoComp02 As String
 OptionsPDFCompressionMonoComp03 As String
 OptionsPDFCompressionMonoComp04 As String
 OptionsPDFCompressionMonoRes As String
 OptionsPDFCompressionMonoResample As String
 OptionsPDFCompressionMonoResample01 As String
 OptionsPDFCompressionMonoResample02 As String
 OptionsPDFCompressionMonoResample03 As String
 OptionsPDFCompressionTextComp As String
 OptionsPDFDescription As String
 OptionsPDFDisallowCopy As String
 OptionsPDFDisallowModify As String
 OptionsPDFDisallowModifyComments As String
 OptionsPDFDisallowPrint As String
 OptionsPDFDisallowUser As String
 OptionsPDFEncryptionHigh As String
 OptionsPDFEncryptionLevel As String
 OptionsPDFEncryptionLow As String
 OptionsPDFEncryptor As String
 OptionsPDFEnhancedPermissions As String
 OptionsPDFEnterPasswords As String
 OptionsPDFFonts As String
 OptionsPDFFontsCaption As String
 OptionsPDFFontsEmbedAll As String
 OptionsPDFFontsSubSetFonts As String
 OptionsPDFGeneral As String
 OptionsPDFGeneralASCII85 As String
 OptionsPDFGeneralAutorotate As String
 OptionsPDFGeneralCaption As String
 OptionsPDFGeneralCompatibility As String
 OptionsPDFGeneralCompatibility01 As String
 OptionsPDFGeneralCompatibility02 As String
 OptionsPDFGeneralCompatibility03 As String
 OptionsPDFGeneralOverprint As String
 OptionsPDFGeneralOverprint01 As String
 OptionsPDFGeneralOverprint02 As String
 OptionsPDFGeneralResolution As String
 OptionsPDFGeneralRotate01 As String
 OptionsPDFGeneralRotate02 As String
 OptionsPDFGeneralRotate03 As String
 OptionsPDFOptions As String
 OptionsPDFOwnerPass As String
 OptionsPDFPasswords As String
 OptionsPDFRepeatPassword As String
 OptionsPDFSecurity As String
 OptionsPDFSecurityCaption As String
 OptionsPDFSetPassword As String
 OptionsPDFSymbol As String
 OptionsPDFUserPass As String
 OptionsPDFUseSecurity As String
 OptionsPNGColorscount01 As String
 OptionsPNGColorscount02 As String
 OptionsPNGColorscount03 As String
 OptionsPNGColorscount04 As String
 OptionsPNGColorscount05 As String
 OptionsPNGDescription As String
 OptionsPNGFiles As String
 OptionsPNGSymbol As String
 OptionsPrintertempDirectoryPrompt As String
 OptionsPrintTestpage As String
 OptionsProcesspriority As String
 OptionsProcesspriorityHigh As String
 OptionsProcesspriorityIdle As String
 OptionsProcesspriorityNormal As String
 OptionsProcesspriorityRealtime As String
 OptionsProgramAutosaveDescription As String
 OptionsProgramAutosaveSymbol As String
 OptionsProgramDirectoriesDescription As String
 OptionsProgramDirectoriesSymbol As String
 OptionsProgramDocumentDescription As String
 OptionsProgramDocumentSymbol As String
 OptionsProgramFont As String
 OptionsProgramFontCancelTest As String
 OptionsProgramFontcharset As String
 OptionsProgramFontDescription As String
 OptionsProgramFontSize As String
 OptionsProgramFontSymbol As String
 OptionsProgramFontTest As String
 OptionsProgramFontTestdescription As String
 OptionsProgramGeneralDescription As String
 OptionsProgramGeneralSymbol As String
 OptionsProgramGhostscriptDescription As String
 OptionsProgramGhostscriptSymbol As String
 OptionsProgramSaveDescription As String
 OptionsProgramSaveSymbol As String
 OptionsProgramSwitchingDefaultprinter As String
 OptionsPSDescription As String
 OptionsPSFiles As String
 OptionsPSLanguageLevel As String
 OptionsPSSymbol As String
 OptionsRemoveSpaces As String
 OptionsReset As String
 OptionsSave As String
 OptionsSaveFilename As String
 OptionsSaveFilenameAdd As String
 OptionsSaveFilenameChange As String
 OptionsSaveFilenameDelete As String
 OptionsSaveFilenameSubstitutions As String
 OptionsSaveFilenameSubstitutionsTitle As String
 OptionsSaveFilenameTokens As String
 OptionsSavePasswords As String
 OptionsShellIntegration As String
 OptionsShellIntegrationAdd As String
 OptionsShellIntegrationCaption As String
 OptionsShellIntegrationRemove As String
 OptionsStandardAuthorToken As String
 OptionsTestpage As String
 OptionsTIFFColorscount01 As String
 OptionsTIFFColorscount02 As String
 OptionsTIFFColorscount03 As String
 OptionsTIFFColorscount04 As String
 OptionsTIFFColorscount05 As String
 OptionsTIFFColorscount06 As String
 OptionsTIFFColorscount07 As String
 OptionsTIFFColorscount08 As String
 OptionsTIFFDescription As String
 OptionsTIFFSymbol As String
 OptionsTreeFormats As String
 OptionsTreeProgram As String
 OptionsUseAutosave As String
 OptionsUseAutosaveDirectory As String
 OptionsUseCreationDateNow As String
 OptionsUserPass As String
 OptionsUseStandardauthor As String

 PrintingAuthor As String
 PrintingBMPFiles As String
 PrintingCancel As String
 PrintingCollect As String
 PrintingCreationDate As String
 PrintingDocumentTitle As String
 PrintingEMail As String
 PrintingEPSFiles As String
 PrintingJPEGFiles As String
 PrintingKeywords As String
 PrintingModifyDate As String
 PrintingNow As String
 PrintingPCXFiles As String
 PrintingPDFFiles As String
 PrintingPNGFiles As String
 PrintingPSFiles As String
 PrintingSave As String
 PrintingStartStandardProgram As String
 PrintingStatus As String
 PrintingSubject As String
 PrintingTIFFFiles As String

 SaveOpenAttributes As String
 SaveOpenCancel As String
 SaveOpenFilename As String
 SaveOpenOpen As String
 SaveOpenOpenTitle As String
 SaveOpenSave As String
 SaveOpenSaveTitle As String
 SaveOpenSize As String

End Type

Public LanguageStrings As tLanguageStrings

Public Sub LoadLanguage(ByVal Languagefile As String)
 InitLanguagesStrings
 LoadCommonStrings Languagefile
 LoadDialogStrings Languagefile
 LoadListStrings Languagefile
 LoadLoggingStrings Languagefile
 LoadMessagesStrings Languagefile
 LoadOptionsStrings Languagefile
 LoadPrintingStrings Languagefile
 LoadSaveOpenStrings Languagefile
End Sub

Private Sub LoadCommonStrings(ByVal Languagefile As String)
 Dim hLang As New clsHash
 ReadINISection Languagefile, "Common", hLang
 With LanguageStrings
  .CommonAuthor = Replace$(hLang.Retrieve("Author", .CommonAuthor),"/n",vbCrLf)
  .CommonLanguagename = Replace$(hLang.Retrieve("Languagename", .CommonLanguagename),"/n",vbCrLf)
  .CommonTitle = Replace$(hLang.Retrieve("Title", .CommonTitle),"/n",vbCrLf)
  .CommonVersion = Replace$(hLang.Retrieve("Version", .CommonVersion),"/n",vbCrLf)
 End With
 Set hLang = Nothing
End Sub

Private Sub LoadDialogStrings(ByVal Languagefile As String)
 Dim hLang As New clsHash
 ReadINISection Languagefile, "Dialog", hLang
 With LanguageStrings
  .DialogDocument = Replace$(hLang.Retrieve("Document", .DialogDocument),"/n",vbCrLf)
  .DialogDocumentAdd = Replace$(hLang.Retrieve("DocumentAdd", .DialogDocumentAdd),"/n",vbCrLf)
  .DialogDocumentBottom = Replace$(hLang.Retrieve("DocumentBottom", .DialogDocumentBottom),"/n",vbCrLf)
  .DialogDocumentCombine = Replace$(hLang.Retrieve("DocumentCombine", .DialogDocumentCombine),"/n",vbCrLf)
  .DialogDocumentCombineAll = Replace$(hLang.Retrieve("DocumentCombineAll", .DialogDocumentCombineAll),"/n",vbCrLf)
  .DialogDocumentCombineAllSend = Replace$(hLang.Retrieve("DocumentCombineAllSend", .DialogDocumentCombineAllSend),"/n",vbCrLf)
  .DialogDocumentDelete = Replace$(hLang.Retrieve("DocumentDelete", .DialogDocumentDelete),"/n",vbCrLf)
  .DialogDocumentDown = Replace$(hLang.Retrieve("DocumentDown", .DialogDocumentDown),"/n",vbCrLf)
  .DialogDocumentPrint = Replace$(hLang.Retrieve("DocumentPrint", .DialogDocumentPrint),"/n",vbCrLf)
  .DialogDocumentSave = Replace$(hLang.Retrieve("DocumentSave", .DialogDocumentSave),"/n",vbCrLf)
  .DialogDocumentSend = Replace$(hLang.Retrieve("DocumentSend", .DialogDocumentSend),"/n",vbCrLf)
  .DialogDocumentTop = Replace$(hLang.Retrieve("DocumentTop", .DialogDocumentTop),"/n",vbCrLf)
  .DialogDocumentUp = Replace$(hLang.Retrieve("DocumentUp", .DialogDocumentUp),"/n",vbCrLf)
  .DialogEmailAddress = Replace$(hLang.Retrieve("EmailAddress", .DialogEmailAddress),"/n",vbCrLf)
  .DialogInfo = Replace$(hLang.Retrieve("Info", .DialogInfo),"/n",vbCrLf)
  .DialogInfoCheckUpdates = Replace$(hLang.Retrieve("InfoCheckUpdates", .DialogInfoCheckUpdates),"/n",vbCrLf)
  .DialogInfoHomepage = Replace$(hLang.Retrieve("InfoHomepage", .DialogInfoHomepage),"/n",vbCrLf)
  .DialogInfoInfo = Replace$(hLang.Retrieve("InfoInfo", .DialogInfoInfo),"/n",vbCrLf)
  .DialogInfoPaypal = Replace$(hLang.Retrieve("InfoPaypal", .DialogInfoPaypal),"/n",vbCrLf)
  .DialogInfoPDFCreatorSourceforge = Replace$(hLang.Retrieve("InfoPDFCreatorSourceforge", .DialogInfoPDFCreatorSourceforge),"/n",vbCrLf)
  .DialogLanguage = Replace$(hLang.Retrieve("Language", .DialogLanguage),"/n",vbCrLf)
  .DialogPrinter = Replace$(hLang.Retrieve("Printer", .DialogPrinter),"/n",vbCrLf)
  .DialogPrinterClose = Replace$(hLang.Retrieve("PrinterClose", .DialogPrinterClose),"/n",vbCrLf)
  .DialogPrinterLogfile = Replace$(hLang.Retrieve("PrinterLogfile", .DialogPrinterLogfile),"/n",vbCrLf)
  .DialogPrinterLogfiles = Replace$(hLang.Retrieve("PrinterLogfiles", .DialogPrinterLogfiles),"/n",vbCrLf)
  .DialogPrinterLogging = Replace$(hLang.Retrieve("PrinterLogging", .DialogPrinterLogging),"/n",vbCrLf)
  .DialogPrinterOptions = Replace$(hLang.Retrieve("PrinterOptions", .DialogPrinterOptions),"/n",vbCrLf)
  .DialogPrinterPrinterStop = Replace$(hLang.Retrieve("PrinterPrinterStop", .DialogPrinterPrinterStop),"/n",vbCrLf)
  .DialogView = Replace$(hLang.Retrieve("View", .DialogView),"/n",vbCrLf)
  .DialogViewStatusbar = Replace$(hLang.Retrieve("ViewStatusbar", .DialogViewStatusbar),"/n",vbCrLf)
  .DialogViewToolbars = Replace$(hLang.Retrieve("ViewToolbars", .DialogViewToolbars),"/n",vbCrLf)
  .DialogViewToolbarsEmail = Replace$(hLang.Retrieve("ViewToolbarsEmail", .DialogViewToolbarsEmail),"/n",vbCrLf)
  .DialogViewToolbarsStandard = Replace$(hLang.Retrieve("ViewToolbarsStandard", .DialogViewToolbarsStandard),"/n",vbCrLf)
 End With
 Set hLang = Nothing
End Sub

Private Sub LoadListStrings(ByVal Languagefile As String)
 Dim hLang As New clsHash
 ReadINISection Languagefile, "List", hLang
 With LanguageStrings
  .ListAddFile = Replace$(hLang.Retrieve("AddFile", .ListAddFile),"/n",vbCrLf)
  .ListAllFiles = Replace$(hLang.Retrieve("AllFiles", .ListAllFiles),"/n",vbCrLf)
  .ListBytes = Replace$(hLang.Retrieve("Bytes", .ListBytes),"/n",vbCrLf)
  .ListDate = Replace$(hLang.Retrieve("Date", .ListDate),"/n",vbCrLf)
  .ListDocumenttitle = Replace$(hLang.Retrieve("Documenttitle", .ListDocumenttitle),"/n",vbCrLf)
  .ListFilename = Replace$(hLang.Retrieve("Filename", .ListFilename),"/n",vbCrLf)
  .ListGBytes = Replace$(hLang.Retrieve("GBytes", .ListGBytes),"/n",vbCrLf)
  .ListKBytes = Replace$(hLang.Retrieve("KBytes", .ListKBytes),"/n",vbCrLf)
  .ListMBytes = Replace$(hLang.Retrieve("MBytes", .ListMBytes),"/n",vbCrLf)
  .ListPDFFiles = Replace$(hLang.Retrieve("PDFFiles", .ListPDFFiles),"/n",vbCrLf)
  .ListPostscriptFiles = Replace$(hLang.Retrieve("PostscriptFiles", .ListPostscriptFiles),"/n",vbCrLf)
  .ListPrinting = Replace$(hLang.Retrieve("Printing", .ListPrinting),"/n",vbCrLf)
  .ListSize = Replace$(hLang.Retrieve("Size", .ListSize),"/n",vbCrLf)
  .ListStatus = Replace$(hLang.Retrieve("Status", .ListStatus),"/n",vbCrLf)
  .ListWaiting = Replace$(hLang.Retrieve("Waiting", .ListWaiting),"/n",vbCrLf)
 End With
 Set hLang = Nothing
End Sub

Private Sub LoadLoggingStrings(ByVal Languagefile As String)
 Dim hLang As New clsHash
 ReadINISection Languagefile, "Logging", hLang
 With LanguageStrings
  .LoggingClear = Replace$(hLang.Retrieve("Clear", .LoggingClear),"/n",vbCrLf)
  .LoggingClose = Replace$(hLang.Retrieve("Close", .LoggingClose),"/n",vbCrLf)
  .LoggingLogfile = Replace$(hLang.Retrieve("Logfile", .LoggingLogfile),"/n",vbCrLf)
 End With
 Set hLang = Nothing
End Sub

Private Sub LoadMessagesStrings(ByVal Languagefile As String)
 Dim hLang As New clsHash
 ReadINISection Languagefile, "Messages", hLang
 With LanguageStrings
  .MessagesMsg01 = Replace$(hLang.Retrieve("Msg01", .MessagesMsg01),"/n",vbCrLf)
  .MessagesMsg02 = Replace$(hLang.Retrieve("Msg02", .MessagesMsg02),"/n",vbCrLf)
  .MessagesMsg03 = Replace$(hLang.Retrieve("Msg03", .MessagesMsg03),"/n",vbCrLf)
  .MessagesMsg04 = Replace$(hLang.Retrieve("Msg04", .MessagesMsg04),"/n",vbCrLf)
  .MessagesMsg05 = Replace$(hLang.Retrieve("Msg05", .MessagesMsg05),"/n",vbCrLf)
  .MessagesMsg06 = Replace$(hLang.Retrieve("Msg06", .MessagesMsg06),"/n",vbCrLf)
  .MessagesMsg07 = Replace$(hLang.Retrieve("Msg07", .MessagesMsg07),"/n",vbCrLf)
  .MessagesMsg08 = Replace$(hLang.Retrieve("Msg08", .MessagesMsg08),"/n",vbCrLf)
  .MessagesMsg09 = Replace$(hLang.Retrieve("Msg09", .MessagesMsg09),"/n",vbCrLf)
  .MessagesMsg10 = Replace$(hLang.Retrieve("Msg10", .MessagesMsg10),"/n",vbCrLf)
  .MessagesMsg11 = Replace$(hLang.Retrieve("Msg11", .MessagesMsg11),"/n",vbCrLf)
  .MessagesMsg12 = Replace$(hLang.Retrieve("Msg12", .MessagesMsg12),"/n",vbCrLf)
  .MessagesMsg13 = Replace$(hLang.Retrieve("Msg13", .MessagesMsg13),"/n",vbCrLf)
  .MessagesMsg14 = Replace$(hLang.Retrieve("Msg14", .MessagesMsg14),"/n",vbCrLf)
  .MessagesMsg15 = Replace$(hLang.Retrieve("Msg15", .MessagesMsg15),"/n",vbCrLf)
  .MessagesMsg16 = Replace$(hLang.Retrieve("Msg16", .MessagesMsg16),"/n",vbCrLf)
  .MessagesMsg17 = Replace$(hLang.Retrieve("Msg17", .MessagesMsg17),"/n",vbCrLf)
  .MessagesMsg19 = Replace$(hLang.Retrieve("Msg19", .MessagesMsg19),"/n",vbCrLf)
  .MessagesMsg20 = Replace$(hLang.Retrieve("Msg20", .MessagesMsg20),"/n",vbCrLf)
  .MessagesMsg21 = Replace$(hLang.Retrieve("Msg21", .MessagesMsg21),"/n",vbCrLf)
  .MessagesMsg22 = Replace$(hLang.Retrieve("Msg22", .MessagesMsg22),"/n",vbCrLf)
  .MessagesMsg23 = Replace$(hLang.Retrieve("Msg23", .MessagesMsg23),"/n",vbCrLf)
  .MessagesMsg24 = Replace$(hLang.Retrieve("Msg24", .MessagesMsg24),"/n",vbCrLf)
  .MessagesMsg25 = Replace$(hLang.Retrieve("Msg25", .MessagesMsg25),"/n",vbCrLf)
  .MessagesMsg26 = Replace$(hLang.Retrieve("Msg26", .MessagesMsg26),"/n",vbCrLf)
  .MessagesMsg27 = Replace$(hLang.Retrieve("Msg27", .MessagesMsg27),"/n",vbCrLf)
  .MessagesMsg28 = Replace$(hLang.Retrieve("Msg28", .MessagesMsg28),"/n",vbCrLf)
  .MessagesMsg29 = Replace$(hLang.Retrieve("Msg29", .MessagesMsg29),"/n",vbCrLf)
  .MessagesMsg30 = Replace$(hLang.Retrieve("Msg30", .MessagesMsg30),"/n",vbCrLf)
  .MessagesMsg31 = Replace$(hLang.Retrieve("Msg31", .MessagesMsg31),"/n",vbCrLf)
  .MessagesMsg32 = Replace$(hLang.Retrieve("Msg32", .MessagesMsg32),"/n",vbCrLf)
  .MessagesMsg33 = Replace$(hLang.Retrieve("Msg33", .MessagesMsg33),"/n",vbCrLf)
  .MessagesMsg34 = Replace$(hLang.Retrieve("Msg34", .MessagesMsg34),"/n",vbCrLf)
  .MessagesMsg35 = Replace$(hLang.Retrieve("Msg35", .MessagesMsg35),"/n",vbCrLf)
  .MessagesMsg36 = Replace$(hLang.Retrieve("Msg36", .MessagesMsg36),"/n",vbCrLf)
 End With
 Set hLang = Nothing
End Sub

Private Sub LoadOptionsStrings(ByVal Languagefile As String)
 Dim hLang As New clsHash
 ReadINISection Languagefile, "Options", hLang
 With LanguageStrings
  .OptionsAssociatePSFiles = Replace$(hLang.Retrieve("AssociatePSFiles", .OptionsAssociatePSFiles),"/n",vbCrLf)
  .OptionsAutosaveDirectoryPrompt = Replace$(hLang.Retrieve("AutosaveDirectoryPrompt", .OptionsAutosaveDirectoryPrompt),"/n",vbCrLf)
  .OptionsAutosaveFilename = Replace$(hLang.Retrieve("AutosaveFilename", .OptionsAutosaveFilename),"/n",vbCrLf)
  .OptionsAutosaveFilenameTokens = Replace$(hLang.Retrieve("AutosaveFilenameTokens", .OptionsAutosaveFilenameTokens),"/n",vbCrLf)
  .OptionsAutosaveFormat = Replace$(hLang.Retrieve("AutosaveFormat", .OptionsAutosaveFormat),"/n",vbCrLf)
  .OptionsAutosaveStartStandardProgram = Replace$(hLang.Retrieve("AutosaveStartStandardProgram", .OptionsAutosaveStartStandardProgram),"/n",vbCrLf)
  .OptionsBitmapResolution = Replace$(hLang.Retrieve("BitmapResolution", .OptionsBitmapResolution),"/n",vbCrLf)
  .OptionsBMPColorscount01 = Replace$(hLang.Retrieve("BMPColorscount01", .OptionsBMPColorscount01),"/n",vbCrLf)
  .OptionsBMPColorscount02 = Replace$(hLang.Retrieve("BMPColorscount02", .OptionsBMPColorscount02),"/n",vbCrLf)
  .OptionsBMPColorscount03 = Replace$(hLang.Retrieve("BMPColorscount03", .OptionsBMPColorscount03),"/n",vbCrLf)
  .OptionsBMPColorscount04 = Replace$(hLang.Retrieve("BMPColorscount04", .OptionsBMPColorscount04),"/n",vbCrLf)
  .OptionsBMPColorscount05 = Replace$(hLang.Retrieve("BMPColorscount05", .OptionsBMPColorscount05),"/n",vbCrLf)
  .OptionsBMPColorscount06 = Replace$(hLang.Retrieve("BMPColorscount06", .OptionsBMPColorscount06),"/n",vbCrLf)
  .OptionsBMPColorscount07 = Replace$(hLang.Retrieve("BMPColorscount07", .OptionsBMPColorscount07),"/n",vbCrLf)
  .OptionsBMPDescription = Replace$(hLang.Retrieve("BMPDescription", .OptionsBMPDescription),"/n",vbCrLf)
  .OptionsBMPSymbol = Replace$(hLang.Retrieve("BMPSymbol", .OptionsBMPSymbol),"/n",vbCrLf)
  .OptionsCancel = Replace$(hLang.Retrieve("Cancel", .OptionsCancel),"/n",vbCrLf)
  .OptionsDirectoriesGSBin = Replace$(hLang.Retrieve("DirectoriesGSBin", .OptionsDirectoriesGSBin),"/n",vbCrLf)
  .OptionsDirectoriesGSFonts = Replace$(hLang.Retrieve("DirectoriesGSFonts", .OptionsDirectoriesGSFonts),"/n",vbCrLf)
  .OptionsDirectoriesGSLibraries = Replace$(hLang.Retrieve("DirectoriesGSLibraries", .OptionsDirectoriesGSLibraries),"/n",vbCrLf)
  .OptionsDirectoriesTempPath = Replace$(hLang.Retrieve("DirectoriesTempPath", .OptionsDirectoriesTempPath),"/n",vbCrLf)
  .OptionsDocument = Replace$(hLang.Retrieve("Document", .OptionsDocument),"/n",vbCrLf)
  .OptionsEPSDescription = Replace$(hLang.Retrieve("EPSDescription", .OptionsEPSDescription),"/n",vbCrLf)
  .OptionsEPSFiles = Replace$(hLang.Retrieve("EPSFiles", .OptionsEPSFiles),"/n",vbCrLf)
  .OptionsEPSSymbol = Replace$(hLang.Retrieve("EPSSymbol", .OptionsEPSSymbol),"/n",vbCrLf)
  .OptionsGhostscriptBinariesDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptBinariesDirectoryPrompt", .OptionsGhostscriptBinariesDirectoryPrompt),"/n",vbCrLf)
  .OptionsGhostscriptFontsDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptFontsDirectoryPrompt", .OptionsGhostscriptFontsDirectoryPrompt),"/n",vbCrLf)
  .OptionsGhostscriptInternal = Replace$(hLang.Retrieve("GhostscriptInternal", .OptionsGhostscriptInternal),"/n",vbCrLf)
  .OptionsGhostscriptLibrariesDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptLibrariesDirectoryPrompt", .OptionsGhostscriptLibrariesDirectoryPrompt),"/n",vbCrLf)
  .OptionsGhostscriptResourceDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptResourceDirectoryPrompt", .OptionsGhostscriptResourceDirectoryPrompt),"/n",vbCrLf)
  .OptionsGhostscriptversion = Replace$(hLang.Retrieve("Ghostscriptversion", .OptionsGhostscriptversion),"/n",vbCrLf)
  .OptionsImageSettings = Replace$(hLang.Retrieve("ImageSettings", .OptionsImageSettings),"/n",vbCrLf)
  .OptionsJavaPath = Replace$(hLang.Retrieve("JavaPath", .OptionsJavaPath),"/n",vbCrLf)
  .OptionsJPEGColorscount01 = Replace$(hLang.Retrieve("JPEGColorscount01", .OptionsJPEGColorscount01),"/n",vbCrLf)
  .OptionsJPEGColorscount02 = Replace$(hLang.Retrieve("JPEGColorscount02", .OptionsJPEGColorscount02),"/n",vbCrLf)
  .OptionsJPEGDescription = Replace$(hLang.Retrieve("JPEGDescription", .OptionsJPEGDescription),"/n",vbCrLf)
  .OptionsJPEGQuality = Replace$(hLang.Retrieve("JPEGQuality", .OptionsJPEGQuality),"/n",vbCrLf)
  .OptionsJPEGSymbol = Replace$(hLang.Retrieve("JPEGSymbol", .OptionsJPEGSymbol),"/n",vbCrLf)
  .OptionsOwnerPass = Replace$(hLang.Retrieve("OwnerPass", .OptionsOwnerPass),"/n",vbCrLf)
  .OptionsPassCancel = Replace$(hLang.Retrieve("PassCancel", .OptionsPassCancel),"/n",vbCrLf)
  .OptionsPassOK = Replace$(hLang.Retrieve("PassOK", .OptionsPassOK),"/n",vbCrLf)
  .OptionsPCXColorscount01 = Replace$(hLang.Retrieve("PCXColorscount01", .OptionsPCXColorscount01),"/n",vbCrLf)
  .OptionsPCXColorscount02 = Replace$(hLang.Retrieve("PCXColorscount02", .OptionsPCXColorscount02),"/n",vbCrLf)
  .OptionsPCXColorscount03 = Replace$(hLang.Retrieve("PCXColorscount03", .OptionsPCXColorscount03),"/n",vbCrLf)
  .OptionsPCXColorscount04 = Replace$(hLang.Retrieve("PCXColorscount04", .OptionsPCXColorscount04),"/n",vbCrLf)
  .OptionsPCXColorscount05 = Replace$(hLang.Retrieve("PCXColorscount05", .OptionsPCXColorscount05),"/n",vbCrLf)
  .OptionsPCXColorscount06 = Replace$(hLang.Retrieve("PCXColorscount06", .OptionsPCXColorscount06),"/n",vbCrLf)
  .OptionsPCXDescription = Replace$(hLang.Retrieve("PCXDescription", .OptionsPCXDescription),"/n",vbCrLf)
  .OptionsPCXSymbol = Replace$(hLang.Retrieve("PCXSymbol", .OptionsPCXSymbol),"/n",vbCrLf)
  .OptionsPDFAllowAssembly = Replace$(hLang.Retrieve("PDFAllowAssembly", .OptionsPDFAllowAssembly),"/n",vbCrLf)
  .OptionsPDFAllowDegradedPrinting = Replace$(hLang.Retrieve("PDFAllowDegradedPrinting", .OptionsPDFAllowDegradedPrinting),"/n",vbCrLf)
  .OptionsPDFAllowFillIn = Replace$(hLang.Retrieve("PDFAllowFillIn", .OptionsPDFAllowFillIn),"/n",vbCrLf)
  .OptionsPDFAllowScreenReaders = Replace$(hLang.Retrieve("PDFAllowScreenReaders", .OptionsPDFAllowScreenReaders),"/n",vbCrLf)
  .OptionsPDFColors = Replace$(hLang.Retrieve("PDFColors", .OptionsPDFColors),"/n",vbCrLf)
  .OptionsPDFColorsCaption = Replace$(hLang.Retrieve("PDFColorsCaption", .OptionsPDFColorsCaption),"/n",vbCrLf)
  .OptionsPDFColorsCMYKtoRGB = Replace$(hLang.Retrieve("PDFColorsCMYKtoRGB", .OptionsPDFColorsCMYKtoRGB),"/n",vbCrLf)
  .OptionsPDFColorsColorModel01 = Replace$(hLang.Retrieve("PDFColorsColorModel01", .OptionsPDFColorsColorModel01),"/n",vbCrLf)
  .OptionsPDFColorsColorModel02 = Replace$(hLang.Retrieve("PDFColorsColorModel02", .OptionsPDFColorsColorModel02),"/n",vbCrLf)
  .OptionsPDFColorsColorModel03 = Replace$(hLang.Retrieve("PDFColorsColorModel03", .OptionsPDFColorsColorModel03),"/n",vbCrLf)
  .OptionsPDFColorsColorOptions = Replace$(hLang.Retrieve("PDFColorsColorOptions", .OptionsPDFColorsColorOptions),"/n",vbCrLf)
  .OptionsPDFColorsPreserveHalftone = Replace$(hLang.Retrieve("PDFColorsPreserveHalftone", .OptionsPDFColorsPreserveHalftone),"/n",vbCrLf)
  .OptionsPDFColorsPreserveOverprint = Replace$(hLang.Retrieve("PDFColorsPreserveOverprint", .OptionsPDFColorsPreserveOverprint),"/n",vbCrLf)
  .OptionsPDFColorsPreserveTransfer = Replace$(hLang.Retrieve("PDFColorsPreserveTransfer", .OptionsPDFColorsPreserveTransfer),"/n",vbCrLf)
  .OptionsPDFCompression = Replace$(hLang.Retrieve("PDFCompression", .OptionsPDFCompression),"/n",vbCrLf)
  .OptionsPDFCompressionCaption = Replace$(hLang.Retrieve("PDFCompressionCaption", .OptionsPDFCompressionCaption),"/n",vbCrLf)
  .OptionsPDFCompressionColor = Replace$(hLang.Retrieve("PDFCompressionColor", .OptionsPDFCompressionColor),"/n",vbCrLf)
  .OptionsPDFCompressionColorComp = Replace$(hLang.Retrieve("PDFCompressionColorComp", .OptionsPDFCompressionColorComp),"/n",vbCrLf)
  .OptionsPDFCompressionColorComp01 = Replace$(hLang.Retrieve("PDFCompressionColorComp01", .OptionsPDFCompressionColorComp01),"/n",vbCrLf)
  .OptionsPDFCompressionColorComp02 = Replace$(hLang.Retrieve("PDFCompressionColorComp02", .OptionsPDFCompressionColorComp02),"/n",vbCrLf)
  .OptionsPDFCompressionColorComp03 = Replace$(hLang.Retrieve("PDFCompressionColorComp03", .OptionsPDFCompressionColorComp03),"/n",vbCrLf)
  .OptionsPDFCompressionColorComp04 = Replace$(hLang.Retrieve("PDFCompressionColorComp04", .OptionsPDFCompressionColorComp04),"/n",vbCrLf)
  .OptionsPDFCompressionColorComp05 = Replace$(hLang.Retrieve("PDFCompressionColorComp05", .OptionsPDFCompressionColorComp05),"/n",vbCrLf)
  .OptionsPDFCompressionColorComp06 = Replace$(hLang.Retrieve("PDFCompressionColorComp06", .OptionsPDFCompressionColorComp06),"/n",vbCrLf)
  .OptionsPDFCompressionColorComp07 = Replace$(hLang.Retrieve("PDFCompressionColorComp07", .OptionsPDFCompressionColorComp07),"/n",vbCrLf)
  .OptionsPDFCompressionColorComp08 = Replace$(hLang.Retrieve("PDFCompressionColorComp08", .OptionsPDFCompressionColorComp08),"/n",vbCrLf)
  .OptionsPDFCompressionColorRes = Replace$(hLang.Retrieve("PDFCompressionColorRes", .OptionsPDFCompressionColorRes),"/n",vbCrLf)
  .OptionsPDFCompressionColorResample = Replace$(hLang.Retrieve("PDFCompressionColorResample", .OptionsPDFCompressionColorResample),"/n",vbCrLf)
  .OptionsPDFCompressionColorResample01 = Replace$(hLang.Retrieve("PDFCompressionColorResample01", .OptionsPDFCompressionColorResample01),"/n",vbCrLf)
  .OptionsPDFCompressionColorResample02 = Replace$(hLang.Retrieve("PDFCompressionColorResample02", .OptionsPDFCompressionColorResample02),"/n",vbCrLf)
  .OptionsPDFCompressionColorResample03 = Replace$(hLang.Retrieve("PDFCompressionColorResample03", .OptionsPDFCompressionColorResample03),"/n",vbCrLf)
  .OptionsPDFCompressionGrey = Replace$(hLang.Retrieve("PDFCompressionGrey", .OptionsPDFCompressionGrey),"/n",vbCrLf)
  .OptionsPDFCompressionGreyComp = Replace$(hLang.Retrieve("PDFCompressionGreyComp", .OptionsPDFCompressionGreyComp),"/n",vbCrLf)
  .OptionsPDFCompressionGreyComp01 = Replace$(hLang.Retrieve("PDFCompressionGreyComp01", .OptionsPDFCompressionGreyComp01),"/n",vbCrLf)
  .OptionsPDFCompressionGreyComp02 = Replace$(hLang.Retrieve("PDFCompressionGreyComp02", .OptionsPDFCompressionGreyComp02),"/n",vbCrLf)
  .OptionsPDFCompressionGreyComp03 = Replace$(hLang.Retrieve("PDFCompressionGreyComp03", .OptionsPDFCompressionGreyComp03),"/n",vbCrLf)
  .OptionsPDFCompressionGreyComp04 = Replace$(hLang.Retrieve("PDFCompressionGreyComp04", .OptionsPDFCompressionGreyComp04),"/n",vbCrLf)
  .OptionsPDFCompressionGreyComp05 = Replace$(hLang.Retrieve("PDFCompressionGreyComp05", .OptionsPDFCompressionGreyComp05),"/n",vbCrLf)
  .OptionsPDFCompressionGreyComp06 = Replace$(hLang.Retrieve("PDFCompressionGreyComp06", .OptionsPDFCompressionGreyComp06),"/n",vbCrLf)
  .OptionsPDFCompressionGreyComp07 = Replace$(hLang.Retrieve("PDFCompressionGreyComp07", .OptionsPDFCompressionGreyComp07),"/n",vbCrLf)
  .OptionsPDFCompressionGreyComp08 = Replace$(hLang.Retrieve("PDFCompressionGreyComp08", .OptionsPDFCompressionGreyComp08),"/n",vbCrLf)
  .OptionsPDFCompressionGreyRes = Replace$(hLang.Retrieve("PDFCompressionGreyRes", .OptionsPDFCompressionGreyRes),"/n",vbCrLf)
  .OptionsPDFCompressionGreyResample = Replace$(hLang.Retrieve("PDFCompressionGreyResample", .OptionsPDFCompressionGreyResample),"/n",vbCrLf)
  .OptionsPDFCompressionGreyResample01 = Replace$(hLang.Retrieve("PDFCompressionGreyResample01", .OptionsPDFCompressionGreyResample01),"/n",vbCrLf)
  .OptionsPDFCompressionGreyResample02 = Replace$(hLang.Retrieve("PDFCompressionGreyResample02", .OptionsPDFCompressionGreyResample02),"/n",vbCrLf)
  .OptionsPDFCompressionGreyResample03 = Replace$(hLang.Retrieve("PDFCompressionGreyResample03", .OptionsPDFCompressionGreyResample03),"/n",vbCrLf)
  .OptionsPDFCompressionMono = Replace$(hLang.Retrieve("PDFCompressionMono", .OptionsPDFCompressionMono),"/n",vbCrLf)
  .OptionsPDFCompressionMonoComp = Replace$(hLang.Retrieve("PDFCompressionMonoComp", .OptionsPDFCompressionMonoComp),"/n",vbCrLf)
  .OptionsPDFCompressionMonoComp01 = Replace$(hLang.Retrieve("PDFCompressionMonoComp01", .OptionsPDFCompressionMonoComp01),"/n",vbCrLf)
  .OptionsPDFCompressionMonoComp02 = Replace$(hLang.Retrieve("PDFCompressionMonoComp02", .OptionsPDFCompressionMonoComp02),"/n",vbCrLf)
  .OptionsPDFCompressionMonoComp03 = Replace$(hLang.Retrieve("PDFCompressionMonoComp03", .OptionsPDFCompressionMonoComp03),"/n",vbCrLf)
  .OptionsPDFCompressionMonoComp04 = Replace$(hLang.Retrieve("PDFCompressionMonoComp04", .OptionsPDFCompressionMonoComp04),"/n",vbCrLf)
  .OptionsPDFCompressionMonoRes = Replace$(hLang.Retrieve("PDFCompressionMonoRes", .OptionsPDFCompressionMonoRes),"/n",vbCrLf)
  .OptionsPDFCompressionMonoResample = Replace$(hLang.Retrieve("PDFCompressionMonoResample", .OptionsPDFCompressionMonoResample),"/n",vbCrLf)
  .OptionsPDFCompressionMonoResample01 = Replace$(hLang.Retrieve("PDFCompressionMonoResample01", .OptionsPDFCompressionMonoResample01),"/n",vbCrLf)
  .OptionsPDFCompressionMonoResample02 = Replace$(hLang.Retrieve("PDFCompressionMonoResample02", .OptionsPDFCompressionMonoResample02),"/n",vbCrLf)
  .OptionsPDFCompressionMonoResample03 = Replace$(hLang.Retrieve("PDFCompressionMonoResample03", .OptionsPDFCompressionMonoResample03),"/n",vbCrLf)
  .OptionsPDFCompressionTextComp = Replace$(hLang.Retrieve("PDFCompressionTextComp", .OptionsPDFCompressionTextComp),"/n",vbCrLf)
  .OptionsPDFDescription = Replace$(hLang.Retrieve("PDFDescription", .OptionsPDFDescription),"/n",vbCrLf)
  .OptionsPDFDisallowCopy = Replace$(hLang.Retrieve("PDFDisallowCopy", .OptionsPDFDisallowCopy),"/n",vbCrLf)
  .OptionsPDFDisallowModify = Replace$(hLang.Retrieve("PDFDisallowModify", .OptionsPDFDisallowModify),"/n",vbCrLf)
  .OptionsPDFDisallowModifyComments = Replace$(hLang.Retrieve("PDFDisallowModifyComments", .OptionsPDFDisallowModifyComments),"/n",vbCrLf)
  .OptionsPDFDisallowPrint = Replace$(hLang.Retrieve("PDFDisallowPrint", .OptionsPDFDisallowPrint),"/n",vbCrLf)
  .OptionsPDFDisallowUser = Replace$(hLang.Retrieve("PDFDisallowUser", .OptionsPDFDisallowUser),"/n",vbCrLf)
  .OptionsPDFEncryptionHigh = Replace$(hLang.Retrieve("PDFEncryptionHigh", .OptionsPDFEncryptionHigh),"/n",vbCrLf)
  .OptionsPDFEncryptionLevel = Replace$(hLang.Retrieve("PDFEncryptionLevel", .OptionsPDFEncryptionLevel),"/n",vbCrLf)
  .OptionsPDFEncryptionLow = Replace$(hLang.Retrieve("PDFEncryptionLow", .OptionsPDFEncryptionLow),"/n",vbCrLf)
  .OptionsPDFEncryptor = Replace$(hLang.Retrieve("PDFEncryptor", .OptionsPDFEncryptor),"/n",vbCrLf)
  .OptionsPDFEnhancedPermissions = Replace$(hLang.Retrieve("PDFEnhancedPermissions", .OptionsPDFEnhancedPermissions),"/n",vbCrLf)
  .OptionsPDFEnterPasswords = Replace$(hLang.Retrieve("PDFEnterPasswords", .OptionsPDFEnterPasswords),"/n",vbCrLf)
  .OptionsPDFFonts = Replace$(hLang.Retrieve("PDFFonts", .OptionsPDFFonts),"/n",vbCrLf)
  .OptionsPDFFontsCaption = Replace$(hLang.Retrieve("PDFFontsCaption", .OptionsPDFFontsCaption),"/n",vbCrLf)
  .OptionsPDFFontsEmbedAll = Replace$(hLang.Retrieve("PDFFontsEmbedAll", .OptionsPDFFontsEmbedAll),"/n",vbCrLf)
  .OptionsPDFFontsSubSetFonts = Replace$(hLang.Retrieve("PDFFontsSubSetFonts", .OptionsPDFFontsSubSetFonts),"/n",vbCrLf)
  .OptionsPDFGeneral = Replace$(hLang.Retrieve("PDFGeneral", .OptionsPDFGeneral),"/n",vbCrLf)
  .OptionsPDFGeneralASCII85 = Replace$(hLang.Retrieve("PDFGeneralASCII85", .OptionsPDFGeneralASCII85),"/n",vbCrLf)
  .OptionsPDFGeneralAutorotate = Replace$(hLang.Retrieve("PDFGeneralAutorotate", .OptionsPDFGeneralAutorotate),"/n",vbCrLf)
  .OptionsPDFGeneralCaption = Replace$(hLang.Retrieve("PDFGeneralCaption", .OptionsPDFGeneralCaption),"/n",vbCrLf)
  .OptionsPDFGeneralCompatibility = Replace$(hLang.Retrieve("PDFGeneralCompatibility", .OptionsPDFGeneralCompatibility),"/n",vbCrLf)
  .OptionsPDFGeneralCompatibility01 = Replace$(hLang.Retrieve("PDFGeneralCompatibility01", .OptionsPDFGeneralCompatibility01),"/n",vbCrLf)
  .OptionsPDFGeneralCompatibility02 = Replace$(hLang.Retrieve("PDFGeneralCompatibility02", .OptionsPDFGeneralCompatibility02),"/n",vbCrLf)
  .OptionsPDFGeneralCompatibility03 = Replace$(hLang.Retrieve("PDFGeneralCompatibility03", .OptionsPDFGeneralCompatibility03),"/n",vbCrLf)
  .OptionsPDFGeneralOverprint = Replace$(hLang.Retrieve("PDFGeneralOverprint", .OptionsPDFGeneralOverprint),"/n",vbCrLf)
  .OptionsPDFGeneralOverprint01 = Replace$(hLang.Retrieve("PDFGeneralOverprint01", .OptionsPDFGeneralOverprint01),"/n",vbCrLf)
  .OptionsPDFGeneralOverprint02 = Replace$(hLang.Retrieve("PDFGeneralOverprint02", .OptionsPDFGeneralOverprint02),"/n",vbCrLf)
  .OptionsPDFGeneralResolution = Replace$(hLang.Retrieve("PDFGeneralResolution", .OptionsPDFGeneralResolution),"/n",vbCrLf)
  .OptionsPDFGeneralRotate01 = Replace$(hLang.Retrieve("PDFGeneralRotate01", .OptionsPDFGeneralRotate01),"/n",vbCrLf)
  .OptionsPDFGeneralRotate02 = Replace$(hLang.Retrieve("PDFGeneralRotate02", .OptionsPDFGeneralRotate02),"/n",vbCrLf)
  .OptionsPDFGeneralRotate03 = Replace$(hLang.Retrieve("PDFGeneralRotate03", .OptionsPDFGeneralRotate03),"/n",vbCrLf)
  .OptionsPDFOptions = Replace$(hLang.Retrieve("PDFOptions", .OptionsPDFOptions),"/n",vbCrLf)
  .OptionsPDFOwnerPass = Replace$(hLang.Retrieve("PDFOwnerPass", .OptionsPDFOwnerPass),"/n",vbCrLf)
  .OptionsPDFPasswords = Replace$(hLang.Retrieve("PDFPasswords", .OptionsPDFPasswords),"/n",vbCrLf)
  .OptionsPDFRepeatPassword = Replace$(hLang.Retrieve("PDFRepeatPassword", .OptionsPDFRepeatPassword),"/n",vbCrLf)
  .OptionsPDFSecurity = Replace$(hLang.Retrieve("PDFSecurity", .OptionsPDFSecurity),"/n",vbCrLf)
  .OptionsPDFSecurityCaption = Replace$(hLang.Retrieve("PDFSecurityCaption", .OptionsPDFSecurityCaption),"/n",vbCrLf)
  .OptionsPDFSetPassword = Replace$(hLang.Retrieve("PDFSetPassword", .OptionsPDFSetPassword),"/n",vbCrLf)
  .OptionsPDFSymbol = Replace$(hLang.Retrieve("PDFSymbol", .OptionsPDFSymbol),"/n",vbCrLf)
  .OptionsPDFUserPass = Replace$(hLang.Retrieve("PDFUserPass", .OptionsPDFUserPass),"/n",vbCrLf)
  .OptionsPDFUseSecurity = Replace$(hLang.Retrieve("PDFUseSecurity", .OptionsPDFUseSecurity),"/n",vbCrLf)
  .OptionsPNGColorscount01 = Replace$(hLang.Retrieve("PNGColorscount01", .OptionsPNGColorscount01),"/n",vbCrLf)
  .OptionsPNGColorscount02 = Replace$(hLang.Retrieve("PNGColorscount02", .OptionsPNGColorscount02),"/n",vbCrLf)
  .OptionsPNGColorscount03 = Replace$(hLang.Retrieve("PNGColorscount03", .OptionsPNGColorscount03),"/n",vbCrLf)
  .OptionsPNGColorscount04 = Replace$(hLang.Retrieve("PNGColorscount04", .OptionsPNGColorscount04),"/n",vbCrLf)
  .OptionsPNGColorscount05 = Replace$(hLang.Retrieve("PNGColorscount05", .OptionsPNGColorscount05),"/n",vbCrLf)
  .OptionsPNGDescription = Replace$(hLang.Retrieve("PNGDescription", .OptionsPNGDescription),"/n",vbCrLf)
  .OptionsPNGFiles = Replace$(hLang.Retrieve("PNGFiles", .OptionsPNGFiles),"/n",vbCrLf)
  .OptionsPNGSymbol = Replace$(hLang.Retrieve("PNGSymbol", .OptionsPNGSymbol),"/n",vbCrLf)
  .OptionsPrintertempDirectoryPrompt = Replace$(hLang.Retrieve("PrintertempDirectoryPrompt", .OptionsPrintertempDirectoryPrompt),"/n",vbCrLf)
  .OptionsPrintTestpage = Replace$(hLang.Retrieve("PrintTestpage", .OptionsPrintTestpage),"/n",vbCrLf)
  .OptionsProcesspriority = Replace$(hLang.Retrieve("Processpriority", .OptionsProcesspriority),"/n",vbCrLf)
  .OptionsProcesspriorityHigh = Replace$(hLang.Retrieve("ProcesspriorityHigh", .OptionsProcesspriorityHigh),"/n",vbCrLf)
  .OptionsProcesspriorityIdle = Replace$(hLang.Retrieve("ProcesspriorityIdle", .OptionsProcesspriorityIdle),"/n",vbCrLf)
  .OptionsProcesspriorityNormal = Replace$(hLang.Retrieve("ProcesspriorityNormal", .OptionsProcesspriorityNormal),"/n",vbCrLf)
  .OptionsProcesspriorityRealtime = Replace$(hLang.Retrieve("ProcesspriorityRealtime", .OptionsProcesspriorityRealtime),"/n",vbCrLf)
  .OptionsProgramAutosaveDescription = Replace$(hLang.Retrieve("ProgramAutosaveDescription", .OptionsProgramAutosaveDescription),"/n",vbCrLf)
  .OptionsProgramAutosaveSymbol = Replace$(hLang.Retrieve("ProgramAutosaveSymbol", .OptionsProgramAutosaveSymbol),"/n",vbCrLf)
  .OptionsProgramDirectoriesDescription = Replace$(hLang.Retrieve("ProgramDirectoriesDescription", .OptionsProgramDirectoriesDescription),"/n",vbCrLf)
  .OptionsProgramDirectoriesSymbol = Replace$(hLang.Retrieve("ProgramDirectoriesSymbol", .OptionsProgramDirectoriesSymbol),"/n",vbCrLf)
  .OptionsProgramDocumentDescription = Replace$(hLang.Retrieve("ProgramDocumentDescription", .OptionsProgramDocumentDescription),"/n",vbCrLf)
  .OptionsProgramDocumentSymbol = Replace$(hLang.Retrieve("ProgramDocumentSymbol", .OptionsProgramDocumentSymbol),"/n",vbCrLf)
  .OptionsProgramFont = Replace$(hLang.Retrieve("ProgramFont", .OptionsProgramFont),"/n",vbCrLf)
  .OptionsProgramFontCancelTest = Replace$(hLang.Retrieve("ProgramFontCancelTest", .OptionsProgramFontCancelTest),"/n",vbCrLf)
  .OptionsProgramFontcharset = Replace$(hLang.Retrieve("ProgramFontcharset", .OptionsProgramFontcharset),"/n",vbCrLf)
  .OptionsProgramFontDescription = Replace$(hLang.Retrieve("ProgramFontDescription", .OptionsProgramFontDescription),"/n",vbCrLf)
  .OptionsProgramFontSize = Replace$(hLang.Retrieve("ProgramFontSize", .OptionsProgramFontSize),"/n",vbCrLf)
  .OptionsProgramFontSymbol = Replace$(hLang.Retrieve("ProgramFontSymbol", .OptionsProgramFontSymbol),"/n",vbCrLf)
  .OptionsProgramFontTest = Replace$(hLang.Retrieve("ProgramFontTest", .OptionsProgramFontTest),"/n",vbCrLf)
  .OptionsProgramFontTestdescription = Replace$(hLang.Retrieve("ProgramFontTestdescription", .OptionsProgramFontTestdescription),"/n",vbCrLf)
  .OptionsProgramGeneralDescription = Replace$(hLang.Retrieve("ProgramGeneralDescription", .OptionsProgramGeneralDescription),"/n",vbCrLf)
  .OptionsProgramGeneralSymbol = Replace$(hLang.Retrieve("ProgramGeneralSymbol", .OptionsProgramGeneralSymbol),"/n",vbCrLf)
  .OptionsProgramGhostscriptDescription = Replace$(hLang.Retrieve("ProgramGhostscriptDescription", .OptionsProgramGhostscriptDescription),"/n",vbCrLf)
  .OptionsProgramGhostscriptSymbol = Replace$(hLang.Retrieve("ProgramGhostscriptSymbol", .OptionsProgramGhostscriptSymbol),"/n",vbCrLf)
  .OptionsProgramSaveDescription = Replace$(hLang.Retrieve("ProgramSaveDescription", .OptionsProgramSaveDescription),"/n",vbCrLf)
  .OptionsProgramSaveSymbol = Replace$(hLang.Retrieve("ProgramSaveSymbol", .OptionsProgramSaveSymbol),"/n",vbCrLf)
  .OptionsProgramSwitchingDefaultprinter = Replace$(hLang.Retrieve("ProgramSwitchingDefaultprinter", .OptionsProgramSwitchingDefaultprinter),"/n",vbCrLf)
  .OptionsPSDescription = Replace$(hLang.Retrieve("PSDescription", .OptionsPSDescription),"/n",vbCrLf)
  .OptionsPSFiles = Replace$(hLang.Retrieve("PSFiles", .OptionsPSFiles),"/n",vbCrLf)
  .OptionsPSLanguageLevel = Replace$(hLang.Retrieve("PSLanguageLevel", .OptionsPSLanguageLevel),"/n",vbCrLf)
  .OptionsPSSymbol = Replace$(hLang.Retrieve("PSSymbol", .OptionsPSSymbol),"/n",vbCrLf)
  .OptionsRemoveSpaces = Replace$(hLang.Retrieve("RemoveSpaces", .OptionsRemoveSpaces),"/n",vbCrLf)
  .OptionsReset = Replace$(hLang.Retrieve("Reset", .OptionsReset),"/n",vbCrLf)
  .OptionsSave = Replace$(hLang.Retrieve("Save", .OptionsSave),"/n",vbCrLf)
  .OptionsSaveFilename = Replace$(hLang.Retrieve("SaveFilename", .OptionsSaveFilename),"/n",vbCrLf)
  .OptionsSaveFilenameAdd = Replace$(hLang.Retrieve("SaveFilenameAdd", .OptionsSaveFilenameAdd),"/n",vbCrLf)
  .OptionsSaveFilenameChange = Replace$(hLang.Retrieve("SaveFilenameChange", .OptionsSaveFilenameChange),"/n",vbCrLf)
  .OptionsSaveFilenameDelete = Replace$(hLang.Retrieve("SaveFilenameDelete", .OptionsSaveFilenameDelete),"/n",vbCrLf)
  .OptionsSaveFilenameSubstitutions = Replace$(hLang.Retrieve("SaveFilenameSubstitutions", .OptionsSaveFilenameSubstitutions),"/n",vbCrLf)
  .OptionsSaveFilenameSubstitutionsTitle = Replace$(hLang.Retrieve("SaveFilenameSubstitutionsTitle", .OptionsSaveFilenameSubstitutionsTitle),"/n",vbCrLf)
  .OptionsSaveFilenameTokens = Replace$(hLang.Retrieve("SaveFilenameTokens", .OptionsSaveFilenameTokens),"/n",vbCrLf)
  .OptionsSavePasswords = Replace$(hLang.Retrieve("SavePasswords", .OptionsSavePasswords),"/n",vbCrLf)
  .OptionsShellIntegration = Replace$(hLang.Retrieve("ShellIntegration", .OptionsShellIntegration),"/n",vbCrLf)
  .OptionsShellIntegrationAdd = Replace$(hLang.Retrieve("ShellIntegrationAdd", .OptionsShellIntegrationAdd),"/n",vbCrLf)
  .OptionsShellIntegrationCaption = Replace$(hLang.Retrieve("ShellIntegrationCaption", .OptionsShellIntegrationCaption),"/n",vbCrLf)
  .OptionsShellIntegrationRemove = Replace$(hLang.Retrieve("ShellIntegrationRemove", .OptionsShellIntegrationRemove),"/n",vbCrLf)
  .OptionsStandardAuthorToken = Replace$(hLang.Retrieve("StandardAuthorToken", .OptionsStandardAuthorToken),"/n",vbCrLf)
  .OptionsTestpage = Replace$(hLang.Retrieve("Testpage", .OptionsTestpage),"/n",vbCrLf)
  .OptionsTIFFColorscount01 = Replace$(hLang.Retrieve("TIFFColorscount01", .OptionsTIFFColorscount01),"/n",vbCrLf)
  .OptionsTIFFColorscount02 = Replace$(hLang.Retrieve("TIFFColorscount02", .OptionsTIFFColorscount02),"/n",vbCrLf)
  .OptionsTIFFColorscount03 = Replace$(hLang.Retrieve("TIFFColorscount03", .OptionsTIFFColorscount03),"/n",vbCrLf)
  .OptionsTIFFColorscount04 = Replace$(hLang.Retrieve("TIFFColorscount04", .OptionsTIFFColorscount04),"/n",vbCrLf)
  .OptionsTIFFColorscount05 = Replace$(hLang.Retrieve("TIFFColorscount05", .OptionsTIFFColorscount05),"/n",vbCrLf)
  .OptionsTIFFColorscount06 = Replace$(hLang.Retrieve("TIFFColorscount06", .OptionsTIFFColorscount06),"/n",vbCrLf)
  .OptionsTIFFColorscount07 = Replace$(hLang.Retrieve("TIFFColorscount07", .OptionsTIFFColorscount07),"/n",vbCrLf)
  .OptionsTIFFColorscount08 = Replace$(hLang.Retrieve("TIFFColorscount08", .OptionsTIFFColorscount08),"/n",vbCrLf)
  .OptionsTIFFDescription = Replace$(hLang.Retrieve("TIFFDescription", .OptionsTIFFDescription),"/n",vbCrLf)
  .OptionsTIFFSymbol = Replace$(hLang.Retrieve("TIFFSymbol", .OptionsTIFFSymbol),"/n",vbCrLf)
  .OptionsTreeFormats = Replace$(hLang.Retrieve("TreeFormats", .OptionsTreeFormats),"/n",vbCrLf)
  .OptionsTreeProgram = Replace$(hLang.Retrieve("TreeProgram", .OptionsTreeProgram),"/n",vbCrLf)
  .OptionsUseAutosave = Replace$(hLang.Retrieve("UseAutosave", .OptionsUseAutosave),"/n",vbCrLf)
  .OptionsUseAutosaveDirectory = Replace$(hLang.Retrieve("UseAutosaveDirectory", .OptionsUseAutosaveDirectory),"/n",vbCrLf)
  .OptionsUseCreationDateNow = Replace$(hLang.Retrieve("UseCreationDateNow", .OptionsUseCreationDateNow),"/n",vbCrLf)
  .OptionsUserPass = Replace$(hLang.Retrieve("UserPass", .OptionsUserPass),"/n",vbCrLf)
  .OptionsUseStandardauthor = Replace$(hLang.Retrieve("UseStandardauthor", .OptionsUseStandardauthor),"/n",vbCrLf)
 End With
 Set hLang = Nothing
End Sub

Private Sub LoadPrintingStrings(ByVal Languagefile As String)
 Dim hLang As New clsHash
 ReadINISection Languagefile, "Printing", hLang
 With LanguageStrings
  .PrintingAuthor = Replace$(hLang.Retrieve("Author", .PrintingAuthor),"/n",vbCrLf)
  .PrintingBMPFiles = Replace$(hLang.Retrieve("BMPFiles", .PrintingBMPFiles),"/n",vbCrLf)
  .PrintingCancel = Replace$(hLang.Retrieve("Cancel", .PrintingCancel),"/n",vbCrLf)
  .PrintingCollect = Replace$(hLang.Retrieve("Collect", .PrintingCollect),"/n",vbCrLf)
  .PrintingCreationDate = Replace$(hLang.Retrieve("CreationDate", .PrintingCreationDate),"/n",vbCrLf)
  .PrintingDocumentTitle = Replace$(hLang.Retrieve("DocumentTitle", .PrintingDocumentTitle),"/n",vbCrLf)
  .PrintingEMail = Replace$(hLang.Retrieve("EMail", .PrintingEMail),"/n",vbCrLf)
  .PrintingEPSFiles = Replace$(hLang.Retrieve("EPSFiles", .PrintingEPSFiles),"/n",vbCrLf)
  .PrintingJPEGFiles = Replace$(hLang.Retrieve("JPEGFiles", .PrintingJPEGFiles),"/n",vbCrLf)
  .PrintingKeywords = Replace$(hLang.Retrieve("Keywords", .PrintingKeywords),"/n",vbCrLf)
  .PrintingModifyDate = Replace$(hLang.Retrieve("ModifyDate", .PrintingModifyDate),"/n",vbCrLf)
  .PrintingNow = Replace$(hLang.Retrieve("Now", .PrintingNow),"/n",vbCrLf)
  .PrintingPCXFiles = Replace$(hLang.Retrieve("PCXFiles", .PrintingPCXFiles),"/n",vbCrLf)
  .PrintingPDFFiles = Replace$(hLang.Retrieve("PDFFiles", .PrintingPDFFiles),"/n",vbCrLf)
  .PrintingPNGFiles = Replace$(hLang.Retrieve("PNGFiles", .PrintingPNGFiles),"/n",vbCrLf)
  .PrintingPSFiles = Replace$(hLang.Retrieve("PSFiles", .PrintingPSFiles),"/n",vbCrLf)
  .PrintingSave = Replace$(hLang.Retrieve("Save", .PrintingSave),"/n",vbCrLf)
  .PrintingStartStandardProgram = Replace$(hLang.Retrieve("StartStandardProgram", .PrintingStartStandardProgram),"/n",vbCrLf)
  .PrintingStatus = Replace$(hLang.Retrieve("Status", .PrintingStatus),"/n",vbCrLf)
  .PrintingSubject = Replace$(hLang.Retrieve("Subject", .PrintingSubject),"/n",vbCrLf)
  .PrintingTIFFFiles = Replace$(hLang.Retrieve("TIFFFiles", .PrintingTIFFFiles),"/n",vbCrLf)
 End With
 Set hLang = Nothing
End Sub

Private Sub LoadSaveOpenStrings(ByVal Languagefile As String)
 Dim hLang As New clsHash
 ReadINISection Languagefile, "SaveOpen", hLang
 With LanguageStrings
  .SaveOpenAttributes = Replace$(hLang.Retrieve("Attributes", .SaveOpenAttributes),"/n",vbCrLf)
  .SaveOpenCancel = Replace$(hLang.Retrieve("Cancel", .SaveOpenCancel),"/n",vbCrLf)
  .SaveOpenFilename = Replace$(hLang.Retrieve("Filename", .SaveOpenFilename),"/n",vbCrLf)
  .SaveOpenOpen = Replace$(hLang.Retrieve("Open", .SaveOpenOpen),"/n",vbCrLf)
  .SaveOpenOpenTitle = Replace$(hLang.Retrieve("OpenTitle", .SaveOpenOpenTitle),"/n",vbCrLf)
  .SaveOpenSave = Replace$(hLang.Retrieve("Save", .SaveOpenSave),"/n",vbCrLf)
  .SaveOpenSaveTitle = Replace$(hLang.Retrieve("SaveTitle", .SaveOpenSaveTitle),"/n",vbCrLf)
  .SaveOpenSize = Replace$(hLang.Retrieve("Size", .SaveOpenSize),"/n",vbCrLf)
 End With
 Set hLang = Nothing
End Sub

Public Sub InitLanguagesStrings()
 With LanguageStrings
  .CommonAuthor = "Philip Chinery, Frank Heindörfer"
  .CommonLanguagename = "English"
  .CommonTitle = "PDF Print monitor"
  .CommonVersion = "0.9.0"

  .DialogDocument = "&Document"
  .DialogDocumentAdd = "Add"
  .DialogDocumentBottom = "Bottom"
  .DialogDocumentCombine = "Combine"
  .DialogDocumentCombineAll = "Combine all"
  .DialogDocumentCombineAllSend = "Combine all and send"
  .DialogDocumentDelete = "Delete"
  .DialogDocumentDown = "Down"
  .DialogDocumentPrint = "Print"
  .DialogDocumentSave = "Save"
  .DialogDocumentSend = "Send"
  .DialogDocumentTop = "Top"
  .DialogDocumentUp = "Up"
  .DialogEmailAddress = "Email address"
  .DialogInfo = "&?"
  .DialogInfoCheckUpdates = "Check for Updates"
  .DialogInfoHomepage = "Product Homepage"
  .DialogInfoInfo = "About"
  .DialogInfoPaypal = "Paypal"
  .DialogInfoPDFCreatorSourceforge = "PDFCreator on Sourceforge"
  .DialogLanguage = "&Language"
  .DialogPrinter = "&Printer"
  .DialogPrinterClose = "Close"
  .DialogPrinterLogfile = "Logfile"
  .DialogPrinterLogfiles = "Logfiles"
  .DialogPrinterLogging = "Logging"
  .DialogPrinterOptions = "Options"
  .DialogPrinterPrinterStop = "Printer stop"
  .DialogView = "&View"
  .DialogViewStatusbar = "Status Bar"
  .DialogViewToolbars = "&Toolbars"
  .DialogViewToolbarsEmail = "Email"
  .DialogViewToolbarsStandard = "Standard"

  .ListAddFile = "Add a file"
  .ListAllFiles = "All files"
  .ListBytes = "Bytes"
  .ListDate = "Created on"
  .ListDocumenttitle = "Document Title"
  .ListFilename = "Filename"
  .ListGBytes = "GBytes"
  .ListKBytes = "kBytes"
  .ListMBytes = "MBytes"
  .ListPDFFiles = "PDF Files"
  .ListPostscriptFiles = "PostScript Files"
  .ListPrinting = "Printing"
  .ListSize = "Size"
  .ListStatus = "Status"
  .ListWaiting = "Waiting"

  .LoggingClear = "Cl&ear"
  .LoggingClose = "&Close"
  .LoggingLogfile = "Logfile"

  .MessagesMsg01 = "Document in queue."
  .MessagesMsg02 = "Documents in queue."
  .MessagesMsg03 = "Do you wish to reset all settings?"
  .MessagesMsg04 = "Error: Cannot send Email!"
  .MessagesMsg05 = "File already exists. Do you want to overwrite it?"
  .MessagesMsg06 = "This file does not seem to be a postscript file!"
  .MessagesMsg07 = "There is a problem when trying to access this drive or directory!"
  .MessagesMsg08 = "Cannot find gsdll32.dll. Please check the ghostscript-program directory (see options)!"
  .MessagesMsg09 = "The output path does not exist. Do you want to create it?"
  .MessagesMsg10 = "This is not a valid path!"
  .MessagesMsg11 = "There is already such an entry!"
  .MessagesMsg12 = "Please don't use these forbidden characters for a filename!"
  .MessagesMsg13 = "Delete all program settings?"
  .MessagesMsg14 = "The file can not be found!"
  .MessagesMsg15 = "Cannot find gsdll32.dll in this directory!"
  .MessagesMsg16 = "No ghostscript font found in this directory!"
  .MessagesMsg17 = "No files in this directory!"
  .MessagesMsg19 = "You need either pdfenc or AFPL Ghostscript greater than, or equal to, version 8.14!"
  .MessagesMsg20 = "There was a problem sending an email with the standard emailclient!"
  .MessagesMsg21 = "User passwords do not match!"
  .MessagesMsg22 = "Owner passwords do not match!"
  .MessagesMsg23 = "The document is not protected!"
  .MessagesMsg24 = "The user password is empty! Continue?"
  .MessagesMsg25 = "The owner password is empty! Continue?"
  .MessagesMsg26 = "Unknown error"
  .MessagesMsg27 = "Cannot find the file/page."
  .MessagesMsg28 = "The filesize is 0 byte."
  .MessagesMsg29 = "Server not found."
  .MessagesMsg30 = "The url isn not interpretable."
  .MessagesMsg31 = "An error has occured"
  .MessagesMsg32 = "The new version %1 is available. Would you like download the new version from the Sourceforge pages?"
  .MessagesMsg33 = "You already have the most recent version."
  .MessagesMsg34 = "The file is in use. Please close the file first or choose another filename."
  .MessagesMsg35 = "It is necessary to temporarily set PDFCreator as defaultprinter."
  .MessagesMsg36 = "Don't ask me again."

  .OptionsAssociatePSFiles = "Associate PDFCreator with postscript files"
  .OptionsAutosaveDirectoryPrompt = "Select Autosave Directory"
  .OptionsAutosaveFilename = "Filename"
  .OptionsAutosaveFilenameTokens = "Add a Filename-Token"
  .OptionsAutosaveFormat = "Autosave format"
  .OptionsAutosaveStartStandardProgram = "After auto-saving open the document with the default program."
  .OptionsBitmapResolution = "Resolution"
  .OptionsBMPColorscount01 = "4294967296 colors (32 Bit)"
  .OptionsBMPColorscount02 = "16777216 colors (24 Bit)"
  .OptionsBMPColorscount03 = "256 colors (8 Bit)"
  .OptionsBMPColorscount04 = "16 colors (4 Bit)"
  .OptionsBMPColorscount05 = "8 colors (3 Bit)"
  .OptionsBMPColorscount06 = "2 colors (Black/White)"
  .OptionsBMPColorscount07 = "Greyscale (8 Bit)"
  .OptionsBMPDescription = "Windows Bitmap Format. Please use only for single pages."
  .OptionsBMPSymbol = "BMP"
  .OptionsCancel = "&Cancel"
  .OptionsDirectoriesGSBin = "Ghostscript Binaries"
  .OptionsDirectoriesGSFonts = "Ghostscript Fonts"
  .OptionsDirectoriesGSLibraries = "Ghostscript Libraries"
  .OptionsDirectoriesTempPath = "Temporary Files"
  .OptionsDocument = "Document"
  .OptionsEPSDescription = "Encapsulated Postscript Format"
  .OptionsEPSFiles = "Encapsulated Postscript-Files"
  .OptionsEPSSymbol = "EPS"
  .OptionsGhostscriptBinariesDirectoryPrompt = "Select Ghostscript Binaries Directory"
  .OptionsGhostscriptFontsDirectoryPrompt = "Select Ghostscript Fonts Directory"
  .OptionsGhostscriptInternal = "Internal Ghostscript: %1 Ghostscript %2"
  .OptionsGhostscriptLibrariesDirectoryPrompt = "Select Ghostscript Libraries Directory"
  .OptionsGhostscriptResourceDirectoryPrompt = "Select Ghostscript Resource Directory"
  .OptionsGhostscriptversion = "Ghostscript Version"
  .OptionsImageSettings = "Settings"
  .OptionsJavaPath = "Path to Java Interpreter"
  .OptionsJPEGColorscount01 = "16777216 colors (24 Bit)"
  .OptionsJPEGColorscount02 = "Greyscale (8 Bit)"
  .OptionsJPEGDescription = "JPEG (JFIF) Format. Please use only for single pages."
  .OptionsJPEGQuality = "Quality:"
  .OptionsJPEGSymbol = "JPEG"
  .OptionsOwnerPass = "Owner Password"
  .OptionsPassCancel = "Cancel"
  .OptionsPassOK = "OK"
  .OptionsPCXColorscount01 = "4294967296 colors (32 Bit) CMYK"
  .OptionsPCXColorscount02 = "16777216 colors (24 Bit)"
  .OptionsPCXColorscount03 = "256 colors (8 Bit)"
  .OptionsPCXColorscount04 = "16 colors (4 Bit)"
  .OptionsPCXColorscount05 = "2 colors (Black/White)"
  .OptionsPCXColorscount06 = "Greyscale (8 Bit)"
  .OptionsPCXDescription = "PCX Format. Please use only for single pages."
  .OptionsPCXSymbol = "PCX"
  .OptionsPDFAllowAssembly = "Allow changes to the assembly"
  .OptionsPDFAllowDegradedPrinting = "Allow printing in low resolution"
  .OptionsPDFAllowFillIn = "Allow filling in form fields"
  .OptionsPDFAllowScreenReaders = "Allow screen readers"
  .OptionsPDFColors = "Colors"
  .OptionsPDFColorsCaption = "Color Options"
  .OptionsPDFColorsCMYKtoRGB = "Convert CMYK images to RGB"
  .OptionsPDFColorsColorModel01 = "Use Color Model Device RGB"
  .OptionsPDFColorsColorModel02 = "Use Color Model Device CMYK"
  .OptionsPDFColorsColorModel03 = "Use Color Model Device Grayscale"
  .OptionsPDFColorsColorOptions = "Options"
  .OptionsPDFColorsPreserveHalftone = "Preserve Halftone Information"
  .OptionsPDFColorsPreserveOverprint = "Preserve Overprint Settings"
  .OptionsPDFColorsPreserveTransfer = "Preserve Transfer Functions"
  .OptionsPDFCompression = "Compression"
  .OptionsPDFCompressionCaption = "PDF Compression"
  .OptionsPDFCompressionColor = "Color Images"
  .OptionsPDFCompressionColorComp = "Compress"
  .OptionsPDFCompressionColorComp01 = "Automatic"
  .OptionsPDFCompressionColorComp02 = "JPEG-Maximum"
  .OptionsPDFCompressionColorComp03 = "JPEG-High"
  .OptionsPDFCompressionColorComp04 = "JPEG-Medium"
  .OptionsPDFCompressionColorComp05 = "JPEG-Low"
  .OptionsPDFCompressionColorComp06 = "JPEG-Minimum"
  .OptionsPDFCompressionColorComp07 = "ZIP"
  .OptionsPDFCompressionColorComp08 = "LZW-Compression"
  .OptionsPDFCompressionColorRes = "Resolution"
  .OptionsPDFCompressionColorResample = "Resample"
  .OptionsPDFCompressionColorResample01 = "Downsample"
  .OptionsPDFCompressionColorResample02 = "Average Downsample"
  .OptionsPDFCompressionColorResample03 = "Bicubic"
  .OptionsPDFCompressionGrey = "Greyscale Images"
  .OptionsPDFCompressionGreyComp = "Compress"
  .OptionsPDFCompressionGreyComp01 = "Automatic"
  .OptionsPDFCompressionGreyComp02 = "JPEG-Maximum"
  .OptionsPDFCompressionGreyComp03 = "JPEG-High"
  .OptionsPDFCompressionGreyComp04 = "JPEG-Medium"
  .OptionsPDFCompressionGreyComp05 = "JPEG-Low"
  .OptionsPDFCompressionGreyComp06 = "JPEG-Minimum"
  .OptionsPDFCompressionGreyComp07 = "ZIP"
  .OptionsPDFCompressionGreyComp08 = "LZW-Compression"
  .OptionsPDFCompressionGreyRes = "Resolution"
  .OptionsPDFCompressionGreyResample = "Resample"
  .OptionsPDFCompressionGreyResample01 = "Downsample"
  .OptionsPDFCompressionGreyResample02 = "Average Downsample"
  .OptionsPDFCompressionGreyResample03 = "Bicubic"
  .OptionsPDFCompressionMono = "Monochrome Images"
  .OptionsPDFCompressionMonoComp = "Compress"
  .OptionsPDFCompressionMonoComp01 = "CCITT Fax Compression"
  .OptionsPDFCompressionMonoComp02 = "ZIP"
  .OptionsPDFCompressionMonoComp03 = "Run-Length-Encoding"
  .OptionsPDFCompressionMonoComp04 = "LZW-Compression"
  .OptionsPDFCompressionMonoRes = "Resolution"
  .OptionsPDFCompressionMonoResample = "Resample"
  .OptionsPDFCompressionMonoResample01 = "Downsample"
  .OptionsPDFCompressionMonoResample02 = "Average Downsample"
  .OptionsPDFCompressionMonoResample03 = "Bicubic"
  .OptionsPDFCompressionTextComp = "Compress Text Objects"
  .OptionsPDFDescription = "Adobe PDF Format"
  .OptionsPDFDisallowCopy = "Copy text and images"
  .OptionsPDFDisallowModify = "Modify the document"
  .OptionsPDFDisallowModifyComments = "Modify comments"
  .OptionsPDFDisallowPrint = "Print the document"
  .OptionsPDFDisallowUser = "Disallow User to"
  .OptionsPDFEncryptionHigh = "High (128 Bit - Adobe Acrobat 5.0 and above)"
  .OptionsPDFEncryptionLevel = "Encryption Level"
  .OptionsPDFEncryptionLow = "Low (40 Bit - Adobe Acrobat 3.0 and above)"
  .OptionsPDFEncryptor = "Encryptor"
  .OptionsPDFEnhancedPermissions = "Enhanced Permissions (128 Bit only)"
  .OptionsPDFEnterPasswords = "Enter Passwords"
  .OptionsPDFFonts = "Fonts"
  .OptionsPDFFontsCaption = "Font Options"
  .OptionsPDFFontsEmbedAll = "Embed all fonts"
  .OptionsPDFFontsSubSetFonts = "Subset fonts when percentage of used characters below:"
  .OptionsPDFGeneral = "General"
  .OptionsPDFGeneralASCII85 = "Convert binary data to ASCII85"
  .OptionsPDFGeneralAutorotate = "Auto-Rotate Pages:"
  .OptionsPDFGeneralCaption = "General Options"
  .OptionsPDFGeneralCompatibility = "Compatibility:"
  .OptionsPDFGeneralCompatibility01 = "Adobe Acrobat 3.0 (PDF 1.2)"
  .OptionsPDFGeneralCompatibility02 = "Adobe Acrobat 4.0 (PDF 1.3)"
  .OptionsPDFGeneralCompatibility03 = "Adobe Acrobat 5.0 (PDF 1.4)"
  .OptionsPDFGeneralOverprint = "Overprint:"
  .OptionsPDFGeneralOverprint01 = "Non-Zero Overprint"
  .OptionsPDFGeneralOverprint02 = "Full Overprint"
  .OptionsPDFGeneralResolution = "Resolution:"
  .OptionsPDFGeneralRotate01 = "None"
  .OptionsPDFGeneralRotate02 = "All"
  .OptionsPDFGeneralRotate03 = "Single Page"
  .OptionsPDFOptions = "PDF Options"
  .OptionsPDFOwnerPass = "Password required to change permissions and passwords"
  .OptionsPDFPasswords = "Passwords"
  .OptionsPDFRepeatPassword = "Repeat"
  .OptionsPDFSecurity = "Security"
  .OptionsPDFSecurityCaption = "Security"
  .OptionsPDFSetPassword = "Password"
  .OptionsPDFSymbol = "PDF"
  .OptionsPDFUserPass = "Password required to open document"
  .OptionsPDFUseSecurity = "Use Security"
  .OptionsPNGColorscount01 = "16777216 colors (24 Bit)"
  .OptionsPNGColorscount02 = "256 colors (8 Bit)"
  .OptionsPNGColorscount03 = "16 colors (4 Bit)"
  .OptionsPNGColorscount04 = "2 colors (2 Bit - Black/White)"
  .OptionsPNGColorscount05 = "Greyscale (8 Bit)"
  .OptionsPNGDescription = "PNG Format. Please use only for single pages."
  .OptionsPNGFiles = "Bitmap PNG-Files"
  .OptionsPNGSymbol = "PNG"
  .OptionsPrintertempDirectoryPrompt = "Select Printer Temp-Directory"
  .OptionsPrintTestpage = "Print Test Page"
  .OptionsProcesspriority = "Process priority"
  .OptionsProcesspriorityHigh = "High"
  .OptionsProcesspriorityIdle = "Idle"
  .OptionsProcesspriorityNormal = "Normal"
  .OptionsProcesspriorityRealtime = "Realtime"
  .OptionsProgramAutosaveDescription = "Auto-save mode. Auto-save does not prompt for a filename and file location. It automatically saves all PDF files to a single directory with a predefined filename."
  .OptionsProgramAutosaveSymbol = "Auto-save"
  .OptionsProgramDirectoriesDescription = "Directories for Ghostscript, temporary files and others."
  .OptionsProgramDirectoriesSymbol = "Directories"
  .OptionsProgramDocumentDescription = "Document properties"
  .OptionsProgramDocumentSymbol = "Document"
  .OptionsProgramFont = "Program Font"
  .OptionsProgramFontCancelTest = "Cancel Test"
  .OptionsProgramFontcharset = "Character Set"
  .OptionsProgramFontDescription = "Font for labels, captions and values. For the program menu use the general settings in your Windows OS."
  .OptionsProgramFontSize = "Size"
  .OptionsProgramFontSymbol = "Program font"
  .OptionsProgramFontTest = "Test"
  .OptionsProgramFontTestdescription = "Here you can test the font."
  .OptionsProgramGeneralDescription = "General Settings."
  .OptionsProgramGeneralSymbol = "General settings"
  .OptionsProgramGhostscriptDescription = "Ghostscript"
  .OptionsProgramGhostscriptSymbol = "Ghostscript"
  .OptionsProgramSaveDescription = "Save"
  .OptionsProgramSaveSymbol = "Save"
  .OptionsProgramSwitchingDefaultprinter = "No confirm message switching PDFCreator temporarily as default printer."
  .OptionsPSDescription = "Postscript Format"
  .OptionsPSFiles = "Postscript-Files"
  .OptionsPSLanguageLevel = "Language Level:"
  .OptionsPSSymbol = "PS"
  .OptionsRemoveSpaces = "Remove leading and trailing spaces"
  .OptionsReset = "&Reset all settings"
  .OptionsSave = "&Save"
  .OptionsSaveFilename = "Filename"
  .OptionsSaveFilenameAdd = "Add"
  .OptionsSaveFilenameChange = "Change"
  .OptionsSaveFilenameDelete = "Delete"
  .OptionsSaveFilenameSubstitutions = "Filename substitution"
  .OptionsSaveFilenameSubstitutionsTitle = "Filename substitution only in <Title>"
  .OptionsSaveFilenameTokens = "Add a Filename-Token"
  .OptionsSavePasswords = "Save passwords temporarily for this session."
  .OptionsShellIntegration = "Shell integration"
  .OptionsShellIntegrationAdd = "Integrate PDFCreator into shell"
  .OptionsShellIntegrationCaption = "Create &PDF with PDFCreator"
  .OptionsShellIntegrationRemove = "Remove shell integration"
  .OptionsStandardAuthorToken = "Add a Author-Token"
  .OptionsTestpage = "PDFCreator Testpage"
  .OptionsTIFFColorscount01 = "16777216 (24 Bit)"
  .OptionsTIFFColorscount02 = "4096 (12 Bit)"
  .OptionsTIFFColorscount03 = "2 colors (Black/White) G3 fax encoding with no EOLs"
  .OptionsTIFFColorscount04 = "2 colors (Black/White) G3 fax encoding with EOLs"
  .OptionsTIFFColorscount05 = "2 colors (Black/White) 2-D G3 fax encoding"
  .OptionsTIFFColorscount06 = "2 colors (Black/White) G4 fax encoding"
  .OptionsTIFFColorscount07 = "2 colors (Black/White) LZW-compatible"
  .OptionsTIFFColorscount08 = "2 colors (Black/White) PackBits"
  .OptionsTIFFDescription = "TIFF Format. For multipages use the tiff-format."
  .OptionsTIFFSymbol = "TIFF"
  .OptionsTreeFormats = "Formats"
  .OptionsTreeProgram = "Program"
  .OptionsUseAutosave = "Use Auto-save"
  .OptionsUseAutosaveDirectory = "Use this directory for auto-save"
  .OptionsUseCreationDateNow = "Use the current Date/Time for 'Creation Date'"
  .OptionsUserPass = "User Password"
  .OptionsUseStandardauthor = "Use standard author"

  .PrintingAuthor = "A&uthor:"
  .PrintingBMPFiles = "BMP-Files"
  .PrintingCancel = "&Cancel"
  .PrintingCollect = "&Wait - Collect"
  .PrintingCreationDate = "Creation &Date:"
  .PrintingDocumentTitle = "Document &Title:"
  .PrintingEMail = "&eMail"
  .PrintingEPSFiles = "Encapsulated Postscript-Files"
  .PrintingJPEGFiles = "JPEG-Files"
  .PrintingKeywords = "&Keywords:"
  .PrintingModifyDate = "&Modify Date:"
  .PrintingNow = "Now"
  .PrintingPCXFiles = "PCX-Files"
  .PrintingPDFFiles = "PDF-Files"
  .PrintingPNGFiles = "PNG-Files"
  .PrintingPSFiles = "Postscript-Files"
  .PrintingSave = "&Save"
  .PrintingStartStandardProgram = "&After saving open the document with the default program."
  .PrintingStatus = "Creating file..."
  .PrintingSubject = "Su&bject:"
  .PrintingTIFFFiles = "TIFF-Files"

  .SaveOpenAttributes = "Attributes"
  .SaveOpenCancel = "Cancel"
  .SaveOpenFilename = "Filename"
  .SaveOpenOpen = "Open"
  .SaveOpenOpenTitle = "Open"
  .SaveOpenSave = "Save"
  .SaveOpenSaveTitle = "Save as"
  .SaveOpenSize = "Size"

 End With
End Sub

