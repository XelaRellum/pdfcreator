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
 DialogDocumentDelete As String
 DialogDocumentDown As String
 DialogDocumentPrint As String
 DialogDocumentSave As String
 DialogDocumentTop As String
 DialogDocumentUp As String
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
 PrintingWaiting As String

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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  InitLanguagesStrings
50020  LoadCommonStrings Languagefile
50030  LoadDialogStrings Languagefile
50040  LoadListStrings Languagefile
50050  LoadLoggingStrings Languagefile
50060  LoadMessagesStrings Languagefile
50070  LoadOptionsStrings Languagefile
50080  LoadPrintingStrings Languagefile
50090  LoadSaveOpenStrings Languagefile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLanguage", "LoadLanguage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub LoadCommonStrings(ByVal Languagefile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hLang As New clsHash
50020  ReadINISection Languagefile, "Common", hLang
50030  With LanguageStrings
50040   .CommonAuthor = Replace$(hLang.Retrieve("Author", .CommonAuthor), "/n", vbCrLf)
50050   .CommonLanguagename = Replace$(hLang.Retrieve("Languagename", .CommonLanguagename), "/n", vbCrLf)
50060   .CommonTitle = Replace$(hLang.Retrieve("Title", .CommonTitle), "/n", vbCrLf)
50070   .CommonVersion = Replace$(hLang.Retrieve("Version", .CommonVersion), "/n", vbCrLf)
50080  End With
50090  Set hLang = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLanguage", "LoadCommonStrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub LoadDialogStrings(ByVal Languagefile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hLang As New clsHash
50020  ReadINISection Languagefile, "Dialog", hLang
50030  With LanguageStrings
50040   .DialogDocument = Replace$(hLang.Retrieve("Document", .DialogDocument), "/n", vbCrLf)
50050   .DialogDocumentAdd = Replace$(hLang.Retrieve("DocumentAdd", .DialogDocumentAdd), "/n", vbCrLf)
50060   .DialogDocumentBottom = Replace$(hLang.Retrieve("DocumentBottom", .DialogDocumentBottom), "/n", vbCrLf)
50070   .DialogDocumentCombine = Replace$(hLang.Retrieve("DocumentCombine", .DialogDocumentCombine), "/n", vbCrLf)
50080   .DialogDocumentDelete = Replace$(hLang.Retrieve("DocumentDelete", .DialogDocumentDelete), "/n", vbCrLf)
50090   .DialogDocumentDown = Replace$(hLang.Retrieve("DocumentDown", .DialogDocumentDown), "/n", vbCrLf)
50100   .DialogDocumentPrint = Replace$(hLang.Retrieve("DocumentPrint", .DialogDocumentPrint), "/n", vbCrLf)
50110   .DialogDocumentSave = Replace$(hLang.Retrieve("DocumentSave", .DialogDocumentSave), "/n", vbCrLf)
50120   .DialogDocumentTop = Replace$(hLang.Retrieve("DocumentTop", .DialogDocumentTop), "/n", vbCrLf)
50130   .DialogDocumentUp = Replace$(hLang.Retrieve("DocumentUp", .DialogDocumentUp), "/n", vbCrLf)
50140   .DialogInfo = Replace$(hLang.Retrieve("Info", .DialogInfo), "/n", vbCrLf)
50150   .DialogInfoCheckUpdates = Replace$(hLang.Retrieve("InfoCheckUpdates", .DialogInfoCheckUpdates), "/n", vbCrLf)
50160   .DialogInfoHomepage = Replace$(hLang.Retrieve("InfoHomepage", .DialogInfoHomepage), "/n", vbCrLf)
50170   .DialogInfoInfo = Replace$(hLang.Retrieve("InfoInfo", .DialogInfoInfo), "/n", vbCrLf)
50180   .DialogInfoPaypal = Replace$(hLang.Retrieve("InfoPaypal", .DialogInfoPaypal), "/n", vbCrLf)
50190   .DialogInfoPDFCreatorSourceforge = Replace$(hLang.Retrieve("InfoPDFCreatorSourceforge", .DialogInfoPDFCreatorSourceforge), "/n", vbCrLf)
50200   .DialogLanguage = Replace$(hLang.Retrieve("Language", .DialogLanguage), "/n", vbCrLf)
50210   .DialogPrinter = Replace$(hLang.Retrieve("Printer", .DialogPrinter), "/n", vbCrLf)
50220   .DialogPrinterClose = Replace$(hLang.Retrieve("PrinterClose", .DialogPrinterClose), "/n", vbCrLf)
50230   .DialogPrinterLogfile = Replace$(hLang.Retrieve("PrinterLogfile", .DialogPrinterLogfile), "/n", vbCrLf)
50240   .DialogPrinterLogfiles = Replace$(hLang.Retrieve("PrinterLogfiles", .DialogPrinterLogfiles), "/n", vbCrLf)
50250   .DialogPrinterLogging = Replace$(hLang.Retrieve("PrinterLogging", .DialogPrinterLogging), "/n", vbCrLf)
50260   .DialogPrinterOptions = Replace$(hLang.Retrieve("PrinterOptions", .DialogPrinterOptions), "/n", vbCrLf)
50270   .DialogPrinterPrinterStop = Replace$(hLang.Retrieve("PrinterPrinterStop", .DialogPrinterPrinterStop), "/n", vbCrLf)
50280   .DialogView = Replace$(hLang.Retrieve("View", .DialogView), "/n", vbCrLf)
50290   .DialogViewStatusbar = Replace$(hLang.Retrieve("ViewStatusbar", .DialogViewStatusbar), "/n", vbCrLf)
50300  End With
50310  Set hLang = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLanguage", "LoadDialogStrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub LoadListStrings(ByVal Languagefile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hLang As New clsHash
50020  ReadINISection Languagefile, "List", hLang
50030  With LanguageStrings
50040   .ListAddFile = Replace$(hLang.Retrieve("AddFile", .ListAddFile), "/n", vbCrLf)
50050   .ListAllFiles = Replace$(hLang.Retrieve("AllFiles", .ListAllFiles), "/n", vbCrLf)
50060   .ListBytes = Replace$(hLang.Retrieve("Bytes", .ListBytes), "/n", vbCrLf)
50070   .ListDate = Replace$(hLang.Retrieve("Date", .ListDate), "/n", vbCrLf)
50080   .ListDocumenttitle = Replace$(hLang.Retrieve("Documenttitle", .ListDocumenttitle), "/n", vbCrLf)
50090   .ListFilename = Replace$(hLang.Retrieve("Filename", .ListFilename), "/n", vbCrLf)
50100   .ListGBytes = Replace$(hLang.Retrieve("GBytes", .ListGBytes), "/n", vbCrLf)
50110   .ListKBytes = Replace$(hLang.Retrieve("KBytes", .ListKBytes), "/n", vbCrLf)
50120   .ListMBytes = Replace$(hLang.Retrieve("MBytes", .ListMBytes), "/n", vbCrLf)
50130   .ListPDFFiles = Replace$(hLang.Retrieve("PDFFiles", .ListPDFFiles), "/n", vbCrLf)
50140   .ListPostscriptFiles = Replace$(hLang.Retrieve("PostscriptFiles", .ListPostscriptFiles), "/n", vbCrLf)
50150   .ListPrinting = Replace$(hLang.Retrieve("Printing", .ListPrinting), "/n", vbCrLf)
50160   .ListSize = Replace$(hLang.Retrieve("Size", .ListSize), "/n", vbCrLf)
50170   .ListStatus = Replace$(hLang.Retrieve("Status", .ListStatus), "/n", vbCrLf)
50180   .ListWaiting = Replace$(hLang.Retrieve("Waiting", .ListWaiting), "/n", vbCrLf)
50190  End With
50200  Set hLang = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLanguage", "LoadListStrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub LoadLoggingStrings(ByVal Languagefile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hLang As New clsHash
50020  ReadINISection Languagefile, "Logging", hLang
50030  With LanguageStrings
50040   .LoggingClear = Replace$(hLang.Retrieve("Clear", .LoggingClear), "/n", vbCrLf)
50050   .LoggingClose = Replace$(hLang.Retrieve("Close", .LoggingClose), "/n", vbCrLf)
50060   .LoggingLogfile = Replace$(hLang.Retrieve("Logfile", .LoggingLogfile), "/n", vbCrLf)
50070  End With
50080  Set hLang = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLanguage", "LoadLoggingStrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub LoadMessagesStrings(ByVal Languagefile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hLang As New clsHash
50020  ReadINISection Languagefile, "Messages", hLang
50030  With LanguageStrings
50040   .MessagesMsg01 = Replace$(hLang.Retrieve("Msg01", .MessagesMsg01), "/n", vbCrLf)
50050   .MessagesMsg02 = Replace$(hLang.Retrieve("Msg02", .MessagesMsg02), "/n", vbCrLf)
50060   .MessagesMsg03 = Replace$(hLang.Retrieve("Msg03", .MessagesMsg03), "/n", vbCrLf)
50070   .MessagesMsg04 = Replace$(hLang.Retrieve("Msg04", .MessagesMsg04), "/n", vbCrLf)
50080   .MessagesMsg05 = Replace$(hLang.Retrieve("Msg05", .MessagesMsg05), "/n", vbCrLf)
50090   .MessagesMsg06 = Replace$(hLang.Retrieve("Msg06", .MessagesMsg06), "/n", vbCrLf)
50100   .MessagesMsg07 = Replace$(hLang.Retrieve("Msg07", .MessagesMsg07), "/n", vbCrLf)
50110   .MessagesMsg08 = Replace$(hLang.Retrieve("Msg08", .MessagesMsg08), "/n", vbCrLf)
50120   .MessagesMsg09 = Replace$(hLang.Retrieve("Msg09", .MessagesMsg09), "/n", vbCrLf)
50130   .MessagesMsg10 = Replace$(hLang.Retrieve("Msg10", .MessagesMsg10), "/n", vbCrLf)
50140   .MessagesMsg11 = Replace$(hLang.Retrieve("Msg11", .MessagesMsg11), "/n", vbCrLf)
50150   .MessagesMsg12 = Replace$(hLang.Retrieve("Msg12", .MessagesMsg12), "/n", vbCrLf)
50160   .MessagesMsg13 = Replace$(hLang.Retrieve("Msg13", .MessagesMsg13), "/n", vbCrLf)
50170   .MessagesMsg14 = Replace$(hLang.Retrieve("Msg14", .MessagesMsg14), "/n", vbCrLf)
50180   .MessagesMsg15 = Replace$(hLang.Retrieve("Msg15", .MessagesMsg15), "/n", vbCrLf)
50190   .MessagesMsg16 = Replace$(hLang.Retrieve("Msg16", .MessagesMsg16), "/n", vbCrLf)
50200   .MessagesMsg17 = Replace$(hLang.Retrieve("Msg17", .MessagesMsg17), "/n", vbCrLf)
50210   .MessagesMsg19 = Replace$(hLang.Retrieve("Msg19", .MessagesMsg19), "/n", vbCrLf)
50220   .MessagesMsg20 = Replace$(hLang.Retrieve("Msg20", .MessagesMsg20), "/n", vbCrLf)
50230   .MessagesMsg21 = Replace$(hLang.Retrieve("Msg21", .MessagesMsg21), "/n", vbCrLf)
50240   .MessagesMsg22 = Replace$(hLang.Retrieve("Msg22", .MessagesMsg22), "/n", vbCrLf)
50250   .MessagesMsg23 = Replace$(hLang.Retrieve("Msg23", .MessagesMsg23), "/n", vbCrLf)
50260   .MessagesMsg24 = Replace$(hLang.Retrieve("Msg24", .MessagesMsg24), "/n", vbCrLf)
50270   .MessagesMsg25 = Replace$(hLang.Retrieve("Msg25", .MessagesMsg25), "/n", vbCrLf)
50280   .MessagesMsg26 = Replace$(hLang.Retrieve("Msg26", .MessagesMsg26), "/n", vbCrLf)
50290   .MessagesMsg27 = Replace$(hLang.Retrieve("Msg27", .MessagesMsg27), "/n", vbCrLf)
50300   .MessagesMsg28 = Replace$(hLang.Retrieve("Msg28", .MessagesMsg28), "/n", vbCrLf)
50310   .MessagesMsg29 = Replace$(hLang.Retrieve("Msg29", .MessagesMsg29), "/n", vbCrLf)
50320   .MessagesMsg30 = Replace$(hLang.Retrieve("Msg30", .MessagesMsg30), "/n", vbCrLf)
50330   .MessagesMsg31 = Replace$(hLang.Retrieve("Msg31", .MessagesMsg31), "/n", vbCrLf)
50340   .MessagesMsg32 = Replace$(hLang.Retrieve("Msg32", .MessagesMsg32), "/n", vbCrLf)
50350   .MessagesMsg33 = Replace$(hLang.Retrieve("Msg33", .MessagesMsg33), "/n", vbCrLf)
50360   .MessagesMsg34 = Replace$(hLang.Retrieve("Msg34", .MessagesMsg34), "/n", vbCrLf)
50370   .MessagesMsg35 = Replace$(hLang.Retrieve("Msg35", .MessagesMsg35), "/n", vbCrLf)
50380   .MessagesMsg36 = Replace$(hLang.Retrieve("Msg36", .MessagesMsg36), "/n", vbCrLf)
50390  End With
50400  Set hLang = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLanguage", "LoadMessagesStrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub LoadOptionsStrings(ByVal Languagefile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hLang As New clsHash
50020  ReadINISection Languagefile, "Options", hLang
50030  With LanguageStrings
50040   .OptionsAssociatePSFiles = Replace$(hLang.Retrieve("AssociatePSFiles", .OptionsAssociatePSFiles), "/n", vbCrLf)
50050   .OptionsAutosaveDirectoryPrompt = Replace$(hLang.Retrieve("AutosaveDirectoryPrompt", .OptionsAutosaveDirectoryPrompt), "/n", vbCrLf)
50060   .OptionsAutosaveFilename = Replace$(hLang.Retrieve("AutosaveFilename", .OptionsAutosaveFilename), "/n", vbCrLf)
50070   .OptionsAutosaveFilenameTokens = Replace$(hLang.Retrieve("AutosaveFilenameTokens", .OptionsAutosaveFilenameTokens), "/n", vbCrLf)
50080   .OptionsAutosaveFormat = Replace$(hLang.Retrieve("AutosaveFormat", .OptionsAutosaveFormat), "/n", vbCrLf)
50090   .OptionsBitmapResolution = Replace$(hLang.Retrieve("BitmapResolution", .OptionsBitmapResolution), "/n", vbCrLf)
50100   .OptionsBMPColorscount01 = Replace$(hLang.Retrieve("BMPColorscount01", .OptionsBMPColorscount01), "/n", vbCrLf)
50110   .OptionsBMPColorscount02 = Replace$(hLang.Retrieve("BMPColorscount02", .OptionsBMPColorscount02), "/n", vbCrLf)
50120   .OptionsBMPColorscount03 = Replace$(hLang.Retrieve("BMPColorscount03", .OptionsBMPColorscount03), "/n", vbCrLf)
50130   .OptionsBMPColorscount04 = Replace$(hLang.Retrieve("BMPColorscount04", .OptionsBMPColorscount04), "/n", vbCrLf)
50140   .OptionsBMPColorscount05 = Replace$(hLang.Retrieve("BMPColorscount05", .OptionsBMPColorscount05), "/n", vbCrLf)
50150   .OptionsBMPColorscount06 = Replace$(hLang.Retrieve("BMPColorscount06", .OptionsBMPColorscount06), "/n", vbCrLf)
50160   .OptionsBMPColorscount07 = Replace$(hLang.Retrieve("BMPColorscount07", .OptionsBMPColorscount07), "/n", vbCrLf)
50170   .OptionsBMPDescription = Replace$(hLang.Retrieve("BMPDescription", .OptionsBMPDescription), "/n", vbCrLf)
50180   .OptionsBMPSymbol = Replace$(hLang.Retrieve("BMPSymbol", .OptionsBMPSymbol), "/n", vbCrLf)
50190   .OptionsCancel = Replace$(hLang.Retrieve("Cancel", .OptionsCancel), "/n", vbCrLf)
50200   .OptionsDirectoriesGSBin = Replace$(hLang.Retrieve("DirectoriesGSBin", .OptionsDirectoriesGSBin), "/n", vbCrLf)
50210   .OptionsDirectoriesGSFonts = Replace$(hLang.Retrieve("DirectoriesGSFonts", .OptionsDirectoriesGSFonts), "/n", vbCrLf)
50220   .OptionsDirectoriesGSLibraries = Replace$(hLang.Retrieve("DirectoriesGSLibraries", .OptionsDirectoriesGSLibraries), "/n", vbCrLf)
50230   .OptionsDirectoriesTempPath = Replace$(hLang.Retrieve("DirectoriesTempPath", .OptionsDirectoriesTempPath), "/n", vbCrLf)
50240   .OptionsDocument = Replace$(hLang.Retrieve("Document", .OptionsDocument), "/n", vbCrLf)
50250   .OptionsEPSDescription = Replace$(hLang.Retrieve("EPSDescription", .OptionsEPSDescription), "/n", vbCrLf)
50260   .OptionsEPSFiles = Replace$(hLang.Retrieve("EPSFiles", .OptionsEPSFiles), "/n", vbCrLf)
50270   .OptionsEPSSymbol = Replace$(hLang.Retrieve("EPSSymbol", .OptionsEPSSymbol), "/n", vbCrLf)
50280   .OptionsGhostscriptBinariesDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptBinariesDirectoryPrompt", .OptionsGhostscriptBinariesDirectoryPrompt), "/n", vbCrLf)
50290   .OptionsGhostscriptFontsDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptFontsDirectoryPrompt", .OptionsGhostscriptFontsDirectoryPrompt), "/n", vbCrLf)
50300   .OptionsGhostscriptInternal = Replace$(hLang.Retrieve("GhostscriptInternal", .OptionsGhostscriptInternal), "/n", vbCrLf)
50310   .OptionsGhostscriptLibrariesDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptLibrariesDirectoryPrompt", .OptionsGhostscriptLibrariesDirectoryPrompt), "/n", vbCrLf)
50320   .OptionsGhostscriptResourceDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptResourceDirectoryPrompt", .OptionsGhostscriptResourceDirectoryPrompt), "/n", vbCrLf)
50330   .OptionsGhostscriptversion = Replace$(hLang.Retrieve("Ghostscriptversion", .OptionsGhostscriptversion), "/n", vbCrLf)
50340   .OptionsImageSettings = Replace$(hLang.Retrieve("ImageSettings", .OptionsImageSettings), "/n", vbCrLf)
50350   .OptionsJavaPath = Replace$(hLang.Retrieve("JavaPath", .OptionsJavaPath), "/n", vbCrLf)
50360   .OptionsJPEGColorscount01 = Replace$(hLang.Retrieve("JPEGColorscount01", .OptionsJPEGColorscount01), "/n", vbCrLf)
50370   .OptionsJPEGColorscount02 = Replace$(hLang.Retrieve("JPEGColorscount02", .OptionsJPEGColorscount02), "/n", vbCrLf)
50380   .OptionsJPEGDescription = Replace$(hLang.Retrieve("JPEGDescription", .OptionsJPEGDescription), "/n", vbCrLf)
50390   .OptionsJPEGQuality = Replace$(hLang.Retrieve("JPEGQuality", .OptionsJPEGQuality), "/n", vbCrLf)
50400   .OptionsJPEGSymbol = Replace$(hLang.Retrieve("JPEGSymbol", .OptionsJPEGSymbol), "/n", vbCrLf)
50410   .OptionsOwnerPass = Replace$(hLang.Retrieve("OwnerPass", .OptionsOwnerPass), "/n", vbCrLf)
50420   .OptionsPassCancel = Replace$(hLang.Retrieve("PassCancel", .OptionsPassCancel), "/n", vbCrLf)
50430   .OptionsPassOK = Replace$(hLang.Retrieve("PassOK", .OptionsPassOK), "/n", vbCrLf)
50440   .OptionsPCXColorscount01 = Replace$(hLang.Retrieve("PCXColorscount01", .OptionsPCXColorscount01), "/n", vbCrLf)
50450   .OptionsPCXColorscount02 = Replace$(hLang.Retrieve("PCXColorscount02", .OptionsPCXColorscount02), "/n", vbCrLf)
50460   .OptionsPCXColorscount03 = Replace$(hLang.Retrieve("PCXColorscount03", .OptionsPCXColorscount03), "/n", vbCrLf)
50470   .OptionsPCXColorscount04 = Replace$(hLang.Retrieve("PCXColorscount04", .OptionsPCXColorscount04), "/n", vbCrLf)
50480   .OptionsPCXColorscount05 = Replace$(hLang.Retrieve("PCXColorscount05", .OptionsPCXColorscount05), "/n", vbCrLf)
50490   .OptionsPCXColorscount06 = Replace$(hLang.Retrieve("PCXColorscount06", .OptionsPCXColorscount06), "/n", vbCrLf)
50500   .OptionsPCXDescription = Replace$(hLang.Retrieve("PCXDescription", .OptionsPCXDescription), "/n", vbCrLf)
50510   .OptionsPCXSymbol = Replace$(hLang.Retrieve("PCXSymbol", .OptionsPCXSymbol), "/n", vbCrLf)
50520   .OptionsPDFAllowAssembly = Replace$(hLang.Retrieve("PDFAllowAssembly", .OptionsPDFAllowAssembly), "/n", vbCrLf)
50530   .OptionsPDFAllowDegradedPrinting = Replace$(hLang.Retrieve("PDFAllowDegradedPrinting", .OptionsPDFAllowDegradedPrinting), "/n", vbCrLf)
50540   .OptionsPDFAllowFillIn = Replace$(hLang.Retrieve("PDFAllowFillIn", .OptionsPDFAllowFillIn), "/n", vbCrLf)
50550   .OptionsPDFAllowScreenReaders = Replace$(hLang.Retrieve("PDFAllowScreenReaders", .OptionsPDFAllowScreenReaders), "/n", vbCrLf)
50560   .OptionsPDFColors = Replace$(hLang.Retrieve("PDFColors", .OptionsPDFColors), "/n", vbCrLf)
50570   .OptionsPDFColorsCaption = Replace$(hLang.Retrieve("PDFColorsCaption", .OptionsPDFColorsCaption), "/n", vbCrLf)
50580   .OptionsPDFColorsCMYKtoRGB = Replace$(hLang.Retrieve("PDFColorsCMYKtoRGB", .OptionsPDFColorsCMYKtoRGB), "/n", vbCrLf)
50590   .OptionsPDFColorsColorModel01 = Replace$(hLang.Retrieve("PDFColorsColorModel01", .OptionsPDFColorsColorModel01), "/n", vbCrLf)
50600   .OptionsPDFColorsColorModel02 = Replace$(hLang.Retrieve("PDFColorsColorModel02", .OptionsPDFColorsColorModel02), "/n", vbCrLf)
50610   .OptionsPDFColorsColorModel03 = Replace$(hLang.Retrieve("PDFColorsColorModel03", .OptionsPDFColorsColorModel03), "/n", vbCrLf)
50620   .OptionsPDFColorsColorOptions = Replace$(hLang.Retrieve("PDFColorsColorOptions", .OptionsPDFColorsColorOptions), "/n", vbCrLf)
50630   .OptionsPDFColorsPreserveHalftone = Replace$(hLang.Retrieve("PDFColorsPreserveHalftone", .OptionsPDFColorsPreserveHalftone), "/n", vbCrLf)
50640   .OptionsPDFColorsPreserveOverprint = Replace$(hLang.Retrieve("PDFColorsPreserveOverprint", .OptionsPDFColorsPreserveOverprint), "/n", vbCrLf)
50650   .OptionsPDFColorsPreserveTransfer = Replace$(hLang.Retrieve("PDFColorsPreserveTransfer", .OptionsPDFColorsPreserveTransfer), "/n", vbCrLf)
50660   .OptionsPDFCompression = Replace$(hLang.Retrieve("PDFCompression", .OptionsPDFCompression), "/n", vbCrLf)
50670   .OptionsPDFCompressionCaption = Replace$(hLang.Retrieve("PDFCompressionCaption", .OptionsPDFCompressionCaption), "/n", vbCrLf)
50680   .OptionsPDFCompressionColor = Replace$(hLang.Retrieve("PDFCompressionColor", .OptionsPDFCompressionColor), "/n", vbCrLf)
50690   .OptionsPDFCompressionColorComp = Replace$(hLang.Retrieve("PDFCompressionColorComp", .OptionsPDFCompressionColorComp), "/n", vbCrLf)
50700   .OptionsPDFCompressionColorComp01 = Replace$(hLang.Retrieve("PDFCompressionColorComp01", .OptionsPDFCompressionColorComp01), "/n", vbCrLf)
50710   .OptionsPDFCompressionColorComp02 = Replace$(hLang.Retrieve("PDFCompressionColorComp02", .OptionsPDFCompressionColorComp02), "/n", vbCrLf)
50720   .OptionsPDFCompressionColorComp03 = Replace$(hLang.Retrieve("PDFCompressionColorComp03", .OptionsPDFCompressionColorComp03), "/n", vbCrLf)
50730   .OptionsPDFCompressionColorComp04 = Replace$(hLang.Retrieve("PDFCompressionColorComp04", .OptionsPDFCompressionColorComp04), "/n", vbCrLf)
50740   .OptionsPDFCompressionColorComp05 = Replace$(hLang.Retrieve("PDFCompressionColorComp05", .OptionsPDFCompressionColorComp05), "/n", vbCrLf)
50750   .OptionsPDFCompressionColorComp06 = Replace$(hLang.Retrieve("PDFCompressionColorComp06", .OptionsPDFCompressionColorComp06), "/n", vbCrLf)
50760   .OptionsPDFCompressionColorComp07 = Replace$(hLang.Retrieve("PDFCompressionColorComp07", .OptionsPDFCompressionColorComp07), "/n", vbCrLf)
50770   .OptionsPDFCompressionColorComp08 = Replace$(hLang.Retrieve("PDFCompressionColorComp08", .OptionsPDFCompressionColorComp08), "/n", vbCrLf)
50780   .OptionsPDFCompressionColorRes = Replace$(hLang.Retrieve("PDFCompressionColorRes", .OptionsPDFCompressionColorRes), "/n", vbCrLf)
50790   .OptionsPDFCompressionColorResample = Replace$(hLang.Retrieve("PDFCompressionColorResample", .OptionsPDFCompressionColorResample), "/n", vbCrLf)
50800   .OptionsPDFCompressionColorResample01 = Replace$(hLang.Retrieve("PDFCompressionColorResample01", .OptionsPDFCompressionColorResample01), "/n", vbCrLf)
50810   .OptionsPDFCompressionColorResample02 = Replace$(hLang.Retrieve("PDFCompressionColorResample02", .OptionsPDFCompressionColorResample02), "/n", vbCrLf)
50820   .OptionsPDFCompressionColorResample03 = Replace$(hLang.Retrieve("PDFCompressionColorResample03", .OptionsPDFCompressionColorResample03), "/n", vbCrLf)
50830   .OptionsPDFCompressionGrey = Replace$(hLang.Retrieve("PDFCompressionGrey", .OptionsPDFCompressionGrey), "/n", vbCrLf)
50840   .OptionsPDFCompressionGreyComp = Replace$(hLang.Retrieve("PDFCompressionGreyComp", .OptionsPDFCompressionGreyComp), "/n", vbCrLf)
50850   .OptionsPDFCompressionGreyComp01 = Replace$(hLang.Retrieve("PDFCompressionGreyComp01", .OptionsPDFCompressionGreyComp01), "/n", vbCrLf)
50860   .OptionsPDFCompressionGreyComp02 = Replace$(hLang.Retrieve("PDFCompressionGreyComp02", .OptionsPDFCompressionGreyComp02), "/n", vbCrLf)
50870   .OptionsPDFCompressionGreyComp03 = Replace$(hLang.Retrieve("PDFCompressionGreyComp03", .OptionsPDFCompressionGreyComp03), "/n", vbCrLf)
50880   .OptionsPDFCompressionGreyComp04 = Replace$(hLang.Retrieve("PDFCompressionGreyComp04", .OptionsPDFCompressionGreyComp04), "/n", vbCrLf)
50890   .OptionsPDFCompressionGreyComp05 = Replace$(hLang.Retrieve("PDFCompressionGreyComp05", .OptionsPDFCompressionGreyComp05), "/n", vbCrLf)
50900   .OptionsPDFCompressionGreyComp06 = Replace$(hLang.Retrieve("PDFCompressionGreyComp06", .OptionsPDFCompressionGreyComp06), "/n", vbCrLf)
50910   .OptionsPDFCompressionGreyComp07 = Replace$(hLang.Retrieve("PDFCompressionGreyComp07", .OptionsPDFCompressionGreyComp07), "/n", vbCrLf)
50920   .OptionsPDFCompressionGreyComp08 = Replace$(hLang.Retrieve("PDFCompressionGreyComp08", .OptionsPDFCompressionGreyComp08), "/n", vbCrLf)
50930   .OptionsPDFCompressionGreyRes = Replace$(hLang.Retrieve("PDFCompressionGreyRes", .OptionsPDFCompressionGreyRes), "/n", vbCrLf)
50940   .OptionsPDFCompressionGreyResample = Replace$(hLang.Retrieve("PDFCompressionGreyResample", .OptionsPDFCompressionGreyResample), "/n", vbCrLf)
50950   .OptionsPDFCompressionGreyResample01 = Replace$(hLang.Retrieve("PDFCompressionGreyResample01", .OptionsPDFCompressionGreyResample01), "/n", vbCrLf)
50960   .OptionsPDFCompressionGreyResample02 = Replace$(hLang.Retrieve("PDFCompressionGreyResample02", .OptionsPDFCompressionGreyResample02), "/n", vbCrLf)
50970   .OptionsPDFCompressionGreyResample03 = Replace$(hLang.Retrieve("PDFCompressionGreyResample03", .OptionsPDFCompressionGreyResample03), "/n", vbCrLf)
50980   .OptionsPDFCompressionMono = Replace$(hLang.Retrieve("PDFCompressionMono", .OptionsPDFCompressionMono), "/n", vbCrLf)
50990   .OptionsPDFCompressionMonoComp = Replace$(hLang.Retrieve("PDFCompressionMonoComp", .OptionsPDFCompressionMonoComp), "/n", vbCrLf)
51000   .OptionsPDFCompressionMonoComp01 = Replace$(hLang.Retrieve("PDFCompressionMonoComp01", .OptionsPDFCompressionMonoComp01), "/n", vbCrLf)
51010   .OptionsPDFCompressionMonoComp02 = Replace$(hLang.Retrieve("PDFCompressionMonoComp02", .OptionsPDFCompressionMonoComp02), "/n", vbCrLf)
51020   .OptionsPDFCompressionMonoComp03 = Replace$(hLang.Retrieve("PDFCompressionMonoComp03", .OptionsPDFCompressionMonoComp03), "/n", vbCrLf)
51030   .OptionsPDFCompressionMonoComp04 = Replace$(hLang.Retrieve("PDFCompressionMonoComp04", .OptionsPDFCompressionMonoComp04), "/n", vbCrLf)
51040   .OptionsPDFCompressionMonoRes = Replace$(hLang.Retrieve("PDFCompressionMonoRes", .OptionsPDFCompressionMonoRes), "/n", vbCrLf)
51050   .OptionsPDFCompressionMonoResample = Replace$(hLang.Retrieve("PDFCompressionMonoResample", .OptionsPDFCompressionMonoResample), "/n", vbCrLf)
51060   .OptionsPDFCompressionMonoResample01 = Replace$(hLang.Retrieve("PDFCompressionMonoResample01", .OptionsPDFCompressionMonoResample01), "/n", vbCrLf)
51070   .OptionsPDFCompressionMonoResample02 = Replace$(hLang.Retrieve("PDFCompressionMonoResample02", .OptionsPDFCompressionMonoResample02), "/n", vbCrLf)
51080   .OptionsPDFCompressionMonoResample03 = Replace$(hLang.Retrieve("PDFCompressionMonoResample03", .OptionsPDFCompressionMonoResample03), "/n", vbCrLf)
51090   .OptionsPDFCompressionTextComp = Replace$(hLang.Retrieve("PDFCompressionTextComp", .OptionsPDFCompressionTextComp), "/n", vbCrLf)
51100   .OptionsPDFDescription = Replace$(hLang.Retrieve("PDFDescription", .OptionsPDFDescription), "/n", vbCrLf)
51110   .OptionsPDFDisallowCopy = Replace$(hLang.Retrieve("PDFDisallowCopy", .OptionsPDFDisallowCopy), "/n", vbCrLf)
51120   .OptionsPDFDisallowModify = Replace$(hLang.Retrieve("PDFDisallowModify", .OptionsPDFDisallowModify), "/n", vbCrLf)
51130   .OptionsPDFDisallowModifyComments = Replace$(hLang.Retrieve("PDFDisallowModifyComments", .OptionsPDFDisallowModifyComments), "/n", vbCrLf)
51140   .OptionsPDFDisallowPrint = Replace$(hLang.Retrieve("PDFDisallowPrint", .OptionsPDFDisallowPrint), "/n", vbCrLf)
51150   .OptionsPDFDisallowUser = Replace$(hLang.Retrieve("PDFDisallowUser", .OptionsPDFDisallowUser), "/n", vbCrLf)
51160   .OptionsPDFEncryptionHigh = Replace$(hLang.Retrieve("PDFEncryptionHigh", .OptionsPDFEncryptionHigh), "/n", vbCrLf)
51170   .OptionsPDFEncryptionLevel = Replace$(hLang.Retrieve("PDFEncryptionLevel", .OptionsPDFEncryptionLevel), "/n", vbCrLf)
51180   .OptionsPDFEncryptionLow = Replace$(hLang.Retrieve("PDFEncryptionLow", .OptionsPDFEncryptionLow), "/n", vbCrLf)
51190   .OptionsPDFEncryptor = Replace$(hLang.Retrieve("PDFEncryptor", .OptionsPDFEncryptor), "/n", vbCrLf)
51200   .OptionsPDFEnhancedPermissions = Replace$(hLang.Retrieve("PDFEnhancedPermissions", .OptionsPDFEnhancedPermissions), "/n", vbCrLf)
51210   .OptionsPDFEnterPasswords = Replace$(hLang.Retrieve("PDFEnterPasswords", .OptionsPDFEnterPasswords), "/n", vbCrLf)
51220   .OptionsPDFFonts = Replace$(hLang.Retrieve("PDFFonts", .OptionsPDFFonts), "/n", vbCrLf)
51230   .OptionsPDFFontsCaption = Replace$(hLang.Retrieve("PDFFontsCaption", .OptionsPDFFontsCaption), "/n", vbCrLf)
51240   .OptionsPDFFontsEmbedAll = Replace$(hLang.Retrieve("PDFFontsEmbedAll", .OptionsPDFFontsEmbedAll), "/n", vbCrLf)
51250   .OptionsPDFFontsSubSetFonts = Replace$(hLang.Retrieve("PDFFontsSubSetFonts", .OptionsPDFFontsSubSetFonts), "/n", vbCrLf)
51260   .OptionsPDFGeneral = Replace$(hLang.Retrieve("PDFGeneral", .OptionsPDFGeneral), "/n", vbCrLf)
51270   .OptionsPDFGeneralASCII85 = Replace$(hLang.Retrieve("PDFGeneralASCII85", .OptionsPDFGeneralASCII85), "/n", vbCrLf)
51280   .OptionsPDFGeneralAutorotate = Replace$(hLang.Retrieve("PDFGeneralAutorotate", .OptionsPDFGeneralAutorotate), "/n", vbCrLf)
51290   .OptionsPDFGeneralCaption = Replace$(hLang.Retrieve("PDFGeneralCaption", .OptionsPDFGeneralCaption), "/n", vbCrLf)
51300   .OptionsPDFGeneralCompatibility = Replace$(hLang.Retrieve("PDFGeneralCompatibility", .OptionsPDFGeneralCompatibility), "/n", vbCrLf)
51310   .OptionsPDFGeneralCompatibility01 = Replace$(hLang.Retrieve("PDFGeneralCompatibility01", .OptionsPDFGeneralCompatibility01), "/n", vbCrLf)
51320   .OptionsPDFGeneralCompatibility02 = Replace$(hLang.Retrieve("PDFGeneralCompatibility02", .OptionsPDFGeneralCompatibility02), "/n", vbCrLf)
51330   .OptionsPDFGeneralCompatibility03 = Replace$(hLang.Retrieve("PDFGeneralCompatibility03", .OptionsPDFGeneralCompatibility03), "/n", vbCrLf)
51340   .OptionsPDFGeneralOverprint = Replace$(hLang.Retrieve("PDFGeneralOverprint", .OptionsPDFGeneralOverprint), "/n", vbCrLf)
51350   .OptionsPDFGeneralOverprint01 = Replace$(hLang.Retrieve("PDFGeneralOverprint01", .OptionsPDFGeneralOverprint01), "/n", vbCrLf)
51360   .OptionsPDFGeneralOverprint02 = Replace$(hLang.Retrieve("PDFGeneralOverprint02", .OptionsPDFGeneralOverprint02), "/n", vbCrLf)
51370   .OptionsPDFGeneralResolution = Replace$(hLang.Retrieve("PDFGeneralResolution", .OptionsPDFGeneralResolution), "/n", vbCrLf)
51380   .OptionsPDFGeneralRotate01 = Replace$(hLang.Retrieve("PDFGeneralRotate01", .OptionsPDFGeneralRotate01), "/n", vbCrLf)
51390   .OptionsPDFGeneralRotate02 = Replace$(hLang.Retrieve("PDFGeneralRotate02", .OptionsPDFGeneralRotate02), "/n", vbCrLf)
51400   .OptionsPDFGeneralRotate03 = Replace$(hLang.Retrieve("PDFGeneralRotate03", .OptionsPDFGeneralRotate03), "/n", vbCrLf)
51410   .OptionsPDFOptions = Replace$(hLang.Retrieve("PDFOptions", .OptionsPDFOptions), "/n", vbCrLf)
51420   .OptionsPDFOwnerPass = Replace$(hLang.Retrieve("PDFOwnerPass", .OptionsPDFOwnerPass), "/n", vbCrLf)
51430   .OptionsPDFPasswords = Replace$(hLang.Retrieve("PDFPasswords", .OptionsPDFPasswords), "/n", vbCrLf)
51440   .OptionsPDFRepeatPassword = Replace$(hLang.Retrieve("PDFRepeatPassword", .OptionsPDFRepeatPassword), "/n", vbCrLf)
51450   .OptionsPDFSecurity = Replace$(hLang.Retrieve("PDFSecurity", .OptionsPDFSecurity), "/n", vbCrLf)
51460   .OptionsPDFSecurityCaption = Replace$(hLang.Retrieve("PDFSecurityCaption", .OptionsPDFSecurityCaption), "/n", vbCrLf)
51470   .OptionsPDFSetPassword = Replace$(hLang.Retrieve("PDFSetPassword", .OptionsPDFSetPassword), "/n", vbCrLf)
51480   .OptionsPDFSymbol = Replace$(hLang.Retrieve("PDFSymbol", .OptionsPDFSymbol), "/n", vbCrLf)
51490   .OptionsPDFUserPass = Replace$(hLang.Retrieve("PDFUserPass", .OptionsPDFUserPass), "/n", vbCrLf)
51500   .OptionsPDFUseSecurity = Replace$(hLang.Retrieve("PDFUseSecurity", .OptionsPDFUseSecurity), "/n", vbCrLf)
51510   .OptionsPNGColorscount01 = Replace$(hLang.Retrieve("PNGColorscount01", .OptionsPNGColorscount01), "/n", vbCrLf)
51520   .OptionsPNGColorscount02 = Replace$(hLang.Retrieve("PNGColorscount02", .OptionsPNGColorscount02), "/n", vbCrLf)
51530   .OptionsPNGColorscount03 = Replace$(hLang.Retrieve("PNGColorscount03", .OptionsPNGColorscount03), "/n", vbCrLf)
51540   .OptionsPNGColorscount04 = Replace$(hLang.Retrieve("PNGColorscount04", .OptionsPNGColorscount04), "/n", vbCrLf)
51550   .OptionsPNGColorscount05 = Replace$(hLang.Retrieve("PNGColorscount05", .OptionsPNGColorscount05), "/n", vbCrLf)
51560   .OptionsPNGDescription = Replace$(hLang.Retrieve("PNGDescription", .OptionsPNGDescription), "/n", vbCrLf)
51570   .OptionsPNGFiles = Replace$(hLang.Retrieve("PNGFiles", .OptionsPNGFiles), "/n", vbCrLf)
51580   .OptionsPNGSymbol = Replace$(hLang.Retrieve("PNGSymbol", .OptionsPNGSymbol), "/n", vbCrLf)
51590   .OptionsPrintertempDirectoryPrompt = Replace$(hLang.Retrieve("PrintertempDirectoryPrompt", .OptionsPrintertempDirectoryPrompt), "/n", vbCrLf)
51600   .OptionsPrintTestpage = Replace$(hLang.Retrieve("PrintTestpage", .OptionsPrintTestpage), "/n", vbCrLf)
51610   .OptionsProcesspriority = Replace$(hLang.Retrieve("Processpriority", .OptionsProcesspriority), "/n", vbCrLf)
51620   .OptionsProcesspriorityHigh = Replace$(hLang.Retrieve("ProcesspriorityHigh", .OptionsProcesspriorityHigh), "/n", vbCrLf)
51630   .OptionsProcesspriorityIdle = Replace$(hLang.Retrieve("ProcesspriorityIdle", .OptionsProcesspriorityIdle), "/n", vbCrLf)
51640   .OptionsProcesspriorityNormal = Replace$(hLang.Retrieve("ProcesspriorityNormal", .OptionsProcesspriorityNormal), "/n", vbCrLf)
51650   .OptionsProcesspriorityRealtime = Replace$(hLang.Retrieve("ProcesspriorityRealtime", .OptionsProcesspriorityRealtime), "/n", vbCrLf)
51660   .OptionsProgramAutosaveDescription = Replace$(hLang.Retrieve("ProgramAutosaveDescription", .OptionsProgramAutosaveDescription), "/n", vbCrLf)
51670   .OptionsProgramAutosaveSymbol = Replace$(hLang.Retrieve("ProgramAutosaveSymbol", .OptionsProgramAutosaveSymbol), "/n", vbCrLf)
51680   .OptionsProgramDirectoriesDescription = Replace$(hLang.Retrieve("ProgramDirectoriesDescription", .OptionsProgramDirectoriesDescription), "/n", vbCrLf)
51690   .OptionsProgramDirectoriesSymbol = Replace$(hLang.Retrieve("ProgramDirectoriesSymbol", .OptionsProgramDirectoriesSymbol), "/n", vbCrLf)
51700   .OptionsProgramDocumentDescription = Replace$(hLang.Retrieve("ProgramDocumentDescription", .OptionsProgramDocumentDescription), "/n", vbCrLf)
51710   .OptionsProgramDocumentSymbol = Replace$(hLang.Retrieve("ProgramDocumentSymbol", .OptionsProgramDocumentSymbol), "/n", vbCrLf)
51720   .OptionsProgramFont = Replace$(hLang.Retrieve("ProgramFont", .OptionsProgramFont), "/n", vbCrLf)
51730   .OptionsProgramFontCancelTest = Replace$(hLang.Retrieve("ProgramFontCancelTest", .OptionsProgramFontCancelTest), "/n", vbCrLf)
51740   .OptionsProgramFontcharset = Replace$(hLang.Retrieve("ProgramFontcharset", .OptionsProgramFontcharset), "/n", vbCrLf)
51750   .OptionsProgramFontDescription = Replace$(hLang.Retrieve("ProgramFontDescription", .OptionsProgramFontDescription), "/n", vbCrLf)
51760   .OptionsProgramFontSize = Replace$(hLang.Retrieve("ProgramFontSize", .OptionsProgramFontSize), "/n", vbCrLf)
51770   .OptionsProgramFontSymbol = Replace$(hLang.Retrieve("ProgramFontSymbol", .OptionsProgramFontSymbol), "/n", vbCrLf)
51780   .OptionsProgramFontTest = Replace$(hLang.Retrieve("ProgramFontTest", .OptionsProgramFontTest), "/n", vbCrLf)
51790   .OptionsProgramFontTestdescription = Replace$(hLang.Retrieve("ProgramFontTestdescription", .OptionsProgramFontTestdescription), "/n", vbCrLf)
51800   .OptionsProgramGeneralDescription = Replace$(hLang.Retrieve("ProgramGeneralDescription", .OptionsProgramGeneralDescription), "/n", vbCrLf)
51810   .OptionsProgramGeneralSymbol = Replace$(hLang.Retrieve("ProgramGeneralSymbol", .OptionsProgramGeneralSymbol), "/n", vbCrLf)
51820   .OptionsProgramGhostscriptDescription = Replace$(hLang.Retrieve("ProgramGhostscriptDescription", .OptionsProgramGhostscriptDescription), "/n", vbCrLf)
51830   .OptionsProgramGhostscriptSymbol = Replace$(hLang.Retrieve("ProgramGhostscriptSymbol", .OptionsProgramGhostscriptSymbol), "/n", vbCrLf)
51840   .OptionsProgramSaveDescription = Replace$(hLang.Retrieve("ProgramSaveDescription", .OptionsProgramSaveDescription), "/n", vbCrLf)
51850   .OptionsProgramSaveSymbol = Replace$(hLang.Retrieve("ProgramSaveSymbol", .OptionsProgramSaveSymbol), "/n", vbCrLf)
51860   .OptionsProgramSwitchingDefaultprinter = Replace$(hLang.Retrieve("ProgramSwitchingDefaultprinter", .OptionsProgramSwitchingDefaultprinter), "/n", vbCrLf)
51870   .OptionsPSDescription = Replace$(hLang.Retrieve("PSDescription", .OptionsPSDescription), "/n", vbCrLf)
51880   .OptionsPSFiles = Replace$(hLang.Retrieve("PSFiles", .OptionsPSFiles), "/n", vbCrLf)
51890   .OptionsPSLanguageLevel = Replace$(hLang.Retrieve("PSLanguageLevel", .OptionsPSLanguageLevel), "/n", vbCrLf)
51900   .OptionsPSSymbol = Replace$(hLang.Retrieve("PSSymbol", .OptionsPSSymbol), "/n", vbCrLf)
51910   .OptionsRemoveSpaces = Replace$(hLang.Retrieve("RemoveSpaces", .OptionsRemoveSpaces), "/n", vbCrLf)
51920   .OptionsReset = Replace$(hLang.Retrieve("Reset", .OptionsReset), "/n", vbCrLf)
51930   .OptionsSave = Replace$(hLang.Retrieve("Save", .OptionsSave), "/n", vbCrLf)
51940   .OptionsSaveFilename = Replace$(hLang.Retrieve("SaveFilename", .OptionsSaveFilename), "/n", vbCrLf)
51950   .OptionsSaveFilenameAdd = Replace$(hLang.Retrieve("SaveFilenameAdd", .OptionsSaveFilenameAdd), "/n", vbCrLf)
51960   .OptionsSaveFilenameChange = Replace$(hLang.Retrieve("SaveFilenameChange", .OptionsSaveFilenameChange), "/n", vbCrLf)
51970   .OptionsSaveFilenameDelete = Replace$(hLang.Retrieve("SaveFilenameDelete", .OptionsSaveFilenameDelete), "/n", vbCrLf)
51980   .OptionsSaveFilenameSubstitutions = Replace$(hLang.Retrieve("SaveFilenameSubstitutions", .OptionsSaveFilenameSubstitutions), "/n", vbCrLf)
51990   .OptionsSaveFilenameSubstitutionsTitle = Replace$(hLang.Retrieve("SaveFilenameSubstitutionsTitle", .OptionsSaveFilenameSubstitutionsTitle), "/n", vbCrLf)
52000   .OptionsSaveFilenameTokens = Replace$(hLang.Retrieve("SaveFilenameTokens", .OptionsSaveFilenameTokens), "/n", vbCrLf)
52010   .OptionsSavePasswords = Replace$(hLang.Retrieve("SavePasswords", .OptionsSavePasswords), "/n", vbCrLf)
52020   .OptionsShellIntegration = Replace$(hLang.Retrieve("ShellIntegration", .OptionsShellIntegration), "/n", vbCrLf)
52030   .OptionsShellIntegrationAdd = Replace$(hLang.Retrieve("ShellIntegrationAdd", .OptionsShellIntegrationAdd), "/n", vbCrLf)
52040   .OptionsShellIntegrationCaption = Replace$(hLang.Retrieve("ShellIntegrationCaption", .OptionsShellIntegrationCaption), "/n", vbCrLf)
52050   .OptionsShellIntegrationRemove = Replace$(hLang.Retrieve("ShellIntegrationRemove", .OptionsShellIntegrationRemove), "/n", vbCrLf)
52060   .OptionsStandardAuthorToken = Replace$(hLang.Retrieve("StandardAuthorToken", .OptionsStandardAuthorToken), "/n", vbCrLf)
52070   .OptionsTestpage = Replace$(hLang.Retrieve("Testpage", .OptionsTestpage), "/n", vbCrLf)
52080   .OptionsTIFFColorscount01 = Replace$(hLang.Retrieve("TIFFColorscount01", .OptionsTIFFColorscount01), "/n", vbCrLf)
52090   .OptionsTIFFColorscount02 = Replace$(hLang.Retrieve("TIFFColorscount02", .OptionsTIFFColorscount02), "/n", vbCrLf)
52100   .OptionsTIFFColorscount03 = Replace$(hLang.Retrieve("TIFFColorscount03", .OptionsTIFFColorscount03), "/n", vbCrLf)
52110   .OptionsTIFFColorscount04 = Replace$(hLang.Retrieve("TIFFColorscount04", .OptionsTIFFColorscount04), "/n", vbCrLf)
52120   .OptionsTIFFColorscount05 = Replace$(hLang.Retrieve("TIFFColorscount05", .OptionsTIFFColorscount05), "/n", vbCrLf)
52130   .OptionsTIFFColorscount06 = Replace$(hLang.Retrieve("TIFFColorscount06", .OptionsTIFFColorscount06), "/n", vbCrLf)
52140   .OptionsTIFFColorscount07 = Replace$(hLang.Retrieve("TIFFColorscount07", .OptionsTIFFColorscount07), "/n", vbCrLf)
52150   .OptionsTIFFColorscount08 = Replace$(hLang.Retrieve("TIFFColorscount08", .OptionsTIFFColorscount08), "/n", vbCrLf)
52160   .OptionsTIFFDescription = Replace$(hLang.Retrieve("TIFFDescription", .OptionsTIFFDescription), "/n", vbCrLf)
52170   .OptionsTIFFSymbol = Replace$(hLang.Retrieve("TIFFSymbol", .OptionsTIFFSymbol), "/n", vbCrLf)
52180   .OptionsTreeFormats = Replace$(hLang.Retrieve("TreeFormats", .OptionsTreeFormats), "/n", vbCrLf)
52190   .OptionsTreeProgram = Replace$(hLang.Retrieve("TreeProgram", .OptionsTreeProgram), "/n", vbCrLf)
52200   .OptionsUseAutosave = Replace$(hLang.Retrieve("UseAutosave", .OptionsUseAutosave), "/n", vbCrLf)
52210   .OptionsUseAutosaveDirectory = Replace$(hLang.Retrieve("UseAutosaveDirectory", .OptionsUseAutosaveDirectory), "/n", vbCrLf)
52220   .OptionsUseCreationDateNow = Replace$(hLang.Retrieve("UseCreationDateNow", .OptionsUseCreationDateNow), "/n", vbCrLf)
52230   .OptionsUserPass = Replace$(hLang.Retrieve("UserPass", .OptionsUserPass), "/n", vbCrLf)
52240   .OptionsUseStandardauthor = Replace$(hLang.Retrieve("UseStandardauthor", .OptionsUseStandardauthor), "/n", vbCrLf)
52250  End With
52260  Set hLang = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLanguage", "LoadOptionsStrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub LoadPrintingStrings(ByVal Languagefile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hLang As New clsHash
50020  ReadINISection Languagefile, "Printing", hLang
50030  With LanguageStrings
50040   .PrintingAuthor = Replace$(hLang.Retrieve("Author", .PrintingAuthor), "/n", vbCrLf)
50050   .PrintingBMPFiles = Replace$(hLang.Retrieve("BMPFiles", .PrintingBMPFiles), "/n", vbCrLf)
50060   .PrintingCreationDate = Replace$(hLang.Retrieve("CreationDate", .PrintingCreationDate), "/n", vbCrLf)
50070   .PrintingDocumentTitle = Replace$(hLang.Retrieve("DocumentTitle", .PrintingDocumentTitle), "/n", vbCrLf)
50080   .PrintingEMail = Replace$(hLang.Retrieve("EMail", .PrintingEMail), "/n", vbCrLf)
50090   .PrintingEPSFiles = Replace$(hLang.Retrieve("EPSFiles", .PrintingEPSFiles), "/n", vbCrLf)
50100   .PrintingJPEGFiles = Replace$(hLang.Retrieve("JPEGFiles", .PrintingJPEGFiles), "/n", vbCrLf)
50110   .PrintingKeywords = Replace$(hLang.Retrieve("Keywords", .PrintingKeywords), "/n", vbCrLf)
50120   .PrintingModifyDate = Replace$(hLang.Retrieve("ModifyDate", .PrintingModifyDate), "/n", vbCrLf)
50130   .PrintingNow = Replace$(hLang.Retrieve("Now", .PrintingNow), "/n", vbCrLf)
50140   .PrintingPCXFiles = Replace$(hLang.Retrieve("PCXFiles", .PrintingPCXFiles), "/n", vbCrLf)
50150   .PrintingPDFFiles = Replace$(hLang.Retrieve("PDFFiles", .PrintingPDFFiles), "/n", vbCrLf)
50160   .PrintingPNGFiles = Replace$(hLang.Retrieve("PNGFiles", .PrintingPNGFiles), "/n", vbCrLf)
50170   .PrintingPSFiles = Replace$(hLang.Retrieve("PSFiles", .PrintingPSFiles), "/n", vbCrLf)
50180   .PrintingSave = Replace$(hLang.Retrieve("Save", .PrintingSave), "/n", vbCrLf)
50190   .PrintingStartStandardProgram = Replace$(hLang.Retrieve("StartStandardProgram", .PrintingStartStandardProgram), "/n", vbCrLf)
50200   .PrintingStatus = Replace$(hLang.Retrieve("Status", .PrintingStatus), "/n", vbCrLf)
50210   .PrintingSubject = Replace$(hLang.Retrieve("Subject", .PrintingSubject), "/n", vbCrLf)
50220   .PrintingTIFFFiles = Replace$(hLang.Retrieve("TIFFFiles", .PrintingTIFFFiles), "/n", vbCrLf)
50230   .PrintingWaiting = Replace$(hLang.Retrieve("Waiting", .PrintingWaiting), "/n", vbCrLf)
50240  End With
50250  Set hLang = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLanguage", "LoadPrintingStrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub LoadSaveOpenStrings(ByVal Languagefile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hLang As New clsHash
50020  ReadINISection Languagefile, "SaveOpen", hLang
50030  With LanguageStrings
50040   .SaveOpenAttributes = Replace$(hLang.Retrieve("Attributes", .SaveOpenAttributes), "/n", vbCrLf)
50050   .SaveOpenCancel = Replace$(hLang.Retrieve("Cancel", .SaveOpenCancel), "/n", vbCrLf)
50060   .SaveOpenFilename = Replace$(hLang.Retrieve("Filename", .SaveOpenFilename), "/n", vbCrLf)
50070   .SaveOpenOpen = Replace$(hLang.Retrieve("Open", .SaveOpenOpen), "/n", vbCrLf)
50080   .SaveOpenOpenTitle = Replace$(hLang.Retrieve("OpenTitle", .SaveOpenOpenTitle), "/n", vbCrLf)
50090   .SaveOpenSave = Replace$(hLang.Retrieve("Save", .SaveOpenSave), "/n", vbCrLf)
50100   .SaveOpenSaveTitle = Replace$(hLang.Retrieve("SaveTitle", .SaveOpenSaveTitle), "/n", vbCrLf)
50110   .SaveOpenSize = Replace$(hLang.Retrieve("Size", .SaveOpenSize), "/n", vbCrLf)
50120  End With
50130  Set hLang = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLanguage", "LoadSaveOpenStrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub InitLanguagesStrings()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With LanguageStrings
50020   .CommonAuthor = "Philip Chinery, Frank Heindörfer"
50030   .CommonLanguagename = "English"
50040   .CommonTitle = "PDF Print monitor"
50050   .CommonVersion = "0.8.0"
50060
50070   .DialogDocument = "Document"
50080   .DialogDocumentAdd = "Add"
50090   .DialogDocumentBottom = "Bottom"
50100   .DialogDocumentCombine = "Combine"
50110   .DialogDocumentDelete = "Delete"
50120   .DialogDocumentDown = "Down"
50130   .DialogDocumentPrint = "Print"
50140   .DialogDocumentSave = "Save"
50150   .DialogDocumentTop = "Top"
50160   .DialogDocumentUp = "Up"
50170   .DialogInfo = "?"
50180   .DialogInfoCheckUpdates = "Check for Updates"
50190   .DialogInfoHomepage = "Product Homepage"
50200   .DialogInfoInfo = "About"
50210   .DialogInfoPaypal = "Paypal"
50220   .DialogInfoPDFCreatorSourceforge = "PDFCreator on Sourceforge"
50230   .DialogLanguage = "Language"
50240   .DialogPrinter = "Printer"
50250   .DialogPrinterClose = "Close"
50260   .DialogPrinterLogfile = "Logfile"
50270   .DialogPrinterLogfiles = "Logfiles"
50280   .DialogPrinterLogging = "Logging"
50290   .DialogPrinterOptions = "Options"
50300   .DialogPrinterPrinterStop = "Printer stop"
50310   .DialogView = "View"
50320   .DialogViewStatusbar = "Status Bar"
50330
50340   .ListAddFile = "Add a file"
50350   .ListAllFiles = "All files"
50360   .ListBytes = "Bytes"
50370   .ListDate = "Created on"
50380   .ListDocumenttitle = "Document Title"
50390   .ListFilename = "Filename"
50400   .ListGBytes = "GBytes"
50410   .ListKBytes = "kBytes"
50420   .ListMBytes = "MBytes"
50430   .ListPDFFiles = "PDF Files"
50440   .ListPostscriptFiles = "PostScript Files"
50450   .ListPrinting = "Printing"
50460   .ListSize = "Size"
50470   .ListStatus = "Status"
50480   .ListWaiting = "Waiting"
50490
50500   .LoggingClear = "Clear"
50510   .LoggingClose = "Close"
50520   .LoggingLogfile = "Logfile"
50530
50540   .MessagesMsg01 = "Document in queue."
50550   .MessagesMsg02 = "Documents in queue."
50560   .MessagesMsg03 = "Do you wish to reset all settings?"
50570   .MessagesMsg04 = "Error: Cannot send Email!"
50580   .MessagesMsg05 = "File already exists. Do you want to overwrite it?"
50590   .MessagesMsg06 = "This file does not seem to be a postscript file!"
50600   .MessagesMsg07 = "There is a problem when trying to access this drive or directory!"
50610   .MessagesMsg08 = "Cannot find gsdll32.dll. Please check the ghostscript-program directory (see options)!"
50620   .MessagesMsg09 = "The output path does not exist. Do you want to create it?"
50630   .MessagesMsg10 = "This is not a valid path!"
50640   .MessagesMsg11 = "There is already such an entry!"
50650   .MessagesMsg12 = "Please don't use these forbidden characters for a filename!"
50660   .MessagesMsg13 = "Delete all program settings?"
50670   .MessagesMsg14 = "The file can not be found!"
50680   .MessagesMsg15 = "Cannot find gsdll32.dll in this directory!"
50690   .MessagesMsg16 = "No ghostscript font found in this directory!"
50700   .MessagesMsg17 = "No files in this directory!"
50710   .MessagesMsg19 = "You need either pdfenc or AFPL Ghostscript greater than, or equal to, version 8.14!"
50720   .MessagesMsg20 = "There was a problem sending an email with the standard emailclient!"
50730   .MessagesMsg21 = "User passwords do not match!"
50740   .MessagesMsg22 = "Owner passwords do not match!"
50750   .MessagesMsg23 = "The document is not protected!"
50760   .MessagesMsg24 = "The user password is empty! Continue?"
50770   .MessagesMsg25 = "The owner password is empty! Continue?"
50780   .MessagesMsg26 = "Unknown error"
50790   .MessagesMsg27 = "Cannot find the file/page."
50800   .MessagesMsg28 = "The filesize is 0 byte."
50810   .MessagesMsg29 = "Server not found."
50820   .MessagesMsg30 = "The url isn not interpretable."
50830   .MessagesMsg31 = "An error has occured"
50840   .MessagesMsg32 = "The new version %1 is available. Would you like download the new version from the Sourceforge pages?"
50850   .MessagesMsg33 = "You already have the most recent version."
50860   .MessagesMsg34 = "The file is in use. Please close the file first or choose another filename."
50870   .MessagesMsg35 = "It is necessary to temporarily set PDFCreator as defaultprinter."
50880   .MessagesMsg36 = "Don't ask me again."
50890
50900   .OptionsAssociatePSFiles = "Associate PDFCreator with postscript files"
50910   .OptionsAutosaveDirectoryPrompt = "Select Autosave Directory"
50920   .OptionsAutosaveFilename = "Filename"
50930   .OptionsAutosaveFilenameTokens = "Add a Filename-Token"
50940   .OptionsAutosaveFormat = "Autosave format"
50950   .OptionsBitmapResolution = "Resolution"
50960   .OptionsBMPColorscount01 = "4294967296 colors (32 Bit)"
50970   .OptionsBMPColorscount02 = "16777216 colors (24 Bit)"
50980   .OptionsBMPColorscount03 = "256 colors (8 Bit)"
50990   .OptionsBMPColorscount04 = "16 colors (4 Bit)"
51000   .OptionsBMPColorscount05 = "8 colors (3 Bit)"
51010   .OptionsBMPColorscount06 = "2 colors (Black/White)"
51020   .OptionsBMPColorscount07 = "Greyscale (8 Bit)"
51030   .OptionsBMPDescription = "Windows Bitmap Format. Please use only for single pages."
51040   .OptionsBMPSymbol = "BMP"
51050   .OptionsCancel = "Cancel"
51060   .OptionsDirectoriesGSBin = "Ghostscript Binaries"
51070   .OptionsDirectoriesGSFonts = "Ghostscript Fonts"
51080   .OptionsDirectoriesGSLibraries = "Ghostscript Libraries"
51090   .OptionsDirectoriesTempPath = "Temporary Files"
51100   .OptionsDocument = "Document"
51110   .OptionsEPSDescription = "Encapsulated Postscript Format"
51120   .OptionsEPSFiles = "Encapsulated Postscript-Files"
51130   .OptionsEPSSymbol = "EPS"
51140   .OptionsGhostscriptBinariesDirectoryPrompt = "Select Ghostscript Binaries Directory"
51150   .OptionsGhostscriptFontsDirectoryPrompt = "Select Ghostscript Fonts Directory"
51160   .OptionsGhostscriptInternal = "Internal Ghostscript: %1 Ghostscript %2"
51170   .OptionsGhostscriptLibrariesDirectoryPrompt = "Select Ghostscript Libraries Directory"
51180   .OptionsGhostscriptResourceDirectoryPrompt = "Select Ghostscript Resource Directory"
51190   .OptionsGhostscriptversion = "Ghostscript Version"
51200   .OptionsImageSettings = "Settings"
51210   .OptionsJavaPath = "Path to Java Interpreter"
51220   .OptionsJPEGColorscount01 = "16777216 colors (24 Bit)"
51230   .OptionsJPEGColorscount02 = "Greyscale (8 Bit)"
51240   .OptionsJPEGDescription = "JPEG (JFIF) Format. Please use only for single pages."
51250   .OptionsJPEGQuality = "Quality:"
51260   .OptionsJPEGSymbol = "JPEG"
51270   .OptionsOwnerPass = "Owner Password"
51280   .OptionsPassCancel = "Cancel"
51290   .OptionsPassOK = "OK"
51300   .OptionsPCXColorscount01 = "4294967296 colors (32 Bit) CMYK"
51310   .OptionsPCXColorscount02 = "16777216 colors (24 Bit)"
51320   .OptionsPCXColorscount03 = "256 colors (8 Bit)"
51330   .OptionsPCXColorscount04 = "16 colors (4 Bit)"
51340   .OptionsPCXColorscount05 = "2 colors (Black/White)"
51350   .OptionsPCXColorscount06 = "Greyscale (8 Bit)"
51360   .OptionsPCXDescription = "PCX Format. Please use only for single pages."
51370   .OptionsPCXSymbol = "PCX"
51380   .OptionsPDFAllowAssembly = "Allow changes to the assembly"
51390   .OptionsPDFAllowDegradedPrinting = "Allow printing in low resolution"
51400   .OptionsPDFAllowFillIn = "Allow filling in form fields"
51410   .OptionsPDFAllowScreenReaders = "Allow screen readers"
51420   .OptionsPDFColors = "Colors"
51430   .OptionsPDFColorsCaption = "Color Options"
51440   .OptionsPDFColorsCMYKtoRGB = "Convert CMYK images to RGB"
51450   .OptionsPDFColorsColorModel01 = "Use Color Model Device RGB"
51460   .OptionsPDFColorsColorModel02 = "Use Color Model Device CMYK"
51470   .OptionsPDFColorsColorModel03 = "Use Color Model Device Grayscale"
51480   .OptionsPDFColorsColorOptions = "Options"
51490   .OptionsPDFColorsPreserveHalftone = "Preserve Halftone Information"
51500   .OptionsPDFColorsPreserveOverprint = "Preserve Overprint Settings"
51510   .OptionsPDFColorsPreserveTransfer = "Preserve Transfer Functions"
51520   .OptionsPDFCompression = "Compression"
51530   .OptionsPDFCompressionCaption = "PDF Compression"
51540   .OptionsPDFCompressionColor = "Color Images"
51550   .OptionsPDFCompressionColorComp = "Compress"
51560   .OptionsPDFCompressionColorComp01 = "Automatic"
51570   .OptionsPDFCompressionColorComp02 = "JPEG-Maximum"
51580   .OptionsPDFCompressionColorComp03 = "JPEG-High"
51590   .OptionsPDFCompressionColorComp04 = "JPEG-Medium"
51600   .OptionsPDFCompressionColorComp05 = "JPEG-Low"
51610   .OptionsPDFCompressionColorComp06 = "JPEG-Minimum"
51620   .OptionsPDFCompressionColorComp07 = "ZIP"
51630   .OptionsPDFCompressionColorComp08 = "LZW-Compression"
51640   .OptionsPDFCompressionColorRes = "Resolution"
51650   .OptionsPDFCompressionColorResample = "Resample"
51660   .OptionsPDFCompressionColorResample01 = "Bicubic"
51670   .OptionsPDFCompressionColorResample02 = "Downsample"
51680   .OptionsPDFCompressionColorResample03 = "Average Downsample"
51690   .OptionsPDFCompressionGrey = "Greyscale Images"
51700   .OptionsPDFCompressionGreyComp = "Compress"
51710   .OptionsPDFCompressionGreyComp01 = "Automatic"
51720   .OptionsPDFCompressionGreyComp02 = "JPEG-Maximum"
51730   .OptionsPDFCompressionGreyComp03 = "JPEG-High"
51740   .OptionsPDFCompressionGreyComp04 = "JPEG-Medium"
51750   .OptionsPDFCompressionGreyComp05 = "JPEG-Low"
51760   .OptionsPDFCompressionGreyComp06 = "JPEG-Minimum"
51770   .OptionsPDFCompressionGreyComp07 = "ZIP"
51780   .OptionsPDFCompressionGreyComp08 = "LZW-Compression"
51790   .OptionsPDFCompressionGreyRes = "Resolution"
51800   .OptionsPDFCompressionGreyResample = "Resample"
51810   .OptionsPDFCompressionGreyResample01 = "Bicubic"
51820   .OptionsPDFCompressionGreyResample02 = "Downsample"
51830   .OptionsPDFCompressionGreyResample03 = "Average Downsample"
51840   .OptionsPDFCompressionMono = "Monochrome Images"
51850   .OptionsPDFCompressionMonoComp = "Compress"
51860   .OptionsPDFCompressionMonoComp01 = "CCITT Fax Compression"
51870   .OptionsPDFCompressionMonoComp02 = "ZIP"
51880   .OptionsPDFCompressionMonoComp03 = "LZW-Compression"
51890   .OptionsPDFCompressionMonoComp04 = "Run-Length-Encoding"
51900   .OptionsPDFCompressionMonoRes = "Resolution"
51910   .OptionsPDFCompressionMonoResample = "Resample"
51920   .OptionsPDFCompressionMonoResample01 = "Bicubic"
51930   .OptionsPDFCompressionMonoResample02 = "Downsample"
51940   .OptionsPDFCompressionMonoResample03 = "Average Downsample"
51950   .OptionsPDFCompressionTextComp = "Compress Text Objects"
51960   .OptionsPDFDescription = "Adobe PDF Format"
51970   .OptionsPDFDisallowCopy = "Copy text and images"
51980   .OptionsPDFDisallowModify = "Modify the document"
51990   .OptionsPDFDisallowModifyComments = "Modify comments"
52000   .OptionsPDFDisallowPrint = "Print the document"
52010   .OptionsPDFDisallowUser = "Disallow User to"
52020   .OptionsPDFEncryptionHigh = "High (128 Bit - Adobe Acrobat 5.0 and above)"
52030   .OptionsPDFEncryptionLevel = "Encryption Level"
52040   .OptionsPDFEncryptionLow = "Low (40 Bit - Adobe Acrobat 3.0 and above)"
52050   .OptionsPDFEncryptor = "Enrcyptor"
52060   .OptionsPDFEnhancedPermissions = "Enhanced Permissions (128 Bit only)"
52070   .OptionsPDFEnterPasswords = "Enter Passwords"
52080   .OptionsPDFFonts = "Fonts"
52090   .OptionsPDFFontsCaption = "Font Options"
52100   .OptionsPDFFontsEmbedAll = "Embed all fonts"
52110   .OptionsPDFFontsSubSetFonts = "Subset fonts when percentage of used characters below:"
52120   .OptionsPDFGeneral = "General"
52130   .OptionsPDFGeneralASCII85 = "Convert binary data to ASCII85"
52140   .OptionsPDFGeneralAutorotate = "Auto-Rotate Pages:"
52150   .OptionsPDFGeneralCaption = "General Options"
52160   .OptionsPDFGeneralCompatibility = "Compatibility:"
52170   .OptionsPDFGeneralCompatibility01 = "Adobe Acrobat 3.0 (PDF 1.2)"
52180   .OptionsPDFGeneralCompatibility02 = "Adobe Acrobat 4.0 (PDF 1.3)"
52190   .OptionsPDFGeneralCompatibility03 = "Adobe Acrobat 5.0 (PDF 1.4)"
52200   .OptionsPDFGeneralOverprint = "Overprint:"
52210   .OptionsPDFGeneralOverprint01 = "Non-Zero Overprint"
52220   .OptionsPDFGeneralOverprint02 = "Full Overprint"
52230   .OptionsPDFGeneralResolution = "Resolution:"
52240   .OptionsPDFGeneralRotate01 = "None"
52250   .OptionsPDFGeneralRotate02 = "All"
52260   .OptionsPDFGeneralRotate03 = "Single Page"
52270   .OptionsPDFOptions = "PDF Options"
52280   .OptionsPDFOwnerPass = "Password required to change permissions and passwords"
52290   .OptionsPDFPasswords = "Passwords"
52300   .OptionsPDFRepeatPassword = "Repeat"
52310   .OptionsPDFSecurity = "Security"
52320   .OptionsPDFSecurityCaption = "Security"
52330   .OptionsPDFSetPassword = "Password"
52340   .OptionsPDFSymbol = "PDF"
52350   .OptionsPDFUserPass = "Password required to open document"
52360   .OptionsPDFUseSecurity = "Use Security"
52370   .OptionsPNGColorscount01 = "16777216 colors (24 Bit)"
52380   .OptionsPNGColorscount02 = "256 colors (8 Bit)"
52390   .OptionsPNGColorscount03 = "16 colors (4 Bit)"
52400   .OptionsPNGColorscount04 = "2 colors (2 Bit - Black/White)"
52410   .OptionsPNGColorscount05 = "Greyscale (8 Bit)"
52420   .OptionsPNGDescription = "PNG Format. Please use only for single pages."
52430   .OptionsPNGFiles = "Bitmap PNG-Files"
52440   .OptionsPNGSymbol = "PNG"
52450   .OptionsPrintertempDirectoryPrompt = "Select Printer Temp-Directory"
52460   .OptionsPrintTestpage = "Print Test Page"
52470   .OptionsProcesspriority = "Processpriority"
52480   .OptionsProcesspriorityHigh = "High"
52490   .OptionsProcesspriorityIdle = "Idle"
52500   .OptionsProcesspriorityNormal = "Normal"
52510   .OptionsProcesspriorityRealtime = "Realtime"
52520   .OptionsProgramAutosaveDescription = "Auto-save mode. Auto-save does not prompt for a filename and file location. It automatically saves all PDF files to a single directory with a predefined filename."
52530   .OptionsProgramAutosaveSymbol = "Auto-save"
52540   .OptionsProgramDirectoriesDescription = "Directories for Ghostscript, temporary files and others."
52550   .OptionsProgramDirectoriesSymbol = "Directories"
52560   .OptionsProgramDocumentDescription = "Document properties"
52570   .OptionsProgramDocumentSymbol = "Document"
52580   .OptionsProgramFont = "Program Font"
52590   .OptionsProgramFontCancelTest = "Cancel Test"
52600   .OptionsProgramFontcharset = "Character Set"
52610   .OptionsProgramFontDescription = "Font for labels, captions and values. For the program menu use the general settings in your Windows OS."
52620   .OptionsProgramFontSize = "Size"
52630   .OptionsProgramFontSymbol = "Fonts"
52640   .OptionsProgramFontTest = "Test"
52650   .OptionsProgramFontTestdescription = "Here you can test the font."
52660   .OptionsProgramGeneralDescription = "General Settings."
52670   .OptionsProgramGeneralSymbol = "General settings"
52680   .OptionsProgramGhostscriptDescription = "Ghostscript"
52690   .OptionsProgramGhostscriptSymbol = "Ghostscript"
52700   .OptionsProgramSaveDescription = "Save"
52710   .OptionsProgramSaveSymbol = "Save"
52720   .OptionsProgramSwitchingDefaultprinter = "No confirm message switching PDFCreator temporarily as default printer."
52730   .OptionsPSDescription = "Postscript Format"
52740   .OptionsPSFiles = "Postscript-Files"
52750   .OptionsPSLanguageLevel = "Language Level:"
52760   .OptionsPSSymbol = "PS"
52770   .OptionsRemoveSpaces = "Remove leading and trailing spaces"
52780   .OptionsReset = "Reset all settings"
52790   .OptionsSave = "Save"
52800   .OptionsSaveFilename = "Filename"
52810   .OptionsSaveFilenameAdd = "Add"
52820   .OptionsSaveFilenameChange = "Change"
52830   .OptionsSaveFilenameDelete = "Delete"
52840   .OptionsSaveFilenameSubstitutions = "Filename substitution"
52850   .OptionsSaveFilenameSubstitutionsTitle = "Filename substitution only in <Title>"
52860   .OptionsSaveFilenameTokens = "Add a Filename-Token"
52870   .OptionsSavePasswords = "Save passwords temporarily for this sesion."
52880   .OptionsShellIntegration = "Shell integration"
52890   .OptionsShellIntegrationAdd = "Integrate PDFCreator into shell"
52900   .OptionsShellIntegrationCaption = "Create &PDF with PDFCreator"
52910   .OptionsShellIntegrationRemove = "Remove shell integration"
52920   .OptionsStandardAuthorToken = "Add a Author-Token"
52930   .OptionsTestpage = "PDFCreator Testpage"
52940   .OptionsTIFFColorscount01 = "16777216 (24 Bit)"
52950   .OptionsTIFFColorscount02 = "4096 (12 Bit)"
52960   .OptionsTIFFColorscount03 = "2 colors (Black/White) G3 fax encoding with no EOLs"
52970   .OptionsTIFFColorscount04 = "2 colors (Black/White) G3 fax encoding with EOLs"
52980   .OptionsTIFFColorscount05 = "2 colors (Black/White) 2-D G3 fax encoding"
52990   .OptionsTIFFColorscount06 = "2 colors (Black/White) G4 fax encoding"
53000   .OptionsTIFFColorscount07 = "2 colors (Black/White) LZW-compatible"
53010   .OptionsTIFFColorscount08 = "2 colors (Black/White) PackBits"
53020   .OptionsTIFFDescription = "TIFF Format. For multipages use the tiff-format."
53030   .OptionsTIFFSymbol = "TIFF"
53040   .OptionsTreeFormats = "Formats"
53050   .OptionsTreeProgram = "Program"
53060   .OptionsUseAutosave = "Use Auto-save"
53070   .OptionsUseAutosaveDirectory = "Use this directory for auto-save"
53080   .OptionsUseCreationDateNow = "Use the current Date/Time for 'Creation Date'"
53090   .OptionsUserPass = "User Password"
53100   .OptionsUseStandardauthor = "Use standard author"
53110
53120   .PrintingAuthor = "Author:"
53130   .PrintingBMPFiles = "BMP-Files"
53140   .PrintingCreationDate = "Creation Date:"
53150   .PrintingDocumentTitle = "Document Title:"
53160   .PrintingEMail = "eMail"
53170   .PrintingEPSFiles = "Encapsulated Postscript-Files"
53180   .PrintingJPEGFiles = "JPEG-Files"
53190   .PrintingKeywords = "Keywords:"
53200   .PrintingModifyDate = "Modify Date:"
53210   .PrintingNow = "Now"
53220   .PrintingPCXFiles = "PCX-Files"
53230   .PrintingPDFFiles = "PDF-Files"
53240   .PrintingPNGFiles = "PNG-Files"
53250   .PrintingPSFiles = "Postscript-Files"
53260   .PrintingSave = "Save"
53270   .PrintingStartStandardProgram = "After saving, open the document with the default program."
53280   .PrintingStatus = "Creating file..."
53290   .PrintingSubject = "Subject:"
53300   .PrintingTIFFFiles = "TIFF-Files"
53310   .PrintingWaiting = "Waiting"
53320
53330   .SaveOpenAttributes = "Attributes"
53340   .SaveOpenCancel = "Cancel"
53350   .SaveOpenFilename = "Filename"
53360   .SaveOpenOpen = "Open"
53370   .SaveOpenOpenTitle = "Open"
53380   .SaveOpenSave = "Save"
53390   .SaveOpenSaveTitle = "Save as"
53400   .SaveOpenSize = "Size"
53410
53420  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLanguage", "InitLanguagesStrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

