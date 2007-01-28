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
 MessagesMsg37 As String
 MessagesMsg38 As String

 OptionsAdditionalGhostscriptParameters As String
 OptionsAdditionalGhostscriptSearchpath As String
 OptionsAddWindowsFontpath As String
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
 OptionsCustomPapersizeHeight As String
 OptionsCustomPapersizeInfo As String
 OptionsCustomPapersizeWidth As String
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
 OptionsLanguagesCurrentLanguage As String
 OptionsLanguagesDownloadMoreLanguages As String
 OptionsLanguagesInstall As String
 OptionsLanguagesRefresh As String
 OptionsLanguagesTranslation As String
 OptionsLanguagesVersion As String
 OptionsOnePagePerFile As String
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
 OptionsPDFOptimize As String
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
 OptionsPrintAfterSaving As String
 OptionsPrintAfterSavingDuplex As String
 OptionsPrintAfterSavingDuplexTumbleOff As String
 OptionsPrintAfterSavingDuplexTumbleOn As String
 OptionsPrintAfterSavingNoCancel As String
 OptionsPrintAfterSavingPrinter As String
 OptionsPrintAfterSavingQueryUser As String
 OptionsPrintAfterSavingQueryUserDefaultPrinter As String
 OptionsPrintAfterSavingQueryUserOff As String
 OptionsPrintAfterSavingQueryUserPrinterSetupDialog As String
 OptionsPrintAfterSavingQueryUserStandardPrinterDialog As String
 OptionsPrintertempDirectoryPrompt As String
 OptionsPrintTestpage As String
 OptionsProcesspriority As String
 OptionsProcesspriorityHigh As String
 OptionsProcesspriorityIdle As String
 OptionsProcesspriorityNormal As String
 OptionsProcesspriorityRealtime As String
 OptionsProgramActionsDescription As String
 OptionsProgramActionsSymbol As String
 OptionsProgramAutosaveDescription As String
 OptionsProgramAutosaveSymbol As String
 OptionsProgramDirectoriesDescription As String
 OptionsProgramDirectoriesSymbol As String
 OptionsProgramDocumentDescription As String
 OptionsProgramDocumentDescription1 As String
 OptionsProgramDocumentDescription2 As String
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
 OptionsProgramGeneralDescription1 As String
 OptionsProgramGeneralDescription2 As String
 OptionsProgramGeneralSymbol As String
 OptionsProgramGhostscriptDescription As String
 OptionsProgramGhostscriptSymbol As String
 OptionsProgramLanguagesDescription As String
 OptionsProgramLanguagesSymbol As String
 OptionsProgramNoProcessingAtStartup As String
 OptionsProgramOptionsDesign As String
 OptionsProgramOptionsDesignGradient As String
 OptionsProgramOptionsDesignSimple As String
 OptionsProgramPrintDescription As String
 OptionsProgramPrintSymbol As String
 OptionsProgramRunProgramAfterSavingCaption As String
 OptionsProgramRunProgramAfterSavingProgram As String
 OptionsProgramRunProgramAfterSavingProgramParameters As String
 OptionsProgramRunProgramAfterSavingWaitUntilReady As String
 OptionsProgramRunProgramAfterSavingWindowstyle As String
 OptionsProgramRunProgramAfterSavingWindowstyleHide As String
 OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus As String
 OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus As String
 OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus As String
 OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus As String
 OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus As String
 OptionsProgramRunProgramBeforeSavingCaption As String
 OptionsProgramRunProgramBeforeSavingProgram As String
 OptionsProgramRunProgramBeforeSavingProgramParameters As String
 OptionsProgramRunProgramBeforeSavingWaitUntilReady As String
 OptionsProgramRunProgramBeforeSavingWindowstyle As String
 OptionsProgramRunProgramBeforeSavingWindowstyleHide As String
 OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus As String
 OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus As String
 OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus As String
 OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus As String
 OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus As String
 OptionsProgramSaveDescription As String
 OptionsProgramSaveSymbol As String
 OptionsProgramShowAnimation As String
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
 OptionsSendEmailAfterAutosave As String
 OptionsSendMailMethod As String
 OptionsSendMailMethodAutomatic As String
 OptionsSendMailMethodMapi As String
 OptionsSendMailMethodSendmailDLL As String
 OptionsShellIntegration As String
 OptionsShellIntegrationAdd As String
 OptionsShellIntegrationCaption As String
 OptionsShellIntegrationRemove As String
 OptionsStamp As String
 OptionsStampFontColor As String
 OptionsStampOutlineFontThickness As String
 OptionsStampString As String
 OptionsStampUseOutlineFont As String
 OptionsStandardAuthorToken As String
 OptionsStandardSaveFormat As String
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
 OptionsUseCustomPapersize As String
 OptionsUseFixPapersize As String
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
50080   .DialogDocumentCombineAll = Replace$(hLang.Retrieve("DocumentCombineAll", .DialogDocumentCombineAll), "/n", vbCrLf)
50090   .DialogDocumentCombineAllSend = Replace$(hLang.Retrieve("DocumentCombineAllSend", .DialogDocumentCombineAllSend), "/n", vbCrLf)
50100   .DialogDocumentDelete = Replace$(hLang.Retrieve("DocumentDelete", .DialogDocumentDelete), "/n", vbCrLf)
50110   .DialogDocumentDown = Replace$(hLang.Retrieve("DocumentDown", .DialogDocumentDown), "/n", vbCrLf)
50120   .DialogDocumentPrint = Replace$(hLang.Retrieve("DocumentPrint", .DialogDocumentPrint), "/n", vbCrLf)
50130   .DialogDocumentSave = Replace$(hLang.Retrieve("DocumentSave", .DialogDocumentSave), "/n", vbCrLf)
50140   .DialogDocumentSend = Replace$(hLang.Retrieve("DocumentSend", .DialogDocumentSend), "/n", vbCrLf)
50150   .DialogDocumentTop = Replace$(hLang.Retrieve("DocumentTop", .DialogDocumentTop), "/n", vbCrLf)
50160   .DialogDocumentUp = Replace$(hLang.Retrieve("DocumentUp", .DialogDocumentUp), "/n", vbCrLf)
50170   .DialogEmailAddress = Replace$(hLang.Retrieve("EmailAddress", .DialogEmailAddress), "/n", vbCrLf)
50180   .DialogInfo = Replace$(hLang.Retrieve("Info", .DialogInfo), "/n", vbCrLf)
50190   .DialogInfoCheckUpdates = Replace$(hLang.Retrieve("InfoCheckUpdates", .DialogInfoCheckUpdates), "/n", vbCrLf)
50200   .DialogInfoHomepage = Replace$(hLang.Retrieve("InfoHomepage", .DialogInfoHomepage), "/n", vbCrLf)
50210   .DialogInfoInfo = Replace$(hLang.Retrieve("InfoInfo", .DialogInfoInfo), "/n", vbCrLf)
50220   .DialogInfoPaypal = Replace$(hLang.Retrieve("InfoPaypal", .DialogInfoPaypal), "/n", vbCrLf)
50230   .DialogInfoPDFCreatorSourceforge = Replace$(hLang.Retrieve("InfoPDFCreatorSourceforge", .DialogInfoPDFCreatorSourceforge), "/n", vbCrLf)
50240   .DialogLanguage = Replace$(hLang.Retrieve("Language", .DialogLanguage), "/n", vbCrLf)
50250   .DialogPrinter = Replace$(hLang.Retrieve("Printer", .DialogPrinter), "/n", vbCrLf)
50260   .DialogPrinterClose = Replace$(hLang.Retrieve("PrinterClose", .DialogPrinterClose), "/n", vbCrLf)
50270   .DialogPrinterLogfile = Replace$(hLang.Retrieve("PrinterLogfile", .DialogPrinterLogfile), "/n", vbCrLf)
50280   .DialogPrinterLogfiles = Replace$(hLang.Retrieve("PrinterLogfiles", .DialogPrinterLogfiles), "/n", vbCrLf)
50290   .DialogPrinterLogging = Replace$(hLang.Retrieve("PrinterLogging", .DialogPrinterLogging), "/n", vbCrLf)
50300   .DialogPrinterOptions = Replace$(hLang.Retrieve("PrinterOptions", .DialogPrinterOptions), "/n", vbCrLf)
50310   .DialogPrinterPrinterStop = Replace$(hLang.Retrieve("PrinterPrinterStop", .DialogPrinterPrinterStop), "/n", vbCrLf)
50320   .DialogView = Replace$(hLang.Retrieve("View", .DialogView), "/n", vbCrLf)
50330   .DialogViewStatusbar = Replace$(hLang.Retrieve("ViewStatusbar", .DialogViewStatusbar), "/n", vbCrLf)
50340   .DialogViewToolbars = Replace$(hLang.Retrieve("ViewToolbars", .DialogViewToolbars), "/n", vbCrLf)
50350   .DialogViewToolbarsEmail = Replace$(hLang.Retrieve("ViewToolbarsEmail", .DialogViewToolbarsEmail), "/n", vbCrLf)
50360   .DialogViewToolbarsStandard = Replace$(hLang.Retrieve("ViewToolbarsStandard", .DialogViewToolbarsStandard), "/n", vbCrLf)
50370  End With
50380  Set hLang = Nothing
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
50390   .MessagesMsg37 = Replace$(hLang.Retrieve("Msg37", .MessagesMsg37), "/n", vbCrLf)
50400   .MessagesMsg38 = Replace$(hLang.Retrieve("Msg38", .MessagesMsg38), "/n", vbCrLf)
50410  End With
50420  Set hLang = Nothing
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
50040   .OptionsAdditionalGhostscriptParameters = Replace$(hLang.Retrieve("AdditionalGhostscriptParameters", .OptionsAdditionalGhostscriptParameters), "/n", vbCrLf)
50050   .OptionsAdditionalGhostscriptSearchpath = Replace$(hLang.Retrieve("AdditionalGhostscriptSearchpath", .OptionsAdditionalGhostscriptSearchpath), "/n", vbCrLf)
50060   .OptionsAddWindowsFontpath = Replace$(hLang.Retrieve("AddWindowsFontpath", .OptionsAddWindowsFontpath), "/n", vbCrLf)
50070   .OptionsAssociatePSFiles = Replace$(hLang.Retrieve("AssociatePSFiles", .OptionsAssociatePSFiles), "/n", vbCrLf)
50080   .OptionsAutosaveDirectoryPrompt = Replace$(hLang.Retrieve("AutosaveDirectoryPrompt", .OptionsAutosaveDirectoryPrompt), "/n", vbCrLf)
50090   .OptionsAutosaveFilename = Replace$(hLang.Retrieve("AutosaveFilename", .OptionsAutosaveFilename), "/n", vbCrLf)
50100   .OptionsAutosaveFilenameTokens = Replace$(hLang.Retrieve("AutosaveFilenameTokens", .OptionsAutosaveFilenameTokens), "/n", vbCrLf)
50110   .OptionsAutosaveFormat = Replace$(hLang.Retrieve("AutosaveFormat", .OptionsAutosaveFormat), "/n", vbCrLf)
50120   .OptionsAutosaveStartStandardProgram = Replace$(hLang.Retrieve("AutosaveStartStandardProgram", .OptionsAutosaveStartStandardProgram), "/n", vbCrLf)
50130   .OptionsBitmapResolution = Replace$(hLang.Retrieve("BitmapResolution", .OptionsBitmapResolution), "/n", vbCrLf)
50140   .OptionsBMPColorscount01 = Replace$(hLang.Retrieve("BMPColorscount01", .OptionsBMPColorscount01), "/n", vbCrLf)
50150   .OptionsBMPColorscount02 = Replace$(hLang.Retrieve("BMPColorscount02", .OptionsBMPColorscount02), "/n", vbCrLf)
50160   .OptionsBMPColorscount03 = Replace$(hLang.Retrieve("BMPColorscount03", .OptionsBMPColorscount03), "/n", vbCrLf)
50170   .OptionsBMPColorscount04 = Replace$(hLang.Retrieve("BMPColorscount04", .OptionsBMPColorscount04), "/n", vbCrLf)
50180   .OptionsBMPColorscount05 = Replace$(hLang.Retrieve("BMPColorscount05", .OptionsBMPColorscount05), "/n", vbCrLf)
50190   .OptionsBMPColorscount06 = Replace$(hLang.Retrieve("BMPColorscount06", .OptionsBMPColorscount06), "/n", vbCrLf)
50200   .OptionsBMPColorscount07 = Replace$(hLang.Retrieve("BMPColorscount07", .OptionsBMPColorscount07), "/n", vbCrLf)
50210   .OptionsBMPDescription = Replace$(hLang.Retrieve("BMPDescription", .OptionsBMPDescription), "/n", vbCrLf)
50220   .OptionsBMPSymbol = Replace$(hLang.Retrieve("BMPSymbol", .OptionsBMPSymbol), "/n", vbCrLf)
50230   .OptionsCancel = Replace$(hLang.Retrieve("Cancel", .OptionsCancel), "/n", vbCrLf)
50240   .OptionsCustomPapersizeHeight = Replace$(hLang.Retrieve("CustomPapersizeHeight", .OptionsCustomPapersizeHeight), "/n", vbCrLf)
50250   .OptionsCustomPapersizeInfo = Replace$(hLang.Retrieve("CustomPapersizeInfo", .OptionsCustomPapersizeInfo), "/n", vbCrLf)
50260   .OptionsCustomPapersizeWidth = Replace$(hLang.Retrieve("CustomPapersizeWidth", .OptionsCustomPapersizeWidth), "/n", vbCrLf)
50270   .OptionsDirectoriesGSBin = Replace$(hLang.Retrieve("DirectoriesGSBin", .OptionsDirectoriesGSBin), "/n", vbCrLf)
50280   .OptionsDirectoriesGSFonts = Replace$(hLang.Retrieve("DirectoriesGSFonts", .OptionsDirectoriesGSFonts), "/n", vbCrLf)
50290   .OptionsDirectoriesGSLibraries = Replace$(hLang.Retrieve("DirectoriesGSLibraries", .OptionsDirectoriesGSLibraries), "/n", vbCrLf)
50300   .OptionsDirectoriesTempPath = Replace$(hLang.Retrieve("DirectoriesTempPath", .OptionsDirectoriesTempPath), "/n", vbCrLf)
50310   .OptionsDocument = Replace$(hLang.Retrieve("Document", .OptionsDocument), "/n", vbCrLf)
50320   .OptionsEPSDescription = Replace$(hLang.Retrieve("EPSDescription", .OptionsEPSDescription), "/n", vbCrLf)
50330   .OptionsEPSFiles = Replace$(hLang.Retrieve("EPSFiles", .OptionsEPSFiles), "/n", vbCrLf)
50340   .OptionsEPSSymbol = Replace$(hLang.Retrieve("EPSSymbol", .OptionsEPSSymbol), "/n", vbCrLf)
50350   .OptionsGhostscriptBinariesDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptBinariesDirectoryPrompt", .OptionsGhostscriptBinariesDirectoryPrompt), "/n", vbCrLf)
50360   .OptionsGhostscriptFontsDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptFontsDirectoryPrompt", .OptionsGhostscriptFontsDirectoryPrompt), "/n", vbCrLf)
50370   .OptionsGhostscriptInternal = Replace$(hLang.Retrieve("GhostscriptInternal", .OptionsGhostscriptInternal), "/n", vbCrLf)
50380   .OptionsGhostscriptLibrariesDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptLibrariesDirectoryPrompt", .OptionsGhostscriptLibrariesDirectoryPrompt), "/n", vbCrLf)
50390   .OptionsGhostscriptResourceDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptResourceDirectoryPrompt", .OptionsGhostscriptResourceDirectoryPrompt), "/n", vbCrLf)
50400   .OptionsGhostscriptversion = Replace$(hLang.Retrieve("Ghostscriptversion", .OptionsGhostscriptversion), "/n", vbCrLf)
50410   .OptionsImageSettings = Replace$(hLang.Retrieve("ImageSettings", .OptionsImageSettings), "/n", vbCrLf)
50420   .OptionsJavaPath = Replace$(hLang.Retrieve("JavaPath", .OptionsJavaPath), "/n", vbCrLf)
50430   .OptionsJPEGColorscount01 = Replace$(hLang.Retrieve("JPEGColorscount01", .OptionsJPEGColorscount01), "/n", vbCrLf)
50440   .OptionsJPEGColorscount02 = Replace$(hLang.Retrieve("JPEGColorscount02", .OptionsJPEGColorscount02), "/n", vbCrLf)
50450   .OptionsJPEGDescription = Replace$(hLang.Retrieve("JPEGDescription", .OptionsJPEGDescription), "/n", vbCrLf)
50460   .OptionsJPEGQuality = Replace$(hLang.Retrieve("JPEGQuality", .OptionsJPEGQuality), "/n", vbCrLf)
50470   .OptionsJPEGSymbol = Replace$(hLang.Retrieve("JPEGSymbol", .OptionsJPEGSymbol), "/n", vbCrLf)
50480   .OptionsLanguagesCurrentLanguage = Replace$(hLang.Retrieve("LanguagesCurrentLanguage", .OptionsLanguagesCurrentLanguage), "/n", vbCrLf)
50490   .OptionsLanguagesDownloadMoreLanguages = Replace$(hLang.Retrieve("LanguagesDownloadMoreLanguages", .OptionsLanguagesDownloadMoreLanguages), "/n", vbCrLf)
50500   .OptionsLanguagesInstall = Replace$(hLang.Retrieve("LanguagesInstall", .OptionsLanguagesInstall), "/n", vbCrLf)
50510   .OptionsLanguagesRefresh = Replace$(hLang.Retrieve("LanguagesRefresh", .OptionsLanguagesRefresh), "/n", vbCrLf)
50520   .OptionsLanguagesTranslation = Replace$(hLang.Retrieve("LanguagesTranslation", .OptionsLanguagesTranslation), "/n", vbCrLf)
50530   .OptionsLanguagesVersion = Replace$(hLang.Retrieve("LanguagesVersion", .OptionsLanguagesVersion), "/n", vbCrLf)
50540   .OptionsOnePagePerFile = Replace$(hLang.Retrieve("OnePagePerFile", .OptionsOnePagePerFile), "/n", vbCrLf)
50550   .OptionsOwnerPass = Replace$(hLang.Retrieve("OwnerPass", .OptionsOwnerPass), "/n", vbCrLf)
50560   .OptionsPassCancel = Replace$(hLang.Retrieve("PassCancel", .OptionsPassCancel), "/n", vbCrLf)
50570   .OptionsPassOK = Replace$(hLang.Retrieve("PassOK", .OptionsPassOK), "/n", vbCrLf)
50580   .OptionsPCXColorscount01 = Replace$(hLang.Retrieve("PCXColorscount01", .OptionsPCXColorscount01), "/n", vbCrLf)
50590   .OptionsPCXColorscount02 = Replace$(hLang.Retrieve("PCXColorscount02", .OptionsPCXColorscount02), "/n", vbCrLf)
50600   .OptionsPCXColorscount03 = Replace$(hLang.Retrieve("PCXColorscount03", .OptionsPCXColorscount03), "/n", vbCrLf)
50610   .OptionsPCXColorscount04 = Replace$(hLang.Retrieve("PCXColorscount04", .OptionsPCXColorscount04), "/n", vbCrLf)
50620   .OptionsPCXColorscount05 = Replace$(hLang.Retrieve("PCXColorscount05", .OptionsPCXColorscount05), "/n", vbCrLf)
50630   .OptionsPCXColorscount06 = Replace$(hLang.Retrieve("PCXColorscount06", .OptionsPCXColorscount06), "/n", vbCrLf)
50640   .OptionsPCXDescription = Replace$(hLang.Retrieve("PCXDescription", .OptionsPCXDescription), "/n", vbCrLf)
50650   .OptionsPCXSymbol = Replace$(hLang.Retrieve("PCXSymbol", .OptionsPCXSymbol), "/n", vbCrLf)
50660   .OptionsPDFAllowAssembly = Replace$(hLang.Retrieve("PDFAllowAssembly", .OptionsPDFAllowAssembly), "/n", vbCrLf)
50670   .OptionsPDFAllowDegradedPrinting = Replace$(hLang.Retrieve("PDFAllowDegradedPrinting", .OptionsPDFAllowDegradedPrinting), "/n", vbCrLf)
50680   .OptionsPDFAllowFillIn = Replace$(hLang.Retrieve("PDFAllowFillIn", .OptionsPDFAllowFillIn), "/n", vbCrLf)
50690   .OptionsPDFAllowScreenReaders = Replace$(hLang.Retrieve("PDFAllowScreenReaders", .OptionsPDFAllowScreenReaders), "/n", vbCrLf)
50700   .OptionsPDFColors = Replace$(hLang.Retrieve("PDFColors", .OptionsPDFColors), "/n", vbCrLf)
50710   .OptionsPDFColorsCaption = Replace$(hLang.Retrieve("PDFColorsCaption", .OptionsPDFColorsCaption), "/n", vbCrLf)
50720   .OptionsPDFColorsCMYKtoRGB = Replace$(hLang.Retrieve("PDFColorsCMYKtoRGB", .OptionsPDFColorsCMYKtoRGB), "/n", vbCrLf)
50730   .OptionsPDFColorsColorModel01 = Replace$(hLang.Retrieve("PDFColorsColorModel01", .OptionsPDFColorsColorModel01), "/n", vbCrLf)
50740   .OptionsPDFColorsColorModel02 = Replace$(hLang.Retrieve("PDFColorsColorModel02", .OptionsPDFColorsColorModel02), "/n", vbCrLf)
50750   .OptionsPDFColorsColorModel03 = Replace$(hLang.Retrieve("PDFColorsColorModel03", .OptionsPDFColorsColorModel03), "/n", vbCrLf)
50760   .OptionsPDFColorsColorOptions = Replace$(hLang.Retrieve("PDFColorsColorOptions", .OptionsPDFColorsColorOptions), "/n", vbCrLf)
50770   .OptionsPDFColorsPreserveHalftone = Replace$(hLang.Retrieve("PDFColorsPreserveHalftone", .OptionsPDFColorsPreserveHalftone), "/n", vbCrLf)
50780   .OptionsPDFColorsPreserveOverprint = Replace$(hLang.Retrieve("PDFColorsPreserveOverprint", .OptionsPDFColorsPreserveOverprint), "/n", vbCrLf)
50790   .OptionsPDFColorsPreserveTransfer = Replace$(hLang.Retrieve("PDFColorsPreserveTransfer", .OptionsPDFColorsPreserveTransfer), "/n", vbCrLf)
50800   .OptionsPDFCompression = Replace$(hLang.Retrieve("PDFCompression", .OptionsPDFCompression), "/n", vbCrLf)
50810   .OptionsPDFCompressionCaption = Replace$(hLang.Retrieve("PDFCompressionCaption", .OptionsPDFCompressionCaption), "/n", vbCrLf)
50820   .OptionsPDFCompressionColor = Replace$(hLang.Retrieve("PDFCompressionColor", .OptionsPDFCompressionColor), "/n", vbCrLf)
50830   .OptionsPDFCompressionColorComp = Replace$(hLang.Retrieve("PDFCompressionColorComp", .OptionsPDFCompressionColorComp), "/n", vbCrLf)
50840   .OptionsPDFCompressionColorComp01 = Replace$(hLang.Retrieve("PDFCompressionColorComp01", .OptionsPDFCompressionColorComp01), "/n", vbCrLf)
50850   .OptionsPDFCompressionColorComp02 = Replace$(hLang.Retrieve("PDFCompressionColorComp02", .OptionsPDFCompressionColorComp02), "/n", vbCrLf)
50860   .OptionsPDFCompressionColorComp03 = Replace$(hLang.Retrieve("PDFCompressionColorComp03", .OptionsPDFCompressionColorComp03), "/n", vbCrLf)
50870   .OptionsPDFCompressionColorComp04 = Replace$(hLang.Retrieve("PDFCompressionColorComp04", .OptionsPDFCompressionColorComp04), "/n", vbCrLf)
50880   .OptionsPDFCompressionColorComp05 = Replace$(hLang.Retrieve("PDFCompressionColorComp05", .OptionsPDFCompressionColorComp05), "/n", vbCrLf)
50890   .OptionsPDFCompressionColorComp06 = Replace$(hLang.Retrieve("PDFCompressionColorComp06", .OptionsPDFCompressionColorComp06), "/n", vbCrLf)
50900   .OptionsPDFCompressionColorComp07 = Replace$(hLang.Retrieve("PDFCompressionColorComp07", .OptionsPDFCompressionColorComp07), "/n", vbCrLf)
50910   .OptionsPDFCompressionColorComp08 = Replace$(hLang.Retrieve("PDFCompressionColorComp08", .OptionsPDFCompressionColorComp08), "/n", vbCrLf)
50920   .OptionsPDFCompressionColorRes = Replace$(hLang.Retrieve("PDFCompressionColorRes", .OptionsPDFCompressionColorRes), "/n", vbCrLf)
50930   .OptionsPDFCompressionColorResample = Replace$(hLang.Retrieve("PDFCompressionColorResample", .OptionsPDFCompressionColorResample), "/n", vbCrLf)
50940   .OptionsPDFCompressionColorResample01 = Replace$(hLang.Retrieve("PDFCompressionColorResample01", .OptionsPDFCompressionColorResample01), "/n", vbCrLf)
50950   .OptionsPDFCompressionColorResample02 = Replace$(hLang.Retrieve("PDFCompressionColorResample02", .OptionsPDFCompressionColorResample02), "/n", vbCrLf)
50960   .OptionsPDFCompressionColorResample03 = Replace$(hLang.Retrieve("PDFCompressionColorResample03", .OptionsPDFCompressionColorResample03), "/n", vbCrLf)
50970   .OptionsPDFCompressionGrey = Replace$(hLang.Retrieve("PDFCompressionGrey", .OptionsPDFCompressionGrey), "/n", vbCrLf)
50980   .OptionsPDFCompressionGreyComp = Replace$(hLang.Retrieve("PDFCompressionGreyComp", .OptionsPDFCompressionGreyComp), "/n", vbCrLf)
50990   .OptionsPDFCompressionGreyComp01 = Replace$(hLang.Retrieve("PDFCompressionGreyComp01", .OptionsPDFCompressionGreyComp01), "/n", vbCrLf)
51000   .OptionsPDFCompressionGreyComp02 = Replace$(hLang.Retrieve("PDFCompressionGreyComp02", .OptionsPDFCompressionGreyComp02), "/n", vbCrLf)
51010   .OptionsPDFCompressionGreyComp03 = Replace$(hLang.Retrieve("PDFCompressionGreyComp03", .OptionsPDFCompressionGreyComp03), "/n", vbCrLf)
51020   .OptionsPDFCompressionGreyComp04 = Replace$(hLang.Retrieve("PDFCompressionGreyComp04", .OptionsPDFCompressionGreyComp04), "/n", vbCrLf)
51030   .OptionsPDFCompressionGreyComp05 = Replace$(hLang.Retrieve("PDFCompressionGreyComp05", .OptionsPDFCompressionGreyComp05), "/n", vbCrLf)
51040   .OptionsPDFCompressionGreyComp06 = Replace$(hLang.Retrieve("PDFCompressionGreyComp06", .OptionsPDFCompressionGreyComp06), "/n", vbCrLf)
51050   .OptionsPDFCompressionGreyComp07 = Replace$(hLang.Retrieve("PDFCompressionGreyComp07", .OptionsPDFCompressionGreyComp07), "/n", vbCrLf)
51060   .OptionsPDFCompressionGreyComp08 = Replace$(hLang.Retrieve("PDFCompressionGreyComp08", .OptionsPDFCompressionGreyComp08), "/n", vbCrLf)
51070   .OptionsPDFCompressionGreyRes = Replace$(hLang.Retrieve("PDFCompressionGreyRes", .OptionsPDFCompressionGreyRes), "/n", vbCrLf)
51080   .OptionsPDFCompressionGreyResample = Replace$(hLang.Retrieve("PDFCompressionGreyResample", .OptionsPDFCompressionGreyResample), "/n", vbCrLf)
51090   .OptionsPDFCompressionGreyResample01 = Replace$(hLang.Retrieve("PDFCompressionGreyResample01", .OptionsPDFCompressionGreyResample01), "/n", vbCrLf)
51100   .OptionsPDFCompressionGreyResample02 = Replace$(hLang.Retrieve("PDFCompressionGreyResample02", .OptionsPDFCompressionGreyResample02), "/n", vbCrLf)
51110   .OptionsPDFCompressionGreyResample03 = Replace$(hLang.Retrieve("PDFCompressionGreyResample03", .OptionsPDFCompressionGreyResample03), "/n", vbCrLf)
51120   .OptionsPDFCompressionMono = Replace$(hLang.Retrieve("PDFCompressionMono", .OptionsPDFCompressionMono), "/n", vbCrLf)
51130   .OptionsPDFCompressionMonoComp = Replace$(hLang.Retrieve("PDFCompressionMonoComp", .OptionsPDFCompressionMonoComp), "/n", vbCrLf)
51140   .OptionsPDFCompressionMonoComp01 = Replace$(hLang.Retrieve("PDFCompressionMonoComp01", .OptionsPDFCompressionMonoComp01), "/n", vbCrLf)
51150   .OptionsPDFCompressionMonoComp02 = Replace$(hLang.Retrieve("PDFCompressionMonoComp02", .OptionsPDFCompressionMonoComp02), "/n", vbCrLf)
51160   .OptionsPDFCompressionMonoComp03 = Replace$(hLang.Retrieve("PDFCompressionMonoComp03", .OptionsPDFCompressionMonoComp03), "/n", vbCrLf)
51170   .OptionsPDFCompressionMonoComp04 = Replace$(hLang.Retrieve("PDFCompressionMonoComp04", .OptionsPDFCompressionMonoComp04), "/n", vbCrLf)
51180   .OptionsPDFCompressionMonoRes = Replace$(hLang.Retrieve("PDFCompressionMonoRes", .OptionsPDFCompressionMonoRes), "/n", vbCrLf)
51190   .OptionsPDFCompressionMonoResample = Replace$(hLang.Retrieve("PDFCompressionMonoResample", .OptionsPDFCompressionMonoResample), "/n", vbCrLf)
51200   .OptionsPDFCompressionMonoResample01 = Replace$(hLang.Retrieve("PDFCompressionMonoResample01", .OptionsPDFCompressionMonoResample01), "/n", vbCrLf)
51210   .OptionsPDFCompressionMonoResample02 = Replace$(hLang.Retrieve("PDFCompressionMonoResample02", .OptionsPDFCompressionMonoResample02), "/n", vbCrLf)
51220   .OptionsPDFCompressionMonoResample03 = Replace$(hLang.Retrieve("PDFCompressionMonoResample03", .OptionsPDFCompressionMonoResample03), "/n", vbCrLf)
51230   .OptionsPDFCompressionTextComp = Replace$(hLang.Retrieve("PDFCompressionTextComp", .OptionsPDFCompressionTextComp), "/n", vbCrLf)
51240   .OptionsPDFDescription = Replace$(hLang.Retrieve("PDFDescription", .OptionsPDFDescription), "/n", vbCrLf)
51250   .OptionsPDFDisallowCopy = Replace$(hLang.Retrieve("PDFDisallowCopy", .OptionsPDFDisallowCopy), "/n", vbCrLf)
51260   .OptionsPDFDisallowModify = Replace$(hLang.Retrieve("PDFDisallowModify", .OptionsPDFDisallowModify), "/n", vbCrLf)
51270   .OptionsPDFDisallowModifyComments = Replace$(hLang.Retrieve("PDFDisallowModifyComments", .OptionsPDFDisallowModifyComments), "/n", vbCrLf)
51280   .OptionsPDFDisallowPrint = Replace$(hLang.Retrieve("PDFDisallowPrint", .OptionsPDFDisallowPrint), "/n", vbCrLf)
51290   .OptionsPDFDisallowUser = Replace$(hLang.Retrieve("PDFDisallowUser", .OptionsPDFDisallowUser), "/n", vbCrLf)
51300   .OptionsPDFEncryptionHigh = Replace$(hLang.Retrieve("PDFEncryptionHigh", .OptionsPDFEncryptionHigh), "/n", vbCrLf)
51310   .OptionsPDFEncryptionLevel = Replace$(hLang.Retrieve("PDFEncryptionLevel", .OptionsPDFEncryptionLevel), "/n", vbCrLf)
51320   .OptionsPDFEncryptionLow = Replace$(hLang.Retrieve("PDFEncryptionLow", .OptionsPDFEncryptionLow), "/n", vbCrLf)
51330   .OptionsPDFEncryptor = Replace$(hLang.Retrieve("PDFEncryptor", .OptionsPDFEncryptor), "/n", vbCrLf)
51340   .OptionsPDFEnhancedPermissions = Replace$(hLang.Retrieve("PDFEnhancedPermissions", .OptionsPDFEnhancedPermissions), "/n", vbCrLf)
51350   .OptionsPDFEnterPasswords = Replace$(hLang.Retrieve("PDFEnterPasswords", .OptionsPDFEnterPasswords), "/n", vbCrLf)
51360   .OptionsPDFFonts = Replace$(hLang.Retrieve("PDFFonts", .OptionsPDFFonts), "/n", vbCrLf)
51370   .OptionsPDFFontsCaption = Replace$(hLang.Retrieve("PDFFontsCaption", .OptionsPDFFontsCaption), "/n", vbCrLf)
51380   .OptionsPDFFontsEmbedAll = Replace$(hLang.Retrieve("PDFFontsEmbedAll", .OptionsPDFFontsEmbedAll), "/n", vbCrLf)
51390   .OptionsPDFFontsSubSetFonts = Replace$(hLang.Retrieve("PDFFontsSubSetFonts", .OptionsPDFFontsSubSetFonts), "/n", vbCrLf)
51400   .OptionsPDFGeneral = Replace$(hLang.Retrieve("PDFGeneral", .OptionsPDFGeneral), "/n", vbCrLf)
51410   .OptionsPDFGeneralASCII85 = Replace$(hLang.Retrieve("PDFGeneralASCII85", .OptionsPDFGeneralASCII85), "/n", vbCrLf)
51420   .OptionsPDFGeneralAutorotate = Replace$(hLang.Retrieve("PDFGeneralAutorotate", .OptionsPDFGeneralAutorotate), "/n", vbCrLf)
51430   .OptionsPDFGeneralCaption = Replace$(hLang.Retrieve("PDFGeneralCaption", .OptionsPDFGeneralCaption), "/n", vbCrLf)
51440   .OptionsPDFGeneralCompatibility = Replace$(hLang.Retrieve("PDFGeneralCompatibility", .OptionsPDFGeneralCompatibility), "/n", vbCrLf)
51450   .OptionsPDFGeneralCompatibility01 = Replace$(hLang.Retrieve("PDFGeneralCompatibility01", .OptionsPDFGeneralCompatibility01), "/n", vbCrLf)
51460   .OptionsPDFGeneralCompatibility02 = Replace$(hLang.Retrieve("PDFGeneralCompatibility02", .OptionsPDFGeneralCompatibility02), "/n", vbCrLf)
51470   .OptionsPDFGeneralCompatibility03 = Replace$(hLang.Retrieve("PDFGeneralCompatibility03", .OptionsPDFGeneralCompatibility03), "/n", vbCrLf)
51480   .OptionsPDFGeneralOverprint = Replace$(hLang.Retrieve("PDFGeneralOverprint", .OptionsPDFGeneralOverprint), "/n", vbCrLf)
51490   .OptionsPDFGeneralOverprint01 = Replace$(hLang.Retrieve("PDFGeneralOverprint01", .OptionsPDFGeneralOverprint01), "/n", vbCrLf)
51500   .OptionsPDFGeneralOverprint02 = Replace$(hLang.Retrieve("PDFGeneralOverprint02", .OptionsPDFGeneralOverprint02), "/n", vbCrLf)
51510   .OptionsPDFGeneralResolution = Replace$(hLang.Retrieve("PDFGeneralResolution", .OptionsPDFGeneralResolution), "/n", vbCrLf)
51520   .OptionsPDFGeneralRotate01 = Replace$(hLang.Retrieve("PDFGeneralRotate01", .OptionsPDFGeneralRotate01), "/n", vbCrLf)
51530   .OptionsPDFGeneralRotate02 = Replace$(hLang.Retrieve("PDFGeneralRotate02", .OptionsPDFGeneralRotate02), "/n", vbCrLf)
51540   .OptionsPDFGeneralRotate03 = Replace$(hLang.Retrieve("PDFGeneralRotate03", .OptionsPDFGeneralRotate03), "/n", vbCrLf)
51550   .OptionsPDFOptimize = Replace$(hLang.Retrieve("PDFOptimize", .OptionsPDFOptimize), "/n", vbCrLf)
51560   .OptionsPDFOptions = Replace$(hLang.Retrieve("PDFOptions", .OptionsPDFOptions), "/n", vbCrLf)
51570   .OptionsPDFOwnerPass = Replace$(hLang.Retrieve("PDFOwnerPass", .OptionsPDFOwnerPass), "/n", vbCrLf)
51580   .OptionsPDFPasswords = Replace$(hLang.Retrieve("PDFPasswords", .OptionsPDFPasswords), "/n", vbCrLf)
51590   .OptionsPDFRepeatPassword = Replace$(hLang.Retrieve("PDFRepeatPassword", .OptionsPDFRepeatPassword), "/n", vbCrLf)
51600   .OptionsPDFSecurity = Replace$(hLang.Retrieve("PDFSecurity", .OptionsPDFSecurity), "/n", vbCrLf)
51610   .OptionsPDFSecurityCaption = Replace$(hLang.Retrieve("PDFSecurityCaption", .OptionsPDFSecurityCaption), "/n", vbCrLf)
51620   .OptionsPDFSetPassword = Replace$(hLang.Retrieve("PDFSetPassword", .OptionsPDFSetPassword), "/n", vbCrLf)
51630   .OptionsPDFSymbol = Replace$(hLang.Retrieve("PDFSymbol", .OptionsPDFSymbol), "/n", vbCrLf)
51640   .OptionsPDFUserPass = Replace$(hLang.Retrieve("PDFUserPass", .OptionsPDFUserPass), "/n", vbCrLf)
51650   .OptionsPDFUseSecurity = Replace$(hLang.Retrieve("PDFUseSecurity", .OptionsPDFUseSecurity), "/n", vbCrLf)
51660   .OptionsPNGColorscount01 = Replace$(hLang.Retrieve("PNGColorscount01", .OptionsPNGColorscount01), "/n", vbCrLf)
51670   .OptionsPNGColorscount02 = Replace$(hLang.Retrieve("PNGColorscount02", .OptionsPNGColorscount02), "/n", vbCrLf)
51680   .OptionsPNGColorscount03 = Replace$(hLang.Retrieve("PNGColorscount03", .OptionsPNGColorscount03), "/n", vbCrLf)
51690   .OptionsPNGColorscount04 = Replace$(hLang.Retrieve("PNGColorscount04", .OptionsPNGColorscount04), "/n", vbCrLf)
51700   .OptionsPNGColorscount05 = Replace$(hLang.Retrieve("PNGColorscount05", .OptionsPNGColorscount05), "/n", vbCrLf)
51710   .OptionsPNGDescription = Replace$(hLang.Retrieve("PNGDescription", .OptionsPNGDescription), "/n", vbCrLf)
51720   .OptionsPNGFiles = Replace$(hLang.Retrieve("PNGFiles", .OptionsPNGFiles), "/n", vbCrLf)
51730   .OptionsPNGSymbol = Replace$(hLang.Retrieve("PNGSymbol", .OptionsPNGSymbol), "/n", vbCrLf)
51740   .OptionsPrintAfterSaving = Replace$(hLang.Retrieve("PrintAfterSaving", .OptionsPrintAfterSaving), "/n", vbCrLf)
51750   .OptionsPrintAfterSavingDuplex = Replace$(hLang.Retrieve("PrintAfterSavingDuplex", .OptionsPrintAfterSavingDuplex), "/n", vbCrLf)
51760   .OptionsPrintAfterSavingDuplexTumbleOff = Replace$(hLang.Retrieve("PrintAfterSavingDuplexTumbleOff", .OptionsPrintAfterSavingDuplexTumbleOff), "/n", vbCrLf)
51770   .OptionsPrintAfterSavingDuplexTumbleOn = Replace$(hLang.Retrieve("PrintAfterSavingDuplexTumbleOn", .OptionsPrintAfterSavingDuplexTumbleOn), "/n", vbCrLf)
51780   .OptionsPrintAfterSavingNoCancel = Replace$(hLang.Retrieve("PrintAfterSavingNoCancel", .OptionsPrintAfterSavingNoCancel), "/n", vbCrLf)
51790   .OptionsPrintAfterSavingPrinter = Replace$(hLang.Retrieve("PrintAfterSavingPrinter", .OptionsPrintAfterSavingPrinter), "/n", vbCrLf)
51800   .OptionsPrintAfterSavingQueryUser = Replace$(hLang.Retrieve("PrintAfterSavingQueryUser", .OptionsPrintAfterSavingQueryUser), "/n", vbCrLf)
51810   .OptionsPrintAfterSavingQueryUserDefaultPrinter = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserDefaultPrinter", .OptionsPrintAfterSavingQueryUserDefaultPrinter), "/n", vbCrLf)
51820   .OptionsPrintAfterSavingQueryUserOff = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserOff", .OptionsPrintAfterSavingQueryUserOff), "/n", vbCrLf)
51830   .OptionsPrintAfterSavingQueryUserPrinterSetupDialog = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserPrinterSetupDialog", .OptionsPrintAfterSavingQueryUserPrinterSetupDialog), "/n", vbCrLf)
51840   .OptionsPrintAfterSavingQueryUserStandardPrinterDialog = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserStandardPrinterDialog", .OptionsPrintAfterSavingQueryUserStandardPrinterDialog), "/n", vbCrLf)
51850   .OptionsPrintertempDirectoryPrompt = Replace$(hLang.Retrieve("PrintertempDirectoryPrompt", .OptionsPrintertempDirectoryPrompt), "/n", vbCrLf)
51860   .OptionsPrintTestpage = Replace$(hLang.Retrieve("PrintTestpage", .OptionsPrintTestpage), "/n", vbCrLf)
51870   .OptionsProcesspriority = Replace$(hLang.Retrieve("Processpriority", .OptionsProcesspriority), "/n", vbCrLf)
51880   .OptionsProcesspriorityHigh = Replace$(hLang.Retrieve("ProcesspriorityHigh", .OptionsProcesspriorityHigh), "/n", vbCrLf)
51890   .OptionsProcesspriorityIdle = Replace$(hLang.Retrieve("ProcesspriorityIdle", .OptionsProcesspriorityIdle), "/n", vbCrLf)
51900   .OptionsProcesspriorityNormal = Replace$(hLang.Retrieve("ProcesspriorityNormal", .OptionsProcesspriorityNormal), "/n", vbCrLf)
51910   .OptionsProcesspriorityRealtime = Replace$(hLang.Retrieve("ProcesspriorityRealtime", .OptionsProcesspriorityRealtime), "/n", vbCrLf)
51920   .OptionsProgramActionsDescription = Replace$(hLang.Retrieve("ProgramActionsDescription", .OptionsProgramActionsDescription), "/n", vbCrLf)
51930   .OptionsProgramActionsSymbol = Replace$(hLang.Retrieve("ProgramActionsSymbol", .OptionsProgramActionsSymbol), "/n", vbCrLf)
51940   .OptionsProgramAutosaveDescription = Replace$(hLang.Retrieve("ProgramAutosaveDescription", .OptionsProgramAutosaveDescription), "/n", vbCrLf)
51950   .OptionsProgramAutosaveSymbol = Replace$(hLang.Retrieve("ProgramAutosaveSymbol", .OptionsProgramAutosaveSymbol), "/n", vbCrLf)
51960   .OptionsProgramDirectoriesDescription = Replace$(hLang.Retrieve("ProgramDirectoriesDescription", .OptionsProgramDirectoriesDescription), "/n", vbCrLf)
51970   .OptionsProgramDirectoriesSymbol = Replace$(hLang.Retrieve("ProgramDirectoriesSymbol", .OptionsProgramDirectoriesSymbol), "/n", vbCrLf)
51980   .OptionsProgramDocumentDescription = Replace$(hLang.Retrieve("ProgramDocumentDescription", .OptionsProgramDocumentDescription), "/n", vbCrLf)
51990   .OptionsProgramDocumentDescription1 = Replace$(hLang.Retrieve("ProgramDocumentDescription1", .OptionsProgramDocumentDescription1), "/n", vbCrLf)
52000   .OptionsProgramDocumentDescription2 = Replace$(hLang.Retrieve("ProgramDocumentDescription2", .OptionsProgramDocumentDescription2), "/n", vbCrLf)
52010   .OptionsProgramDocumentSymbol = Replace$(hLang.Retrieve("ProgramDocumentSymbol", .OptionsProgramDocumentSymbol), "/n", vbCrLf)
52020   .OptionsProgramFont = Replace$(hLang.Retrieve("ProgramFont", .OptionsProgramFont), "/n", vbCrLf)
52030   .OptionsProgramFontCancelTest = Replace$(hLang.Retrieve("ProgramFontCancelTest", .OptionsProgramFontCancelTest), "/n", vbCrLf)
52040   .OptionsProgramFontcharset = Replace$(hLang.Retrieve("ProgramFontcharset", .OptionsProgramFontcharset), "/n", vbCrLf)
52050   .OptionsProgramFontDescription = Replace$(hLang.Retrieve("ProgramFontDescription", .OptionsProgramFontDescription), "/n", vbCrLf)
52060   .OptionsProgramFontSize = Replace$(hLang.Retrieve("ProgramFontSize", .OptionsProgramFontSize), "/n", vbCrLf)
52070   .OptionsProgramFontSymbol = Replace$(hLang.Retrieve("ProgramFontSymbol", .OptionsProgramFontSymbol), "/n", vbCrLf)
52080   .OptionsProgramFontTest = Replace$(hLang.Retrieve("ProgramFontTest", .OptionsProgramFontTest), "/n", vbCrLf)
52090   .OptionsProgramFontTestdescription = Replace$(hLang.Retrieve("ProgramFontTestdescription", .OptionsProgramFontTestdescription), "/n", vbCrLf)
52100   .OptionsProgramGeneralDescription = Replace$(hLang.Retrieve("ProgramGeneralDescription", .OptionsProgramGeneralDescription), "/n", vbCrLf)
52110   .OptionsProgramGeneralDescription1 = Replace$(hLang.Retrieve("ProgramGeneralDescription1", .OptionsProgramGeneralDescription1), "/n", vbCrLf)
52120   .OptionsProgramGeneralDescription2 = Replace$(hLang.Retrieve("ProgramGeneralDescription2", .OptionsProgramGeneralDescription2), "/n", vbCrLf)
52130   .OptionsProgramGeneralSymbol = Replace$(hLang.Retrieve("ProgramGeneralSymbol", .OptionsProgramGeneralSymbol), "/n", vbCrLf)
52140   .OptionsProgramGhostscriptDescription = Replace$(hLang.Retrieve("ProgramGhostscriptDescription", .OptionsProgramGhostscriptDescription), "/n", vbCrLf)
52150   .OptionsProgramGhostscriptSymbol = Replace$(hLang.Retrieve("ProgramGhostscriptSymbol", .OptionsProgramGhostscriptSymbol), "/n", vbCrLf)
52160   .OptionsProgramLanguagesDescription = Replace$(hLang.Retrieve("ProgramLanguagesDescription", .OptionsProgramLanguagesDescription), "/n", vbCrLf)
52170   .OptionsProgramLanguagesSymbol = Replace$(hLang.Retrieve("ProgramLanguagesSymbol", .OptionsProgramLanguagesSymbol), "/n", vbCrLf)
52180   .OptionsProgramNoProcessingAtStartup = Replace$(hLang.Retrieve("ProgramNoProcessingAtStartup", .OptionsProgramNoProcessingAtStartup), "/n", vbCrLf)
52190   .OptionsProgramOptionsDesign = Replace$(hLang.Retrieve("ProgramOptionsDesign", .OptionsProgramOptionsDesign), "/n", vbCrLf)
52200   .OptionsProgramOptionsDesignGradient = Replace$(hLang.Retrieve("ProgramOptionsDesignGradient", .OptionsProgramOptionsDesignGradient), "/n", vbCrLf)
52210   .OptionsProgramOptionsDesignSimple = Replace$(hLang.Retrieve("ProgramOptionsDesignSimple", .OptionsProgramOptionsDesignSimple), "/n", vbCrLf)
52220   .OptionsProgramPrintDescription = Replace$(hLang.Retrieve("ProgramPrintDescription", .OptionsProgramPrintDescription), "/n", vbCrLf)
52230   .OptionsProgramPrintSymbol = Replace$(hLang.Retrieve("ProgramPrintSymbol", .OptionsProgramPrintSymbol), "/n", vbCrLf)
52240   .OptionsProgramRunProgramAfterSavingCaption = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingCaption", .OptionsProgramRunProgramAfterSavingCaption), "/n", vbCrLf)
52250   .OptionsProgramRunProgramAfterSavingProgram = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingProgram", .OptionsProgramRunProgramAfterSavingProgram), "/n", vbCrLf)
52260   .OptionsProgramRunProgramAfterSavingProgramParameters = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingProgramParameters", .OptionsProgramRunProgramAfterSavingProgramParameters), "/n", vbCrLf)
52270   .OptionsProgramRunProgramAfterSavingWaitUntilReady = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWaitUntilReady", .OptionsProgramRunProgramAfterSavingWaitUntilReady), "/n", vbCrLf)
52280   .OptionsProgramRunProgramAfterSavingWindowstyle = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyle", .OptionsProgramRunProgramAfterSavingWindowstyle), "/n", vbCrLf)
52290   .OptionsProgramRunProgramAfterSavingWindowstyleHide = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleHide", .OptionsProgramRunProgramAfterSavingWindowstyleHide), "/n", vbCrLf)
52300   .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleMaximizedFocus", .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus), "/n", vbCrLf)
52310   .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleMinimizedFocus", .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus), "/n", vbCrLf)
52320   .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus", .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus), "/n", vbCrLf)
52330   .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleNormalFocus", .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus), "/n", vbCrLf)
52340   .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleNormalNoFocus", .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus), "/n", vbCrLf)
52350   .OptionsProgramRunProgramBeforeSavingCaption = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingCaption", .OptionsProgramRunProgramBeforeSavingCaption), "/n", vbCrLf)
52360   .OptionsProgramRunProgramBeforeSavingProgram = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingProgram", .OptionsProgramRunProgramBeforeSavingProgram), "/n", vbCrLf)
52370   .OptionsProgramRunProgramBeforeSavingProgramParameters = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingProgramParameters", .OptionsProgramRunProgramBeforeSavingProgramParameters), "/n", vbCrLf)
52380   .OptionsProgramRunProgramBeforeSavingWaitUntilReady = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWaitUntilReady", .OptionsProgramRunProgramBeforeSavingWaitUntilReady), "/n", vbCrLf)
52390   .OptionsProgramRunProgramBeforeSavingWindowstyle = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyle", .OptionsProgramRunProgramBeforeSavingWindowstyle), "/n", vbCrLf)
52400   .OptionsProgramRunProgramBeforeSavingWindowstyleHide = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleHide", .OptionsProgramRunProgramBeforeSavingWindowstyleHide), "/n", vbCrLf)
52410   .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleMaximizedFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus), "/n", vbCrLf)
52420   .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleMinimizedFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus), "/n", vbCrLf)
52430   .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus), "/n", vbCrLf)
52440   .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleNormalFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus), "/n", vbCrLf)
52450   .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleNormalNoFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus), "/n", vbCrLf)
52460   .OptionsProgramSaveDescription = Replace$(hLang.Retrieve("ProgramSaveDescription", .OptionsProgramSaveDescription), "/n", vbCrLf)
52470   .OptionsProgramSaveSymbol = Replace$(hLang.Retrieve("ProgramSaveSymbol", .OptionsProgramSaveSymbol), "/n", vbCrLf)
52480   .OptionsProgramShowAnimation = Replace$(hLang.Retrieve("ProgramShowAnimation", .OptionsProgramShowAnimation), "/n", vbCrLf)
52490   .OptionsProgramSwitchingDefaultprinter = Replace$(hLang.Retrieve("ProgramSwitchingDefaultprinter", .OptionsProgramSwitchingDefaultprinter), "/n", vbCrLf)
52500   .OptionsPSDescription = Replace$(hLang.Retrieve("PSDescription", .OptionsPSDescription), "/n", vbCrLf)
52510   .OptionsPSFiles = Replace$(hLang.Retrieve("PSFiles", .OptionsPSFiles), "/n", vbCrLf)
52520   .OptionsPSLanguageLevel = Replace$(hLang.Retrieve("PSLanguageLevel", .OptionsPSLanguageLevel), "/n", vbCrLf)
52530   .OptionsPSSymbol = Replace$(hLang.Retrieve("PSSymbol", .OptionsPSSymbol), "/n", vbCrLf)
52540   .OptionsRemoveSpaces = Replace$(hLang.Retrieve("RemoveSpaces", .OptionsRemoveSpaces), "/n", vbCrLf)
52550   .OptionsReset = Replace$(hLang.Retrieve("Reset", .OptionsReset), "/n", vbCrLf)
52560   .OptionsSave = Replace$(hLang.Retrieve("Save", .OptionsSave), "/n", vbCrLf)
52570   .OptionsSaveFilename = Replace$(hLang.Retrieve("SaveFilename", .OptionsSaveFilename), "/n", vbCrLf)
52580   .OptionsSaveFilenameAdd = Replace$(hLang.Retrieve("SaveFilenameAdd", .OptionsSaveFilenameAdd), "/n", vbCrLf)
52590   .OptionsSaveFilenameChange = Replace$(hLang.Retrieve("SaveFilenameChange", .OptionsSaveFilenameChange), "/n", vbCrLf)
52600   .OptionsSaveFilenameDelete = Replace$(hLang.Retrieve("SaveFilenameDelete", .OptionsSaveFilenameDelete), "/n", vbCrLf)
52610   .OptionsSaveFilenameSubstitutions = Replace$(hLang.Retrieve("SaveFilenameSubstitutions", .OptionsSaveFilenameSubstitutions), "/n", vbCrLf)
52620   .OptionsSaveFilenameSubstitutionsTitle = Replace$(hLang.Retrieve("SaveFilenameSubstitutionsTitle", .OptionsSaveFilenameSubstitutionsTitle), "/n", vbCrLf)
52630   .OptionsSaveFilenameTokens = Replace$(hLang.Retrieve("SaveFilenameTokens", .OptionsSaveFilenameTokens), "/n", vbCrLf)
52640   .OptionsSavePasswords = Replace$(hLang.Retrieve("SavePasswords", .OptionsSavePasswords), "/n", vbCrLf)
52650   .OptionsSendEmailAfterAutosave = Replace$(hLang.Retrieve("SendEmailAfterAutosave", .OptionsSendEmailAfterAutosave), "/n", vbCrLf)
52660   .OptionsSendMailMethod = Replace$(hLang.Retrieve("SendMailMethod", .OptionsSendMailMethod), "/n", vbCrLf)
52670   .OptionsSendMailMethodAutomatic = Replace$(hLang.Retrieve("SendMailMethodAutomatic", .OptionsSendMailMethodAutomatic), "/n", vbCrLf)
52680   .OptionsSendMailMethodMapi = Replace$(hLang.Retrieve("SendMailMethodMapi", .OptionsSendMailMethodMapi), "/n", vbCrLf)
52690   .OptionsSendMailMethodSendmailDLL = Replace$(hLang.Retrieve("SendMailMethodSendmailDLL", .OptionsSendMailMethodSendmailDLL), "/n", vbCrLf)
52700   .OptionsShellIntegration = Replace$(hLang.Retrieve("ShellIntegration", .OptionsShellIntegration), "/n", vbCrLf)
52710   .OptionsShellIntegrationAdd = Replace$(hLang.Retrieve("ShellIntegrationAdd", .OptionsShellIntegrationAdd), "/n", vbCrLf)
52720   .OptionsShellIntegrationCaption = Replace$(hLang.Retrieve("ShellIntegrationCaption", .OptionsShellIntegrationCaption), "/n", vbCrLf)
52730   .OptionsShellIntegrationRemove = Replace$(hLang.Retrieve("ShellIntegrationRemove", .OptionsShellIntegrationRemove), "/n", vbCrLf)
52740   .OptionsStamp = Replace$(hLang.Retrieve("Stamp", .OptionsStamp), "/n", vbCrLf)
52750   .OptionsStampFontColor = Replace$(hLang.Retrieve("StampFontColor", .OptionsStampFontColor), "/n", vbCrLf)
52760   .OptionsStampOutlineFontThickness = Replace$(hLang.Retrieve("StampOutlineFontThickness", .OptionsStampOutlineFontThickness), "/n", vbCrLf)
52770   .OptionsStampString = Replace$(hLang.Retrieve("StampString", .OptionsStampString), "/n", vbCrLf)
52780   .OptionsStampUseOutlineFont = Replace$(hLang.Retrieve("StampUseOutlineFont", .OptionsStampUseOutlineFont), "/n", vbCrLf)
52790   .OptionsStandardAuthorToken = Replace$(hLang.Retrieve("StandardAuthorToken", .OptionsStandardAuthorToken), "/n", vbCrLf)
52800   .OptionsStandardSaveFormat = Replace$(hLang.Retrieve("StandardSaveFormat", .OptionsStandardSaveFormat), "/n", vbCrLf)
52810   .OptionsTestpage = Replace$(hLang.Retrieve("Testpage", .OptionsTestpage), "/n", vbCrLf)
52820   .OptionsTIFFColorscount01 = Replace$(hLang.Retrieve("TIFFColorscount01", .OptionsTIFFColorscount01), "/n", vbCrLf)
52830   .OptionsTIFFColorscount02 = Replace$(hLang.Retrieve("TIFFColorscount02", .OptionsTIFFColorscount02), "/n", vbCrLf)
52840   .OptionsTIFFColorscount03 = Replace$(hLang.Retrieve("TIFFColorscount03", .OptionsTIFFColorscount03), "/n", vbCrLf)
52850   .OptionsTIFFColorscount04 = Replace$(hLang.Retrieve("TIFFColorscount04", .OptionsTIFFColorscount04), "/n", vbCrLf)
52860   .OptionsTIFFColorscount05 = Replace$(hLang.Retrieve("TIFFColorscount05", .OptionsTIFFColorscount05), "/n", vbCrLf)
52870   .OptionsTIFFColorscount06 = Replace$(hLang.Retrieve("TIFFColorscount06", .OptionsTIFFColorscount06), "/n", vbCrLf)
52880   .OptionsTIFFColorscount07 = Replace$(hLang.Retrieve("TIFFColorscount07", .OptionsTIFFColorscount07), "/n", vbCrLf)
52890   .OptionsTIFFColorscount08 = Replace$(hLang.Retrieve("TIFFColorscount08", .OptionsTIFFColorscount08), "/n", vbCrLf)
52900   .OptionsTIFFDescription = Replace$(hLang.Retrieve("TIFFDescription", .OptionsTIFFDescription), "/n", vbCrLf)
52910   .OptionsTIFFSymbol = Replace$(hLang.Retrieve("TIFFSymbol", .OptionsTIFFSymbol), "/n", vbCrLf)
52920   .OptionsTreeFormats = Replace$(hLang.Retrieve("TreeFormats", .OptionsTreeFormats), "/n", vbCrLf)
52930   .OptionsTreeProgram = Replace$(hLang.Retrieve("TreeProgram", .OptionsTreeProgram), "/n", vbCrLf)
52940   .OptionsUseAutosave = Replace$(hLang.Retrieve("UseAutosave", .OptionsUseAutosave), "/n", vbCrLf)
52950   .OptionsUseAutosaveDirectory = Replace$(hLang.Retrieve("UseAutosaveDirectory", .OptionsUseAutosaveDirectory), "/n", vbCrLf)
52960   .OptionsUseCreationDateNow = Replace$(hLang.Retrieve("UseCreationDateNow", .OptionsUseCreationDateNow), "/n", vbCrLf)
52970   .OptionsUseCustomPapersize = Replace$(hLang.Retrieve("UseCustomPapersize", .OptionsUseCustomPapersize), "/n", vbCrLf)
52980   .OptionsUseFixPapersize = Replace$(hLang.Retrieve("UseFixPapersize", .OptionsUseFixPapersize), "/n", vbCrLf)
52990   .OptionsUserPass = Replace$(hLang.Retrieve("UserPass", .OptionsUserPass), "/n", vbCrLf)
53000   .OptionsUseStandardauthor = Replace$(hLang.Retrieve("UseStandardauthor", .OptionsUseStandardauthor), "/n", vbCrLf)
53010  End With
53020  Set hLang = Nothing
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
50060   .PrintingCancel = Replace$(hLang.Retrieve("Cancel", .PrintingCancel), "/n", vbCrLf)
50070   .PrintingCollect = Replace$(hLang.Retrieve("Collect", .PrintingCollect), "/n", vbCrLf)
50080   .PrintingCreationDate = Replace$(hLang.Retrieve("CreationDate", .PrintingCreationDate), "/n", vbCrLf)
50090   .PrintingDocumentTitle = Replace$(hLang.Retrieve("DocumentTitle", .PrintingDocumentTitle), "/n", vbCrLf)
50100   .PrintingEMail = Replace$(hLang.Retrieve("EMail", .PrintingEMail), "/n", vbCrLf)
50110   .PrintingEPSFiles = Replace$(hLang.Retrieve("EPSFiles", .PrintingEPSFiles), "/n", vbCrLf)
50120   .PrintingJPEGFiles = Replace$(hLang.Retrieve("JPEGFiles", .PrintingJPEGFiles), "/n", vbCrLf)
50130   .PrintingKeywords = Replace$(hLang.Retrieve("Keywords", .PrintingKeywords), "/n", vbCrLf)
50140   .PrintingModifyDate = Replace$(hLang.Retrieve("ModifyDate", .PrintingModifyDate), "/n", vbCrLf)
50150   .PrintingNow = Replace$(hLang.Retrieve("Now", .PrintingNow), "/n", vbCrLf)
50160   .PrintingPCXFiles = Replace$(hLang.Retrieve("PCXFiles", .PrintingPCXFiles), "/n", vbCrLf)
50170   .PrintingPDFFiles = Replace$(hLang.Retrieve("PDFFiles", .PrintingPDFFiles), "/n", vbCrLf)
50180   .PrintingPNGFiles = Replace$(hLang.Retrieve("PNGFiles", .PrintingPNGFiles), "/n", vbCrLf)
50190   .PrintingPSFiles = Replace$(hLang.Retrieve("PSFiles", .PrintingPSFiles), "/n", vbCrLf)
50200   .PrintingSave = Replace$(hLang.Retrieve("Save", .PrintingSave), "/n", vbCrLf)
50210   .PrintingStartStandardProgram = Replace$(hLang.Retrieve("StartStandardProgram", .PrintingStartStandardProgram), "/n", vbCrLf)
50220   .PrintingStatus = Replace$(hLang.Retrieve("Status", .PrintingStatus), "/n", vbCrLf)
50230   .PrintingSubject = Replace$(hLang.Retrieve("Subject", .PrintingSubject), "/n", vbCrLf)
50240   .PrintingTIFFFiles = Replace$(hLang.Retrieve("TIFFFiles", .PrintingTIFFFiles), "/n", vbCrLf)
50250  End With
50260  Set hLang = Nothing
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
50050   .CommonVersion = "0.9.4"
50060
50070   .DialogDocument = "&Document"
50080   .DialogDocumentAdd = "Add"
50090   .DialogDocumentBottom = "Bottom"
50100   .DialogDocumentCombine = "Combine"
50110   .DialogDocumentCombineAll = "Combine all"
50120   .DialogDocumentCombineAllSend = "Combine all and send"
50130   .DialogDocumentDelete = "Delete"
50140   .DialogDocumentDown = "Down"
50150   .DialogDocumentPrint = "Print"
50160   .DialogDocumentSave = "Save"
50170   .DialogDocumentSend = "Send"
50180   .DialogDocumentTop = "Top"
50190   .DialogDocumentUp = "Up"
50200   .DialogEmailAddress = "Email address"
50210   .DialogInfo = "&?"
50220   .DialogInfoCheckUpdates = "Check for Updates"
50230   .DialogInfoHomepage = "Product Homepage"
50240   .DialogInfoInfo = "About"
50250   .DialogInfoPaypal = "Paypal"
50260   .DialogInfoPDFCreatorSourceforge = "PDFCreator on Sourceforge"
50270   .DialogLanguage = "&Language"
50280   .DialogPrinter = "&Printer"
50290   .DialogPrinterClose = "Close"
50300   .DialogPrinterLogfile = "Logfile"
50310   .DialogPrinterLogfiles = "Logfiles"
50320   .DialogPrinterLogging = "Logging"
50330   .DialogPrinterOptions = "Options"
50340   .DialogPrinterPrinterStop = "Printer stop"
50350   .DialogView = "&View"
50360   .DialogViewStatusbar = "Status Bar"
50370   .DialogViewToolbars = "&Toolbars"
50380   .DialogViewToolbarsEmail = "Email"
50390   .DialogViewToolbarsStandard = "Standard"
50400
50410   .ListAddFile = "Add a file"
50420   .ListAllFiles = "All files"
50430   .ListBytes = "Bytes"
50440   .ListDate = "Created on"
50450   .ListDocumenttitle = "Document Title"
50460   .ListFilename = "Filename"
50470   .ListGBytes = "GBytes"
50480   .ListKBytes = "kBytes"
50490   .ListMBytes = "MBytes"
50500   .ListPDFFiles = "PDF Files"
50510   .ListPostscriptFiles = "PostScript Files"
50520   .ListPrinting = "Printing"
50530   .ListSize = "Size"
50540   .ListStatus = "Status"
50550   .ListWaiting = "Waiting"
50560
50570   .LoggingClear = "Cl&ear"
50580   .LoggingClose = "&Close"
50590   .LoggingLogfile = "Logfile"
50600
50610   .MessagesMsg01 = "Document in queue."
50620   .MessagesMsg02 = "Documents in queue."
50630   .MessagesMsg03 = "Do you wish to reset all settings?"
50640   .MessagesMsg04 = "Error: Cannot send Email!"
50650   .MessagesMsg05 = "File already exists. Do you want to overwrite it?"
50660   .MessagesMsg06 = "This file does not seem to be a postscript file!"
50670   .MessagesMsg07 = "There is a problem when trying to access this drive or directory!"
50680   .MessagesMsg08 = "Cannot find gsdll32.dll. Please check the ghostscript-program directory (see options)!"
50690   .MessagesMsg09 = "The output path does not exist. Do you want to create it?"
50700   .MessagesMsg10 = "This is not a valid path!"
50710   .MessagesMsg11 = "There is already such an entry!"
50720   .MessagesMsg12 = "Please don't use these forbidden characters for a filename!"
50730   .MessagesMsg13 = "Delete all program settings?"
50740   .MessagesMsg14 = "The file can not be found!"
50750   .MessagesMsg15 = "Cannot find gsdll32.dll in this directory!"
50760   .MessagesMsg16 = "No ghostscript font found in this directory!"
50770   .MessagesMsg17 = "No files in this directory!"
50780   .MessagesMsg19 = "You need either pdfenc or AFPL Ghostscript greater than, or equal to, version 8.14!"
50790   .MessagesMsg20 = "There was a problem sending an email with the standard emailclient!"
50800   .MessagesMsg21 = "User passwords do not match!"
50810   .MessagesMsg22 = "Owner passwords do not match!"
50820   .MessagesMsg23 = "The document is not protected!"
50830   .MessagesMsg24 = "The user password is empty! Continue?"
50840   .MessagesMsg25 = "The owner password is empty! Continue?"
50850   .MessagesMsg26 = "Unknown error"
50860   .MessagesMsg27 = "Cannot find the file/page."
50870   .MessagesMsg28 = "The filesize is 0 byte."
50880   .MessagesMsg29 = "Server not found."
50890   .MessagesMsg30 = "The url isn not interpretable."
50900   .MessagesMsg31 = "An error has occured"
50910   .MessagesMsg32 = "The new version %1 is available. Would you like download the new version from the Sourceforge pages?"
50920   .MessagesMsg33 = "You already have the most recent version."
50930   .MessagesMsg34 = "The file is in use. Please close the file first or choose another filename."
50940   .MessagesMsg35 = "It is necessary to temporarily set PDFCreator as defaultprinter."
50950   .MessagesMsg36 = "Don't ask me again."
50960   .MessagesMsg37 = "The downloaded file is not a valid language file!"
50970   .MessagesMsg38 = "The language file has been successfully installed!"
50980
50990   .OptionsAdditionalGhostscriptParameters = "Additional Ghostscript parameters"
51000   .OptionsAdditionalGhostscriptSearchpath = "Additional Ghostscript searchpath"
51010   .OptionsAddWindowsFontpath = "Use Windows fonts"
51020   .OptionsAssociatePSFiles = "Associate PDFCreator with postscript files"
51030   .OptionsAutosaveDirectoryPrompt = "Select Autosave Directory"
51040   .OptionsAutosaveFilename = "Filename"
51050   .OptionsAutosaveFilenameTokens = "Add a Filename-Token"
51060   .OptionsAutosaveFormat = "Autosave format"
51070   .OptionsAutosaveStartStandardProgram = "After auto-saving open the document with the default program."
51080   .OptionsBitmapResolution = "Resolution"
51090   .OptionsBMPColorscount01 = "4294967296 colors (32 Bit)"
51100   .OptionsBMPColorscount02 = "16777216 colors (24 Bit)"
51110   .OptionsBMPColorscount03 = "256 colors (8 Bit)"
51120   .OptionsBMPColorscount04 = "16 colors (4 Bit)"
51130   .OptionsBMPColorscount05 = "8 colors (3 Bit)"
51140   .OptionsBMPColorscount06 = "2 colors (Black/White)"
51150   .OptionsBMPColorscount07 = "Greyscale (8 Bit)"
51160   .OptionsBMPDescription = "Windows Bitmap Format. Please use only for single pages."
51170   .OptionsBMPSymbol = "BMP"
51180   .OptionsCancel = "&Cancel"
51190   .OptionsCustomPapersizeHeight = "Height"
51200   .OptionsCustomPapersizeInfo = "Units of 1/72 of an inch."
51210   .OptionsCustomPapersizeWidth = "Width"
51220   .OptionsDirectoriesGSBin = "Ghostscript Binaries"
51230   .OptionsDirectoriesGSFonts = "Ghostscript Fonts"
51240   .OptionsDirectoriesGSLibraries = "Ghostscript Libraries"
51250   .OptionsDirectoriesTempPath = "Temporary Files"
51260   .OptionsDocument = "Document"
51270   .OptionsEPSDescription = "Encapsulated Postscript Format"
51280   .OptionsEPSFiles = "Encapsulated Postscript-Files"
51290   .OptionsEPSSymbol = "EPS"
51300   .OptionsGhostscriptBinariesDirectoryPrompt = "Select Ghostscript Binaries Directory"
51310   .OptionsGhostscriptFontsDirectoryPrompt = "Select Ghostscript Fonts Directory"
51320   .OptionsGhostscriptInternal = "Internal Ghostscript: %1 Ghostscript %2"
51330   .OptionsGhostscriptLibrariesDirectoryPrompt = "Select Ghostscript Libraries Directory"
51340   .OptionsGhostscriptResourceDirectoryPrompt = "Select Ghostscript Resource Directory"
51350   .OptionsGhostscriptversion = "Ghostscript Version"
51360   .OptionsImageSettings = "Settings"
51370   .OptionsJavaPath = "Path to Java Interpreter"
51380   .OptionsJPEGColorscount01 = "16777216 colors (24 Bit)"
51390   .OptionsJPEGColorscount02 = "Greyscale (8 Bit)"
51400   .OptionsJPEGDescription = "JPEG (JFIF) Format. Please use only for single pages."
51410   .OptionsJPEGQuality = "Quality:"
51420   .OptionsJPEGSymbol = "JPEG"
51430   .OptionsLanguagesCurrentLanguage = "Current language"
51440   .OptionsLanguagesDownloadMoreLanguages = "Load more languages from the internet"
51450   .OptionsLanguagesInstall = "Install"
51460   .OptionsLanguagesRefresh = "Refresh List"
51470   .OptionsLanguagesTranslation = "Translation"
51480   .OptionsLanguagesVersion = "Version"
51490   .OptionsOnePagePerFile = "One page per file (not for pdf and eps files)"
51500   .OptionsOwnerPass = "Owner Password"
51510   .OptionsPassCancel = "Cancel"
51520   .OptionsPassOK = "OK"
51530   .OptionsPCXColorscount01 = "4294967296 colors (32 Bit) CMYK"
51540   .OptionsPCXColorscount02 = "16777216 colors (24 Bit)"
51550   .OptionsPCXColorscount03 = "256 colors (8 Bit)"
51560   .OptionsPCXColorscount04 = "16 colors (4 Bit)"
51570   .OptionsPCXColorscount05 = "2 colors (Black/White)"
51580   .OptionsPCXColorscount06 = "Greyscale (8 Bit)"
51590   .OptionsPCXDescription = "PCX Format. Please use only for single pages."
51600   .OptionsPCXSymbol = "PCX"
51610   .OptionsPDFAllowAssembly = "Allow changes to the assembly"
51620   .OptionsPDFAllowDegradedPrinting = "Allow printing in low resolution"
51630   .OptionsPDFAllowFillIn = "Allow filling in form fields"
51640   .OptionsPDFAllowScreenReaders = "Allow screen readers"
51650   .OptionsPDFColors = "Colors"
51660   .OptionsPDFColorsCaption = "Color Options"
51670   .OptionsPDFColorsCMYKtoRGB = "Convert CMYK images to RGB"
51680   .OptionsPDFColorsColorModel01 = "Use Color Model Device RGB"
51690   .OptionsPDFColorsColorModel02 = "Use Color Model Device CMYK"
51700   .OptionsPDFColorsColorModel03 = "Use Color Model Device Grayscale"
51710   .OptionsPDFColorsColorOptions = "Options"
51720   .OptionsPDFColorsPreserveHalftone = "Preserve Halftone Information"
51730   .OptionsPDFColorsPreserveOverprint = "Preserve Overprint Settings"
51740   .OptionsPDFColorsPreserveTransfer = "Preserve Transfer Functions"
51750   .OptionsPDFCompression = "Compression"
51760   .OptionsPDFCompressionCaption = "PDF Compression"
51770   .OptionsPDFCompressionColor = "Color Images"
51780   .OptionsPDFCompressionColorComp = "Compress"
51790   .OptionsPDFCompressionColorComp01 = "Automatic"
51800   .OptionsPDFCompressionColorComp02 = "JPEG-Maximum"
51810   .OptionsPDFCompressionColorComp03 = "JPEG-High"
51820   .OptionsPDFCompressionColorComp04 = "JPEG-Medium"
51830   .OptionsPDFCompressionColorComp05 = "JPEG-Low"
51840   .OptionsPDFCompressionColorComp06 = "JPEG-Minimum"
51850   .OptionsPDFCompressionColorComp07 = "ZIP"
51860   .OptionsPDFCompressionColorComp08 = "LZW-Compression"
51870   .OptionsPDFCompressionColorRes = "Resolution"
51880   .OptionsPDFCompressionColorResample = "Resample"
51890   .OptionsPDFCompressionColorResample01 = "Downsample"
51900   .OptionsPDFCompressionColorResample02 = "Average Downsample"
51910   .OptionsPDFCompressionColorResample03 = "Bicubic"
51920   .OptionsPDFCompressionGrey = "Greyscale Images"
51930   .OptionsPDFCompressionGreyComp = "Compress"
51940   .OptionsPDFCompressionGreyComp01 = "Automatic"
51950   .OptionsPDFCompressionGreyComp02 = "JPEG-Maximum"
51960   .OptionsPDFCompressionGreyComp03 = "JPEG-High"
51970   .OptionsPDFCompressionGreyComp04 = "JPEG-Medium"
51980   .OptionsPDFCompressionGreyComp05 = "JPEG-Low"
51990   .OptionsPDFCompressionGreyComp06 = "JPEG-Minimum"
52000   .OptionsPDFCompressionGreyComp07 = "ZIP"
52010   .OptionsPDFCompressionGreyComp08 = "LZW-Compression"
52020   .OptionsPDFCompressionGreyRes = "Resolution"
52030   .OptionsPDFCompressionGreyResample = "Resample"
52040   .OptionsPDFCompressionGreyResample01 = "Downsample"
52050   .OptionsPDFCompressionGreyResample02 = "Average Downsample"
52060   .OptionsPDFCompressionGreyResample03 = "Bicubic"
52070   .OptionsPDFCompressionMono = "Monochrome Images"
52080   .OptionsPDFCompressionMonoComp = "Compress"
52090   .OptionsPDFCompressionMonoComp01 = "CCITT Fax Compression"
52100   .OptionsPDFCompressionMonoComp02 = "ZIP"
52110   .OptionsPDFCompressionMonoComp03 = "Run-Length-Encoding"
52120   .OptionsPDFCompressionMonoComp04 = "LZW-Compression"
52130   .OptionsPDFCompressionMonoRes = "Resolution"
52140   .OptionsPDFCompressionMonoResample = "Resample"
52150   .OptionsPDFCompressionMonoResample01 = "Downsample"
52160   .OptionsPDFCompressionMonoResample02 = "Average Downsample"
52170   .OptionsPDFCompressionMonoResample03 = "Bicubic"
52180   .OptionsPDFCompressionTextComp = "Compress Text Objects"
52190   .OptionsPDFDescription = "Adobe PDF Format"
52200   .OptionsPDFDisallowCopy = "Copy text and images"
52210   .OptionsPDFDisallowModify = "Modify the document"
52220   .OptionsPDFDisallowModifyComments = "Modify comments"
52230   .OptionsPDFDisallowPrint = "Print the document"
52240   .OptionsPDFDisallowUser = "Disallow User to"
52250   .OptionsPDFEncryptionHigh = "High (128 Bit - Adobe Acrobat 5.0 and above)"
52260   .OptionsPDFEncryptionLevel = "Encryption Level"
52270   .OptionsPDFEncryptionLow = "Low (40 Bit - Adobe Acrobat 3.0 and above)"
52280   .OptionsPDFEncryptor = "Encryptor"
52290   .OptionsPDFEnhancedPermissions = "Enhanced Permissions (128 Bit only)"
52300   .OptionsPDFEnterPasswords = "Enter Passwords"
52310   .OptionsPDFFonts = "Fonts"
52320   .OptionsPDFFontsCaption = "Font Options"
52330   .OptionsPDFFontsEmbedAll = "Embed all fonts"
52340   .OptionsPDFFontsSubSetFonts = "Subset fonts when percentage of used characters below:"
52350   .OptionsPDFGeneral = "General"
52360   .OptionsPDFGeneralASCII85 = "Convert binary data to ASCII85"
52370   .OptionsPDFGeneralAutorotate = "Auto-Rotate Pages:"
52380   .OptionsPDFGeneralCaption = "General Options"
52390   .OptionsPDFGeneralCompatibility = "Compatibility:"
52400   .OptionsPDFGeneralCompatibility01 = "Adobe Acrobat 3.0 (PDF 1.2)"
52410   .OptionsPDFGeneralCompatibility02 = "Adobe Acrobat 4.0 (PDF 1.3)"
52420   .OptionsPDFGeneralCompatibility03 = "Adobe Acrobat 5.0 (PDF 1.4)"
52430   .OptionsPDFGeneralOverprint = "Overprint:"
52440   .OptionsPDFGeneralOverprint01 = "Non-Zero Overprint"
52450   .OptionsPDFGeneralOverprint02 = "Full Overprint"
52460   .OptionsPDFGeneralResolution = "Resolution:"
52470   .OptionsPDFGeneralRotate01 = "None"
52480   .OptionsPDFGeneralRotate02 = "All"
52490   .OptionsPDFGeneralRotate03 = "Single Page"
52500   .OptionsPDFOptimize = "Fast web view"
52510   .OptionsPDFOptions = "PDF Options"
52520   .OptionsPDFOwnerPass = "Password required to change permissions and passwords"
52530   .OptionsPDFPasswords = "Passwords"
52540   .OptionsPDFRepeatPassword = "Repeat"
52550   .OptionsPDFSecurity = "Security"
52560   .OptionsPDFSecurityCaption = "Security"
52570   .OptionsPDFSetPassword = "Password"
52580   .OptionsPDFSymbol = "PDF"
52590   .OptionsPDFUserPass = "Password required to open document"
52600   .OptionsPDFUseSecurity = "Use Security"
52610   .OptionsPNGColorscount01 = "16777216 colors (24 Bit)"
52620   .OptionsPNGColorscount02 = "256 colors (8 Bit)"
52630   .OptionsPNGColorscount03 = "16 colors (4 Bit)"
52640   .OptionsPNGColorscount04 = "2 colors (2 Bit - Black/White)"
52650   .OptionsPNGColorscount05 = "Greyscale (8 Bit)"
52660   .OptionsPNGDescription = "PNG Format. Please use only for single pages."
52670   .OptionsPNGFiles = "Bitmap PNG-Files"
52680   .OptionsPNGSymbol = "PNG"
52690   .OptionsPrintAfterSaving = "Print after saving"
52700   .OptionsPrintAfterSavingDuplex = "Duplex"
52710   .OptionsPrintAfterSavingDuplexTumbleOff = "Don't use tumble (Default)"
52720   .OptionsPrintAfterSavingDuplexTumbleOn = "Use tumble"
52730   .OptionsPrintAfterSavingNoCancel = "Hide the progress dialog during printing"
52740   .OptionsPrintAfterSavingPrinter = "Printer"
52750   .OptionsPrintAfterSavingQueryUser = "Query user"
52760   .OptionsPrintAfterSavingQueryUserDefaultPrinter = "Select the default Windows printer without any user interaction"
52770   .OptionsPrintAfterSavingQueryUserOff = "Off (Default)"
52780   .OptionsPrintAfterSavingQueryUserPrinterSetupDialog = "Shows the printer setup dialog"
52790   .OptionsPrintAfterSavingQueryUserStandardPrinterDialog = "Show the standard printer dialog"
52800   .OptionsPrintertempDirectoryPrompt = "Select Printer Temp-Directory"
52810   .OptionsPrintTestpage = "Print Test Page"
52820   .OptionsProcesspriority = "Process priority"
52830   .OptionsProcesspriorityHigh = "High"
52840   .OptionsProcesspriorityIdle = "Idle"
52850   .OptionsProcesspriorityNormal = "Normal"
52860   .OptionsProcesspriorityRealtime = "Realtime"
52870   .OptionsProgramActionsDescription = "Define an action before and after saving a file."
52880   .OptionsProgramActionsSymbol = "Actions"
52890   .OptionsProgramAutosaveDescription = "Auto-save mode. Auto-save does not prompt for a filename and file location. It automatically saves all PDF files to a single directory with a predefined filename."
52900   .OptionsProgramAutosaveSymbol = "Auto-save"
52910   .OptionsProgramDirectoriesDescription = "Directories for Ghostscript, temporary files and others."
52920   .OptionsProgramDirectoriesSymbol = "Directories"
52930   .OptionsProgramDocumentDescription = "Document properties"
52940   .OptionsProgramDocumentDescription1 = "Document properties 1"
52950   .OptionsProgramDocumentDescription2 = "Document properties 2"
52960   .OptionsProgramDocumentSymbol = "Document"
52970   .OptionsProgramFont = "Program Font"
52980   .OptionsProgramFontCancelTest = "Cancel Test"
52990   .OptionsProgramFontcharset = "Character Set"
53000   .OptionsProgramFontDescription = "Font for labels, captions and values. For the program menu use the general settings in your Windows OS."
53010   .OptionsProgramFontSize = "Size"
53020   .OptionsProgramFontSymbol = "Program font"
53030   .OptionsProgramFontTest = "Test"
53040   .OptionsProgramFontTestdescription = "Here you can test the font."
53050   .OptionsProgramGeneralDescription = "General Settings"
53060   .OptionsProgramGeneralDescription1 = "General Settings 1"
53070   .OptionsProgramGeneralDescription2 = "General Settings 2"
53080   .OptionsProgramGeneralSymbol = "General settings"
53090   .OptionsProgramGhostscriptDescription = "Ghostscript"
53100   .OptionsProgramGhostscriptSymbol = "Ghostscript"
53110   .OptionsProgramLanguagesDescription = "Define the language and download another languages from the internet."
53120   .OptionsProgramLanguagesSymbol = "Languages"
53130   .OptionsProgramNoProcessingAtStartup = "No processing at startup"
53140   .OptionsProgramOptionsDesign = "Frame color of the options dialog"
53150   .OptionsProgramOptionsDesignGradient = "Red and blue gradient (Default)"
53160   .OptionsProgramOptionsDesignSimple = "Simple red and blue color"
53170   .OptionsProgramPrintDescription = "Print after saving"
53180   .OptionsProgramPrintSymbol = "Print"
53190   .OptionsProgramRunProgramAfterSavingCaption = "Action after saving"
53200   .OptionsProgramRunProgramAfterSavingProgram = "Program/Script"
53210   .OptionsProgramRunProgramAfterSavingProgramParameters = "Program parameters"
53220   .OptionsProgramRunProgramAfterSavingWaitUntilReady = "Wait until the program/script is ready"
53230   .OptionsProgramRunProgramAfterSavingWindowstyle = "Window style"
53240   .OptionsProgramRunProgramAfterSavingWindowstyleHide = "Hide"
53250   .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus = "Maximized/Focus"
53260   .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus = "Minimized/Focus"
53270   .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus = "Minimized/No focus"
53280   .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus = "Normal/Focus"
53290   .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus = "Normal/No focus"
53300   .OptionsProgramRunProgramBeforeSavingCaption = "Action before saving"
53310   .OptionsProgramRunProgramBeforeSavingProgram = "Program/Script"
53320   .OptionsProgramRunProgramBeforeSavingProgramParameters = "Program parameters"
53330   .OptionsProgramRunProgramBeforeSavingWaitUntilReady = "Wait until the program/script is ready"
53340   .OptionsProgramRunProgramBeforeSavingWindowstyle = "Window style"
53350   .OptionsProgramRunProgramBeforeSavingWindowstyleHide = "Hide"
53360   .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus = "Maximized/Focus"
53370   .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus = "Minimized/Focus"
53380   .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus = "Minimized/NoFocus"
53390   .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus = "Normal/Focus"
53400   .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus = "Normal/NoFocus"
53410   .OptionsProgramSaveDescription = "Save"
53420   .OptionsProgramSaveSymbol = "Save"
53430   .OptionsProgramShowAnimation = "Show an animation during the process"
53440   .OptionsProgramSwitchingDefaultprinter = "No confirm message switching PDFCreator temporarily as default printer."
53450   .OptionsPSDescription = "Postscript Format"
53460   .OptionsPSFiles = "Postscript-Files"
53470   .OptionsPSLanguageLevel = "Language Level:"
53480   .OptionsPSSymbol = "PS"
53490   .OptionsRemoveSpaces = "Remove leading and trailing spaces"
53500   .OptionsReset = "&Reset all settings"
53510   .OptionsSave = "&Save"
53520   .OptionsSaveFilename = "Filename"
53530   .OptionsSaveFilenameAdd = "Add"
53540   .OptionsSaveFilenameChange = "Change"
53550   .OptionsSaveFilenameDelete = "Delete"
53560   .OptionsSaveFilenameSubstitutions = "Filename substitution"
53570   .OptionsSaveFilenameSubstitutionsTitle = "Filename substitution only in <Title>"
53580   .OptionsSaveFilenameTokens = "Add a Filename-Token"
53590   .OptionsSavePasswords = "Save passwords temporarily for this session."
53600   .OptionsSendEmailAfterAutosave = "Send an email after auto-saving"
53610   .OptionsSendMailMethod = "Methode to send an email"
53620   .OptionsSendMailMethodAutomatic = "Automatic"
53630   .OptionsSendMailMethodMapi = "Mapi interface"
53640   .OptionsSendMailMethodSendmailDLL = "Using sendmail.dll"
53650   .OptionsShellIntegration = "Shell integration"
53660   .OptionsShellIntegrationAdd = "Integrate PDFCreator into shell"
53670   .OptionsShellIntegrationCaption = "Create &PDF with PDFCreator"
53680   .OptionsShellIntegrationRemove = "Remove shell integration"
53690   .OptionsStamp = "Stamp"
53700   .OptionsStampFontColor = "Font-color"
53710   .OptionsStampOutlineFontThickness = "Outline font thickness"
53720   .OptionsStampString = "Stampstring"
53730   .OptionsStampUseOutlineFont = "Use outline font"
53740   .OptionsStandardAuthorToken = "Add a Author-Token"
53750   .OptionsStandardSaveFormat = "Standard save format"
53760   .OptionsTestpage = "PDFCreator Testpage"
53770   .OptionsTIFFColorscount01 = "16777216 (24 Bit)"
53780   .OptionsTIFFColorscount02 = "4096 (12 Bit)"
53790   .OptionsTIFFColorscount03 = "2 colors (Black/White) G3 fax encoding with no EOLs"
53800   .OptionsTIFFColorscount04 = "2 colors (Black/White) G3 fax encoding with EOLs"
53810   .OptionsTIFFColorscount05 = "2 colors (Black/White) 2-D G3 fax encoding"
53820   .OptionsTIFFColorscount06 = "2 colors (Black/White) G4 fax encoding"
53830   .OptionsTIFFColorscount07 = "2 colors (Black/White) LZW-compatible"
53840   .OptionsTIFFColorscount08 = "2 colors (Black/White) PackBits"
53850   .OptionsTIFFDescription = "TIFF Format. For multipages use the tiff-format."
53860   .OptionsTIFFSymbol = "TIFF"
53870   .OptionsTreeFormats = "Formats"
53880   .OptionsTreeProgram = "Program"
53890   .OptionsUseAutosave = "Use Auto-save"
53900   .OptionsUseAutosaveDirectory = "Use this directory for auto-save"
53910   .OptionsUseCreationDateNow = "Use the current Date/Time for 'Creation Date'"
53920   .OptionsUseCustomPapersize = "Use custom paper size"
53930   .OptionsUseFixPapersize = "Use fixed paper size"
53940   .OptionsUserPass = "User Password"
53950   .OptionsUseStandardauthor = "Use standard author"
53960
53970   .PrintingAuthor = "A&uthor:"
53980   .PrintingBMPFiles = "BMP-Files"
53990   .PrintingCancel = "&Cancel"
54000   .PrintingCollect = "&Wait - Collect"
54010   .PrintingCreationDate = "Creation &Date:"
54020   .PrintingDocumentTitle = "Document &Title:"
54030   .PrintingEMail = "&eMail"
54040   .PrintingEPSFiles = "Encapsulated Postscript-Files"
54050   .PrintingJPEGFiles = "JPEG-Files"
54060   .PrintingKeywords = "&Keywords:"
54070   .PrintingModifyDate = "&Modify Date:"
54080   .PrintingNow = "Now"
54090   .PrintingPCXFiles = "PCX-Files"
54100   .PrintingPDFFiles = "PDF-Files"
54110   .PrintingPNGFiles = "PNG-Files"
54120   .PrintingPSFiles = "Postscript-Files"
54130   .PrintingSave = "&Save"
54140   .PrintingStartStandardProgram = "&After saving open the document with the default program."
54150   .PrintingStatus = "Creating file..."
54160   .PrintingSubject = "Su&bject:"
54170   .PrintingTIFFFiles = "TIFF-Files"
54180
54190   .SaveOpenAttributes = "Attributes"
54200   .SaveOpenCancel = "Cancel"
54210   .SaveOpenFilename = "Filename"
54220   .SaveOpenOpen = "Open"
54230   .SaveOpenOpenTitle = "Open"
54240   .SaveOpenSave = "Save"
54250   .SaveOpenSaveTitle = "Save as"
54260   .SaveOpenSize = "Size"
54270
54280  End With
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

