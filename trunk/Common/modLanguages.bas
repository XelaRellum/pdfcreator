Attribute VB_Name = "modLanguage"
Option Explicit

' Automatically generated with DeveloperTool by Frank Heindörfer
' 2003 - 2007
' Email: thesmilyface@users.sourceforge.net

Public Type tLanguageStrings
 CommonAuthor As String
 CommonLanguagename As String
 CommonTitle As String
 CommonVersion As String

 DialogDocument As String
 DialogDocumentAdd As String
 DialogDocumentAddFromClipboard As String
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
 MessagesMsg39 As String

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
 OptionsBrowserAddOn As String
 OptionsBrowserAddOnInstall As String
 OptionsCancel As String
 OptionsCheckUpdateDescription As String
 OptionsCheckUpdateInterval As String
 OptionsCheckUpdateInterval01 As String
 OptionsCheckUpdateInterval02 As String
 OptionsCheckUpdateInterval03 As String
 OptionsCheckUpdateInterval04 As String
 OptionsCheckUpdateNow As String
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
 OptionsNothingToConfigure As String
 OptionsOnePagePerFile As String
 OptionsOwnerPass As String
 OptionsPassCancel As String
 OptionsPassOK As String
 OptionsPCLColorscount01 As String
 OptionsPCLColorscount02 As String
 OptionsPCLDescription As String
 OptionsPCLSymbol As String
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
 OptionsPDFGeneralCompatibility04 As String
 OptionsPDFGeneralDefaultSettings As String
 OptionsPDFGeneralDefaultSettingsDefault As String
 OptionsPDFGeneralDefaultSettingsEbook As String
 OptionsPDFGeneralDefaultSettingsPrepress As String
 OptionsPDFGeneralDefaultSettingsPrinter As String
 OptionsPDFGeneralDefaultSettingsScreen As String
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
 OptionsPDFSigning As String
 OptionsPDFSigningCaption As String
 OptionsPDFSigningPfxFile As String
 OptionsPDFSigningSignatureContact As String
 OptionsPDFSigningSignatureLocation As String
 OptionsPDFSigningSignatureMultiSignature As String
 OptionsPDFSigningSignaturePosition As String
 OptionsPDFSigningSignaturePositionLeftX As String
 OptionsPDFSigningSignaturePositionLeftY As String
 OptionsPDFSigningSignaturePositionRightX As String
 OptionsPDFSigningSignaturePositionRightY As String
 OptionsPDFSigningSignatureReason As String
 OptionsPDFSigningSignatureVisible As String
 OptionsPDFSigningSignPdfFile As String
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
 OptionsPSDColorsCount01 As String
 OptionsPSDColorscount02 As String
 OptionsPSDDescription As String
 OptionsPSDescription As String
 OptionsPSDSymbol As String
 OptionsPSFiles As String
 OptionsPSLanguageLevel As String
 OptionsPSSymbol As String
 OptionsRAWColorsCount01 As String
 OptionsRAWColorscount02 As String
 OptionsRAWColorscount03 As String
 OptionsRAWDescription As String
 OptionsRAWSymbol As String
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
 OptionsTXTDescription As String
 OptionsTXTSymbol As String
 OptionsUseAutosave As String
 OptionsUseAutosaveDirectory As String
 OptionsUseCreationDateNow As String
 OptionsUseCustomPapersize As String
 OptionsUseFixPapersize As String
 OptionsUserPass As String
 OptionsUseStandardauthor As String
 OptionsXCFColorsCount01 As String
 OptionsXCFColorscount02 As String
 OptionsXCFDescription As String
 OptionsXCFSymbol As String

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
 PrintingPCLFiles As String
 PrintingPCXFiles As String
 PrintingPDFAFiles As String
 PrintingPDFFiles As String
 PrintingPDFXFiles As String
 PrintingPNGFiles As String
 PrintingPSDFiles As String
 PrintingPSFiles As String
 PrintingRAWFiles As String
 PrintingSave As String
 PrintingStartStandardProgram As String
 PrintingStatus As String
 PrintingSubject As String
 PrintingTIFFFiles As String
 PrintingTXTFiles As String
 PrintingXCFFiles As String

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
50060   .DialogDocumentAddFromClipboard = Replace$(hLang.Retrieve("DocumentAddFromClipboard", .DialogDocumentAddFromClipboard), "/n", vbCrLf)
50070   .DialogDocumentBottom = Replace$(hLang.Retrieve("DocumentBottom", .DialogDocumentBottom), "/n", vbCrLf)
50080   .DialogDocumentCombine = Replace$(hLang.Retrieve("DocumentCombine", .DialogDocumentCombine), "/n", vbCrLf)
50090   .DialogDocumentCombineAll = Replace$(hLang.Retrieve("DocumentCombineAll", .DialogDocumentCombineAll), "/n", vbCrLf)
50100   .DialogDocumentCombineAllSend = Replace$(hLang.Retrieve("DocumentCombineAllSend", .DialogDocumentCombineAllSend), "/n", vbCrLf)
50110   .DialogDocumentDelete = Replace$(hLang.Retrieve("DocumentDelete", .DialogDocumentDelete), "/n", vbCrLf)
50120   .DialogDocumentDown = Replace$(hLang.Retrieve("DocumentDown", .DialogDocumentDown), "/n", vbCrLf)
50130   .DialogDocumentPrint = Replace$(hLang.Retrieve("DocumentPrint", .DialogDocumentPrint), "/n", vbCrLf)
50140   .DialogDocumentSave = Replace$(hLang.Retrieve("DocumentSave", .DialogDocumentSave), "/n", vbCrLf)
50150   .DialogDocumentSend = Replace$(hLang.Retrieve("DocumentSend", .DialogDocumentSend), "/n", vbCrLf)
50160   .DialogDocumentTop = Replace$(hLang.Retrieve("DocumentTop", .DialogDocumentTop), "/n", vbCrLf)
50170   .DialogDocumentUp = Replace$(hLang.Retrieve("DocumentUp", .DialogDocumentUp), "/n", vbCrLf)
50180   .DialogEmailAddress = Replace$(hLang.Retrieve("EmailAddress", .DialogEmailAddress), "/n", vbCrLf)
50190   .DialogInfo = Replace$(hLang.Retrieve("Info", .DialogInfo), "/n", vbCrLf)
50200   .DialogInfoCheckUpdates = Replace$(hLang.Retrieve("InfoCheckUpdates", .DialogInfoCheckUpdates), "/n", vbCrLf)
50210   .DialogInfoHomepage = Replace$(hLang.Retrieve("InfoHomepage", .DialogInfoHomepage), "/n", vbCrLf)
50220   .DialogInfoInfo = Replace$(hLang.Retrieve("InfoInfo", .DialogInfoInfo), "/n", vbCrLf)
50230   .DialogInfoPaypal = Replace$(hLang.Retrieve("InfoPaypal", .DialogInfoPaypal), "/n", vbCrLf)
50240   .DialogInfoPDFCreatorSourceforge = Replace$(hLang.Retrieve("InfoPDFCreatorSourceforge", .DialogInfoPDFCreatorSourceforge), "/n", vbCrLf)
50250   .DialogLanguage = Replace$(hLang.Retrieve("Language", .DialogLanguage), "/n", vbCrLf)
50260   .DialogPrinter = Replace$(hLang.Retrieve("Printer", .DialogPrinter), "/n", vbCrLf)
50270   .DialogPrinterClose = Replace$(hLang.Retrieve("PrinterClose", .DialogPrinterClose), "/n", vbCrLf)
50280   .DialogPrinterLogfile = Replace$(hLang.Retrieve("PrinterLogfile", .DialogPrinterLogfile), "/n", vbCrLf)
50290   .DialogPrinterLogfiles = Replace$(hLang.Retrieve("PrinterLogfiles", .DialogPrinterLogfiles), "/n", vbCrLf)
50300   .DialogPrinterLogging = Replace$(hLang.Retrieve("PrinterLogging", .DialogPrinterLogging), "/n", vbCrLf)
50310   .DialogPrinterOptions = Replace$(hLang.Retrieve("PrinterOptions", .DialogPrinterOptions), "/n", vbCrLf)
50320   .DialogPrinterPrinterStop = Replace$(hLang.Retrieve("PrinterPrinterStop", .DialogPrinterPrinterStop), "/n", vbCrLf)
50330   .DialogView = Replace$(hLang.Retrieve("View", .DialogView), "/n", vbCrLf)
50340   .DialogViewStatusbar = Replace$(hLang.Retrieve("ViewStatusbar", .DialogViewStatusbar), "/n", vbCrLf)
50350   .DialogViewToolbars = Replace$(hLang.Retrieve("ViewToolbars", .DialogViewToolbars), "/n", vbCrLf)
50360   .DialogViewToolbarsEmail = Replace$(hLang.Retrieve("ViewToolbarsEmail", .DialogViewToolbarsEmail), "/n", vbCrLf)
50370   .DialogViewToolbarsStandard = Replace$(hLang.Retrieve("ViewToolbarsStandard", .DialogViewToolbarsStandard), "/n", vbCrLf)
50380  End With
50390  Set hLang = Nothing
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
50410   .MessagesMsg39 = Replace$(hLang.Retrieve("Msg39", .MessagesMsg39), "/n", vbCrLf)
50420  End With
50430  Set hLang = Nothing
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
50230   .OptionsBrowserAddOn = Replace$(hLang.Retrieve("BrowserAddOn", .OptionsBrowserAddOn), "/n", vbCrLf)
50240   .OptionsBrowserAddOnInstall = Replace$(hLang.Retrieve("BrowserAddOnInstall", .OptionsBrowserAddOnInstall), "/n", vbCrLf)
50250   .OptionsCancel = Replace$(hLang.Retrieve("Cancel", .OptionsCancel), "/n", vbCrLf)
50260   .OptionsCheckUpdateDescription = Replace$(hLang.Retrieve("CheckUpdateDescription", .OptionsCheckUpdateDescription), "/n", vbCrLf)
50270   .OptionsCheckUpdateInterval = Replace$(hLang.Retrieve("CheckUpdateInterval", .OptionsCheckUpdateInterval), "/n", vbCrLf)
50280   .OptionsCheckUpdateInterval01 = Replace$(hLang.Retrieve("CheckUpdateInterval01", .OptionsCheckUpdateInterval01), "/n", vbCrLf)
50290   .OptionsCheckUpdateInterval02 = Replace$(hLang.Retrieve("CheckUpdateInterval02", .OptionsCheckUpdateInterval02), "/n", vbCrLf)
50300   .OptionsCheckUpdateInterval03 = Replace$(hLang.Retrieve("CheckUpdateInterval03", .OptionsCheckUpdateInterval03), "/n", vbCrLf)
50310   .OptionsCheckUpdateInterval04 = Replace$(hLang.Retrieve("CheckUpdateInterval04", .OptionsCheckUpdateInterval04), "/n", vbCrLf)
50320   .OptionsCheckUpdateNow = Replace$(hLang.Retrieve("CheckUpdateNow", .OptionsCheckUpdateNow), "/n", vbCrLf)
50330   .OptionsCustomPapersizeHeight = Replace$(hLang.Retrieve("CustomPapersizeHeight", .OptionsCustomPapersizeHeight), "/n", vbCrLf)
50340   .OptionsCustomPapersizeInfo = Replace$(hLang.Retrieve("CustomPapersizeInfo", .OptionsCustomPapersizeInfo), "/n", vbCrLf)
50350   .OptionsCustomPapersizeWidth = Replace$(hLang.Retrieve("CustomPapersizeWidth", .OptionsCustomPapersizeWidth), "/n", vbCrLf)
50360   .OptionsDirectoriesGSBin = Replace$(hLang.Retrieve("DirectoriesGSBin", .OptionsDirectoriesGSBin), "/n", vbCrLf)
50370   .OptionsDirectoriesGSFonts = Replace$(hLang.Retrieve("DirectoriesGSFonts", .OptionsDirectoriesGSFonts), "/n", vbCrLf)
50380   .OptionsDirectoriesGSLibraries = Replace$(hLang.Retrieve("DirectoriesGSLibraries", .OptionsDirectoriesGSLibraries), "/n", vbCrLf)
50390   .OptionsDirectoriesTempPath = Replace$(hLang.Retrieve("DirectoriesTempPath", .OptionsDirectoriesTempPath), "/n", vbCrLf)
50400   .OptionsDocument = Replace$(hLang.Retrieve("Document", .OptionsDocument), "/n", vbCrLf)
50410   .OptionsEPSDescription = Replace$(hLang.Retrieve("EPSDescription", .OptionsEPSDescription), "/n", vbCrLf)
50420   .OptionsEPSFiles = Replace$(hLang.Retrieve("EPSFiles", .OptionsEPSFiles), "/n", vbCrLf)
50430   .OptionsEPSSymbol = Replace$(hLang.Retrieve("EPSSymbol", .OptionsEPSSymbol), "/n", vbCrLf)
50440   .OptionsGhostscriptBinariesDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptBinariesDirectoryPrompt", .OptionsGhostscriptBinariesDirectoryPrompt), "/n", vbCrLf)
50450   .OptionsGhostscriptFontsDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptFontsDirectoryPrompt", .OptionsGhostscriptFontsDirectoryPrompt), "/n", vbCrLf)
50460   .OptionsGhostscriptInternal = Replace$(hLang.Retrieve("GhostscriptInternal", .OptionsGhostscriptInternal), "/n", vbCrLf)
50470   .OptionsGhostscriptLibrariesDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptLibrariesDirectoryPrompt", .OptionsGhostscriptLibrariesDirectoryPrompt), "/n", vbCrLf)
50480   .OptionsGhostscriptResourceDirectoryPrompt = Replace$(hLang.Retrieve("GhostscriptResourceDirectoryPrompt", .OptionsGhostscriptResourceDirectoryPrompt), "/n", vbCrLf)
50490   .OptionsGhostscriptversion = Replace$(hLang.Retrieve("Ghostscriptversion", .OptionsGhostscriptversion), "/n", vbCrLf)
50500   .OptionsImageSettings = Replace$(hLang.Retrieve("ImageSettings", .OptionsImageSettings), "/n", vbCrLf)
50510   .OptionsJavaPath = Replace$(hLang.Retrieve("JavaPath", .OptionsJavaPath), "/n", vbCrLf)
50520   .OptionsJPEGColorscount01 = Replace$(hLang.Retrieve("JPEGColorscount01", .OptionsJPEGColorscount01), "/n", vbCrLf)
50530   .OptionsJPEGColorscount02 = Replace$(hLang.Retrieve("JPEGColorscount02", .OptionsJPEGColorscount02), "/n", vbCrLf)
50540   .OptionsJPEGDescription = Replace$(hLang.Retrieve("JPEGDescription", .OptionsJPEGDescription), "/n", vbCrLf)
50550   .OptionsJPEGQuality = Replace$(hLang.Retrieve("JPEGQuality", .OptionsJPEGQuality), "/n", vbCrLf)
50560   .OptionsJPEGSymbol = Replace$(hLang.Retrieve("JPEGSymbol", .OptionsJPEGSymbol), "/n", vbCrLf)
50570   .OptionsLanguagesCurrentLanguage = Replace$(hLang.Retrieve("LanguagesCurrentLanguage", .OptionsLanguagesCurrentLanguage), "/n", vbCrLf)
50580   .OptionsLanguagesDownloadMoreLanguages = Replace$(hLang.Retrieve("LanguagesDownloadMoreLanguages", .OptionsLanguagesDownloadMoreLanguages), "/n", vbCrLf)
50590   .OptionsLanguagesInstall = Replace$(hLang.Retrieve("LanguagesInstall", .OptionsLanguagesInstall), "/n", vbCrLf)
50600   .OptionsLanguagesRefresh = Replace$(hLang.Retrieve("LanguagesRefresh", .OptionsLanguagesRefresh), "/n", vbCrLf)
50610   .OptionsLanguagesTranslation = Replace$(hLang.Retrieve("LanguagesTranslation", .OptionsLanguagesTranslation), "/n", vbCrLf)
50620   .OptionsLanguagesVersion = Replace$(hLang.Retrieve("LanguagesVersion", .OptionsLanguagesVersion), "/n", vbCrLf)
50630   .OptionsNothingToConfigure = Replace$(hLang.Retrieve("NothingToConfigure", .OptionsNothingToConfigure), "/n", vbCrLf)
50640   .OptionsOnePagePerFile = Replace$(hLang.Retrieve("OnePagePerFile", .OptionsOnePagePerFile), "/n", vbCrLf)
50650   .OptionsOwnerPass = Replace$(hLang.Retrieve("OwnerPass", .OptionsOwnerPass), "/n", vbCrLf)
50660   .OptionsPassCancel = Replace$(hLang.Retrieve("PassCancel", .OptionsPassCancel), "/n", vbCrLf)
50670   .OptionsPassOK = Replace$(hLang.Retrieve("PassOK", .OptionsPassOK), "/n", vbCrLf)
50680   .OptionsPCLColorscount01 = Replace$(hLang.Retrieve("PCLColorscount01", .OptionsPCLColorscount01), "/n", vbCrLf)
50690   .OptionsPCLColorscount02 = Replace$(hLang.Retrieve("PCLColorscount02", .OptionsPCLColorscount02), "/n", vbCrLf)
50700   .OptionsPCLDescription = Replace$(hLang.Retrieve("PCLDescription", .OptionsPCLDescription), "/n", vbCrLf)
50710   .OptionsPCLSymbol = Replace$(hLang.Retrieve("PCLSymbol", .OptionsPCLSymbol), "/n", vbCrLf)
50720   .OptionsPCXColorscount01 = Replace$(hLang.Retrieve("PCXColorscount01", .OptionsPCXColorscount01), "/n", vbCrLf)
50730   .OptionsPCXColorscount02 = Replace$(hLang.Retrieve("PCXColorscount02", .OptionsPCXColorscount02), "/n", vbCrLf)
50740   .OptionsPCXColorscount03 = Replace$(hLang.Retrieve("PCXColorscount03", .OptionsPCXColorscount03), "/n", vbCrLf)
50750   .OptionsPCXColorscount04 = Replace$(hLang.Retrieve("PCXColorscount04", .OptionsPCXColorscount04), "/n", vbCrLf)
50760   .OptionsPCXColorscount05 = Replace$(hLang.Retrieve("PCXColorscount05", .OptionsPCXColorscount05), "/n", vbCrLf)
50770   .OptionsPCXColorscount06 = Replace$(hLang.Retrieve("PCXColorscount06", .OptionsPCXColorscount06), "/n", vbCrLf)
50780   .OptionsPCXDescription = Replace$(hLang.Retrieve("PCXDescription", .OptionsPCXDescription), "/n", vbCrLf)
50790   .OptionsPCXSymbol = Replace$(hLang.Retrieve("PCXSymbol", .OptionsPCXSymbol), "/n", vbCrLf)
50800   .OptionsPDFAllowAssembly = Replace$(hLang.Retrieve("PDFAllowAssembly", .OptionsPDFAllowAssembly), "/n", vbCrLf)
50810   .OptionsPDFAllowDegradedPrinting = Replace$(hLang.Retrieve("PDFAllowDegradedPrinting", .OptionsPDFAllowDegradedPrinting), "/n", vbCrLf)
50820   .OptionsPDFAllowFillIn = Replace$(hLang.Retrieve("PDFAllowFillIn", .OptionsPDFAllowFillIn), "/n", vbCrLf)
50830   .OptionsPDFAllowScreenReaders = Replace$(hLang.Retrieve("PDFAllowScreenReaders", .OptionsPDFAllowScreenReaders), "/n", vbCrLf)
50840   .OptionsPDFColors = Replace$(hLang.Retrieve("PDFColors", .OptionsPDFColors), "/n", vbCrLf)
50850   .OptionsPDFColorsCaption = Replace$(hLang.Retrieve("PDFColorsCaption", .OptionsPDFColorsCaption), "/n", vbCrLf)
50860   .OptionsPDFColorsCMYKtoRGB = Replace$(hLang.Retrieve("PDFColorsCMYKtoRGB", .OptionsPDFColorsCMYKtoRGB), "/n", vbCrLf)
50870   .OptionsPDFColorsColorModel01 = Replace$(hLang.Retrieve("PDFColorsColorModel01", .OptionsPDFColorsColorModel01), "/n", vbCrLf)
50880   .OptionsPDFColorsColorModel02 = Replace$(hLang.Retrieve("PDFColorsColorModel02", .OptionsPDFColorsColorModel02), "/n", vbCrLf)
50890   .OptionsPDFColorsColorModel03 = Replace$(hLang.Retrieve("PDFColorsColorModel03", .OptionsPDFColorsColorModel03), "/n", vbCrLf)
50900   .OptionsPDFColorsColorOptions = Replace$(hLang.Retrieve("PDFColorsColorOptions", .OptionsPDFColorsColorOptions), "/n", vbCrLf)
50910   .OptionsPDFColorsPreserveHalftone = Replace$(hLang.Retrieve("PDFColorsPreserveHalftone", .OptionsPDFColorsPreserveHalftone), "/n", vbCrLf)
50920   .OptionsPDFColorsPreserveOverprint = Replace$(hLang.Retrieve("PDFColorsPreserveOverprint", .OptionsPDFColorsPreserveOverprint), "/n", vbCrLf)
50930   .OptionsPDFColorsPreserveTransfer = Replace$(hLang.Retrieve("PDFColorsPreserveTransfer", .OptionsPDFColorsPreserveTransfer), "/n", vbCrLf)
50940   .OptionsPDFCompression = Replace$(hLang.Retrieve("PDFCompression", .OptionsPDFCompression), "/n", vbCrLf)
50950   .OptionsPDFCompressionCaption = Replace$(hLang.Retrieve("PDFCompressionCaption", .OptionsPDFCompressionCaption), "/n", vbCrLf)
50960   .OptionsPDFCompressionColor = Replace$(hLang.Retrieve("PDFCompressionColor", .OptionsPDFCompressionColor), "/n", vbCrLf)
50970   .OptionsPDFCompressionColorComp = Replace$(hLang.Retrieve("PDFCompressionColorComp", .OptionsPDFCompressionColorComp), "/n", vbCrLf)
50980   .OptionsPDFCompressionColorComp01 = Replace$(hLang.Retrieve("PDFCompressionColorComp01", .OptionsPDFCompressionColorComp01), "/n", vbCrLf)
50990   .OptionsPDFCompressionColorComp02 = Replace$(hLang.Retrieve("PDFCompressionColorComp02", .OptionsPDFCompressionColorComp02), "/n", vbCrLf)
51000   .OptionsPDFCompressionColorComp03 = Replace$(hLang.Retrieve("PDFCompressionColorComp03", .OptionsPDFCompressionColorComp03), "/n", vbCrLf)
51010   .OptionsPDFCompressionColorComp04 = Replace$(hLang.Retrieve("PDFCompressionColorComp04", .OptionsPDFCompressionColorComp04), "/n", vbCrLf)
51020   .OptionsPDFCompressionColorComp05 = Replace$(hLang.Retrieve("PDFCompressionColorComp05", .OptionsPDFCompressionColorComp05), "/n", vbCrLf)
51030   .OptionsPDFCompressionColorComp06 = Replace$(hLang.Retrieve("PDFCompressionColorComp06", .OptionsPDFCompressionColorComp06), "/n", vbCrLf)
51040   .OptionsPDFCompressionColorComp07 = Replace$(hLang.Retrieve("PDFCompressionColorComp07", .OptionsPDFCompressionColorComp07), "/n", vbCrLf)
51050   .OptionsPDFCompressionColorComp08 = Replace$(hLang.Retrieve("PDFCompressionColorComp08", .OptionsPDFCompressionColorComp08), "/n", vbCrLf)
51060   .OptionsPDFCompressionColorRes = Replace$(hLang.Retrieve("PDFCompressionColorRes", .OptionsPDFCompressionColorRes), "/n", vbCrLf)
51070   .OptionsPDFCompressionColorResample = Replace$(hLang.Retrieve("PDFCompressionColorResample", .OptionsPDFCompressionColorResample), "/n", vbCrLf)
51080   .OptionsPDFCompressionColorResample01 = Replace$(hLang.Retrieve("PDFCompressionColorResample01", .OptionsPDFCompressionColorResample01), "/n", vbCrLf)
51090   .OptionsPDFCompressionColorResample02 = Replace$(hLang.Retrieve("PDFCompressionColorResample02", .OptionsPDFCompressionColorResample02), "/n", vbCrLf)
51100   .OptionsPDFCompressionColorResample03 = Replace$(hLang.Retrieve("PDFCompressionColorResample03", .OptionsPDFCompressionColorResample03), "/n", vbCrLf)
51110   .OptionsPDFCompressionGrey = Replace$(hLang.Retrieve("PDFCompressionGrey", .OptionsPDFCompressionGrey), "/n", vbCrLf)
51120   .OptionsPDFCompressionGreyComp = Replace$(hLang.Retrieve("PDFCompressionGreyComp", .OptionsPDFCompressionGreyComp), "/n", vbCrLf)
51130   .OptionsPDFCompressionGreyComp01 = Replace$(hLang.Retrieve("PDFCompressionGreyComp01", .OptionsPDFCompressionGreyComp01), "/n", vbCrLf)
51140   .OptionsPDFCompressionGreyComp02 = Replace$(hLang.Retrieve("PDFCompressionGreyComp02", .OptionsPDFCompressionGreyComp02), "/n", vbCrLf)
51150   .OptionsPDFCompressionGreyComp03 = Replace$(hLang.Retrieve("PDFCompressionGreyComp03", .OptionsPDFCompressionGreyComp03), "/n", vbCrLf)
51160   .OptionsPDFCompressionGreyComp04 = Replace$(hLang.Retrieve("PDFCompressionGreyComp04", .OptionsPDFCompressionGreyComp04), "/n", vbCrLf)
51170   .OptionsPDFCompressionGreyComp05 = Replace$(hLang.Retrieve("PDFCompressionGreyComp05", .OptionsPDFCompressionGreyComp05), "/n", vbCrLf)
51180   .OptionsPDFCompressionGreyComp06 = Replace$(hLang.Retrieve("PDFCompressionGreyComp06", .OptionsPDFCompressionGreyComp06), "/n", vbCrLf)
51190   .OptionsPDFCompressionGreyComp07 = Replace$(hLang.Retrieve("PDFCompressionGreyComp07", .OptionsPDFCompressionGreyComp07), "/n", vbCrLf)
51200   .OptionsPDFCompressionGreyComp08 = Replace$(hLang.Retrieve("PDFCompressionGreyComp08", .OptionsPDFCompressionGreyComp08), "/n", vbCrLf)
51210   .OptionsPDFCompressionGreyRes = Replace$(hLang.Retrieve("PDFCompressionGreyRes", .OptionsPDFCompressionGreyRes), "/n", vbCrLf)
51220   .OptionsPDFCompressionGreyResample = Replace$(hLang.Retrieve("PDFCompressionGreyResample", .OptionsPDFCompressionGreyResample), "/n", vbCrLf)
51230   .OptionsPDFCompressionGreyResample01 = Replace$(hLang.Retrieve("PDFCompressionGreyResample01", .OptionsPDFCompressionGreyResample01), "/n", vbCrLf)
51240   .OptionsPDFCompressionGreyResample02 = Replace$(hLang.Retrieve("PDFCompressionGreyResample02", .OptionsPDFCompressionGreyResample02), "/n", vbCrLf)
51250   .OptionsPDFCompressionGreyResample03 = Replace$(hLang.Retrieve("PDFCompressionGreyResample03", .OptionsPDFCompressionGreyResample03), "/n", vbCrLf)
51260   .OptionsPDFCompressionMono = Replace$(hLang.Retrieve("PDFCompressionMono", .OptionsPDFCompressionMono), "/n", vbCrLf)
51270   .OptionsPDFCompressionMonoComp = Replace$(hLang.Retrieve("PDFCompressionMonoComp", .OptionsPDFCompressionMonoComp), "/n", vbCrLf)
51280   .OptionsPDFCompressionMonoComp01 = Replace$(hLang.Retrieve("PDFCompressionMonoComp01", .OptionsPDFCompressionMonoComp01), "/n", vbCrLf)
51290   .OptionsPDFCompressionMonoComp02 = Replace$(hLang.Retrieve("PDFCompressionMonoComp02", .OptionsPDFCompressionMonoComp02), "/n", vbCrLf)
51300   .OptionsPDFCompressionMonoComp03 = Replace$(hLang.Retrieve("PDFCompressionMonoComp03", .OptionsPDFCompressionMonoComp03), "/n", vbCrLf)
51310   .OptionsPDFCompressionMonoComp04 = Replace$(hLang.Retrieve("PDFCompressionMonoComp04", .OptionsPDFCompressionMonoComp04), "/n", vbCrLf)
51320   .OptionsPDFCompressionMonoRes = Replace$(hLang.Retrieve("PDFCompressionMonoRes", .OptionsPDFCompressionMonoRes), "/n", vbCrLf)
51330   .OptionsPDFCompressionMonoResample = Replace$(hLang.Retrieve("PDFCompressionMonoResample", .OptionsPDFCompressionMonoResample), "/n", vbCrLf)
51340   .OptionsPDFCompressionMonoResample01 = Replace$(hLang.Retrieve("PDFCompressionMonoResample01", .OptionsPDFCompressionMonoResample01), "/n", vbCrLf)
51350   .OptionsPDFCompressionMonoResample02 = Replace$(hLang.Retrieve("PDFCompressionMonoResample02", .OptionsPDFCompressionMonoResample02), "/n", vbCrLf)
51360   .OptionsPDFCompressionMonoResample03 = Replace$(hLang.Retrieve("PDFCompressionMonoResample03", .OptionsPDFCompressionMonoResample03), "/n", vbCrLf)
51370   .OptionsPDFCompressionTextComp = Replace$(hLang.Retrieve("PDFCompressionTextComp", .OptionsPDFCompressionTextComp), "/n", vbCrLf)
51380   .OptionsPDFDescription = Replace$(hLang.Retrieve("PDFDescription", .OptionsPDFDescription), "/n", vbCrLf)
51390   .OptionsPDFDisallowCopy = Replace$(hLang.Retrieve("PDFDisallowCopy", .OptionsPDFDisallowCopy), "/n", vbCrLf)
51400   .OptionsPDFDisallowModify = Replace$(hLang.Retrieve("PDFDisallowModify", .OptionsPDFDisallowModify), "/n", vbCrLf)
51410   .OptionsPDFDisallowModifyComments = Replace$(hLang.Retrieve("PDFDisallowModifyComments", .OptionsPDFDisallowModifyComments), "/n", vbCrLf)
51420   .OptionsPDFDisallowPrint = Replace$(hLang.Retrieve("PDFDisallowPrint", .OptionsPDFDisallowPrint), "/n", vbCrLf)
51430   .OptionsPDFDisallowUser = Replace$(hLang.Retrieve("PDFDisallowUser", .OptionsPDFDisallowUser), "/n", vbCrLf)
51440   .OptionsPDFEncryptionHigh = Replace$(hLang.Retrieve("PDFEncryptionHigh", .OptionsPDFEncryptionHigh), "/n", vbCrLf)
51450   .OptionsPDFEncryptionLevel = Replace$(hLang.Retrieve("PDFEncryptionLevel", .OptionsPDFEncryptionLevel), "/n", vbCrLf)
51460   .OptionsPDFEncryptionLow = Replace$(hLang.Retrieve("PDFEncryptionLow", .OptionsPDFEncryptionLow), "/n", vbCrLf)
51470   .OptionsPDFEncryptor = Replace$(hLang.Retrieve("PDFEncryptor", .OptionsPDFEncryptor), "/n", vbCrLf)
51480   .OptionsPDFEnhancedPermissions = Replace$(hLang.Retrieve("PDFEnhancedPermissions", .OptionsPDFEnhancedPermissions), "/n", vbCrLf)
51490   .OptionsPDFEnterPasswords = Replace$(hLang.Retrieve("PDFEnterPasswords", .OptionsPDFEnterPasswords), "/n", vbCrLf)
51500   .OptionsPDFFonts = Replace$(hLang.Retrieve("PDFFonts", .OptionsPDFFonts), "/n", vbCrLf)
51510   .OptionsPDFFontsCaption = Replace$(hLang.Retrieve("PDFFontsCaption", .OptionsPDFFontsCaption), "/n", vbCrLf)
51520   .OptionsPDFFontsEmbedAll = Replace$(hLang.Retrieve("PDFFontsEmbedAll", .OptionsPDFFontsEmbedAll), "/n", vbCrLf)
51530   .OptionsPDFFontsSubSetFonts = Replace$(hLang.Retrieve("PDFFontsSubSetFonts", .OptionsPDFFontsSubSetFonts), "/n", vbCrLf)
51540   .OptionsPDFGeneral = Replace$(hLang.Retrieve("PDFGeneral", .OptionsPDFGeneral), "/n", vbCrLf)
51550   .OptionsPDFGeneralASCII85 = Replace$(hLang.Retrieve("PDFGeneralASCII85", .OptionsPDFGeneralASCII85), "/n", vbCrLf)
51560   .OptionsPDFGeneralAutorotate = Replace$(hLang.Retrieve("PDFGeneralAutorotate", .OptionsPDFGeneralAutorotate), "/n", vbCrLf)
51570   .OptionsPDFGeneralCaption = Replace$(hLang.Retrieve("PDFGeneralCaption", .OptionsPDFGeneralCaption), "/n", vbCrLf)
51580   .OptionsPDFGeneralCompatibility = Replace$(hLang.Retrieve("PDFGeneralCompatibility", .OptionsPDFGeneralCompatibility), "/n", vbCrLf)
51590   .OptionsPDFGeneralCompatibility01 = Replace$(hLang.Retrieve("PDFGeneralCompatibility01", .OptionsPDFGeneralCompatibility01), "/n", vbCrLf)
51600   .OptionsPDFGeneralCompatibility02 = Replace$(hLang.Retrieve("PDFGeneralCompatibility02", .OptionsPDFGeneralCompatibility02), "/n", vbCrLf)
51610   .OptionsPDFGeneralCompatibility03 = Replace$(hLang.Retrieve("PDFGeneralCompatibility03", .OptionsPDFGeneralCompatibility03), "/n", vbCrLf)
51620   .OptionsPDFGeneralCompatibility04 = Replace$(hLang.Retrieve("PDFGeneralCompatibility04", .OptionsPDFGeneralCompatibility04), "/n", vbCrLf)
51630   .OptionsPDFGeneralDefaultSettings = Replace$(hLang.Retrieve("PDFGeneralDefaultSettings", .OptionsPDFGeneralDefaultSettings), "/n", vbCrLf)
51640   .OptionsPDFGeneralDefaultSettingsDefault = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsDefault", .OptionsPDFGeneralDefaultSettingsDefault), "/n", vbCrLf)
51650   .OptionsPDFGeneralDefaultSettingsEbook = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsEbook", .OptionsPDFGeneralDefaultSettingsEbook), "/n", vbCrLf)
51660   .OptionsPDFGeneralDefaultSettingsPrepress = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsPrepress", .OptionsPDFGeneralDefaultSettingsPrepress), "/n", vbCrLf)
51670   .OptionsPDFGeneralDefaultSettingsPrinter = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsPrinter", .OptionsPDFGeneralDefaultSettingsPrinter), "/n", vbCrLf)
51680   .OptionsPDFGeneralDefaultSettingsScreen = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsScreen", .OptionsPDFGeneralDefaultSettingsScreen), "/n", vbCrLf)
51690   .OptionsPDFGeneralOverprint = Replace$(hLang.Retrieve("PDFGeneralOverprint", .OptionsPDFGeneralOverprint), "/n", vbCrLf)
51700   .OptionsPDFGeneralOverprint01 = Replace$(hLang.Retrieve("PDFGeneralOverprint01", .OptionsPDFGeneralOverprint01), "/n", vbCrLf)
51710   .OptionsPDFGeneralOverprint02 = Replace$(hLang.Retrieve("PDFGeneralOverprint02", .OptionsPDFGeneralOverprint02), "/n", vbCrLf)
51720   .OptionsPDFGeneralResolution = Replace$(hLang.Retrieve("PDFGeneralResolution", .OptionsPDFGeneralResolution), "/n", vbCrLf)
51730   .OptionsPDFGeneralRotate01 = Replace$(hLang.Retrieve("PDFGeneralRotate01", .OptionsPDFGeneralRotate01), "/n", vbCrLf)
51740   .OptionsPDFGeneralRotate02 = Replace$(hLang.Retrieve("PDFGeneralRotate02", .OptionsPDFGeneralRotate02), "/n", vbCrLf)
51750   .OptionsPDFGeneralRotate03 = Replace$(hLang.Retrieve("PDFGeneralRotate03", .OptionsPDFGeneralRotate03), "/n", vbCrLf)
51760   .OptionsPDFOptimize = Replace$(hLang.Retrieve("PDFOptimize", .OptionsPDFOptimize), "/n", vbCrLf)
51770   .OptionsPDFOptions = Replace$(hLang.Retrieve("PDFOptions", .OptionsPDFOptions), "/n", vbCrLf)
51780   .OptionsPDFOwnerPass = Replace$(hLang.Retrieve("PDFOwnerPass", .OptionsPDFOwnerPass), "/n", vbCrLf)
51790   .OptionsPDFPasswords = Replace$(hLang.Retrieve("PDFPasswords", .OptionsPDFPasswords), "/n", vbCrLf)
51800   .OptionsPDFRepeatPassword = Replace$(hLang.Retrieve("PDFRepeatPassword", .OptionsPDFRepeatPassword), "/n", vbCrLf)
51810   .OptionsPDFSecurity = Replace$(hLang.Retrieve("PDFSecurity", .OptionsPDFSecurity), "/n", vbCrLf)
51820   .OptionsPDFSecurityCaption = Replace$(hLang.Retrieve("PDFSecurityCaption", .OptionsPDFSecurityCaption), "/n", vbCrLf)
51830   .OptionsPDFSetPassword = Replace$(hLang.Retrieve("PDFSetPassword", .OptionsPDFSetPassword), "/n", vbCrLf)
51840   .OptionsPDFSigning = Replace$(hLang.Retrieve("PDFSigning", .OptionsPDFSigning), "/n", vbCrLf)
51850   .OptionsPDFSigningCaption = Replace$(hLang.Retrieve("PDFSigningCaption", .OptionsPDFSigningCaption), "/n", vbCrLf)
51860   .OptionsPDFSigningPfxFile = Replace$(hLang.Retrieve("PDFSigningPfxFile", .OptionsPDFSigningPfxFile), "/n", vbCrLf)
51870   .OptionsPDFSigningSignatureContact = Replace$(hLang.Retrieve("PDFSigningSignatureContact", .OptionsPDFSigningSignatureContact), "/n", vbCrLf)
51880   .OptionsPDFSigningSignatureLocation = Replace$(hLang.Retrieve("PDFSigningSignatureLocation", .OptionsPDFSigningSignatureLocation), "/n", vbCrLf)
51890   .OptionsPDFSigningSignatureMultiSignature = Replace$(hLang.Retrieve("PDFSigningSignatureMultiSignature", .OptionsPDFSigningSignatureMultiSignature), "/n", vbCrLf)
51900   .OptionsPDFSigningSignaturePosition = Replace$(hLang.Retrieve("PDFSigningSignaturePosition", .OptionsPDFSigningSignaturePosition), "/n", vbCrLf)
51910   .OptionsPDFSigningSignaturePositionLeftX = Replace$(hLang.Retrieve("PDFSigningSignaturePositionLeftX", .OptionsPDFSigningSignaturePositionLeftX), "/n", vbCrLf)
51920   .OptionsPDFSigningSignaturePositionLeftY = Replace$(hLang.Retrieve("PDFSigningSignaturePositionLeftY", .OptionsPDFSigningSignaturePositionLeftY), "/n", vbCrLf)
51930   .OptionsPDFSigningSignaturePositionRightX = Replace$(hLang.Retrieve("PDFSigningSignaturePositionRightX", .OptionsPDFSigningSignaturePositionRightX), "/n", vbCrLf)
51940   .OptionsPDFSigningSignaturePositionRightY = Replace$(hLang.Retrieve("PDFSigningSignaturePositionRightY", .OptionsPDFSigningSignaturePositionRightY), "/n", vbCrLf)
51950   .OptionsPDFSigningSignatureReason = Replace$(hLang.Retrieve("PDFSigningSignatureReason", .OptionsPDFSigningSignatureReason), "/n", vbCrLf)
51960   .OptionsPDFSigningSignatureVisible = Replace$(hLang.Retrieve("PDFSigningSignatureVisible", .OptionsPDFSigningSignatureVisible), "/n", vbCrLf)
51970   .OptionsPDFSigningSignPdfFile = Replace$(hLang.Retrieve("PDFSigningSignPdfFile", .OptionsPDFSigningSignPdfFile), "/n", vbCrLf)
51980   .OptionsPDFSymbol = Replace$(hLang.Retrieve("PDFSymbol", .OptionsPDFSymbol), "/n", vbCrLf)
51990   .OptionsPDFUserPass = Replace$(hLang.Retrieve("PDFUserPass", .OptionsPDFUserPass), "/n", vbCrLf)
52000   .OptionsPDFUseSecurity = Replace$(hLang.Retrieve("PDFUseSecurity", .OptionsPDFUseSecurity), "/n", vbCrLf)
52010   .OptionsPNGColorscount01 = Replace$(hLang.Retrieve("PNGColorscount01", .OptionsPNGColorscount01), "/n", vbCrLf)
52020   .OptionsPNGColorscount02 = Replace$(hLang.Retrieve("PNGColorscount02", .OptionsPNGColorscount02), "/n", vbCrLf)
52030   .OptionsPNGColorscount03 = Replace$(hLang.Retrieve("PNGColorscount03", .OptionsPNGColorscount03), "/n", vbCrLf)
52040   .OptionsPNGColorscount04 = Replace$(hLang.Retrieve("PNGColorscount04", .OptionsPNGColorscount04), "/n", vbCrLf)
52050   .OptionsPNGColorscount05 = Replace$(hLang.Retrieve("PNGColorscount05", .OptionsPNGColorscount05), "/n", vbCrLf)
52060   .OptionsPNGDescription = Replace$(hLang.Retrieve("PNGDescription", .OptionsPNGDescription), "/n", vbCrLf)
52070   .OptionsPNGFiles = Replace$(hLang.Retrieve("PNGFiles", .OptionsPNGFiles), "/n", vbCrLf)
52080   .OptionsPNGSymbol = Replace$(hLang.Retrieve("PNGSymbol", .OptionsPNGSymbol), "/n", vbCrLf)
52090   .OptionsPrintAfterSaving = Replace$(hLang.Retrieve("PrintAfterSaving", .OptionsPrintAfterSaving), "/n", vbCrLf)
52100   .OptionsPrintAfterSavingDuplex = Replace$(hLang.Retrieve("PrintAfterSavingDuplex", .OptionsPrintAfterSavingDuplex), "/n", vbCrLf)
52110   .OptionsPrintAfterSavingDuplexTumbleOff = Replace$(hLang.Retrieve("PrintAfterSavingDuplexTumbleOff", .OptionsPrintAfterSavingDuplexTumbleOff), "/n", vbCrLf)
52120   .OptionsPrintAfterSavingDuplexTumbleOn = Replace$(hLang.Retrieve("PrintAfterSavingDuplexTumbleOn", .OptionsPrintAfterSavingDuplexTumbleOn), "/n", vbCrLf)
52130   .OptionsPrintAfterSavingNoCancel = Replace$(hLang.Retrieve("PrintAfterSavingNoCancel", .OptionsPrintAfterSavingNoCancel), "/n", vbCrLf)
52140   .OptionsPrintAfterSavingPrinter = Replace$(hLang.Retrieve("PrintAfterSavingPrinter", .OptionsPrintAfterSavingPrinter), "/n", vbCrLf)
52150   .OptionsPrintAfterSavingQueryUser = Replace$(hLang.Retrieve("PrintAfterSavingQueryUser", .OptionsPrintAfterSavingQueryUser), "/n", vbCrLf)
52160   .OptionsPrintAfterSavingQueryUserDefaultPrinter = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserDefaultPrinter", .OptionsPrintAfterSavingQueryUserDefaultPrinter), "/n", vbCrLf)
52170   .OptionsPrintAfterSavingQueryUserOff = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserOff", .OptionsPrintAfterSavingQueryUserOff), "/n", vbCrLf)
52180   .OptionsPrintAfterSavingQueryUserPrinterSetupDialog = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserPrinterSetupDialog", .OptionsPrintAfterSavingQueryUserPrinterSetupDialog), "/n", vbCrLf)
52190   .OptionsPrintAfterSavingQueryUserStandardPrinterDialog = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserStandardPrinterDialog", .OptionsPrintAfterSavingQueryUserStandardPrinterDialog), "/n", vbCrLf)
52200   .OptionsPrintertempDirectoryPrompt = Replace$(hLang.Retrieve("PrintertempDirectoryPrompt", .OptionsPrintertempDirectoryPrompt), "/n", vbCrLf)
52210   .OptionsPrintTestpage = Replace$(hLang.Retrieve("PrintTestpage", .OptionsPrintTestpage), "/n", vbCrLf)
52220   .OptionsProcesspriority = Replace$(hLang.Retrieve("Processpriority", .OptionsProcesspriority), "/n", vbCrLf)
52230   .OptionsProcesspriorityHigh = Replace$(hLang.Retrieve("ProcesspriorityHigh", .OptionsProcesspriorityHigh), "/n", vbCrLf)
52240   .OptionsProcesspriorityIdle = Replace$(hLang.Retrieve("ProcesspriorityIdle", .OptionsProcesspriorityIdle), "/n", vbCrLf)
52250   .OptionsProcesspriorityNormal = Replace$(hLang.Retrieve("ProcesspriorityNormal", .OptionsProcesspriorityNormal), "/n", vbCrLf)
52260   .OptionsProcesspriorityRealtime = Replace$(hLang.Retrieve("ProcesspriorityRealtime", .OptionsProcesspriorityRealtime), "/n", vbCrLf)
52270   .OptionsProgramActionsDescription = Replace$(hLang.Retrieve("ProgramActionsDescription", .OptionsProgramActionsDescription), "/n", vbCrLf)
52280   .OptionsProgramActionsSymbol = Replace$(hLang.Retrieve("ProgramActionsSymbol", .OptionsProgramActionsSymbol), "/n", vbCrLf)
52290   .OptionsProgramAutosaveDescription = Replace$(hLang.Retrieve("ProgramAutosaveDescription", .OptionsProgramAutosaveDescription), "/n", vbCrLf)
52300   .OptionsProgramAutosaveSymbol = Replace$(hLang.Retrieve("ProgramAutosaveSymbol", .OptionsProgramAutosaveSymbol), "/n", vbCrLf)
52310   .OptionsProgramDirectoriesDescription = Replace$(hLang.Retrieve("ProgramDirectoriesDescription", .OptionsProgramDirectoriesDescription), "/n", vbCrLf)
52320   .OptionsProgramDirectoriesSymbol = Replace$(hLang.Retrieve("ProgramDirectoriesSymbol", .OptionsProgramDirectoriesSymbol), "/n", vbCrLf)
52330   .OptionsProgramDocumentDescription = Replace$(hLang.Retrieve("ProgramDocumentDescription", .OptionsProgramDocumentDescription), "/n", vbCrLf)
52340   .OptionsProgramDocumentDescription1 = Replace$(hLang.Retrieve("ProgramDocumentDescription1", .OptionsProgramDocumentDescription1), "/n", vbCrLf)
52350   .OptionsProgramDocumentDescription2 = Replace$(hLang.Retrieve("ProgramDocumentDescription2", .OptionsProgramDocumentDescription2), "/n", vbCrLf)
52360   .OptionsProgramDocumentSymbol = Replace$(hLang.Retrieve("ProgramDocumentSymbol", .OptionsProgramDocumentSymbol), "/n", vbCrLf)
52370   .OptionsProgramFont = Replace$(hLang.Retrieve("ProgramFont", .OptionsProgramFont), "/n", vbCrLf)
52380   .OptionsProgramFontCancelTest = Replace$(hLang.Retrieve("ProgramFontCancelTest", .OptionsProgramFontCancelTest), "/n", vbCrLf)
52390   .OptionsProgramFontcharset = Replace$(hLang.Retrieve("ProgramFontcharset", .OptionsProgramFontcharset), "/n", vbCrLf)
52400   .OptionsProgramFontDescription = Replace$(hLang.Retrieve("ProgramFontDescription", .OptionsProgramFontDescription), "/n", vbCrLf)
52410   .OptionsProgramFontSize = Replace$(hLang.Retrieve("ProgramFontSize", .OptionsProgramFontSize), "/n", vbCrLf)
52420   .OptionsProgramFontSymbol = Replace$(hLang.Retrieve("ProgramFontSymbol", .OptionsProgramFontSymbol), "/n", vbCrLf)
52430   .OptionsProgramFontTest = Replace$(hLang.Retrieve("ProgramFontTest", .OptionsProgramFontTest), "/n", vbCrLf)
52440   .OptionsProgramFontTestdescription = Replace$(hLang.Retrieve("ProgramFontTestdescription", .OptionsProgramFontTestdescription), "/n", vbCrLf)
52450   .OptionsProgramGeneralDescription = Replace$(hLang.Retrieve("ProgramGeneralDescription", .OptionsProgramGeneralDescription), "/n", vbCrLf)
52460   .OptionsProgramGeneralDescription1 = Replace$(hLang.Retrieve("ProgramGeneralDescription1", .OptionsProgramGeneralDescription1), "/n", vbCrLf)
52470   .OptionsProgramGeneralDescription2 = Replace$(hLang.Retrieve("ProgramGeneralDescription2", .OptionsProgramGeneralDescription2), "/n", vbCrLf)
52480   .OptionsProgramGeneralSymbol = Replace$(hLang.Retrieve("ProgramGeneralSymbol", .OptionsProgramGeneralSymbol), "/n", vbCrLf)
52490   .OptionsProgramGhostscriptDescription = Replace$(hLang.Retrieve("ProgramGhostscriptDescription", .OptionsProgramGhostscriptDescription), "/n", vbCrLf)
52500   .OptionsProgramGhostscriptSymbol = Replace$(hLang.Retrieve("ProgramGhostscriptSymbol", .OptionsProgramGhostscriptSymbol), "/n", vbCrLf)
52510   .OptionsProgramLanguagesDescription = Replace$(hLang.Retrieve("ProgramLanguagesDescription", .OptionsProgramLanguagesDescription), "/n", vbCrLf)
52520   .OptionsProgramLanguagesSymbol = Replace$(hLang.Retrieve("ProgramLanguagesSymbol", .OptionsProgramLanguagesSymbol), "/n", vbCrLf)
52530   .OptionsProgramNoProcessingAtStartup = Replace$(hLang.Retrieve("ProgramNoProcessingAtStartup", .OptionsProgramNoProcessingAtStartup), "/n", vbCrLf)
52540   .OptionsProgramOptionsDesign = Replace$(hLang.Retrieve("ProgramOptionsDesign", .OptionsProgramOptionsDesign), "/n", vbCrLf)
52550   .OptionsProgramOptionsDesignGradient = Replace$(hLang.Retrieve("ProgramOptionsDesignGradient", .OptionsProgramOptionsDesignGradient), "/n", vbCrLf)
52560   .OptionsProgramOptionsDesignSimple = Replace$(hLang.Retrieve("ProgramOptionsDesignSimple", .OptionsProgramOptionsDesignSimple), "/n", vbCrLf)
52570   .OptionsProgramPrintDescription = Replace$(hLang.Retrieve("ProgramPrintDescription", .OptionsProgramPrintDescription), "/n", vbCrLf)
52580   .OptionsProgramPrintSymbol = Replace$(hLang.Retrieve("ProgramPrintSymbol", .OptionsProgramPrintSymbol), "/n", vbCrLf)
52590   .OptionsProgramRunProgramAfterSavingCaption = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingCaption", .OptionsProgramRunProgramAfterSavingCaption), "/n", vbCrLf)
52600   .OptionsProgramRunProgramAfterSavingProgram = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingProgram", .OptionsProgramRunProgramAfterSavingProgram), "/n", vbCrLf)
52610   .OptionsProgramRunProgramAfterSavingProgramParameters = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingProgramParameters", .OptionsProgramRunProgramAfterSavingProgramParameters), "/n", vbCrLf)
52620   .OptionsProgramRunProgramAfterSavingWaitUntilReady = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWaitUntilReady", .OptionsProgramRunProgramAfterSavingWaitUntilReady), "/n", vbCrLf)
52630   .OptionsProgramRunProgramAfterSavingWindowstyle = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyle", .OptionsProgramRunProgramAfterSavingWindowstyle), "/n", vbCrLf)
52640   .OptionsProgramRunProgramAfterSavingWindowstyleHide = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleHide", .OptionsProgramRunProgramAfterSavingWindowstyleHide), "/n", vbCrLf)
52650   .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleMaximizedFocus", .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus), "/n", vbCrLf)
52660   .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleMinimizedFocus", .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus), "/n", vbCrLf)
52670   .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus", .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus), "/n", vbCrLf)
52680   .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleNormalFocus", .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus), "/n", vbCrLf)
52690   .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleNormalNoFocus", .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus), "/n", vbCrLf)
52700   .OptionsProgramRunProgramBeforeSavingCaption = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingCaption", .OptionsProgramRunProgramBeforeSavingCaption), "/n", vbCrLf)
52710   .OptionsProgramRunProgramBeforeSavingProgram = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingProgram", .OptionsProgramRunProgramBeforeSavingProgram), "/n", vbCrLf)
52720   .OptionsProgramRunProgramBeforeSavingProgramParameters = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingProgramParameters", .OptionsProgramRunProgramBeforeSavingProgramParameters), "/n", vbCrLf)
52730   .OptionsProgramRunProgramBeforeSavingWaitUntilReady = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWaitUntilReady", .OptionsProgramRunProgramBeforeSavingWaitUntilReady), "/n", vbCrLf)
52740   .OptionsProgramRunProgramBeforeSavingWindowstyle = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyle", .OptionsProgramRunProgramBeforeSavingWindowstyle), "/n", vbCrLf)
52750   .OptionsProgramRunProgramBeforeSavingWindowstyleHide = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleHide", .OptionsProgramRunProgramBeforeSavingWindowstyleHide), "/n", vbCrLf)
52760   .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleMaximizedFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus), "/n", vbCrLf)
52770   .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleMinimizedFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus), "/n", vbCrLf)
52780   .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus), "/n", vbCrLf)
52790   .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleNormalFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus), "/n", vbCrLf)
52800   .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleNormalNoFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus), "/n", vbCrLf)
52810   .OptionsProgramSaveDescription = Replace$(hLang.Retrieve("ProgramSaveDescription", .OptionsProgramSaveDescription), "/n", vbCrLf)
52820   .OptionsProgramSaveSymbol = Replace$(hLang.Retrieve("ProgramSaveSymbol", .OptionsProgramSaveSymbol), "/n", vbCrLf)
52830   .OptionsProgramShowAnimation = Replace$(hLang.Retrieve("ProgramShowAnimation", .OptionsProgramShowAnimation), "/n", vbCrLf)
52840   .OptionsProgramSwitchingDefaultprinter = Replace$(hLang.Retrieve("ProgramSwitchingDefaultprinter", .OptionsProgramSwitchingDefaultprinter), "/n", vbCrLf)
52850   .OptionsPSDColorsCount01 = Replace$(hLang.Retrieve("PSDColorsCount01", .OptionsPSDColorsCount01), "/n", vbCrLf)
52860   .OptionsPSDColorscount02 = Replace$(hLang.Retrieve("PSDColorscount02", .OptionsPSDColorscount02), "/n", vbCrLf)
52870   .OptionsPSDDescription = Replace$(hLang.Retrieve("PSDDescription", .OptionsPSDDescription), "/n", vbCrLf)
52880   .OptionsPSDescription = Replace$(hLang.Retrieve("PSDescription", .OptionsPSDescription), "/n", vbCrLf)
52890   .OptionsPSDSymbol = Replace$(hLang.Retrieve("PSDSymbol", .OptionsPSDSymbol), "/n", vbCrLf)
52900   .OptionsPSFiles = Replace$(hLang.Retrieve("PSFiles", .OptionsPSFiles), "/n", vbCrLf)
52910   .OptionsPSLanguageLevel = Replace$(hLang.Retrieve("PSLanguageLevel", .OptionsPSLanguageLevel), "/n", vbCrLf)
52920   .OptionsPSSymbol = Replace$(hLang.Retrieve("PSSymbol", .OptionsPSSymbol), "/n", vbCrLf)
52930   .OptionsRAWColorsCount01 = Replace$(hLang.Retrieve("RAWColorsCount01", .OptionsRAWColorsCount01), "/n", vbCrLf)
52940   .OptionsRAWColorscount02 = Replace$(hLang.Retrieve("RAWColorscount02", .OptionsRAWColorscount02), "/n", vbCrLf)
52950   .OptionsRAWColorscount03 = Replace$(hLang.Retrieve("RAWColorscount03", .OptionsRAWColorscount03), "/n", vbCrLf)
52960   .OptionsRAWDescription = Replace$(hLang.Retrieve("RAWDescription", .OptionsRAWDescription), "/n", vbCrLf)
52970   .OptionsRAWSymbol = Replace$(hLang.Retrieve("RAWSymbol", .OptionsRAWSymbol), "/n", vbCrLf)
52980   .OptionsRemoveSpaces = Replace$(hLang.Retrieve("RemoveSpaces", .OptionsRemoveSpaces), "/n", vbCrLf)
52990   .OptionsReset = Replace$(hLang.Retrieve("Reset", .OptionsReset), "/n", vbCrLf)
53000   .OptionsSave = Replace$(hLang.Retrieve("Save", .OptionsSave), "/n", vbCrLf)
53010   .OptionsSaveFilename = Replace$(hLang.Retrieve("SaveFilename", .OptionsSaveFilename), "/n", vbCrLf)
53020   .OptionsSaveFilenameAdd = Replace$(hLang.Retrieve("SaveFilenameAdd", .OptionsSaveFilenameAdd), "/n", vbCrLf)
53030   .OptionsSaveFilenameChange = Replace$(hLang.Retrieve("SaveFilenameChange", .OptionsSaveFilenameChange), "/n", vbCrLf)
53040   .OptionsSaveFilenameDelete = Replace$(hLang.Retrieve("SaveFilenameDelete", .OptionsSaveFilenameDelete), "/n", vbCrLf)
53050   .OptionsSaveFilenameSubstitutions = Replace$(hLang.Retrieve("SaveFilenameSubstitutions", .OptionsSaveFilenameSubstitutions), "/n", vbCrLf)
53060   .OptionsSaveFilenameSubstitutionsTitle = Replace$(hLang.Retrieve("SaveFilenameSubstitutionsTitle", .OptionsSaveFilenameSubstitutionsTitle), "/n", vbCrLf)
53070   .OptionsSaveFilenameTokens = Replace$(hLang.Retrieve("SaveFilenameTokens", .OptionsSaveFilenameTokens), "/n", vbCrLf)
53080   .OptionsSavePasswords = Replace$(hLang.Retrieve("SavePasswords", .OptionsSavePasswords), "/n", vbCrLf)
53090   .OptionsSendEmailAfterAutosave = Replace$(hLang.Retrieve("SendEmailAfterAutosave", .OptionsSendEmailAfterAutosave), "/n", vbCrLf)
53100   .OptionsSendMailMethod = Replace$(hLang.Retrieve("SendMailMethod", .OptionsSendMailMethod), "/n", vbCrLf)
53110   .OptionsSendMailMethodAutomatic = Replace$(hLang.Retrieve("SendMailMethodAutomatic", .OptionsSendMailMethodAutomatic), "/n", vbCrLf)
53120   .OptionsSendMailMethodMapi = Replace$(hLang.Retrieve("SendMailMethodMapi", .OptionsSendMailMethodMapi), "/n", vbCrLf)
53130   .OptionsSendMailMethodSendmailDLL = Replace$(hLang.Retrieve("SendMailMethodSendmailDLL", .OptionsSendMailMethodSendmailDLL), "/n", vbCrLf)
53140   .OptionsShellIntegration = Replace$(hLang.Retrieve("ShellIntegration", .OptionsShellIntegration), "/n", vbCrLf)
53150   .OptionsShellIntegrationAdd = Replace$(hLang.Retrieve("ShellIntegrationAdd", .OptionsShellIntegrationAdd), "/n", vbCrLf)
53160   .OptionsShellIntegrationCaption = Replace$(hLang.Retrieve("ShellIntegrationCaption", .OptionsShellIntegrationCaption), "/n", vbCrLf)
53170   .OptionsShellIntegrationRemove = Replace$(hLang.Retrieve("ShellIntegrationRemove", .OptionsShellIntegrationRemove), "/n", vbCrLf)
53180   .OptionsStamp = Replace$(hLang.Retrieve("Stamp", .OptionsStamp), "/n", vbCrLf)
53190   .OptionsStampFontColor = Replace$(hLang.Retrieve("StampFontColor", .OptionsStampFontColor), "/n", vbCrLf)
53200   .OptionsStampOutlineFontThickness = Replace$(hLang.Retrieve("StampOutlineFontThickness", .OptionsStampOutlineFontThickness), "/n", vbCrLf)
53210   .OptionsStampString = Replace$(hLang.Retrieve("StampString", .OptionsStampString), "/n", vbCrLf)
53220   .OptionsStampUseOutlineFont = Replace$(hLang.Retrieve("StampUseOutlineFont", .OptionsStampUseOutlineFont), "/n", vbCrLf)
53230   .OptionsStandardAuthorToken = Replace$(hLang.Retrieve("StandardAuthorToken", .OptionsStandardAuthorToken), "/n", vbCrLf)
53240   .OptionsStandardSaveFormat = Replace$(hLang.Retrieve("StandardSaveFormat", .OptionsStandardSaveFormat), "/n", vbCrLf)
53250   .OptionsTestpage = Replace$(hLang.Retrieve("Testpage", .OptionsTestpage), "/n", vbCrLf)
53260   .OptionsTIFFColorscount01 = Replace$(hLang.Retrieve("TIFFColorscount01", .OptionsTIFFColorscount01), "/n", vbCrLf)
53270   .OptionsTIFFColorscount02 = Replace$(hLang.Retrieve("TIFFColorscount02", .OptionsTIFFColorscount02), "/n", vbCrLf)
53280   .OptionsTIFFColorscount03 = Replace$(hLang.Retrieve("TIFFColorscount03", .OptionsTIFFColorscount03), "/n", vbCrLf)
53290   .OptionsTIFFColorscount04 = Replace$(hLang.Retrieve("TIFFColorscount04", .OptionsTIFFColorscount04), "/n", vbCrLf)
53300   .OptionsTIFFColorscount05 = Replace$(hLang.Retrieve("TIFFColorscount05", .OptionsTIFFColorscount05), "/n", vbCrLf)
53310   .OptionsTIFFColorscount06 = Replace$(hLang.Retrieve("TIFFColorscount06", .OptionsTIFFColorscount06), "/n", vbCrLf)
53320   .OptionsTIFFColorscount07 = Replace$(hLang.Retrieve("TIFFColorscount07", .OptionsTIFFColorscount07), "/n", vbCrLf)
53330   .OptionsTIFFColorscount08 = Replace$(hLang.Retrieve("TIFFColorscount08", .OptionsTIFFColorscount08), "/n", vbCrLf)
53340   .OptionsTIFFDescription = Replace$(hLang.Retrieve("TIFFDescription", .OptionsTIFFDescription), "/n", vbCrLf)
53350   .OptionsTIFFSymbol = Replace$(hLang.Retrieve("TIFFSymbol", .OptionsTIFFSymbol), "/n", vbCrLf)
53360   .OptionsTreeFormats = Replace$(hLang.Retrieve("TreeFormats", .OptionsTreeFormats), "/n", vbCrLf)
53370   .OptionsTreeProgram = Replace$(hLang.Retrieve("TreeProgram", .OptionsTreeProgram), "/n", vbCrLf)
53380   .OptionsTXTDescription = Replace$(hLang.Retrieve("TXTDescription", .OptionsTXTDescription), "/n", vbCrLf)
53390   .OptionsTXTSymbol = Replace$(hLang.Retrieve("TXTSymbol", .OptionsTXTSymbol), "/n", vbCrLf)
53400   .OptionsUseAutosave = Replace$(hLang.Retrieve("UseAutosave", .OptionsUseAutosave), "/n", vbCrLf)
53410   .OptionsUseAutosaveDirectory = Replace$(hLang.Retrieve("UseAutosaveDirectory", .OptionsUseAutosaveDirectory), "/n", vbCrLf)
53420   .OptionsUseCreationDateNow = Replace$(hLang.Retrieve("UseCreationDateNow", .OptionsUseCreationDateNow), "/n", vbCrLf)
53430   .OptionsUseCustomPapersize = Replace$(hLang.Retrieve("UseCustomPapersize", .OptionsUseCustomPapersize), "/n", vbCrLf)
53440   .OptionsUseFixPapersize = Replace$(hLang.Retrieve("UseFixPapersize", .OptionsUseFixPapersize), "/n", vbCrLf)
53450   .OptionsUserPass = Replace$(hLang.Retrieve("UserPass", .OptionsUserPass), "/n", vbCrLf)
53460   .OptionsUseStandardauthor = Replace$(hLang.Retrieve("UseStandardauthor", .OptionsUseStandardauthor), "/n", vbCrLf)
53470   .OptionsXCFColorsCount01 = Replace$(hLang.Retrieve("XCFColorsCount01", .OptionsXCFColorsCount01), "/n", vbCrLf)
53480   .OptionsXCFColorscount02 = Replace$(hLang.Retrieve("XCFColorscount02", .OptionsXCFColorscount02), "/n", vbCrLf)
53490   .OptionsXCFDescription = Replace$(hLang.Retrieve("XCFDescription", .OptionsXCFDescription), "/n", vbCrLf)
53500   .OptionsXCFSymbol = Replace$(hLang.Retrieve("XCFSymbol", .OptionsXCFSymbol), "/n", vbCrLf)
53510  End With
53520  Set hLang = Nothing
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
50160   .PrintingPCLFiles = Replace$(hLang.Retrieve("PCLFiles", .PrintingPCLFiles), "/n", vbCrLf)
50170   .PrintingPCXFiles = Replace$(hLang.Retrieve("PCXFiles", .PrintingPCXFiles), "/n", vbCrLf)
50180   .PrintingPDFAFiles = Replace$(hLang.Retrieve("PDFAFiles", .PrintingPDFAFiles), "/n", vbCrLf)
50190   .PrintingPDFFiles = Replace$(hLang.Retrieve("PDFFiles", .PrintingPDFFiles), "/n", vbCrLf)
50200   .PrintingPDFXFiles = Replace$(hLang.Retrieve("PDFXFiles", .PrintingPDFXFiles), "/n", vbCrLf)
50210   .PrintingPNGFiles = Replace$(hLang.Retrieve("PNGFiles", .PrintingPNGFiles), "/n", vbCrLf)
50220   .PrintingPSDFiles = Replace$(hLang.Retrieve("PSDFiles", .PrintingPSDFiles), "/n", vbCrLf)
50230   .PrintingPSFiles = Replace$(hLang.Retrieve("PSFiles", .PrintingPSFiles), "/n", vbCrLf)
50240   .PrintingRAWFiles = Replace$(hLang.Retrieve("RAWFiles", .PrintingRAWFiles), "/n", vbCrLf)
50250   .PrintingSave = Replace$(hLang.Retrieve("Save", .PrintingSave), "/n", vbCrLf)
50260   .PrintingStartStandardProgram = Replace$(hLang.Retrieve("StartStandardProgram", .PrintingStartStandardProgram), "/n", vbCrLf)
50270   .PrintingStatus = Replace$(hLang.Retrieve("Status", .PrintingStatus), "/n", vbCrLf)
50280   .PrintingSubject = Replace$(hLang.Retrieve("Subject", .PrintingSubject), "/n", vbCrLf)
50290   .PrintingTIFFFiles = Replace$(hLang.Retrieve("TIFFFiles", .PrintingTIFFFiles), "/n", vbCrLf)
50300   .PrintingTXTFiles = Replace$(hLang.Retrieve("TXTFiles", .PrintingTXTFiles), "/n", vbCrLf)
50310   .PrintingXCFFiles = Replace$(hLang.Retrieve("XCFFiles", .PrintingXCFFiles), "/n", vbCrLf)
50320  End With
50330  Set hLang = Nothing
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
50050   .CommonVersion = "0.9.7"
50060
50070   .DialogDocument = "&Document"
50080   .DialogDocumentAdd = "Add"
50090   .DialogDocumentAddFromClipboard = "Add from clipboard"
50100   .DialogDocumentBottom = "Bottom"
50110   .DialogDocumentCombine = "Combine"
50120   .DialogDocumentCombineAll = "Combine all"
50130   .DialogDocumentCombineAllSend = "Combine all and send"
50140   .DialogDocumentDelete = "Delete"
50150   .DialogDocumentDown = "Down"
50160   .DialogDocumentPrint = "Print"
50170   .DialogDocumentSave = "Save"
50180   .DialogDocumentSend = "Send"
50190   .DialogDocumentTop = "Top"
50200   .DialogDocumentUp = "Up"
50210   .DialogEmailAddress = "Email address"
50220   .DialogInfo = "&?"
50230   .DialogInfoCheckUpdates = "Check for Updates"
50240   .DialogInfoHomepage = "Product Homepage"
50250   .DialogInfoInfo = "About"
50260   .DialogInfoPaypal = "Paypal"
50270   .DialogInfoPDFCreatorSourceforge = "PDFCreator on Sourceforge"
50280   .DialogLanguage = "&Language"
50290   .DialogPrinter = "&Printer"
50300   .DialogPrinterClose = "Close"
50310   .DialogPrinterLogfile = "Logfile"
50320   .DialogPrinterLogfiles = "Logfiles"
50330   .DialogPrinterLogging = "Logging"
50340   .DialogPrinterOptions = "Options"
50350   .DialogPrinterPrinterStop = "Printer stop"
50360   .DialogView = "&View"
50370   .DialogViewStatusbar = "Status Bar"
50380   .DialogViewToolbars = "&Toolbars"
50390   .DialogViewToolbarsEmail = "Email"
50400   .DialogViewToolbarsStandard = "Standard"
50410
50420   .ListAddFile = "Add a file"
50430   .ListAllFiles = "All files"
50440   .ListBytes = "Bytes"
50450   .ListDate = "Created on"
50460   .ListDocumenttitle = "Document Title"
50470   .ListFilename = "Filename"
50480   .ListGBytes = "GBytes"
50490   .ListKBytes = "kBytes"
50500   .ListMBytes = "MBytes"
50510   .ListPDFFiles = "PDF Files"
50520   .ListPostscriptFiles = "PostScript Files"
50530   .ListPrinting = "Printing"
50540   .ListSize = "Size"
50550   .ListStatus = "Status"
50560   .ListWaiting = "Waiting"
50570
50580   .LoggingClear = "Cl&ear"
50590   .LoggingClose = "&Close"
50600   .LoggingLogfile = "Logfile"
50610
50620   .MessagesMsg01 = "Document in queue."
50630   .MessagesMsg02 = "Documents in queue."
50640   .MessagesMsg03 = "Do you wish to reset all settings?"
50650   .MessagesMsg04 = "Error: Cannot send Email!"
50660   .MessagesMsg05 = "File already exists. Do you want to overwrite it?"
50670   .MessagesMsg06 = "This file does not seem to be a postscript file!"
50680   .MessagesMsg07 = "There is a problem when trying to access this drive or directory!"
50690   .MessagesMsg08 = "Cannot find gsdll32.dll. Please check the ghostscript-program directory (see options)!"
50700   .MessagesMsg09 = "The output path does not exist. Do you want to create it?"
50710   .MessagesMsg10 = "This is not a valid path!"
50720   .MessagesMsg11 = "There is already such an entry!"
50730   .MessagesMsg12 = "Please don't use these forbidden characters for a filename!"
50740   .MessagesMsg13 = "Delete all program settings?"
50750   .MessagesMsg14 = "The file can not be found!"
50760   .MessagesMsg15 = "Cannot find gsdll32.dll in this directory!"
50770   .MessagesMsg16 = "No ghostscript font found in this directory!"
50780   .MessagesMsg17 = "No files in this directory!"
50790   .MessagesMsg19 = "You need either pdfenc or AFPL Ghostscript greater than, or equal to, version 8.14!"
50800   .MessagesMsg20 = "There was a problem sending an email with the standard emailclient!"
50810   .MessagesMsg21 = "User passwords do not match!"
50820   .MessagesMsg22 = "Owner passwords do not match!"
50830   .MessagesMsg23 = "The document is not protected!"
50840   .MessagesMsg24 = "The user password is empty! Continue?"
50850   .MessagesMsg25 = "The owner password is empty! Continue?"
50860   .MessagesMsg26 = "Unknown error"
50870   .MessagesMsg27 = "Cannot find the file/page."
50880   .MessagesMsg28 = "The filesize is 0 byte."
50890   .MessagesMsg29 = "Server not found."
50900   .MessagesMsg30 = "The url isn not interpretable."
50910   .MessagesMsg31 = "An error has occured"
50920   .MessagesMsg32 = "The new version %1 is available. Would you like download the new version from the Sourceforge pages?"
50930   .MessagesMsg33 = "You already have the most recent version."
50940   .MessagesMsg34 = "The file is in use. Please close the file first or choose another filename."
50950   .MessagesMsg35 = "It is necessary to temporarily set PDFCreator as defaultprinter."
50960   .MessagesMsg36 = "Don't ask me again."
50970   .MessagesMsg37 = "The downloaded file is not a valid language file!"
50980   .MessagesMsg38 = "The language file has been successfully installed!"
50990   .MessagesMsg39 = "pdfforge.dll is not installed! You can find more information in the help file."
51000
51010   .OptionsAdditionalGhostscriptParameters = "Additional Ghostscript parameters"
51020   .OptionsAdditionalGhostscriptSearchpath = "Additional Ghostscript searchpath"
51030   .OptionsAddWindowsFontpath = "Use Windows fonts"
51040   .OptionsAssociatePSFiles = "Associate PDFCreator with postscript files"
51050   .OptionsAutosaveDirectoryPrompt = "Select Autosave Directory"
51060   .OptionsAutosaveFilename = "Filename"
51070   .OptionsAutosaveFilenameTokens = "Add a Filename-Token"
51080   .OptionsAutosaveFormat = "Autosave format"
51090   .OptionsAutosaveStartStandardProgram = "After auto-saving open the document with the default program."
51100   .OptionsBitmapResolution = "Resolution"
51110   .OptionsBMPColorscount01 = "4294967296 colors (32 Bit)"
51120   .OptionsBMPColorscount02 = "16777216 colors (24 Bit)"
51130   .OptionsBMPColorscount03 = "256 colors (8 Bit)"
51140   .OptionsBMPColorscount04 = "16 colors (4 Bit)"
51150   .OptionsBMPColorscount05 = "8 colors (3 Bit)"
51160   .OptionsBMPColorscount06 = "2 colors (Black/White)"
51170   .OptionsBMPColorscount07 = "Greyscale (8 Bit)"
51180   .OptionsBMPDescription = "Windows Bitmap Format. Please use only for single pages."
51190   .OptionsBMPSymbol = "BMP"
51200   .OptionsBrowserAddOn = "Browser Add On"
51210   .OptionsBrowserAddOnInstall = "Install PDFCreator Browser Add On"
51220   .OptionsCancel = "&Cancel"
51230   .OptionsCheckUpdateDescription = "Check update"
51240   .OptionsCheckUpdateInterval = "Update interval"
51250   .OptionsCheckUpdateInterval01 = "Never"
51260   .OptionsCheckUpdateInterval02 = "Once a day"
51270   .OptionsCheckUpdateInterval03 = "Once a week"
51280   .OptionsCheckUpdateInterval04 = "Once a month"
51290   .OptionsCheckUpdateNow = "Check now"
51300   .OptionsCustomPapersizeHeight = "Height"
51310   .OptionsCustomPapersizeInfo = "Units of 1/72 of an inch."
51320   .OptionsCustomPapersizeWidth = "Width"
51330   .OptionsDirectoriesGSBin = "Ghostscript Binaries"
51340   .OptionsDirectoriesGSFonts = "Ghostscript Fonts"
51350   .OptionsDirectoriesGSLibraries = "Ghostscript Libraries"
51360   .OptionsDirectoriesTempPath = "Temporary Files"
51370   .OptionsDocument = "Document"
51380   .OptionsEPSDescription = "Encapsulated Postscript Format"
51390   .OptionsEPSFiles = "Encapsulated Postscript-Files"
51400   .OptionsEPSSymbol = "EPS"
51410   .OptionsGhostscriptBinariesDirectoryPrompt = "Select Ghostscript Binaries Directory"
51420   .OptionsGhostscriptFontsDirectoryPrompt = "Select Ghostscript Fonts Directory"
51430   .OptionsGhostscriptInternal = "Internal Ghostscript: %1 Ghostscript %2"
51440   .OptionsGhostscriptLibrariesDirectoryPrompt = "Select Ghostscript Libraries Directory"
51450   .OptionsGhostscriptResourceDirectoryPrompt = "Select Ghostscript Resource Directory"
51460   .OptionsGhostscriptversion = "Ghostscript Version"
51470   .OptionsImageSettings = "Settings"
51480   .OptionsJavaPath = "Path to Java Interpreter"
51490   .OptionsJPEGColorscount01 = "16777216 colors (24 Bit)"
51500   .OptionsJPEGColorscount02 = "Greyscale (8 Bit)"
51510   .OptionsJPEGDescription = "JPEG (JFIF) Format. Please use only for single pages."
51520   .OptionsJPEGQuality = "Quality:"
51530   .OptionsJPEGSymbol = "JPEG"
51540   .OptionsLanguagesCurrentLanguage = "Current language"
51550   .OptionsLanguagesDownloadMoreLanguages = "Load more languages from the internet"
51560   .OptionsLanguagesInstall = "Install"
51570   .OptionsLanguagesRefresh = "Refresh List"
51580   .OptionsLanguagesTranslation = "Translation"
51590   .OptionsLanguagesVersion = "Version"
51600   .OptionsNothingToConfigure = "There is nothing to configure."
51610   .OptionsOnePagePerFile = "One page per file (not for pdf and eps files)"
51620   .OptionsOwnerPass = "Owner Password"
51630   .OptionsPassCancel = "Cancel"
51640   .OptionsPassOK = "OK"
51650   .OptionsPCLColorscount01 = "16777216 colors (24bit)"
51660   .OptionsPCLColorscount02 = "2 colors (Black/White)"
51670   .OptionsPCLDescription = "HP PCL-XL Format"
51680   .OptionsPCLSymbol = "PCL"
51690   .OptionsPCXColorscount01 = "4294967296 colors (32 Bit) CMYK"
51700   .OptionsPCXColorscount02 = "16777216 colors (24 Bit)"
51710   .OptionsPCXColorscount03 = "256 colors (8 Bit)"
51720   .OptionsPCXColorscount04 = "16 colors (4 Bit)"
51730   .OptionsPCXColorscount05 = "2 colors (Black\White)"
51740   .OptionsPCXColorscount06 = "Greyscale (8 Bit)"
51750   .OptionsPCXDescription = "PCX Format. Please use only for single pages."
51760   .OptionsPCXSymbol = "PCX"
51770   .OptionsPDFAllowAssembly = "Allow changes to the assembly"
51780   .OptionsPDFAllowDegradedPrinting = "Allow printing in low resolution"
51790   .OptionsPDFAllowFillIn = "Allow filling in form fields"
51800   .OptionsPDFAllowScreenReaders = "Allow screen readers"
51810   .OptionsPDFColors = "Colors"
51820   .OptionsPDFColorsCaption = "Color Options"
51830   .OptionsPDFColorsCMYKtoRGB = "Convert CMYK images to RGB"
51840   .OptionsPDFColorsColorModel01 = "Use Color Model Device RGB"
51850   .OptionsPDFColorsColorModel02 = "Use Color Model Device CMYK"
51860   .OptionsPDFColorsColorModel03 = "Use Color Model Device Grayscale"
51870   .OptionsPDFColorsColorOptions = "Options"
51880   .OptionsPDFColorsPreserveHalftone = "Preserve Halftone Information"
51890   .OptionsPDFColorsPreserveOverprint = "Preserve Overprint Settings"
51900   .OptionsPDFColorsPreserveTransfer = "Preserve Transfer Functions"
51910   .OptionsPDFCompression = "Compression"
51920   .OptionsPDFCompressionCaption = "PDF Compression"
51930   .OptionsPDFCompressionColor = "Color Images"
51940   .OptionsPDFCompressionColorComp = "Compress"
51950   .OptionsPDFCompressionColorComp01 = "Automatic"
51960   .OptionsPDFCompressionColorComp02 = "JPEG-Maximum"
51970   .OptionsPDFCompressionColorComp03 = "JPEG-High"
51980   .OptionsPDFCompressionColorComp04 = "JPEG-Medium"
51990   .OptionsPDFCompressionColorComp05 = "JPEG-Low"
52000   .OptionsPDFCompressionColorComp06 = "JPEG-Minimum"
52010   .OptionsPDFCompressionColorComp07 = "ZIP"
52020   .OptionsPDFCompressionColorComp08 = "LZW-Compression"
52030   .OptionsPDFCompressionColorRes = "Resolution"
52040   .OptionsPDFCompressionColorResample = "Resample"
52050   .OptionsPDFCompressionColorResample01 = "Downsample"
52060   .OptionsPDFCompressionColorResample02 = "Average Downsample"
52070   .OptionsPDFCompressionColorResample03 = "Bicubic"
52080   .OptionsPDFCompressionGrey = "Greyscale Images"
52090   .OptionsPDFCompressionGreyComp = "Compress"
52100   .OptionsPDFCompressionGreyComp01 = "Automatic"
52110   .OptionsPDFCompressionGreyComp02 = "JPEG-Maximum"
52120   .OptionsPDFCompressionGreyComp03 = "JPEG-High"
52130   .OptionsPDFCompressionGreyComp04 = "JPEG-Medium"
52140   .OptionsPDFCompressionGreyComp05 = "JPEG-Low"
52150   .OptionsPDFCompressionGreyComp06 = "JPEG-Minimum"
52160   .OptionsPDFCompressionGreyComp07 = "ZIP"
52170   .OptionsPDFCompressionGreyComp08 = "LZW-Compression"
52180   .OptionsPDFCompressionGreyRes = "Resolution"
52190   .OptionsPDFCompressionGreyResample = "Resample"
52200   .OptionsPDFCompressionGreyResample01 = "Downsample"
52210   .OptionsPDFCompressionGreyResample02 = "Average Downsample"
52220   .OptionsPDFCompressionGreyResample03 = "Bicubic"
52230   .OptionsPDFCompressionMono = "Monochrome Images"
52240   .OptionsPDFCompressionMonoComp = "Compress"
52250   .OptionsPDFCompressionMonoComp01 = "CCITT Fax Compression"
52260   .OptionsPDFCompressionMonoComp02 = "ZIP"
52270   .OptionsPDFCompressionMonoComp03 = "Run-Length-Encoding"
52280   .OptionsPDFCompressionMonoComp04 = "LZW-Compression"
52290   .OptionsPDFCompressionMonoRes = "Resolution"
52300   .OptionsPDFCompressionMonoResample = "Resample"
52310   .OptionsPDFCompressionMonoResample01 = "Downsample"
52320   .OptionsPDFCompressionMonoResample02 = "Average Downsample"
52330   .OptionsPDFCompressionMonoResample03 = "Bicubic"
52340   .OptionsPDFCompressionTextComp = "Compress Text Objects"
52350   .OptionsPDFDescription = "Adobe PDF Format"
52360   .OptionsPDFDisallowCopy = "Copy text and images"
52370   .OptionsPDFDisallowModify = "Modify the document"
52380   .OptionsPDFDisallowModifyComments = "Modify comments"
52390   .OptionsPDFDisallowPrint = "Print the document"
52400   .OptionsPDFDisallowUser = "Disallow User to"
52410   .OptionsPDFEncryptionHigh = "High (128 Bit - Adobe Acrobat 5.0 and above)"
52420   .OptionsPDFEncryptionLevel = "Encryption Level"
52430   .OptionsPDFEncryptionLow = "Low (40 Bit - Adobe Acrobat 3.0 and above)"
52440   .OptionsPDFEncryptor = "Encryptor"
52450   .OptionsPDFEnhancedPermissions = "Enhanced Permissions (128 Bit only)"
52460   .OptionsPDFEnterPasswords = "Enter Passwords"
52470   .OptionsPDFFonts = "Fonts"
52480   .OptionsPDFFontsCaption = "Font Options"
52490   .OptionsPDFFontsEmbedAll = "Embed all fonts"
52500   .OptionsPDFFontsSubSetFonts = "Subset fonts when percentage of used characters below:"
52510   .OptionsPDFGeneral = "General"
52520   .OptionsPDFGeneralASCII85 = "Convert binary data to ASCII85"
52530   .OptionsPDFGeneralAutorotate = "Auto-Rotate Pages:"
52540   .OptionsPDFGeneralCaption = "General Options"
52550   .OptionsPDFGeneralCompatibility = "Compatibility:"
52560   .OptionsPDFGeneralCompatibility01 = "Adobe Acrobat 3.0 (PDF 1.2)"
52570   .OptionsPDFGeneralCompatibility02 = "Adobe Acrobat 4.0 (PDF 1.3)"
52580   .OptionsPDFGeneralCompatibility03 = "Adobe Acrobat 5.0 (PDF 1.4)"
52590   .OptionsPDFGeneralCompatibility04 = "Adobe Acrobat 6.0 (PDF 1.5)"
52600   .OptionsPDFGeneralDefaultSettings = "Default settings"
52610   .OptionsPDFGeneralDefaultSettingsDefault = "Default"
52620   .OptionsPDFGeneralDefaultSettingsEbook = "Ebook"
52630   .OptionsPDFGeneralDefaultSettingsPrepress = "Pre-press"
52640   .OptionsPDFGeneralDefaultSettingsPrinter = "Printer"
52650   .OptionsPDFGeneralDefaultSettingsScreen = "Screen"
52660   .OptionsPDFGeneralOverprint = "Overprint:"
52670   .OptionsPDFGeneralOverprint01 = "Non-Zero Overprint"
52680   .OptionsPDFGeneralOverprint02 = "Full Overprint"
52690   .OptionsPDFGeneralResolution = "Resolution:"
52700   .OptionsPDFGeneralRotate01 = "None"
52710   .OptionsPDFGeneralRotate02 = "All"
52720   .OptionsPDFGeneralRotate03 = "Single Page"
52730   .OptionsPDFOptimize = "Fast web view"
52740   .OptionsPDFOptions = "PDF Options"
52750   .OptionsPDFOwnerPass = "Password required to change permissions and passwords"
52760   .OptionsPDFPasswords = "Passwords"
52770   .OptionsPDFRepeatPassword = "Repeat"
52780   .OptionsPDFSecurity = "Security"
52790   .OptionsPDFSecurityCaption = "Security"
52800   .OptionsPDFSetPassword = "Password"
52810   .OptionsPDFSigning = "Signing"
52820   .OptionsPDFSigningCaption = "Signing of PDFs"
52830   .OptionsPDFSigningPfxFile = "Pfx\P12 file"
52840   .OptionsPDFSigningSignatureContact = "Signature contact"
52850   .OptionsPDFSigningSignatureLocation = "Signature location"
52860   .OptionsPDFSigningSignatureMultiSignature = "Multi signature allowed"
52870   .OptionsPDFSigningSignaturePosition = "Signature position"
52880   .OptionsPDFSigningSignaturePositionLeftX = "LeftX"
52890   .OptionsPDFSigningSignaturePositionLeftY = "LeftY"
52900   .OptionsPDFSigningSignaturePositionRightX = "RightX"
52910   .OptionsPDFSigningSignaturePositionRightY = "RightY"
52920   .OptionsPDFSigningSignatureReason = "Signature reason"
52930   .OptionsPDFSigningSignatureVisible = "Signature visible in pdf file"
52940   .OptionsPDFSigningSignPdfFile = "Sign pdf file"
52950   .OptionsPDFSymbol = "PDF"
52960   .OptionsPDFUserPass = "Password required to open document"
52970   .OptionsPDFUseSecurity = "Use Security"
52980   .OptionsPNGColorscount01 = "16777216 colors (24 Bit)"
52990   .OptionsPNGColorscount02 = "256 colors (8 Bit)"
53000   .OptionsPNGColorscount03 = "16 colors (4 Bit)"
53010   .OptionsPNGColorscount04 = "2 colors (2 Bit - Black/White)"
53020   .OptionsPNGColorscount05 = "Greyscale (8 Bit)"
53030   .OptionsPNGDescription = "PNG Format. Please use only for single pages."
53040   .OptionsPNGFiles = "Bitmap PNG-Files"
53050   .OptionsPNGSymbol = "PNG"
53060   .OptionsPrintAfterSaving = "Print after saving"
53070   .OptionsPrintAfterSavingDuplex = "Duplex"
53080   .OptionsPrintAfterSavingDuplexTumbleOff = "Don't use tumble (Default)"
53090   .OptionsPrintAfterSavingDuplexTumbleOn = "Use tumble"
53100   .OptionsPrintAfterSavingNoCancel = "Hide the progress dialog during printing"
53110   .OptionsPrintAfterSavingPrinter = "Printer"
53120   .OptionsPrintAfterSavingQueryUser = "Query user"
53130   .OptionsPrintAfterSavingQueryUserDefaultPrinter = "Select the default Windows printer without any user interaction"
53140   .OptionsPrintAfterSavingQueryUserOff = "Off (Default)"
53150   .OptionsPrintAfterSavingQueryUserPrinterSetupDialog = "Shows the printer setup dialog"
53160   .OptionsPrintAfterSavingQueryUserStandardPrinterDialog = "Show the standard printer dialog"
53170   .OptionsPrintertempDirectoryPrompt = "Select Printer Temp-Directory"
53180   .OptionsPrintTestpage = "Print Test Page"
53190   .OptionsProcesspriority = "Process priority"
53200   .OptionsProcesspriorityHigh = "High"
53210   .OptionsProcesspriorityIdle = "Idle"
53220   .OptionsProcesspriorityNormal = "Normal"
53230   .OptionsProcesspriorityRealtime = "Realtime"
53240   .OptionsProgramActionsDescription = "Define an action before and after saving a file."
53250   .OptionsProgramActionsSymbol = "Actions"
53260   .OptionsProgramAutosaveDescription = "Auto-save mode. Auto-save does not prompt for a filename and file location. It automatically saves all PDF files to a single directory with a predefined filename."
53270   .OptionsProgramAutosaveSymbol = "Auto-save"
53280   .OptionsProgramDirectoriesDescription = "Directories for Ghostscript, temporary files and others."
53290   .OptionsProgramDirectoriesSymbol = "Directories"
53300   .OptionsProgramDocumentDescription = "Document properties"
53310   .OptionsProgramDocumentDescription1 = "Document properties 1"
53320   .OptionsProgramDocumentDescription2 = "Document properties 2"
53330   .OptionsProgramDocumentSymbol = "Document"
53340   .OptionsProgramFont = "Program Font"
53350   .OptionsProgramFontCancelTest = "Cancel Test"
53360   .OptionsProgramFontcharset = "Character Set"
53370   .OptionsProgramFontDescription = "Font for labels, captions and values. For the program menu use the general settings in your Windows OS."
53380   .OptionsProgramFontSize = "Size"
53390   .OptionsProgramFontSymbol = "Program font"
53400   .OptionsProgramFontTest = "Test"
53410   .OptionsProgramFontTestdescription = "Here you can test the font."
53420   .OptionsProgramGeneralDescription = "General Settings"
53430   .OptionsProgramGeneralDescription1 = "General Settings 1"
53440   .OptionsProgramGeneralDescription2 = "General Settings 2"
53450   .OptionsProgramGeneralSymbol = "General settings"
53460   .OptionsProgramGhostscriptDescription = "Ghostscript"
53470   .OptionsProgramGhostscriptSymbol = "Ghostscript"
53480   .OptionsProgramLanguagesDescription = "Define the language and download another languages from the internet."
53490   .OptionsProgramLanguagesSymbol = "Languages"
53500   .OptionsProgramNoProcessingAtStartup = "No processing at startup"
53510   .OptionsProgramOptionsDesign = "Frame color of the options dialog"
53520   .OptionsProgramOptionsDesignGradient = "Red and blue gradient (Default)"
53530   .OptionsProgramOptionsDesignSimple = "Simple red and blue color"
53540   .OptionsProgramPrintDescription = "Print after saving"
53550   .OptionsProgramPrintSymbol = "Print"
53560   .OptionsProgramRunProgramAfterSavingCaption = "Action after saving"
53570   .OptionsProgramRunProgramAfterSavingProgram = "Program/Script"
53580   .OptionsProgramRunProgramAfterSavingProgramParameters = "Program parameters"
53590   .OptionsProgramRunProgramAfterSavingWaitUntilReady = "Wait until the program/script is ready"
53600   .OptionsProgramRunProgramAfterSavingWindowstyle = "Window style"
53610   .OptionsProgramRunProgramAfterSavingWindowstyleHide = "Hide"
53620   .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus = "Maximized/Focus"
53630   .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus = "Minimized/Focus"
53640   .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus = "Minimized/No focus"
53650   .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus = "Normal/Focus"
53660   .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus = "Normal/No focus"
53670   .OptionsProgramRunProgramBeforeSavingCaption = "Action before saving"
53680   .OptionsProgramRunProgramBeforeSavingProgram = "Program/Script"
53690   .OptionsProgramRunProgramBeforeSavingProgramParameters = "Program parameters"
53700   .OptionsProgramRunProgramBeforeSavingWaitUntilReady = "Wait until the program/script is ready"
53710   .OptionsProgramRunProgramBeforeSavingWindowstyle = "Window style"
53720   .OptionsProgramRunProgramBeforeSavingWindowstyleHide = "Hide"
53730   .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus = "Maximized/Focus"
53740   .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus = "Minimized/Focus"
53750   .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus = "Minimized/NoFocus"
53760   .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus = "Normal/Focus"
53770   .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus = "Normal/NoFocus"
53780   .OptionsProgramSaveDescription = "Save"
53790   .OptionsProgramSaveSymbol = "Save"
53800   .OptionsProgramShowAnimation = "Show an animation during the process"
53810   .OptionsProgramSwitchingDefaultprinter = "No confirm message switching PDFCreator temporarily as default printer."
53820   .OptionsPSDColorsCount01 = "4294967296 colors (32 Bit) CMYK"
53830   .OptionsPSDColorscount02 = "16777216 colors (24 Bit)"
53840   .OptionsPSDDescription = "Photoshop Format"
53850   .OptionsPSDescription = "Postscript Format"
53860   .OptionsPSDSymbol = "PSD"
53870   .OptionsPSFiles = "Postscript-Files"
53880   .OptionsPSLanguageLevel = "Language Level:"
53890   .OptionsPSSymbol = "PS"
53900   .OptionsRAWColorsCount01 = "4294967296 colors (32 Bit) CMYK"
53910   .OptionsRAWColorscount02 = "16777216 colors (24 Bit)"
53920   .OptionsRAWColorscount03 = "2 colors (Black/White)"
53930   .OptionsRAWDescription = "Raw Format"
53940   .OptionsRAWSymbol = "Raw"
53950   .OptionsRemoveSpaces = "Remove leading and trailing spaces"
53960   .OptionsReset = "&Reset all settings"
53970   .OptionsSave = "&Save"
53980   .OptionsSaveFilename = "Filename"
53990   .OptionsSaveFilenameAdd = "Add"
54000   .OptionsSaveFilenameChange = "Change"
54010   .OptionsSaveFilenameDelete = "Delete"
54020   .OptionsSaveFilenameSubstitutions = "Filename substitution"
54030   .OptionsSaveFilenameSubstitutionsTitle = "Filename substitution only in <Title>"
54040   .OptionsSaveFilenameTokens = "Add a Filename-Token"
54050   .OptionsSavePasswords = "Save passwords temporarily for this session."
54060   .OptionsSendEmailAfterAutosave = "Send an email after auto-saving"
54070   .OptionsSendMailMethod = "Methode to send an email"
54080   .OptionsSendMailMethodAutomatic = "Automatic"
54090   .OptionsSendMailMethodMapi = "Mapi interface"
54100   .OptionsSendMailMethodSendmailDLL = "Using sendmail.dll"
54110   .OptionsShellIntegration = "Shell integration"
54120   .OptionsShellIntegrationAdd = "Integrate PDFCreator into shell"
54130   .OptionsShellIntegrationCaption = "Create &PDF with PDFCreator"
54140   .OptionsShellIntegrationRemove = "Remove shell integration"
54150   .OptionsStamp = "Stamp"
54160   .OptionsStampFontColor = "Font-color"
54170   .OptionsStampOutlineFontThickness = "Outline font thickness"
54180   .OptionsStampString = "Stampstring"
54190   .OptionsStampUseOutlineFont = "Use outline font"
54200   .OptionsStandardAuthorToken = "Add a Author-Token"
54210   .OptionsStandardSaveFormat = "Standard save format"
54220   .OptionsTestpage = "PDFCreator Testpage"
54230   .OptionsTIFFColorscount01 = "16777216 (24 Bit)"
54240   .OptionsTIFFColorscount02 = "4096 (12 Bit)"
54250   .OptionsTIFFColorscount03 = "2 colors (Black/White) G3 fax encoding with no EOLs"
54260   .OptionsTIFFColorscount04 = "2 colors (Black/White) G3 fax encoding with EOLs"
54270   .OptionsTIFFColorscount05 = "2 colors (Black/White) 2-D G3 fax encoding"
54280   .OptionsTIFFColorscount06 = "2 colors (Black/White) G4 fax encoding"
54290   .OptionsTIFFColorscount07 = "2 colors (Black/White) LZW-compatible"
54300   .OptionsTIFFColorscount08 = "2 colors (Black/White) PackBits"
54310   .OptionsTIFFDescription = "TIFF Format. For multipages use the tiff-format."
54320   .OptionsTIFFSymbol = "TIFF"
54330   .OptionsTreeFormats = "Formats"
54340   .OptionsTreeProgram = "Program"
54350   .OptionsTXTDescription = "Text Format"
54360   .OptionsTXTSymbol = "TXT"
54370   .OptionsUseAutosave = "Use Auto-save"
54380   .OptionsUseAutosaveDirectory = "Use this directory for auto-save"
54390   .OptionsUseCreationDateNow = "Use the current Date/Time for 'Creation Date'"
54400   .OptionsUseCustomPapersize = "Use custom paper size"
54410   .OptionsUseFixPapersize = "Use fixed paper size"
54420   .OptionsUserPass = "User Password"
54430   .OptionsUseStandardauthor = "Use standard author"
54440   .OptionsXCFColorsCount01 = "4294967296 colors (32 Bit) CMYK"
54450   .OptionsXCFColorscount02 = "16777216 colors (24 Bit)"
54460   .OptionsXCFDescription = "Gimp Format"
54470   .OptionsXCFSymbol = "XCF"
54480
54490   .PrintingAuthor = "A&uthor:"
54500   .PrintingBMPFiles = "BMP-Files"
54510   .PrintingCancel = "&Cancel"
54520   .PrintingCollect = "&Wait - Collect"
54530   .PrintingCreationDate = "Creation &Date:"
54540   .PrintingDocumentTitle = "Document &Title:"
54550   .PrintingEMail = "&eMail"
54560   .PrintingEPSFiles = "Encapsulated Postscript-Files"
54570   .PrintingJPEGFiles = "JPEG-Files"
54580   .PrintingKeywords = "&Keywords:"
54590   .PrintingModifyDate = "&Modify Date:"
54600   .PrintingNow = "Now"
54610   .PrintingPCLFiles = "PCL (HP PCL-XL)-Files"
54620   .PrintingPCXFiles = "PCX-Files"
54630   .PrintingPDFAFiles = "PDF/A-1b-Files"
54640   .PrintingPDFFiles = "PDF-Files"
54650   .PrintingPDFXFiles = "PDF/X-Files"
54660   .PrintingPNGFiles = "PNG-Files"
54670   .PrintingPSDFiles = "PSD (Adobe Photoshop)-Files"
54680   .PrintingPSFiles = "Postscript-Files"
54690   .PrintingRAWFiles = "RAW (binary format)-Files"
54700   .PrintingSave = "&Save"
54710   .PrintingStartStandardProgram = "&After saving open the document with the default program."
54720   .PrintingStatus = "Creating file..."
54730   .PrintingSubject = "Su&bject:"
54740   .PrintingTIFFFiles = "TIFF-Files"
54750   .PrintingTXTFiles = "Text-Files"
54760   .PrintingXCFFiles = "XCF (Gimp)-Files"
54770
54780   .SaveOpenAttributes = "Attributes"
54790   .SaveOpenCancel = "Cancel"
54800   .SaveOpenFilename = "Filename"
54810   .SaveOpenOpen = "Open"
54820   .SaveOpenOpenTitle = "Open"
54830   .SaveOpenSave = "Save"
54840   .SaveOpenSaveTitle = "Save as"
54850   .SaveOpenSize = "Size"
54860
54870  End With
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

