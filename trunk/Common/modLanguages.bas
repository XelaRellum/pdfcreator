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
50540   .OptionsNothingToConfigure = Replace$(hLang.Retrieve("NothingToConfigure", .OptionsNothingToConfigure), "/n", vbCrLf)
50550   .OptionsOnePagePerFile = Replace$(hLang.Retrieve("OnePagePerFile", .OptionsOnePagePerFile), "/n", vbCrLf)
50560   .OptionsOwnerPass = Replace$(hLang.Retrieve("OwnerPass", .OptionsOwnerPass), "/n", vbCrLf)
50570   .OptionsPassCancel = Replace$(hLang.Retrieve("PassCancel", .OptionsPassCancel), "/n", vbCrLf)
50580   .OptionsPassOK = Replace$(hLang.Retrieve("PassOK", .OptionsPassOK), "/n", vbCrLf)
50590   .OptionsPCLColorscount01 = Replace$(hLang.Retrieve("PCLColorscount01", .OptionsPCLColorscount01), "/n", vbCrLf)
50600   .OptionsPCLColorscount02 = Replace$(hLang.Retrieve("PCLColorscount02", .OptionsPCLColorscount02), "/n", vbCrLf)
50610   .OptionsPCLDescription = Replace$(hLang.Retrieve("PCLDescription", .OptionsPCLDescription), "/n", vbCrLf)
50620   .OptionsPCLSymbol = Replace$(hLang.Retrieve("PCLSymbol", .OptionsPCLSymbol), "/n", vbCrLf)
50630   .OptionsPCXColorscount01 = Replace$(hLang.Retrieve("PCXColorscount01", .OptionsPCXColorscount01), "/n", vbCrLf)
50640   .OptionsPCXColorscount02 = Replace$(hLang.Retrieve("PCXColorscount02", .OptionsPCXColorscount02), "/n", vbCrLf)
50650   .OptionsPCXColorscount03 = Replace$(hLang.Retrieve("PCXColorscount03", .OptionsPCXColorscount03), "/n", vbCrLf)
50660   .OptionsPCXColorscount04 = Replace$(hLang.Retrieve("PCXColorscount04", .OptionsPCXColorscount04), "/n", vbCrLf)
50670   .OptionsPCXColorscount05 = Replace$(hLang.Retrieve("PCXColorscount05", .OptionsPCXColorscount05), "/n", vbCrLf)
50680   .OptionsPCXColorscount06 = Replace$(hLang.Retrieve("PCXColorscount06", .OptionsPCXColorscount06), "/n", vbCrLf)
50690   .OptionsPCXDescription = Replace$(hLang.Retrieve("PCXDescription", .OptionsPCXDescription), "/n", vbCrLf)
50700   .OptionsPCXSymbol = Replace$(hLang.Retrieve("PCXSymbol", .OptionsPCXSymbol), "/n", vbCrLf)
50710   .OptionsPDFAllowAssembly = Replace$(hLang.Retrieve("PDFAllowAssembly", .OptionsPDFAllowAssembly), "/n", vbCrLf)
50720   .OptionsPDFAllowDegradedPrinting = Replace$(hLang.Retrieve("PDFAllowDegradedPrinting", .OptionsPDFAllowDegradedPrinting), "/n", vbCrLf)
50730   .OptionsPDFAllowFillIn = Replace$(hLang.Retrieve("PDFAllowFillIn", .OptionsPDFAllowFillIn), "/n", vbCrLf)
50740   .OptionsPDFAllowScreenReaders = Replace$(hLang.Retrieve("PDFAllowScreenReaders", .OptionsPDFAllowScreenReaders), "/n", vbCrLf)
50750   .OptionsPDFColors = Replace$(hLang.Retrieve("PDFColors", .OptionsPDFColors), "/n", vbCrLf)
50760   .OptionsPDFColorsCaption = Replace$(hLang.Retrieve("PDFColorsCaption", .OptionsPDFColorsCaption), "/n", vbCrLf)
50770   .OptionsPDFColorsCMYKtoRGB = Replace$(hLang.Retrieve("PDFColorsCMYKtoRGB", .OptionsPDFColorsCMYKtoRGB), "/n", vbCrLf)
50780   .OptionsPDFColorsColorModel01 = Replace$(hLang.Retrieve("PDFColorsColorModel01", .OptionsPDFColorsColorModel01), "/n", vbCrLf)
50790   .OptionsPDFColorsColorModel02 = Replace$(hLang.Retrieve("PDFColorsColorModel02", .OptionsPDFColorsColorModel02), "/n", vbCrLf)
50800   .OptionsPDFColorsColorModel03 = Replace$(hLang.Retrieve("PDFColorsColorModel03", .OptionsPDFColorsColorModel03), "/n", vbCrLf)
50810   .OptionsPDFColorsColorOptions = Replace$(hLang.Retrieve("PDFColorsColorOptions", .OptionsPDFColorsColorOptions), "/n", vbCrLf)
50820   .OptionsPDFColorsPreserveHalftone = Replace$(hLang.Retrieve("PDFColorsPreserveHalftone", .OptionsPDFColorsPreserveHalftone), "/n", vbCrLf)
50830   .OptionsPDFColorsPreserveOverprint = Replace$(hLang.Retrieve("PDFColorsPreserveOverprint", .OptionsPDFColorsPreserveOverprint), "/n", vbCrLf)
50840   .OptionsPDFColorsPreserveTransfer = Replace$(hLang.Retrieve("PDFColorsPreserveTransfer", .OptionsPDFColorsPreserveTransfer), "/n", vbCrLf)
50850   .OptionsPDFCompression = Replace$(hLang.Retrieve("PDFCompression", .OptionsPDFCompression), "/n", vbCrLf)
50860   .OptionsPDFCompressionCaption = Replace$(hLang.Retrieve("PDFCompressionCaption", .OptionsPDFCompressionCaption), "/n", vbCrLf)
50870   .OptionsPDFCompressionColor = Replace$(hLang.Retrieve("PDFCompressionColor", .OptionsPDFCompressionColor), "/n", vbCrLf)
50880   .OptionsPDFCompressionColorComp = Replace$(hLang.Retrieve("PDFCompressionColorComp", .OptionsPDFCompressionColorComp), "/n", vbCrLf)
50890   .OptionsPDFCompressionColorComp01 = Replace$(hLang.Retrieve("PDFCompressionColorComp01", .OptionsPDFCompressionColorComp01), "/n", vbCrLf)
50900   .OptionsPDFCompressionColorComp02 = Replace$(hLang.Retrieve("PDFCompressionColorComp02", .OptionsPDFCompressionColorComp02), "/n", vbCrLf)
50910   .OptionsPDFCompressionColorComp03 = Replace$(hLang.Retrieve("PDFCompressionColorComp03", .OptionsPDFCompressionColorComp03), "/n", vbCrLf)
50920   .OptionsPDFCompressionColorComp04 = Replace$(hLang.Retrieve("PDFCompressionColorComp04", .OptionsPDFCompressionColorComp04), "/n", vbCrLf)
50930   .OptionsPDFCompressionColorComp05 = Replace$(hLang.Retrieve("PDFCompressionColorComp05", .OptionsPDFCompressionColorComp05), "/n", vbCrLf)
50940   .OptionsPDFCompressionColorComp06 = Replace$(hLang.Retrieve("PDFCompressionColorComp06", .OptionsPDFCompressionColorComp06), "/n", vbCrLf)
50950   .OptionsPDFCompressionColorComp07 = Replace$(hLang.Retrieve("PDFCompressionColorComp07", .OptionsPDFCompressionColorComp07), "/n", vbCrLf)
50960   .OptionsPDFCompressionColorComp08 = Replace$(hLang.Retrieve("PDFCompressionColorComp08", .OptionsPDFCompressionColorComp08), "/n", vbCrLf)
50970   .OptionsPDFCompressionColorRes = Replace$(hLang.Retrieve("PDFCompressionColorRes", .OptionsPDFCompressionColorRes), "/n", vbCrLf)
50980   .OptionsPDFCompressionColorResample = Replace$(hLang.Retrieve("PDFCompressionColorResample", .OptionsPDFCompressionColorResample), "/n", vbCrLf)
50990   .OptionsPDFCompressionColorResample01 = Replace$(hLang.Retrieve("PDFCompressionColorResample01", .OptionsPDFCompressionColorResample01), "/n", vbCrLf)
51000   .OptionsPDFCompressionColorResample02 = Replace$(hLang.Retrieve("PDFCompressionColorResample02", .OptionsPDFCompressionColorResample02), "/n", vbCrLf)
51010   .OptionsPDFCompressionColorResample03 = Replace$(hLang.Retrieve("PDFCompressionColorResample03", .OptionsPDFCompressionColorResample03), "/n", vbCrLf)
51020   .OptionsPDFCompressionGrey = Replace$(hLang.Retrieve("PDFCompressionGrey", .OptionsPDFCompressionGrey), "/n", vbCrLf)
51030   .OptionsPDFCompressionGreyComp = Replace$(hLang.Retrieve("PDFCompressionGreyComp", .OptionsPDFCompressionGreyComp), "/n", vbCrLf)
51040   .OptionsPDFCompressionGreyComp01 = Replace$(hLang.Retrieve("PDFCompressionGreyComp01", .OptionsPDFCompressionGreyComp01), "/n", vbCrLf)
51050   .OptionsPDFCompressionGreyComp02 = Replace$(hLang.Retrieve("PDFCompressionGreyComp02", .OptionsPDFCompressionGreyComp02), "/n", vbCrLf)
51060   .OptionsPDFCompressionGreyComp03 = Replace$(hLang.Retrieve("PDFCompressionGreyComp03", .OptionsPDFCompressionGreyComp03), "/n", vbCrLf)
51070   .OptionsPDFCompressionGreyComp04 = Replace$(hLang.Retrieve("PDFCompressionGreyComp04", .OptionsPDFCompressionGreyComp04), "/n", vbCrLf)
51080   .OptionsPDFCompressionGreyComp05 = Replace$(hLang.Retrieve("PDFCompressionGreyComp05", .OptionsPDFCompressionGreyComp05), "/n", vbCrLf)
51090   .OptionsPDFCompressionGreyComp06 = Replace$(hLang.Retrieve("PDFCompressionGreyComp06", .OptionsPDFCompressionGreyComp06), "/n", vbCrLf)
51100   .OptionsPDFCompressionGreyComp07 = Replace$(hLang.Retrieve("PDFCompressionGreyComp07", .OptionsPDFCompressionGreyComp07), "/n", vbCrLf)
51110   .OptionsPDFCompressionGreyComp08 = Replace$(hLang.Retrieve("PDFCompressionGreyComp08", .OptionsPDFCompressionGreyComp08), "/n", vbCrLf)
51120   .OptionsPDFCompressionGreyRes = Replace$(hLang.Retrieve("PDFCompressionGreyRes", .OptionsPDFCompressionGreyRes), "/n", vbCrLf)
51130   .OptionsPDFCompressionGreyResample = Replace$(hLang.Retrieve("PDFCompressionGreyResample", .OptionsPDFCompressionGreyResample), "/n", vbCrLf)
51140   .OptionsPDFCompressionGreyResample01 = Replace$(hLang.Retrieve("PDFCompressionGreyResample01", .OptionsPDFCompressionGreyResample01), "/n", vbCrLf)
51150   .OptionsPDFCompressionGreyResample02 = Replace$(hLang.Retrieve("PDFCompressionGreyResample02", .OptionsPDFCompressionGreyResample02), "/n", vbCrLf)
51160   .OptionsPDFCompressionGreyResample03 = Replace$(hLang.Retrieve("PDFCompressionGreyResample03", .OptionsPDFCompressionGreyResample03), "/n", vbCrLf)
51170   .OptionsPDFCompressionMono = Replace$(hLang.Retrieve("PDFCompressionMono", .OptionsPDFCompressionMono), "/n", vbCrLf)
51180   .OptionsPDFCompressionMonoComp = Replace$(hLang.Retrieve("PDFCompressionMonoComp", .OptionsPDFCompressionMonoComp), "/n", vbCrLf)
51190   .OptionsPDFCompressionMonoComp01 = Replace$(hLang.Retrieve("PDFCompressionMonoComp01", .OptionsPDFCompressionMonoComp01), "/n", vbCrLf)
51200   .OptionsPDFCompressionMonoComp02 = Replace$(hLang.Retrieve("PDFCompressionMonoComp02", .OptionsPDFCompressionMonoComp02), "/n", vbCrLf)
51210   .OptionsPDFCompressionMonoComp03 = Replace$(hLang.Retrieve("PDFCompressionMonoComp03", .OptionsPDFCompressionMonoComp03), "/n", vbCrLf)
51220   .OptionsPDFCompressionMonoComp04 = Replace$(hLang.Retrieve("PDFCompressionMonoComp04", .OptionsPDFCompressionMonoComp04), "/n", vbCrLf)
51230   .OptionsPDFCompressionMonoRes = Replace$(hLang.Retrieve("PDFCompressionMonoRes", .OptionsPDFCompressionMonoRes), "/n", vbCrLf)
51240   .OptionsPDFCompressionMonoResample = Replace$(hLang.Retrieve("PDFCompressionMonoResample", .OptionsPDFCompressionMonoResample), "/n", vbCrLf)
51250   .OptionsPDFCompressionMonoResample01 = Replace$(hLang.Retrieve("PDFCompressionMonoResample01", .OptionsPDFCompressionMonoResample01), "/n", vbCrLf)
51260   .OptionsPDFCompressionMonoResample02 = Replace$(hLang.Retrieve("PDFCompressionMonoResample02", .OptionsPDFCompressionMonoResample02), "/n", vbCrLf)
51270   .OptionsPDFCompressionMonoResample03 = Replace$(hLang.Retrieve("PDFCompressionMonoResample03", .OptionsPDFCompressionMonoResample03), "/n", vbCrLf)
51280   .OptionsPDFCompressionTextComp = Replace$(hLang.Retrieve("PDFCompressionTextComp", .OptionsPDFCompressionTextComp), "/n", vbCrLf)
51290   .OptionsPDFDescription = Replace$(hLang.Retrieve("PDFDescription", .OptionsPDFDescription), "/n", vbCrLf)
51300   .OptionsPDFDisallowCopy = Replace$(hLang.Retrieve("PDFDisallowCopy", .OptionsPDFDisallowCopy), "/n", vbCrLf)
51310   .OptionsPDFDisallowModify = Replace$(hLang.Retrieve("PDFDisallowModify", .OptionsPDFDisallowModify), "/n", vbCrLf)
51320   .OptionsPDFDisallowModifyComments = Replace$(hLang.Retrieve("PDFDisallowModifyComments", .OptionsPDFDisallowModifyComments), "/n", vbCrLf)
51330   .OptionsPDFDisallowPrint = Replace$(hLang.Retrieve("PDFDisallowPrint", .OptionsPDFDisallowPrint), "/n", vbCrLf)
51340   .OptionsPDFDisallowUser = Replace$(hLang.Retrieve("PDFDisallowUser", .OptionsPDFDisallowUser), "/n", vbCrLf)
51350   .OptionsPDFEncryptionHigh = Replace$(hLang.Retrieve("PDFEncryptionHigh", .OptionsPDFEncryptionHigh), "/n", vbCrLf)
51360   .OptionsPDFEncryptionLevel = Replace$(hLang.Retrieve("PDFEncryptionLevel", .OptionsPDFEncryptionLevel), "/n", vbCrLf)
51370   .OptionsPDFEncryptionLow = Replace$(hLang.Retrieve("PDFEncryptionLow", .OptionsPDFEncryptionLow), "/n", vbCrLf)
51380   .OptionsPDFEncryptor = Replace$(hLang.Retrieve("PDFEncryptor", .OptionsPDFEncryptor), "/n", vbCrLf)
51390   .OptionsPDFEnhancedPermissions = Replace$(hLang.Retrieve("PDFEnhancedPermissions", .OptionsPDFEnhancedPermissions), "/n", vbCrLf)
51400   .OptionsPDFEnterPasswords = Replace$(hLang.Retrieve("PDFEnterPasswords", .OptionsPDFEnterPasswords), "/n", vbCrLf)
51410   .OptionsPDFFonts = Replace$(hLang.Retrieve("PDFFonts", .OptionsPDFFonts), "/n", vbCrLf)
51420   .OptionsPDFFontsCaption = Replace$(hLang.Retrieve("PDFFontsCaption", .OptionsPDFFontsCaption), "/n", vbCrLf)
51430   .OptionsPDFFontsEmbedAll = Replace$(hLang.Retrieve("PDFFontsEmbedAll", .OptionsPDFFontsEmbedAll), "/n", vbCrLf)
51440   .OptionsPDFFontsSubSetFonts = Replace$(hLang.Retrieve("PDFFontsSubSetFonts", .OptionsPDFFontsSubSetFonts), "/n", vbCrLf)
51450   .OptionsPDFGeneral = Replace$(hLang.Retrieve("PDFGeneral", .OptionsPDFGeneral), "/n", vbCrLf)
51460   .OptionsPDFGeneralASCII85 = Replace$(hLang.Retrieve("PDFGeneralASCII85", .OptionsPDFGeneralASCII85), "/n", vbCrLf)
51470   .OptionsPDFGeneralAutorotate = Replace$(hLang.Retrieve("PDFGeneralAutorotate", .OptionsPDFGeneralAutorotate), "/n", vbCrLf)
51480   .OptionsPDFGeneralCaption = Replace$(hLang.Retrieve("PDFGeneralCaption", .OptionsPDFGeneralCaption), "/n", vbCrLf)
51490   .OptionsPDFGeneralCompatibility = Replace$(hLang.Retrieve("PDFGeneralCompatibility", .OptionsPDFGeneralCompatibility), "/n", vbCrLf)
51500   .OptionsPDFGeneralCompatibility01 = Replace$(hLang.Retrieve("PDFGeneralCompatibility01", .OptionsPDFGeneralCompatibility01), "/n", vbCrLf)
51510   .OptionsPDFGeneralCompatibility02 = Replace$(hLang.Retrieve("PDFGeneralCompatibility02", .OptionsPDFGeneralCompatibility02), "/n", vbCrLf)
51520   .OptionsPDFGeneralCompatibility03 = Replace$(hLang.Retrieve("PDFGeneralCompatibility03", .OptionsPDFGeneralCompatibility03), "/n", vbCrLf)
51530   .OptionsPDFGeneralCompatibility04 = Replace$(hLang.Retrieve("PDFGeneralCompatibility04", .OptionsPDFGeneralCompatibility04), "/n", vbCrLf)
51540   .OptionsPDFGeneralDefaultSettings = Replace$(hLang.Retrieve("PDFGeneralDefaultSettings", .OptionsPDFGeneralDefaultSettings), "/n", vbCrLf)
51550   .OptionsPDFGeneralDefaultSettingsDefault = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsDefault", .OptionsPDFGeneralDefaultSettingsDefault), "/n", vbCrLf)
51560   .OptionsPDFGeneralDefaultSettingsEbook = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsEbook", .OptionsPDFGeneralDefaultSettingsEbook), "/n", vbCrLf)
51570   .OptionsPDFGeneralDefaultSettingsPrepress = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsPrepress", .OptionsPDFGeneralDefaultSettingsPrepress), "/n", vbCrLf)
51580   .OptionsPDFGeneralDefaultSettingsPrinter = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsPrinter", .OptionsPDFGeneralDefaultSettingsPrinter), "/n", vbCrLf)
51590   .OptionsPDFGeneralDefaultSettingsScreen = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsScreen", .OptionsPDFGeneralDefaultSettingsScreen), "/n", vbCrLf)
51600   .OptionsPDFGeneralOverprint = Replace$(hLang.Retrieve("PDFGeneralOverprint", .OptionsPDFGeneralOverprint), "/n", vbCrLf)
51610   .OptionsPDFGeneralOverprint01 = Replace$(hLang.Retrieve("PDFGeneralOverprint01", .OptionsPDFGeneralOverprint01), "/n", vbCrLf)
51620   .OptionsPDFGeneralOverprint02 = Replace$(hLang.Retrieve("PDFGeneralOverprint02", .OptionsPDFGeneralOverprint02), "/n", vbCrLf)
51630   .OptionsPDFGeneralResolution = Replace$(hLang.Retrieve("PDFGeneralResolution", .OptionsPDFGeneralResolution), "/n", vbCrLf)
51640   .OptionsPDFGeneralRotate01 = Replace$(hLang.Retrieve("PDFGeneralRotate01", .OptionsPDFGeneralRotate01), "/n", vbCrLf)
51650   .OptionsPDFGeneralRotate02 = Replace$(hLang.Retrieve("PDFGeneralRotate02", .OptionsPDFGeneralRotate02), "/n", vbCrLf)
51660   .OptionsPDFGeneralRotate03 = Replace$(hLang.Retrieve("PDFGeneralRotate03", .OptionsPDFGeneralRotate03), "/n", vbCrLf)
51670   .OptionsPDFOptimize = Replace$(hLang.Retrieve("PDFOptimize", .OptionsPDFOptimize), "/n", vbCrLf)
51680   .OptionsPDFOptions = Replace$(hLang.Retrieve("PDFOptions", .OptionsPDFOptions), "/n", vbCrLf)
51690   .OptionsPDFOwnerPass = Replace$(hLang.Retrieve("PDFOwnerPass", .OptionsPDFOwnerPass), "/n", vbCrLf)
51700   .OptionsPDFPasswords = Replace$(hLang.Retrieve("PDFPasswords", .OptionsPDFPasswords), "/n", vbCrLf)
51710   .OptionsPDFRepeatPassword = Replace$(hLang.Retrieve("PDFRepeatPassword", .OptionsPDFRepeatPassword), "/n", vbCrLf)
51720   .OptionsPDFSecurity = Replace$(hLang.Retrieve("PDFSecurity", .OptionsPDFSecurity), "/n", vbCrLf)
51730   .OptionsPDFSecurityCaption = Replace$(hLang.Retrieve("PDFSecurityCaption", .OptionsPDFSecurityCaption), "/n", vbCrLf)
51740   .OptionsPDFSetPassword = Replace$(hLang.Retrieve("PDFSetPassword", .OptionsPDFSetPassword), "/n", vbCrLf)
51750   .OptionsPDFSigning = Replace$(hLang.Retrieve("PDFSigning", .OptionsPDFSigning), "/n", vbCrLf)
51760   .OptionsPDFSigningCaption = Replace$(hLang.Retrieve("PDFSigningCaption", .OptionsPDFSigningCaption), "/n", vbCrLf)
51770   .OptionsPDFSigningPfxFile = Replace$(hLang.Retrieve("PDFSigningPfxFile", .OptionsPDFSigningPfxFile), "/n", vbCrLf)
51780   .OptionsPDFSigningSignatureContact = Replace$(hLang.Retrieve("PDFSigningSignatureContact", .OptionsPDFSigningSignatureContact), "/n", vbCrLf)
51790   .OptionsPDFSigningSignatureLocation = Replace$(hLang.Retrieve("PDFSigningSignatureLocation", .OptionsPDFSigningSignatureLocation), "/n", vbCrLf)
51800   .OptionsPDFSigningSignatureMultiSignature = Replace$(hLang.Retrieve("PDFSigningSignatureMultiSignature", .OptionsPDFSigningSignatureMultiSignature), "/n", vbCrLf)
51810   .OptionsPDFSigningSignaturePosition = Replace$(hLang.Retrieve("PDFSigningSignaturePosition", .OptionsPDFSigningSignaturePosition), "/n", vbCrLf)
51820   .OptionsPDFSigningSignaturePositionLeftX = Replace$(hLang.Retrieve("PDFSigningSignaturePositionLeftX", .OptionsPDFSigningSignaturePositionLeftX), "/n", vbCrLf)
51830   .OptionsPDFSigningSignaturePositionLeftY = Replace$(hLang.Retrieve("PDFSigningSignaturePositionLeftY", .OptionsPDFSigningSignaturePositionLeftY), "/n", vbCrLf)
51840   .OptionsPDFSigningSignaturePositionRightX = Replace$(hLang.Retrieve("PDFSigningSignaturePositionRightX", .OptionsPDFSigningSignaturePositionRightX), "/n", vbCrLf)
51850   .OptionsPDFSigningSignaturePositionRightY = Replace$(hLang.Retrieve("PDFSigningSignaturePositionRightY", .OptionsPDFSigningSignaturePositionRightY), "/n", vbCrLf)
51860   .OptionsPDFSigningSignatureReason = Replace$(hLang.Retrieve("PDFSigningSignatureReason", .OptionsPDFSigningSignatureReason), "/n", vbCrLf)
51870   .OptionsPDFSigningSignatureVisible = Replace$(hLang.Retrieve("PDFSigningSignatureVisible", .OptionsPDFSigningSignatureVisible), "/n", vbCrLf)
51880   .OptionsPDFSigningSignPdfFile = Replace$(hLang.Retrieve("PDFSigningSignPdfFile", .OptionsPDFSigningSignPdfFile), "/n", vbCrLf)
51890   .OptionsPDFSymbol = Replace$(hLang.Retrieve("PDFSymbol", .OptionsPDFSymbol), "/n", vbCrLf)
51900   .OptionsPDFUserPass = Replace$(hLang.Retrieve("PDFUserPass", .OptionsPDFUserPass), "/n", vbCrLf)
51910   .OptionsPDFUseSecurity = Replace$(hLang.Retrieve("PDFUseSecurity", .OptionsPDFUseSecurity), "/n", vbCrLf)
51920   .OptionsPNGColorscount01 = Replace$(hLang.Retrieve("PNGColorscount01", .OptionsPNGColorscount01), "/n", vbCrLf)
51930   .OptionsPNGColorscount02 = Replace$(hLang.Retrieve("PNGColorscount02", .OptionsPNGColorscount02), "/n", vbCrLf)
51940   .OptionsPNGColorscount03 = Replace$(hLang.Retrieve("PNGColorscount03", .OptionsPNGColorscount03), "/n", vbCrLf)
51950   .OptionsPNGColorscount04 = Replace$(hLang.Retrieve("PNGColorscount04", .OptionsPNGColorscount04), "/n", vbCrLf)
51960   .OptionsPNGColorscount05 = Replace$(hLang.Retrieve("PNGColorscount05", .OptionsPNGColorscount05), "/n", vbCrLf)
51970   .OptionsPNGDescription = Replace$(hLang.Retrieve("PNGDescription", .OptionsPNGDescription), "/n", vbCrLf)
51980   .OptionsPNGFiles = Replace$(hLang.Retrieve("PNGFiles", .OptionsPNGFiles), "/n", vbCrLf)
51990   .OptionsPNGSymbol = Replace$(hLang.Retrieve("PNGSymbol", .OptionsPNGSymbol), "/n", vbCrLf)
52000   .OptionsPrintAfterSaving = Replace$(hLang.Retrieve("PrintAfterSaving", .OptionsPrintAfterSaving), "/n", vbCrLf)
52010   .OptionsPrintAfterSavingDuplex = Replace$(hLang.Retrieve("PrintAfterSavingDuplex", .OptionsPrintAfterSavingDuplex), "/n", vbCrLf)
52020   .OptionsPrintAfterSavingDuplexTumbleOff = Replace$(hLang.Retrieve("PrintAfterSavingDuplexTumbleOff", .OptionsPrintAfterSavingDuplexTumbleOff), "/n", vbCrLf)
52030   .OptionsPrintAfterSavingDuplexTumbleOn = Replace$(hLang.Retrieve("PrintAfterSavingDuplexTumbleOn", .OptionsPrintAfterSavingDuplexTumbleOn), "/n", vbCrLf)
52040   .OptionsPrintAfterSavingNoCancel = Replace$(hLang.Retrieve("PrintAfterSavingNoCancel", .OptionsPrintAfterSavingNoCancel), "/n", vbCrLf)
52050   .OptionsPrintAfterSavingPrinter = Replace$(hLang.Retrieve("PrintAfterSavingPrinter", .OptionsPrintAfterSavingPrinter), "/n", vbCrLf)
52060   .OptionsPrintAfterSavingQueryUser = Replace$(hLang.Retrieve("PrintAfterSavingQueryUser", .OptionsPrintAfterSavingQueryUser), "/n", vbCrLf)
52070   .OptionsPrintAfterSavingQueryUserDefaultPrinter = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserDefaultPrinter", .OptionsPrintAfterSavingQueryUserDefaultPrinter), "/n", vbCrLf)
52080   .OptionsPrintAfterSavingQueryUserOff = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserOff", .OptionsPrintAfterSavingQueryUserOff), "/n", vbCrLf)
52090   .OptionsPrintAfterSavingQueryUserPrinterSetupDialog = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserPrinterSetupDialog", .OptionsPrintAfterSavingQueryUserPrinterSetupDialog), "/n", vbCrLf)
52100   .OptionsPrintAfterSavingQueryUserStandardPrinterDialog = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserStandardPrinterDialog", .OptionsPrintAfterSavingQueryUserStandardPrinterDialog), "/n", vbCrLf)
52110   .OptionsPrintertempDirectoryPrompt = Replace$(hLang.Retrieve("PrintertempDirectoryPrompt", .OptionsPrintertempDirectoryPrompt), "/n", vbCrLf)
52120   .OptionsPrintTestpage = Replace$(hLang.Retrieve("PrintTestpage", .OptionsPrintTestpage), "/n", vbCrLf)
52130   .OptionsProcesspriority = Replace$(hLang.Retrieve("Processpriority", .OptionsProcesspriority), "/n", vbCrLf)
52140   .OptionsProcesspriorityHigh = Replace$(hLang.Retrieve("ProcesspriorityHigh", .OptionsProcesspriorityHigh), "/n", vbCrLf)
52150   .OptionsProcesspriorityIdle = Replace$(hLang.Retrieve("ProcesspriorityIdle", .OptionsProcesspriorityIdle), "/n", vbCrLf)
52160   .OptionsProcesspriorityNormal = Replace$(hLang.Retrieve("ProcesspriorityNormal", .OptionsProcesspriorityNormal), "/n", vbCrLf)
52170   .OptionsProcesspriorityRealtime = Replace$(hLang.Retrieve("ProcesspriorityRealtime", .OptionsProcesspriorityRealtime), "/n", vbCrLf)
52180   .OptionsProgramActionsDescription = Replace$(hLang.Retrieve("ProgramActionsDescription", .OptionsProgramActionsDescription), "/n", vbCrLf)
52190   .OptionsProgramActionsSymbol = Replace$(hLang.Retrieve("ProgramActionsSymbol", .OptionsProgramActionsSymbol), "/n", vbCrLf)
52200   .OptionsProgramAutosaveDescription = Replace$(hLang.Retrieve("ProgramAutosaveDescription", .OptionsProgramAutosaveDescription), "/n", vbCrLf)
52210   .OptionsProgramAutosaveSymbol = Replace$(hLang.Retrieve("ProgramAutosaveSymbol", .OptionsProgramAutosaveSymbol), "/n", vbCrLf)
52220   .OptionsProgramDirectoriesDescription = Replace$(hLang.Retrieve("ProgramDirectoriesDescription", .OptionsProgramDirectoriesDescription), "/n", vbCrLf)
52230   .OptionsProgramDirectoriesSymbol = Replace$(hLang.Retrieve("ProgramDirectoriesSymbol", .OptionsProgramDirectoriesSymbol), "/n", vbCrLf)
52240   .OptionsProgramDocumentDescription = Replace$(hLang.Retrieve("ProgramDocumentDescription", .OptionsProgramDocumentDescription), "/n", vbCrLf)
52250   .OptionsProgramDocumentDescription1 = Replace$(hLang.Retrieve("ProgramDocumentDescription1", .OptionsProgramDocumentDescription1), "/n", vbCrLf)
52260   .OptionsProgramDocumentDescription2 = Replace$(hLang.Retrieve("ProgramDocumentDescription2", .OptionsProgramDocumentDescription2), "/n", vbCrLf)
52270   .OptionsProgramDocumentSymbol = Replace$(hLang.Retrieve("ProgramDocumentSymbol", .OptionsProgramDocumentSymbol), "/n", vbCrLf)
52280   .OptionsProgramFont = Replace$(hLang.Retrieve("ProgramFont", .OptionsProgramFont), "/n", vbCrLf)
52290   .OptionsProgramFontCancelTest = Replace$(hLang.Retrieve("ProgramFontCancelTest", .OptionsProgramFontCancelTest), "/n", vbCrLf)
52300   .OptionsProgramFontcharset = Replace$(hLang.Retrieve("ProgramFontcharset", .OptionsProgramFontcharset), "/n", vbCrLf)
52310   .OptionsProgramFontDescription = Replace$(hLang.Retrieve("ProgramFontDescription", .OptionsProgramFontDescription), "/n", vbCrLf)
52320   .OptionsProgramFontSize = Replace$(hLang.Retrieve("ProgramFontSize", .OptionsProgramFontSize), "/n", vbCrLf)
52330   .OptionsProgramFontSymbol = Replace$(hLang.Retrieve("ProgramFontSymbol", .OptionsProgramFontSymbol), "/n", vbCrLf)
52340   .OptionsProgramFontTest = Replace$(hLang.Retrieve("ProgramFontTest", .OptionsProgramFontTest), "/n", vbCrLf)
52350   .OptionsProgramFontTestdescription = Replace$(hLang.Retrieve("ProgramFontTestdescription", .OptionsProgramFontTestdescription), "/n", vbCrLf)
52360   .OptionsProgramGeneralDescription = Replace$(hLang.Retrieve("ProgramGeneralDescription", .OptionsProgramGeneralDescription), "/n", vbCrLf)
52370   .OptionsProgramGeneralDescription1 = Replace$(hLang.Retrieve("ProgramGeneralDescription1", .OptionsProgramGeneralDescription1), "/n", vbCrLf)
52380   .OptionsProgramGeneralDescription2 = Replace$(hLang.Retrieve("ProgramGeneralDescription2", .OptionsProgramGeneralDescription2), "/n", vbCrLf)
52390   .OptionsProgramGeneralSymbol = Replace$(hLang.Retrieve("ProgramGeneralSymbol", .OptionsProgramGeneralSymbol), "/n", vbCrLf)
52400   .OptionsProgramGhostscriptDescription = Replace$(hLang.Retrieve("ProgramGhostscriptDescription", .OptionsProgramGhostscriptDescription), "/n", vbCrLf)
52410   .OptionsProgramGhostscriptSymbol = Replace$(hLang.Retrieve("ProgramGhostscriptSymbol", .OptionsProgramGhostscriptSymbol), "/n", vbCrLf)
52420   .OptionsProgramLanguagesDescription = Replace$(hLang.Retrieve("ProgramLanguagesDescription", .OptionsProgramLanguagesDescription), "/n", vbCrLf)
52430   .OptionsProgramLanguagesSymbol = Replace$(hLang.Retrieve("ProgramLanguagesSymbol", .OptionsProgramLanguagesSymbol), "/n", vbCrLf)
52440   .OptionsProgramNoProcessingAtStartup = Replace$(hLang.Retrieve("ProgramNoProcessingAtStartup", .OptionsProgramNoProcessingAtStartup), "/n", vbCrLf)
52450   .OptionsProgramOptionsDesign = Replace$(hLang.Retrieve("ProgramOptionsDesign", .OptionsProgramOptionsDesign), "/n", vbCrLf)
52460   .OptionsProgramOptionsDesignGradient = Replace$(hLang.Retrieve("ProgramOptionsDesignGradient", .OptionsProgramOptionsDesignGradient), "/n", vbCrLf)
52470   .OptionsProgramOptionsDesignSimple = Replace$(hLang.Retrieve("ProgramOptionsDesignSimple", .OptionsProgramOptionsDesignSimple), "/n", vbCrLf)
52480   .OptionsProgramPrintDescription = Replace$(hLang.Retrieve("ProgramPrintDescription", .OptionsProgramPrintDescription), "/n", vbCrLf)
52490   .OptionsProgramPrintSymbol = Replace$(hLang.Retrieve("ProgramPrintSymbol", .OptionsProgramPrintSymbol), "/n", vbCrLf)
52500   .OptionsProgramRunProgramAfterSavingCaption = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingCaption", .OptionsProgramRunProgramAfterSavingCaption), "/n", vbCrLf)
52510   .OptionsProgramRunProgramAfterSavingProgram = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingProgram", .OptionsProgramRunProgramAfterSavingProgram), "/n", vbCrLf)
52520   .OptionsProgramRunProgramAfterSavingProgramParameters = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingProgramParameters", .OptionsProgramRunProgramAfterSavingProgramParameters), "/n", vbCrLf)
52530   .OptionsProgramRunProgramAfterSavingWaitUntilReady = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWaitUntilReady", .OptionsProgramRunProgramAfterSavingWaitUntilReady), "/n", vbCrLf)
52540   .OptionsProgramRunProgramAfterSavingWindowstyle = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyle", .OptionsProgramRunProgramAfterSavingWindowstyle), "/n", vbCrLf)
52550   .OptionsProgramRunProgramAfterSavingWindowstyleHide = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleHide", .OptionsProgramRunProgramAfterSavingWindowstyleHide), "/n", vbCrLf)
52560   .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleMaximizedFocus", .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus), "/n", vbCrLf)
52570   .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleMinimizedFocus", .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus), "/n", vbCrLf)
52580   .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus", .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus), "/n", vbCrLf)
52590   .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleNormalFocus", .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus), "/n", vbCrLf)
52600   .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleNormalNoFocus", .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus), "/n", vbCrLf)
52610   .OptionsProgramRunProgramBeforeSavingCaption = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingCaption", .OptionsProgramRunProgramBeforeSavingCaption), "/n", vbCrLf)
52620   .OptionsProgramRunProgramBeforeSavingProgram = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingProgram", .OptionsProgramRunProgramBeforeSavingProgram), "/n", vbCrLf)
52630   .OptionsProgramRunProgramBeforeSavingProgramParameters = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingProgramParameters", .OptionsProgramRunProgramBeforeSavingProgramParameters), "/n", vbCrLf)
52640   .OptionsProgramRunProgramBeforeSavingWaitUntilReady = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWaitUntilReady", .OptionsProgramRunProgramBeforeSavingWaitUntilReady), "/n", vbCrLf)
52650   .OptionsProgramRunProgramBeforeSavingWindowstyle = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyle", .OptionsProgramRunProgramBeforeSavingWindowstyle), "/n", vbCrLf)
52660   .OptionsProgramRunProgramBeforeSavingWindowstyleHide = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleHide", .OptionsProgramRunProgramBeforeSavingWindowstyleHide), "/n", vbCrLf)
52670   .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleMaximizedFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus), "/n", vbCrLf)
52680   .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleMinimizedFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus), "/n", vbCrLf)
52690   .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus), "/n", vbCrLf)
52700   .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleNormalFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus), "/n", vbCrLf)
52710   .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleNormalNoFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus), "/n", vbCrLf)
52720   .OptionsProgramSaveDescription = Replace$(hLang.Retrieve("ProgramSaveDescription", .OptionsProgramSaveDescription), "/n", vbCrLf)
52730   .OptionsProgramSaveSymbol = Replace$(hLang.Retrieve("ProgramSaveSymbol", .OptionsProgramSaveSymbol), "/n", vbCrLf)
52740   .OptionsProgramShowAnimation = Replace$(hLang.Retrieve("ProgramShowAnimation", .OptionsProgramShowAnimation), "/n", vbCrLf)
52750   .OptionsProgramSwitchingDefaultprinter = Replace$(hLang.Retrieve("ProgramSwitchingDefaultprinter", .OptionsProgramSwitchingDefaultprinter), "/n", vbCrLf)
52760   .OptionsPSDColorsCount01 = Replace$(hLang.Retrieve("PSDColorsCount01", .OptionsPSDColorsCount01), "/n", vbCrLf)
52770   .OptionsPSDColorscount02 = Replace$(hLang.Retrieve("PSDColorscount02", .OptionsPSDColorscount02), "/n", vbCrLf)
52780   .OptionsPSDDescription = Replace$(hLang.Retrieve("PSDDescription", .OptionsPSDDescription), "/n", vbCrLf)
52790   .OptionsPSDescription = Replace$(hLang.Retrieve("PSDescription", .OptionsPSDescription), "/n", vbCrLf)
52800   .OptionsPSDSymbol = Replace$(hLang.Retrieve("PSDSymbol", .OptionsPSDSymbol), "/n", vbCrLf)
52810   .OptionsPSFiles = Replace$(hLang.Retrieve("PSFiles", .OptionsPSFiles), "/n", vbCrLf)
52820   .OptionsPSLanguageLevel = Replace$(hLang.Retrieve("PSLanguageLevel", .OptionsPSLanguageLevel), "/n", vbCrLf)
52830   .OptionsPSSymbol = Replace$(hLang.Retrieve("PSSymbol", .OptionsPSSymbol), "/n", vbCrLf)
52840   .OptionsRAWColorsCount01 = Replace$(hLang.Retrieve("RAWColorsCount01", .OptionsRAWColorsCount01), "/n", vbCrLf)
52850   .OptionsRAWColorscount02 = Replace$(hLang.Retrieve("RAWColorscount02", .OptionsRAWColorscount02), "/n", vbCrLf)
52860   .OptionsRAWColorscount03 = Replace$(hLang.Retrieve("RAWColorscount03", .OptionsRAWColorscount03), "/n", vbCrLf)
52870   .OptionsRAWDescription = Replace$(hLang.Retrieve("RAWDescription", .OptionsRAWDescription), "/n", vbCrLf)
52880   .OptionsRAWSymbol = Replace$(hLang.Retrieve("RAWSymbol", .OptionsRAWSymbol), "/n", vbCrLf)
52890   .OptionsRemoveSpaces = Replace$(hLang.Retrieve("RemoveSpaces", .OptionsRemoveSpaces), "/n", vbCrLf)
52900   .OptionsReset = Replace$(hLang.Retrieve("Reset", .OptionsReset), "/n", vbCrLf)
52910   .OptionsSave = Replace$(hLang.Retrieve("Save", .OptionsSave), "/n", vbCrLf)
52920   .OptionsSaveFilename = Replace$(hLang.Retrieve("SaveFilename", .OptionsSaveFilename), "/n", vbCrLf)
52930   .OptionsSaveFilenameAdd = Replace$(hLang.Retrieve("SaveFilenameAdd", .OptionsSaveFilenameAdd), "/n", vbCrLf)
52940   .OptionsSaveFilenameChange = Replace$(hLang.Retrieve("SaveFilenameChange", .OptionsSaveFilenameChange), "/n", vbCrLf)
52950   .OptionsSaveFilenameDelete = Replace$(hLang.Retrieve("SaveFilenameDelete", .OptionsSaveFilenameDelete), "/n", vbCrLf)
52960   .OptionsSaveFilenameSubstitutions = Replace$(hLang.Retrieve("SaveFilenameSubstitutions", .OptionsSaveFilenameSubstitutions), "/n", vbCrLf)
52970   .OptionsSaveFilenameSubstitutionsTitle = Replace$(hLang.Retrieve("SaveFilenameSubstitutionsTitle", .OptionsSaveFilenameSubstitutionsTitle), "/n", vbCrLf)
52980   .OptionsSaveFilenameTokens = Replace$(hLang.Retrieve("SaveFilenameTokens", .OptionsSaveFilenameTokens), "/n", vbCrLf)
52990   .OptionsSavePasswords = Replace$(hLang.Retrieve("SavePasswords", .OptionsSavePasswords), "/n", vbCrLf)
53000   .OptionsSendEmailAfterAutosave = Replace$(hLang.Retrieve("SendEmailAfterAutosave", .OptionsSendEmailAfterAutosave), "/n", vbCrLf)
53010   .OptionsSendMailMethod = Replace$(hLang.Retrieve("SendMailMethod", .OptionsSendMailMethod), "/n", vbCrLf)
53020   .OptionsSendMailMethodAutomatic = Replace$(hLang.Retrieve("SendMailMethodAutomatic", .OptionsSendMailMethodAutomatic), "/n", vbCrLf)
53030   .OptionsSendMailMethodMapi = Replace$(hLang.Retrieve("SendMailMethodMapi", .OptionsSendMailMethodMapi), "/n", vbCrLf)
53040   .OptionsSendMailMethodSendmailDLL = Replace$(hLang.Retrieve("SendMailMethodSendmailDLL", .OptionsSendMailMethodSendmailDLL), "/n", vbCrLf)
53050   .OptionsShellIntegration = Replace$(hLang.Retrieve("ShellIntegration", .OptionsShellIntegration), "/n", vbCrLf)
53060   .OptionsShellIntegrationAdd = Replace$(hLang.Retrieve("ShellIntegrationAdd", .OptionsShellIntegrationAdd), "/n", vbCrLf)
53070   .OptionsShellIntegrationCaption = Replace$(hLang.Retrieve("ShellIntegrationCaption", .OptionsShellIntegrationCaption), "/n", vbCrLf)
53080   .OptionsShellIntegrationRemove = Replace$(hLang.Retrieve("ShellIntegrationRemove", .OptionsShellIntegrationRemove), "/n", vbCrLf)
53090   .OptionsStamp = Replace$(hLang.Retrieve("Stamp", .OptionsStamp), "/n", vbCrLf)
53100   .OptionsStampFontColor = Replace$(hLang.Retrieve("StampFontColor", .OptionsStampFontColor), "/n", vbCrLf)
53110   .OptionsStampOutlineFontThickness = Replace$(hLang.Retrieve("StampOutlineFontThickness", .OptionsStampOutlineFontThickness), "/n", vbCrLf)
53120   .OptionsStampString = Replace$(hLang.Retrieve("StampString", .OptionsStampString), "/n", vbCrLf)
53130   .OptionsStampUseOutlineFont = Replace$(hLang.Retrieve("StampUseOutlineFont", .OptionsStampUseOutlineFont), "/n", vbCrLf)
53140   .OptionsStandardAuthorToken = Replace$(hLang.Retrieve("StandardAuthorToken", .OptionsStandardAuthorToken), "/n", vbCrLf)
53150   .OptionsStandardSaveFormat = Replace$(hLang.Retrieve("StandardSaveFormat", .OptionsStandardSaveFormat), "/n", vbCrLf)
53160   .OptionsTestpage = Replace$(hLang.Retrieve("Testpage", .OptionsTestpage), "/n", vbCrLf)
53170   .OptionsTIFFColorscount01 = Replace$(hLang.Retrieve("TIFFColorscount01", .OptionsTIFFColorscount01), "/n", vbCrLf)
53180   .OptionsTIFFColorscount02 = Replace$(hLang.Retrieve("TIFFColorscount02", .OptionsTIFFColorscount02), "/n", vbCrLf)
53190   .OptionsTIFFColorscount03 = Replace$(hLang.Retrieve("TIFFColorscount03", .OptionsTIFFColorscount03), "/n", vbCrLf)
53200   .OptionsTIFFColorscount04 = Replace$(hLang.Retrieve("TIFFColorscount04", .OptionsTIFFColorscount04), "/n", vbCrLf)
53210   .OptionsTIFFColorscount05 = Replace$(hLang.Retrieve("TIFFColorscount05", .OptionsTIFFColorscount05), "/n", vbCrLf)
53220   .OptionsTIFFColorscount06 = Replace$(hLang.Retrieve("TIFFColorscount06", .OptionsTIFFColorscount06), "/n", vbCrLf)
53230   .OptionsTIFFColorscount07 = Replace$(hLang.Retrieve("TIFFColorscount07", .OptionsTIFFColorscount07), "/n", vbCrLf)
53240   .OptionsTIFFColorscount08 = Replace$(hLang.Retrieve("TIFFColorscount08", .OptionsTIFFColorscount08), "/n", vbCrLf)
53250   .OptionsTIFFDescription = Replace$(hLang.Retrieve("TIFFDescription", .OptionsTIFFDescription), "/n", vbCrLf)
53260   .OptionsTIFFSymbol = Replace$(hLang.Retrieve("TIFFSymbol", .OptionsTIFFSymbol), "/n", vbCrLf)
53270   .OptionsTreeFormats = Replace$(hLang.Retrieve("TreeFormats", .OptionsTreeFormats), "/n", vbCrLf)
53280   .OptionsTreeProgram = Replace$(hLang.Retrieve("TreeProgram", .OptionsTreeProgram), "/n", vbCrLf)
53290   .OptionsTXTDescription = Replace$(hLang.Retrieve("TXTDescription", .OptionsTXTDescription), "/n", vbCrLf)
53300   .OptionsTXTSymbol = Replace$(hLang.Retrieve("TXTSymbol", .OptionsTXTSymbol), "/n", vbCrLf)
53310   .OptionsUseAutosave = Replace$(hLang.Retrieve("UseAutosave", .OptionsUseAutosave), "/n", vbCrLf)
53320   .OptionsUseAutosaveDirectory = Replace$(hLang.Retrieve("UseAutosaveDirectory", .OptionsUseAutosaveDirectory), "/n", vbCrLf)
53330   .OptionsUseCreationDateNow = Replace$(hLang.Retrieve("UseCreationDateNow", .OptionsUseCreationDateNow), "/n", vbCrLf)
53340   .OptionsUseCustomPapersize = Replace$(hLang.Retrieve("UseCustomPapersize", .OptionsUseCustomPapersize), "/n", vbCrLf)
53350   .OptionsUseFixPapersize = Replace$(hLang.Retrieve("UseFixPapersize", .OptionsUseFixPapersize), "/n", vbCrLf)
53360   .OptionsUserPass = Replace$(hLang.Retrieve("UserPass", .OptionsUserPass), "/n", vbCrLf)
53370   .OptionsUseStandardauthor = Replace$(hLang.Retrieve("UseStandardauthor", .OptionsUseStandardauthor), "/n", vbCrLf)
53380   .OptionsXCFColorsCount01 = Replace$(hLang.Retrieve("XCFColorsCount01", .OptionsXCFColorsCount01), "/n", vbCrLf)
53390   .OptionsXCFColorscount02 = Replace$(hLang.Retrieve("XCFColorscount02", .OptionsXCFColorscount02), "/n", vbCrLf)
53400   .OptionsXCFDescription = Replace$(hLang.Retrieve("XCFDescription", .OptionsXCFDescription), "/n", vbCrLf)
53410   .OptionsXCFSymbol = Replace$(hLang.Retrieve("XCFSymbol", .OptionsXCFSymbol), "/n", vbCrLf)
53420  End With
53430  Set hLang = Nothing
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
50050   .CommonVersion = "0.9.6"
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
51200   .OptionsCancel = "&Cancel"
51210   .OptionsCustomPapersizeHeight = "Height"
51220   .OptionsCustomPapersizeInfo = "Units of 1/72 of an inch."
51230   .OptionsCustomPapersizeWidth = "Width"
51240   .OptionsDirectoriesGSBin = "Ghostscript Binaries"
51250   .OptionsDirectoriesGSFonts = "Ghostscript Fonts"
51260   .OptionsDirectoriesGSLibraries = "Ghostscript Libraries"
51270   .OptionsDirectoriesTempPath = "Temporary Files"
51280   .OptionsDocument = "Document"
51290   .OptionsEPSDescription = "Encapsulated Postscript Format"
51300   .OptionsEPSFiles = "Encapsulated Postscript-Files"
51310   .OptionsEPSSymbol = "EPS"
51320   .OptionsGhostscriptBinariesDirectoryPrompt = "Select Ghostscript Binaries Directory"
51330   .OptionsGhostscriptFontsDirectoryPrompt = "Select Ghostscript Fonts Directory"
51340   .OptionsGhostscriptInternal = "Internal Ghostscript: %1 Ghostscript %2"
51350   .OptionsGhostscriptLibrariesDirectoryPrompt = "Select Ghostscript Libraries Directory"
51360   .OptionsGhostscriptResourceDirectoryPrompt = "Select Ghostscript Resource Directory"
51370   .OptionsGhostscriptversion = "Ghostscript Version"
51380   .OptionsImageSettings = "Settings"
51390   .OptionsJavaPath = "Path to Java Interpreter"
51400   .OptionsJPEGColorscount01 = "16777216 colors (24 Bit)"
51410   .OptionsJPEGColorscount02 = "Greyscale (8 Bit)"
51420   .OptionsJPEGDescription = "JPEG (JFIF) Format. Please use only for single pages."
51430   .OptionsJPEGQuality = "Quality:"
51440   .OptionsJPEGSymbol = "JPEG"
51450   .OptionsLanguagesCurrentLanguage = "Current language"
51460   .OptionsLanguagesDownloadMoreLanguages = "Load more languages from the internet"
51470   .OptionsLanguagesInstall = "Install"
51480   .OptionsLanguagesRefresh = "Refresh List"
51490   .OptionsLanguagesTranslation = "Translation"
51500   .OptionsLanguagesVersion = "Version"
51510   .OptionsNothingToConfigure = "There is nothing to configure."
51520   .OptionsOnePagePerFile = "One page per file (not for pdf and eps files)"
51530   .OptionsOwnerPass = "Owner Password"
51540   .OptionsPassCancel = "Cancel"
51550   .OptionsPassOK = "OK"
51560   .OptionsPCLColorscount01 = "16777216 colors (24bit)"
51570   .OptionsPCLColorscount02 = "2 colors (Black/White)"
51580   .OptionsPCLDescription = "HP PCL-XL Format"
51590   .OptionsPCLSymbol = "PCL"
51600   .OptionsPCXColorscount01 = "4294967296 colors (32 Bit) CMYK"
51610   .OptionsPCXColorscount02 = "16777216 colors (24 Bit)"
51620   .OptionsPCXColorscount03 = "256 colors (8 Bit)"
51630   .OptionsPCXColorscount04 = "16 colors (4 Bit)"
51640   .OptionsPCXColorscount05 = "2 colors (Black\White)"
51650   .OptionsPCXColorscount06 = "Greyscale (8 Bit)"
51660   .OptionsPCXDescription = "PCX Format. Please use only for single pages."
51670   .OptionsPCXSymbol = "PCX"
51680   .OptionsPDFAllowAssembly = "Allow changes to the assembly"
51690   .OptionsPDFAllowDegradedPrinting = "Allow printing in low resolution"
51700   .OptionsPDFAllowFillIn = "Allow filling in form fields"
51710   .OptionsPDFAllowScreenReaders = "Allow screen readers"
51720   .OptionsPDFColors = "Colors"
51730   .OptionsPDFColorsCaption = "Color Options"
51740   .OptionsPDFColorsCMYKtoRGB = "Convert CMYK images to RGB"
51750   .OptionsPDFColorsColorModel01 = "Use Color Model Device RGB"
51760   .OptionsPDFColorsColorModel02 = "Use Color Model Device CMYK"
51770   .OptionsPDFColorsColorModel03 = "Use Color Model Device Grayscale"
51780   .OptionsPDFColorsColorOptions = "Options"
51790   .OptionsPDFColorsPreserveHalftone = "Preserve Halftone Information"
51800   .OptionsPDFColorsPreserveOverprint = "Preserve Overprint Settings"
51810   .OptionsPDFColorsPreserveTransfer = "Preserve Transfer Functions"
51820   .OptionsPDFCompression = "Compression"
51830   .OptionsPDFCompressionCaption = "PDF Compression"
51840   .OptionsPDFCompressionColor = "Color Images"
51850   .OptionsPDFCompressionColorComp = "Compress"
51860   .OptionsPDFCompressionColorComp01 = "Automatic"
51870   .OptionsPDFCompressionColorComp02 = "JPEG-Maximum"
51880   .OptionsPDFCompressionColorComp03 = "JPEG-High"
51890   .OptionsPDFCompressionColorComp04 = "JPEG-Medium"
51900   .OptionsPDFCompressionColorComp05 = "JPEG-Low"
51910   .OptionsPDFCompressionColorComp06 = "JPEG-Minimum"
51920   .OptionsPDFCompressionColorComp07 = "ZIP"
51930   .OptionsPDFCompressionColorComp08 = "LZW-Compression"
51940   .OptionsPDFCompressionColorRes = "Resolution"
51950   .OptionsPDFCompressionColorResample = "Resample"
51960   .OptionsPDFCompressionColorResample01 = "Downsample"
51970   .OptionsPDFCompressionColorResample02 = "Average Downsample"
51980   .OptionsPDFCompressionColorResample03 = "Bicubic"
51990   .OptionsPDFCompressionGrey = "Greyscale Images"
52000   .OptionsPDFCompressionGreyComp = "Compress"
52010   .OptionsPDFCompressionGreyComp01 = "Automatic"
52020   .OptionsPDFCompressionGreyComp02 = "JPEG-Maximum"
52030   .OptionsPDFCompressionGreyComp03 = "JPEG-High"
52040   .OptionsPDFCompressionGreyComp04 = "JPEG-Medium"
52050   .OptionsPDFCompressionGreyComp05 = "JPEG-Low"
52060   .OptionsPDFCompressionGreyComp06 = "JPEG-Minimum"
52070   .OptionsPDFCompressionGreyComp07 = "ZIP"
52080   .OptionsPDFCompressionGreyComp08 = "LZW-Compression"
52090   .OptionsPDFCompressionGreyRes = "Resolution"
52100   .OptionsPDFCompressionGreyResample = "Resample"
52110   .OptionsPDFCompressionGreyResample01 = "Downsample"
52120   .OptionsPDFCompressionGreyResample02 = "Average Downsample"
52130   .OptionsPDFCompressionGreyResample03 = "Bicubic"
52140   .OptionsPDFCompressionMono = "Monochrome Images"
52150   .OptionsPDFCompressionMonoComp = "Compress"
52160   .OptionsPDFCompressionMonoComp01 = "CCITT Fax Compression"
52170   .OptionsPDFCompressionMonoComp02 = "ZIP"
52180   .OptionsPDFCompressionMonoComp03 = "Run-Length-Encoding"
52190   .OptionsPDFCompressionMonoComp04 = "LZW-Compression"
52200   .OptionsPDFCompressionMonoRes = "Resolution"
52210   .OptionsPDFCompressionMonoResample = "Resample"
52220   .OptionsPDFCompressionMonoResample01 = "Downsample"
52230   .OptionsPDFCompressionMonoResample02 = "Average Downsample"
52240   .OptionsPDFCompressionMonoResample03 = "Bicubic"
52250   .OptionsPDFCompressionTextComp = "Compress Text Objects"
52260   .OptionsPDFDescription = "Adobe PDF Format"
52270   .OptionsPDFDisallowCopy = "Copy text and images"
52280   .OptionsPDFDisallowModify = "Modify the document"
52290   .OptionsPDFDisallowModifyComments = "Modify comments"
52300   .OptionsPDFDisallowPrint = "Print the document"
52310   .OptionsPDFDisallowUser = "Disallow User to"
52320   .OptionsPDFEncryptionHigh = "High (128 Bit - Adobe Acrobat 5.0 and above)"
52330   .OptionsPDFEncryptionLevel = "Encryption Level"
52340   .OptionsPDFEncryptionLow = "Low (40 Bit - Adobe Acrobat 3.0 and above)"
52350   .OptionsPDFEncryptor = "Encryptor"
52360   .OptionsPDFEnhancedPermissions = "Enhanced Permissions (128 Bit only)"
52370   .OptionsPDFEnterPasswords = "Enter Passwords"
52380   .OptionsPDFFonts = "Fonts"
52390   .OptionsPDFFontsCaption = "Font Options"
52400   .OptionsPDFFontsEmbedAll = "Embed all fonts"
52410   .OptionsPDFFontsSubSetFonts = "Subset fonts when percentage of used characters below:"
52420   .OptionsPDFGeneral = "General"
52430   .OptionsPDFGeneralASCII85 = "Convert binary data to ASCII85"
52440   .OptionsPDFGeneralAutorotate = "Auto-Rotate Pages:"
52450   .OptionsPDFGeneralCaption = "General Options"
52460   .OptionsPDFGeneralCompatibility = "Compatibility:"
52470   .OptionsPDFGeneralCompatibility01 = "Adobe Acrobat 3.0 (PDF 1.2)"
52480   .OptionsPDFGeneralCompatibility02 = "Adobe Acrobat 4.0 (PDF 1.3)"
52490   .OptionsPDFGeneralCompatibility03 = "Adobe Acrobat 5.0 (PDF 1.4)"
52500   .OptionsPDFGeneralCompatibility04 = "Adobe Acrobat 6.0 (PDF 1.5)"
52510   .OptionsPDFGeneralDefaultSettings = "Default settings"
52520   .OptionsPDFGeneralDefaultSettingsDefault = "Default"
52530   .OptionsPDFGeneralDefaultSettingsEbook = "Ebook"
52540   .OptionsPDFGeneralDefaultSettingsPrepress = "Pre-press"
52550   .OptionsPDFGeneralDefaultSettingsPrinter = "Printer"
52560   .OptionsPDFGeneralDefaultSettingsScreen = "Screen"
52570   .OptionsPDFGeneralOverprint = "Overprint:"
52580   .OptionsPDFGeneralOverprint01 = "Non-Zero Overprint"
52590   .OptionsPDFGeneralOverprint02 = "Full Overprint"
52600   .OptionsPDFGeneralResolution = "Resolution:"
52610   .OptionsPDFGeneralRotate01 = "None"
52620   .OptionsPDFGeneralRotate02 = "All"
52630   .OptionsPDFGeneralRotate03 = "Single Page"
52640   .OptionsPDFOptimize = "Fast web view"
52650   .OptionsPDFOptions = "PDF Options"
52660   .OptionsPDFOwnerPass = "Password required to change permissions and passwords"
52670   .OptionsPDFPasswords = "Passwords"
52680   .OptionsPDFRepeatPassword = "Repeat"
52690   .OptionsPDFSecurity = "Security"
52700   .OptionsPDFSecurityCaption = "Security"
52710   .OptionsPDFSetPassword = "Password"
52720   .OptionsPDFSigning = "Signing"
52730   .OptionsPDFSigningCaption = "Signing of PDFs"
52740   .OptionsPDFSigningPfxFile = "Pfx\P12 file"
52750   .OptionsPDFSigningSignatureContact = "Signature contact"
52760   .OptionsPDFSigningSignatureLocation = "Signature location"
52770   .OptionsPDFSigningSignatureMultiSignature = "Multi signature allowed"
52780   .OptionsPDFSigningSignaturePosition = "Signature position"
52790   .OptionsPDFSigningSignaturePositionLeftX = "LeftX"
52800   .OptionsPDFSigningSignaturePositionLeftY = "LeftY"
52810   .OptionsPDFSigningSignaturePositionRightX = "RightX"
52820   .OptionsPDFSigningSignaturePositionRightY = "RightY"
52830   .OptionsPDFSigningSignatureReason = "Signature reason"
52840   .OptionsPDFSigningSignatureVisible = "Signature visible in pdf file"
52850   .OptionsPDFSigningSignPdfFile = "Sign pdf file"
52860   .OptionsPDFSymbol = "PDF"
52870   .OptionsPDFUserPass = "Password required to open document"
52880   .OptionsPDFUseSecurity = "Use Security"
52890   .OptionsPNGColorscount01 = "16777216 colors (24 Bit)"
52900   .OptionsPNGColorscount02 = "256 colors (8 Bit)"
52910   .OptionsPNGColorscount03 = "16 colors (4 Bit)"
52920   .OptionsPNGColorscount04 = "2 colors (2 Bit - Black/White)"
52930   .OptionsPNGColorscount05 = "Greyscale (8 Bit)"
52940   .OptionsPNGDescription = "PNG Format. Please use only for single pages."
52950   .OptionsPNGFiles = "Bitmap PNG-Files"
52960   .OptionsPNGSymbol = "PNG"
52970   .OptionsPrintAfterSaving = "Print after saving"
52980   .OptionsPrintAfterSavingDuplex = "Duplex"
52990   .OptionsPrintAfterSavingDuplexTumbleOff = "Don't use tumble (Default)"
53000   .OptionsPrintAfterSavingDuplexTumbleOn = "Use tumble"
53010   .OptionsPrintAfterSavingNoCancel = "Hide the progress dialog during printing"
53020   .OptionsPrintAfterSavingPrinter = "Printer"
53030   .OptionsPrintAfterSavingQueryUser = "Query user"
53040   .OptionsPrintAfterSavingQueryUserDefaultPrinter = "Select the default Windows printer without any user interaction"
53050   .OptionsPrintAfterSavingQueryUserOff = "Off (Default)"
53060   .OptionsPrintAfterSavingQueryUserPrinterSetupDialog = "Shows the printer setup dialog"
53070   .OptionsPrintAfterSavingQueryUserStandardPrinterDialog = "Show the standard printer dialog"
53080   .OptionsPrintertempDirectoryPrompt = "Select Printer Temp-Directory"
53090   .OptionsPrintTestpage = "Print Test Page"
53100   .OptionsProcesspriority = "Process priority"
53110   .OptionsProcesspriorityHigh = "High"
53120   .OptionsProcesspriorityIdle = "Idle"
53130   .OptionsProcesspriorityNormal = "Normal"
53140   .OptionsProcesspriorityRealtime = "Realtime"
53150   .OptionsProgramActionsDescription = "Define an action before and after saving a file."
53160   .OptionsProgramActionsSymbol = "Actions"
53170   .OptionsProgramAutosaveDescription = "Auto-save mode. Auto-save does not prompt for a filename and file location. It automatically saves all PDF files to a single directory with a predefined filename."
53180   .OptionsProgramAutosaveSymbol = "Auto-save"
53190   .OptionsProgramDirectoriesDescription = "Directories for Ghostscript, temporary files and others."
53200   .OptionsProgramDirectoriesSymbol = "Directories"
53210   .OptionsProgramDocumentDescription = "Document properties"
53220   .OptionsProgramDocumentDescription1 = "Document properties 1"
53230   .OptionsProgramDocumentDescription2 = "Document properties 2"
53240   .OptionsProgramDocumentSymbol = "Document"
53250   .OptionsProgramFont = "Program Font"
53260   .OptionsProgramFontCancelTest = "Cancel Test"
53270   .OptionsProgramFontcharset = "Character Set"
53280   .OptionsProgramFontDescription = "Font for labels, captions and values. For the program menu use the general settings in your Windows OS."
53290   .OptionsProgramFontSize = "Size"
53300   .OptionsProgramFontSymbol = "Program font"
53310   .OptionsProgramFontTest = "Test"
53320   .OptionsProgramFontTestdescription = "Here you can test the font."
53330   .OptionsProgramGeneralDescription = "General Settings"
53340   .OptionsProgramGeneralDescription1 = "General Settings 1"
53350   .OptionsProgramGeneralDescription2 = "General Settings 2"
53360   .OptionsProgramGeneralSymbol = "General settings"
53370   .OptionsProgramGhostscriptDescription = "Ghostscript"
53380   .OptionsProgramGhostscriptSymbol = "Ghostscript"
53390   .OptionsProgramLanguagesDescription = "Define the language and download another languages from the internet."
53400   .OptionsProgramLanguagesSymbol = "Languages"
53410   .OptionsProgramNoProcessingAtStartup = "No processing at startup"
53420   .OptionsProgramOptionsDesign = "Frame color of the options dialog"
53430   .OptionsProgramOptionsDesignGradient = "Red and blue gradient (Default)"
53440   .OptionsProgramOptionsDesignSimple = "Simple red and blue color"
53450   .OptionsProgramPrintDescription = "Print after saving"
53460   .OptionsProgramPrintSymbol = "Print"
53470   .OptionsProgramRunProgramAfterSavingCaption = "Action after saving"
53480   .OptionsProgramRunProgramAfterSavingProgram = "Program/Script"
53490   .OptionsProgramRunProgramAfterSavingProgramParameters = "Program parameters"
53500   .OptionsProgramRunProgramAfterSavingWaitUntilReady = "Wait until the program/script is ready"
53510   .OptionsProgramRunProgramAfterSavingWindowstyle = "Window style"
53520   .OptionsProgramRunProgramAfterSavingWindowstyleHide = "Hide"
53530   .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus = "Maximized/Focus"
53540   .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus = "Minimized/Focus"
53550   .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus = "Minimized/No focus"
53560   .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus = "Normal/Focus"
53570   .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus = "Normal/No focus"
53580   .OptionsProgramRunProgramBeforeSavingCaption = "Action before saving"
53590   .OptionsProgramRunProgramBeforeSavingProgram = "Program/Script"
53600   .OptionsProgramRunProgramBeforeSavingProgramParameters = "Program parameters"
53610   .OptionsProgramRunProgramBeforeSavingWaitUntilReady = "Wait until the program/script is ready"
53620   .OptionsProgramRunProgramBeforeSavingWindowstyle = "Window style"
53630   .OptionsProgramRunProgramBeforeSavingWindowstyleHide = "Hide"
53640   .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus = "Maximized/Focus"
53650   .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus = "Minimized/Focus"
53660   .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus = "Minimized/NoFocus"
53670   .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus = "Normal/Focus"
53680   .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus = "Normal/NoFocus"
53690   .OptionsProgramSaveDescription = "Save"
53700   .OptionsProgramSaveSymbol = "Save"
53710   .OptionsProgramShowAnimation = "Show an animation during the process"
53720   .OptionsProgramSwitchingDefaultprinter = "No confirm message switching PDFCreator temporarily as default printer."
53730   .OptionsPSDColorsCount01 = "4294967296 colors (32 Bit) CMYK"
53740   .OptionsPSDColorscount02 = "16777216 colors (24 Bit)"
53750   .OptionsPSDDescription = "Photoshop Format"
53760   .OptionsPSDescription = "Postscript Format"
53770   .OptionsPSDSymbol = "PSD"
53780   .OptionsPSFiles = "Postscript-Files"
53790   .OptionsPSLanguageLevel = "Language Level:"
53800   .OptionsPSSymbol = "PS"
53810   .OptionsRAWColorsCount01 = "4294967296 colors (32 Bit) CMYK"
53820   .OptionsRAWColorscount02 = "16777216 colors (24 Bit)"
53830   .OptionsRAWColorscount03 = "2 colors (Black/White)"
53840   .OptionsRAWDescription = "Raw Format"
53850   .OptionsRAWSymbol = "Raw"
53860   .OptionsRemoveSpaces = "Remove leading and trailing spaces"
53870   .OptionsReset = "&Reset all settings"
53880   .OptionsSave = "&Save"
53890   .OptionsSaveFilename = "Filename"
53900   .OptionsSaveFilenameAdd = "Add"
53910   .OptionsSaveFilenameChange = "Change"
53920   .OptionsSaveFilenameDelete = "Delete"
53930   .OptionsSaveFilenameSubstitutions = "Filename substitution"
53940   .OptionsSaveFilenameSubstitutionsTitle = "Filename substitution only in <Title>"
53950   .OptionsSaveFilenameTokens = "Add a Filename-Token"
53960   .OptionsSavePasswords = "Save passwords temporarily for this session."
53970   .OptionsSendEmailAfterAutosave = "Send an email after auto-saving"
53980   .OptionsSendMailMethod = "Methode to send an email"
53990   .OptionsSendMailMethodAutomatic = "Automatic"
54000   .OptionsSendMailMethodMapi = "Mapi interface"
54010   .OptionsSendMailMethodSendmailDLL = "Using sendmail.dll"
54020   .OptionsShellIntegration = "Shell integration"
54030   .OptionsShellIntegrationAdd = "Integrate PDFCreator into shell"
54040   .OptionsShellIntegrationCaption = "Create &PDF with PDFCreator"
54050   .OptionsShellIntegrationRemove = "Remove shell integration"
54060   .OptionsStamp = "Stamp"
54070   .OptionsStampFontColor = "Font-color"
54080   .OptionsStampOutlineFontThickness = "Outline font thickness"
54090   .OptionsStampString = "Stampstring"
54100   .OptionsStampUseOutlineFont = "Use outline font"
54110   .OptionsStandardAuthorToken = "Add a Author-Token"
54120   .OptionsStandardSaveFormat = "Standard save format"
54130   .OptionsTestpage = "PDFCreator Testpage"
54140   .OptionsTIFFColorscount01 = "16777216 (24 Bit)"
54150   .OptionsTIFFColorscount02 = "4096 (12 Bit)"
54160   .OptionsTIFFColorscount03 = "2 colors (Black/White) G3 fax encoding with no EOLs"
54170   .OptionsTIFFColorscount04 = "2 colors (Black/White) G3 fax encoding with EOLs"
54180   .OptionsTIFFColorscount05 = "2 colors (Black/White) 2-D G3 fax encoding"
54190   .OptionsTIFFColorscount06 = "2 colors (Black/White) G4 fax encoding"
54200   .OptionsTIFFColorscount07 = "2 colors (Black/White) LZW-compatible"
54210   .OptionsTIFFColorscount08 = "2 colors (Black/White) PackBits"
54220   .OptionsTIFFDescription = "TIFF Format. For multipages use the tiff-format."
54230   .OptionsTIFFSymbol = "TIFF"
54240   .OptionsTreeFormats = "Formats"
54250   .OptionsTreeProgram = "Program"
54260   .OptionsTXTDescription = "Text Format"
54270   .OptionsTXTSymbol = "TXT"
54280   .OptionsUseAutosave = "Use Auto-save"
54290   .OptionsUseAutosaveDirectory = "Use this directory for auto-save"
54300   .OptionsUseCreationDateNow = "Use the current Date/Time for 'Creation Date'"
54310   .OptionsUseCustomPapersize = "Use custom paper size"
54320   .OptionsUseFixPapersize = "Use fixed paper size"
54330   .OptionsUserPass = "User Password"
54340   .OptionsUseStandardauthor = "Use standard author"
54350   .OptionsXCFColorsCount01 = "4294967296 colors (32 Bit) CMYK"
54360   .OptionsXCFColorscount02 = "16777216 colors (24 Bit)"
54370   .OptionsXCFDescription = "Gimp Format"
54380   .OptionsXCFSymbol = "XCF"
54390
54400   .PrintingAuthor = "A&uthor:"
54410   .PrintingBMPFiles = "BMP-Files"
54420   .PrintingCancel = "&Cancel"
54430   .PrintingCollect = "&Wait - Collect"
54440   .PrintingCreationDate = "Creation &Date:"
54450   .PrintingDocumentTitle = "Document &Title:"
54460   .PrintingEMail = "&eMail"
54470   .PrintingEPSFiles = "Encapsulated Postscript-Files"
54480   .PrintingJPEGFiles = "JPEG-Files"
54490   .PrintingKeywords = "&Keywords:"
54500   .PrintingModifyDate = "&Modify Date:"
54510   .PrintingNow = "Now"
54520   .PrintingPCLFiles = "PCL (HP PCL-XL)-Files"
54530   .PrintingPCXFiles = "PCX-Files"
54540   .PrintingPDFAFiles = "PDF/A-1b-Files"
54550   .PrintingPDFFiles = "PDF-Files"
54560   .PrintingPDFXFiles = "PDF/X-Files"
54570   .PrintingPNGFiles = "PNG-Files"
54580   .PrintingPSDFiles = "PSD (Adobe Photoshop)-Files"
54590   .PrintingPSFiles = "Postscript-Files"
54600   .PrintingRAWFiles = "RAW (binary format)-Files"
54610   .PrintingSave = "&Save"
54620   .PrintingStartStandardProgram = "&After saving open the document with the default program."
54630   .PrintingStatus = "Creating file..."
54640   .PrintingSubject = "Su&bject:"
54650   .PrintingTIFFFiles = "TIFF-Files"
54660   .PrintingTXTFiles = "Text-Files"
54670   .PrintingXCFFiles = "XCF (Gimp)-Files"
54680
54690   .SaveOpenAttributes = "Attributes"
54700   .SaveOpenCancel = "Cancel"
54710   .SaveOpenFilename = "Filename"
54720   .SaveOpenOpen = "Open"
54730   .SaveOpenOpenTitle = "Open"
54740   .SaveOpenSave = "Save"
54750   .SaveOpenSaveTitle = "Save as"
54760   .SaveOpenSize = "Size"
54770
54780  End With
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

