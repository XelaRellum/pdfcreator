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
 DialogInfoTitle As String
 DialogLanguage As String
 DialogPrinter As String
 DialogPrinterClose As String
 DialogPrinterLogfile As String
 DialogPrinterLogfiles As String
 DialogPrinterLogging As String
 DialogPrinterOptions As String
 DialogPrinterPrinters As String
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
 MessagesMsg40 As String
 MessagesMsg41 As String
 MessagesMsg42 As String
 MessagesMsg43 As String
 MessagesMsg44 As String

 OptionsAdditionalGhostscriptParameters As String
 OptionsAdditionalGhostscriptSearchpath As String
 OptionsAddWindowsFontpath As String
 OptionsAllowSpecialGSCharsInFilenames As String
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
 OptionsBMPColorscount05_2 As String
 OptionsBMPColorscount06_2 As String
 OptionsBMPColorscount07 As String
 OptionsBMPColorscount08 As String
 OptionsBMPDescription As String
 OptionsBMPSymbol As String
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
 OptionsEnableNotice As String
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
 OptionsPDFCompressionColorComp09 As String
 OptionsPDFCompressionColorCompFac As String
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
 OptionsPDFCompressionGreyComp09 As String
 OptionsPDFCompressionGreyCompFac As String
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
 OptionsPDFEncryptionAes128 As String
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
 OptionsPDFOwnerPasswordShowChars As String
 OptionsPDFPasswords As String
 OptionsPDFRepeatPassword As String
 OptionsPDFSecurity As String
 OptionsPDFSecurityCaption As String
 OptionsPDFSetPassword As String
 OptionsPDFSigning As String
 OptionsPDFSigningCaption As String
 OptionsPDFSigningCerticatePassword As String
 OptionsPDFSigningCerticatePasswordCancel As String
 OptionsPDFSigningCerticatePasswordOk As String
 OptionsPDFSigningCerticatePasswordShowPassword As String
 OptionsPDFSigningCertificateEmptyPassword As String
 OptionsPDFSigningCertificateFile As String
 OptionsPDFSigningChooseCertifcateFile As String
 OptionsPDFSigningEnterCerticatePassword As String
 OptionsPDFSigningP12Files As String
 OptionsPDFSigningPfxFiles As String
 OptionsPDFSigningPfxP12Files As String
 OptionsPDFSigningSignatureContact As String
 OptionsPDFSigningSignatureLocation As String
 OptionsPDFSigningSignatureMultiSignature As String
 OptionsPDFSigningSignatureOnPage As String
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
 OptionsPDFUserPasswordShowChars As String
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
 OptionsPrintAfterSavingBitsPerPixel As String
 OptionsPrintAfterSavingBitsPerPixelCMYK As String
 OptionsPrintAfterSavingBitsPerPixelMono As String
 OptionsPrintAfterSavingBitsPerPixelTrueColor As String
 OptionsPrintAfterSavingDuplex As String
 OptionsPrintAfterSavingDuplexTumbleOff As String
 OptionsPrintAfterSavingDuplexTumbleOn As String
 OptionsPrintAfterSavingMaxResolution As String
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
 OptionsProfile As String
 OptionsProfileAdd As String
 OptionsProfileCancel As String
 OptionsProfileDefaultName As String
 OptionsProfileDel As String
 OptionsProfileLoadFromDisc As String
 OptionsProfileNewProfile As String
 OptionsProfileOk As String
 OptionsProfileRenameProfile As String
 OptionsProfileSaveToDisc As String
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
 OptionsSVGDescription As String
 OptionsSVGSymbol As String
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
 OptionsToolbar As String
 OptionsToolbarInstall As String
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

 PrintersAdminNotice As String
 PrintersClose As String
 PrintersNewPrinterName As String
 PrintersPrinter As String
 PrintersPrinterAdd As String
 PrintersPrinterDel As String
 PrintersPrinters As String
 PrintersProfile As String
 PrintersSave As String

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
 PrintingProfile As String
 PrintingPSDFiles As String
 PrintingPSFiles As String
 PrintingRAWFiles As String
 PrintingSave As String
 PrintingStartStandardProgram As String
 PrintingStatus As String
 PrintingSubject As String
 PrintingSVGFiles As String
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
 InitLanguagesStrings
 LoadCommonStrings Languagefile
 LoadDialogStrings Languagefile
 LoadListStrings Languagefile
 LoadLoggingStrings Languagefile
 LoadMessagesStrings Languagefile
 LoadOptionsStrings Languagefile
 LoadPrintersStrings Languagefile
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
  .DialogDocumentAddFromClipboard = Replace$(hLang.Retrieve("DocumentAddFromClipboard", .DialogDocumentAddFromClipboard),"/n",vbCrLf)
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
  .DialogInfoTitle = Replace$(hLang.Retrieve("InfoTitle", .DialogInfoTitle),"/n",vbCrLf)
  .DialogLanguage = Replace$(hLang.Retrieve("Language", .DialogLanguage),"/n",vbCrLf)
  .DialogPrinter = Replace$(hLang.Retrieve("Printer", .DialogPrinter),"/n",vbCrLf)
  .DialogPrinterClose = Replace$(hLang.Retrieve("PrinterClose", .DialogPrinterClose),"/n",vbCrLf)
  .DialogPrinterLogfile = Replace$(hLang.Retrieve("PrinterLogfile", .DialogPrinterLogfile),"/n",vbCrLf)
  .DialogPrinterLogfiles = Replace$(hLang.Retrieve("PrinterLogfiles", .DialogPrinterLogfiles),"/n",vbCrLf)
  .DialogPrinterLogging = Replace$(hLang.Retrieve("PrinterLogging", .DialogPrinterLogging),"/n",vbCrLf)
  .DialogPrinterOptions = Replace$(hLang.Retrieve("PrinterOptions", .DialogPrinterOptions),"/n",vbCrLf)
  .DialogPrinterPrinters = Replace$(hLang.Retrieve("PrinterPrinters", .DialogPrinterPrinters),"/n",vbCrLf)
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
  .MessagesMsg37 = Replace$(hLang.Retrieve("Msg37", .MessagesMsg37),"/n",vbCrLf)
  .MessagesMsg38 = Replace$(hLang.Retrieve("Msg38", .MessagesMsg38),"/n",vbCrLf)
  .MessagesMsg39 = Replace$(hLang.Retrieve("Msg39", .MessagesMsg39),"/n",vbCrLf)
  .MessagesMsg40 = Replace$(hLang.Retrieve("Msg40", .MessagesMsg40),"/n",vbCrLf)
  .MessagesMsg41 = Replace$(hLang.Retrieve("Msg41", .MessagesMsg41),"/n",vbCrLf)
  .MessagesMsg42 = Replace$(hLang.Retrieve("Msg42", .MessagesMsg42),"/n",vbCrLf)
  .MessagesMsg43 = Replace$(hLang.Retrieve("Msg43", .MessagesMsg43),"/n",vbCrLf)
  .MessagesMsg44 = Replace$(hLang.Retrieve("Msg44", .MessagesMsg44),"/n",vbCrLf)
 End With
 Set hLang = Nothing
End Sub

Private Sub LoadOptionsStrings(ByVal Languagefile As String)
 Dim hLang As New clsHash
 ReadINISection Languagefile, "Options", hLang
 With LanguageStrings
  .OptionsAdditionalGhostscriptParameters = Replace$(hLang.Retrieve("AdditionalGhostscriptParameters", .OptionsAdditionalGhostscriptParameters),"/n",vbCrLf)
  .OptionsAdditionalGhostscriptSearchpath = Replace$(hLang.Retrieve("AdditionalGhostscriptSearchpath", .OptionsAdditionalGhostscriptSearchpath),"/n",vbCrLf)
  .OptionsAddWindowsFontpath = Replace$(hLang.Retrieve("AddWindowsFontpath", .OptionsAddWindowsFontpath),"/n",vbCrLf)
  .OptionsAllowSpecialGSCharsInFilenames = Replace$(hLang.Retrieve("AllowSpecialGSCharsInFilenames", .OptionsAllowSpecialGSCharsInFilenames),"/n",vbCrLf)
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
  .OptionsBMPColorscount05_2 = Replace$(hLang.Retrieve("BMPColorscount05_2", .OptionsBMPColorscount05_2),"/n",vbCrLf)
  .OptionsBMPColorscount06_2 = Replace$(hLang.Retrieve("BMPColorscount06_2", .OptionsBMPColorscount06_2),"/n",vbCrLf)
  .OptionsBMPColorscount07 = Replace$(hLang.Retrieve("BMPColorscount07", .OptionsBMPColorscount07),"/n",vbCrLf)
  .OptionsBMPColorscount08 = Replace$(hLang.Retrieve("BMPColorscount08", .OptionsBMPColorscount08),"/n",vbCrLf)
  .OptionsBMPDescription = Replace$(hLang.Retrieve("BMPDescription", .OptionsBMPDescription),"/n",vbCrLf)
  .OptionsBMPSymbol = Replace$(hLang.Retrieve("BMPSymbol", .OptionsBMPSymbol),"/n",vbCrLf)
  .OptionsCancel = Replace$(hLang.Retrieve("Cancel", .OptionsCancel),"/n",vbCrLf)
  .OptionsCheckUpdateDescription = Replace$(hLang.Retrieve("CheckUpdateDescription", .OptionsCheckUpdateDescription),"/n",vbCrLf)
  .OptionsCheckUpdateInterval = Replace$(hLang.Retrieve("CheckUpdateInterval", .OptionsCheckUpdateInterval),"/n",vbCrLf)
  .OptionsCheckUpdateInterval01 = Replace$(hLang.Retrieve("CheckUpdateInterval01", .OptionsCheckUpdateInterval01),"/n",vbCrLf)
  .OptionsCheckUpdateInterval02 = Replace$(hLang.Retrieve("CheckUpdateInterval02", .OptionsCheckUpdateInterval02),"/n",vbCrLf)
  .OptionsCheckUpdateInterval03 = Replace$(hLang.Retrieve("CheckUpdateInterval03", .OptionsCheckUpdateInterval03),"/n",vbCrLf)
  .OptionsCheckUpdateInterval04 = Replace$(hLang.Retrieve("CheckUpdateInterval04", .OptionsCheckUpdateInterval04),"/n",vbCrLf)
  .OptionsCheckUpdateNow = Replace$(hLang.Retrieve("CheckUpdateNow", .OptionsCheckUpdateNow),"/n",vbCrLf)
  .OptionsCustomPapersizeHeight = Replace$(hLang.Retrieve("CustomPapersizeHeight", .OptionsCustomPapersizeHeight),"/n",vbCrLf)
  .OptionsCustomPapersizeInfo = Replace$(hLang.Retrieve("CustomPapersizeInfo", .OptionsCustomPapersizeInfo),"/n",vbCrLf)
  .OptionsCustomPapersizeWidth = Replace$(hLang.Retrieve("CustomPapersizeWidth", .OptionsCustomPapersizeWidth),"/n",vbCrLf)
  .OptionsDirectoriesGSBin = Replace$(hLang.Retrieve("DirectoriesGSBin", .OptionsDirectoriesGSBin),"/n",vbCrLf)
  .OptionsDirectoriesGSFonts = Replace$(hLang.Retrieve("DirectoriesGSFonts", .OptionsDirectoriesGSFonts),"/n",vbCrLf)
  .OptionsDirectoriesGSLibraries = Replace$(hLang.Retrieve("DirectoriesGSLibraries", .OptionsDirectoriesGSLibraries),"/n",vbCrLf)
  .OptionsDirectoriesTempPath = Replace$(hLang.Retrieve("DirectoriesTempPath", .OptionsDirectoriesTempPath),"/n",vbCrLf)
  .OptionsDocument = Replace$(hLang.Retrieve("Document", .OptionsDocument),"/n",vbCrLf)
  .OptionsEnableNotice = Replace$(hLang.Retrieve("EnableNotice", .OptionsEnableNotice),"/n",vbCrLf)
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
  .OptionsLanguagesCurrentLanguage = Replace$(hLang.Retrieve("LanguagesCurrentLanguage", .OptionsLanguagesCurrentLanguage),"/n",vbCrLf)
  .OptionsLanguagesDownloadMoreLanguages = Replace$(hLang.Retrieve("LanguagesDownloadMoreLanguages", .OptionsLanguagesDownloadMoreLanguages),"/n",vbCrLf)
  .OptionsLanguagesInstall = Replace$(hLang.Retrieve("LanguagesInstall", .OptionsLanguagesInstall),"/n",vbCrLf)
  .OptionsLanguagesRefresh = Replace$(hLang.Retrieve("LanguagesRefresh", .OptionsLanguagesRefresh),"/n",vbCrLf)
  .OptionsLanguagesTranslation = Replace$(hLang.Retrieve("LanguagesTranslation", .OptionsLanguagesTranslation),"/n",vbCrLf)
  .OptionsLanguagesVersion = Replace$(hLang.Retrieve("LanguagesVersion", .OptionsLanguagesVersion),"/n",vbCrLf)
  .OptionsNothingToConfigure = Replace$(hLang.Retrieve("NothingToConfigure", .OptionsNothingToConfigure),"/n",vbCrLf)
  .OptionsOnePagePerFile = Replace$(hLang.Retrieve("OnePagePerFile", .OptionsOnePagePerFile),"/n",vbCrLf)
  .OptionsOwnerPass = Replace$(hLang.Retrieve("OwnerPass", .OptionsOwnerPass),"/n",vbCrLf)
  .OptionsPassCancel = Replace$(hLang.Retrieve("PassCancel", .OptionsPassCancel),"/n",vbCrLf)
  .OptionsPassOK = Replace$(hLang.Retrieve("PassOK", .OptionsPassOK),"/n",vbCrLf)
  .OptionsPCLColorscount01 = Replace$(hLang.Retrieve("PCLColorscount01", .OptionsPCLColorscount01),"/n",vbCrLf)
  .OptionsPCLColorscount02 = Replace$(hLang.Retrieve("PCLColorscount02", .OptionsPCLColorscount02),"/n",vbCrLf)
  .OptionsPCLDescription = Replace$(hLang.Retrieve("PCLDescription", .OptionsPCLDescription),"/n",vbCrLf)
  .OptionsPCLSymbol = Replace$(hLang.Retrieve("PCLSymbol", .OptionsPCLSymbol),"/n",vbCrLf)
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
  .OptionsPDFCompressionColorComp09 = Replace$(hLang.Retrieve("PDFCompressionColorComp09", .OptionsPDFCompressionColorComp09),"/n",vbCrLf)
  .OptionsPDFCompressionColorCompFac = Replace$(hLang.Retrieve("PDFCompressionColorCompFac", .OptionsPDFCompressionColorCompFac),"/n",vbCrLf)
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
  .OptionsPDFCompressionGreyComp09 = Replace$(hLang.Retrieve("PDFCompressionGreyComp09", .OptionsPDFCompressionGreyComp09),"/n",vbCrLf)
  .OptionsPDFCompressionGreyCompFac = Replace$(hLang.Retrieve("PDFCompressionGreyCompFac", .OptionsPDFCompressionGreyCompFac),"/n",vbCrLf)
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
  .OptionsPDFEncryptionAes128 = Replace$(hLang.Retrieve("PDFEncryptionAes128", .OptionsPDFEncryptionAes128),"/n",vbCrLf)
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
  .OptionsPDFGeneralCompatibility04 = Replace$(hLang.Retrieve("PDFGeneralCompatibility04", .OptionsPDFGeneralCompatibility04),"/n",vbCrLf)
  .OptionsPDFGeneralDefaultSettings = Replace$(hLang.Retrieve("PDFGeneralDefaultSettings", .OptionsPDFGeneralDefaultSettings),"/n",vbCrLf)
  .OptionsPDFGeneralDefaultSettingsDefault = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsDefault", .OptionsPDFGeneralDefaultSettingsDefault),"/n",vbCrLf)
  .OptionsPDFGeneralDefaultSettingsEbook = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsEbook", .OptionsPDFGeneralDefaultSettingsEbook),"/n",vbCrLf)
  .OptionsPDFGeneralDefaultSettingsPrepress = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsPrepress", .OptionsPDFGeneralDefaultSettingsPrepress),"/n",vbCrLf)
  .OptionsPDFGeneralDefaultSettingsPrinter = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsPrinter", .OptionsPDFGeneralDefaultSettingsPrinter),"/n",vbCrLf)
  .OptionsPDFGeneralDefaultSettingsScreen = Replace$(hLang.Retrieve("PDFGeneralDefaultSettingsScreen", .OptionsPDFGeneralDefaultSettingsScreen),"/n",vbCrLf)
  .OptionsPDFGeneralOverprint = Replace$(hLang.Retrieve("PDFGeneralOverprint", .OptionsPDFGeneralOverprint),"/n",vbCrLf)
  .OptionsPDFGeneralOverprint01 = Replace$(hLang.Retrieve("PDFGeneralOverprint01", .OptionsPDFGeneralOverprint01),"/n",vbCrLf)
  .OptionsPDFGeneralOverprint02 = Replace$(hLang.Retrieve("PDFGeneralOverprint02", .OptionsPDFGeneralOverprint02),"/n",vbCrLf)
  .OptionsPDFGeneralResolution = Replace$(hLang.Retrieve("PDFGeneralResolution", .OptionsPDFGeneralResolution),"/n",vbCrLf)
  .OptionsPDFGeneralRotate01 = Replace$(hLang.Retrieve("PDFGeneralRotate01", .OptionsPDFGeneralRotate01),"/n",vbCrLf)
  .OptionsPDFGeneralRotate02 = Replace$(hLang.Retrieve("PDFGeneralRotate02", .OptionsPDFGeneralRotate02),"/n",vbCrLf)
  .OptionsPDFGeneralRotate03 = Replace$(hLang.Retrieve("PDFGeneralRotate03", .OptionsPDFGeneralRotate03),"/n",vbCrLf)
  .OptionsPDFOptimize = Replace$(hLang.Retrieve("PDFOptimize", .OptionsPDFOptimize),"/n",vbCrLf)
  .OptionsPDFOptions = Replace$(hLang.Retrieve("PDFOptions", .OptionsPDFOptions),"/n",vbCrLf)
  .OptionsPDFOwnerPass = Replace$(hLang.Retrieve("PDFOwnerPass", .OptionsPDFOwnerPass),"/n",vbCrLf)
  .OptionsPDFOwnerPasswordShowChars = Replace$(hLang.Retrieve("PDFOwnerPasswordShowChars", .OptionsPDFOwnerPasswordShowChars),"/n",vbCrLf)
  .OptionsPDFPasswords = Replace$(hLang.Retrieve("PDFPasswords", .OptionsPDFPasswords),"/n",vbCrLf)
  .OptionsPDFRepeatPassword = Replace$(hLang.Retrieve("PDFRepeatPassword", .OptionsPDFRepeatPassword),"/n",vbCrLf)
  .OptionsPDFSecurity = Replace$(hLang.Retrieve("PDFSecurity", .OptionsPDFSecurity),"/n",vbCrLf)
  .OptionsPDFSecurityCaption = Replace$(hLang.Retrieve("PDFSecurityCaption", .OptionsPDFSecurityCaption),"/n",vbCrLf)
  .OptionsPDFSetPassword = Replace$(hLang.Retrieve("PDFSetPassword", .OptionsPDFSetPassword),"/n",vbCrLf)
  .OptionsPDFSigning = Replace$(hLang.Retrieve("PDFSigning", .OptionsPDFSigning),"/n",vbCrLf)
  .OptionsPDFSigningCaption = Replace$(hLang.Retrieve("PDFSigningCaption", .OptionsPDFSigningCaption),"/n",vbCrLf)
  .OptionsPDFSigningCerticatePassword = Replace$(hLang.Retrieve("PDFSigningCerticatePassword", .OptionsPDFSigningCerticatePassword),"/n",vbCrLf)
  .OptionsPDFSigningCerticatePasswordCancel = Replace$(hLang.Retrieve("PDFSigningCerticatePasswordCancel", .OptionsPDFSigningCerticatePasswordCancel),"/n",vbCrLf)
  .OptionsPDFSigningCerticatePasswordOk = Replace$(hLang.Retrieve("PDFSigningCerticatePasswordOk", .OptionsPDFSigningCerticatePasswordOk),"/n",vbCrLf)
  .OptionsPDFSigningCerticatePasswordShowPassword = Replace$(hLang.Retrieve("PDFSigningCerticatePasswordShowPassword", .OptionsPDFSigningCerticatePasswordShowPassword),"/n",vbCrLf)
  .OptionsPDFSigningCertificateEmptyPassword = Replace$(hLang.Retrieve("PDFSigningCertificateEmptyPassword", .OptionsPDFSigningCertificateEmptyPassword),"/n",vbCrLf)
  .OptionsPDFSigningCertificateFile = Replace$(hLang.Retrieve("PDFSigningCertificateFile", .OptionsPDFSigningCertificateFile),"/n",vbCrLf)
  .OptionsPDFSigningChooseCertifcateFile = Replace$(hLang.Retrieve("PDFSigningChooseCertifcateFile", .OptionsPDFSigningChooseCertifcateFile),"/n",vbCrLf)
  .OptionsPDFSigningEnterCerticatePassword = Replace$(hLang.Retrieve("PDFSigningEnterCerticatePassword", .OptionsPDFSigningEnterCerticatePassword),"/n",vbCrLf)
  .OptionsPDFSigningP12Files = Replace$(hLang.Retrieve("PDFSigningP12Files", .OptionsPDFSigningP12Files),"/n",vbCrLf)
  .OptionsPDFSigningPfxFiles = Replace$(hLang.Retrieve("PDFSigningPfxFiles", .OptionsPDFSigningPfxFiles),"/n",vbCrLf)
  .OptionsPDFSigningPfxP12Files = Replace$(hLang.Retrieve("PDFSigningPfxP12Files", .OptionsPDFSigningPfxP12Files),"/n",vbCrLf)
  .OptionsPDFSigningSignatureContact = Replace$(hLang.Retrieve("PDFSigningSignatureContact", .OptionsPDFSigningSignatureContact),"/n",vbCrLf)
  .OptionsPDFSigningSignatureLocation = Replace$(hLang.Retrieve("PDFSigningSignatureLocation", .OptionsPDFSigningSignatureLocation),"/n",vbCrLf)
  .OptionsPDFSigningSignatureMultiSignature = Replace$(hLang.Retrieve("PDFSigningSignatureMultiSignature", .OptionsPDFSigningSignatureMultiSignature),"/n",vbCrLf)
  .OptionsPDFSigningSignatureOnPage = Replace$(hLang.Retrieve("PDFSigningSignatureOnPage", .OptionsPDFSigningSignatureOnPage),"/n",vbCrLf)
  .OptionsPDFSigningSignaturePosition = Replace$(hLang.Retrieve("PDFSigningSignaturePosition", .OptionsPDFSigningSignaturePosition),"/n",vbCrLf)
  .OptionsPDFSigningSignaturePositionLeftX = Replace$(hLang.Retrieve("PDFSigningSignaturePositionLeftX", .OptionsPDFSigningSignaturePositionLeftX),"/n",vbCrLf)
  .OptionsPDFSigningSignaturePositionLeftY = Replace$(hLang.Retrieve("PDFSigningSignaturePositionLeftY", .OptionsPDFSigningSignaturePositionLeftY),"/n",vbCrLf)
  .OptionsPDFSigningSignaturePositionRightX = Replace$(hLang.Retrieve("PDFSigningSignaturePositionRightX", .OptionsPDFSigningSignaturePositionRightX),"/n",vbCrLf)
  .OptionsPDFSigningSignaturePositionRightY = Replace$(hLang.Retrieve("PDFSigningSignaturePositionRightY", .OptionsPDFSigningSignaturePositionRightY),"/n",vbCrLf)
  .OptionsPDFSigningSignatureReason = Replace$(hLang.Retrieve("PDFSigningSignatureReason", .OptionsPDFSigningSignatureReason),"/n",vbCrLf)
  .OptionsPDFSigningSignatureVisible = Replace$(hLang.Retrieve("PDFSigningSignatureVisible", .OptionsPDFSigningSignatureVisible),"/n",vbCrLf)
  .OptionsPDFSigningSignPdfFile = Replace$(hLang.Retrieve("PDFSigningSignPdfFile", .OptionsPDFSigningSignPdfFile),"/n",vbCrLf)
  .OptionsPDFSymbol = Replace$(hLang.Retrieve("PDFSymbol", .OptionsPDFSymbol),"/n",vbCrLf)
  .OptionsPDFUserPass = Replace$(hLang.Retrieve("PDFUserPass", .OptionsPDFUserPass),"/n",vbCrLf)
  .OptionsPDFUserPasswordShowChars = Replace$(hLang.Retrieve("PDFUserPasswordShowChars", .OptionsPDFUserPasswordShowChars),"/n",vbCrLf)
  .OptionsPDFUseSecurity = Replace$(hLang.Retrieve("PDFUseSecurity", .OptionsPDFUseSecurity),"/n",vbCrLf)
  .OptionsPNGColorscount01 = Replace$(hLang.Retrieve("PNGColorscount01", .OptionsPNGColorscount01),"/n",vbCrLf)
  .OptionsPNGColorscount02 = Replace$(hLang.Retrieve("PNGColorscount02", .OptionsPNGColorscount02),"/n",vbCrLf)
  .OptionsPNGColorscount03 = Replace$(hLang.Retrieve("PNGColorscount03", .OptionsPNGColorscount03),"/n",vbCrLf)
  .OptionsPNGColorscount04 = Replace$(hLang.Retrieve("PNGColorscount04", .OptionsPNGColorscount04),"/n",vbCrLf)
  .OptionsPNGColorscount05 = Replace$(hLang.Retrieve("PNGColorscount05", .OptionsPNGColorscount05),"/n",vbCrLf)
  .OptionsPNGDescription = Replace$(hLang.Retrieve("PNGDescription", .OptionsPNGDescription),"/n",vbCrLf)
  .OptionsPNGFiles = Replace$(hLang.Retrieve("PNGFiles", .OptionsPNGFiles),"/n",vbCrLf)
  .OptionsPNGSymbol = Replace$(hLang.Retrieve("PNGSymbol", .OptionsPNGSymbol),"/n",vbCrLf)
  .OptionsPrintAfterSaving = Replace$(hLang.Retrieve("PrintAfterSaving", .OptionsPrintAfterSaving),"/n",vbCrLf)
  .OptionsPrintAfterSavingBitsPerPixel = Replace$(hLang.Retrieve("PrintAfterSavingBitsPerPixel", .OptionsPrintAfterSavingBitsPerPixel),"/n",vbCrLf)
  .OptionsPrintAfterSavingBitsPerPixelCMYK = Replace$(hLang.Retrieve("PrintAfterSavingBitsPerPixelCMYK", .OptionsPrintAfterSavingBitsPerPixelCMYK),"/n",vbCrLf)
  .OptionsPrintAfterSavingBitsPerPixelMono = Replace$(hLang.Retrieve("PrintAfterSavingBitsPerPixelMono", .OptionsPrintAfterSavingBitsPerPixelMono),"/n",vbCrLf)
  .OptionsPrintAfterSavingBitsPerPixelTrueColor = Replace$(hLang.Retrieve("PrintAfterSavingBitsPerPixelTrueColor", .OptionsPrintAfterSavingBitsPerPixelTrueColor),"/n",vbCrLf)
  .OptionsPrintAfterSavingDuplex = Replace$(hLang.Retrieve("PrintAfterSavingDuplex", .OptionsPrintAfterSavingDuplex),"/n",vbCrLf)
  .OptionsPrintAfterSavingDuplexTumbleOff = Replace$(hLang.Retrieve("PrintAfterSavingDuplexTumbleOff", .OptionsPrintAfterSavingDuplexTumbleOff),"/n",vbCrLf)
  .OptionsPrintAfterSavingDuplexTumbleOn = Replace$(hLang.Retrieve("PrintAfterSavingDuplexTumbleOn", .OptionsPrintAfterSavingDuplexTumbleOn),"/n",vbCrLf)
  .OptionsPrintAfterSavingMaxResolution = Replace$(hLang.Retrieve("PrintAfterSavingMaxResolution", .OptionsPrintAfterSavingMaxResolution),"/n",vbCrLf)
  .OptionsPrintAfterSavingNoCancel = Replace$(hLang.Retrieve("PrintAfterSavingNoCancel", .OptionsPrintAfterSavingNoCancel),"/n",vbCrLf)
  .OptionsPrintAfterSavingPrinter = Replace$(hLang.Retrieve("PrintAfterSavingPrinter", .OptionsPrintAfterSavingPrinter),"/n",vbCrLf)
  .OptionsPrintAfterSavingQueryUser = Replace$(hLang.Retrieve("PrintAfterSavingQueryUser", .OptionsPrintAfterSavingQueryUser),"/n",vbCrLf)
  .OptionsPrintAfterSavingQueryUserDefaultPrinter = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserDefaultPrinter", .OptionsPrintAfterSavingQueryUserDefaultPrinter),"/n",vbCrLf)
  .OptionsPrintAfterSavingQueryUserOff = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserOff", .OptionsPrintAfterSavingQueryUserOff),"/n",vbCrLf)
  .OptionsPrintAfterSavingQueryUserPrinterSetupDialog = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserPrinterSetupDialog", .OptionsPrintAfterSavingQueryUserPrinterSetupDialog),"/n",vbCrLf)
  .OptionsPrintAfterSavingQueryUserStandardPrinterDialog = Replace$(hLang.Retrieve("PrintAfterSavingQueryUserStandardPrinterDialog", .OptionsPrintAfterSavingQueryUserStandardPrinterDialog),"/n",vbCrLf)
  .OptionsPrintertempDirectoryPrompt = Replace$(hLang.Retrieve("PrintertempDirectoryPrompt", .OptionsPrintertempDirectoryPrompt),"/n",vbCrLf)
  .OptionsPrintTestpage = Replace$(hLang.Retrieve("PrintTestpage", .OptionsPrintTestpage),"/n",vbCrLf)
  .OptionsProcesspriority = Replace$(hLang.Retrieve("Processpriority", .OptionsProcesspriority),"/n",vbCrLf)
  .OptionsProcesspriorityHigh = Replace$(hLang.Retrieve("ProcesspriorityHigh", .OptionsProcesspriorityHigh),"/n",vbCrLf)
  .OptionsProcesspriorityIdle = Replace$(hLang.Retrieve("ProcesspriorityIdle", .OptionsProcesspriorityIdle),"/n",vbCrLf)
  .OptionsProcesspriorityNormal = Replace$(hLang.Retrieve("ProcesspriorityNormal", .OptionsProcesspriorityNormal),"/n",vbCrLf)
  .OptionsProcesspriorityRealtime = Replace$(hLang.Retrieve("ProcesspriorityRealtime", .OptionsProcesspriorityRealtime),"/n",vbCrLf)
  .OptionsProfile = Replace$(hLang.Retrieve("Profile", .OptionsProfile),"/n",vbCrLf)
  .OptionsProfileAdd = Replace$(hLang.Retrieve("ProfileAdd", .OptionsProfileAdd),"/n",vbCrLf)
  .OptionsProfileCancel = Replace$(hLang.Retrieve("ProfileCancel", .OptionsProfileCancel),"/n",vbCrLf)
  .OptionsProfileDefaultName = Replace$(hLang.Retrieve("ProfileDefaultName", .OptionsProfileDefaultName),"/n",vbCrLf)
  .OptionsProfileDel = Replace$(hLang.Retrieve("ProfileDel", .OptionsProfileDel),"/n",vbCrLf)
  .OptionsProfileLoadFromDisc = Replace$(hLang.Retrieve("ProfileLoadFromDisc", .OptionsProfileLoadFromDisc),"/n",vbCrLf)
  .OptionsProfileNewProfile = Replace$(hLang.Retrieve("ProfileNewProfile", .OptionsProfileNewProfile),"/n",vbCrLf)
  .OptionsProfileOk = Replace$(hLang.Retrieve("ProfileOk", .OptionsProfileOk),"/n",vbCrLf)
  .OptionsProfileRenameProfile = Replace$(hLang.Retrieve("ProfileRenameProfile", .OptionsProfileRenameProfile),"/n",vbCrLf)
  .OptionsProfileSaveToDisc = Replace$(hLang.Retrieve("ProfileSaveToDisc", .OptionsProfileSaveToDisc),"/n",vbCrLf)
  .OptionsProgramActionsDescription = Replace$(hLang.Retrieve("ProgramActionsDescription", .OptionsProgramActionsDescription),"/n",vbCrLf)
  .OptionsProgramActionsSymbol = Replace$(hLang.Retrieve("ProgramActionsSymbol", .OptionsProgramActionsSymbol),"/n",vbCrLf)
  .OptionsProgramAutosaveDescription = Replace$(hLang.Retrieve("ProgramAutosaveDescription", .OptionsProgramAutosaveDescription),"/n",vbCrLf)
  .OptionsProgramAutosaveSymbol = Replace$(hLang.Retrieve("ProgramAutosaveSymbol", .OptionsProgramAutosaveSymbol),"/n",vbCrLf)
  .OptionsProgramDirectoriesDescription = Replace$(hLang.Retrieve("ProgramDirectoriesDescription", .OptionsProgramDirectoriesDescription),"/n",vbCrLf)
  .OptionsProgramDirectoriesSymbol = Replace$(hLang.Retrieve("ProgramDirectoriesSymbol", .OptionsProgramDirectoriesSymbol),"/n",vbCrLf)
  .OptionsProgramDocumentDescription = Replace$(hLang.Retrieve("ProgramDocumentDescription", .OptionsProgramDocumentDescription),"/n",vbCrLf)
  .OptionsProgramDocumentDescription1 = Replace$(hLang.Retrieve("ProgramDocumentDescription1", .OptionsProgramDocumentDescription1),"/n",vbCrLf)
  .OptionsProgramDocumentDescription2 = Replace$(hLang.Retrieve("ProgramDocumentDescription2", .OptionsProgramDocumentDescription2),"/n",vbCrLf)
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
  .OptionsProgramGeneralDescription1 = Replace$(hLang.Retrieve("ProgramGeneralDescription1", .OptionsProgramGeneralDescription1),"/n",vbCrLf)
  .OptionsProgramGeneralDescription2 = Replace$(hLang.Retrieve("ProgramGeneralDescription2", .OptionsProgramGeneralDescription2),"/n",vbCrLf)
  .OptionsProgramGeneralSymbol = Replace$(hLang.Retrieve("ProgramGeneralSymbol", .OptionsProgramGeneralSymbol),"/n",vbCrLf)
  .OptionsProgramGhostscriptDescription = Replace$(hLang.Retrieve("ProgramGhostscriptDescription", .OptionsProgramGhostscriptDescription),"/n",vbCrLf)
  .OptionsProgramGhostscriptSymbol = Replace$(hLang.Retrieve("ProgramGhostscriptSymbol", .OptionsProgramGhostscriptSymbol),"/n",vbCrLf)
  .OptionsProgramLanguagesDescription = Replace$(hLang.Retrieve("ProgramLanguagesDescription", .OptionsProgramLanguagesDescription),"/n",vbCrLf)
  .OptionsProgramLanguagesSymbol = Replace$(hLang.Retrieve("ProgramLanguagesSymbol", .OptionsProgramLanguagesSymbol),"/n",vbCrLf)
  .OptionsProgramNoProcessingAtStartup = Replace$(hLang.Retrieve("ProgramNoProcessingAtStartup", .OptionsProgramNoProcessingAtStartup),"/n",vbCrLf)
  .OptionsProgramOptionsDesign = Replace$(hLang.Retrieve("ProgramOptionsDesign", .OptionsProgramOptionsDesign),"/n",vbCrLf)
  .OptionsProgramOptionsDesignGradient = Replace$(hLang.Retrieve("ProgramOptionsDesignGradient", .OptionsProgramOptionsDesignGradient),"/n",vbCrLf)
  .OptionsProgramOptionsDesignSimple = Replace$(hLang.Retrieve("ProgramOptionsDesignSimple", .OptionsProgramOptionsDesignSimple),"/n",vbCrLf)
  .OptionsProgramPrintDescription = Replace$(hLang.Retrieve("ProgramPrintDescription", .OptionsProgramPrintDescription),"/n",vbCrLf)
  .OptionsProgramPrintSymbol = Replace$(hLang.Retrieve("ProgramPrintSymbol", .OptionsProgramPrintSymbol),"/n",vbCrLf)
  .OptionsProgramRunProgramAfterSavingCaption = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingCaption", .OptionsProgramRunProgramAfterSavingCaption),"/n",vbCrLf)
  .OptionsProgramRunProgramAfterSavingProgram = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingProgram", .OptionsProgramRunProgramAfterSavingProgram),"/n",vbCrLf)
  .OptionsProgramRunProgramAfterSavingProgramParameters = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingProgramParameters", .OptionsProgramRunProgramAfterSavingProgramParameters),"/n",vbCrLf)
  .OptionsProgramRunProgramAfterSavingWaitUntilReady = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWaitUntilReady", .OptionsProgramRunProgramAfterSavingWaitUntilReady),"/n",vbCrLf)
  .OptionsProgramRunProgramAfterSavingWindowstyle = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyle", .OptionsProgramRunProgramAfterSavingWindowstyle),"/n",vbCrLf)
  .OptionsProgramRunProgramAfterSavingWindowstyleHide = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleHide", .OptionsProgramRunProgramAfterSavingWindowstyleHide),"/n",vbCrLf)
  .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleMaximizedFocus", .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus),"/n",vbCrLf)
  .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleMinimizedFocus", .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus),"/n",vbCrLf)
  .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus", .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus),"/n",vbCrLf)
  .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleNormalFocus", .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus),"/n",vbCrLf)
  .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramAfterSavingWindowstyleNormalNoFocus", .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus),"/n",vbCrLf)
  .OptionsProgramRunProgramBeforeSavingCaption = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingCaption", .OptionsProgramRunProgramBeforeSavingCaption),"/n",vbCrLf)
  .OptionsProgramRunProgramBeforeSavingProgram = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingProgram", .OptionsProgramRunProgramBeforeSavingProgram),"/n",vbCrLf)
  .OptionsProgramRunProgramBeforeSavingProgramParameters = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingProgramParameters", .OptionsProgramRunProgramBeforeSavingProgramParameters),"/n",vbCrLf)
  .OptionsProgramRunProgramBeforeSavingWaitUntilReady = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWaitUntilReady", .OptionsProgramRunProgramBeforeSavingWaitUntilReady),"/n",vbCrLf)
  .OptionsProgramRunProgramBeforeSavingWindowstyle = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyle", .OptionsProgramRunProgramBeforeSavingWindowstyle),"/n",vbCrLf)
  .OptionsProgramRunProgramBeforeSavingWindowstyleHide = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleHide", .OptionsProgramRunProgramBeforeSavingWindowstyleHide),"/n",vbCrLf)
  .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleMaximizedFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus),"/n",vbCrLf)
  .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleMinimizedFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus),"/n",vbCrLf)
  .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus),"/n",vbCrLf)
  .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleNormalFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus),"/n",vbCrLf)
  .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus = Replace$(hLang.Retrieve("ProgramRunProgramBeforeSavingWindowstyleNormalNoFocus", .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus),"/n",vbCrLf)
  .OptionsProgramSaveDescription = Replace$(hLang.Retrieve("ProgramSaveDescription", .OptionsProgramSaveDescription),"/n",vbCrLf)
  .OptionsProgramSaveSymbol = Replace$(hLang.Retrieve("ProgramSaveSymbol", .OptionsProgramSaveSymbol),"/n",vbCrLf)
  .OptionsProgramShowAnimation = Replace$(hLang.Retrieve("ProgramShowAnimation", .OptionsProgramShowAnimation),"/n",vbCrLf)
  .OptionsProgramSwitchingDefaultprinter = Replace$(hLang.Retrieve("ProgramSwitchingDefaultprinter", .OptionsProgramSwitchingDefaultprinter),"/n",vbCrLf)
  .OptionsPSDColorsCount01 = Replace$(hLang.Retrieve("PSDColorsCount01", .OptionsPSDColorsCount01),"/n",vbCrLf)
  .OptionsPSDColorscount02 = Replace$(hLang.Retrieve("PSDColorscount02", .OptionsPSDColorscount02),"/n",vbCrLf)
  .OptionsPSDDescription = Replace$(hLang.Retrieve("PSDDescription", .OptionsPSDDescription),"/n",vbCrLf)
  .OptionsPSDescription = Replace$(hLang.Retrieve("PSDescription", .OptionsPSDescription),"/n",vbCrLf)
  .OptionsPSDSymbol = Replace$(hLang.Retrieve("PSDSymbol", .OptionsPSDSymbol),"/n",vbCrLf)
  .OptionsPSFiles = Replace$(hLang.Retrieve("PSFiles", .OptionsPSFiles),"/n",vbCrLf)
  .OptionsPSLanguageLevel = Replace$(hLang.Retrieve("PSLanguageLevel", .OptionsPSLanguageLevel),"/n",vbCrLf)
  .OptionsPSSymbol = Replace$(hLang.Retrieve("PSSymbol", .OptionsPSSymbol),"/n",vbCrLf)
  .OptionsRAWColorsCount01 = Replace$(hLang.Retrieve("RAWColorsCount01", .OptionsRAWColorsCount01),"/n",vbCrLf)
  .OptionsRAWColorscount02 = Replace$(hLang.Retrieve("RAWColorscount02", .OptionsRAWColorscount02),"/n",vbCrLf)
  .OptionsRAWColorscount03 = Replace$(hLang.Retrieve("RAWColorscount03", .OptionsRAWColorscount03),"/n",vbCrLf)
  .OptionsRAWDescription = Replace$(hLang.Retrieve("RAWDescription", .OptionsRAWDescription),"/n",vbCrLf)
  .OptionsRAWSymbol = Replace$(hLang.Retrieve("RAWSymbol", .OptionsRAWSymbol),"/n",vbCrLf)
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
  .OptionsSendEmailAfterAutosave = Replace$(hLang.Retrieve("SendEmailAfterAutosave", .OptionsSendEmailAfterAutosave),"/n",vbCrLf)
  .OptionsSendMailMethod = Replace$(hLang.Retrieve("SendMailMethod", .OptionsSendMailMethod),"/n",vbCrLf)
  .OptionsSendMailMethodAutomatic = Replace$(hLang.Retrieve("SendMailMethodAutomatic", .OptionsSendMailMethodAutomatic),"/n",vbCrLf)
  .OptionsSendMailMethodMapi = Replace$(hLang.Retrieve("SendMailMethodMapi", .OptionsSendMailMethodMapi),"/n",vbCrLf)
  .OptionsSendMailMethodSendmailDLL = Replace$(hLang.Retrieve("SendMailMethodSendmailDLL", .OptionsSendMailMethodSendmailDLL),"/n",vbCrLf)
  .OptionsShellIntegration = Replace$(hLang.Retrieve("ShellIntegration", .OptionsShellIntegration),"/n",vbCrLf)
  .OptionsShellIntegrationAdd = Replace$(hLang.Retrieve("ShellIntegrationAdd", .OptionsShellIntegrationAdd),"/n",vbCrLf)
  .OptionsShellIntegrationCaption = Replace$(hLang.Retrieve("ShellIntegrationCaption", .OptionsShellIntegrationCaption),"/n",vbCrLf)
  .OptionsShellIntegrationRemove = Replace$(hLang.Retrieve("ShellIntegrationRemove", .OptionsShellIntegrationRemove),"/n",vbCrLf)
  .OptionsStamp = Replace$(hLang.Retrieve("Stamp", .OptionsStamp),"/n",vbCrLf)
  .OptionsStampFontColor = Replace$(hLang.Retrieve("StampFontColor", .OptionsStampFontColor),"/n",vbCrLf)
  .OptionsStampOutlineFontThickness = Replace$(hLang.Retrieve("StampOutlineFontThickness", .OptionsStampOutlineFontThickness),"/n",vbCrLf)
  .OptionsStampString = Replace$(hLang.Retrieve("StampString", .OptionsStampString),"/n",vbCrLf)
  .OptionsStampUseOutlineFont = Replace$(hLang.Retrieve("StampUseOutlineFont", .OptionsStampUseOutlineFont),"/n",vbCrLf)
  .OptionsStandardAuthorToken = Replace$(hLang.Retrieve("StandardAuthorToken", .OptionsStandardAuthorToken),"/n",vbCrLf)
  .OptionsStandardSaveFormat = Replace$(hLang.Retrieve("StandardSaveFormat", .OptionsStandardSaveFormat),"/n",vbCrLf)
  .OptionsSVGDescription = Replace$(hLang.Retrieve("SVGDescription", .OptionsSVGDescription),"/n",vbCrLf)
  .OptionsSVGSymbol = Replace$(hLang.Retrieve("SVGSymbol", .OptionsSVGSymbol),"/n",vbCrLf)
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
  .OptionsToolbar = Replace$(hLang.Retrieve("Toolbar", .OptionsToolbar),"/n",vbCrLf)
  .OptionsToolbarInstall = Replace$(hLang.Retrieve("ToolbarInstall", .OptionsToolbarInstall),"/n",vbCrLf)
  .OptionsTreeFormats = Replace$(hLang.Retrieve("TreeFormats", .OptionsTreeFormats),"/n",vbCrLf)
  .OptionsTreeProgram = Replace$(hLang.Retrieve("TreeProgram", .OptionsTreeProgram),"/n",vbCrLf)
  .OptionsTXTDescription = Replace$(hLang.Retrieve("TXTDescription", .OptionsTXTDescription),"/n",vbCrLf)
  .OptionsTXTSymbol = Replace$(hLang.Retrieve("TXTSymbol", .OptionsTXTSymbol),"/n",vbCrLf)
  .OptionsUseAutosave = Replace$(hLang.Retrieve("UseAutosave", .OptionsUseAutosave),"/n",vbCrLf)
  .OptionsUseAutosaveDirectory = Replace$(hLang.Retrieve("UseAutosaveDirectory", .OptionsUseAutosaveDirectory),"/n",vbCrLf)
  .OptionsUseCreationDateNow = Replace$(hLang.Retrieve("UseCreationDateNow", .OptionsUseCreationDateNow),"/n",vbCrLf)
  .OptionsUseCustomPapersize = Replace$(hLang.Retrieve("UseCustomPapersize", .OptionsUseCustomPapersize),"/n",vbCrLf)
  .OptionsUseFixPapersize = Replace$(hLang.Retrieve("UseFixPapersize", .OptionsUseFixPapersize),"/n",vbCrLf)
  .OptionsUserPass = Replace$(hLang.Retrieve("UserPass", .OptionsUserPass),"/n",vbCrLf)
  .OptionsUseStandardauthor = Replace$(hLang.Retrieve("UseStandardauthor", .OptionsUseStandardauthor),"/n",vbCrLf)
  .OptionsXCFColorsCount01 = Replace$(hLang.Retrieve("XCFColorsCount01", .OptionsXCFColorsCount01),"/n",vbCrLf)
  .OptionsXCFColorscount02 = Replace$(hLang.Retrieve("XCFColorscount02", .OptionsXCFColorscount02),"/n",vbCrLf)
  .OptionsXCFDescription = Replace$(hLang.Retrieve("XCFDescription", .OptionsXCFDescription),"/n",vbCrLf)
  .OptionsXCFSymbol = Replace$(hLang.Retrieve("XCFSymbol", .OptionsXCFSymbol),"/n",vbCrLf)
 End With
 Set hLang = Nothing
End Sub

Private Sub LoadPrintersStrings(ByVal Languagefile As String)
 Dim hLang As New clsHash
 ReadINISection Languagefile, "Printers", hLang
 With LanguageStrings
  .PrintersAdminNotice = Replace$(hLang.Retrieve("AdminNotice", .PrintersAdminNotice),"/n",vbCrLf)
  .PrintersClose = Replace$(hLang.Retrieve("Close", .PrintersClose),"/n",vbCrLf)
  .PrintersNewPrinterName = Replace$(hLang.Retrieve("NewPrinterName", .PrintersNewPrinterName),"/n",vbCrLf)
  .PrintersPrinter = Replace$(hLang.Retrieve("Printer", .PrintersPrinter),"/n",vbCrLf)
  .PrintersPrinterAdd = Replace$(hLang.Retrieve("PrinterAdd", .PrintersPrinterAdd),"/n",vbCrLf)
  .PrintersPrinterDel = Replace$(hLang.Retrieve("PrinterDel", .PrintersPrinterDel),"/n",vbCrLf)
  .PrintersPrinters = Replace$(hLang.Retrieve("Printers", .PrintersPrinters),"/n",vbCrLf)
  .PrintersProfile = Replace$(hLang.Retrieve("Profile", .PrintersProfile),"/n",vbCrLf)
  .PrintersSave = Replace$(hLang.Retrieve("Save", .PrintersSave),"/n",vbCrLf)
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
  .PrintingPCLFiles = Replace$(hLang.Retrieve("PCLFiles", .PrintingPCLFiles),"/n",vbCrLf)
  .PrintingPCXFiles = Replace$(hLang.Retrieve("PCXFiles", .PrintingPCXFiles),"/n",vbCrLf)
  .PrintingPDFAFiles = Replace$(hLang.Retrieve("PDFAFiles", .PrintingPDFAFiles),"/n",vbCrLf)
  .PrintingPDFFiles = Replace$(hLang.Retrieve("PDFFiles", .PrintingPDFFiles),"/n",vbCrLf)
  .PrintingPDFXFiles = Replace$(hLang.Retrieve("PDFXFiles", .PrintingPDFXFiles),"/n",vbCrLf)
  .PrintingPNGFiles = Replace$(hLang.Retrieve("PNGFiles", .PrintingPNGFiles),"/n",vbCrLf)
  .PrintingProfile = Replace$(hLang.Retrieve("Profile", .PrintingProfile),"/n",vbCrLf)
  .PrintingPSDFiles = Replace$(hLang.Retrieve("PSDFiles", .PrintingPSDFiles),"/n",vbCrLf)
  .PrintingPSFiles = Replace$(hLang.Retrieve("PSFiles", .PrintingPSFiles),"/n",vbCrLf)
  .PrintingRAWFiles = Replace$(hLang.Retrieve("RAWFiles", .PrintingRAWFiles),"/n",vbCrLf)
  .PrintingSave = Replace$(hLang.Retrieve("Save", .PrintingSave),"/n",vbCrLf)
  .PrintingStartStandardProgram = Replace$(hLang.Retrieve("StartStandardProgram", .PrintingStartStandardProgram),"/n",vbCrLf)
  .PrintingStatus = Replace$(hLang.Retrieve("Status", .PrintingStatus),"/n",vbCrLf)
  .PrintingSubject = Replace$(hLang.Retrieve("Subject", .PrintingSubject),"/n",vbCrLf)
  .PrintingSVGFiles = Replace$(hLang.Retrieve("SVGFiles", .PrintingSVGFiles),"/n",vbCrLf)
  .PrintingTIFFFiles = Replace$(hLang.Retrieve("TIFFFiles", .PrintingTIFFFiles),"/n",vbCrLf)
  .PrintingTXTFiles = Replace$(hLang.Retrieve("TXTFiles", .PrintingTXTFiles),"/n",vbCrLf)
  .PrintingXCFFiles = Replace$(hLang.Retrieve("XCFFiles", .PrintingXCFFiles),"/n",vbCrLf)
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
  .CommonVersion = "1.1.0"

  .DialogDocument = "&Document"
  .DialogDocumentAdd = "Add"
  .DialogDocumentAddFromClipboard = "Add from clipboard"
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
  .DialogInfoTitle = "About"
  .DialogLanguage = "&Language"
  .DialogPrinter = "&Printer"
  .DialogPrinterClose = "Close"
  .DialogPrinterLogfile = "Logfile"
  .DialogPrinterLogfiles = "Logfiles"
  .DialogPrinterLogging = "Logging"
  .DialogPrinterOptions = "Options"
  .DialogPrinterPrinters = "Printers"
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
  .MessagesMsg31 = "An error has occurred"
  .MessagesMsg32 = "The new version %1 is available. Would you like download the new version from the Sourceforge pages?"
  .MessagesMsg33 = "You already have the most recent version."
  .MessagesMsg34 = "The file is in use. Please close the file first or choose another filename."
  .MessagesMsg35 = "It is necessary to temporarily set PDFCreator as defaultprinter."
  .MessagesMsg36 = "Don't ask me again."
  .MessagesMsg37 = "The downloaded file is not a valid language file!"
  .MessagesMsg38 = "The language file has been successfully installed!"
  .MessagesMsg39 = "pdfforge.dll is not installed! You can find more information in the help file."
  .MessagesMsg40 = "A printer with this name is installed already!"
  .MessagesMsg41 = "A profile with this name exists already!"
  .MessagesMsg42 = "Do you want delete the profile: '%1'?"
  .MessagesMsg43 = "You can't delete this profile because it is associate with at least one printer."
  .MessagesMsg44 = "Could not connect to the update server."

  .OptionsAdditionalGhostscriptParameters = "Additional Ghostscript parameters"
  .OptionsAdditionalGhostscriptSearchpath = "Additional Ghostscript searchpath"
  .OptionsAddWindowsFontpath = "Use Windows fonts"
  .OptionsAllowSpecialGSCharsInFilenames = "Allow special Ghostscript chars in filename"
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
  .OptionsBMPColorscount05_2 = "Separated 8-bit CMYK"
  .OptionsBMPColorscount06_2 = "Separated 1-bit CMYK"
  .OptionsBMPColorscount07 = "Greyscale (8 Bit)"
  .OptionsBMPColorscount08 = "Monochrome"
  .OptionsBMPDescription = "Windows Bitmap Format. Please use only for single pages."
  .OptionsBMPSymbol = "BMP"
  .OptionsCancel = "&Cancel"
  .OptionsCheckUpdateDescription = "Check update"
  .OptionsCheckUpdateInterval = "Update interval"
  .OptionsCheckUpdateInterval01 = "Never"
  .OptionsCheckUpdateInterval02 = "Once a day"
  .OptionsCheckUpdateInterval03 = "Once a week"
  .OptionsCheckUpdateInterval04 = "Once a month"
  .OptionsCheckUpdateNow = "Check now"
  .OptionsCustomPapersizeHeight = "Height"
  .OptionsCustomPapersizeInfo = "Units of 1/72 of an inch."
  .OptionsCustomPapersizeWidth = "Width"
  .OptionsDirectoriesGSBin = "Ghostscript Binaries"
  .OptionsDirectoriesGSFonts = "Ghostscript Fonts"
  .OptionsDirectoriesGSLibraries = "Ghostscript Libraries"
  .OptionsDirectoriesTempPath = "Temporary Files"
  .OptionsDocument = "Document"
  .OptionsEnableNotice = "You can set these options in the default profile only."
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
  .OptionsLanguagesCurrentLanguage = "Current language"
  .OptionsLanguagesDownloadMoreLanguages = "Load more languages from the internet"
  .OptionsLanguagesInstall = "Install"
  .OptionsLanguagesRefresh = "Refresh List"
  .OptionsLanguagesTranslation = "Translation"
  .OptionsLanguagesVersion = "Version"
  .OptionsNothingToConfigure = "There is nothing to configure."
  .OptionsOnePagePerFile = "One page per file (not for pdf and eps files)"
  .OptionsOwnerPass = "Owner Password"
  .OptionsPassCancel = "Cancel"
  .OptionsPassOK = "OK"
  .OptionsPCLColorscount01 = "16777216 colors (24bit)"
  .OptionsPCLColorscount02 = "2 colors (Black/White)"
  .OptionsPCLDescription = "HP PCL-XL Format"
  .OptionsPCLSymbol = "PCL"
  .OptionsPCXColorscount01 = "4294967296 colors (32 Bit) CMYK"
  .OptionsPCXColorscount02 = "16777216 colors (24 Bit)"
  .OptionsPCXColorscount03 = "256 colors (8 Bit)"
  .OptionsPCXColorscount04 = "16 colors (4 Bit)"
  .OptionsPCXColorscount05 = "2 colors (Black\White)"
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
  .OptionsPDFCompressionColorComp09 = "JPEG-Manual"
  .OptionsPDFCompressionColorCompFac = "Factor"
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
  .OptionsPDFCompressionGreyComp09 = "JPEG-Manual"
  .OptionsPDFCompressionGreyCompFac = "Factor"
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
  .OptionsPDFEncryptionAes128 = "Very high (AES 128 Bit - Adobe Acrobat 7.0 and above)"
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
  .OptionsPDFGeneralCompatibility04 = "Adobe Acrobat 6.0 (PDF 1.5)"
  .OptionsPDFGeneralDefaultSettings = "Default settings"
  .OptionsPDFGeneralDefaultSettingsDefault = "Default"
  .OptionsPDFGeneralDefaultSettingsEbook = "Ebook"
  .OptionsPDFGeneralDefaultSettingsPrepress = "Pre-press"
  .OptionsPDFGeneralDefaultSettingsPrinter = "Printer"
  .OptionsPDFGeneralDefaultSettingsScreen = "Screen"
  .OptionsPDFGeneralOverprint = "Overprint:"
  .OptionsPDFGeneralOverprint01 = "Non-Zero Overprint"
  .OptionsPDFGeneralOverprint02 = "Full Overprint"
  .OptionsPDFGeneralResolution = "Resolution:"
  .OptionsPDFGeneralRotate01 = "None"
  .OptionsPDFGeneralRotate02 = "All"
  .OptionsPDFGeneralRotate03 = "Single Page"
  .OptionsPDFOptimize = "Fast web view"
  .OptionsPDFOptions = "PDF Options"
  .OptionsPDFOwnerPass = "Password required to change permissions and passwords"
  .OptionsPDFOwnerPasswordShowChars = "Show password"
  .OptionsPDFPasswords = "Passwords"
  .OptionsPDFRepeatPassword = "Repeat"
  .OptionsPDFSecurity = "Security"
  .OptionsPDFSecurityCaption = "Security"
  .OptionsPDFSetPassword = "Password"
  .OptionsPDFSigning = "Signing"
  .OptionsPDFSigningCaption = "Signing of PDFs"
  .OptionsPDFSigningCerticatePassword = "Certificate password"
  .OptionsPDFSigningCerticatePasswordCancel = "&Cancel"
  .OptionsPDFSigningCerticatePasswordOk = "&Ok"
  .OptionsPDFSigningCerticatePasswordShowPassword = "Show password"
  .OptionsPDFSigningCertificateEmptyPassword = "No password is entered. The pdf file will not be signed."
  .OptionsPDFSigningCertificateFile = "Certificate file"
  .OptionsPDFSigningChooseCertifcateFile = "Choose a certificate"
  .OptionsPDFSigningEnterCerticatePassword = "Enter certificate password"
  .OptionsPDFSigningP12Files = "P12 files"
  .OptionsPDFSigningPfxFiles = "Pfx files"
  .OptionsPDFSigningPfxP12Files = "Pfx/P12 files"
  .OptionsPDFSigningSignatureContact = "Signature contact"
  .OptionsPDFSigningSignatureLocation = "Signature location"
  .OptionsPDFSigningSignatureMultiSignature = "Multi signature allowed"
  .OptionsPDFSigningSignatureOnPage = "Show signature on page"
  .OptionsPDFSigningSignaturePosition = "Signature position"
  .OptionsPDFSigningSignaturePositionLeftX = "LeftX"
  .OptionsPDFSigningSignaturePositionLeftY = "LeftY"
  .OptionsPDFSigningSignaturePositionRightX = "RightX"
  .OptionsPDFSigningSignaturePositionRightY = "RightY"
  .OptionsPDFSigningSignatureReason = "Signature reason"
  .OptionsPDFSigningSignatureVisible = "Signature visible in pdf file"
  .OptionsPDFSigningSignPdfFile = "Sign pdf file"
  .OptionsPDFSymbol = "PDF"
  .OptionsPDFUserPass = "Password required to open document"
  .OptionsPDFUserPasswordShowChars = "Show password"
  .OptionsPDFUseSecurity = "Use Security"
  .OptionsPNGColorscount01 = "16777216 colors (24 Bit)"
  .OptionsPNGColorscount02 = "256 colors (8 Bit)"
  .OptionsPNGColorscount03 = "16 colors (4 Bit)"
  .OptionsPNGColorscount04 = "2 colors (2 Bit - Black/White)"
  .OptionsPNGColorscount05 = "Greyscale (8 Bit)"
  .OptionsPNGDescription = "PNG Format. Please use only for single pages."
  .OptionsPNGFiles = "Bitmap PNG-Files"
  .OptionsPNGSymbol = "PNG"
  .OptionsPrintAfterSaving = "Print after saving"
  .OptionsPrintAfterSavingBitsPerPixel = "Bits per pixel"
  .OptionsPrintAfterSavingBitsPerPixelCMYK = "CMYK"
  .OptionsPrintAfterSavingBitsPerPixelMono = "Mono"
  .OptionsPrintAfterSavingBitsPerPixelTrueColor = "True Color"
  .OptionsPrintAfterSavingDuplex = "Duplex"
  .OptionsPrintAfterSavingDuplexTumbleOff = "Don't use tumble (Default)"
  .OptionsPrintAfterSavingDuplexTumbleOn = "Use tumble"
  .OptionsPrintAfterSavingMaxResolution = "Set maximum print resolution"
  .OptionsPrintAfterSavingNoCancel = "Hide the progress dialog during printing"
  .OptionsPrintAfterSavingPrinter = "Printer"
  .OptionsPrintAfterSavingQueryUser = "Query user"
  .OptionsPrintAfterSavingQueryUserDefaultPrinter = "Select the default Windows printer without any user interaction"
  .OptionsPrintAfterSavingQueryUserOff = "Off (Default)"
  .OptionsPrintAfterSavingQueryUserPrinterSetupDialog = "Shows the printer setup dialog"
  .OptionsPrintAfterSavingQueryUserStandardPrinterDialog = "Show the standard printer dialog"
  .OptionsPrintertempDirectoryPrompt = "Select Printer Temp-Directory"
  .OptionsPrintTestpage = "Print Test Page"
  .OptionsProcesspriority = "Process priority"
  .OptionsProcesspriorityHigh = "High"
  .OptionsProcesspriorityIdle = "Idle"
  .OptionsProcesspriorityNormal = "Normal"
  .OptionsProcesspriorityRealtime = "Realtime"
  .OptionsProfile = "Profile"
  .OptionsProfileAdd = "Add profile"
  .OptionsProfileCancel = "&Cancel"
  .OptionsProfileDefaultName = "Default"
  .OptionsProfileDel = "Delete profile"
  .OptionsProfileLoadFromDisc = "Load profile from disc"
  .OptionsProfileNewProfile = "New profile"
  .OptionsProfileOk = "&Ok"
  .OptionsProfileRenameProfile = "Rename profile"
  .OptionsProfileSaveToDisc = "Save profile to disc"
  .OptionsProgramActionsDescription = "Define an action before and after saving a file."
  .OptionsProgramActionsSymbol = "Actions"
  .OptionsProgramAutosaveDescription = "Auto-save mode. Auto-save does not prompt for a filename and file location. It automatically saves all PDF files to a single directory with a predefined filename."
  .OptionsProgramAutosaveSymbol = "Auto-save"
  .OptionsProgramDirectoriesDescription = "Directories for Ghostscript, temporary files and others."
  .OptionsProgramDirectoriesSymbol = "Directories"
  .OptionsProgramDocumentDescription = "Document properties"
  .OptionsProgramDocumentDescription1 = "Document properties 1"
  .OptionsProgramDocumentDescription2 = "Document properties 2"
  .OptionsProgramDocumentSymbol = "Document"
  .OptionsProgramFont = "Program Font"
  .OptionsProgramFontCancelTest = "Cancel Test"
  .OptionsProgramFontcharset = "Character Set"
  .OptionsProgramFontDescription = "Font for labels, captions and values. For the program menu use the general settings in your Windows OS."
  .OptionsProgramFontSize = "Size"
  .OptionsProgramFontSymbol = "Program font"
  .OptionsProgramFontTest = "Test"
  .OptionsProgramFontTestdescription = "Here you can test the font."
  .OptionsProgramGeneralDescription = "General Settings"
  .OptionsProgramGeneralDescription1 = "General Settings 1"
  .OptionsProgramGeneralDescription2 = "General Settings 2"
  .OptionsProgramGeneralSymbol = "General settings"
  .OptionsProgramGhostscriptDescription = "Ghostscript"
  .OptionsProgramGhostscriptSymbol = "Ghostscript"
  .OptionsProgramLanguagesDescription = "Define the language and download another languages from the internet."
  .OptionsProgramLanguagesSymbol = "Languages"
  .OptionsProgramNoProcessingAtStartup = "No processing at startup"
  .OptionsProgramOptionsDesign = "Frame color of the options dialog"
  .OptionsProgramOptionsDesignGradient = "Red and blue gradient (Default)"
  .OptionsProgramOptionsDesignSimple = "Simple red and blue color"
  .OptionsProgramPrintDescription = "Print after saving"
  .OptionsProgramPrintSymbol = "Print"
  .OptionsProgramRunProgramAfterSavingCaption = "Action after saving"
  .OptionsProgramRunProgramAfterSavingProgram = "Program/Script"
  .OptionsProgramRunProgramAfterSavingProgramParameters = "Program parameters"
  .OptionsProgramRunProgramAfterSavingWaitUntilReady = "Wait until the program/script is ready"
  .OptionsProgramRunProgramAfterSavingWindowstyle = "Window style"
  .OptionsProgramRunProgramAfterSavingWindowstyleHide = "Hide"
  .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus = "Maximized/Focus"
  .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus = "Minimized/Focus"
  .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus = "Minimized/No focus"
  .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus = "Normal/Focus"
  .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus = "Normal/No focus"
  .OptionsProgramRunProgramBeforeSavingCaption = "Action before saving"
  .OptionsProgramRunProgramBeforeSavingProgram = "Program/Script"
  .OptionsProgramRunProgramBeforeSavingProgramParameters = "Program parameters"
  .OptionsProgramRunProgramBeforeSavingWaitUntilReady = "Wait until the program/script is ready"
  .OptionsProgramRunProgramBeforeSavingWindowstyle = "Window style"
  .OptionsProgramRunProgramBeforeSavingWindowstyleHide = "Hide"
  .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus = "Maximized/Focus"
  .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus = "Minimized/Focus"
  .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus = "Minimized/NoFocus"
  .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus = "Normal/Focus"
  .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus = "Normal/NoFocus"
  .OptionsProgramSaveDescription = "Save"
  .OptionsProgramSaveSymbol = "Save"
  .OptionsProgramShowAnimation = "Show an animation during the process"
  .OptionsProgramSwitchingDefaultprinter = "No confirm message switching PDFCreator temporarily as default printer."
  .OptionsPSDColorsCount01 = "4294967296 colors (32 Bit) CMYK"
  .OptionsPSDColorscount02 = "16777216 colors (24 Bit)"
  .OptionsPSDDescription = "Photoshop Format"
  .OptionsPSDescription = "Postscript Format"
  .OptionsPSDSymbol = "PSD"
  .OptionsPSFiles = "Postscript-Files"
  .OptionsPSLanguageLevel = "Language Level:"
  .OptionsPSSymbol = "PS"
  .OptionsRAWColorsCount01 = "4294967296 colors (32 Bit) CMYK"
  .OptionsRAWColorscount02 = "16777216 colors (24 Bit)"
  .OptionsRAWColorscount03 = "2 colors (Black/White)"
  .OptionsRAWDescription = "Raw Format"
  .OptionsRAWSymbol = "Raw"
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
  .OptionsSendEmailAfterAutosave = "Send an email after auto-saving"
  .OptionsSendMailMethod = "Method to send an email"
  .OptionsSendMailMethodAutomatic = "Automatic"
  .OptionsSendMailMethodMapi = "Mapi interface"
  .OptionsSendMailMethodSendmailDLL = "Using sendmail.dll"
  .OptionsShellIntegration = "Shell integration"
  .OptionsShellIntegrationAdd = "Integrate PDFCreator into shell"
  .OptionsShellIntegrationCaption = "Create &PDF with PDFCreator"
  .OptionsShellIntegrationRemove = "Remove shell integration"
  .OptionsStamp = "Stamp"
  .OptionsStampFontColor = "Font-color"
  .OptionsStampOutlineFontThickness = "Outline font thickness"
  .OptionsStampString = "Stampstring"
  .OptionsStampUseOutlineFont = "Use outline font"
  .OptionsStandardAuthorToken = "Add a Author-Token"
  .OptionsStandardSaveFormat = "Standard save format"
  .OptionsSVGDescription = "SVG Format"
  .OptionsSVGSymbol = "SVG"
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
  .OptionsToolbar = "Toolbar"
  .OptionsToolbarInstall = "Install pdfforge Toolbar"
  .OptionsTreeFormats = "Formats"
  .OptionsTreeProgram = "Program"
  .OptionsTXTDescription = "Text Format"
  .OptionsTXTSymbol = "TXT"
  .OptionsUseAutosave = "Use Auto-save"
  .OptionsUseAutosaveDirectory = "Use this directory for auto-save"
  .OptionsUseCreationDateNow = "Use the current Date/Time for 'Creation Date'"
  .OptionsUseCustomPapersize = "Use custom paper size"
  .OptionsUseFixPapersize = "Use fixed paper size"
  .OptionsUserPass = "User Password"
  .OptionsUseStandardauthor = "Use standard author"
  .OptionsXCFColorsCount01 = "4294967296 colors (32 Bit) CMYK"
  .OptionsXCFColorscount02 = "16777216 colors (24 Bit)"
  .OptionsXCFDescription = "Gimp Format"
  .OptionsXCFSymbol = "XCF"

  .PrintersAdminNotice = "You must be an administrator to install or delete a printer!"
  .PrintersClose = "Close"
  .PrintersNewPrinterName = "New printer name"
  .PrintersPrinter = "Printer"
  .PrintersPrinterAdd = "Add printer"
  .PrintersPrinterDel = "Del printer"
  .PrintersPrinters = "Printers"
  .PrintersProfile = "Profile"
  .PrintersSave = "Save"

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
  .PrintingPCLFiles = "PCL (HP PCL-XL)-Files"
  .PrintingPCXFiles = "PCX-Files"
  .PrintingPDFAFiles = "PDF/A-1b-Files"
  .PrintingPDFFiles = "PDF-Files"
  .PrintingPDFXFiles = "PDF/X-Files"
  .PrintingPNGFiles = "PNG-Files"
  .PrintingProfile = "Profile"
  .PrintingPSDFiles = "PSD (Adobe Photoshop)-Files"
  .PrintingPSFiles = "Postscript-Files"
  .PrintingRAWFiles = "RAW (binary format)-Files"
  .PrintingSave = "&Save"
  .PrintingStartStandardProgram = "&After saving open the document with the default program."
  .PrintingStatus = "Creating file..."
  .PrintingSubject = "Su&bject:"
  .PrintingSVGFiles = "SVG-Files"
  .PrintingTIFFFiles = "TIFF-Files"
  .PrintingTXTFiles = "Text-Files"
  .PrintingXCFFiles = "XCF (Gimp)-Files"

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

