Attribute VB_Name = "modOptions2"
Option Explicit

' Automatically generated with DeveloperTool by Frank Heindörfer
' 2003 - 2007
' Email: thesmilyface@users.sourceforge.net

Public Sub ShowOptions(Frm As Form, sOptions As tOptions)
 On Error Resume Next
 Dim i As Long, tList() As String, tStrA() As String, lsv As ListView
 With sOptions
  Frm.cmbAdditionalGhostscriptParameters.Text = .AdditionalGhostscriptParameters
  Frm.txtAdditionalGhostscriptSearchpath.Text = .AdditionalGhostscriptSearchpath
  Frm.chkAddWindowsFontpath.Value = .AddWindowsFontpath
  Frm.txtAutosaveDirectory.Text = .AutosaveDirectory
  Frm.txtAutosaveFilename.Text = .AutosaveFilename
  Frm.cmbAutosaveFormat.ListIndex = .AutosaveFormat
  Frm.chkAutosaveStartStandardProgram.Value = .AutosaveStartStandardProgram
  Frm.txtBitmapResolution.Text = .BitmapResolution
  Frm.cmbBMPColors.ListIndex = .BMPColorscount
  Frm.txtCustomPapersizeHeight.Text = .DeviceHeightPoints
  Frm.txtCustomPapersizeWidth.Text = .DeviceWidthPoints
  Frm.txtGSbin.Text = .DirectoryGhostscriptBinaries
  Frm.txtGSfonts.Text = .DirectoryGhostscriptFonts
  Frm.txtGSlib.Text = .DirectoryGhostscriptLibraries
  Frm.txtGSResource.Text = .DirectoryGhostscriptResource
  Frm.cmbEPSLanguageLevel.ListIndex = .EPSLanguageLevel
  Set lsv = Frm.lsvFilenameSubst
  lsv.ListItems.Clear
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
  Frm.chkNoConfirmMessageSwitchingDefaultprinter = .NoConfirmMessageSwitchingDefaultprinter
  Frm.chkNoProcessingAtStartup = .NoProcessingAtStartup
  Frm.chkOnePagePerFile.Value = .OnePagePerFile
  Frm.cmbOptionsDesign.ListIndex = .OptionsDesign
  With Frm.cmbDocumentPapersizes
   For i = 0 To .ListCount - 1
    If UCase$(.List(i)) = UCase$(Options.Papersize) Then
     .ListIndex = i
     Exit For
    End If
   Next i
  End With
  Frm.cmbPCXColors.ListIndex = .PCXColorscount
  Frm.chkAllowAssembly.Value = .PDFAllowAssembly
  Frm.chkAllowDegradedPrinting.Value = .PDFAllowDegradedPrinting
  Frm.chkAllowFillIn.Value = .PDFAllowFillIn
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
  Frm.chkAllowCopy.Value = .PDFDisallowCopy
  Frm.chkAllowModifyAnnotations.Value = .PDFDisallowModifyAnnotations
  Frm.chkAllowModifyContents.Value = .PDFDisallowModifyContents
  Frm.chkAllowPrinting.Value = .PDFDisallowPrinting
  Frm.cmbPDFEncryptor.ItemData(Frm.cmbPDFEncryptor.ListIndex) = .PDFEncryptor
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
  Frm.chkPDFOptimize.Value = .PDFOptimize
  Frm.chkOwnerPass.Value = .PDFOwnerPass
  Frm.chkUserPass.Value = .PDFUserPass
  Frm.chkUseSecurity.Value = .PDFUseSecurity
  Frm.cmbPNGColors.ListIndex = .PNGColorscount
  Frm.chkPrintAfterSaving.Value = .PrintAfterSaving
  Frm.chkPrintAfterSavingDuplex.Value = .PrintAfterSavingDuplex
  Frm.chkPrintAfterSavingNoCancel.Value = .PrintAfterSavingNoCancel
  Frm.cmbPrintAfterSavingPrinter.Text = .PrintAfterSavingPrinter
  Frm.cmbPrintAfterSavingQueryUser.ListIndex = .PrintAfterSavingQueryUser
  Frm.cmbPrintAfterSavingTumble.ListIndex = .PrintAfterSavingTumble
  Frm.txtTemppath.Text = .PrinterTemppath
  Frm.sldProcessPriority.Value = .ProcessPriority
  For i = 0 To Frm.cmbFonts.ListCount - 1
    If UCase$(Frm.cmbFonts.List(i)) = UCase$(.ProgramFont) Then
     Frm.cmbFonts.ListIndex = i
     Exit For
    End If
  Next i
  Frm.cmbCharset.Text = .ProgramFontCharset
  Frm.cmbProgramFontSize.Text = .ProgramFontSize
  Frm.cmbPSLanguageLevel.ListIndex = .PSLanguageLevel
  Frm.chkSpaces.Value = .RemoveSpaces
  Frm.chkRunProgramAfterSaving.Value = .RunProgramAfterSaving
  Frm.cmbRunProgramAfterSavingProgramname.Text = .RunProgramAfterSavingProgramname
  Frm.txtRunProgramAfterSavingProgramParameters.Text = .RunProgramAfterSavingProgramParameters
  Frm.chkRunProgramAfterSavingWaitUntilReady.Value = .RunProgramAfterSavingWaitUntilReady
  Frm.cmbRunProgramAfterSavingWindowstyle.ListIndex = .RunProgramAfterSavingWindowstyle
  Frm.chkRunProgramBeforeSaving.Value = .RunProgramBeforeSaving
  Frm.cmbRunProgramBeforeSavingProgramname.Text = .RunProgramBeforeSavingProgramname
  Frm.txtRunProgramBeforeSavingProgramParameters.Text = .RunProgramBeforeSavingProgramParameters
  Frm.cmbRunProgramBeforeSavingWindowstyle.ListIndex = .RunProgramBeforeSavingWindowstyle
  Frm.txtSaveFilename.Text = .SaveFilename
  Frm.chkAutosaveSendEmail.Value = .SendEmailAfterAutoSaving
  Frm.cmbSendMailMethod.ListIndex = .SendMailMethod
  Frm.chkShowAnimation.Value = .ShowAnimation
  Frm.picStampFontColor.BackColor = HTMLColorToOleColor(.StampFontColor)
  Frm.lblFontNameSize.Caption = .StampFontname & ", " & .StampFontsize
  Frm.txtOutlineFontThickness.Text = .StampOutlineFontthickness
  Frm.txtStampString.Text = .StampString
  Frm.chkStampUseOutlineFont.Value = .StampUseOutlineFont
  Frm.txtStandardAuthor.Text = .StandardAuthor
  Frm.cmbStandardSaveformat.ListIndex = .StandardSaveformat
  Frm.cmbTIFFColors.ListIndex = .TIFFColorscount
  Frm.chkUseAutosave.Value = .UseAutosave
  Frm.chkUseAutosaveDirectory.Value = .UseAutosaveDirectory
  Frm.chkUseCreationDateNow.Value = .UseCreationDateNow
  Frm.chkUseCustomPapersize.Value = .UseCustomPaperSize
  Frm.chkUseFixPaperSize.Value = .UseFixPapersize
  Frm.chkUseStandardAuthor.Value = .UseStandardAuthor
 End With
End Sub

Public Sub GetOptions(Frm As Form, sOptions As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, tStr As String, lsv As ListView
50020  sOptions = StandardOptions
50030  With sOptions
50040  .AdditionalGhostscriptParameters = Frm.cmbAdditionalGhostscriptParameters.Text
50050  .AdditionalGhostscriptSearchpath = Frm.txtAdditionalGhostscriptSearchpath.Text
50060  .AddWindowsFontpath = Abs(Frm.chkAddWindowsFontpath.Value)
50070  .AutosaveDirectory = Frm.txtAutosaveDirectory.Text
50080  .AutosaveFilename = Frm.txtAutosaveFilename.Text
50090  If LenB(Frm.cmbAutosaveFormat.ListIndex) > 0 Then
50100   .AutosaveFormat = Frm.cmbAutosaveFormat.ListIndex
50110  End If
50120  .AutosaveStartStandardProgram = Abs(Frm.chkAutosaveStartStandardProgram.Value)
50130  If LenB(Frm.txtBitmapResolution.Text) > 0 Then
50140   .BitmapResolution = Frm.txtBitmapResolution.Text
50150  End If
50160  If LenB(Frm.cmbBMPColors.ListIndex) > 0 Then
50170   .BMPColorscount = Frm.cmbBMPColors.ListIndex
50180  End If
50190  If LenB(Frm.txtCustomPapersizeHeight.Text) > 0 Then
50200   .DeviceHeightPoints = Frm.txtCustomPapersizeHeight.Text
50210  End If
50220  If LenB(Frm.txtCustomPapersizeWidth.Text) > 0 Then
50230   .DeviceWidthPoints = Frm.txtCustomPapersizeWidth.Text
50240  End If
50250  .DirectoryGhostscriptBinaries = Frm.txtGSbin.Text
50260  .DirectoryGhostscriptFonts = Frm.txtGSfonts.Text
50270  .DirectoryGhostscriptLibraries = Frm.txtGSlib.Text
50280  .DirectoryGhostscriptResource = Frm.txtGSResource.Text
50290  If LenB(Frm.cmbEPSLanguageLevel.ListIndex) > 0 Then
50300   .EPSLanguageLevel = Frm.cmbEPSLanguageLevel.ListIndex
50310  End If
50320  tStr = ""
50330  Set lsv = Frm.lsvFilenameSubst
50340  For i = 1 To lsv.ListItems.Count
50350   If i < lsv.ListItems.Count Then
50360     tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1) & "\"
50370    Else
50380     tStr = tStr & lsv.ListItems(i).Text & "|" & lsv.ListItems(i).SubItems(1)
50390   End If
50400  Next i
50410  .FilenameSubstitutions = tStr
50420  .FilenameSubstitutionsOnlyInTitle = Abs(Frm.chkFilenameSubst.Value)
50430  If LenB(Frm.cmbJPEGColors.ListIndex) > 0 Then
50440   .JPEGColorscount = Frm.cmbJPEGColors.ListIndex
50450  End If
50460  If LenB(Frm.txtJPEGQuality.Text) > 0 Then
50470   .JPEGQuality = Frm.txtJPEGQuality.Text
50480  End If
50490  .NoConfirmMessageSwitchingDefaultprinter = Abs(Frm.chkNoConfirmMessageSwitchingDefaultprinter)
50500  .NoProcessingAtStartup = Abs(Frm.chkNoProcessingAtStartup)
50510  .OnePagePerFile = Abs(Frm.chkOnePagePerFile.Value)
50520  If LenB(Frm.cmbOptionsDesign.ListIndex) > 0 Then
50530   .OptionsDesign = Frm.cmbOptionsDesign.ListIndex
50540  End If
50550  If Frm.cmbDocumentPapersizes.ListCount > 0 Then
50560   If Frm.cmbDocumentPapersizes.ListIndex > 0 Then
50570    .Papersize = Frm.cmbDocumentPapersizes.List(Frm.cmbDocumentPapersizes.ListIndex)
50580   End If
50590  End If
50600  If LenB(Frm.cmbPCXColors.ListIndex) > 0 Then
50610   .PCXColorscount = Frm.cmbPCXColors.ListIndex
50620  End If
50630  .PDFAllowAssembly = Abs(Frm.chkAllowAssembly.Value)
50640  .PDFAllowDegradedPrinting = Abs(Frm.chkAllowDegradedPrinting.Value)
50650  .PDFAllowFillIn = Abs(Frm.chkAllowFillIn.Value)
50660  .PDFAllowScreenReaders = Abs(Frm.chkAllowScreenReaders.Value)
50670  .PDFColorsCMYKToRGB = Abs(Frm.chkPDFCMYKtoRGB.Value)
50680  If LenB(Frm.cmbPDFColorModel.ListIndex) > 0 Then
50690   .PDFColorsColorModel = Frm.cmbPDFColorModel.ListIndex
50700  End If
50710  .PDFColorsPreserveHalftone = Abs(Frm.chkPDFPreserveHalftone.Value)
50720  .PDFColorsPreserveOverprint = Abs(Frm.chkPDFPreserveOverprint.Value)
50730  .PDFColorsPreserveTransfer = Abs(Frm.chkPDFPreserveTransfer.Value)
50740  .PDFCompressionColorCompression = Abs(Frm.chkPDFColorComp.Value)
50750  If LenB(Frm.cmbPDFColorComp.ListIndex) > 0 Then
50760   .PDFCompressionColorCompressionChoice = Frm.cmbPDFColorComp.ListIndex
50770  End If
50780  .PDFCompressionColorResample = Abs(Frm.chkPDFColorResample.Value)
50790  If LenB(Frm.cmbPDFColorResample.ListIndex) > 0 Then
50800   .PDFCompressionColorResampleChoice = Frm.cmbPDFColorResample.ListIndex
50810  End If
50820  If LenB(Frm.txtPDFColorRes.Text) > 0 Then
50830   .PDFCompressionColorResolution = Frm.txtPDFColorRes.Text
50840  End If
50850  .PDFCompressionGreyCompression = Abs(Frm.chkPDFGreyComp.Value)
50860  If LenB(Frm.cmbPDFGreyComp.ListIndex) > 0 Then
50870   .PDFCompressionGreyCompressionChoice = Frm.cmbPDFGreyComp.ListIndex
50880  End If
50890  .PDFCompressionGreyResample = Abs(Frm.chkPDFGreyResample.Value)
50900  If LenB(Frm.cmbPDFGreyResample.ListIndex) > 0 Then
50910   .PDFCompressionGreyResampleChoice = Frm.cmbPDFGreyResample.ListIndex
50920  End If
50930  If LenB(Frm.txtPDFGreyRes.Text) > 0 Then
50940   .PDFCompressionGreyResolution = Frm.txtPDFGreyRes.Text
50950  End If
50960  .PDFCompressionMonoCompression = Abs(Frm.chkPDFMonoComp.Value)
50970  If LenB(Frm.cmbPDFMonoComp.ListIndex) > 0 Then
50980   .PDFCompressionMonoCompressionChoice = Frm.cmbPDFMonoComp.ListIndex
50990  End If
51000  .PDFCompressionMonoResample = Abs(Frm.chkPDFMonoResample.Value)
51010  If LenB(Frm.cmbPDFMonoResample.ListIndex) > 0 Then
51020   .PDFCompressionMonoResampleChoice = Frm.cmbPDFMonoResample.ListIndex
51030  End If
51040  If LenB(Frm.txtPDFMonoRes.Text) > 0 Then
51050   .PDFCompressionMonoResolution = Frm.txtPDFMonoRes.Text
51060  End If
51070  .PDFCompressionTextCompression = Abs(Frm.chkPDFTextComp.Value)
51080  .PDFDisallowCopy = Abs(Frm.chkAllowCopy.Value)
51090  .PDFDisallowModifyAnnotations = Abs(Frm.chkAllowModifyAnnotations.Value)
51100  .PDFDisallowModifyContents = Abs(Frm.chkAllowModifyContents.Value)
51110  .PDFDisallowPrinting = Abs(Frm.chkAllowPrinting.Value)
51120  If Frm.cmbPDFEncryptor.ListIndex < 0 Then
51130    .PDFEncryptor = 0
51140   Else
51150    .PDFEncryptor = Frm.cmbPDFEncryptor.ItemData(Frm.cmbPDFEncryptor.ListIndex)
51160  End If
51170  .PDFFontsEmbedAll = Abs(Frm.chkPDFEmbedAll.Value)
51180  .PDFFontsSubSetFonts = Abs(Frm.chkPDFSubSetFonts.Value)
51190  If LenB(Frm.txtPDFSubSetPerc.Text) > 0 Then
51200   .PDFFontsSubSetFontsPercent = Frm.txtPDFSubSetPerc.Text
51210  End If
51220  .PDFGeneralASCII85 = Abs(Frm.chkPDFASCII85.Value)
51230  If LenB(Frm.cmbPDFRotate.ListIndex) > 0 Then
51240   .PDFGeneralAutorotate = Frm.cmbPDFRotate.ListIndex
51250  End If
51260  If LenB(Frm.cmbPDFCompat.ListIndex) > 0 Then
51270   .PDFGeneralCompatibility = Frm.cmbPDFCompat.ListIndex
51280  End If
51290  If LenB(Frm.cmbPDFOverprint.ListIndex) > 0 Then
51300   .PDFGeneralOverprint = Frm.cmbPDFOverprint.ListIndex
51310  End If
51320  If LenB(Frm.txtPDFRes.Text) > 0 Then
51330   .PDFGeneralResolution = Frm.txtPDFRes.Text
51340  End If
51350  .PDFHighEncryption = Abs(Frm.optEncHigh.Value)
51360  .PDFLowEncryption = Abs(Frm.optEncLow.Value)
51370  .PDFOptimize = Abs(Frm.chkPDFOptimize.Value)
51380  .PDFOwnerPass = Abs(Frm.chkOwnerPass.Value)
51390  .PDFUserPass = Abs(Frm.chkUserPass.Value)
51400  .PDFUseSecurity = Abs(Frm.chkUseSecurity.Value)
51410  If LenB(Frm.cmbPNGColors.ListIndex) > 0 Then
51420   .PNGColorscount = Frm.cmbPNGColors.ListIndex
51430  End If
51440  .PrintAfterSaving = Abs(Frm.chkPrintAfterSaving.Value)
51450  .PrintAfterSavingDuplex = Abs(Frm.chkPrintAfterSavingDuplex.Value)
51460  .PrintAfterSavingNoCancel = Abs(Frm.chkPrintAfterSavingNoCancel.Value)
51470  .PrintAfterSavingPrinter = Frm.cmbPrintAfterSavingPrinter.Text
51480  If LenB(Frm.cmbPrintAfterSavingQueryUser.ListIndex) > 0 Then
51490   .PrintAfterSavingQueryUser = Frm.cmbPrintAfterSavingQueryUser.ListIndex
51500  End If
51510  If LenB(Frm.cmbPrintAfterSavingTumble.ListIndex) > 0 Then
51520   .PrintAfterSavingTumble = Frm.cmbPrintAfterSavingTumble.ListIndex
51530  End If
51540  .PrinterTemppath = Frm.txtTemppath.Text
51550  If LenB(Frm.sldProcessPriority.Value) > 0 Then
51560   .ProcessPriority = Frm.sldProcessPriority.Value
51570  End If
51580  .ProgramFont = Frm.cmbFonts.List(Frm.cmbFonts.ListIndex)
51590  If LenB(Frm.cmbCharset.Text) > 0 Then
51600   .ProgramFontCharset = Frm.cmbCharset.Text
51610  End If
51620  If LenB(Frm.cmbProgramFontSize.Text) > 0 Then
51630   .ProgramFontSize = Frm.cmbProgramFontSize.Text
51640  End If
51650  If LenB(Frm.cmbPSLanguageLevel.ListIndex) > 0 Then
51660   .PSLanguageLevel = Frm.cmbPSLanguageLevel.ListIndex
51670  End If
51680  .RemoveSpaces = Abs(Frm.chkSpaces.Value)
51690  .RunProgramAfterSaving = Abs(Frm.chkRunProgramAfterSaving.Value)
51700  .RunProgramAfterSavingProgramname = Frm.cmbRunProgramAfterSavingProgramname.Text
51710  .RunProgramAfterSavingProgramParameters = Frm.txtRunProgramAfterSavingProgramParameters.Text
51720  .RunProgramAfterSavingWaitUntilReady = Abs(Frm.chkRunProgramAfterSavingWaitUntilReady.Value)
51730  If LenB(Frm.cmbRunProgramAfterSavingWindowstyle.ListIndex) > 0 Then
51740   .RunProgramAfterSavingWindowstyle = Frm.cmbRunProgramAfterSavingWindowstyle.ListIndex
51750  End If
51760  .RunProgramBeforeSaving = Abs(Frm.chkRunProgramBeforeSaving.Value)
51770  .RunProgramBeforeSavingProgramname = Frm.cmbRunProgramBeforeSavingProgramname.Text
51780  .RunProgramBeforeSavingProgramParameters = Frm.txtRunProgramBeforeSavingProgramParameters.Text
51790  If LenB(Frm.cmbRunProgramBeforeSavingWindowstyle.ListIndex) > 0 Then
51800   .RunProgramBeforeSavingWindowstyle = Frm.cmbRunProgramBeforeSavingWindowstyle.ListIndex
51810  End If
51820  .SaveFilename = Frm.txtSaveFilename.Text
51830  .SendEmailAfterAutoSaving = Abs(Frm.chkAutosaveSendEmail.Value)
51840  If LenB(Frm.cmbSendMailMethod.ListIndex) > 0 Then
51850   .SendMailMethod = Frm.cmbSendMailMethod.ListIndex
51860  End If
51870  .ShowAnimation = Abs(Frm.chkShowAnimation.Value)
51880  .StampFontColor = OleColorToHTMLColor(Frm.picStampFontColor.BackColor)
51890  If LenB(Frm.txtOutlineFontThickness.Text) > 0 Then
51900   .StampOutlineFontthickness = Frm.txtOutlineFontThickness.Text
51910  End If
51920  .StampString = Frm.txtStampString.Text
51930  .StampUseOutlineFont = Abs(Frm.chkStampUseOutlineFont.Value)
51940  .StandardAuthor = Frm.txtStandardAuthor.Text
51950  If LenB(Frm.cmbStandardSaveformat.ListIndex) > 0 Then
51960   .StandardSaveformat = Frm.cmbStandardSaveformat.ListIndex
51970  End If
51980  If LenB(Frm.cmbTIFFColors.ListIndex) > 0 Then
51990   .TIFFColorscount = Frm.cmbTIFFColors.ListIndex
52000  End If
52010  .UseAutosave = Abs(Frm.chkUseAutosave.Value)
52020  .UseAutosaveDirectory = Abs(Frm.chkUseAutosaveDirectory.Value)
52030  .UseCreationDateNow = Abs(Frm.chkUseCreationDateNow.Value)
52040  .UseCustomPaperSize = Frm.chkUseCustomPapersize.Value
52050  .UseFixPapersize = Abs(Frm.chkUseFixPaperSize.Value)
52060  .UseStandardAuthor = Abs(Frm.chkUseStandardAuthor.Value)
52070  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOptions2", "GetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
