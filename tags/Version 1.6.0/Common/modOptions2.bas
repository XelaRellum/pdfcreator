Attribute VB_Name = "modOptions2"
Option Explicit

' Automatically generated with DeveloperTool by Frank Heindörfer
' 2003 - 2007
' Email: thesmilyface@users.sourceforge.net

Public Sub ShowOptions(Frm as Form, sOptions as tOptions)
 On Error Resume Next
 Dim i as Long, tList() as String, tStrA() As String, lsv As ListView
 With sOptions
  frm.cmbAdditionalGhostscriptParameters.Text = .AdditionalGhostscriptParameters
  frm.txtAdditionalGhostscriptSearchpath.Text = .AdditionalGhostscriptSearchpath
  frm.chkAddWindowsFontpath.Value = .AddWindowsFontpath
  frm.txtAutosaveDirectory.Text = .AutosaveDirectory
  frm.txtAutosaveFilename.Text = .AutosaveFilename
  frm.cmbAutosaveFormat.Listindex = .AutosaveFormat
  frm.chkAutosaveStartStandardProgram.Value = .AutosaveStartStandardProgram
  frm.txtBitmapResolution.Text = .BitmapResolution
  frm.cmbBMPColors.Listindex = .BMPColorscount
  frm.txtCustomPapersizeHeight.Text = .DeviceHeightPoints
  frm.txtCustomPapersizeWidth.Text = .DeviceWidthPoints
  frm.txtGSbin.Text = .DirectoryGhostscriptBinaries
  frm.txtGSfonts.Text = .DirectoryGhostscriptFonts
  frm.txtGSlib.Text = .DirectoryGhostscriptLibraries
  frm.txtGSResource.Text = .DirectoryGhostscriptResource
  frm.cmbEPSLanguageLevel.Listindex = .EPSLanguageLevel
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
  frm.chkFilenameSubst.Value = .FilenameSubstitutionsOnlyInTitle
  frm.cmbJPEGColors.Listindex = .JPEGColorscount
  frm.txtJPEGQuality.Text = .JPEGQuality
  frm.chkNoConfirmMessageSwitchingDefaultprinter = .NoConfirmMessageSwitchingDefaultprinter
  frm.chkNoProcessingAtStartup = .NoProcessingAtStartup
  frm.chkOnePagePerFile.Value = .OnePagePerFile
  frm.cmbOptionsDesign.Listindex = .OptionsDesign
  With Frm.cmbDocumentPapersizes
   For i = 0 To .ListCount - 1
    If UCase$(.List(i)) = UCase$(Options.Papersize) Then
     .ListIndex = i
     Exit For
    End If
   Next i
  End With
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
  frm.chkPDFOptimize.value = .PDFOptimize
  frm.chkOwnerPass.Value = .PDFOwnerPass
  frm.chkUserPass.Value = .PDFUserPass
  frm.chkUseSecurity.Value = .PDFUseSecurity
  frm.cmbPNGColors.Listindex = .PNGColorscount
  frm.chkPrintAfterSaving.Value = .PrintAfterSaving
  frm.chkPrintAfterSavingDuplex.Value = .PrintAfterSavingDuplex
  frm.chkPrintAfterSavingNoCancel.Value = .PrintAfterSavingNoCancel
  frm.cmbPrintAfterSavingPrinter.Text = .PrintAfterSavingPrinter
  frm.cmbPrintAfterSavingQueryUser.Listindex = .PrintAfterSavingQueryUser
  frm.cmbPrintAfterSavingTumble.Listindex = .PrintAfterSavingTumble
  frm.txtTemppath.Text = .PrinterTemppath
  frm.sldProcessPriority.Value = .ProcessPriority
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
  frm.chkRunProgramAfterSaving.Value = .RunProgramAfterSaving
  frm.cmbRunProgramAfterSavingProgramname.Text = .RunProgramAfterSavingProgramname
  frm.txtRunProgramAfterSavingProgramParameters.Text = .RunProgramAfterSavingProgramParameters
  frm.chkRunProgramAfterSavingWaitUntilReady.Value = .RunProgramAfterSavingWaitUntilReady
  frm.cmbRunProgramAfterSavingWindowstyle.Listindex = .RunProgramAfterSavingWindowstyle
  frm.chkRunProgramBeforeSaving.Value = .RunProgramBeforeSaving
  frm.cmbRunProgramBeforeSavingProgramname.Text = .RunProgramBeforeSavingProgramname
  frm.txtRunProgramBeforeSavingProgramParameters.Text = .RunProgramBeforeSavingProgramParameters
  frm.cmbRunProgramBeforeSavingWindowstyle.Listindex = .RunProgramBeforeSavingWindowstyle
  frm.txtSaveFilename.Text = .SaveFilename
  frm.chkAutosaveSendEmail.Value = .SendEmailAfterAutoSaving
  frm.cmbSendMailMethod.Listindex = .SendMailMethod
  frm.chkShowAnimation.Value = .ShowAnimation
  Frm.picStampFontColor.BackColor = HTMLcolorToOleColor(.StampFontColor)
  Frm.lblFontNameSize.Caption = .StampFontname & ", " & .StampFontsize
  frm.txtOutlineFontThickness.Text = .StampOutlineFontthickness
  frm.txtStampString.Text = .StampString
  frm.chkStampUseOutlineFont.Value = .StampUseOutlineFont
  frm.txtStandardAuthor.Text = .StandardAuthor
  frm.cmbStandardSaveformat.Listindex = .StandardSaveformat
  frm.cmbTIFFColors.Listindex = .TIFFColorscount
  frm.chkUseAutosave.Value = .UseAutosave
  frm.chkUseAutosaveDirectory.Value = .UseAutosaveDirectory
  frm.chkUseCreationDateNow.Value = .UseCreationDateNow
  frm.chkUseCustomPapersize.Value = .UseCustomPaperSize
  frm.chkUseFixPaperSize.Value = .UseFixPapersize
  frm.chkUseStandardAuthor.Value = .UseStandardAuthor
 End With
End Sub

Public Sub GetOptions(Frm as Form, sOptions as tOptions)
 Dim i as Long, tStr as String, lsv As ListView
 sOptions = StandardOptions
 With sOptions
 .AdditionalGhostscriptParameters =  Frm.cmbAdditionalGhostscriptParameters.Text
 .AdditionalGhostscriptSearchpath =  Frm.txtAdditionalGhostscriptSearchpath.Text
 .AddWindowsFontpath =  Abs(Frm.chkAddWindowsFontpath.Value)
 .AutosaveDirectory =  Frm.txtAutosaveDirectory.Text
 .AutosaveFilename =  Frm.txtAutosaveFilename.Text
 If LenB(Frm.cmbAutosaveFormat.Listindex) > 0 Then
  .AutosaveFormat =  Frm.cmbAutosaveFormat.Listindex
 End If
 .AutosaveStartStandardProgram =  Abs(Frm.chkAutosaveStartStandardProgram.Value)
 If LenB(Frm.txtBitmapResolution.Text) > 0 Then
  .BitmapResolution =  Frm.txtBitmapResolution.Text
 End If
 If LenB(Frm.cmbBMPColors.Listindex) > 0 Then
  .BMPColorscount =  Frm.cmbBMPColors.Listindex
 End If
 If LenB(Frm.txtCustomPapersizeHeight.Text) > 0 Then
  .DeviceHeightPoints =  Frm.txtCustomPapersizeHeight.Text
 End If
 If LenB(Frm.txtCustomPapersizeWidth.Text) > 0 Then
  .DeviceWidthPoints =  Frm.txtCustomPapersizeWidth.Text
 End If
 .DirectoryGhostscriptBinaries =  Frm.txtGSbin.Text
 .DirectoryGhostscriptFonts =  Frm.txtGSfonts.Text
 .DirectoryGhostscriptLibraries =  Frm.txtGSlib.Text
 .DirectoryGhostscriptResource =  Frm.txtGSResource.Text
 If LenB(Frm.cmbEPSLanguageLevel.Listindex) > 0 Then
  .EPSLanguageLevel =  Frm.cmbEPSLanguageLevel.Listindex
 End If
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
 .FilenameSubstitutionsOnlyInTitle =  Abs(Frm.chkFilenameSubst.Value)
 If LenB(Frm.cmbJPEGColors.Listindex) > 0 Then
  .JPEGColorscount =  Frm.cmbJPEGColors.Listindex
 End If
 If LenB(Frm.txtJPEGQuality.Text) > 0 Then
  .JPEGQuality =  Frm.txtJPEGQuality.Text
 End If
 .NoConfirmMessageSwitchingDefaultprinter =  Abs(Frm.chkNoConfirmMessageSwitchingDefaultprinter)
 .NoProcessingAtStartup =  Abs(Frm.chkNoProcessingAtStartup)
 .OnePagePerFile =  Abs(Frm.chkOnePagePerFile.Value)
 If LenB(Frm.cmbOptionsDesign.Listindex) > 0 Then
  .OptionsDesign =  Frm.cmbOptionsDesign.Listindex
 End If
 If Frm.cmbDocumentPapersizes.ListCount > 0 Then
  If Frm.cmbDocumentPapersizes.ListIndex > 0 Then
   .Papersize = Frm.cmbDocumentPapersizes.List(Frm.cmbDocumentPapersizes.ListIndex)
  End If
 End If
 If LenB(Frm.cmbPCXColors.Listindex) > 0 Then
  .PCXColorscount =  Frm.cmbPCXColors.Listindex
 End If
 .PDFAllowAssembly =  Abs(Frm.chkAllowAssembly.Value)
 .PDFAllowDegradedPrinting =  Abs(Frm.chkAllowDegradedPrinting.Value)
 .PDFAllowFillIn =  Abs(Frm.chkAllowFillIn.Value)
 .PDFAllowScreenReaders =  Abs(Frm.chkAllowScreenReaders.Value)
 .PDFColorsCMYKToRGB =  Abs(Frm.chkPDFCMYKtoRGB.Value)
 If LenB(Frm.cmbPDFColorModel.Listindex) > 0 Then
  .PDFColorsColorModel =  Frm.cmbPDFColorModel.Listindex
 End If
 .PDFColorsPreserveHalftone =  Abs(Frm.chkPDFPreserveHalftone.Value)
 .PDFColorsPreserveOverprint =  Abs(Frm.chkPDFPreserveOverprint.Value)
 .PDFColorsPreserveTransfer =  Abs(Frm.chkPDFPreserveTransfer.Value)
 .PDFCompressionColorCompression =  Abs(Frm.chkPDFColorComp.Value)
 If LenB(Frm.cmbPDFColorComp.Listindex) > 0 Then
  .PDFCompressionColorCompressionChoice =  Frm.cmbPDFColorComp.Listindex
 End If
 .PDFCompressionColorResample =  Abs(Frm.chkPDFColorResample.Value)
 If LenB(Frm.cmbPDFColorResample.Listindex) > 0 Then
  .PDFCompressionColorResampleChoice =  Frm.cmbPDFColorResample.Listindex
 End If
 If LenB(Frm.txtPDFColorRes.Text) > 0 Then
  .PDFCompressionColorResolution =  Frm.txtPDFColorRes.Text
 End If
 .PDFCompressionGreyCompression =  Abs(Frm.chkPDFGreyComp.Value)
 If LenB(Frm.cmbPDFGreyComp.Listindex) > 0 Then
  .PDFCompressionGreyCompressionChoice =  Frm.cmbPDFGreyComp.Listindex
 End If
 .PDFCompressionGreyResample =  Abs(Frm.chkPDFGreyResample.Value)
 If LenB(Frm.cmbPDFGreyResample.Listindex) > 0 Then
  .PDFCompressionGreyResampleChoice =  Frm.cmbPDFGreyResample.Listindex
 End If
 If LenB(Frm.txtPDFGreyRes.Text) > 0 Then
  .PDFCompressionGreyResolution =  Frm.txtPDFGreyRes.Text
 End If
 .PDFCompressionMonoCompression =  Abs(Frm.chkPDFMonoComp.Value)
 If LenB(Frm.cmbPDFMonoComp.Listindex) > 0 Then
  .PDFCompressionMonoCompressionChoice =  Frm.cmbPDFMonoComp.Listindex
 End If
 .PDFCompressionMonoResample =  Abs(Frm.chkPDFMonoResample.Value)
 If LenB(Frm.cmbPDFMonoResample.Listindex) > 0 Then
  .PDFCompressionMonoResampleChoice =  Frm.cmbPDFMonoResample.Listindex
 End If
 If LenB(Frm.txtPDFMonoRes.Text) > 0 Then
  .PDFCompressionMonoResolution =  Frm.txtPDFMonoRes.Text
 End If
 .PDFCompressionTextCompression =  Abs(Frm.chkPDFTextComp.Value)
 .PDFDisallowCopy =  Abs(Frm.chkAllowCopy.Value)
 .PDFDisallowModifyAnnotations =  Abs(Frm.chkAllowModifyAnnotations.Value)
 .PDFDisallowModifyContents =  Abs(Frm.chkAllowModifyContents.Value)
 .PDFDisallowPrinting =  Abs(Frm.chkAllowPrinting.Value)
 If Frm.cmbPDFEncryptor.ListIndex < 0 Then
   .PDFEncryptor = 0
  Else
   .PDFEncryptor =  frm.cmbPDFEncryptor.Itemdata(Frm.cmbPDFEncryptor.Listindex)
 End If
 .PDFFontsEmbedAll =  Abs(Frm.chkPDFEmbedAll.Value)
 .PDFFontsSubSetFonts =  Abs(Frm.chkPDFSubSetFonts.Value)
 If LenB(Frm.txtPDFSubSetPerc.Text) > 0 Then
  .PDFFontsSubSetFontsPercent =  Frm.txtPDFSubSetPerc.Text
 End If
 .PDFGeneralASCII85 =  Abs(Frm.chkPDFASCII85.Value)
 If LenB(Frm.cmbPDFRotate.Listindex) > 0 Then
  .PDFGeneralAutorotate =  Frm.cmbPDFRotate.Listindex
 End If
 If LenB(Frm.cmbPDFCompat.Listindex) > 0 Then
  .PDFGeneralCompatibility =  Frm.cmbPDFCompat.Listindex
 End If
 If LenB(Frm.cmbPDFOverprint.Listindex) > 0 Then
  .PDFGeneralOverprint =  Frm.cmbPDFOverprint.Listindex
 End If
 If LenB(Frm.txtPDFRes.Text) > 0 Then
  .PDFGeneralResolution =  Frm.txtPDFRes.Text
 End If
 .PDFHighEncryption =  Abs(Frm.optEncHigh.Value)
 .PDFLowEncryption =  Abs(Frm.optEncLow.Value)
 .PDFOptimize =  Abs(Frm.chkPDFOptimize.value)
 .PDFOwnerPass =  Abs(Frm.chkOwnerPass.Value)
 .PDFUserPass =  Abs(Frm.chkUserPass.Value)
 .PDFUseSecurity =  Abs(Frm.chkUseSecurity.Value)
 If LenB(Frm.cmbPNGColors.Listindex) > 0 Then
  .PNGColorscount =  Frm.cmbPNGColors.Listindex
 End If
 .PrintAfterSaving =  Abs(Frm.chkPrintAfterSaving.Value)
 .PrintAfterSavingDuplex =  Abs(Frm.chkPrintAfterSavingDuplex.Value)
 .PrintAfterSavingNoCancel =  Abs(Frm.chkPrintAfterSavingNoCancel.Value)
 .PrintAfterSavingPrinter =  Frm.cmbPrintAfterSavingPrinter.Text
 If LenB(Frm.cmbPrintAfterSavingQueryUser.Listindex) > 0 Then
  .PrintAfterSavingQueryUser =  Frm.cmbPrintAfterSavingQueryUser.Listindex
 End If
 If LenB(Frm.cmbPrintAfterSavingTumble.Listindex) > 0 Then
  .PrintAfterSavingTumble =  Frm.cmbPrintAfterSavingTumble.Listindex
 End If
 .PrinterTemppath =  Frm.txtTemppath.Text
 If LenB(Frm.sldProcessPriority.Value) > 0 Then
  .ProcessPriority =  Frm.sldProcessPriority.Value
 End If
 .ProgramFont =  Frm.cmbFonts.List(frm.cmbFonts.Listindex)
 If LenB(Frm.cmbCharset.Text) > 0 Then
  .ProgramFontCharset =  Frm.cmbCharset.Text
 End If
 If LenB(Frm.cmbProgramFontSize.text) > 0 Then
  .ProgramFontSize =  Frm.cmbProgramFontSize.text
 End If
 If LenB(Frm.cmbPSLanguageLevel.Listindex) > 0 Then
  .PSLanguageLevel =  Frm.cmbPSLanguageLevel.Listindex
 End If
 .RemoveSpaces =  Abs(Frm.chkSpaces.Value)
 .RunProgramAfterSaving =  Abs(Frm.chkRunProgramAfterSaving.Value)
 .RunProgramAfterSavingProgramname =  Frm.cmbRunProgramAfterSavingProgramname.Text
 .RunProgramAfterSavingProgramParameters =  Frm.txtRunProgramAfterSavingProgramParameters.Text
 .RunProgramAfterSavingWaitUntilReady =  Abs(Frm.chkRunProgramAfterSavingWaitUntilReady.Value)
 If LenB(Frm.cmbRunProgramAfterSavingWindowstyle.Listindex) > 0 Then
  .RunProgramAfterSavingWindowstyle =  Frm.cmbRunProgramAfterSavingWindowstyle.Listindex
 End If
 .RunProgramBeforeSaving =  Abs(Frm.chkRunProgramBeforeSaving.Value)
 .RunProgramBeforeSavingProgramname =  Frm.cmbRunProgramBeforeSavingProgramname.Text
 .RunProgramBeforeSavingProgramParameters =  Frm.txtRunProgramBeforeSavingProgramParameters.Text
 If LenB(Frm.cmbRunProgramBeforeSavingWindowstyle.Listindex) > 0 Then
  .RunProgramBeforeSavingWindowstyle =  Frm.cmbRunProgramBeforeSavingWindowstyle.Listindex
 End If
 .SaveFilename =  Frm.txtSaveFilename.Text
 .SendEmailAfterAutoSaving =  Abs(Frm.chkAutosaveSendEmail.Value)
 If LenB(Frm.cmbSendMailMethod.Listindex) > 0 Then
  .SendMailMethod =  Frm.cmbSendMailMethod.Listindex
 End If
 .ShowAnimation =  Abs(Frm.chkShowAnimation.Value)
 .StampFontColor = OleColorToHTMLColor(frm.picStampFontColor.BackColor)
 If LenB(Frm.txtOutlineFontThickness.Text) > 0 Then
  .StampOutlineFontthickness =  Frm.txtOutlineFontThickness.Text
 End If
 .StampString =  Frm.txtStampString.Text
 .StampUseOutlineFont =  Abs(Frm.chkStampUseOutlineFont.Value)
 .StandardAuthor =  Frm.txtStandardAuthor.Text
 If LenB(Frm.cmbStandardSaveformat.Listindex) > 0 Then
  .StandardSaveformat =  Frm.cmbStandardSaveformat.Listindex
 End If
 If LenB(Frm.cmbTIFFColors.Listindex) > 0 Then
  .TIFFColorscount =  Frm.cmbTIFFColors.Listindex
 End If
 .UseAutosave =  Abs(Frm.chkUseAutosave.Value)
 .UseAutosaveDirectory =  Abs(Frm.chkUseAutosaveDirectory.Value)
 .UseCreationDateNow =  Abs(Frm.chkUseCreationDateNow.Value)
 .UseCustomPaperSize =  Frm.chkUseCustomPapersize.Value
 .UseFixPapersize =  Abs(Frm.chkUseFixPaperSize.Value)
 .UseStandardAuthor =  Abs(Frm.chkUseStandardAuthor.Value)
 End With
End Sub
