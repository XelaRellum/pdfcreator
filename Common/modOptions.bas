Attribute VB_Name = "modOptions"
Option Explicit

' Module automatically generated with LanguagesTool from Frank Heindörfer, Philip Chinery
' 2003
' Email: thesmilyface@users.sourceforge.net

Global Options As tOptions

Public Type tOptions
 AutosaveDirectory As String
 AutosaveFilename As String
 AutosaveFormat As Long
 BitmapResolution As Long
 BMPColorscount As Long
 JPEGColorscount As Long
 JPEGQuality As Long
 Language As String
 LastSaveDirectory As String
 Logging As Long
 LogLines As Long
 PCXColorscount As Long
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
 PDFFontsEmbedAll As Long
 PDFFontsSubSetFonts As Long
 PDFFontsSubSetFontsPercent As Long
 PDFGeneralASCII85 As Long
 PDFGeneralAutorotate As Long
 PDFGeneralCompatibility As Long
 PDFGeneralOverprint As Long
 PDFGeneralResolution As Long
 PNGColorscount As Long
 PrinterStop As Long
 ProgramFont As String
 ProgramFontCharset As Long
 ProgramFontSize As Long
 StandardAuthor As String
 StartStandardProgram As Long
 TIFFColorscount As Long
 UseAutosave As Long
 UseAutosaveDirectory As Long
 UseCreationDateNow As Long
 UseStandardAuthor As Long
End Type

Public Function StandardOptions() As tOptions
 Dim myOptions As tOptions
 With myOptions
  .AutosaveDirectory = ""
  .AutosaveFilename = "<DateTime>"
  .AutosaveFormat = "0"
  .BitmapResolution = "150"
  .BMPColorscount = "0"
  .JPEGColorscount = "0"
  .JPEGQuality = "75"
  .Language = "english"
  .LastSaveDirectory = GetMyFiles
  .Logging = "0"
  .LogLines = "100"
  .PCXColorscount = "0"
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
  .PDFFontsEmbedAll = "1"
  .PDFFontsSubSetFonts = "1"
  .PDFFontsSubSetFontsPercent = "100"
  .PDFGeneralASCII85 = "0"
  .PDFGeneralAutorotate = "0"
  .PDFGeneralCompatibility = "1"
  .PDFGeneralOverprint = "0"
  .PDFGeneralResolution = "600"
  .PNGColorscount = "0"
  .PrinterStop = "0"
  .ProgramFont = "MS Sans Serif"
  .ProgramFontCharset = "0"
  .ProgramFontSize = "8"
  .StandardAuthor = ""
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
 On Error Resume Next
 Dim ini As clsINI, myOptions As tOptions, tStr As String, hOpt As New clsHash
 Set ini = New clsINI
 ini.FileName = PDFCreatorINIFile
 ini.Section = "Options"
 If ini.CheckIniFile = False Then
  ReadOptions = StandardOptions
  Exit Function
 End If
 ReadINISection PDFCreatorINIFile, "Options", hOpt
 With myOptions
  tStr = hOpt.Retrieve("AutosaveDirectory")
  If Len(tStr) > 0 Then
    .AutosaveDirectory = tStr
   Else
    .AutosaveDirectory = GetSpecialFolder(ssfPERSONAL)
  End If
  tStr = hOpt.Retrieve("AutosaveFilename")
  If Len(tStr) > 0 Then
    .AutosaveFilename = tStr
   Else
    .AutosaveFilename = "<DateTime>"
  End If
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
    .BMPColorscount = 0
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
  tStr = hOpt.Retrieve("Language")
  If Len(tStr) > 0 Then
    .Language = tStr
   Else
    .Language = "english"
  End If
  tStr = hOpt.Retrieve("LastSaveDirectory")
  If Len(tStr) > 0 Then
    If Dir(tStr, vbDirectory) <> "" Then
      .LastSaveDirectory = tStr
     Else
      .LastSaveDirectory = GetSpecialFolder(ssfPERSONAL)
    End If
   Else
    .LastSaveDirectory = GetSpecialFolder(ssfPERSONAL)
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
  tStr = hOpt.Retrieve("PCXColorscount")
  If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
    .PCXColorscount = CLng(tStr)
   Else
    .PCXColorscount = 0
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
  If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
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
  If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
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
  If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
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
    .PDFGeneralAutorotate = 0
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
  tStr = hOpt.Retrieve("ProgramFont")
  If Len(tStr) > 0 Then
    .ProgramFont = tStr
   Else
    .ProgramFont = "MS Sans Serif"
  End If
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
  tStr = hOpt.Retrieve("StandardAuthor")
  If Len(tStr) > 0 Then
    .StandardAuthor = tStr
   Else
    .StandardAuthor = ""
  End If
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

Public Sub SaveOptions(sOptions As tOptions)
 Dim ini As clsINI
 Set ini = New clsINI
 ini.FileName = PDFCreatorINIFile
 ini.Section = "Options"
 If ini.CheckIniFile = False Then
  ini.CreateIniFile
 End If
 With sOptions
  ini.SaveKey CStr(.AutosaveDirectory), "AutosaveDirectory"
  ini.SaveKey CStr(.AutosaveFilename), "AutosaveFilename"
  ini.SaveKey CStr(.AutosaveFormat), "AutosaveFormat"
  ini.SaveKey CStr(.BitmapResolution), "BitmapResolution"
  ini.SaveKey CStr(.BMPColorscount), "BMPColorscount"
  ini.SaveKey CStr(.JPEGColorscount), "JPEGColorscount"
  ini.SaveKey CStr(.JPEGQuality), "JPEGQuality"
  ini.SaveKey CStr(.Language), "Language"
  ini.SaveKey CStr(.LastSaveDirectory), "LastSaveDirectory"
  ini.SaveKey CStr(.Logging), "Logging"
  ini.SaveKey CStr(.LogLines), "LogLines"
  ini.SaveKey CStr(.PCXColorscount), "PCXColorscount"
  ini.SaveKey CStr(.PDFColorsCMYKToRGB), "PDFColorsCMYKToRGB"
  ini.SaveKey CStr(.PDFColorsColorModel), "PDFColorsColorModel"
  ini.SaveKey CStr(.PDFColorsPreserveHalftone), "PDFColorsPreserveHalftone"
  ini.SaveKey CStr(.PDFColorsPreserveOverprint), "PDFColorsPreserveOverprint"
  ini.SaveKey CStr(.PDFColorsPreserveTransfer), "PDFColorsPreserveTransfer"
  ini.SaveKey CStr(.PDFCompressionColorCompression), "PDFCompressionColorCompression"
  ini.SaveKey CStr(.PDFCompressionColorCompressionChoice), "PDFCompressionColorCompressionChoice"
  ini.SaveKey CStr(.PDFCompressionColorResample), "PDFCompressionColorResample"
  ini.SaveKey CStr(.PDFCompressionColorResampleChoice), "PDFCompressionColorResampleChoice"
  ini.SaveKey CStr(.PDFCompressionColorResolution), "PDFCompressionColorResolution"
  ini.SaveKey CStr(.PDFCompressionGreyCompression), "PDFCompressionGreyCompression"
  ini.SaveKey CStr(.PDFCompressionGreyCompressionChoice), "PDFCompressionGreyCompressionChoice"
  ini.SaveKey CStr(.PDFCompressionGreyResample), "PDFCompressionGreyResample"
  ini.SaveKey CStr(.PDFCompressionGreyResampleChoice), "PDFCompressionGreyResampleChoice"
  ini.SaveKey CStr(.PDFCompressionGreyResolution), "PDFCompressionGreyResolution"
  ini.SaveKey CStr(.PDFCompressionMonoCompression), "PDFCompressionMonoCompression"
  ini.SaveKey CStr(.PDFCompressionMonoCompressionChoice), "PDFCompressionMonoCompressionChoice"
  ini.SaveKey CStr(.PDFCompressionMonoResample), "PDFCompressionMonoResample"
  ini.SaveKey CStr(.PDFCompressionMonoResampleChoice), "PDFCompressionMonoResampleChoice"
  ini.SaveKey CStr(.PDFCompressionMonoResolution), "PDFCompressionMonoResolution"
  ini.SaveKey CStr(.PDFCompressionTextCompression), "PDFCompressionTextCompression"
  ini.SaveKey CStr(.PDFFontsEmbedAll), "PDFFontsEmbedAll"
  ini.SaveKey CStr(.PDFFontsSubSetFonts), "PDFFontsSubSetFonts"
  ini.SaveKey CStr(.PDFFontsSubSetFontsPercent), "PDFFontsSubSetFontsPercent"
  ini.SaveKey CStr(.PDFGeneralASCII85), "PDFGeneralASCII85"
  ini.SaveKey CStr(.PDFGeneralAutorotate), "PDFGeneralAutorotate"
  ini.SaveKey CStr(.PDFGeneralCompatibility), "PDFGeneralCompatibility"
  ini.SaveKey CStr(.PDFGeneralOverprint), "PDFGeneralOverprint"
  ini.SaveKey CStr(.PDFGeneralResolution), "PDFGeneralResolution"
  ini.SaveKey CStr(.PNGColorscount), "PNGColorscount"
  ini.SaveKey CStr(.PrinterStop), "PrinterStop"
  ini.SaveKey CStr(.ProgramFont), "ProgramFont"
  ini.SaveKey CStr(.ProgramFontCharset), "ProgramFontCharset"
  ini.SaveKey CStr(.ProgramFontSize), "ProgramFontSize"
  ini.SaveKey CStr(.StandardAuthor), "StandardAuthor"
  ini.SaveKey CStr(.StartStandardProgram), "StartStandardProgram"
  ini.SaveKey CStr(.TIFFColorscount), "TIFFColorscount"
  ini.SaveKey CStr(.UseAutosave), "UseAutosave"
  ini.SaveKey CStr(.UseAutosaveDirectory), "UseAutosaveDirectory"
  ini.SaveKey CStr(.UseCreationDateNow), "UseCreationDateNow"
  ini.SaveKey CStr(.UseStandardAuthor), "UseStandardAuthor"
 End With
 Set ini = Nothing
End Sub

Public Sub ShowOptions(Frm As Form, sOptions As tOptions)
 On Local Error Resume Next
 Dim i As Long
 With sOptions
  Frm.txtAutosaveDirectory.Text = .AutosaveDirectory
  Frm.txtAutosaveFilename.Text = .AutosaveFilename
  Frm.cmbAutosaveFormat.ListIndex = .AutosaveFormat
  Frm.txtBitmapResolution.Text = .BitmapResolution
  Frm.cmbBMPColors.ListIndex = .BMPColorscount
  Frm.cmbJPEGColors.ListIndex = .JPEGColorscount
  Frm.txtJPEGQuality.Text = .JPEGQuality
  Frm.cmbPCXColors.ListIndex = .PCXColorscount
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
  Frm.chkPDFEmbedAll.Value = .PDFFontsEmbedAll
  Frm.chkPDFSubSetFonts.Value = .PDFFontsSubSetFonts
  Frm.txtPDFSubSetPerc.Text = .PDFFontsSubSetFontsPercent
  Frm.chkPDFASCII85.Value = .PDFGeneralASCII85
  Frm.cmbPDFRotate.ListIndex = .PDFGeneralAutorotate
  Frm.cmbPDFCompat.ListIndex = .PDFGeneralCompatibility
  Frm.cmbPDFOverprint.ListIndex = .PDFGeneralOverprint
  Frm.txtPDFRes.Text = .PDFGeneralResolution
  Frm.cmbPNGColors.ListIndex = .PNGColorscount
  For i = 0 To Frm.cmbFonts.ListCount - 1
    If UCase$(Frm.cmbFonts.List(i)) = UCase$(.ProgramFont) Then
     Frm.cmbFonts.ListIndex = i
     Exit For
    End If
  Next i
  Frm.cmbCharset.Text = .ProgramFontCharset
  Frm.txtProgramFontSize.Text = .ProgramFontSize
  Frm.txtStandardAuthor.Text = .StandardAuthor
  Frm.cmbTIFFColors.ListIndex = .TIFFColorscount
  Frm.chkUseAutosave.Value = .UseAutosave
  Frm.chkUseAutosaveDirectory.Value = .UseAutosaveDirectory
  Frm.chkUseCreationDateNow.Value = .UseCreationDateNow
  Frm.chkUseStandardAuthor.Value = .UseStandardAuthor
 End With
End Sub

Public Sub GetOptions(Frm As Form, sOptions As tOptions)
 With sOptions
  .AutosaveDirectory = Frm.txtAutosaveDirectory.Text
  .AutosaveFilename = Frm.txtAutosaveFilename.Text
  .AutosaveFormat = Frm.cmbAutosaveFormat.ListIndex
  .BitmapResolution = Frm.txtBitmapResolution.Text
  .BMPColorscount = Frm.cmbBMPColors.ListIndex
  .JPEGColorscount = Frm.cmbJPEGColors.ListIndex
  .JPEGQuality = Frm.txtJPEGQuality.Text
  .PCXColorscount = Frm.cmbPCXColors.ListIndex
  .PDFColorsCMYKToRGB = Frm.chkPDFCMYKtoRGB.Value
  .PDFColorsColorModel = Frm.cmbPDFColorModel.ListIndex
  .PDFColorsPreserveHalftone = Frm.chkPDFPreserveHalftone.Value
  .PDFColorsPreserveOverprint = Frm.chkPDFPreserveOverprint.Value
  .PDFColorsPreserveTransfer = Frm.chkPDFPreserveTransfer.Value
  .PDFCompressionColorCompression = Frm.chkPDFColorComp.Value
  .PDFCompressionColorCompressionChoice = Frm.cmbPDFColorComp.ListIndex
  .PDFCompressionColorResample = Frm.chkPDFColorResample.Value
  .PDFCompressionColorResampleChoice = Frm.cmbPDFColorResample.ListIndex
  .PDFCompressionColorResolution = Frm.txtPDFColorRes.Text
  .PDFCompressionGreyCompression = Frm.chkPDFGreyComp.Value
  .PDFCompressionGreyCompressionChoice = Frm.cmbPDFGreyComp.ListIndex
  .PDFCompressionGreyResample = Frm.chkPDFGreyResample.Value
  .PDFCompressionGreyResampleChoice = Frm.cmbPDFGreyResample.ListIndex
  .PDFCompressionGreyResolution = Frm.txtPDFGreyRes.Text
  .PDFCompressionMonoCompression = Frm.chkPDFMonoComp.Value
  .PDFCompressionMonoCompressionChoice = Frm.cmbPDFMonoComp.ListIndex
  .PDFCompressionMonoResample = Frm.chkPDFMonoResample.Value
  .PDFCompressionMonoResampleChoice = Frm.cmbPDFMonoResample.ListIndex
  .PDFCompressionMonoResolution = Frm.txtPDFMonoRes.Text
  .PDFCompressionTextCompression = Frm.chkPDFTextComp.Value
  .PDFFontsEmbedAll = Frm.chkPDFEmbedAll.Value
  .PDFFontsSubSetFonts = Frm.chkPDFSubSetFonts.Value
  .PDFFontsSubSetFontsPercent = Frm.txtPDFSubSetPerc.Text
  .PDFGeneralASCII85 = Frm.chkPDFASCII85.Value
  .PDFGeneralAutorotate = Frm.cmbPDFRotate.ListIndex
  .PDFGeneralCompatibility = Frm.cmbPDFCompat.ListIndex
  .PDFGeneralOverprint = Frm.cmbPDFOverprint.ListIndex
  .PDFGeneralResolution = Frm.txtPDFRes.Text
  .PNGColorscount = Frm.cmbPNGColors.ListIndex
  .ProgramFont = Frm.cmbFonts.List(Frm.cmbFonts.ListIndex)
  .ProgramFontCharset = Frm.cmbCharset.Text
  .ProgramFontSize = Frm.txtProgramFontSize.Text
  .StandardAuthor = Frm.txtStandardAuthor.Text
  .TIFFColorscount = Frm.cmbTIFFColors.ListIndex
  .UseAutosave = Frm.chkUseAutosave.Value
  .UseAutosaveDirectory = Frm.chkUseAutosaveDirectory.Value
  .UseCreationDateNow = Frm.chkUseCreationDateNow.Value
  .UseStandardAuthor = Frm.chkUseStandardAuthor.Value
 End With
End Sub

Public Sub SetPrinterStop(StopPrinter As Boolean)
 If StopPrinter = True Then
   Options.PrinterStop = 1
   PrinterStop = True
  Else
   Options.PrinterStop = 0
   PrinterStop = False
 End If
 SaveOptions Options
End Sub

Public Sub SetLogging(Logging As Boolean)
 If Logging = True Then
   Options.Logging = 1
  Else
   Options.Logging = 0
 End If
 SaveOptions Options
End Sub

Public Sub SetLanguage(Language As String)
 Options.Language = Language
 SaveOptions Options
End Sub
