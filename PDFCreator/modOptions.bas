Attribute VB_Name = "modOptions"
Option Explicit

' Modul automatic generated with LanguagesTool from Frank Heindörfer
' 2003
' Email: thesmilyface@users.sourceforge.net

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
 Dim Options As tOptions
 With Options
  .AutosaveDirectory = ""
  .AutosaveFilename = "<DateTime>"
  .AutosaveFormat = "0"
  .BitmapResolution = "150"
  .BMPColorscount = "0"
  .JPEGColorscount = "0"
  .JPEGQuality = "75"
  .Language = "english"
  .LastSaveDirectory = GetSpecialFolder(ssfPERSONAL)
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
 StandardOptions = Options
End Function

Public Function ReadOptions() As tOptions
 On Error Resume Next
 Dim ini As clsINI, Options As tOptions, tStr as String
 Set ini = New clsINI
 ini.Filename = PDFCreatorINIFile
 ini.Section = "Options"
 If ini.Checkinifile = False Then
  ReadOptions = StandardOptions
  Exit Function
 End If
 With Options
  tStr =  ini.GetKeyFromSection("AutosaveDirectory")
  If Len(tStr) > 0 Then
    .AutosaveDirectory = tStr
   Else
    .AutosaveDirectory = GetSpecialFolder(ssfPERSONAL)
  End If
  tStr =  ini.GetKeyFromSection("AutosaveFilename")
  If Len(tStr) > 0 Then
    .AutosaveFilename = tStr
   Else
    .AutosaveFilename = ""
  End If
  tStr =  ini.GetKeyFromSection("AutosaveFormat")
  If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
    .AutosaveFormat = CLng(tStr)
   Else
    .AutosaveFormat = 0
  End If
  tStr =  ini.GetKeyFromSection("BitmapResolution")
  If CLng(tStr) >= 1 Then
    .BitmapResolution = CLng(tStr)
   Else
    .BitmapResolution = 150
  End If
  tStr =  ini.GetKeyFromSection("BMPColorscount")
  If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
    .BMPColorscount = CLng(tStr)
   Else
    .BMPColorscount = 0
  End If
  tStr =  ini.GetKeyFromSection("JPEGColorscount")
  If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
    .JPEGColorscount = CLng(tStr)
   Else
    .JPEGColorscount = 0
  End If
  tStr =  ini.GetKeyFromSection("JPEGQuality")
  If CLng(tStr) >= 0 And CLng(tStr) <= 100 Then
    .JPEGQuality = CLng(tStr)
   Else
    .JPEGQuality = 75
  End If
  tStr =  ini.GetKeyFromSection("Language")
  If Len(tStr) > 0 Then
    .Language = tStr
   Else
    .Language = ""
  End If
  tStr =  ini.GetKeyFromSection("LastSaveDirectory")
  If Len(tStr) > 0 Then
    If Dir(tStr,vbDirectory)<>"" Then
      .LastSaveDirectory = tStr
     Else
      .LastSaveDirectory = GetSpecialFolder(ssfPERSONAL)
    End If
   Else
    .LastSaveDirectory = GetSpecialFolder(ssfPERSONAL)
  End If
  tStr =  ini.GetKeyFromSection("Logging")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .Logging = CLng(tStr)
   Else
    .Logging = 0
  End If
  tStr =  ini.GetKeyFromSection("LogLines")
  If CLng(tStr) >= 100 And CLng(tStr) <= 1000 Then
    .LogLines = CLng(tStr)
   Else
    .LogLines = 100
  End If
  tStr =  ini.GetKeyFromSection("PCXColorscount")
  If CLng(tStr) >= 0 And CLng(tStr) <= 5 Then
    .PCXColorscount = CLng(tStr)
   Else
    .PCXColorscount = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFColorsCMYKToRGB")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFColorsCMYKToRGB = CLng(tStr)
   Else
    .PDFColorsCMYKToRGB = 1
  End If
  tStr =  ini.GetKeyFromSection("PDFColorsColorModel")
  If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
    .PDFColorsColorModel = CLng(tStr)
   Else
    .PDFColorsColorModel = 1
  End If
  tStr =  ini.GetKeyFromSection("PDFColorsPreserveHalftone")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFColorsPreserveHalftone = CLng(tStr)
   Else
    .PDFColorsPreserveHalftone = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFColorsPreserveOverprint")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFColorsPreserveOverprint = CLng(tStr)
   Else
    .PDFColorsPreserveOverprint = 1
  End If
  tStr =  ini.GetKeyFromSection("PDFColorsPreserveTransfer")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFColorsPreserveTransfer = CLng(tStr)
   Else
    .PDFColorsPreserveTransfer = 1
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionColorCompression")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionColorCompression = CLng(tStr)
   Else
    .PDFCompressionColorCompression = 1
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionColorCompressionChoice")
  If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
    .PDFCompressionColorCompressionChoice = CLng(tStr)
   Else
    .PDFCompressionColorCompressionChoice = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionColorResample")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionColorResample = CLng(tStr)
   Else
    .PDFCompressionColorResample = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionColorResampleChoice")
  If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
    .PDFCompressionColorResampleChoice = CLng(tStr)
   Else
    .PDFCompressionColorResampleChoice = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionColorResolution")
  If CLng(tStr) >= 0 Then
    .PDFCompressionColorResolution = CLng(tStr)
   Else
    .PDFCompressionColorResolution = 300
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionGreyCompression")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionGreyCompression = CLng(tStr)
   Else
    .PDFCompressionGreyCompression = 1
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionGreyCompressionChoice")
  If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
    .PDFCompressionGreyCompressionChoice = CLng(tStr)
   Else
    .PDFCompressionGreyCompressionChoice = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionGreyResample")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionGreyResample = CLng(tStr)
   Else
    .PDFCompressionGreyResample = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionGreyResampleChoice")
  If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
    .PDFCompressionGreyResampleChoice = CLng(tStr)
   Else
    .PDFCompressionGreyResampleChoice = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionGreyResolution")
  If CLng(tStr) >= 0 Then
    .PDFCompressionGreyResolution = CLng(tStr)
   Else
    .PDFCompressionGreyResolution = 300
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionMonoCompression")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionMonoCompression = CLng(tStr)
   Else
    .PDFCompressionMonoCompression = 1
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionMonoCompressionChoice")
  If CLng(tStr) >= 0 And CLng(tStr) <= 6 Then
    .PDFCompressionMonoCompressionChoice = CLng(tStr)
   Else
    .PDFCompressionMonoCompressionChoice = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionMonoResample")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionMonoResample = CLng(tStr)
   Else
    .PDFCompressionMonoResample = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionMonoResampleChoice")
  If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
    .PDFCompressionMonoResampleChoice = CLng(tStr)
   Else
    .PDFCompressionMonoResampleChoice = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionMonoResolution")
  If CLng(tStr) >= 0 Then
    .PDFCompressionMonoResolution = CLng(tStr)
   Else
    .PDFCompressionMonoResolution = 1200
  End If
  tStr =  ini.GetKeyFromSection("PDFCompressionTextCompression")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFCompressionTextCompression = CLng(tStr)
   Else
    .PDFCompressionTextCompression = 1
  End If
  tStr =  ini.GetKeyFromSection("PDFFontsEmbedAll")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFFontsEmbedAll = CLng(tStr)
   Else
    .PDFFontsEmbedAll = 1
  End If
  tStr =  ini.GetKeyFromSection("PDFFontsSubSetFonts")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFFontsSubSetFonts = CLng(tStr)
   Else
    .PDFFontsSubSetFonts = 1
  End If
  tStr =  ini.GetKeyFromSection("PDFFontsSubSetFontsPercent")
  If CLng(tStr) >= 0 Then
    .PDFFontsSubSetFontsPercent = CLng(tStr)
   Else
    .PDFFontsSubSetFontsPercent = 100
  End If
  tStr =  ini.GetKeyFromSection("PDFGeneralASCII85")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PDFGeneralASCII85 = CLng(tStr)
   Else
    .PDFGeneralASCII85 = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFGeneralAutorotate")
  If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
    .PDFGeneralAutorotate = CLng(tStr)
   Else
    .PDFGeneralAutorotate = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFGeneralCompatibility")
  If CLng(tStr) >= 0 And CLng(tStr) <= 2 Then
    .PDFGeneralCompatibility = CLng(tStr)
   Else
    .PDFGeneralCompatibility = 1
  End If
  tStr =  ini.GetKeyFromSection("PDFGeneralOverprint")
  If CLng(tStr) >= 0 And CLng(tStr) <= 1 Then
    .PDFGeneralOverprint = CLng(tStr)
   Else
    .PDFGeneralOverprint = 0
  End If
  tStr =  ini.GetKeyFromSection("PDFGeneralResolution")
  If CLng(tStr) >= 0 Then
    .PDFGeneralResolution = CLng(tStr)
   Else
    .PDFGeneralResolution = 600
  End If
  tStr =  ini.GetKeyFromSection("PNGColorscount")
  If CLng(tStr) >= 0 And CLng(tStr) <= 4 Then
    .PNGColorscount = CLng(tStr)
   Else
    .PNGColorscount = 0
  End If
  tStr =  ini.GetKeyFromSection("PrinterStop")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .PrinterStop = CLng(tStr)
   Else
    .PrinterStop = 0
  End If
  tStr =  ini.GetKeyFromSection("ProgramFont")
  If Len(tStr) > 0 Then
    .ProgramFont = tStr
   Else
    .ProgramFont = ""
  End If
  tStr =  ini.GetKeyFromSection("ProgramFontCharset")
  If CLng(tStr) >= 0 Then
    .ProgramFontCharset = CLng(tStr)
   Else
    .ProgramFontCharset = 0
  End If
  tStr =  ini.GetKeyFromSection("ProgramFontSize")
  If CLng(tStr) >= 1 And CLng(tStr) <= 72 Then
    .ProgramFontSize = CLng(tStr)
   Else
    .ProgramFontSize = 8
  End If
  tStr =  ini.GetKeyFromSection("StandardAuthor")
  If Len(tStr) > 0 Then
    .StandardAuthor = tStr
   Else
    .StandardAuthor = ""
  End If
  tStr =  ini.GetKeyFromSection("StartStandardProgram")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .StartStandardProgram = CLng(tStr)
   Else
    .StartStandardProgram = 1
  End If
  tStr =  ini.GetKeyFromSection("TIFFColorscount")
  If CLng(tStr) >= 0 And CLng(tStr) <= 7 Then
    .TIFFColorscount = CLng(tStr)
   Else
    .TIFFColorscount = 0
  End If
  tStr =  ini.GetKeyFromSection("UseAutosave")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .UseAutosave = CLng(tStr)
   Else
    .UseAutosave = 0
  End If
  tStr =  ini.GetKeyFromSection("UseAutosaveDirectory")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .UseAutosaveDirectory = CLng(tStr)
   Else
    .UseAutosaveDirectory = 1
  End If
  tStr =  ini.GetKeyFromSection("UseCreationDateNow")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .UseCreationDateNow = CLng(tStr)
   Else
    .UseCreationDateNow = 0
  End If
  tStr =  ini.GetKeyFromSection("UseStandardAuthor")
  If CLng(tStr) = 0 Or CLng(tStr) = 1 Then
    .UseStandardAuthor = CLng(tStr)
   Else
    .UseStandardAuthor = 0
  End If
 End With
 Set ini = Nothing
 ReadOptions = Options
End Function

Public Sub SaveOptions(Options as tOptions)
 Dim ini As clsINI
 Set ini = New clsINI
 ini.Filename = PDFCreatorINIFile
 ini.Section = "Options"
 If ini.CheckInifile = False Then
  ini.CreateInifile
 End If
 With Options
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

Public Sub ShowOptions(Frm as Form, Options as tOptions)
 On Local Error Resume Next
 Dim i as Long
 With Options
  frm.txtAutosaveDirectory.Text = .AutosaveDirectory
  frm.txtAutosaveFilename.Text = .AutosaveFilename
  frm.cmbAutosaveFormat.Listindex = .AutosaveFormat
  frm.txtBitmapResolution.Text = .BitmapResolution
  frm.cmbBMPColors.Listindex = .BMPColorscount
  frm.cmbJPEGColors.Listindex = .JPEGColorscount
  frm.txtJPEGQuality.Text = .JPEGQuality
  frm.cmbPCXColors.Listindex = .PCXColorscount
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
  frm.chkPDFEmbedAll.Value = .PDFFontsEmbedAll
  frm.chkPDFSubSetFonts.Value = .PDFFontsSubSetFonts
  frm.txtPDFSubSetPerc.Text = .PDFFontsSubSetFontsPercent
  frm.chkPDFASCII85.Value = .PDFGeneralASCII85
  frm.cmbPDFRotate.Listindex = .PDFGeneralAutorotate
  frm.cmbPDFCompat.Listindex = .PDFGeneralCompatibility
  frm.cmbPDFOverprint.Listindex = .PDFGeneralOverprint
  frm.txtPDFRes.Text = .PDFGeneralResolution
  frm.cmbPNGColors.Listindex = .PNGColorscount
  For i=0 to frm.cmbFonts.Listcount - 1
    If Ucase$(frm.cmbFonts.List(i)) = Ucase$(.ProgramFont) Then
     frm.cmbFonts.Listindex = i
     Exit For
    End If
  Next i
  frm.cmbCharset.Text = .ProgramFontCharset
  frm.txtProgramFontSize.text = .ProgramFontSize
  frm.txtStandardAuthor.Text = .StandardAuthor
  frm.cmbTIFFColors.Listindex = .TIFFColorscount
  frm.chkUseAutosave.Value = .UseAutosave
  frm.chkUseAutosaveDirectory.Value = .UseAutosaveDirectory
  frm.chkUseCreationDateNow.Value = .UseCreationDateNow
  frm.chkUseStandardAuthor.Value = .UseStandardAuthor
 End With
End Sub

Public Sub GetOptions(Frm as Form, Options as tOptions)
 Options = Readoptions
 With Options
  .AutosaveDirectory = frm.txtAutosaveDirectory.Text
  .AutosaveFilename = frm.txtAutosaveFilename.Text
  .AutosaveFormat = frm.cmbAutosaveFormat.Listindex
  .BitmapResolution = frm.txtBitmapResolution.Text
  .BMPColorscount = frm.cmbBMPColors.Listindex
  .JPEGColorscount = frm.cmbJPEGColors.Listindex
  .JPEGQuality = frm.txtJPEGQuality.Text
  .PCXColorscount = frm.cmbPCXColors.Listindex
  .PDFColorsCMYKToRGB = frm.chkPDFCMYKtoRGB.Value
  .PDFColorsColorModel = frm.cmbPDFColorModel.Listindex
  .PDFColorsPreserveHalftone = frm.chkPDFPreserveHalftone.Value
  .PDFColorsPreserveOverprint = frm.chkPDFPreserveOverprint.Value
  .PDFColorsPreserveTransfer = frm.chkPDFPreserveTransfer.Value
  .PDFCompressionColorCompression = frm.chkPDFColorComp.Value
  .PDFCompressionColorCompressionChoice = frm.cmbPDFColorComp.Listindex
  .PDFCompressionColorResample = frm.chkPDFColorResample.Value
  .PDFCompressionColorResampleChoice = frm.cmbPDFColorResample.Listindex
  .PDFCompressionColorResolution = frm.txtPDFColorRes.Text
  .PDFCompressionGreyCompression = frm.chkPDFGreyComp.Value
  .PDFCompressionGreyCompressionChoice = frm.cmbPDFGreyComp.Listindex
  .PDFCompressionGreyResample = frm.chkPDFGreyResample.Value
  .PDFCompressionGreyResampleChoice = frm.cmbPDFGreyResample.Listindex
  .PDFCompressionGreyResolution = frm.txtPDFGreyRes.Text
  .PDFCompressionMonoCompression = frm.chkPDFMonoComp.Value
  .PDFCompressionMonoCompressionChoice = frm.cmbPDFMonoComp.Listindex
  .PDFCompressionMonoResample = frm.chkPDFMonoResample.Value
  .PDFCompressionMonoResampleChoice = frm.cmbPDFMonoResample.Listindex
  .PDFCompressionMonoResolution = frm.txtPDFMonoRes.Text
  .PDFCompressionTextCompression = frm.chkPDFTextComp.Value
  .PDFFontsEmbedAll = frm.chkPDFEmbedAll.Value
  .PDFFontsSubSetFonts = frm.chkPDFSubSetFonts.Value
  .PDFFontsSubSetFontsPercent = frm.txtPDFSubSetPerc.Text
  .PDFGeneralASCII85 = frm.chkPDFASCII85.Value
  .PDFGeneralAutorotate = frm.cmbPDFRotate.Listindex
  .PDFGeneralCompatibility = frm.cmbPDFCompat.Listindex
  .PDFGeneralOverprint = frm.cmbPDFOverprint.Listindex
  .PDFGeneralResolution = frm.txtPDFRes.Text
  .PNGColorscount = frm.cmbPNGColors.Listindex
  .ProgramFont = frm.cmbFonts.List(frm.cmbFonts.Listindex)
  .ProgramFontCharset = frm.cmbCharset.Text
  .ProgramFontSize = frm.txtProgramFontSize.text
  .StandardAuthor = frm.txtStandardAuthor.Text
  .TIFFColorscount = frm.cmbTIFFColors.Listindex
  .UseAutosave = frm.chkUseAutosave.Value
  .UseAutosaveDirectory = frm.chkUseAutosaveDirectory.Value
  .UseCreationDateNow = frm.chkUseCreationDateNow.Value
  .UseStandardAuthor = frm.chkUseStandardAuthor.Value
 End With
End Sub

Public Sub SetPrinterStop(StopPrinter as Boolean)
 Dim Options As tOptions
 Options = ReadOptions
 If StopPrinter = True Then
   Options.PrinterStop = 1
   PrinterStop = True
  Else
   Options.PrinterStop = 0
   PrinterStop = False
 End If
 SaveOptions Options
End Sub

Public Sub SetLogging(Logging as Boolean)
 Dim Options As tOptions
 Options = ReadOptions
 If Logging = True Then
   Options.Logging = 1
  Else
   Options.Logging = 0
 End If
 SaveOptions Options
End Sub

Public Sub SetLanguage(Language as String)
 Dim Options As tOptions
 Options = ReadOptions
 Options.Language = Language
 SaveOptions Options
End Sub

