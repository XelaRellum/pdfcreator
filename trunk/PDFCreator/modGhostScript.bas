Attribute VB_Name = "modGhostscript"
Option Explicit

Public GsDllLoaded As Long

Public Enum tGhostscriptDevice
 PDFWriter = 0
 PNGWriter = 1
 JPEGWriter = 2
 BMPWriter = 3
 PCXWriter = 4
 TIFFWriter = 5
 PSWriter = 6
 EPSWriter = 7
End Enum

Private GSParams() As String
Private GSParamsIndex As Integer

Public GS_ERROR
Public UseReturnPipe

'General
Public GS_COMPATIBILITY
Public GS_RESOLUTION
Public GS_AUTOROTATE
Public GS_OVERPRINT
Public GS_ASCII85

'Compression
Public GS_COMPRESSPAGES
Public GS_COMPRESSCOLOR
Public GS_COMPRESSGREY
Public GS_COMPRESSMONO
Public GS_COMPRESSCOLORMETHOD
Public GS_COMPRESSGREYMETHOD
Public GS_COMPRESSMONOMETHOD
Public GS_COMPRESSCOLORVALUE
Public GS_COMPRESSGREYVALUE
Public GS_COMPRESSMONOVALUE
Public GS_COMPRESSCOLORLEVEL
Public GS_COMPRESSGREYLEVEL
Public GS_COMPRESSMONOLEVEL
Public GS_COMPRESSCOLORAUTO
Public GS_COMPRESSGREYAUTO
Public GS_COMPRESSMONOAUTO
Public GS_COLORRESOLUTION
Public GS_GREYRESOLUTION
Public GS_MONORESOLUTION
Public GS_COLORRESAMPLE
Public GS_GREYRESAMPLE
Public GS_MONORESAMPLE
Public GS_COLORRESAMPLEMETHOD
Public GS_GREYRESAMPLEMETHOD
Public GS_MONORESAMPLEMETHOD

'Fonts
Public GS_EMBEDALLFONTS
Public GS_SUBSETFONTS
Public GS_SUBSETFONTPERC
Public GS_KEEPFONTNAMES

'Colors
Public GS_COLORMODEL
Public GS_CMYKTORGB
Public GS_PRESERVEOVERPRINT
Public GS_TRANSFERFUNCTIONS
Public GS_HALFTONE

'Bitmap
Public GS_BitmapRESOLUTION
Public GS_PNGColorscount
Public GS_JPEGColorscount
Public GS_BMPColorscount
Public GS_PCXColorscount
Public GS_TIFFColorscount
Public GS_JPEGQuality

' Postscript
Public GS_PSLanguageLevel
Public GS_EPSLanguageLevel

'** Begin Declarations for Encrypt PDF
Private Enum EncryptionStrength
   encLow = 48
   encStrong = 128
End Enum

Public Type EncryptData
   InputFile As String
   OutputFile As String
   UserPass As String
   OwnerPass As String
   AllowPrinting As Boolean
   AllowModifyContents As Boolean
   AllowCopy As Boolean
   AllowModifyAnnotations As Boolean
   AllowFillIn As Boolean '(128 bit only)
   AllowScreenReaders As Boolean '(128 bit only)
   AllowAssembly As Boolean '(128 bit only)
   AllowDegradedPrinting As Boolean '(128 bit only)
   EncryptionLevel As EncryptionStrength
End Type
'** End Declarations for Encrypt PDF


Public Sub GSInit(Options As tOptions)
 Dim Rotate(2) As String, Resample(2) As String, Colormodel(2) As String, _
  ColorsPreserveTransfer(1) As String
 Dim PNGColorscount(4) As String, JPEGColorscount(1) As String, BMPColorscount(6) As String, _
  PCXColorscount(5) As String, TIFFColorscount(7) As String, _
  PSLanguageLevel(3) As String

 Rotate(0) = "None": Rotate(1) = "All": Rotate(2) = "PageByPage"

 Resample(0) = "Bicubic": Resample(1) = "Subsample": Resample(2) = "Average"

 Colormodel(0) = "RGB": Colormodel(1) = "CMYK": Colormodel(2) = "GRAY"

 ColorsPreserveTransfer(0) = "Remove": ColorsPreserveTransfer(1) = "Preserve"

 PNGColorscount(0) = "png16m": PNGColorscount(1) = "png256"
 PNGColorscount(2) = "png16": PNGColorscount(3) = "png2"
 PNGColorscount(4) = "pnggray"

 JPEGColorscount(0) = "jpeg": JPEGColorscount(1) = "jpeggray"

 BMPColorscount(0) = "bmp32b": BMPColorscount(1) = "bmp16m"
 BMPColorscount(2) = "bmp256": BMPColorscount(3) = "bmp16"
 BMPColorscount(4) = "bmpsep8": BMPColorscount(5) = "bmpsep1"
 BMPColorscount(6) = "bmpgray"

 PCXColorscount(0) = "pcxcmyk": PCXColorscount(1) = "pcx24b"
 PCXColorscount(2) = "pcx256": PCXColorscount(3) = "pcx16"
 PCXColorscount(4) = "pcxmono": PCXColorscount(5) = "pcxgray"

 TIFFColorscount(0) = "tiff24nc": TIFFColorscount(1) = "tiff12nc"
 TIFFColorscount(2) = "tiffcrle": TIFFColorscount(3) = "tiffg3"
 TIFFColorscount(4) = "tiffg32d": TIFFColorscount(5) = "tiffg4"
 TIFFColorscount(6) = "tifflzw": TIFFColorscount(7) = "tiffpack"

 PSLanguageLevel(0) = "1": PSLanguageLevel(1) = "1.5"
 PSLanguageLevel(2) = "2": PSLanguageLevel(3) = "3"

With Options
 'General

 GS_COMPATIBILITY = "1." & (.PDFGeneralCompatibility + 2)
 GS_RESOLUTION = .PDFGeneralResolution
 GS_AUTOROTATE = Rotate(.PDFGeneralAutorotate)
 GS_OVERPRINT = .PDFGeneralOverprint
 GS_ASCII85 = Bool2Text(.PDFGeneralASCII85)

 'Compression
 GS_COMPRESSPAGES = Bool2Text(.PDFCompressionTextCompression)
 GS_COMPRESSCOLOR = Bool2Text(.PDFCompressionColorCompression)
 GS_COMPRESSGREY = Bool2Text(.PDFCompressionGreyCompression)
 GS_COMPRESSMONO = Bool2Text(.PDFCompressionMonoCompression)

 SelectColorCompression .PDFCompressionColorCompressionChoice
 SelectGreyCompression .PDFCompressionGreyCompressionChoice
 SelectMonoCompression .PDFCompressionMonoCompressionChoice

 GS_COMPRESSCOLORVALUE = Bool2Text(.PDFCompressionColorCompression)
 GS_COMPRESSGREYVALUE = Bool2Text(.PDFCompressionGreyCompression)
 GS_COMPRESSMONOVALUE = Bool2Text(.PDFCompressionMonoCompression)

 GS_COLORRESOLUTION = .PDFCompressionColorResolution
 GS_GREYRESOLUTION = .PDFCompressionGreyResolution
 GS_MONORESOLUTION = .PDFCompressionMonoResolution

 GS_COLORRESAMPLE = Bool2Text(.PDFCompressionColorResample)
 GS_GREYRESAMPLE = Bool2Text(.PDFCompressionGreyResample)
 GS_MONORESAMPLE = Bool2Text(.PDFCompressionMonoResample)

 GS_COLORRESAMPLEMETHOD = Resample(.PDFCompressionColorResampleChoice)
 GS_GREYRESAMPLEMETHOD = Resample(.PDFCompressionGreyResampleChoice)
 GS_MONORESAMPLEMETHOD = Resample(.PDFCompressionMonoResampleChoice)

 'Fonts
 GS_EMBEDALLFONTS = Bool2Text(.PDFFontsEmbedAll)
 GS_SUBSETFONTS = Bool2Text(.PDFFontsSubSetFonts)
 GS_SUBSETFONTPERC = .PDFFontsSubSetFontsPercent

 'Colors
 GS_COLORMODEL = Colormodel(.PDFColorsColorModel)
 GS_CMYKTORGB = Bool2Text(.PDFColorsCMYKToRGB)
 GS_PRESERVEOVERPRINT = Bool2Text(.PDFColorsPreserveOverprint)
 GS_TRANSFERFUNCTIONS = ColorsPreserveTransfer(.PDFColorsPreserveTransfer)
 GS_HALFTONE = Bool2Text(.PDFColorsPreserveHalftone)

 'Bitmap
 GS_BitmapRESOLUTION = .BitmapResolution
 GS_PNGColorscount = PNGColorscount(.PNGColorscount)
 GS_JPEGColorscount = JPEGColorscount(.JPEGColorscount)
 GS_BMPColorscount = BMPColorscount(.BMPColorscount)
 GS_PCXColorscount = PCXColorscount(.PCXColorscount)
 GS_TIFFColorscount = TIFFColorscount(.TIFFColorscount)
 GS_JPEGQuality = .JPEGQuality
 GS_PSLanguageLevel = PSLanguageLevel(.PSLanguageLevel)
 GS_EPSLanguageLevel = PSLanguageLevel(.EPSLanguageLevel)
End With
'Other
GS_ERROR = 0
UseReturnPipe = 1
End Sub

Public Function CallGScript(GSInputFile As String, GSOutputFile As String, _
 Options As tOptions, Ghostscriptdevice As tGhostscriptDevice)
Dim gsret As Long, i As Long

GSInit Options
InitParams

'50055 MsgBox _
 "Bin: " & Options.DirectoryGhostscriptBinaries & vbCrLf & _
 "Lib: " & Options.DirectoryGhostscriptLibraries & vbCrLf & _
 "Fonts: " & Options.DirectoryGhostscriptFonts
AddParams "-I" & Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
AddParams "-q"
AddParams "-dNOPAUSE"
AddParams "-dSAFER"
AddParams "-dBATCH"
Select Case Ghostscriptdevice
 Case 0: 'PDF
  AddParams "-sDEVICE=pdfwrite"
  AddParams "-dPDFSETTINGS=/printer"
  AddParams "-dCompatibilityLevel=" & GS_COMPATIBILITY
  AddParams "-r" & GS_RESOLUTION & "x" & GS_RESOLUTION
  AddParams "-dProcessColorModel=/Device" & GS_COLORMODEL
  AddParams "-dAutoRotatePages=/" & GS_AUTOROTATE
  AddParams "-dCompressPages=" & GS_COMPRESSPAGES
  AddParams "-dEmbedAllFonts=" & GS_EMBEDALLFONTS
  AddParams "-dSubsetFonts=" & GS_SUBSETFONTS
  AddParams "-dMaxSubsetPct=" & GS_SUBSETFONTPERC
  AddParams "-dConvertCMYKImagesToRGB=" & GS_CMYKTORGB

  AddParams "-dAutoFilterColorImages=" & GS_COMPRESSCOLORAUTO
  AddParams "-dEncodeColorImages=" & GS_COMPRESSCOLOR
  AddParams "-dColorImageFilter=/" & GS_COMPRESSCOLORMETHOD
  'AddParams "-dColorImageDict " & GS_COMPRESSCOLORLEVEL
  'AddParams "-sColorImageDict=<< /QFactor 0.25 /HSamples [1 1 1 1] /VSamples [1 1 1 1] >>"
  'AddParams "-c<< /QFactor 0.9 >>"


  AddParams "-dEncodeGrayImages=" & GS_COMPRESSGREY
  AddParams "-dGrayImageFilter=/" & GS_COMPRESSGREYMETHOD
  'AddParams "-dGrayACSImageDict " & GS_COMPRESSGREYLEVEL

  AddParams "-dColorImageResolution=" & GS_COLORRESOLUTION
  AddParams "-dGrayImageResolution=" & GS_GREYRESOLUTION
  AddParams "-dMonoImageResolution=" & GS_MONORESOLUTION
  AddParams "-dDownsampleColorImages=" & GS_COLORRESAMPLE
  AddParams "-dDownsampleGrayImages=" & GS_GREYRESAMPLE
  AddParams "-dDownsampleMonoImages=" & GS_MONORESAMPLE
  AddParams "-dColorImageDownsampleType=/" & GS_COLORRESAMPLEMETHOD
  AddParams "-dGrayImageDownsampleType=/" & GS_GREYRESAMPLEMETHOD
  AddParams "-dMonoImageDownsampleType=/" & GS_MONORESAMPLEMETHOD
  AddParams "-dPreserveOverprintSettings=" & GS_PRESERVEOVERPRINT
  AddParams "-dUCRandBGInfo=/Preserve"
  AddParams "-dUseFlateCompression=true"
  AddParams "-dParseDSCCommentsForDocInfo=true"
  AddParams "-dParseDSCComments=true"
  AddParams "-dOPM=" & GS_OVERPRINT
  AddParams "-dOffOptimizations=0"
  AddParams "-dLockDistillerParams=false"
  AddParams "-dGrayImageDepth=-1"
  AddParams "-dASCII85EncodePages=" & GS_ASCII85
  AddParams "-dDefaultRenderingIntent=/Default"
  AddParams "-dTransferFunctionInfo=/" & GS_TRANSFERFUNCTIONS
  AddParams "-dPreserveHalftoneInfo=" & GS_HALFTONE
  AddParams "-dOptimize=true"
  AddParams "-dDetectBlends=true"

'  If Options.PDFUseSecurity <> 0 Then
'   Dim Tempfile As String, Temppath As String, encPDF As EncryptData, retEnc As Boolean
'   Tempfile = GetTempFile(GetTempPath, "~PDF")
'   AddParams "-sOutputFile=" & Tempfile
'  Else
   AddParams "-sOutputFile=" & GSOutputFile
'  End If
  AddParams GSInputFile
  AddParams "-c quit"


 Case 1: 'PNG
  AddParams "-sDEVICE=" & GS_PNGColorscount
  AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
  AddParams "-sOutputFile=" & GSOutputFile
  AddParams GSInputFile
  AddParams "-c quit"
 Case 2: 'JPEG
  AddParams "-sDEVICE=" & GS_JPEGColorscount
  AddParams "-dJPEGQ=" & GS_JPEGQuality
  AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
  AddParams "-sOutputFile=" & GSOutputFile
  AddParams GSInputFile
  AddParams "-c quit"
 Case 3: 'BMP
  AddParams "-sDEVICE=" & GS_BMPColorscount
  AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
  AddParams "-q"
  AddParams "-sOutputFile=" & GSOutputFile
  AddParams GSInputFile
  AddParams "-c quit"
 Case 4: 'PCX
  AddParams "-sDEVICE=" & GS_PCXColorscount
  AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
  AddParams "-sOutputFile=" & GSOutputFile
  AddParams GSInputFile
  AddParams "-c quit"
 Case 5: 'TIFF
  AddParams "-sDEVICE=" & GS_TIFFColorscount
  AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
  AddParams "-sOutputFile=" & GSOutputFile
  AddParams GSInputFile
  AddParams "-c quit"
 Case 6: 'PS
  AddParams "-dLanguageLevel=" & GS_PSLanguageLevel
  AddParams "-sDEVICE=pswrite"
  AddParams "-sOutputFile=" & GSOutputFile
  AddParams GSInputFile
  AddParams "-c quit"
 Case 7: 'EPS
  AddParams "-dLanguageLevel=" & GS_EPSLanguageLevel
  AddParams "-sDEVICE=epswrite"
  AddParams "-sOutputFile=" & GSOutputFile
  AddParams GSInputFile
  AddParams "-c quit"
End Select

gsret = CallGS(GSParams)

'If (Options.PDFUseSecurity <> 0) And (Ghostscriptdevice = PDFWriter) Then
' SetEncryptionParams encPDF, Tempfile, GSOutputFile
' retEnc = EncryptPDF(encPDF)
' If retEnc = False Then
'   FileCopy Tempfile, GSOutputFile
'   IfLoggingWriteLogfile "Error with encryption - using unencrypted file"
' End If
' Kill Tempfile
'End If

End Function

Public Function OptimizePDF(GSInputFile As String, GSOutputFile As String)
Dim gsret

InitParams

AddParams "-I" & App.Path & "\lib;" & App.Path & "\fonts"
AddParams "-q"
AddParams "-dNODISPLAY"
AddParams "-dSAFER"
AddParams "-dDELAYSAFER"
AddParams "-- pdfopt.ps"
AddParams GSInputFile
AddParams GSOutputFile

frmMain.Refresh
gsret = CallGS(GSParams)
End Function

Public Sub ReturnValue(data As String)
 Dim newData As String
 newData = Replace(data, vbLf, "; ")
 IfLoggingWriteLogfile "Error: " & newData
' IfLoggingShowLogfile frmLog, frmMain
End Sub

Public Function Bool2Text(number As Long)
 If number = 1 Then
  Bool2Text = "true"
 Else
  Bool2Text = "false"
 End If
End Function

Private Sub SelectColorCompression(ByVal gsMethod)
 GS_COMPRESSCOLORAUTO = "false"
 Select Case gsMethod
  Case 0
   GS_COMPRESSCOLORAUTO = "true"
   GS_COMPRESSCOLORMETHOD = "null"
   GS_COMPRESSCOLORLEVEL = "null"
  Case 1
   GS_COMPRESSCOLORMETHOD = "DCTEncode"
   GS_COMPRESSCOLORLEVEL = "Maximum"
  Case 2
   GS_COMPRESSCOLORMETHOD = "DCTEncode"
   GS_COMPRESSCOLORLEVEL = "High"
  Case 3
   GS_COMPRESSCOLORMETHOD = "DCTEncode"
   GS_COMPRESSCOLORLEVEL = "Medium"
  Case 4
   GS_COMPRESSCOLORMETHOD = "DCTEncode"
   GS_COMPRESSCOLORLEVEL = "Low"
  Case 5
   GS_COMPRESSCOLORMETHOD = "DCTEncode"
   GS_COMPRESSCOLORLEVEL = "Minimum"
  Case 6
   GS_COMPRESSCOLORMETHOD = "FlateEncode"
   GS_COMPRESSCOLORLEVEL = "Maximum"
 End Select
End Sub

Private Sub SelectGreyCompression(ByVal gsMethod)
 GS_COMPRESSGREYAUTO = "false"
 Select Case gsMethod
  Case 0
   GS_COMPRESSGREYAUTO = "true"
   GS_COMPRESSGREYMETHOD = "null"
   GS_COMPRESSGREYLEVEL = "null"
  Case 1
   GS_COMPRESSGREYMETHOD = "DCTEncode"
   GS_COMPRESSGREYLEVEL = "Maximum"
  Case 2
   GS_COMPRESSGREYMETHOD = "DCTEncode"
   GS_COMPRESSGREYLEVEL = "High"
  Case 3
   GS_COMPRESSGREYMETHOD = "DCTEncode"
   GS_COMPRESSGREYLEVEL = "Medium"
  Case 4
   GS_COMPRESSGREYMETHOD = "DCTEncode"
   GS_COMPRESSGREYLEVEL = "Low"
  Case 5
   GS_COMPRESSGREYMETHOD = "DCTEncode"
   GS_COMPRESSGREYLEVEL = "Minimum"
  Case 6
   GS_COMPRESSGREYMETHOD = "FlateEncode"
   GS_COMPRESSGREYLEVEL = "Maximum"
 End Select
End Sub

Private Sub SelectMonoCompression(ByVal gsMethod)
 GS_COMPRESSGREYAUTO = "false"
 Select Case gsMethod
  Case 0
   GS_COMPRESSGREYAUTO = "true"
   GS_COMPRESSGREYMETHOD = "null"
   GS_COMPRESSGREYLEVEL = "null"
  Case 1
   GS_COMPRESSGREYMETHOD = "DCTEncode"
   GS_COMPRESSGREYLEVEL = "Maximum"
  Case 2
   GS_COMPRESSGREYMETHOD = "DCTEncode"
   GS_COMPRESSGREYLEVEL = "High"
  Case 3
   GS_COMPRESSGREYMETHOD = "DCTEncode"
   GS_COMPRESSGREYLEVEL = "Medium"
  Case 4
   GS_COMPRESSGREYMETHOD = "DCTEncode"
   GS_COMPRESSGREYLEVEL = "Low"
  Case 5
   GS_COMPRESSGREYMETHOD = "DCTEncode"
   GS_COMPRESSGREYLEVEL = "Minimum"
  Case 6
   GS_COMPRESSGREYMETHOD = "FlateEncode"
   GS_COMPRESSGREYLEVEL = "Maximum"
 End Select
End Sub

Private Sub InitParams()
 GSParamsIndex = 0
 ReDim GSParams(GSParamsIndex)
End Sub

Private Sub AddParams(strValue As String)
 GSParamsIndex = GSParamsIndex + 1
 ReDim Preserve GSParams(GSParamsIndex)
 GSParams(GSParamsIndex) = strValue
End Sub

Public Function EncryptPDF(encData As EncryptData) As Boolean
 Dim strPermissions As String, strShell As String, ret As Double

 strPermissions = vbNullString
 strPermissions = strPermissions & Abs(Int(Not encData.AllowPrinting))
 strPermissions = strPermissions & Abs(Int(Not encData.AllowModifyContents))
 strPermissions = strPermissions & Abs(Int(Not encData.AllowCopy))
 strPermissions = strPermissions & Abs(Int(Not encData.AllowModifyAnnotations))
 If Options.PDFHighEncryption Then
   strPermissions = strPermissions & Abs(Int(Not encData.AllowFillIn)) '(128 bit only)
   strPermissions = strPermissions & Abs(Int(Not encData.AllowScreenReaders)) '(128 bit only)
   strPermissions = strPermissions & Abs(Int(Not encData.AllowAssembly)) '(128 bit only)
   strPermissions = strPermissions & Abs(Int(Not encData.AllowDegradedPrinting)) '(128 bit only)
  Else
   strPermissions = strPermissions & "0000"
 End If

 strShell = App.Path & "\pdfencrypt.exe """ & encData.InputFile & """ """ & encData.OutputFile & """ User=" & encData.UserPass & " Owner=" & encData.OwnerPass & " " & strPermissions & " " & encData.EncryptionLevel
 'IfLoggingWriteLogfile strShell

 ret = RunProgramWait(strShell, False)

 If Dir$(encData.OutputFile) <> "" Then
  EncryptPDF = True
 End If
End Function

Public Sub SetEncryptionParams(ByRef encData As EncryptData, InputFile As String, OutputFile As String)
Dim retPasswd As Boolean

encData.InputFile = InputFile
encData.OutputFile = OutputFile

retPasswd = EnterPasswords(encData.UserPass, encData.OwnerPass, frmPassword)

encData.AllowPrinting = Options.PDFAllowPrinting
encData.AllowModifyContents = Options.PDFAllowModifyContents
encData.AllowCopy = Options.PDFAllowCopy
encData.AllowModifyAnnotations = Options.PDFAllowModifyAnnotations
encData.AllowFillIn = Options.PDFAllowFillIn
encData.AllowScreenReaders = Options.PDFAllowScreenReaders
encData.AllowAssembly = Options.PDFAllowAssembly
encData.AllowDegradedPrinting = Options.PDFAllowDegradedPrinting
If Options.PDFHighEncryption = True Then
 encData.EncryptionLevel = encStrong
Else
 encData.EncryptionLevel = encLow
End If
End Sub
