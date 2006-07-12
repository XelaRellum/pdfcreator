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
 TXTWriter = 8
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
Public GS_COMPRESSCOLORAUTO
Public GS_COMPRESSGREYAUTO
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
   encLow = 40
   encStrong = 128
End Enum

'Security
Private Const SEC_RESERVED0 = 2& + 1&
Private Const SEC_PRINT = 4&
Private Const SEC_MODIFY = 8&
Private Const SEC_COPY = 16&
Private Const SEC_FORM = 32&
Private Const SEC_RESERVED1 = 128& + 256&
'Revision 3 only
Private Const SEC_FILLFORM = 512&
Private Const SEC_SCREENREADERS = 1024&
Private Const SEC_ASSEMBLY = 2048&
Private Const SEC_HQPRINT = 4096&
 
Public Type EncryptData
 InputFile As String
 OutputFile As String
 UserPass As String
 OwnerPass As String
 DisallowPrinting As Boolean
 DisallowModifyContents As Boolean
 DisallowCopy As Boolean
 DisallowModifyAnnotations As Boolean
 AllowFillIn As Boolean '(128 bit only)
 AllowScreenReaders As Boolean '(128 bit only)
 AllowAssembly As Boolean '(128 bit only)
 AllowDegradedPrinting As Boolean '(128 bit only)
 EncryptionLevel As EncryptionStrength
End Type
'** End Declarations for Encrypt PDF

Public GS_OutStr As String

Private ParamCommands As Collection

Public Sub GSInit(Options As tOptions)
 Dim Rotate(2) As String, Resample(2) As String, Colormodel(2) As String, _
  ColorsPreserveTransfer(1) As String
 Dim PNGColorscount(4) As String, JPEGColorscount(1) As String, BMPColorscount(6) As String, _
  PCXColorscount(5) As String, TIFFColorscount(7) As String, _
  PSLanguageLevel(3) As String

 Rotate(0) = "None": Rotate(1) = "All": Rotate(2) = "PageByPage"

 Resample(0) = "Bicubic": Resample(1) = "Subsample": Resample(2) = "Average"

 Colormodel(0) = "RGB": Colormodel(1) = "CMYK": Colormodel(2) = "Gray"

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

Private Function CallGhostscript(Comment As String)
 Dim LastStop As Currency, res As Boolean, c As Currency
 LastStop = ExactTimer_Value()
 res = CallGS(GSParams)
 c = ExactTimer_Value() - LastStop
 IfLoggingWriteLogfile "Time for converting [" & Comment & "]: " & _
 Format$(Int(c) * (1 / 86400), "hh:nn:ss:") & Format$(((c) - Int(c)) * 1000, "000")
End Function

Private Function CreatePDF(GSInputFile As String, GSOutputFile As String, Options As tOptions)
 Dim FName As String, tStr As String, encPDF As EncryptData, tEnc As Boolean

 InitParams
 Set ParamCommands = New Collection

 tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
 If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
 End If
 If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
 End If
 AddParams "-I" & tStr
 AddParams "-q"
 AddParams "-dNOPAUSE"
 AddParams "-dSAFER"
 AddParams "-dBATCH"
 If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
  AddParams "-sFONTPATH=" & GetFontsDirectory
 End If
 AddParams "-sDEVICE=pdfwrite"
 If Options.DontUseDocumentSettings = 0 Then
  AddParams "-dCompatibilityLevel=" & GS_COMPATIBILITY
  AddParams "-r" & GS_RESOLUTION & "x" & GS_RESOLUTION
  AddParams "-dProcessColorModel=/Device" & GS_COLORMODEL
  AddParams "-dAutoRotatePages=/" & GS_AUTOROTATE
  AddParams "-dCompressPages=" & GS_COMPRESSPAGES
  AddParams "-dEmbedAllFonts=" & GS_EMBEDALLFONTS
  AddParams "-dSubsetFonts=" & GS_SUBSETFONTS
  AddParams "-dMaxSubsetPct=" & GS_SUBSETFONTPERC
  AddParams "-dConvertCMYKImagesToRGB=" & GS_CMYKTORGB
  
  If Options.UseFixPapersize <> 0 Then
   If Options.UseCustomPaperSize = 0 Then
     If LenB(Trim$(Options.Papersize)) > 0 Then
      AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
      AddParams "-dFIXEDMEDIA"
      AddParams "-dNORANGEPAGESIZE"
     End If
    Else
     If Options.DeviceWidthPoints >= 1 Then
      AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
     End If
     If Options.DeviceHeightPoints >= 1 Then
      AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
     End If
   End If
  End If
 
 End If
 tEnc = False
 If Options.PDFOptimize = 0 And Options.PDFUseSecurity <> 0 And _
    SecurityIsPossible = True And Options.PDFEncryptor = 0 Then
   If SetEncryptionParams(encPDF, "", "") = True Then
    tEnc = True
   If Len(encPDF.OwnerPass) > 0 Then
     AddParams "-sOwnerPassword=" & encPDF.OwnerPass & ""
    End If
    If Len(encPDF.UserPass) > 0 Then
     AddParams "-sUserPassword=" & encPDF.UserPass
    End If
    AddParams "-dPermissions=" & CalculatePermissions(encPDF)
    If GS_COMPATIBILITY = "1.4" Then
      AddParams "-dEncryptionR=3"
     Else
      AddParams "-dEncryptionR=2"
    End If
    If encPDF.EncryptionLevel = encLow Then
      AddParams "-dKeyLength=40"
     Else
      AddParams "-dKeyLength=128"
    End If
   Else
    If Options.UseAutosave = 0 Then
     MsgBox LanguageStrings.MessagesMsg23, vbCritical
    End If
  End If
 End If
 AddParams "-sOutputFile=" & GSOutputFile

 If Options.DontUseDocumentSettings = 0 Then
  SetColorParams
  SetGreyParams
  SetMonoParams

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
  AddParams "-dDetectBlends=true"

  AddAdditionalGhostscriptParameters

  AddParamCommands
 End If

 AddParams "-f"
 AddParams GSInputFile
 ShowParams
 If tEnc = True Then
   CallGhostscript "PDF with encryption"
  Else
   CallGhostscript "PDF without encryption"
 End If
End Function

Private Function CreatePNG(GSInputFile As String, GSOutputFile As String, Options As tOptions)
 Dim Path As String, FName As String, Ext As String, tStr As String

 GSInit Options
 InitParams
 Set ParamCommands = New Collection

 If Options.OnePagePerFile = 1 Then
  SplitPath GSOutputFile, , Path, , FName, Ext
  GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
 End If
 tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
 If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
 End If
 If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
 End If
 AddParams "-I" & tStr
 AddParams "-q"
 AddParams "-dNOPAUSE"
 AddParams "-dSAFER"
 AddParams "-dBATCH"
 If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
  AddParams "-sFONTPATH=" & GetFontsDirectory
 End If

 AddParams "-sDEVICE=" & GS_PNGColorscount

 If Options.DontUseDocumentSettings = 0 Then
  If Options.UseFixPapersize <> 0 Then
   If Options.UseCustomPaperSize = 0 Then
     If LenB(Trim$(Options.Papersize)) > 0 Then
      AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
      AddParams "-dFIXEDMEDIA"
      AddParams "-dNORANGEPAGESIZE"
     End If
    Else
     If Options.DeviceWidthPoints >= 1 Then
      AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
     End If
     If Options.DeviceHeightPoints >= 1 Then
      AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
     End If
   End If
  End If
  AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
  AddParams "-sOutputFile=" & GSOutputFile
 End If

 AddAdditionalGhostscriptParameters

 AddParams "-f"
 AddParams GSInputFile
 ShowParams
 CallGhostscript "PNG"
End Function

Private Function CreateJPEG(GSInputFile As String, GSOutputFile As String, Options As tOptions)
 Dim Path As String, FName As String, Ext As String, tStr As String

 GSInit Options
 InitParams
 Set ParamCommands = New Collection

 If Options.OnePagePerFile = 1 Then
  SplitPath GSOutputFile, , Path, , FName, Ext
  GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
 End If
 tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
 If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
 End If
 If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
 End If
 AddParams "-I" & tStr
 AddParams "-q"
 AddParams "-dNOPAUSE"
 AddParams "-dSAFER"
 AddParams "-dBATCH"
 If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
  AddParams "-sFONTPATH=" & GetFontsDirectory
 End If

 AddParams "-sDEVICE=" & GS_JPEGColorscount
 If Options.DontUseDocumentSettings = 0 Then
  AddParams "-dJPEGQ=" & GS_JPEGQuality
  If Options.UseFixPapersize <> 0 Then
   If Options.UseCustomPaperSize = 0 Then
     If LenB(Trim$(Options.Papersize)) > 0 Then
      AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
      AddParams "-dFIXEDMEDIA"
      AddParams "-dNORANGEPAGESIZE"
     End If
    Else
     If Options.DeviceWidthPoints >= 1 Then
      AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
     End If
     If Options.DeviceHeightPoints >= 1 Then
      AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
     End If
   End If
  End If
  AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
  AddParams "-sOutputFile=" & GSOutputFile
 End If

 AddAdditionalGhostscriptParameters

 AddParams "-f"
 AddParams GSInputFile
 ShowParams
 CallGhostscript "JPEG"
End Function

Private Function CreateBMP(GSInputFile As String, GSOutputFile As String, Options As tOptions)
 Dim Path As String, FName As String, Ext As String, tStr As String

 GSInit Options
 InitParams
 Set ParamCommands = New Collection

 If Options.OnePagePerFile = 1 Then
  SplitPath GSOutputFile, , Path, , FName, Ext
  GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
 End If
 tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
 If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
 End If
 If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
 End If
 AddParams "-I" & tStr
 AddParams "-q"
 AddParams "-dNOPAUSE"
 AddParams "-dSAFER"
 AddParams "-dBATCH"
 If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
  AddParams "-sFONTPATH=" & GetFontsDirectory
 End If

 AddParams "-sDEVICE=" & GS_BMPColorscount
 If Options.DontUseDocumentSettings = 0 Then
  If Options.UseFixPapersize <> 0 Then
   If Options.UseCustomPaperSize = 0 Then
     If LenB(Trim$(Options.Papersize)) > 0 Then
      AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
      AddParams "-dFIXEDMEDIA"
      AddParams "-dNORANGEPAGESIZE"
     End If
    Else
     If Options.DeviceWidthPoints >= 1 Then
      AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
     End If
     If Options.DeviceHeightPoints >= 1 Then
      AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
     End If
   End If
  End If
  AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
 End If
 AddParams "-sOutputFile=" & GSOutputFile

 AddAdditionalGhostscriptParameters

 AddParams "-f"
 AddParams GSInputFile
 ShowParams
 CallGhostscript "BMP"
End Function

Private Function CreatePCX(GSInputFile As String, GSOutputFile As String, Options As tOptions)
 Dim Path As String, FName As String, Ext As String, tStr As String

 GSInit Options
 InitParams
 Set ParamCommands = New Collection

 If Options.OnePagePerFile = 1 Then
  SplitPath GSOutputFile, , Path, , FName, Ext
  GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
 End If
 tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
 If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
 End If
 If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
 End If
 AddParams "-I" & tStr
 AddParams "-q"
 AddParams "-dNOPAUSE"
 AddParams "-dSAFER"
 AddParams "-dBATCH"
 If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
  AddParams "-sFONTPATH=" & GetFontsDirectory
 End If

 AddParams "-sDEVICE=" & GS_PCXColorscount
 If Options.DontUseDocumentSettings = 0 Then
  If Options.UseFixPapersize <> 0 Then
   If Options.UseCustomPaperSize = 0 Then
     If LenB(Trim$(Options.Papersize)) > 0 Then
      AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
      AddParams "-dFIXEDMEDIA"
      AddParams "-dNORANGEPAGESIZE"
     End If
    Else
     If Options.DeviceWidthPoints >= 1 Then
      AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
     End If
     If Options.DeviceHeightPoints >= 1 Then
      AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
     End If
   End If
  End If
  AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
 End If
 AddParams "-sOutputFile=" & GSOutputFile

 AddAdditionalGhostscriptParameters

 AddParams "-f"
 AddParams GSInputFile
 ShowParams
 CallGhostscript "PCX"
End Function

Private Function CreateTIFF(GSInputFile As String, GSOutputFile As String, Options As tOptions)
 Dim Path As String, FName As String, Ext As String, tStr As String

 GSInit Options
 InitParams
 Set ParamCommands = New Collection

 If Options.OnePagePerFile = 1 Then
  SplitPath GSOutputFile, , Path, , FName, Ext
  GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
 End If
 tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
 If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
 End If
 If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
 End If
 AddParams "-I" & tStr
 AddParams "-q"
 AddParams "-dNOPAUSE"
 AddParams "-dSAFER"
 AddParams "-dBATCH"
 If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
  AddParams "-sFONTPATH=" & GetFontsDirectory
 End If

 AddParams "-sDEVICE=" & GS_TIFFColorscount
 If Options.DontUseDocumentSettings = 0 Then
  If Options.UseFixPapersize <> 0 Then
   If Options.UseCustomPaperSize = 0 Then
     If LenB(Trim$(Options.Papersize)) > 0 Then
      AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
      AddParams "-dFIXEDMEDIA"
      AddParams "-dNORANGEPAGESIZE"
     End If
    Else
     If Options.DeviceWidthPoints >= 1 Then
      AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
     End If
     If Options.DeviceHeightPoints >= 1 Then
      AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
     End If
   End If
  End If
  AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
 End If
 AddParams "-sOutputFile=" & GSOutputFile

 AddAdditionalGhostscriptParameters

 AddParams "-f"
 AddParams GSInputFile
 ShowParams
 CallGhostscript "TIFF"
End Function

Private Function CreatePS(GSInputFile As String, GSOutputFile As String, Options As tOptions)
 Dim Path As String, FName As String, Ext As String, tStr As String

 GSInit Options
 InitParams
 Set ParamCommands = New Collection

 If Options.OnePagePerFile = 1 Then
  SplitPath GSOutputFile, , Path, , FName, Ext
  GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
 End If
 tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
 If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
 End If
 If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
 End If
 AddParams "-I" & tStr
 AddParams "-q"
 AddParams "-dNOPAUSE"
 AddParams "-dSAFER"
 AddParams "-dBATCH"
 If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
  AddParams "-sFONTPATH=" & GetFontsDirectory
 End If

 AddParams "-sDEVICE=pswrite"
 If Options.DontUseDocumentSettings = 0 Then
  If Options.UseFixPapersize <> 0 Then
   If Options.UseCustomPaperSize = 0 Then
     If LenB(Trim$(Options.Papersize)) > 0 Then
      AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
      AddParams "-dFIXEDMEDIA"
      AddParams "-dNORANGEPAGESIZE"
     End If
    Else
     If Options.DeviceWidthPoints >= 1 Then
      AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
     End If
     If Options.DeviceHeightPoints >= 1 Then
      AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
     End If
   End If
  End If
  AddParams "-dLanguageLevel=" & GS_PSLanguageLevel
 End If
 AddParams "-sOutputFile=" & GSOutputFile

 AddAdditionalGhostscriptParameters

 AddParams "-f"
 AddParams GSInputFile
 ShowParams
 CallGhostscript "PS"
End Function

Private Function CreateEPS(GSInputFile As String, GSOutputFile As String, Options As tOptions)
 Dim Path As String, FName As String, Ext As String, tStr As String

 GSInit Options
 InitParams
 Set ParamCommands = New Collection

 If Options.OnePagePerFile = 1 Then
  SplitPath GSOutputFile, , Path, , FName, Ext
  GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
 End If
 tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
 If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
 End If
 If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
 End If
 AddParams "-I" & tStr
 AddParams "-q"
 AddParams "-dNOPAUSE"
 AddParams "-dSAFER"
 AddParams "-dBATCH"
 If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
  AddParams "-sFONTPATH=" & GetFontsDirectory
 End If

 AddParams "-sDEVICE=epswrite"
 If Options.DontUseDocumentSettings = 0 Then
  If Options.UseFixPapersize <> 0 Then
   If Options.UseCustomPaperSize = 0 Then
     If LenB(Trim$(Options.Papersize)) > 0 Then
      AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
      AddParams "-dFIXEDMEDIA"
      AddParams "-dNORANGEPAGESIZE"
     End If
    Else
     If Options.DeviceWidthPoints >= 1 Then
      AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
     End If
     If Options.DeviceHeightPoints >= 1 Then
      AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
     End If
   End If
  End If
  AddParams "-dLanguageLevel=" & GS_EPSLanguageLevel
 End If
 AddParams "-sOutputFile=" & GSOutputFile

 AddAdditionalGhostscriptParameters

 AddParams "-f"
 AddParams GSInputFile
 ShowParams
 CallGhostscript "EPS"
End Function

Private Function CreateTXT(GSInputFile As String, Options As tOptions)
 Dim Path As String, FName As String, Ext As String, tStr As String

 GSInit Options
 InitParams
 Set ParamCommands = New Collection

 tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
 If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
 End If
 If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
 End If
 AddParams "-I" & tStr
 AddParams "-q"
 AddParams "-dNOPAUSE"
 AddParams "-dSAFER"
 AddParams "-dBATCH"
 AddParams "-dNODISPLAY"
 AddParams "-dDELAYBIND"
 AddParams "-dWRITESYSTEMDICT"
 AddParams "-dSIMPLE"
 AddParams "ps2ascii.ps"
 AddParams GSInputFile
 ShowParams
 CallGhostscript "TXT"
End Function

Public Function CallGScript(GSInputFile As String, GSOutputFile As String, _
 Options As tOptions, Ghostscriptdevice As tGhostscriptDevice)
 Dim enc As Boolean, encPDF As EncryptData, retEnc As Boolean, _
  Tempfile As String, tL As Long
 GSInit Options
 Select Case Ghostscriptdevice
  Case 0: 'PDF
   With Options
    If .PDFOptimize = 1 And .PDFUseSecurity = 0 Then
      Tempfile = GetTempFile(GetTempPath, "~CP")
      KillFile Tempfile
      CreatePDF GSInputFile, Tempfile, Options
      OptimizePDF Tempfile, GSOutputFile
      KillFile Tempfile
     Else
      If .PDFUseSecurity <> 0 And SecurityIsPossible = True Then
        If .PDFEncryptor = 1 Then
          enc = SetEncryptionParams(encPDF, GSInputFile, GSOutputFile)
          If enc = True Then
           Tempfile = GetTempFile(GetTempPath, "~CP")
           KillFile Tempfile
           CreatePDF GSInputFile, Tempfile, Options
           encPDF.InputFile = Tempfile
           retEnc = EncryptPDF(encPDF)
           KillFile encPDF.InputFile
           If retEnc = False Then
            IfLoggingWriteLogfile "Error with encryption - using unencrypted file"
            Name GSInputFile As GSOutputFile
           End If
          End If
         Else
          tL = .PDFOptimize
          .PDFOptimize = 0
          CreatePDF GSInputFile, GSOutputFile, Options
          .PDFOptimize = tL
        End If
       Else
        CreatePDF GSInputFile, GSOutputFile, Options
      End If
    End If
   End With
  Case 1: 'PNG
   CreatePNG GSInputFile, GSOutputFile, Options
  Case 2: 'JPEG
   CreateJPEG GSInputFile, GSOutputFile, Options
  Case 3: 'BMP
   CreateBMP GSInputFile, GSOutputFile, Options
  Case 4: 'PCX
   CreatePCX GSInputFile, GSOutputFile, Options
  Case 5: 'TIFF
   CreateTIFF GSInputFile, GSOutputFile, Options
  Case 6: 'PS
   CreatePS GSInputFile, GSOutputFile, Options
  Case 7: 'EPS
   CreateEPS GSInputFile, GSOutputFile, Options
  Case 8: 'TXT
   CreateTXT GSInputFile, Options
   CreateTextFile GSOutputFile, GS_OutStr
 End Select
End Function

Public Function OptimizePDF(PDFInputFilename As String, PDFOutputFilename As String) As Boolean
 Dim LastStop As Currency, tStr As String, c As Currency
 InitParams
 tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
 If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
 End If
 If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
 End If
 AddParams "-I" & tStr
 AddParams "-q"
 AddParams "-dNODISPLAY"
 AddParams "-dSAFER"
 AddParams "-dDELAYSAFER"
 AddParams "--"
 AddParams "pdfopt.ps"
 AddParams PDFInputFilename
 AddParams PDFOutputFilename

 GSParams(0) = "pdfopt"
 LastStop = ExactTimer_Value()
 OptimizePDF = CallGS(GSParams)
 c = ExactTimer_Value() - LastStop
 IfLoggingWriteLogfile "Time for optimizing: " & _
 Format$(Int(c) * (1 / 86400), "hh:nn:ss:") & Format$(((c) - Int(c)) * 1000, "000")
End Function

Public Function Bool2Text(Number As Long)
 If Number = 1 Then
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
 Select Case gsMethod
  Case 0
   GS_COMPRESSMONOMETHOD = "CCITTFaxEncode"
  Case 1
   GS_COMPRESSMONOMETHOD = "FlateEncode"
  Case 2
   GS_COMPRESSMONOMETHOD = "RunLengthEncode"
  Case 3
   GS_COMPRESSMONOMETHOD = "LZWEncode"
 End Select
End Sub

Private Sub InitParams()
 GSParamsIndex = 0
 ReDim GSParams(GSParamsIndex)
End Sub

Private Function ShowParams() As String
 Dim i As Long, tStr As String
 tStr = GSParams(LBound(GSParams))
 For i = LBound(GSParams) + 1 To UBound(GSParams)
  tStr = tStr & vbCrLf & GSParams(i)
 Next i
 IfLoggingWriteLogfile "Ghostscriptparameter:" & vbCrLf & tStr
 ShowParams = "Ghostscriptparameter:" & vbCrLf & tStr
End Function

Private Sub AddParams(strValue As String)
 GSParamsIndex = GSParamsIndex + 1
 ReDim Preserve GSParams(GSParamsIndex)
 GSParams(GSParamsIndex) = strValue
End Sub

Private Function BuildPermissionString(encData As EncryptData) As String
 Dim strPermissions As String

 strPermissions = vbNullString
 strPermissions = strPermissions & Abs(Int(Not encData.DisallowPrinting))
 strPermissions = strPermissions & Abs(Int(Not encData.DisallowModifyContents))
 strPermissions = strPermissions & Abs(Int(Not encData.DisallowCopy))
 strPermissions = strPermissions & Abs(Int(Not encData.DisallowModifyAnnotations))
 If Options.PDFHighEncryption Then
   strPermissions = strPermissions & Abs(Int(encData.AllowFillIn)) '(128 bit only)
   strPermissions = strPermissions & Abs(Int(encData.AllowScreenReaders)) '(128 bit only)
   strPermissions = strPermissions & Abs(Int(encData.AllowAssembly)) '(128 bit only)
   strPermissions = strPermissions & Abs(Int(encData.AllowDegradedPrinting)) '(128 bit only)
  Else
   strPermissions = strPermissions & "0000"
 End If
 BuildPermissionString = strPermissions
End Function

Public Function EncryptPDF(encData As EncryptData) As Boolean
 Dim strPermissions As String, strShell As String, ret As Double

 strPermissions = BuildPermissionString(encData)

' strShell = App.Path & "\pdfencrypt.exe """ & encData.InputFile & """ """ & encData.OutputFile & """ User=" & encData.UserPass & " Owner=" & encData.OwnerPass & " " & strPermissions & " " & encData.EncryptionLevel
' strShell = CompletePath(Options.DirectoryJava) & "Java.exe -cp """ & CompletePath(App.Path) & "iText.jar"" com.lowagie.tools.encrypt_pdf """ & encData.InputFile & """ """ & encData.OutputFile & """ """ & encData.UserPass & """ """ & encData.OwnerPass & """ " & strPermissions & " " & encData.EncryptionLevel

 strShell = GetPDFCreatorApplicationPath & "pdfenc.exe """ & encData.InputFile & """ """ & encData.OutputFile & """ """ & encData.UserPass & """ """ & encData.OwnerPass & """ " & strPermissions & " " & encData.EncryptionLevel

 IfLoggingWriteLogfile strShell

 ret = RunProgramWait(strShell, False)

 If Dir$(encData.OutputFile) <> vbNullString Then
  EncryptPDF = True
 End If
End Function

Public Function CalculatePermissions(ByRef encData As EncryptData) As Long
 Dim tB As Long, tB2 As Long
 tB = 192
 With encData
  If Abs(.DisallowPrinting) = 0 Then
   tB = tB + 4
  End If
  If Abs(.DisallowModifyContents) = 0 Then
   tB = tB + 8
  End If
  If Abs(.DisallowCopy) = 0 Then
   tB = tB + 16
  End If
  If Abs(.DisallowModifyAnnotations) = 0 Then
   tB = tB + 32
  End If
  CalculatePermissions = tB - 256
  If .EncryptionLevel = encStrong Then
    tB2 = 240
    If Abs(.AllowFillIn) <> 0 Then
     tB2 = tB2 + 1
    End If
    If Abs(.AllowScreenReaders) <> 0 Then
     tB2 = tB2 + 2
    End If
    If Abs(.AllowAssembly) <> 0 Then
     tB2 = tB2 + 4
    End If
    If Abs(.AllowDegradedPrinting) = 0 Then
     tB2 = tB2 + 8
    End If
   CalculatePermissions = CalculatePermissions - (255 - tB2) * 256&
  End If
 End With
End Function

Public Function SetEncryptionParams(ByRef encData As EncryptData, InputFile As String, OutputFile As String) As Boolean
 Dim retPasswd As Boolean

 encData.InputFile = InputFile
 encData.OutputFile = OutputFile

 If Len(Options.PDFOwnerPasswordString) > 0 Then
   encData.OwnerPass = Options.PDFOwnerPasswordString
   OwnerPassword = Options.PDFOwnerPasswordString
   If Options.PDFUserPass = 1 Then
    encData.UserPass = Options.PDFUserPasswordString
    UserPassword = Options.PDFUserPasswordString
   End If
   retPasswd = True
  Else
   If SavePasswordsForThisSession = False Then
     If Options.UseAutosave = 0 Then
       retPasswd = EnterPasswords(encData.UserPass, encData.OwnerPass, frmPassword)
      Else
       retPasswd = False
     End If
    Else
     encData.OwnerPass = OwnerPassword: encData.UserPass = UserPassword
   End If
 End If
 If retPasswd = True Or SavePasswordsForThisSession = True Then
   With encData
    .DisallowPrinting = Options.PDFDisallowPrinting
    .DisallowModifyContents = Options.PDFDisallowModifyContents
    .DisallowCopy = Options.PDFDisallowCopy
    .DisallowModifyAnnotations = Options.PDFDisallowModifyAnnotations
    .AllowFillIn = Options.PDFAllowFillIn
    .AllowScreenReaders = Options.PDFAllowScreenReaders
    .AllowAssembly = Options.PDFAllowAssembly
    .AllowDegradedPrinting = Options.PDFAllowDegradedPrinting
    If Options.PDFHighEncryption = 1 Then
      .EncryptionLevel = encStrong
     Else
      .EncryptionLevel = encLow
    End If
   End With
   SetEncryptionParams = True
   encData.UserPass = UserPassword
   encData.OwnerPass = OwnerPassword
  Else
   SetEncryptionParams = False
 End If
End Function

Private Sub SetColorParams()
 If Options.PDFCompressionColorCompression = 1 Then
   AddParams "-dEncodeColorImages=true"
   If Options.PDFCompressionColorCompressionChoice = 0 Then
     AddParams "-dAutoFilterColorImages=true"
    Else
     AddParams "-dAutoFilterColorImages=false"
     If Options.PDFCompressionColorResample = 1 Then
       AddParams "-dDownsampleColorImages=true"
       Select Case Options.PDFCompressionColorResampleChoice
        Case 0:
         AddParams "-dColorImageDownsampleType=/Subsample"
        Case 1:
         AddParams "-dColorImageDownsampleType=/Average"
        Case 2:
         AddParams "-dColorImageDownsampleType=/Bicubic"
       End Select
       AddParams "-dColorImageResolution=" & Options.PDFCompressionColorResolution
      Else
       AddParams "-dDownsampleColorImages=false"
     End If
     Select Case Options.PDFCompressionColorCompressionChoice
      Case 1:
       AddParams "-dColorImageFilter=/DCTEncode"
       AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
      Case 2:
       AddParams "-dColorImageFilter=/DCTEncode"
       AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
      Case 3:
       AddParams "-dColorImageFilter=/DCTEncode"
       AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
      Case 4:
       AddParams "-dColorImageFilter=/DCTEncode"
       AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, ".") & _
        " /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
      Case 5:
       AddParams "-dColorImageFilter=/DCTEncode"
       AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor " & _
       Replace$(CStr(Options.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, ".") & _
       " /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
      Case 6:
       AddParams "-dColorImageFilter=/FlateEncode"
      Case 7:
       AddParams "-dColorImageFilter=/LZWEncode"
     End Select
   End If
  Else
   AddParams "-dEncodeColorImages=false"
 End If
End Sub

Private Sub SetGreyParams()
 If Options.PDFCompressionGreyCompression = 1 Then
   AddParams "-dEncodeGrayImages=true"
   If Options.PDFCompressionGreyCompressionChoice = 0 Then
     AddParams "-dAutoFilterGrayImages=true"
    Else
     AddParams "-dAutoFilterGrayImages=false"
     If Options.PDFCompressionGreyResample = 1 Then
       AddParams "-dDownsampleGrayImages=true"
       Select Case Options.PDFCompressionGreyResampleChoice
        Case 0:
         AddParams "-dGrayImageDownsampleType=/Subsample"
        Case 1:
         AddParams "-dGrayImageDownsampleType=/Average"
        Case 2:
         AddParams "-dGrayImageDownsampleType=/Bicubic"
       End Select
       AddParams "-dGrayImageResolution=" & Options.PDFCompressionGreyResolution
      Else
       AddParams "-dDownsampleGrayImages=false"
     End If
     Select Case Options.PDFCompressionGreyCompressionChoice
      Case 1:
       AddParams "-dGrayImageFilter=/DCTEncode"
       AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
      Case 2:
       AddParams "-dGrayImageFilter=/DCTEncode"
       AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
      Case 3:
       AddParams "-dGrayImageFilter=/DCTEncode"
       AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
      Case 4:
       AddParams "-dGrayImageFilter=/DCTEncode"
       AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, ".") & _
        " /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
      Case 5:
       AddParams "-dGrayImageFilter=/DCTEncode"
       AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor " & _
       Replace$(CStr(Options.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, ".") & _
       " /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
      Case 6:
       AddParams "-dGrayImageFilter=/FlateEncode"
      Case 7:
       AddParams "-dGrayImageFilter=/LZWEncode"
     End Select
   End If
  Else
   AddParams "-dEncodeGrayImages=false"
 End If
End Sub

Private Sub SetMonoParams()
 If Options.PDFCompressionMonoCompression = 1 Then
   AddParams "-dEncodeMonoImages=true"
   Select Case Options.PDFCompressionMonoCompressionChoice
    Case 0:
     AddParams "-dMonoImageFilter=/CCITTFaxEncode"
    Case 1:
     AddParams "-dMonoImageFilter=/FlateEncode"
    Case 2:
     AddParams "-dMonoImageFilter=/RunLengthEncode"
    Case 3:
     AddParams "-dMonoImageFilter=/LZWEncode"
   End Select
   If Options.PDFCompressionMonoResample = 1 Then
     AddParams "-dDownsampleMonoImages=true"
     Select Case Options.PDFCompressionMonoResampleChoice
      Case 0:
       AddParams "-dMonoImageDownsampleType=/Subsample"
      Case 1:
       AddParams "-dMonoImageDownsampleType=/Average"
      Case 2:
       AddParams "-dMonoImageDownsampleType=/Bicubic"
     End Select
     AddParams "-dMonoImageResolution=" & Options.PDFCompressionMonoResolution
    Else
     AddParams "-dDownsampleMonoImages=false"
   End If
  Else
   AddParams "-dEncodeMonoImages=false"
 End If
End Sub

Public Function GhostScriptSecurity() As Boolean
 GhostScriptSecurity = False
 If LenB(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll) = 0 Then
  Exit Function
 End If
' GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
 If GsDllLoaded = 0 Then
  Exit Function
 End If
 GSRevision = GetGhostscriptRevision
' UnLoadDLL GsDllLoaded
 If InStr(UCase$(GSRevision.strProduct), "AFPL") > 0 Then
  If GSRevision.intRevision < 814 Then
   Exit Function
  End If
  GhostScriptSecurity = True
  Exit Function
 End If
 If InStr(UCase$(GSRevision.strProduct), "GPL") > 0 Then
  If GSRevision.intRevision < 815 Then
   Exit Function
  End If
  GhostScriptSecurity = True
  Exit Function
 End If
End Function

Public Function GetAllGhostscriptversions() As Collection
 Dim reg As clsRegistry, tStr As String, tColl As Collection, i As Long, _
  tf() As String, GS_DLL As String, GS_LIB As String, tB As Boolean, j As Long
 Set reg = New clsRegistry
 Set GetAllGhostscriptversions = New Collection
 With reg
  .hkey = HKEY_LOCAL_MACHINE
  .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
  If .KeyExists = True Then
   tStr = Trim$(.GetRegistryValue("GhostscriptCopyright"))
   If LenB(tStr) > 0 Then
    tStr = Replace$(LanguageStrings.OptionsGhostscriptInternal, "%1", tStr)
    tStr = Replace$(tStr, "%2", Trim$(.GetRegistryValue("GhostscriptVersion")))
    GetAllGhostscriptversions.Add tStr
   End If
  End If
  tStr = "AFPL Ghostscript"
  .KeyRoot = "SOFTWARE\" & tStr
  Set tColl = .EnumRegistryKeys(HKEY_LOCAL_MACHINE, .KeyRoot)
  For i = 1 To tColl.Count
   .Subkey = tColl.Item(i)
   GS_DLL = .GetRegistryValue("GS_DLL")
   GS_LIB = .GetRegistryValue("GS_LIB")
   If Len(GS_DLL) > 0 Then
    If FileExists(GS_DLL) = True Then
     If Len(GS_LIB) > 0 Then
      If InStr(GS_LIB, ";") > 0 Then
       tf = Split(GS_LIB, ";")
       tB = False
       For j = 0 To UBound(tf)
        If DirExists(tf(j)) = False Then
         tB = True
        End If
       Next j
       If tB = False Then
        GetAllGhostscriptversions.Add tStr & " " & tColl.Item(i)
       End If
      End If
     End If
    End If
   End If
  Next i
  tStr = "GNU Ghostscript"
  .KeyRoot = "SOFTWARE\" & tStr
  Set tColl = .EnumRegistryKeys(HKEY_LOCAL_MACHINE, .KeyRoot)
  For i = 1 To tColl.Count
   .Subkey = tColl.Item(i)
   GS_DLL = .GetRegistryValue("GS_DLL")
   GS_LIB = .GetRegistryValue("GS_LIB")
   If Len(GS_DLL) > 0 Then
    If FileExists(GS_DLL) = True Then
     If Len(GS_LIB) > 0 Then
      If InStr(GS_LIB, ";") > 0 Then
       tf = Split(GS_LIB, ";")
       tB = False
       For j = 0 To UBound(tf)
        If DirExists(tf(j)) = False Then
         tB = True
        End If
       Next j
       If tB = False Then
        GetAllGhostscriptversions.Add tStr & " " & tColl.Item(i)
       End If
      End If
     End If
    End If
   End If
  Next i
  tStr = "GPL Ghostscript"
  .KeyRoot = "SOFTWARE\" & tStr
  Set tColl = .EnumRegistryKeys(HKEY_LOCAL_MACHINE, .KeyRoot)
  For i = 1 To tColl.Count
   .Subkey = tColl.Item(i)
   GS_DLL = .GetRegistryValue("GS_DLL")
   GS_LIB = .GetRegistryValue("GS_LIB")
   If Len(GS_DLL) > 0 Then
    If FileExists(GS_DLL) = True Then
     If Len(GS_LIB) > 0 Then
      If InStr(GS_LIB, ";") > 0 Then
       tf = Split(GS_LIB, ";")
       tB = False
       For j = 0 To UBound(tf)
        If DirExists(tf(j)) = False Then
         tB = True
        End If
       Next j
       If tB = False Then
        GetAllGhostscriptversions.Add tStr & " " & tColl.Item(i)
       End If
      End If
     End If
    End If
   End If
  Next i
 End With
 Set reg = Nothing
End Function

Public Sub CheckForStamping(Filename As String)
 Dim StampPage As String, tStr As String, R As String, G As String, B As String, _
  Stampfile As String, Path As String, ff As Long, Files As Collection, _
  StampString As String, StampFontsize As Double, _
  StampOutlineFontthickness As Double
 StampString = RemoveLeadingAndTrailingQuotes(Trim$(Options.StampString))
 If Len(StampString) > 0 Then
  StampPage = StrConv(LoadResData(101, "STAMPPAGE"), vbUnicode)
  StampPage = Replace(StampPage, vbCrLf, vbCr, , , vbBinaryCompare)
  StampPage = Replace(StampPage, "[STAMPSTRING]", EncodeCharsOctal(StampString), , , vbTextCompare)
  StampPage = Replace(StampPage, "[FONTNAME]", Replace(Trim$(Options.StampFontname), " ", ""), , , vbTextCompare)
  StampFontsize = 48
  If IsNumeric(Options.StampFontsize) = True Then
   If CDbl(Options.StampFontsize) > 0 Then
    StampFontsize = CDbl(Options.StampFontsize)
   End If
  End If
  StampPage = Replace(StampPage, "[FONTSIZE]", StampFontsize, , , vbTextCompare)
  StampOutlineFontthickness = 0
  If IsNumeric(Options.StampOutlineFontthickness) = True Then
   If CDbl(Options.StampOutlineFontthickness) >= 0 Then
    StampOutlineFontthickness = CDbl(Options.StampOutlineFontthickness)
   End If
  End If
  StampPage = Replace(StampPage, "[STAMPOUTLINEFONTTHICKNESS]", StampOutlineFontthickness, , , vbTextCompare)
  If Options.StampUseOutlineFont <> 1 Then
    StampPage = Replace(StampPage, "[USEOUTLINEFONT]", "show", , , vbTextCompare)
   Else
    StampPage = Replace(StampPage, "[USEOUTLINEFONT]", "true charpath stroke", , , vbTextCompare)
  End If
  If Len(Options.StampFontColor) > 0 Then
    tStr = Replace$(Options.StampFontColor, "#", "&H")
    If IsNumeric(tStr) = True Then
      R = Replace$(Format(CDbl((CLng(tStr) And CLng("&HFF0000")) / 65536) / 255#, "0.00"), ",", ".", , 1)
      G = Replace$(Format(CDbl((CLng(tStr) And CLng("&H00FF00")) / 256) / 255#, "0.00"), ",", ".", , 1)
      B = Replace$(Format(CDbl(CLng(tStr) And CLng("&H0000FF")) / 255#, "0.00"), ",", ".", , 1)
      StampPage = Replace(StampPage, "[FONTCOLOR]", R & " " & G & " " & B, , , vbTextCompare)
     Else
      StampPage = Replace(StampPage, "[FONTCOLOR]", "1 0 0", , , vbTextCompare)
    End If
   Else
    StampPage = Replace(StampPage, "[FONTCOLOR]", "1 0 0", , , vbTextCompare)
  End If
  Path = CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & GetUsername
  If DirExists(Path) = False Then
   MakePath Path
  End If
  Stampfile = GetTempFile(Path, "~ST")
  ff = FreeFile
  Open Stampfile For Output As #ff
  Print #ff, StampPage
  Close #ff
  Set Files = New Collection
  Files.Add Stampfile
  Files.Add Filename
  Stampfile = GetTempFile(Path, "~ST")
  KillFile Stampfile
  CombineFiles Stampfile, Files
  Name Stampfile As Filename
 End If
End Sub

Private Sub AddParamCommands()
 Dim i As Long
 If ParamCommands.Count > 0 Then
  AddParams "-c"
  For i = 1 To ParamCommands.Count
   AddParams ParamCommands(i)
  Next i
 End If
End Sub

Private Sub AddParamCommand(GhostscriptCommand As String)
 ParamCommands.Add GhostscriptCommand
End Sub

Private Sub AddAdditionalGhostscriptParameters()
 Dim tStr As String, tStrf() As String, i As Long
 tStr = Trim$(Options.AdditionalGhostscriptParameters)
 If LenB(tStr) > 0 Then
  If InStr(1, tStr, "|") > 0 Then
    tStrf = Split(tStr, "|")
    For i = LBound(tStrf) To UBound(tStrf)
     tStr = Trim$(tStrf(i))
     If LenB(tStr) > 0 Then
      AddParams tStr
     End If
    Next i
   Else
    AddParams tStr
  End If
 End If
End Sub

Public Sub CheckForPrintingAfterSaving(GSInputFile As String, Options As tOptions)
 Dim tStr As String

 If Options.PrintAfterSaving = 0 Then
  Exit Sub
 End If

 GSInit Options
 InitParams
 Set ParamCommands = New Collection

 tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
 If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
 End If
 If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
  tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
 End If
 AddParams "-I" & tStr
 AddParams "-q"
 AddParams "-dNOPAUSE"
 AddParams "-dSAFER"
 AddParams "-dBATCH"
 If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
  AddParams "-sFONTPATH=" & GetFontsDirectory
 End If

 AddParams "-sDEVICE=mswinpr2"
 If Options.PrintAfterSavingNoCancel = 1 Then
  AddParams "-dNoCancel"
 End If
 If Options.PrintAfterSavingQueryUser > 0 Then
  AddParams "-QueryUser=" & Options.PrintAfterSavingQueryUser
 End If
 AddParams "-sOutputFile=\\spool\" & Options.PrintAfterSavingPrinter
 If Options.PrintAfterSavingDuplex = 1 Then
  If Options.PrintAfterSavingTumble = 1 Then
    AddParamCommand "<< /Duplex true /Tumble true >> setpagedevice"
   Else
    AddParamCommand "<< /Duplex true /Tumble false >> setpagedevice"
  End If
 End If
 AddParamCommands
 AddParams "-f"
 AddParams GSInputFile
 ShowParams
 CallGhostscript "mswinpr2"
End Sub

Public Sub ConvertPostscriptFile(InputFilename As String, OutputFilename As String)
 Dim Ext As String, Tempfile As String
 IFIsPS = False
 If LenB(InputFilename) = 0 Then
  Exit Sub
 End If
 If FileExists(InputFilename) = False Then
  If LenB(InputFilename) > 0 Then
   MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & _
    "InputFile -IF" & vbCrLf & ">" & InputFilename & "<", vbExclamation + vbOKOnly
  End If
  Exit Sub
 End If

 If LenB(OutputFilename) > 0 Then
   If IsPostscriptFile(InputFilename) = True Then
    If GsDllLoaded = 0 Then
     Exit Sub
    End If
    SplitPath OutputFilename, , , , , Ext
    GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
    If GsDllLoaded = 0 Then
     MsgBox LanguageStrings.MessagesMsg08
    End If
    Select Case UCase$(Ext)
     Case "PDF"
      CallGScript InputFilename, OutputFilename, Options, PDFWriter
     Case "PNG"
      CallGScript InputFilename, OutputFilename, Options, PNGWriter
     Case "JPG"
      CallGScript InputFilename, OutputFilename, Options, JPEGWriter
     Case "BMP"
      CallGScript InputFilename, OutputFilename, Options, BMPWriter
     Case "PCX"
      CallGScript InputFilename, OutputFilename, Options, PCXWriter
     Case "TIF"
      CallGScript InputFilename, OutputFilename, Options, TIFFWriter
     Case "PS"
      CallGScript InputFilename, OutputFilename, Options, PSWriter
     Case "EPS"
      CallGScript InputFilename, OutputFilename, Options, EPSWriter
     Case "TXT"
      CallGScript InputFilename, OutputFilename, Options, TXTWriter
    End Select
   End If
'   If GsDllLoaded <> 0 Then
'    UnloadDLLComplete GsDllLoaded
'   End If
   ConvertedOutputFilename = OutputFilename
   ReadyConverting = True
   Exit Sub
  Else
   If FileExists(InputFilename) = True Then
    If IsPostscriptFile(InputFilename) = True Then
      IFIsPS = True
     Else
      MsgBox LanguageStrings.MessagesMsg06 & vbCrLf & vbCrLf & InputFilename
    End If
   End If
 End If
 DoEvents
End Sub

