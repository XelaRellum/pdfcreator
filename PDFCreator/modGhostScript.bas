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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Rotate(2) As String, Resample(2) As String, Colormodel(2) As String, _
  ColorsPreserveTransfer(1) As String
50030  Dim PNGColorscount(4) As String, JPEGColorscount(1) As String, BMPColorscount(6) As String, _
  PCXColorscount(5) As String, TIFFColorscount(7) As String, _
  PSLanguageLevel(3) As String
50060
50070  Rotate(0) = "None": Rotate(1) = "All": Rotate(2) = "PageByPage"
50080
50090  Resample(0) = "Bicubic": Resample(1) = "Subsample": Resample(2) = "Average"
50100
50110  Colormodel(0) = "RGB": Colormodel(1) = "CMYK": Colormodel(2) = "GRAY"
50120
50130  ColorsPreserveTransfer(0) = "Remove": ColorsPreserveTransfer(1) = "Preserve"
50140
50150  PNGColorscount(0) = "png16m": PNGColorscount(1) = "png256"
50160  PNGColorscount(2) = "png16": PNGColorscount(3) = "png2"
50170  PNGColorscount(4) = "pnggray"
50180
50190  JPEGColorscount(0) = "jpeg": JPEGColorscount(1) = "jpeggray"
50200
50210  BMPColorscount(0) = "bmp32b": BMPColorscount(1) = "bmp16m"
50220  BMPColorscount(2) = "bmp256": BMPColorscount(3) = "bmp16"
50230  BMPColorscount(4) = "bmpsep8": BMPColorscount(5) = "bmpsep1"
50240  BMPColorscount(6) = "bmpgray"
50250
50260  PCXColorscount(0) = "pcxcmyk": PCXColorscount(1) = "pcx24b"
50270  PCXColorscount(2) = "pcx256": PCXColorscount(3) = "pcx16"
50280  PCXColorscount(4) = "pcxmono": PCXColorscount(5) = "pcxgray"
50290
50300  TIFFColorscount(0) = "tiff24nc": TIFFColorscount(1) = "tiff12nc"
50310  TIFFColorscount(2) = "tiffcrle": TIFFColorscount(3) = "tiffg3"
50320  TIFFColorscount(4) = "tiffg32d": TIFFColorscount(5) = "tiffg4"
50330  TIFFColorscount(6) = "tifflzw": TIFFColorscount(7) = "tiffpack"
50340
50350  PSLanguageLevel(0) = "1": PSLanguageLevel(1) = "1.5"
50360  PSLanguageLevel(2) = "2": PSLanguageLevel(3) = "3"
50370
50380 With Options
50390  'General
50400
50410  GS_COMPATIBILITY = "1." & (.PDFGeneralCompatibility + 2)
50420  GS_RESOLUTION = .PDFGeneralResolution
50430  GS_AUTOROTATE = Rotate(.PDFGeneralAutorotate)
50440  GS_OVERPRINT = .PDFGeneralOverprint
50450  GS_ASCII85 = Bool2Text(.PDFGeneralASCII85)
50460
50470  'Compression
50480  GS_COMPRESSPAGES = Bool2Text(.PDFCompressionTextCompression)
50490  GS_COMPRESSCOLOR = Bool2Text(.PDFCompressionColorCompression)
50500  GS_COMPRESSGREY = Bool2Text(.PDFCompressionGreyCompression)
50510  GS_COMPRESSMONO = Bool2Text(.PDFCompressionMonoCompression)
50520
50530  SelectColorCompression .PDFCompressionColorCompressionChoice
50540  SelectGreyCompression .PDFCompressionGreyCompressionChoice
50550  SelectMonoCompression .PDFCompressionMonoCompressionChoice
50560
50570  GS_COMPRESSCOLORVALUE = Bool2Text(.PDFCompressionColorCompression)
50580  GS_COMPRESSGREYVALUE = Bool2Text(.PDFCompressionGreyCompression)
50590  GS_COMPRESSMONOVALUE = Bool2Text(.PDFCompressionMonoCompression)
50600
50610  GS_COLORRESOLUTION = .PDFCompressionColorResolution
50620  GS_GREYRESOLUTION = .PDFCompressionGreyResolution
50630  GS_MONORESOLUTION = .PDFCompressionMonoResolution
50640
50650  GS_COLORRESAMPLE = Bool2Text(.PDFCompressionColorResample)
50660  GS_GREYRESAMPLE = Bool2Text(.PDFCompressionGreyResample)
50670  GS_MONORESAMPLE = Bool2Text(.PDFCompressionMonoResample)
50680
50690  GS_COLORRESAMPLEMETHOD = Resample(.PDFCompressionColorResampleChoice)
50700  GS_GREYRESAMPLEMETHOD = Resample(.PDFCompressionGreyResampleChoice)
50710  GS_MONORESAMPLEMETHOD = Resample(.PDFCompressionMonoResampleChoice)
50720
50730  'Fonts
50740  GS_EMBEDALLFONTS = Bool2Text(.PDFFontsEmbedAll)
50750  GS_SUBSETFONTS = Bool2Text(.PDFFontsSubSetFonts)
50760  GS_SUBSETFONTPERC = .PDFFontsSubSetFontsPercent
50770
50780  'Colors
50790  GS_COLORMODEL = Colormodel(.PDFColorsColorModel)
50800  GS_CMYKTORGB = Bool2Text(.PDFColorsCMYKToRGB)
50810  GS_PRESERVEOVERPRINT = Bool2Text(.PDFColorsPreserveOverprint)
50820  GS_TRANSFERFUNCTIONS = ColorsPreserveTransfer(.PDFColorsPreserveTransfer)
50830  GS_HALFTONE = Bool2Text(.PDFColorsPreserveHalftone)
50840
50850  'Bitmap
50860  GS_BitmapRESOLUTION = .BitmapResolution
50870  GS_PNGColorscount = PNGColorscount(.PNGColorscount)
50880  GS_JPEGColorscount = JPEGColorscount(.JPEGColorscount)
50890  GS_BMPColorscount = BMPColorscount(.BMPColorscount)
50900  GS_PCXColorscount = PCXColorscount(.PCXColorscount)
50910  GS_TIFFColorscount = TIFFColorscount(.TIFFColorscount)
50920  GS_JPEGQuality = .JPEGQuality
50930  GS_PSLanguageLevel = PSLanguageLevel(.PSLanguageLevel)
50940  GS_EPSLanguageLevel = PSLanguageLevel(.EPSLanguageLevel)
50950 End With
50960 'Other
50970 GS_ERROR = 0
50980 UseReturnPipe = 1
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "GSInit")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function CallGScript(GSInputFile As String, GSOutputFile As String, _
 Options As tOptions, Ghostscriptdevice As tGhostscriptDevice)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Dim gsret As Long, i As Long
50020
50030 GSInit Options
50040 InitParams
50050
50060 '50055 MsgBox _
 "Bin: " & Options.DirectoryGhostscriptBinaries & vbCrLf & _
 "Lib: " & Options.DirectoryGhostscriptLibraries & vbCrLf & _
 "Fonts: " & Options.DirectoryGhostscriptFonts
50100 AddParams "-I" & Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50110 AddParams "-q"
50120 AddParams "-dNOPAUSE"
50130 AddParams "-dSAFER"
50140 AddParams "-dBATCH"
50150 Select Case Ghostscriptdevice
 Case 0: 'PDF
50170   AddParams "-sDEVICE=pdfwrite"
50180   AddParams "-dPDFSETTINGS=/printer"
50190   AddParams "-dCompatibilityLevel=" & GS_COMPATIBILITY
50200   AddParams "-r" & GS_RESOLUTION & "x" & GS_RESOLUTION
50210   AddParams "-dProcessColorModel=/Device" & GS_COLORMODEL
50220   AddParams "-dAutoRotatePages=/" & GS_AUTOROTATE
50230   AddParams "-dCompressPages=" & GS_COMPRESSPAGES
50240   AddParams "-dEmbedAllFonts=" & GS_EMBEDALLFONTS
50250   AddParams "-dSubsetFonts=" & GS_SUBSETFONTS
50260   AddParams "-dMaxSubsetPct=" & GS_SUBSETFONTPERC
50270   AddParams "-dConvertCMYKImagesToRGB=" & GS_CMYKTORGB
50280
50290   AddParams "-dAutoFilterColorImages=" & GS_COMPRESSCOLORAUTO
50300   AddParams "-dEncodeColorImages=" & GS_COMPRESSCOLOR
50310   AddParams "-dColorImageFilter=/" & GS_COMPRESSCOLORMETHOD
50320   'AddParams "-dColorImageDict " & GS_COMPRESSCOLORLEVEL
50330   'AddParams "-sColorImageDict=<< /QFactor 0.25 /HSamples [1 1 1 1] /VSamples [1 1 1 1] >>"
50340   'AddParams "-c<< /QFactor 0.9 >>"
50350
50360
50370   AddParams "-dEncodeGrayImages=" & GS_COMPRESSGREY
50380   AddParams "-dGrayImageFilter=/" & GS_COMPRESSGREYMETHOD
50390   'AddParams "-dGrayACSImageDict " & GS_COMPRESSGREYLEVEL
50400
50410   AddParams "-dColorImageResolution=" & GS_COLORRESOLUTION
50420   AddParams "-dGrayImageResolution=" & GS_GREYRESOLUTION
50430   AddParams "-dMonoImageResolution=" & GS_MONORESOLUTION
50440   AddParams "-dDownsampleColorImages=" & GS_COLORRESAMPLE
50450   AddParams "-dDownsampleGrayImages=" & GS_GREYRESAMPLE
50460   AddParams "-dDownsampleMonoImages=" & GS_MONORESAMPLE
50470   AddParams "-dColorImageDownsampleType=/" & GS_COLORRESAMPLEMETHOD
50480   AddParams "-dGrayImageDownsampleType=/" & GS_GREYRESAMPLEMETHOD
50490   AddParams "-dMonoImageDownsampleType=/" & GS_MONORESAMPLEMETHOD
50500   AddParams "-dPreserveOverprintSettings=" & GS_PRESERVEOVERPRINT
50510   AddParams "-dUCRandBGInfo=/Preserve"
50520   AddParams "-dUseFlateCompression=true"
50530   AddParams "-dParseDSCCommentsForDocInfo=true"
50540   AddParams "-dParseDSCComments=true"
50550   AddParams "-dOPM=" & GS_OVERPRINT
50560   AddParams "-dOffOptimizations=0"
50570   AddParams "-dLockDistillerParams=false"
50580   AddParams "-dGrayImageDepth=-1"
50590   AddParams "-dASCII85EncodePages=" & GS_ASCII85
50600   AddParams "-dDefaultRenderingIntent=/Default"
50610   AddParams "-dTransferFunctionInfo=/" & GS_TRANSFERFUNCTIONS
50620   AddParams "-dPreserveHalftoneInfo=" & GS_HALFTONE
50630   AddParams "-dOptimize=true"
50640   AddParams "-dDetectBlends=true"
50650
50660 '  If Options.PDFUseSecurity <> 0 Then
50670 '   Dim Tempfile As String, Temppath As String, encPDF As EncryptData, retEnc As Boolean
50680 '   Tempfile = GetTempFile(GetTempPath, "~PDF")
50690 '   AddParams "-sOutputFile=" & Tempfile
50700 '  Else
50710    AddParams "-sOutputFile=" & GSOutputFile
50720 '  End If
50730   AddParams GSInputFile
50740   AddParams "-c quit"
50750
50760
50770  Case 1: 'PNG
50780   AddParams "-sDEVICE=" & GS_PNGColorscount
50790   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
50800   AddParams "-sOutputFile=" & GSOutputFile
50810   AddParams GSInputFile
50820   AddParams "-c quit"
50830  Case 2: 'JPEG
50840   AddParams "-sDEVICE=" & GS_JPEGColorscount
50850   AddParams "-dJPEGQ=" & GS_JPEGQuality
50860   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
50870   AddParams "-sOutputFile=" & GSOutputFile
50880   AddParams GSInputFile
50890   AddParams "-c quit"
50900  Case 3: 'BMP
50910   AddParams "-sDEVICE=" & GS_BMPColorscount
50920   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
50930   AddParams "-q"
50940   AddParams "-sOutputFile=" & GSOutputFile
50950   AddParams GSInputFile
50960   AddParams "-c quit"
50970  Case 4: 'PCX
50980   AddParams "-sDEVICE=" & GS_PCXColorscount
50990   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
51000   AddParams "-sOutputFile=" & GSOutputFile
51010   AddParams GSInputFile
51020   AddParams "-c quit"
51030  Case 5: 'TIFF
51040   AddParams "-sDEVICE=" & GS_TIFFColorscount
51050   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
51060   AddParams "-sOutputFile=" & GSOutputFile
51070   AddParams GSInputFile
51080   AddParams "-c quit"
51090  Case 6: 'PS
51100   AddParams "-dLanguageLevel=" & GS_PSLanguageLevel
51110   AddParams "-sDEVICE=pswrite"
51120   AddParams "-sOutputFile=" & GSOutputFile
51130   AddParams GSInputFile
51140   AddParams "-c quit"
51150  Case 7: 'EPS
51160   AddParams "-dLanguageLevel=" & GS_EPSLanguageLevel
51170   AddParams "-sDEVICE=epswrite"
51180   AddParams "-sOutputFile=" & GSOutputFile
51190   AddParams GSInputFile
51200   AddParams "-c quit"
51210 End Select
51220
51230 gsret = CallGS(GSParams)
51240
51250 'If (Options.PDFUseSecurity <> 0) And (Ghostscriptdevice = PDFWriter) Then
51260 ' SetEncryptionParams encPDF, Tempfile, GSOutputFile
51270 ' retEnc = EncryptPDF(encPDF)
51280 ' If retEnc = False Then
51290 '   FileCopy Tempfile, GSOutputFile
51300 '   IfLoggingWriteLogfile "Error with encryption - using unencrypted file"
51310 ' End If
51320 ' Kill Tempfile
51330 'End If
51340
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CallGScript")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function OptimizePDF(GSInputFile As String, GSOutputFile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Dim gsret
50020
50030 InitParams
50040
50050 AddParams "-I" & App.Path & "\lib;" & App.Path & "\fonts"
50060 AddParams "-q"
50070 AddParams "-dNODISPLAY"
50080 AddParams "-dSAFER"
50090 AddParams "-dDELAYSAFER"
50100 AddParams "-- pdfopt.ps"
50110 AddParams GSInputFile
50120 AddParams GSOutputFile
50130
50140 frmMain.Refresh
50150 gsret = CallGS(GSParams)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "OptimizePDF")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub ReturnValue(data As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim newData As String
50020  newData = Replace(data, vbLf, "; ")
50030  IfLoggingWriteLogfile "Error: " & newData
50040 ' IfLoggingShowLogfile frmLog, frmMain
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "ReturnValue")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function Bool2Text(number As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If number = 1 Then
50020   Bool2Text = "true"
50030  Else
50040   Bool2Text = "false"
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "Bool2Text")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub SelectColorCompression(ByVal gsMethod)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GS_COMPRESSCOLORAUTO = "false"
50020  Select Case gsMethod
  Case 0
50040    GS_COMPRESSCOLORAUTO = "true"
50050    GS_COMPRESSCOLORMETHOD = "null"
50060    GS_COMPRESSCOLORLEVEL = "null"
50070   Case 1
50080    GS_COMPRESSCOLORMETHOD = "DCTEncode"
50090    GS_COMPRESSCOLORLEVEL = "Maximum"
50100   Case 2
50110    GS_COMPRESSCOLORMETHOD = "DCTEncode"
50120    GS_COMPRESSCOLORLEVEL = "High"
50130   Case 3
50140    GS_COMPRESSCOLORMETHOD = "DCTEncode"
50150    GS_COMPRESSCOLORLEVEL = "Medium"
50160   Case 4
50170    GS_COMPRESSCOLORMETHOD = "DCTEncode"
50180    GS_COMPRESSCOLORLEVEL = "Low"
50190   Case 5
50200    GS_COMPRESSCOLORMETHOD = "DCTEncode"
50210    GS_COMPRESSCOLORLEVEL = "Minimum"
50220   Case 6
50230    GS_COMPRESSCOLORMETHOD = "FlateEncode"
50240    GS_COMPRESSCOLORLEVEL = "Maximum"
50250  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "SelectColorCompression")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SelectGreyCompression(ByVal gsMethod)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GS_COMPRESSGREYAUTO = "false"
50020  Select Case gsMethod
  Case 0
50040    GS_COMPRESSGREYAUTO = "true"
50050    GS_COMPRESSGREYMETHOD = "null"
50060    GS_COMPRESSGREYLEVEL = "null"
50070   Case 1
50080    GS_COMPRESSGREYMETHOD = "DCTEncode"
50090    GS_COMPRESSGREYLEVEL = "Maximum"
50100   Case 2
50110    GS_COMPRESSGREYMETHOD = "DCTEncode"
50120    GS_COMPRESSGREYLEVEL = "High"
50130   Case 3
50140    GS_COMPRESSGREYMETHOD = "DCTEncode"
50150    GS_COMPRESSGREYLEVEL = "Medium"
50160   Case 4
50170    GS_COMPRESSGREYMETHOD = "DCTEncode"
50180    GS_COMPRESSGREYLEVEL = "Low"
50190   Case 5
50200    GS_COMPRESSGREYMETHOD = "DCTEncode"
50210    GS_COMPRESSGREYLEVEL = "Minimum"
50220   Case 6
50230    GS_COMPRESSGREYMETHOD = "FlateEncode"
50240    GS_COMPRESSGREYLEVEL = "Maximum"
50250  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "SelectGreyCompression")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SelectMonoCompression(ByVal gsMethod)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GS_COMPRESSGREYAUTO = "false"
50020  Select Case gsMethod
  Case 0
50040    GS_COMPRESSGREYAUTO = "true"
50050    GS_COMPRESSGREYMETHOD = "null"
50060    GS_COMPRESSGREYLEVEL = "null"
50070   Case 1
50080    GS_COMPRESSGREYMETHOD = "DCTEncode"
50090    GS_COMPRESSGREYLEVEL = "Maximum"
50100   Case 2
50110    GS_COMPRESSGREYMETHOD = "DCTEncode"
50120    GS_COMPRESSGREYLEVEL = "High"
50130   Case 3
50140    GS_COMPRESSGREYMETHOD = "DCTEncode"
50150    GS_COMPRESSGREYLEVEL = "Medium"
50160   Case 4
50170    GS_COMPRESSGREYMETHOD = "DCTEncode"
50180    GS_COMPRESSGREYLEVEL = "Low"
50190   Case 5
50200    GS_COMPRESSGREYMETHOD = "DCTEncode"
50210    GS_COMPRESSGREYLEVEL = "Minimum"
50220   Case 6
50230    GS_COMPRESSGREYMETHOD = "FlateEncode"
50240    GS_COMPRESSGREYLEVEL = "Maximum"
50250  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "SelectMonoCompression")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub InitParams()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GSParamsIndex = 0
50020  ReDim GSParams(GSParamsIndex)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "InitParams")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AddParams(strValue As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GSParamsIndex = GSParamsIndex + 1
50020  ReDim Preserve GSParams(GSParamsIndex)
50030  GSParams(GSParamsIndex) = strValue
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "AddParams")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function EncryptPDF(encData As EncryptData) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strPermissions As String, strShell As String, ret As Double
50020
50030  strPermissions = vbNullString
50040  strPermissions = strPermissions & Abs(Int(Not encData.AllowPrinting))
50050  strPermissions = strPermissions & Abs(Int(Not encData.AllowModifyContents))
50060  strPermissions = strPermissions & Abs(Int(Not encData.AllowCopy))
50070  strPermissions = strPermissions & Abs(Int(Not encData.AllowModifyAnnotations))
50080  If Options.PDFHighEncryption Then
50090    strPermissions = strPermissions & Abs(Int(Not encData.AllowFillIn)) '(128 bit only)
50100    strPermissions = strPermissions & Abs(Int(Not encData.AllowScreenReaders)) '(128 bit only)
50110    strPermissions = strPermissions & Abs(Int(Not encData.AllowAssembly)) '(128 bit only)
50120    strPermissions = strPermissions & Abs(Int(Not encData.AllowDegradedPrinting)) '(128 bit only)
50130   Else
50140    strPermissions = strPermissions & "0000"
50150  End If
50160
50170  strShell = App.Path & "\pdfencrypt.exe """ & encData.InputFile & """ """ & encData.OutputFile & """ User=" & encData.UserPass & " Owner=" & encData.OwnerPass & " " & strPermissions & " " & encData.EncryptionLevel
50180  'IfLoggingWriteLogfile strShell
50190
50200  ret = RunProgramWait(strShell, False)
50210
50220  If Dir$(encData.OutputFile) <> "" Then
50230   EncryptPDF = True
50240  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "EncryptPDF")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub SetEncryptionParams(ByRef encData As EncryptData, InputFile As String, OutputFile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Dim retPasswd As Boolean
50020
50030 encData.InputFile = InputFile
50040 encData.OutputFile = OutputFile
50050
50060 retPasswd = EnterPasswords(encData.UserPass, encData.OwnerPass, frmPassword)
50070
50080 encData.AllowPrinting = Options.PDFAllowPrinting
50090 encData.AllowModifyContents = Options.PDFAllowModifyContents
50100 encData.AllowCopy = Options.PDFAllowCopy
50110 encData.AllowModifyAnnotations = Options.PDFAllowModifyAnnotations
50120 encData.AllowFillIn = Options.PDFAllowFillIn
50130 encData.AllowScreenReaders = Options.PDFAllowScreenReaders
50140 encData.AllowAssembly = Options.PDFAllowAssembly
50150 encData.AllowDegradedPrinting = Options.PDFAllowDegradedPrinting
50160 If Options.PDFHighEncryption = True Then
50170  encData.EncryptionLevel = encStrong
50180 Else
50190  encData.EncryptionLevel = encLow
50200 End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "SetEncryptionParams")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
