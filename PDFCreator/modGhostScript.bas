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
 PDFAWriter = 9
 PDFXWriter = 10
 PSDWriter = 11
 PCLWriter = 12
 RAWWriter = 13
End Enum

Private GSParams() As String
Private GSParamsIndex As Integer

Public GS_ERROR
Public UseReturnPipe

'General
Public GS_PDFDEFAULT
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
Public GS_PNGColorscount
Public GS_JPEGColorscount
Public GS_BMPColorscount
Public GS_PCXColorscount
Public GS_TIFFColorscount
Public GS_JPEGQuality
Public GS_PSDColorscount
Public GS_PCLColorscount
Public GS_RAWColorscount

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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Rotate(2) As String, Resample(2) As String, Colormodel(2) As String, _
  ColorsPreserveTransfer(1) As String
50030  Dim PNGColorscount(4) As String, JPEGColorscount(1) As String, BMPColorscount(6) As String, _
  PCXColorscount(5) As String, TIFFColorscount(7) As String, _
  PSLanguageLevel(3) As String, PSDColorsCount(1) As String, _
  PCLColorsCount(1) As String, RAWColorsCount(2) As String, _
  PDFDefaultSettings(4) As String
50080
50090  PDFDefaultSettings(0) = "default": PDFDefaultSettings(1) = "screen": PDFDefaultSettings(2) = "ebook"
50100  PDFDefaultSettings(3) = "printer": PDFDefaultSettings(4) = "prepress"
50110  Rotate(0) = "None": Rotate(1) = "All": Rotate(2) = "PageByPage"
50120
50130  Resample(0) = "Bicubic": Resample(1) = "Subsample": Resample(2) = "Average"
50140
50150  Colormodel(0) = "RGB": Colormodel(1) = "CMYK": Colormodel(2) = "Gray"
50160
50170  ColorsPreserveTransfer(0) = "Remove": ColorsPreserveTransfer(1) = "Preserve"
50180
50190  PNGColorscount(0) = "png16m": PNGColorscount(1) = "png256"
50200  PNGColorscount(2) = "png16": PNGColorscount(3) = "pngmono"
50210  PNGColorscount(4) = "pnggray"
50220
50230  JPEGColorscount(0) = "jpeg": JPEGColorscount(1) = "jpeggray"
50240
50250  BMPColorscount(0) = "bmp32b": BMPColorscount(1) = "bmp16m"
50260  BMPColorscount(2) = "bmp256": BMPColorscount(3) = "bmp16"
50270  BMPColorscount(4) = "bmpsep8": BMPColorscount(5) = "bmpsep1"
50280  BMPColorscount(6) = "bmpgray"
50290
50300  PCXColorscount(0) = "pcxcmyk": PCXColorscount(1) = "pcx24b"
50310  PCXColorscount(2) = "pcx256": PCXColorscount(3) = "pcx16"
50320  PCXColorscount(4) = "pcxmono": PCXColorscount(5) = "pcxgray"
50330
50340  TIFFColorscount(0) = "tiff24nc": TIFFColorscount(1) = "tiff12nc"
50350  TIFFColorscount(2) = "tiffcrle": TIFFColorscount(3) = "tiffg3"
50360  TIFFColorscount(4) = "tiffg32d": TIFFColorscount(5) = "tiffg4"
50370  TIFFColorscount(6) = "tifflzw": TIFFColorscount(7) = "tiffpack"
50380
50390  PSLanguageLevel(0) = "1": PSLanguageLevel(1) = "1.5"
50400  PSLanguageLevel(2) = "2": PSLanguageLevel(3) = "3"
50410
50420  PSDColorsCount(0) = "psdcmyk": PSDColorsCount(1) = "psdrgb"
50430  PCLColorsCount(0) = "pxlcolor": PCLColorsCount(1) = "pxlmono"
50440  RAWColorsCount(0) = "bitcmyk": RAWColorsCount(1) = "bitrgb": RAWColorsCount(2) = "bit"
50450
50460 With Options
50470  'General
50480  GS_PDFDEFAULT = PDFDefaultSettings(.PDFGeneralDefault)
50490  GS_COMPATIBILITY = "1." & (.PDFGeneralCompatibility + 2)
50500  GS_RESOLUTION = .PDFGeneralResolution
50510  GS_AUTOROTATE = Rotate(.PDFGeneralAutorotate)
50520  GS_OVERPRINT = .PDFGeneralOverprint
50530  GS_ASCII85 = Bool2Text(.PDFGeneralASCII85)
50540
50550  'Compression
50560  GS_COMPRESSPAGES = Bool2Text(.PDFCompressionTextCompression)
50570  GS_COMPRESSCOLOR = Bool2Text(.PDFCompressionColorCompression)
50580  GS_COMPRESSGREY = Bool2Text(.PDFCompressionGreyCompression)
50590  GS_COMPRESSMONO = Bool2Text(.PDFCompressionMonoCompression)
50600
50610  SelectColorCompression .PDFCompressionColorCompressionChoice
50620  SelectGreyCompression .PDFCompressionGreyCompressionChoice
50630  SelectMonoCompression .PDFCompressionMonoCompressionChoice
50640
50650  GS_COMPRESSCOLORVALUE = Bool2Text(.PDFCompressionColorCompression)
50660  GS_COMPRESSGREYVALUE = Bool2Text(.PDFCompressionGreyCompression)
50670  GS_COMPRESSMONOVALUE = Bool2Text(.PDFCompressionMonoCompression)
50680
50690  GS_COLORRESOLUTION = .PDFCompressionColorResolution
50700  GS_GREYRESOLUTION = .PDFCompressionGreyResolution
50710  GS_MONORESOLUTION = .PDFCompressionMonoResolution
50720
50730  GS_COLORRESAMPLE = Bool2Text(.PDFCompressionColorResample)
50740  GS_GREYRESAMPLE = Bool2Text(.PDFCompressionGreyResample)
50750  GS_MONORESAMPLE = Bool2Text(.PDFCompressionMonoResample)
50760
50770  GS_COLORRESAMPLEMETHOD = Resample(.PDFCompressionColorResampleChoice)
50780  GS_GREYRESAMPLEMETHOD = Resample(.PDFCompressionGreyResampleChoice)
50790  GS_MONORESAMPLEMETHOD = Resample(.PDFCompressionMonoResampleChoice)
50800
50810  'Fonts
50820  GS_EMBEDALLFONTS = Bool2Text(.PDFFontsEmbedAll)
50830  GS_SUBSETFONTS = Bool2Text(.PDFFontsSubSetFonts)
50840  GS_SUBSETFONTPERC = .PDFFontsSubSetFontsPercent
50850
50860  'Colors
50870  GS_COLORMODEL = Colormodel(.PDFColorsColorModel)
50880  GS_CMYKTORGB = Bool2Text(.PDFColorsCMYKToRGB)
50890  GS_PRESERVEOVERPRINT = Bool2Text(.PDFColorsPreserveOverprint)
50900  GS_TRANSFERFUNCTIONS = ColorsPreserveTransfer(.PDFColorsPreserveTransfer)
50910  GS_HALFTONE = Bool2Text(.PDFColorsPreserveHalftone)
50920
50930  'Bitmap
50940  GS_PNGColorscount = PNGColorscount(.PNGColorscount)
50950  GS_JPEGColorscount = JPEGColorscount(.JPEGColorscount)
50960  GS_BMPColorscount = BMPColorscount(.BMPColorscount)
50970  GS_PCXColorscount = PCXColorscount(.PCXColorscount)
50980  GS_TIFFColorscount = TIFFColorscount(.TIFFColorscount)
50990  GS_JPEGQuality = .JPEGQuality
51000  GS_PSLanguageLevel = PSLanguageLevel(.PSLanguageLevel)
51010  GS_EPSLanguageLevel = PSLanguageLevel(.EPSLanguageLevel)
51020  GS_PSDColorscount = PSDColorsCount(.PSDColorsCount)
51030  GS_PCLColorscount = PCLColorsCount(.PCLColorsCount)
51040  GS_RAWColorscount = RAWColorsCount(.RAWColorsCount)
51050 End With
51060 'Other
51070 GS_ERROR = 0
51080 UseReturnPipe = 1
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

Private Function CallGhostscript(Comment As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim LastStop As Currency, res As Boolean, c As Currency
50020  If PerformanceTimer Then
50030   LastStop = ExactTimer_Value()
50040  End If
50050  res = CallGS(GSParams)
50060  If PerformanceTimer Then
50070    c = ExactTimer_Value() - LastStop
50080    IfLoggingWriteLogfile "Time for converting [" & Comment & "]: " & _
    Format$(Int(c) * (1 / 86400), "hh:nn:ss:") & Format$(((c) - Int(c)) * 1000, "000")
50100   Else
50110    IfLoggingWriteLogfile "Time for converting -> No performance timer [" & Comment & "]"
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CallGhostscript")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePDF(GSInputFile As String, GSOutputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim FName As String, tStr As String, encPDF As EncryptData, tEnc As Boolean
50020
50030  InitParams
50040  Set ParamCommands = New Collection
50050
50060  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50070  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50080   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50090  End If
50100  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50110   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50120  End If
50130  AddParams "-I" & tStr
50140  AddParams "-q"
50150  AddParams "-dNOPAUSE"
50160  AddParams "-dSAFER"
50170  AddParams "-dBATCH"
50180  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50190   AddParams "-sFONTPATH=" & GetFontsDirectory
50200  End If
50210  AddParams "-sDEVICE=pdfwrite"
50220  If Options.DontUseDocumentSettings = 0 Then
50230   AddParams "-dPDFSETTINGS=/" & GS_PDFDEFAULT
50240   AddParams "-dCompatibilityLevel=" & GS_COMPATIBILITY
50250   AddParams "-r" & GS_RESOLUTION & "x" & GS_RESOLUTION
50260   AddParams "-dProcessColorModel=/Device" & GS_COLORMODEL
50270   AddParams "-dAutoRotatePages=/" & GS_AUTOROTATE
50280   AddParams "-dCompressPages=" & GS_COMPRESSPAGES
50290   AddParams "-dEmbedAllFonts=" & GS_EMBEDALLFONTS
50300   AddParams "-dSubsetFonts=" & GS_SUBSETFONTS
50310   AddParams "-dMaxSubsetPct=" & GS_SUBSETFONTPERC
50320   AddParams "-dConvertCMYKImagesToRGB=" & GS_CMYKTORGB
50330
50340   If Options.UseFixPapersize <> 0 Then
50350    If Options.UseCustomPaperSize = 0 Then
50360      If LenB(Trim$(Options.Papersize)) > 0 Then
50370       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50380       AddParams "-dFIXEDMEDIA"
50390       AddParams "-dNORANGEPAGESIZE"
50400      End If
50410     Else
50420      If Options.DeviceWidthPoints >= 1 Then
50430       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50440      End If
50450      If Options.DeviceHeightPoints >= 1 Then
50460       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50470      End If
50480    End If
50490   End If
50500
50510  End If
50520  tEnc = False
50530  If Options.PDFOptimize = 0 And Options.PDFUseSecurity <> 0 And SecurityIsPossible = True And Options.PDFEncryptor = 0 Then
50540   If SetEncryptionParams(encPDF, "", "") = True Then
50550     tEnc = True
50560     If Len(encPDF.OwnerPass) > 0 Then
50570       AddParams "-sOwnerPassword=" & encPDF.OwnerPass
50580      Else
50590       If Len(encPDF.UserPass) > 0 Then
50600        AddParams "-sOwnerPassword=" & encPDF.UserPass
50610       End If
50620     End If
50630     If Len(encPDF.UserPass) > 0 Then
50640      AddParams "-sUserPassword=" & encPDF.UserPass
50650     End If
50660     AddParams "-dPermissions=" & CalculatePermissions(encPDF)
50670     If GS_COMPATIBILITY = "1.4" Then
50680       AddParams "-dEncryptionR=3"
50690      Else
50700       AddParams "-dEncryptionR=2"
50710     End If
50720     If encPDF.EncryptionLevel = encLow Then
50730       AddParams "-dKeyLength=40"
50740      Else
50750       AddParams "-dKeyLength=128"
50760     End If
50770    Else
50780     If Options.UseAutosave = 0 Then
50790      MsgBox LanguageStrings.MessagesMsg23, vbCritical
50800     End If
50810   End If
50820  End If
50830  AddParams "-sOutputFile=" & GSOutputFile
50840
50850  If Options.DontUseDocumentSettings = 0 Then
50860   SetColorParams
50870   SetGreyParams
50880   SetMonoParams
50890
50900   AddParams "-dPreserveOverprintSettings=" & GS_PRESERVEOVERPRINT
50910   AddParams "-dUCRandBGInfo=/Preserve"
50920   AddParams "-dUseFlateCompression=true"
50930   AddParams "-dParseDSCCommentsForDocInfo=true"
50940   AddParams "-dParseDSCComments=true"
50950   AddParams "-dOPM=" & GS_OVERPRINT
50960   AddParams "-dOffOptimizations=0"
50970   AddParams "-dLockDistillerParams=false"
50980   AddParams "-dGrayImageDepth=-1"
50990   AddParams "-dASCII85EncodePages=" & GS_ASCII85
51000   AddParams "-dDefaultRenderingIntent=/Default"
51010   AddParams "-dTransferFunctionInfo=/" & GS_TRANSFERFUNCTIONS
51020   AddParams "-dPreserveHalftoneInfo=" & GS_HALFTONE
51030   AddParams "-dDetectBlends=true"
51040
51050   AddAdditionalGhostscriptParameters
51060
51070   AddParamCommands
51080  End If
51090
51100  AddParams "-f"
51110  AddParams GSInputFile
51120  ShowParams
51130  If tEnc = True Then
51140    CallGhostscript "PDF with encryption"
51150   Else
51160    CallGhostscript "PDF without encryption"
51170  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreatePDF")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePNG(GSInputFile As String, GSOutputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, FName As String, Ext As String, tStr As String
50020
50030  GSInit Options
50040  InitParams
50050  Set ParamCommands = New Collection
50060
50070  If Options.OnePagePerFile = 1 Then
50080   SplitPath GSOutputFile, , Path, , FName, Ext
50090   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50100  End If
50110  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50120  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50130   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50140  End If
50150  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50160   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50170  End If
50180  AddParams "-I" & tStr
50190  AddParams "-q"
50200  AddParams "-dNOPAUSE"
50210  AddParams "-dSAFER"
50220  AddParams "-dBATCH"
50230  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50240   AddParams "-sFONTPATH=" & GetFontsDirectory
50250  End If
50260
50270  AddParams "-sDEVICE=" & GS_PNGColorscount
50280
50290  If Options.DontUseDocumentSettings = 0 Then
50300   If Options.UseFixPapersize <> 0 Then
50310    If Options.UseCustomPaperSize = 0 Then
50320      If LenB(Trim$(Options.Papersize)) > 0 Then
50330       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50340       AddParams "-dFIXEDMEDIA"
50350       AddParams "-dNORANGEPAGESIZE"
50360      End If
50370     Else
50380      If Options.DeviceWidthPoints >= 1 Then
50390       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50400      End If
50410      If Options.DeviceHeightPoints >= 1 Then
50420       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50430      End If
50440    End If
50450   End If
50460   AddParams "-r" & Options.PNGResolution & "x" & Options.PNGResolution
50470   AddParams "-sOutputFile=" & GSOutputFile
50480  End If
50490
50500  AddAdditionalGhostscriptParameters
50510
50520  AddParams "-f"
50530  AddParams GSInputFile
50540  ShowParams
50550  CallGhostscript "PNG"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreatePNG")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateJPEG(GSInputFile As String, GSOutputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, FName As String, Ext As String, tStr As String
50020
50030  GSInit Options
50040  InitParams
50050  Set ParamCommands = New Collection
50060
50070  If Options.OnePagePerFile = 1 Then
50080   SplitPath GSOutputFile, , Path, , FName, Ext
50090   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50100  End If
50110  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50120  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50130   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50140  End If
50150  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50160   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50170  End If
50180  AddParams "-I" & tStr
50190  AddParams "-q"
50200  AddParams "-dNOPAUSE"
50210  AddParams "-dSAFER"
50220  AddParams "-dBATCH"
50230  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50240   AddParams "-sFONTPATH=" & GetFontsDirectory
50250  End If
50260
50270  AddParams "-sDEVICE=" & GS_JPEGColorscount
50280  If Options.DontUseDocumentSettings = 0 Then
50290   AddParams "-dJPEGQ=" & GS_JPEGQuality
50300   If Options.UseFixPapersize <> 0 Then
50310    If Options.UseCustomPaperSize = 0 Then
50320      If LenB(Trim$(Options.Papersize)) > 0 Then
50330       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50340       AddParams "-dFIXEDMEDIA"
50350       AddParams "-dNORANGEPAGESIZE"
50360      End If
50370     Else
50380      If Options.DeviceWidthPoints >= 1 Then
50390       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50400      End If
50410      If Options.DeviceHeightPoints >= 1 Then
50420       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50430      End If
50440    End If
50450   End If
50460   AddParams "-r" & Options.JPEGResolution & "x" & Options.JPEGResolution
50470   AddParams "-sOutputFile=" & GSOutputFile
50480  End If
50490
50500  AddAdditionalGhostscriptParameters
50510
50520  AddParams "-f"
50530  AddParams GSInputFile
50540  ShowParams
50550  CallGhostscript "JPEG"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreateJPEG")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateBMP(GSInputFile As String, GSOutputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, FName As String, Ext As String, tStr As String
50020
50030  GSInit Options
50040  InitParams
50050  Set ParamCommands = New Collection
50060
50070  If Options.OnePagePerFile = 1 Then
50080   SplitPath GSOutputFile, , Path, , FName, Ext
50090   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50100  End If
50110  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50120  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50130   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50140  End If
50150  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50160   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50170  End If
50180  AddParams "-I" & tStr
50190  AddParams "-q"
50200  AddParams "-dNOPAUSE"
50210  AddParams "-dSAFER"
50220  AddParams "-dBATCH"
50230  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50240   AddParams "-sFONTPATH=" & GetFontsDirectory
50250  End If
50260
50270  AddParams "-sDEVICE=" & GS_BMPColorscount
50280
50290  If Options.DontUseDocumentSettings = 0 Then
50300   If Options.UseFixPapersize <> 0 Then
50310    If Options.UseCustomPaperSize = 0 Then
50320      If LenB(Trim$(Options.Papersize)) > 0 Then
50330       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50340       AddParams "-dFIXEDMEDIA"
50350       AddParams "-dNORANGEPAGESIZE"
50360      End If
50370     Else
50380      If Options.DeviceWidthPoints >= 1 Then
50390       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50400      End If
50410      If Options.DeviceHeightPoints >= 1 Then
50420       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50430      End If
50440    End If
50450   End If
50460   AddParams "-r" & Options.BMPResolution & "x" & Options.BMPResolution
50470  End If
50480  AddParams "-sOutputFile=" & GSOutputFile
50490
50500  AddAdditionalGhostscriptParameters
50510
50520  AddParams "-f"
50530  AddParams GSInputFile
50540  ShowParams
50550  CallGhostscript "BMP"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreateBMP")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePCX(GSInputFile As String, GSOutputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, FName As String, Ext As String, tStr As String
50020
50030  GSInit Options
50040  InitParams
50050  Set ParamCommands = New Collection
50060
50070  If Options.OnePagePerFile = 1 Then
50080   SplitPath GSOutputFile, , Path, , FName, Ext
50090   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50100  End If
50110  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50120  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50130   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50140  End If
50150  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50160   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50170  End If
50180  AddParams "-I" & tStr
50190  AddParams "-q"
50200  AddParams "-dNOPAUSE"
50210  AddParams "-dSAFER"
50220  AddParams "-dBATCH"
50230  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50240   AddParams "-sFONTPATH=" & GetFontsDirectory
50250  End If
50260
50270  AddParams "-sDEVICE=" & GS_PCXColorscount
50280  If Options.DontUseDocumentSettings = 0 Then
50290   If Options.UseFixPapersize <> 0 Then
50300    If Options.UseCustomPaperSize = 0 Then
50310      If LenB(Trim$(Options.Papersize)) > 0 Then
50320       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50330       AddParams "-dFIXEDMEDIA"
50340       AddParams "-dNORANGEPAGESIZE"
50350      End If
50360     Else
50370      If Options.DeviceWidthPoints >= 1 Then
50380       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50390      End If
50400      If Options.DeviceHeightPoints >= 1 Then
50410       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50420      End If
50430    End If
50440   End If
50450   AddParams "-r" & Options.PCXResolution & "x" & Options.PCXResolution
50460  End If
50470  AddParams "-sOutputFile=" & GSOutputFile
50480
50490  AddAdditionalGhostscriptParameters
50500
50510  AddParams "-f"
50520  AddParams GSInputFile
50530  ShowParams
50540  CallGhostscript "PCX"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreatePCX")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateTIFF(GSInputFile As String, GSOutputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, FName As String, Ext As String, tStr As String
50020
50030  GSInit Options
50040  InitParams
50050  Set ParamCommands = New Collection
50060
50070  If Options.OnePagePerFile = 1 Then
50080   SplitPath GSOutputFile, , Path, , FName, Ext
50090   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50100  End If
50110  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50120  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50130   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50140  End If
50150  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50160   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50170  End If
50180  AddParams "-I" & tStr
50190  AddParams "-q"
50200  AddParams "-dNOPAUSE"
50210  AddParams "-dSAFER"
50220  AddParams "-dBATCH"
50230  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50240   AddParams "-sFONTPATH=" & GetFontsDirectory
50250  End If
50260
50270  AddParams "-sDEVICE=" & GS_TIFFColorscount
50280  If Options.DontUseDocumentSettings = 0 Then
50290   If Options.UseFixPapersize <> 0 Then
50300    If Options.UseCustomPaperSize = 0 Then
50310      If LenB(Trim$(Options.Papersize)) > 0 Then
50320       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50330       AddParams "-dFIXEDMEDIA"
50340       AddParams "-dNORANGEPAGESIZE"
50350      End If
50360     Else
50370      If Options.DeviceWidthPoints >= 1 Then
50380       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50390      End If
50400      If Options.DeviceHeightPoints >= 1 Then
50410       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50420      End If
50430    End If
50440   End If
50450   AddParams "-r" & Options.TIFFResolution & "x" & Options.TIFFResolution
50460  End If
50470  AddParams "-sOutputFile=" & GSOutputFile
50480
50490  AddAdditionalGhostscriptParameters
50500
50510  AddParams "-f"
50520  AddParams GSInputFile
50530  ShowParams
50540  CallGhostscript "TIFF"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreateTIFF")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePS(GSInputFile As String, GSOutputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, FName As String, Ext As String, tStr As String
50020
50030  GSInit Options
50040  InitParams
50050  Set ParamCommands = New Collection
50060
50070  If Options.OnePagePerFile = 1 Then
50080   SplitPath GSOutputFile, , Path, , FName, Ext
50090   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50100  End If
50110  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50120  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50130   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50140  End If
50150  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50160   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50170  End If
50180  AddParams "-I" & tStr
50190  AddParams "-q"
50200  AddParams "-dNOPAUSE"
50210  AddParams "-dSAFER"
50220  AddParams "-dBATCH"
50230  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50240   AddParams "-sFONTPATH=" & GetFontsDirectory
50250  End If
50260
50270  AddParams "-sDEVICE=pswrite"
50280  If Options.DontUseDocumentSettings = 0 Then
50290   If Options.UseFixPapersize <> 0 Then
50300    If Options.UseCustomPaperSize = 0 Then
50310      If LenB(Trim$(Options.Papersize)) > 0 Then
50320       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50330       AddParams "-dFIXEDMEDIA"
50340       AddParams "-dNORANGEPAGESIZE"
50350      End If
50360     Else
50370      If Options.DeviceWidthPoints >= 1 Then
50380       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50390      End If
50400      If Options.DeviceHeightPoints >= 1 Then
50410       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50420      End If
50430    End If
50440   End If
50450   AddParams "-dLanguageLevel=" & GS_PSLanguageLevel
50460  End If
50470  AddParams "-sOutputFile=" & GSOutputFile
50480
50490  AddAdditionalGhostscriptParameters
50500
50510  AddParams "-f"
50520  AddParams GSInputFile
50530  ShowParams
50540  CallGhostscript "PS"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreatePS")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateEPS(GSInputFile As String, GSOutputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, FName As String, Ext As String, tStr As String
50020
50030  GSInit Options
50040  InitParams
50050  Set ParamCommands = New Collection
50060
50070  If Options.OnePagePerFile = 1 Then
50080   SplitPath GSOutputFile, , Path, , FName, Ext
50090   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50100  End If
50110  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50120  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50130   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50140  End If
50150  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50160   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50170  End If
50180  AddParams "-I" & tStr
50190  AddParams "-q"
50200  AddParams "-dNOPAUSE"
50210  AddParams "-dSAFER"
50220  AddParams "-dBATCH"
50230  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50240   AddParams "-sFONTPATH=" & GetFontsDirectory
50250  End If
50260
50270  AddParams "-sDEVICE=epswrite"
50280  If Options.DontUseDocumentSettings = 0 Then
50290   If Options.UseFixPapersize <> 0 Then
50300    If Options.UseCustomPaperSize = 0 Then
50310      If LenB(Trim$(Options.Papersize)) > 0 Then
50320       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50330       AddParams "-dFIXEDMEDIA"
50340       AddParams "-dNORANGEPAGESIZE"
50350      End If
50360     Else
50370      If Options.DeviceWidthPoints >= 1 Then
50380       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50390      End If
50400      If Options.DeviceHeightPoints >= 1 Then
50410       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50420      End If
50430    End If
50440   End If
50450   AddParams "-dLanguageLevel=" & GS_EPSLanguageLevel
50460  End If
50470  AddParams "-sOutputFile=" & GSOutputFile
50480
50490  AddAdditionalGhostscriptParameters
50500
50510  AddParams "-f"
50520  AddParams GSInputFile
50530  ShowParams
50540  CallGhostscript "EPS"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreateEPS")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateTXT(GSInputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, FName As String, Ext As String, tStr As String
50020
50030  GSInit Options
50040  InitParams
50050  Set ParamCommands = New Collection
50060
50070  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50080  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50090   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50100  End If
50110  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50120   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50130  End If
50140  AddParams "-I" & tStr
50150  AddParams "-q"
50160  AddParams "-dNOPAUSE"
50170  AddParams "-dSAFER"
50180  AddParams "-dBATCH"
50190  AddParams "-dNODISPLAY"
50200  AddParams "-dDELAYBIND"
50210  AddParams "-dWRITESYSTEMDICT"
50220  AddParams "-dSIMPLE"
50230  AddParams "ps2ascii.ps"
50240  AddParams GSInputFile
50250  ShowParams
50260  CallGhostscript "TXT"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreateTXT")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePDFA(GSInputFile As String, GSOutputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim FName As String, tStr As String, encPDF As EncryptData, tEnc As Boolean
50020
50030  InitParams
50040  Set ParamCommands = New Collection
50050
50060  tStr = Options.DirectoryGhostscriptLibraries
50070  If LenB(LTrim(Options.DirectoryGhostscriptFonts)) > 0 Then
50080   If DirExists(Options.DirectoryGhostscriptFonts) Then
50090    tStr = tStr & ";" & Options.DirectoryGhostscriptFonts
50100   End If
50110  End If
50120  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50130   If DirExists(Options.DirectoryGhostscriptResource) Then
50140    tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50150   End If
50160  End If
50170  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50180   If DirExists(Options.AdditionalGhostscriptSearchpath) Then
50190    tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50200   End If
50210  End If
50220  AddParams "-I" & tStr
50230  AddParams "-dPDFA"
50240  AddParams "-q"
50250  AddParams "-dNOPAUSE"
50260  AddParams "-dSAFER"
50270  AddParams "-dBATCH"
50280  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50290   AddParams "-sFONTPATH=" & GetFontsDirectory
50300  End If
50310  AddParams "-sDEVICE=pdfwrite"
50320  If Options.DontUseDocumentSettings = 0 Then
50330   AddParams "-dPDFSETTINGS=/" & GS_PDFDEFAULT
50340   AddParams "-dCompatibilityLevel=" & GS_COMPATIBILITY
50350   AddParams "-dNOOUTERSAVE"
50360   AddParams "-dUseCIEColor"
50370   AddParams "-r" & GS_RESOLUTION & "x" & GS_RESOLUTION
50380   AddParams "-sProcessColorModel=Device" & GS_COLORMODEL
50390   AddParams "-dAutoRotatePages=/" & GS_AUTOROTATE
50400   AddParams "-dCompressPages=" & GS_COMPRESSPAGES
50410   AddParams "-dEmbedAllFonts=true"
50420   AddParams "-dSubsetFonts=" & GS_SUBSETFONTS
50430   AddParams "-dMaxSubsetPct=100"
50440   AddParams "-dConvertCMYKImagesToRGB=" & GS_CMYKTORGB
50450
50460   If Options.UseFixPapersize <> 0 Then
50470    If Options.UseCustomPaperSize = 0 Then
50480      If LenB(Trim$(Options.Papersize)) > 0 Then
50490       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50500       AddParams "-dFIXEDMEDIA"
50510       AddParams "-dNORANGEPAGESIZE"
50520      End If
50530     Else
50540      If Options.DeviceWidthPoints >= 1 Then
50550       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50560      End If
50570      If Options.DeviceHeightPoints >= 1 Then
50580       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50590      End If
50600    End If
50610   End If
50620
50630  End If
50640  tEnc = False
50650  AddParams "-sOutputFile=" & GSOutputFile
50660
50670  If Options.DontUseDocumentSettings = 0 Then
50680   SetColorParams
50690   SetGreyParams
50700   SetMonoParams
50710
50720   AddParams "-dPreserveOverprintSettings=" & GS_PRESERVEOVERPRINT
50730   AddParams "-dUCRandBGInfo=/Preserve"
50740   AddParams "-dUseFlateCompression=true"
50750   AddParams "-dParseDSCCommentsForDocInfo=true"
50760   AddParams "-dParseDSCComments=true"
50770   AddParams "-dOPM=1" '& GS_OVERPRINT
50780   AddParams "-dOffOptimizations=0"
50790   AddParams "-dLockDistillerParams=false"
50800   AddParams "-dGrayImageDepth=-1"
50810   AddParams "-dASCII85EncodePages=" & GS_ASCII85
50820   AddParams "-dDefaultRenderingIntent=/Default"
50830   AddParams "-dTransferFunctionInfo=/" & GS_TRANSFERFUNCTIONS
50840   AddParams "-dPreserveHalftoneInfo=" & GS_HALFTONE
50850   AddParams "-dDetectBlends=true"
50860
50870   AddAdditionalGhostscriptParameters
50880
50890   AddParams "-f"
50900   AddParams CompletePath(Options.DirectoryGhostscriptLibraries) + "pdfa_def.ps"
50910
50920   AddParamCommands
50930  End If
50940
50950  AddParams "-f"
50960  AddParams GSInputFile
50970  ShowParams
50980  CallGhostscript "PDF/A (without encryption)"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreatePDFA")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePDFX(GSInputFile As String, GSOutputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim FName As String, tStr As String, encPDF As EncryptData, tEnc As Boolean
50020
50030  InitParams
50040  Set ParamCommands = New Collection
50050
50060  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50070  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50080   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50090  End If
50100  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50110   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50120  End If
50130  AddParams "-I" & tStr
50140  AddParams "-q"
50150  AddParams "-dNOPAUSE"
50160  AddParams "-dSAFER"
50170  AddParams "-dBATCH"
50180  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50190   AddParams "-sFONTPATH=" & GetFontsDirectory
50200  End If
50210  AddParams "-sDEVICE=pdfwrite"
50220  If Options.DontUseDocumentSettings = 0 Then
50230   AddParams "-dPDFSETTINGS=/" & GS_PDFDEFAULT
50240   AddParams "-dCompatibilityLevel=" & GS_COMPATIBILITY
50250   AddParams "-r" & GS_RESOLUTION & "x" & GS_RESOLUTION
50260   AddParams "-dProcessColorModel=/Device" & GS_COLORMODEL
50270   AddParams "-dAutoRotatePages=/" & GS_AUTOROTATE
50280   AddParams "-dCompressPages=" & GS_COMPRESSPAGES
50290   AddParams "-dEmbedAllFonts=" & GS_EMBEDALLFONTS
50300   AddParams "-dSubsetFonts=" & GS_SUBSETFONTS
50310   AddParams "-dMaxSubsetPct=" & GS_SUBSETFONTPERC
50320   AddParams "-dConvertCMYKImagesToRGB=" & GS_CMYKTORGB
50330
50340   If Options.UseFixPapersize <> 0 Then
50350    If Options.UseCustomPaperSize = 0 Then
50360      If LenB(Trim$(Options.Papersize)) > 0 Then
50370       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50380       AddParams "-dFIXEDMEDIA"
50390       AddParams "-dNORANGEPAGESIZE"
50400      End If
50410     Else
50420      If Options.DeviceWidthPoints >= 1 Then
50430       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50440      End If
50450      If Options.DeviceHeightPoints >= 1 Then
50460       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50470      End If
50480    End If
50490   End If
50500
50510  End If
50520  tEnc = False
50530  AddParams "-sOutputFile=" & GSOutputFile
50540
50550  If Options.DontUseDocumentSettings = 0 Then
50560   SetColorParams
50570   SetGreyParams
50580   SetMonoParams
50590
50600   AddParams "-dPreserveOverprintSettings=" & GS_PRESERVEOVERPRINT
50610   AddParams "-dUCRandBGInfo=/Preserve"
50620   AddParams "-dUseFlateCompression=true"
50630   AddParams "-dParseDSCCommentsForDocInfo=true"
50640   AddParams "-dParseDSCComments=true"
50650   AddParams "-dOPM=" & GS_OVERPRINT
50660   AddParams "-dOffOptimizations=0"
50670   AddParams "-dLockDistillerParams=false"
50680   AddParams "-dGrayImageDepth=-1"
50690   AddParams "-dASCII85EncodePages=" & GS_ASCII85
50700   AddParams "-dDefaultRenderingIntent=/Default"
50710   AddParams "-dTransferFunctionInfo=/" & GS_TRANSFERFUNCTIONS
50720   AddParams "-dPreserveHalftoneInfo=" & GS_HALFTONE
50730   AddParams "-dDetectBlends=true"
50740
50750   AddAdditionalGhostscriptParameters
50760
50770   AddParams "-dPDFX"
50780   AddParams "-f"
50790   AddParams CompletePath(Options.DirectoryGhostscriptLibraries) + "pdfx_def.ps"
50800
50810   AddParamCommands
50820  End If
50830
50840  AddParams "-f"
50850  AddParams GSInputFile
50860  ShowParams
50870  CallGhostscript "PDF/X (without encryption)"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreatePDFX")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePSD(GSInputFile As String, GSOutputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, FName As String, Ext As String, tStr As String
50020
50030  GSInit Options
50040  InitParams
50050  Set ParamCommands = New Collection
50060
50070  If Options.OnePagePerFile = 1 Then
50080   SplitPath GSOutputFile, , Path, , FName, Ext
50090   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50100  End If
50110  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50120  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50130   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50140  End If
50150  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50160   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50170  End If
50180  AddParams "-I" & tStr
50190  AddParams "-q"
50200  AddParams "-dNOPAUSE"
50210  AddParams "-dSAFER"
50220  AddParams "-dBATCH"
50230  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50240   AddParams "-sFONTPATH=" & GetFontsDirectory
50250  End If
50260
50270  AddParams "-sDEVICE=" & GS_PSDColorscount
50280  If Options.DontUseDocumentSettings = 0 Then
50290   If Options.UseFixPapersize <> 0 Then
50300    If Options.UseCustomPaperSize = 0 Then
50310      If LenB(Trim$(Options.Papersize)) > 0 Then
50320       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50330       AddParams "-dFIXEDMEDIA"
50340       AddParams "-dNORANGEPAGESIZE"
50350      End If
50360     Else
50370      If Options.DeviceWidthPoints >= 1 Then
50380       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50390      End If
50400      If Options.DeviceHeightPoints >= 1 Then
50410       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50420      End If
50430    End If
50440   End If
50450   AddParams "-r" & Options.PSDResolution & "x" & Options.PSDResolution
50460  End If
50470  AddParams "-sOutputFile=" & GSOutputFile
50480
50490  AddAdditionalGhostscriptParameters
50500
50510  AddParams "-f"
50520  AddParams GSInputFile
50530  ShowParams
50540  CallGhostscript "PSD"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreatePSD")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePCL(GSInputFile As String, GSOutputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, FName As String, Ext As String, tStr As String
50020
50030  GSInit Options
50040  InitParams
50050  Set ParamCommands = New Collection
50060
50070  If Options.OnePagePerFile = 1 Then
50080   SplitPath GSOutputFile, , Path, , FName, Ext
50090   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50100  End If
50110  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50120  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50130   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50140  End If
50150  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50160   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50170  End If
50180  AddParams "-I" & tStr
50190  AddParams "-q"
50200  AddParams "-dNOPAUSE"
50210  AddParams "-dSAFER"
50220  AddParams "-dBATCH"
50230  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50240   AddParams "-sFONTPATH=" & GetFontsDirectory
50250  End If
50260
50270  AddParams "-sDEVICE=" & GS_PCLColorscount
50280  If Options.DontUseDocumentSettings = 0 Then
50290   If Options.UseFixPapersize <> 0 Then
50300    If Options.UseCustomPaperSize = 0 Then
50310      If LenB(Trim$(Options.Papersize)) > 0 Then
50320       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50330       AddParams "-dFIXEDMEDIA"
50340       AddParams "-dNORANGEPAGESIZE"
50350      End If
50360     Else
50370      If Options.DeviceWidthPoints >= 1 Then
50380       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50390      End If
50400      If Options.DeviceHeightPoints >= 1 Then
50410       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50420      End If
50430    End If
50440   End If
50450   AddParams "-r" & Options.PCLResolution & "x" & Options.PCLResolution
50460  End If
50470  AddParams "-sOutputFile=" & GSOutputFile
50480
50490  AddAdditionalGhostscriptParameters
50500
50510  AddParams "-f"
50520  AddParams GSInputFile
50530  ShowParams
50540  CallGhostscript "PCL"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreatePCL")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateRAW(GSInputFile As String, GSOutputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, FName As String, Ext As String, tStr As String
50020
50030  GSInit Options
50040  InitParams
50050  Set ParamCommands = New Collection
50060
50070  If Options.OnePagePerFile = 1 Then
50080   SplitPath GSOutputFile, , Path, , FName, Ext
50090   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50100  End If
50110  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50120  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50130   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50140  End If
50150  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50160   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50170  End If
50180  AddParams "-I" & tStr
50190  AddParams "-q"
50200  AddParams "-dNOPAUSE"
50210  AddParams "-dSAFER"
50220  AddParams "-dBATCH"
50230  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50240   AddParams "-sFONTPATH=" & GetFontsDirectory
50250  End If
50260
50270  AddParams "-sDEVICE=" & GS_RAWColorscount
50280  If Options.DontUseDocumentSettings = 0 Then
50290   If Options.UseFixPapersize <> 0 Then
50300    If Options.UseCustomPaperSize = 0 Then
50310      If LenB(Trim$(Options.Papersize)) > 0 Then
50320       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50330       AddParams "-dFIXEDMEDIA"
50340       AddParams "-dNORANGEPAGESIZE"
50350      End If
50360     Else
50370      If Options.DeviceWidthPoints >= 1 Then
50380       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50390      End If
50400      If Options.DeviceHeightPoints >= 1 Then
50410       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50420      End If
50430    End If
50440   End If
50450   AddParams "-r" & Options.RAWResolution & "x" & Options.RAWResolution
50460  End If
50470  AddParams "-sOutputFile=" & GSOutputFile
50480
50490  AddAdditionalGhostscriptParameters
50500
50510  AddParams "-f"
50520  AddParams GSInputFile
50530  ShowParams
50540  CallGhostscript "RAW"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CreateRAW")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CallGScript(GSInputFile As String, GSOutputFile As String, _
 Options As tOptions, Ghostscriptdevice As tGhostscriptDevice)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim enc As Boolean, encPDF As EncryptData, retEnc As Boolean, _
  Tempfile As String, tL As Long, m As Object
50030  GSInit Options
50041  Select Case Ghostscriptdevice
        Case 0: 'PDF
50060    With Options
50070     If .PDFOptimize = 1 And .PDFUseSecurity = 0 Then
50080       Tempfile = GetTempFile(GetTempPath, "~CP")
50090       KillFile Tempfile
50100       CreatePDF GSInputFile, Tempfile, Options
50110       OptimizePDF Tempfile, GSOutputFile
50120       KillFile Tempfile
50130      Else
50140       If .PDFUseSecurity <> 0 And SecurityIsPossible = True Then
50150         If .PDFEncryptor = 1 Then
50160           enc = SetEncryptionParams(encPDF, GSInputFile, GSOutputFile)
50170           If enc = True Then
50180            Tempfile = GetTempFile(GetTempPath, "~CP")
50190            KillFile Tempfile
50200            CreatePDF GSInputFile, Tempfile, Options
50210            encPDF.InputFile = Tempfile
50220            retEnc = EncryptPDF(encPDF)
50230            KillFile encPDF.InputFile
50240            If retEnc = False Then
50250             IfLoggingWriteLogfile "Error with encryption - using unencrypted file"
50260             Name GSInputFile As GSOutputFile
50270            End If
50280           End If
50290          Else
50300           tL = .PDFOptimize
50310           .PDFOptimize = 0
50320           CreatePDF GSInputFile, GSOutputFile, Options
50330           .PDFOptimize = tL
50340         End If
50350        Else
50360         CreatePDF GSInputFile, GSOutputFile, Options
50370       End If
50380     End If
50390     If PDFUpdateMetadataIsPossible Then
50400      If .PDFUpdateMetadata > 0 Then
50410       If .PDFUpdateMetadata = 2 Or _
       (.PDFUpdateMetadata = 1 And (InStr(1, .AdditionalGhostscriptParameters, "dpdfa", vbTextCompare) > 0)) Then
50430        Set m = CreateObject("pdfForge.pdf.pdf")
50440        Tempfile = GetTempFile(GetTempPath, "~MP")
50450        KillFile Tempfile
50460        Call m.UpdateXMPMetadata(GSOutputFile, Tempfile)
50470        If FileExists(Tempfile) Then
50480         If KillFile(GSOutputFile) Then
50490          Name Tempfile As GSOutputFile
50500         End If
50510        End If
50520       End If
50530      End If
50540     End If
50550     If PDFSigningIsPossible Then
50560      If .PDFSigningSignPDF = 1 Then
50570       SignPDF GSOutputFile
50580      End If
50590     End If
50600    End With
50610   Case 1: 'PNG
50620    CreatePNG GSInputFile, GSOutputFile, Options
50630   Case 2: 'JPEG
50640    CreateJPEG GSInputFile, GSOutputFile, Options
50650   Case 3: 'BMP
50660    CreateBMP GSInputFile, GSOutputFile, Options
50670   Case 4: 'PCX
50680    CreatePCX GSInputFile, GSOutputFile, Options
50690   Case 5: 'TIFF
50700    CreateTIFF GSInputFile, GSOutputFile, Options
50710   Case 6: 'PS
50720    CreatePS GSInputFile, GSOutputFile, Options
50730   Case 7: 'EPS
50740    CreateEPS GSInputFile, GSOutputFile, Options
50750   Case 8: 'TXT
50760    CreateTXT GSInputFile, Options
50770    CreateTextFile GSOutputFile, GS_OutStr
50780   Case 9: 'PDFA
50790    CreatePDFA GSInputFile, GSOutputFile, Options
50800    With Options
50810     If PDFUpdateMetadataIsPossible Then
50820      If .PDFUpdateMetadata > 0 Then
50830       Set m = CreateObject("pdfForge.pdf.pdf")
50840       Tempfile = GetTempFile(GetTempPath, "~MP")
50850       KillFile Tempfile
50860       Call m.UpdateXMPMetadata(GSOutputFile, Tempfile)
50870       If FileExists(Tempfile) Then
50880        If KillFile(GSOutputFile) Then
50890         Name Tempfile As GSOutputFile
50900        End If
50910       End If
50920      End If
50930     End If
50940     If PDFSigningIsPossible Then
50950      If .PDFSigningSignPDF = 1 Then
50960       SignPDF GSOutputFile
50970      End If
50980     End If
50990    End With
51000   Case 10: 'PDFX
51010    CreatePDFX GSInputFile, GSOutputFile, Options
51020    With Options
51030     If PDFSigningIsPossible Then
51040      If .PDFSigningSignPDF = 1 Then
51050       SignPDF GSOutputFile
51060      End If
51070     End If
51080    End With
51090   Case 11: 'PSD
51100    CreatePSD GSInputFile, GSOutputFile, Options
51110   Case 12: 'PCL
51120    CreatePCL GSInputFile, GSOutputFile, Options
51130   Case 13: 'RAW
51140    CreateRAW GSInputFile, GSOutputFile, Options
51150  End Select
51160
51170  Options.Counter = Options.Counter + 1
51180  SaveOption Options, "Counter"
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

Private Sub SignPDF(Filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim m As Object, PFXPassword As String, signatureVisible As Boolean, multiSignatures As Boolean, Tempfile As String, tStr As String
50020   Dim res As Long, files As Collection, certFilename As String
50030  With Options
50040   If LenB(.PDFSigningPFXFile) = 0 Then
50050     res = OpenFileDialog(files, "", "PFX\P12 files (*.pfx,*.p12)|*.pfx;*.p12|PFX files (*.pfx)|*pfx|P12 files (*.p12|*.p12", "*.pfx;*.p12", "C:\", "Choose a certificate", OFN_FILEMUSTEXIST Or OFN_EXPLORER Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST, 0, 1)
50060     If res > 0 Then
50070      certFilename = files(1)
50080     End If
50090    Else
50100     certFilename = .PDFSigningPFXFile
50110   End If
50120   If LenB(.PDFSigningPFXFilePassword) > 0 Then
50130     PFXPassword = .PDFSigningPFXFilePassword
50140    Else
50150     'Ask for the password
50160     PFXPassword = InputBox("Certifacte password")
50170   End If
50180   Tempfile = GetTempFile(GetTempPath, "~MP")
50190   KillFile Tempfile
50200   Set m = CreateObject("pdfForge.pdf.pdf")
50210   If .PDFSigningSignatureVisible = 0 Then
50220     signatureVisible = False
50230    Else
50240     signatureVisible = True
50250   End If
50260   If .PDFSigningMultiSignature = 0 Then
50270     multiSignatures = False
50280    Else
50290     multiSignatures = True
50300   End If
50310   Call m.signPDFFile(Filename, Tempfile, certFilename, PFXPassword, .PDFSigningSignatureReason, .PDFSigningSignatureContact, .PDFSigningSignatureLocation, _
   signatureVisible, .PDFSigningSignatureLeftX, .PDFSigningSignatureLeftY, .PDFSigningSignatureRightX, .PDFSigningSignatureRightY, multiSignatures, Nothing)
50330  End With
50340  If FileExists(Tempfile) Then
50350   If KillFile(Filename) Then
50360    Name Tempfile As Filename
50370   End If
50380  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "SignPDF")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function OptimizePDF(PDFInputFilename As String, PDFOutputFilename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim LastStop As Currency, tStr As String, c As Currency
50020  InitParams
50030  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50040  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50050   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50060  End If
50070  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50080   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50090  End If
50100  AddParams "-I" & tStr
50110  AddParams "-q"
50120  AddParams "-dNODISPLAY"
50130  AddParams "-dSAFER"
50140  AddParams "-dDELAYSAFER"
50150  AddParams "--"
50160  AddParams "pdfopt.ps"
50170  AddParams PDFInputFilename
50180  AddParams PDFOutputFilename
50190
50200  GSParams(0) = "pdfopt"
50210   If PerformanceTimer Then
50220    c = ExactTimer_Value() - LastStop
50230    IfLoggingWriteLogfile "Time for converting: " & _
    Format$(Int(c) * (1 / 86400), "hh:nn:ss:") & Format$(((c) - Int(c)) * 1000, "000")
50250   Else
50260    IfLoggingWriteLogfile "Time for converting -> No performance timer"
50270  End If
50280
50290  If PerformanceTimer Then
50300   LastStop = ExactTimer_Value()
50310  End If
50320  OptimizePDF = CallGS(GSParams)
50330  If PerformanceTimer Then
50340    c = ExactTimer_Value() - LastStop
50350    IfLoggingWriteLogfile "Time for optimizing: " & _
    Format$(Int(c) * (1 / 86400), "hh:nn:ss:") & Format$(((c) - Int(c)) * 1000, "000")
50370   Else
50380    IfLoggingWriteLogfile "Time for optimizing: No performance timer"
50390  End If
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

Public Function Bool2Text(Number As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Number = 1 Then
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
50021  Select Case gsMethod
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
50021  Select Case gsMethod
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
50011  Select Case gsMethod
        Case 0
50030    GS_COMPRESSMONOMETHOD = "CCITTFaxEncode"
50040   Case 1
50050    GS_COMPRESSMONOMETHOD = "FlateEncode"
50060   Case 2
50070    GS_COMPRESSMONOMETHOD = "RunLengthEncode"
50080   Case 3
50090    GS_COMPRESSMONOMETHOD = "LZWEncode"
50100  End Select
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

Private Sub ShowParams()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, tStr As String
50020  If Options.Logging <> 0 Then
50030   tStr = GSParams(LBound(GSParams))
50040   For i = LBound(GSParams) + 1 To UBound(GSParams)
50050    tStr = tStr & vbCrLf & GSParams(i)
50060   Next i
50070   IfLoggingWriteLogfile "Ghostscriptparameter:" & vbCrLf & tStr
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "ShowParams")
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

Private Function BuildPermissionString(encData As EncryptData) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strPermissions As String
50020
50030  strPermissions = vbNullString
50040  strPermissions = strPermissions & Abs(Int(Not encData.DisallowPrinting))
50050  strPermissions = strPermissions & Abs(Int(Not encData.DisallowModifyContents))
50060  strPermissions = strPermissions & Abs(Int(Not encData.DisallowCopy))
50070  strPermissions = strPermissions & Abs(Int(Not encData.DisallowModifyAnnotations))
50080  If Options.PDFHighEncryption Then
50090    strPermissions = strPermissions & Abs(Int(encData.AllowFillIn)) '(128 bit only)
50100    strPermissions = strPermissions & Abs(Int(encData.AllowScreenReaders)) '(128 bit only)
50110    strPermissions = strPermissions & Abs(Int(encData.AllowAssembly)) '(128 bit only)
50120    strPermissions = strPermissions & Abs(Int(encData.AllowDegradedPrinting)) '(128 bit only)
50130   Else
50140    strPermissions = strPermissions & "0000"
50150  End If
50160  BuildPermissionString = strPermissions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "BuildPermissionString")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function EncryptPDF(encData As EncryptData) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strPermissions As String, strShell As String, ret As Double
50020
50030  strPermissions = BuildPermissionString(encData)
50040
50050 ' strShell = App.Path & "\pdfencrypt.exe """ & encData.InputFile & """ """ & encData.OutputFile & """ User=" & encData.UserPass & " Owner=" & encData.OwnerPass & " " & strPermissions & " " & encData.EncryptionLevel
50060 ' strShell = CompletePath(Options.DirectoryJava) & "Java.exe -cp """ & CompletePath(App.Path) & "iText.jar"" com.lowagie.tools.encrypt_pdf """ & encData.InputFile & """ """ & encData.OutputFile & """ """ & encData.UserPass & """ """ & encData.OwnerPass & """ " & strPermissions & " " & encData.EncryptionLevel
50070
50080  strShell = GetPDFCreatorApplicationPath & "pdfenc.exe """ & encData.InputFile & """ """ & encData.OutputFile & """ """ & encData.UserPass & """ """ & encData.OwnerPass & """ " & strPermissions & " " & encData.EncryptionLevel
50090
50100  IfLoggingWriteLogfile strShell
50110
50120  ret = RunProgramWait(strShell, False)
50130
50140  If Dir$(encData.OutputFile) <> vbNullString Then
50150   EncryptPDF = True
50160  End If
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

Public Function CalculatePermissions(ByRef encData As EncryptData) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tB As Long, tB2 As Long
50020  tB = 192
50030  With encData
50040   If Abs(.DisallowPrinting) = 0 Then
50050    tB = tB + 4
50060   End If
50070   If Abs(.DisallowModifyContents) = 0 Then
50080    tB = tB + 8
50090   End If
50100   If Abs(.DisallowCopy) = 0 Then
50110    tB = tB + 16
50120   End If
50130   If Abs(.DisallowModifyAnnotations) = 0 Then
50140    tB = tB + 32
50150   End If
50160   CalculatePermissions = tB - 256
50170   If .EncryptionLevel = encStrong Then
50180     tB2 = 240
50190     If Abs(.AllowFillIn) <> 0 Then
50200      tB2 = tB2 + 1
50210     End If
50220     If Abs(.AllowScreenReaders) <> 0 Then
50230      tB2 = tB2 + 2
50240     End If
50250     If Abs(.AllowAssembly) <> 0 Then
50260      tB2 = tB2 + 4
50270     End If
50280     If Abs(.AllowDegradedPrinting) = 0 Then
50290      tB2 = tB2 + 8
50300     End If
50310    CalculatePermissions = CalculatePermissions - (255 - tB2) * 256&
50320   End If
50330  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CalculatePermissions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function SetEncryptionParams(ByRef encData As EncryptData, InputFile As String, OutputFile As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim retPasswd As Boolean
50020
50030  encData.InputFile = InputFile
50040  encData.OutputFile = OutputFile
50050
50060  If Len(Options.PDFOwnerPasswordString) > 0 Then
50070    encData.OwnerPass = Options.PDFOwnerPasswordString
50080    OwnerPassword = Options.PDFOwnerPasswordString
50090    If Options.PDFUserPass = 1 Then
50100     encData.UserPass = Options.PDFUserPasswordString
50110     UserPassword = Options.PDFUserPasswordString
50120    End If
50130    retPasswd = True
50140   Else
50150    If SavePasswordsForThisSession = False Then
50160      If Options.UseAutosave = 0 Then
50170        retPasswd = EnterPasswords(encData.UserPass, encData.OwnerPass, frmPassword)
50180       Else
50190        retPasswd = False
50200      End If
50210     Else
50220      encData.OwnerPass = OwnerPassword: encData.UserPass = UserPassword
50230    End If
50240  End If
50250  If retPasswd = True Or SavePasswordsForThisSession = True Then
50260    With encData
50270     .DisallowPrinting = Options.PDFDisallowPrinting
50280     .DisallowModifyContents = Options.PDFDisallowModifyContents
50290     .DisallowCopy = Options.PDFDisallowCopy
50300     .DisallowModifyAnnotations = Options.PDFDisallowModifyAnnotations
50310     .AllowFillIn = Options.PDFAllowFillIn
50320     .AllowScreenReaders = Options.PDFAllowScreenReaders
50330     .AllowAssembly = Options.PDFAllowAssembly
50340     .AllowDegradedPrinting = Options.PDFAllowDegradedPrinting
50350     If Options.PDFHighEncryption = 1 Then
50360       .EncryptionLevel = encStrong
50370      Else
50380       .EncryptionLevel = encLow
50390     End If
50400    End With
50410    SetEncryptionParams = True
50420    encData.UserPass = UserPassword
50430    encData.OwnerPass = OwnerPassword
50440   Else
50450    SetEncryptionParams = False
50460  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "SetEncryptionParams")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub SetColorParams()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Options.PDFCompressionColorCompression = 1 Then
50020    AddParams "-dEncodeColorImages=true"
50030    If Options.PDFCompressionColorCompressionChoice = 0 Then
50040      AddParams "-dAutoFilterColorImages=true"
50050     Else
50060      AddParams "-dAutoFilterColorImages=false"
50070      If Options.PDFCompressionColorResample = 1 Then
50080        AddParams "-dDownsampleColorImages=true"
50091        Select Case Options.PDFCompressionColorResampleChoice
              Case 0:
50110          AddParams "-dColorImageDownsampleType=/Subsample"
50120         Case 1:
50130          AddParams "-dColorImageDownsampleType=/Average"
50140         Case 2:
50150          AddParams "-dColorImageDownsampleType=/Bicubic"
50160        End Select
50170        AddParams "-dColorImageResolution=" & Options.PDFCompressionColorResolution
50180       Else
50190        AddParams "-dDownsampleColorImages=false"
50200      End If
50211      Select Case Options.PDFCompressionColorCompressionChoice
            Case 1:
50230        AddParams "-dColorImageFilter=/DCTEncode"
50240        AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionColorCompressionJPEGMaximumFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50270       Case 2:
50280        AddParams "-dColorImageFilter=/DCTEncode"
50290        AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionColorCompressionJPEGHighFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50320       Case 3:
50330        AddParams "-dColorImageFilter=/DCTEncode"
50340        AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionColorCompressionJPEGMediumFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50370       Case 4:
50380        AddParams "-dColorImageFilter=/DCTEncode"
50390        AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionColorCompressionJPEGLowFactor), GetDecimalChar, ".") & _
        " /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
50420       Case 5:
50430        AddParams "-dColorImageFilter=/DCTEncode"
50440        AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor " & _
       Replace$(CStr(Options.PDFCompressionColorCompressionJPEGMinimumFactor), GetDecimalChar, ".") & _
       " /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
50470       Case 6:
50480        AddParams "-dColorImageFilter=/FlateEncode"
50490       Case 7:
50500        AddParams "-dColorImageFilter=/LZWEncode"
50510      End Select
50520    End If
50530   Else
50540    AddParams "-dEncodeColorImages=false"
50550  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "SetColorParams")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetGreyParams()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Options.PDFCompressionGreyCompression = 1 Then
50020    AddParams "-dEncodeGrayImages=true"
50030    If Options.PDFCompressionGreyCompressionChoice = 0 Then
50040      AddParams "-dAutoFilterGrayImages=true"
50050     Else
50060      AddParams "-dAutoFilterGrayImages=false"
50070      If Options.PDFCompressionGreyResample = 1 Then
50080        AddParams "-dDownsampleGrayImages=true"
50091        Select Case Options.PDFCompressionGreyResampleChoice
              Case 0:
50110          AddParams "-dGrayImageDownsampleType=/Subsample"
50120         Case 1:
50130          AddParams "-dGrayImageDownsampleType=/Average"
50140         Case 2:
50150          AddParams "-dGrayImageDownsampleType=/Bicubic"
50160        End Select
50170        AddParams "-dGrayImageResolution=" & Options.PDFCompressionGreyResolution
50180       Else
50190        AddParams "-dDownsampleGrayImages=false"
50200      End If
50211      Select Case Options.PDFCompressionGreyCompressionChoice
            Case 1:
50230        AddParams "-dGrayImageFilter=/DCTEncode"
50240        AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionGreyCompressionJPEGMaximumFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50270       Case 2:
50280        AddParams "-dGrayImageFilter=/DCTEncode"
50290        AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionGreyCompressionJPEGHighFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50320       Case 3:
50330        AddParams "-dGrayImageFilter=/DCTEncode"
50340        AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionGreyCompressionJPEGMediumFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50370       Case 4:
50380        AddParams "-dGrayImageFilter=/DCTEncode"
50390        AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionGreyCompressionJPEGLowFactor), GetDecimalChar, ".") & _
        " /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
50420       Case 5:
50430        AddParams "-dGrayImageFilter=/DCTEncode"
50440        AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor " & _
       Replace$(CStr(Options.PDFCompressionGreyCompressionJPEGMinimumFactor), GetDecimalChar, ".") & _
       " /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
50470       Case 6:
50480        AddParams "-dGrayImageFilter=/FlateEncode"
50490       Case 7:
50500        AddParams "-dGrayImageFilter=/LZWEncode"
50510      End Select
50520    End If
50530   Else
50540    AddParams "-dEncodeGrayImages=false"
50550  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "SetGreyParams")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetMonoParams()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Options.PDFCompressionMonoCompression = 1 Then
50020    AddParams "-dEncodeMonoImages=true"
50031    Select Case Options.PDFCompressionMonoCompressionChoice
          Case 0:
50050      AddParams "-dMonoImageFilter=/CCITTFaxEncode"
50060     Case 1:
50070      AddParams "-dMonoImageFilter=/FlateEncode"
50080     Case 2:
50090      AddParams "-dMonoImageFilter=/RunLengthEncode"
50100     Case 3:
50110      AddParams "-dMonoImageFilter=/LZWEncode"
50120    End Select
50130    If Options.PDFCompressionMonoResample = 1 Then
50140      AddParams "-dDownsampleMonoImages=true"
50151      Select Case Options.PDFCompressionMonoResampleChoice
            Case 0:
50170        AddParams "-dMonoImageDownsampleType=/Subsample"
50180       Case 1:
50190        AddParams "-dMonoImageDownsampleType=/Average"
50200       Case 2:
50210        AddParams "-dMonoImageDownsampleType=/Bicubic"
50220      End Select
50230      AddParams "-dMonoImageResolution=" & Options.PDFCompressionMonoResolution
50240     Else
50250      AddParams "-dDownsampleMonoImages=false"
50260    End If
50270   Else
50280    AddParams "-dEncodeMonoImages=false"
50290  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "SetMonoParams")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GhostScriptSecurity() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GhostScriptSecurity = False
50020  If LenB(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll) = 0 Then
50030   Exit Function
50040  End If
50050 ' GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50060  If GsDllLoaded = 0 Then
50070   Exit Function
50080  End If
50090  GSRevision = GetGhostscriptRevision
50100 ' UnLoadDLL GsDllLoaded
50110  If InStr(UCase$(GSRevision.strProduct), "AFPL") > 0 Then
50120   If GSRevision.intRevision < 814 Then
50130    Exit Function
50140   End If
50150   GhostScriptSecurity = True
50160   Exit Function
50170  End If
50180  If InStr(UCase$(GSRevision.strProduct), "GPL") > 0 Then
50190   If GSRevision.intRevision < 815 Then
50200    Exit Function
50210   End If
50220   GhostScriptSecurity = True
50230   Exit Function
50240  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "GhostScriptSecurity")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetAllGhostscriptversions() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, tStr As String, tColl As Collection, i As Long, _
  tf() As String, GS_DLL As String, GS_LIB As String, tB As Boolean, j As Long
50030  Set reg = New clsRegistry
50040  Set GetAllGhostscriptversions = New Collection
50050  With reg
50060   .hkey = HKEY_LOCAL_MACHINE
50070   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50080   If .KeyExists = True Then
50090    tStr = Trim$(.GetRegistryValue("GhostscriptCopyright"))
50100    If LenB(tStr) > 0 Then
50110     tStr = Replace$(LanguageStrings.OptionsGhostscriptInternal, "%1", tStr)
50120     tStr = Replace$(tStr, "%2", Trim$(.GetRegistryValue("GhostscriptVersion")))
50130     GetAllGhostscriptversions.Add tStr
50140    End If
50150   End If
50160   tStr = "AFPL Ghostscript"
50170   .KeyRoot = "SOFTWARE\" & tStr
50180   Set tColl = .EnumRegistryKeys(HKEY_LOCAL_MACHINE, .KeyRoot)
50190   For i = 1 To tColl.Count
50200    .Subkey = tColl.Item(i)
50210    GS_DLL = .GetRegistryValue("GS_DLL")
50220    GS_LIB = .GetRegistryValue("GS_LIB")
50230    If Len(GS_DLL) > 0 Then
50240     If FileExists(GS_DLL) = True Then
50250      GetAllGhostscriptversions.Add tStr & " " & tColl.Item(i)
50260     End If
50270    End If
50280   Next i
50290   tStr = "GNU Ghostscript"
50300   .KeyRoot = "SOFTWARE\" & tStr
50310   Set tColl = .EnumRegistryKeys(HKEY_LOCAL_MACHINE, .KeyRoot)
50320   For i = 1 To tColl.Count
50330    .Subkey = tColl.Item(i)
50340    GS_DLL = .GetRegistryValue("GS_DLL")
50350    GS_LIB = .GetRegistryValue("GS_LIB")
50360    If Len(GS_DLL) > 0 Then
50370     If FileExists(GS_DLL) = True Then
50380      GetAllGhostscriptversions.Add tStr & " " & tColl.Item(i)
50390     End If
50400    End If
50410   Next i
50420   tStr = "GPL Ghostscript"
50430   .KeyRoot = "SOFTWARE\" & tStr
50440   Set tColl = .EnumRegistryKeys(HKEY_LOCAL_MACHINE, .KeyRoot)
50450   For i = 1 To tColl.Count
50460    .Subkey = tColl.Item(i)
50470    GS_DLL = .GetRegistryValue("GS_DLL")
50480    GS_LIB = .GetRegistryValue("GS_LIB")
50490    If Len(GS_DLL) > 0 Then
50500     If FileExists(GS_DLL) = True Then
50510      GetAllGhostscriptversions.Add tStr & " " & tColl.Item(i)
50520     End If
50530    End If
50540   Next i
50550  End With
50560  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "GetAllGhostscriptversions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub CheckForStamping(Filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim StampPage As String, tStr As String, R As String, G As String, b As String, _
  Stampfile As String, Path As String, ff As Long, files As Collection, _
  StampString As String, StampFontsize As Double, _
  StampOutlineFontthickness As Double
50050  StampString = RemoveLeadingAndTrailingQuotes(Trim$(Options.StampString))
50060  If Len(StampString) > 0 Then
50070   StampPage = StrConv(LoadResData(101, "STAMPPAGE"), vbUnicode)
50080   StampPage = Replace(StampPage, vbCrLf, vbCr, , , vbBinaryCompare)
50090   StampPage = Replace(StampPage, "[STAMPSTRING]", EncodeCharsOctal(StampString), , , vbTextCompare)
50100   StampPage = Replace(StampPage, "[FONTNAME]", Replace(Trim$(Options.StampFontname), " ", ""), , , vbTextCompare)
50110   StampFontsize = 48
50120   If IsNumeric(Options.StampFontsize) = True Then
50130    If CDbl(Options.StampFontsize) > 0 Then
50140     StampFontsize = CDbl(Options.StampFontsize)
50150    End If
50160   End If
50170   StampPage = Replace(StampPage, "[FONTSIZE]", StampFontsize, , , vbTextCompare)
50180   StampOutlineFontthickness = 0
50190   If IsNumeric(Options.StampOutlineFontthickness) = True Then
50200    If CDbl(Options.StampOutlineFontthickness) >= 0 Then
50210     StampOutlineFontthickness = CDbl(Options.StampOutlineFontthickness)
50220    End If
50230   End If
50240   StampPage = Replace(StampPage, "[STAMPOUTLINEFONTTHICKNESS]", StampOutlineFontthickness, , , vbTextCompare)
50250   If Options.StampUseOutlineFont <> 1 Then
50260     StampPage = Replace(StampPage, "[USEOUTLINEFONT]", "show", , , vbTextCompare)
50270    Else
50280     StampPage = Replace(StampPage, "[USEOUTLINEFONT]", "true charpath stroke", , , vbTextCompare)
50290   End If
50300   If Len(Options.StampFontColor) > 0 Then
50310     tStr = Replace$(Options.StampFontColor, "#", "&H")
50320     If IsNumeric(tStr) = True Then
50330       R = Replace$(Format(CDbl((CLng(tStr) And CLng("&HFF0000")) / 65536) / 255#, "0.00"), ",", ".", , 1)
50340       G = Replace$(Format(CDbl((CLng(tStr) And CLng("&H00FF00")) / 256) / 255#, "0.00"), ",", ".", , 1)
50350       b = Replace$(Format(CDbl(CLng(tStr) And CLng("&H0000FF")) / 255#, "0.00"), ",", ".", , 1)
50360       StampPage = Replace(StampPage, "[FONTCOLOR]", R & " " & G & " " & b, , , vbTextCompare)
50370      Else
50380       StampPage = Replace(StampPage, "[FONTCOLOR]", "1 0 0", , , vbTextCompare)
50390     End If
50400    Else
50410     StampPage = Replace(StampPage, "[FONTCOLOR]", "1 0 0", , , vbTextCompare)
50420   End If
50430   Path = CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & GetUsername
50440   If DirExists(Path) = False Then
50450    MakePath Path
50460   End If
50470   Stampfile = GetTempFile(Path, "~ST")
50480   ff = FreeFile
50490   Open Stampfile For Output As #ff
50500   Print #ff, StampPage
50510   Close #ff
50520   Set files = New Collection
50530   files.Add Stampfile
50540   files.Add Filename
50550   Stampfile = GetTempFile(Path, "~ST")
50560   KillFile Stampfile
50570   CombineFiles Stampfile, files
50580   Name Stampfile As Filename
50590  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CheckForStamping")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AddParamCommands()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  If ParamCommands.Count > 0 Then
50030   AddParams "-c"
50040   For i = 1 To ParamCommands.Count
50050    AddParams ParamCommands(i)
50060   Next i
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "AddParamCommands")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AddParamCommand(GhostscriptCommand As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ParamCommands.Add GhostscriptCommand
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "AddParamCommand")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AddAdditionalGhostscriptParameters()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String, tStrf() As String, i As Long
50020  tStr = Replace$(Trim$(Options.AdditionalGhostscriptParameters), "<app>", GetPDFCreatorApplicationPath, , , vbTextCompare)
50030  tStr = Replace$(Trim$(tStr), "<gslib>", CompletePath(Options.DirectoryGhostscriptLibraries), , , vbTextCompare)
50040  If LenB(tStr) > 0 Then
50050   If InStr(1, tStr, "|") > 0 Then
50060     tStrf = Split(tStr, "|")
50070     For i = LBound(tStrf) To UBound(tStrf)
50080      tStr = Trim$(tStrf(i))
50090      If LenB(tStr) > 0 Then
50100       AddParams tStr
50110      End If
50120     Next i
50130    Else
50140     AddParams tStr
50150   End If
50160  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "AddAdditionalGhostscriptParameters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub CheckForPrintingAfterSaving(GSInputFile As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String
50020
50030  If Options.PrintAfterSaving = 0 Then
50040   Exit Sub
50050  End If
50060
50070  GSInit Options
50080  InitParams
50090  Set ParamCommands = New Collection
50100
50110  tStr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50120  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50130   tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50140  End If
50150  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50160   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50170  End If
50180  AddParams "-I" & tStr
50190  AddParams "-q"
50200  AddParams "-dNOPAUSE"
50210  AddParams "-dSAFER"
50220  AddParams "-dBATCH"
50230  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50240   AddParams "-sFONTPATH=" & GetFontsDirectory
50250  End If
50260
50270  AddParams "-sDEVICE=mswinpr2"
50280  tStr = ""
50290  If Options.PrintAfterSavingQueryUser > 0 Then
50300   tStr = "/QueryUser " & Options.PrintAfterSavingQueryUser
50310  End If
50320  If Options.PrintAfterSavingNoCancel = 1 Then
50330   tStr = tStr & " /NoCancel"
50340  End If
50350  tStr = Trim$(tStr)
50360  If LenB(tStr) > 0 Then
50370   AddParamCommand "<< " & tStr & " >> setpagedevice"
50380  End If
50390  AddParams "-sOutputFile=\\spool\" & Options.PrintAfterSavingPrinter
50400  If Options.PrintAfterSavingDuplex = 1 Then
50410   If Options.PrintAfterSavingTumble = 1 Then
50420     AddParamCommand "<< /Duplex true /Tumble true >> setpagedevice"
50430    Else
50440     AddParamCommand "<< /Duplex true /Tumble false >> setpagedevice"
50450   End If
50460  End If
50470  AddParamCommands
50480  AddParams "-f"
50490  AddParams GSInputFile
50500  ShowParams
50510  CallGhostscript "mswinpr2"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "CheckForPrintingAfterSaving")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub ConvertFile(InputFilename As String, OutputFilename As String, Optional SubFormat As String = "")
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Ext As String, Tempfile As String, ivgf As Boolean, inFile As String
50020  IFIsPS = False
50030  If LenB(InputFilename) = 0 Then
50040   Exit Sub
50050  End If
50060  If FileExists(InputFilename) = False Then
50070   If LenB(InputFilename) > 0 Then
50080    MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & _
    "InputFile -IF" & vbCrLf & ">" & InputFilename & "<", vbExclamation + vbOKOnly
50100   End If
50110   Exit Sub
50120  End If
50130  ivgf = IsValidGraphicFile(InputFilename)
50140  If LenB(OutputFilename) > 0 Then
50150    If IsPostscriptFile(InputFilename) = True Or ivgf Then
50160     If GsDllLoaded = 0 Then
50170      Exit Sub
50180     End If
50190     GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50200     If GsDllLoaded = 0 Then
50210      MsgBox LanguageStrings.MessagesMsg08
50220     End If
50230     inFile = InputFilename
50240     If ivgf Then
50250      Tempfile = GetTempFile(GetTempPath, "~p")
50260      Kill Tempfile
50270      If Image2PS(InputFilename, Tempfile) Then
50280        inFile = Tempfile
50290       Else
50300        IfLoggingWriteLogfile "ConvertFile: There is a problem converting '" & InputFilename & "'!"
50310        Exit Sub
50320      End If
50330     End If
50340     SplitPath OutputFilename, , , , , Ext
50351     Select Case UCase$(Ext)
           Case "PDF"
50371       Select Case UCase(SubFormat)
             Case "PDF/A-1B"
50390         CallGScript inFile, OutputFilename, Options, PDFAWriter
50400        Case "PDF/X"
50410         CallGScript inFile, OutputFilename, Options, PDFXWriter
50420        Case Else
50430         CallGScript inFile, OutputFilename, Options, PDFWriter
50440       End Select
50450      Case "PNG"
50460       CallGScript inFile, OutputFilename, Options, PNGWriter
50470      Case "JPG"
50480       CallGScript inFile, OutputFilename, Options, JPEGWriter
50490      Case "BMP"
50500       CallGScript inFile, OutputFilename, Options, BMPWriter
50510      Case "PCX"
50520       CallGScript inFile, OutputFilename, Options, PCXWriter
50530      Case "TIF"
50540       CallGScript inFile, OutputFilename, Options, TIFFWriter
50550      Case "PS"
50560       CallGScript inFile, OutputFilename, Options, PSWriter
50570      Case "EPS"
50580       CallGScript inFile, OutputFilename, Options, EPSWriter
50590      Case "TXT"
50600       CallGScript inFile, OutputFilename, Options, TXTWriter
50610      Case "PCL"
50620       CallGScript inFile, OutputFilename, Options, PCLWriter
50630      Case "PSD"
50640       CallGScript inFile, OutputFilename, Options, PSDWriter
50650      Case "RAW"
50660       CallGScript inFile, OutputFilename, Options, RAWWriter
50670     End Select
50680     If ivgf Then
50690      KillFile Tempfile
50700     End If
50710    End If
50720 '   If GsDllLoaded <> 0 Then
50730 '    UnloadDLLComplete GsDllLoaded
50740 '   End If
50750    ConvertedOutputFilename = OutputFilename
50760    ReadyConverting = True
50770    Exit Sub
50780   Else
50790    If FileExists(InputFilename) = True Then
50800     If IsPostscriptFile(InputFilename) = True Then
50810       IFIsPS = True
50820      Else
50830       MsgBox LanguageStrings.MessagesMsg06 & vbCrLf & vbCrLf & InputFilename
50840     End If
50850    End If
50860  End If
50870  DoEvents
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "ConvertFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
