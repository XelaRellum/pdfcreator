Attribute VB_Name = "modGhostScript"
Option Explicit

Public Type tGhostscriptVersion
 Major As Long
 Minor As Long
End Type

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
 SVGWriter = 14
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

Public GS_OutStr As String, PFXPassword As String

Private ParamCommands As Collection, currentOwnerPassword As String

Public Sub GSInit(Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Rotate(2) As String, Resample(2) As String, Colormodel(2) As String, _
  ColorsPreserveTransfer(1) As String
50030  Dim PNGColorscount(5) As String, JPEGColorscount(1) As String, BMPColorscount(7) As String, _
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
50210  PNGColorscount(4) = "pnggray": PNGColorscount(5) = "pngalpha"
50220
50230  JPEGColorscount(0) = "jpeg": JPEGColorscount(1) = "jpeggray"
50240
50250  BMPColorscount(0) = "bmp32b": BMPColorscount(1) = "bmp16m"
50260  BMPColorscount(2) = "bmp256": BMPColorscount(3) = "bmp16"
50270  BMPColorscount(4) = "bmpsep8": BMPColorscount(5) = "bmpsep1"
50280  BMPColorscount(6) = "bmpgray": BMPColorscount(7) = "bmpmono"
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
50460  With Options
50470  'General
50480   GS_PDFDEFAULT = PDFDefaultSettings(.PDFGeneralDefault)
50490   GS_COMPATIBILITY = "1." & (.PDFGeneralCompatibility + 2)
50500   GS_RESOLUTION = .PDFGeneralResolution
50510   GS_AUTOROTATE = Rotate(.PDFGeneralAutorotate)
50520   GS_OVERPRINT = .PDFGeneralOverprint
50530   GS_ASCII85 = Bool2Text(.PDFGeneralASCII85)
50540
50550   'Compression
50560   GS_COMPRESSPAGES = Bool2Text(.PDFCompressionTextCompression)
50570   GS_COMPRESSCOLOR = Bool2Text(.PDFCompressionColorCompression)
50580   GS_COMPRESSGREY = Bool2Text(.PDFCompressionGreyCompression)
50590   GS_COMPRESSMONO = Bool2Text(.PDFCompressionMonoCompression)
50600
50610   SelectColorCompression .PDFCompressionColorCompressionChoice
50620   SelectGreyCompression .PDFCompressionGreyCompressionChoice
50630   SelectMonoCompression .PDFCompressionMonoCompressionChoice
50640
50650   GS_COMPRESSCOLORVALUE = Bool2Text(.PDFCompressionColorCompression)
50660   GS_COMPRESSGREYVALUE = Bool2Text(.PDFCompressionGreyCompression)
50670   GS_COMPRESSMONOVALUE = Bool2Text(.PDFCompressionMonoCompression)
50680
50690   GS_COLORRESOLUTION = .PDFCompressionColorResolution
50700   GS_GREYRESOLUTION = .PDFCompressionGreyResolution
50710   GS_MONORESOLUTION = .PDFCompressionMonoResolution
50720
50730   GS_COLORRESAMPLE = Bool2Text(.PDFCompressionColorResample)
50740   GS_GREYRESAMPLE = Bool2Text(.PDFCompressionGreyResample)
50750   GS_MONORESAMPLE = Bool2Text(.PDFCompressionMonoResample)
50760
50770   GS_COLORRESAMPLEMETHOD = Resample(.PDFCompressionColorResampleChoice)
50780   GS_GREYRESAMPLEMETHOD = Resample(.PDFCompressionGreyResampleChoice)
50790   GS_MONORESAMPLEMETHOD = Resample(.PDFCompressionMonoResampleChoice)
50800
50810   'Fonts
50820   GS_EMBEDALLFONTS = Bool2Text(.PDFFontsEmbedAll)
50830   GS_SUBSETFONTS = Bool2Text(.PDFFontsSubSetFonts)
50840   GS_SUBSETFONTPERC = .PDFFontsSubSetFontsPercent
50850
50860   'Colors
50870   GS_COLORMODEL = Colormodel(.PDFColorsColorModel)
50880   GS_CMYKTORGB = Bool2Text(.PDFColorsCMYKToRGB)
50890   GS_PRESERVEOVERPRINT = Bool2Text(.PDFColorsPreserveOverprint)
50900   GS_TRANSFERFUNCTIONS = ColorsPreserveTransfer(.PDFColorsPreserveTransfer)
50910   GS_HALFTONE = Bool2Text(.PDFColorsPreserveHalftone)
50920
50930   'Bitmap
50940   GS_PNGColorscount = PNGColorscount(.PNGColorscount)
50950   GS_JPEGColorscount = JPEGColorscount(.JPEGColorscount)
50960   GS_BMPColorscount = BMPColorscount(.BMPColorscount)
50970   GS_PCXColorscount = PCXColorscount(.PCXColorscount)
50980   GS_TIFFColorscount = TIFFColorscount(.TIFFColorscount)
50990   GS_JPEGQuality = .JPEGQuality
51000   GS_PSLanguageLevel = PSLanguageLevel(.PSLanguageLevel)
51010   GS_EPSLanguageLevel = PSLanguageLevel(.EPSLanguageLevel)
51020   GS_PSDColorscount = PSDColorsCount(.PSDColorsCount)
51030   GS_PCLColorscount = PCLColorsCount(.PCLColorsCount)
51040   GS_RAWColorscount = RAWColorsCount(.RAWColorsCount)
51050  End With
51060  'Other
51070  GS_ERROR = 0
51080  UseReturnPipe = 1
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "GSInit")
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
Select Case ErrPtnr.OnError("modGhostScript", "CallGhostscript")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePDF(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim FName As String, tStr As String, encPDF As EncryptData, tEnc As Boolean
50020
50030  InitParams
50040  Set ParamCommands = New Collection
50050
50060  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50070
50080  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50090   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50100  End If
50110  AddParams "-I" & tStr
50120  AddParams "-q"
50130  AddParams "-dNOPAUSE"
50140  'AddParams "-dSAFER"
50150  AddParams "-dBATCH"
50160  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50170   AddParams "-sFONTPATH=" & GetFontsDirectory
50180  End If
50190  AddParams "-sDEVICE=pdfwrite"
50200  If Options.DontUseDocumentSettings = 0 Then
50210   AddParams "-dPDFSETTINGS=/" & GS_PDFDEFAULT
50220   AddParams "-dCompatibilityLevel=" & GS_COMPATIBILITY
50230   AddParams "-r" & GS_RESOLUTION & "x" & GS_RESOLUTION
50240   AddParams "-dProcessColorModel=/Device" & GS_COLORMODEL
50250   AddParams "-dAutoRotatePages=/" & GS_AUTOROTATE
50260   AddParams "-dCompressPages=" & GS_COMPRESSPAGES
50270   AddParams "-dEmbedAllFonts=" & GS_EMBEDALLFONTS
50280   AddParams "-dSubsetFonts=" & GS_SUBSETFONTS
50290   AddParams "-dMaxSubsetPct=" & GS_SUBSETFONTPERC
50300   AddParams "-dConvertCMYKImagesToRGB=" & GS_CMYKTORGB
50310
50320   If Options.UseFixPapersize <> 0 Then
50330    If Options.UseCustomPaperSize = 0 Then
50340      If LenB(Trim$(Options.Papersize)) > 0 Then
50350       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50360       AddParams "-dFIXEDMEDIA"
50370       AddParams "-dNORANGEPAGESIZE"
50380      End If
50390     Else
50400      If Options.DeviceWidthPoints >= 1 Then
50410       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50420      End If
50430      If Options.DeviceHeightPoints >= 1 Then
50440       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50450      End If
50460    End If
50470   End If
50480  End If
50490  tEnc = False
50500  If Options.PDFOptimize = 0 And Options.PDFUseSecurity <> 0 And SecurityIsPossible = True Then
50510   If Options.PDFAes128Encryption = 0 Then
50520    If SetEncryptionParams(encPDF, "", "") = True Then
50530     tEnc = True
50540     If Len(encPDF.OwnerPass) > 0 Then
50550       AddParams "-sOwnerPassword=" & encPDF.OwnerPass: currentOwnerPassword = encPDF.OwnerPass
50560      Else
50570       If Len(encPDF.UserPass) > 0 Then
50580        AddParams "-sOwnerPassword=" & encPDF.UserPass: currentOwnerPassword = encPDF.OwnerPass
50590       End If
50600     End If
50610     If Len(encPDF.UserPass) > 0 Then
50620      AddParams "-sUserPassword=" & encPDF.UserPass
50630     End If
50640     AddParams "-dPermissions=" & CalculatePermissions(encPDF)
50650     If GS_COMPATIBILITY = "1.4" Or GS_COMPATIBILITY = "1.5" Then
50660       AddParams "-dEncryptionR=3"
50670       AddParams "-dKeyLength=128"
50680      Else
50690       AddParams "-dEncryptionR=2"
50700       AddParams "-dKeyLength=40"
50710     End If
50720    Else
50730     If Options.UseAutosave = 0 Then
50740      MsgBox LanguageStrings.MessagesMsg23, vbCritical
50750     End If
50760    End If
50770   End If
50780  End If
50790
50800  If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50810   GSOutputFile = Replace$(GSOutputFile, "%", "%%")
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
51110  If LenB(StampFile) > 0 Then
51120   If FileExists(StampFile) Then
51130    AddParams StampFile
51140   End If
51150  End If
51160  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
51170  Set isf = New clsInfoSpoolFile
51180  isf.ReadInfoFile InfoSpoolFileName
51190  For i = 1 To isf.InfoFiles.Count
51200   Set isfi = isf.InfoFiles(i)
51210   If FileExists(isfi.SpoolFileName) Then
51220    AddParams isfi.SpoolFileName
51230   End If
51240  Next i
51250  If LenB(PDFDocInfoFile) > 0 Then
51260   If FileExists(PDFDocInfoFile) Then
51270    AddParams PDFDocInfoFile
51280   End If
51290  End If
51300
51310  ShowParams
51320  If tEnc = True Then
51330    CallGhostscript "PDF with encryption"
51340   Else
51350    CallGhostscript "PDF without encryption"
51360  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreatePDF")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePNG(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
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
50110
50120  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50130
50140  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50150   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50160  End If
50170  AddParams "-I" & tStr
50180  AddParams "-q"
50190  AddParams "-dNOPAUSE"
50200  'AddParams "-dSAFER"
50210  AddParams "-dBATCH"
50220  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50230   AddParams "-sFONTPATH=" & GetFontsDirectory
50240  End If
50250
50260  AddParams "-sDEVICE=" & GS_PNGColorscount
50270
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
50450   AddParams "-r" & Options.PNGResolution & "x" & Options.PNGResolution
50460
50470   If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50480    GSOutputFile = Replace$(GSOutputFile, "%", "%%")
50490   End If
50500   AddParams "-sOutputFile=" & GSOutputFile
50510  End If
50520
50530  AddAdditionalGhostscriptParameters
50540
50550  AddParams "-f"
50560  If LenB(StampFile) > 0 Then
50570   If FileExists(StampFile) Then
50580    AddParams StampFile
50590   End If
50600  End If
50610  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50620  Set isf = New clsInfoSpoolFile
50630  isf.ReadInfoFile InfoSpoolFileName
50640  For i = 1 To isf.InfoFiles.Count
50650   Set isfi = isf.InfoFiles(i)
50660   If FileExists(isfi.SpoolFileName) Then
50670    AddParams isfi.SpoolFileName
50680   End If
50690  Next i
50700  If LenB(PDFDocInfoFile) > 0 Then
50710   If FileExists(PDFDocInfoFile) Then
50720    AddParams PDFDocInfoFile
50730   End If
50740  End If
50750
50760  ShowParams
50770  CallGhostscript "PNG"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreatePNG")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateJPEG(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
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
50110
50120  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50130
50140  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50150   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50160  End If
50170  AddParams "-I" & tStr
50180  AddParams "-q"
50190  AddParams "-dNOPAUSE"
50200  'AddParams "-dSAFER"
50210  AddParams "-dBATCH"
50220  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50230   AddParams "-sFONTPATH=" & GetFontsDirectory
50240  End If
50250
50260  AddParams "-sDEVICE=" & GS_JPEGColorscount
50270  If Options.DontUseDocumentSettings = 0 Then
50280   AddParams "-dJPEGQ=" & GS_JPEGQuality
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
50450   AddParams "-r" & Options.JPEGResolution & "x" & Options.JPEGResolution
50460
50470   If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50480    GSOutputFile = Replace$(GSOutputFile, "%", "%%")
50490   End If
50500   AddParams "-sOutputFile=" & GSOutputFile
50510  End If
50520
50530  AddAdditionalGhostscriptParameters
50540
50550  AddParams "-f"
50560  If LenB(StampFile) > 0 Then
50570   If FileExists(StampFile) Then
50580    AddParams StampFile
50590   End If
50600  End If
50610  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50620  Set isf = New clsInfoSpoolFile
50630  isf.ReadInfoFile InfoSpoolFileName
50640  For i = 1 To isf.InfoFiles.Count
50650   Set isfi = isf.InfoFiles(i)
50660   If FileExists(isfi.SpoolFileName) Then
50670    AddParams isfi.SpoolFileName
50680   End If
50690  Next i
50700  If LenB(PDFDocInfoFile) > 0 Then
50710   If FileExists(PDFDocInfoFile) Then
50720    AddParams PDFDocInfoFile
50730   End If
50740  End If
50750
50760  ShowParams
50770  CallGhostscript "JPEG"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreateJPEG")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateBMP(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
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
50110
50120  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50130
50140  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50150   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50160  End If
50170  AddParams "-I" & tStr
50180  AddParams "-q"
50190  AddParams "-dNOPAUSE"
50200  'AddParams "-dSAFER"
50210  AddParams "-dBATCH"
50220  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50230   AddParams "-sFONTPATH=" & GetFontsDirectory
50240  End If
50250
50260  AddParams "-sDEVICE=" & GS_BMPColorscount
50270
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
50450   AddParams "-r" & Options.BMPResolution & "x" & Options.BMPResolution
50460  End If
50470
50480  If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50490   GSOutputFile = Replace$(GSOutputFile, "%", "%%")
50500  End If
50510  AddParams "-sOutputFile=" & GSOutputFile
50520
50530  AddAdditionalGhostscriptParameters
50540
50550  AddParams "-f"
50560  If LenB(StampFile) > 0 Then
50570   If FileExists(StampFile) Then
50580    AddParams StampFile
50590   End If
50600  End If
50610  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50620  Set isf = New clsInfoSpoolFile
50630  isf.ReadInfoFile InfoSpoolFileName
50640  For i = 1 To isf.InfoFiles.Count
50650   Set isfi = isf.InfoFiles(i)
50660   If FileExists(isfi.SpoolFileName) Then
50670    AddParams isfi.SpoolFileName
50680   End If
50690  Next i
50700  If LenB(PDFDocInfoFile) > 0 Then
50710   If FileExists(PDFDocInfoFile) Then
50720    AddParams PDFDocInfoFile
50730   End If
50740  End If
50750
50760  ShowParams
50770  CallGhostscript "BMP"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreateBMP")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePCX(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
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
50110
50120  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50130
50140  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50150   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50160  End If
50170  AddParams "-I" & tStr
50180  AddParams "-q"
50190  AddParams "-dNOPAUSE"
50200  'AddParams "-dSAFER"
50210  AddParams "-dBATCH"
50220  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50230   AddParams "-sFONTPATH=" & GetFontsDirectory
50240  End If
50250
50260  AddParams "-sDEVICE=" & GS_PCXColorscount
50270  If Options.DontUseDocumentSettings = 0 Then
50280   If Options.UseFixPapersize <> 0 Then
50290    If Options.UseCustomPaperSize = 0 Then
50300      If LenB(Trim$(Options.Papersize)) > 0 Then
50310       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50320       AddParams "-dFIXEDMEDIA"
50330       AddParams "-dNORANGEPAGESIZE"
50340      End If
50350     Else
50360      If Options.DeviceWidthPoints >= 1 Then
50370       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50380      End If
50390      If Options.DeviceHeightPoints >= 1 Then
50400       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50410      End If
50420    End If
50430   End If
50440   AddParams "-r" & Options.PCXResolution & "x" & Options.PCXResolution
50450  End If
50460
50470  If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50480   GSOutputFile = Replace$(GSOutputFile, "%", "%%")
50490  End If
50500  AddParams "-sOutputFile=" & GSOutputFile
50510
50520  AddAdditionalGhostscriptParameters
50530
50540  AddParams "-f"
50550  If LenB(StampFile) > 0 Then
50560   If FileExists(StampFile) Then
50570    AddParams StampFile
50580   End If
50590  End If
50600  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50610  Set isf = New clsInfoSpoolFile
50620  isf.ReadInfoFile InfoSpoolFileName
50630  For i = 1 To isf.InfoFiles.Count
50640   Set isfi = isf.InfoFiles(i)
50650   If FileExists(isfi.SpoolFileName) Then
50660    AddParams isfi.SpoolFileName
50670   End If
50680  Next i
50690  If LenB(PDFDocInfoFile) > 0 Then
50700   If FileExists(PDFDocInfoFile) Then
50710    AddParams PDFDocInfoFile
50720   End If
50730  End If
50740
50750  ShowParams
50760  CallGhostscript "PCX"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreatePCX")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateTIFF(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
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
50110
50120  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50130
50140  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50150   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50160  End If
50170  AddParams "-I" & tStr
50180  AddParams "-q"
50190  AddParams "-dNOPAUSE"
50200  'AddParams "-dSAFER"
50210  AddParams "-dBATCH"
50220  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50230   AddParams "-sFONTPATH=" & GetFontsDirectory
50240  End If
50250
50260  AddParams "-sDEVICE=" & GS_TIFFColorscount
50270  If Options.DontUseDocumentSettings = 0 Then
50280   If Options.UseFixPapersize <> 0 Then
50290    If Options.UseCustomPaperSize = 0 Then
50300      If LenB(Trim$(Options.Papersize)) > 0 Then
50310       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50320       AddParams "-dFIXEDMEDIA"
50330       AddParams "-dNORANGEPAGESIZE"
50340      End If
50350     Else
50360      If Options.DeviceWidthPoints >= 1 Then
50370       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50380      End If
50390      If Options.DeviceHeightPoints >= 1 Then
50400       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50410      End If
50420    End If
50430   End If
50440   AddParams "-r" & Options.TIFFResolution & "x" & Options.TIFFResolution
50450  End If
50460
50470  If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50480   GSOutputFile = Replace$(GSOutputFile, "%", "%%")
50490  End If
50500  AddParams "-sOutputFile=" & GSOutputFile
50510
50520  AddAdditionalGhostscriptParameters
50530
50540  AddParams "-f"
50550  If LenB(StampFile) > 0 Then
50560   If FileExists(StampFile) Then
50570    AddParams StampFile
50580   End If
50590  End If
50600  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50610  Set isf = New clsInfoSpoolFile
50620  isf.ReadInfoFile InfoSpoolFileName
50630  For i = 1 To isf.InfoFiles.Count
50640   Set isfi = isf.InfoFiles(i)
50650   If FileExists(isfi.SpoolFileName) Then
50660    AddParams isfi.SpoolFileName
50670   End If
50680  Next i
50690  If LenB(PDFDocInfoFile) > 0 Then
50700   If FileExists(PDFDocInfoFile) Then
50710    AddParams PDFDocInfoFile
50720   End If
50730  End If
50740
50750  ShowParams
50760  CallGhostscript "TIFF"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreateTIFF")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePS(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
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
50110
50120  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50130
50140  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50150   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50160  End If
50170  AddParams "-I" & tStr
50180  AddParams "-q"
50190  AddParams "-dNOPAUSE"
50200  'AddParams "-dSAFER"
50210  AddParams "-dBATCH"
50220  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50230   AddParams "-sFONTPATH=" & GetFontsDirectory
50240  End If
50250
50260  AddParams "-sDEVICE=pswrite"
50270  If Options.DontUseDocumentSettings = 0 Then
50280   If Options.UseFixPapersize <> 0 Then
50290    If Options.UseCustomPaperSize = 0 Then
50300      If LenB(Trim$(Options.Papersize)) > 0 Then
50310       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50320       AddParams "-dFIXEDMEDIA"
50330       AddParams "-dNORANGEPAGESIZE"
50340      End If
50350     Else
50360      If Options.DeviceWidthPoints >= 1 Then
50370       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50380      End If
50390      If Options.DeviceHeightPoints >= 1 Then
50400       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50410      End If
50420    End If
50430   End If
50440   AddParams "-dLanguageLevel=" & GS_PSLanguageLevel
50450  End If
50460
50470  If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50480   GSOutputFile = Replace$(GSOutputFile, "%", "%%")
50490  End If
50500  AddParams "-sOutputFile=" & GSOutputFile
50510
50520  AddAdditionalGhostscriptParameters
50530
50540  AddParams "-f"
50550  If LenB(StampFile) > 0 Then
50560   If FileExists(StampFile) Then
50570    AddParams StampFile
50580   End If
50590  End If
50600  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50610  Set isf = New clsInfoSpoolFile
50620  isf.ReadInfoFile InfoSpoolFileName
50630  For i = 1 To isf.InfoFiles.Count
50640   Set isfi = isf.InfoFiles(i)
50650   If FileExists(isfi.SpoolFileName) Then
50660    AddParams isfi.SpoolFileName
50670   End If
50680  Next i
50690  If LenB(PDFDocInfoFile) > 0 Then
50700   If FileExists(PDFDocInfoFile) Then
50710    AddParams PDFDocInfoFile
50720   End If
50730  End If
50740
50750  ShowParams
50760  CallGhostscript "PS"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreatePS")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateEPS(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
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
50110
50120  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50130
50140  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50150   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50160  End If
50170  AddParams "-I" & tStr
50180  AddParams "-q"
50190  AddParams "-dNOPAUSE"
50200  'AddParams "-dSAFER"
50210  AddParams "-dBATCH"
50220  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50230   AddParams "-sFONTPATH=" & GetFontsDirectory
50240  End If
50250
50260  AddParams "-sDEVICE=epswrite"
50270  If Options.DontUseDocumentSettings = 0 Then
50280   If Options.UseFixPapersize <> 0 Then
50290    If Options.UseCustomPaperSize = 0 Then
50300      If LenB(Trim$(Options.Papersize)) > 0 Then
50310       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50320       AddParams "-dFIXEDMEDIA"
50330       AddParams "-dNORANGEPAGESIZE"
50340      End If
50350     Else
50360      If Options.DeviceWidthPoints >= 1 Then
50370       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50380      End If
50390      If Options.DeviceHeightPoints >= 1 Then
50400       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50410      End If
50420    End If
50430   End If
50440   AddParams "-dLanguageLevel=" & GS_EPSLanguageLevel
50450  End If
50460
50470  If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50480   GSOutputFile = Replace$(GSOutputFile, "%", "%%")
50490  End If
50500  AddParams "-sOutputFile=" & GSOutputFile
50510
50520  AddAdditionalGhostscriptParameters
50530
50540  AddParams "-f"
50550  If LenB(StampFile) > 0 Then
50560   If FileExists(StampFile) Then
50570    AddParams StampFile
50580   End If
50590  End If
50600  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50610  Set isf = New clsInfoSpoolFile
50620  isf.ReadInfoFile InfoSpoolFileName
50630  For i = 1 To isf.InfoFiles.Count
50640   Set isfi = isf.InfoFiles(i)
50650   If FileExists(isfi.SpoolFileName) Then
50660    AddParams isfi.SpoolFileName
50670   End If
50680  Next i
50690  If LenB(PDFDocInfoFile) > 0 Then
50700   If FileExists(PDFDocInfoFile) Then
50710    AddParams PDFDocInfoFile
50720   End If
50730  End If
50740
50750  ShowParams
50760  CallGhostscript "EPS"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreateEPS")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateTXT(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, FName As String, Ext As String, tStr As String
50020
50030  GSInit Options
50040  InitParams
50050  Set ParamCommands = New Collection
50060
50070  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50080
50090  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50100   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50110  End If
50120  AddParams "-I" & tStr
50130  AddParams "-q"
50140  AddParams "-dNOPAUSE"
50150  'AddParams "-dSAFER"
50160  AddParams "-dBATCH"
50170  AddParams "-dNODISPLAY"
50180  AddParams "-dDELAYBIND"
50190  AddParams "-dWRITESYSTEMDICT"
50200  AddParams "-dSIMPLE"
50210  AddParams "ps2ascii.ps"
50220
50230  ' TODO: Testen
50240  If LenB(StampFile) > 0 Then
50250   If FileExists(StampFile) Then
50260    AddParams StampFile
50270   End If
50280  End If
50290  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50300  Set isf = New clsInfoSpoolFile
50310  isf.ReadInfoFile InfoSpoolFileName
50320  For i = 1 To isf.InfoFiles.Count
50330   Set isfi = isf.InfoFiles(i)
50340   If FileExists(isfi.SpoolFileName) Then
50350    AddParams isfi.SpoolFileName
50360   End If
50370  Next i
50380  If LenB(PDFDocInfoFile) > 0 Then
50390   If FileExists(PDFDocInfoFile) Then
50400    AddParams PDFDocInfoFile
50410   End If
50420  End If
50430
50440  ShowParams
50450  CallGhostscript "TXT"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreateTXT")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePDFA(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim FName As String, tStr As String, encPDF As EncryptData, tEnc As Boolean
50020
50030  InitParams
50040  Set ParamCommands = New Collection
50050
50060  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50070
50080  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50090   If DirExists(Options.AdditionalGhostscriptSearchpath) Then
50100    tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50110   End If
50120  End If
50130  AddParams "-I" & tStr
50140  AddParams "-dPDFA"
50150  AddParams "-q"
50160  AddParams "-dNOPAUSE"
50170  'AddParams "-dSAFER"
50180  AddParams "-dBATCH"
50190  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50200   AddParams "-sFONTPATH=" & GetFontsDirectory
50210  End If
50220  AddParams "-sDEVICE=pdfwrite"
50230  If Options.DontUseDocumentSettings = 0 Then
50240   AddParams "-dPDFSETTINGS=/" & GS_PDFDEFAULT
50250   AddParams "-dCompatibilityLevel=" & GS_COMPATIBILITY
50260   AddParams "-dNOOUTERSAVE"
50270   AddParams "-dUseCIEColor"
50280   AddParams "-r" & GS_RESOLUTION & "x" & GS_RESOLUTION
50290   AddParams "-sProcessColorModel=Device" & GS_COLORMODEL
50300   AddParams "-dAutoRotatePages=/" & GS_AUTOROTATE
50310   AddParams "-dCompressPages=" & GS_COMPRESSPAGES
50320   AddParams "-dEmbedAllFonts=true"
50330   AddParams "-dSubsetFonts=" & GS_SUBSETFONTS
50340   AddParams "-dMaxSubsetPct=100"
50350   AddParams "-dConvertCMYKImagesToRGB=" & GS_CMYKTORGB
50360
50370   If Options.UseFixPapersize <> 0 Then
50380    If Options.UseCustomPaperSize = 0 Then
50390      If LenB(Trim$(Options.Papersize)) > 0 Then
50400       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50410       AddParams "-dFIXEDMEDIA"
50420       AddParams "-dNORANGEPAGESIZE"
50430      End If
50440     Else
50450      If Options.DeviceWidthPoints >= 1 Then
50460       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50470      End If
50480      If Options.DeviceHeightPoints >= 1 Then
50490       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50500      End If
50510    End If
50520   End If
50530
50540  End If
50550  tEnc = False
50560
50570  If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50580   GSOutputFile = Replace$(GSOutputFile, "%", "%%")
50590  End If
50600  AddParams "-sOutputFile=" & GSOutputFile
50610
50620  If Options.DontUseDocumentSettings = 0 Then
50630   SetColorParams
50640   SetGreyParams
50650   SetMonoParams
50660
50670   AddParams "-dPreserveOverprintSettings=" & GS_PRESERVEOVERPRINT
50680   AddParams "-dUCRandBGInfo=/Preserve"
50690   AddParams "-dUseFlateCompression=true"
50700   AddParams "-dParseDSCCommentsForDocInfo=true"
50710   AddParams "-dParseDSCComments=true"
50720   AddParams "-dOPM=1" '& GS_OVERPRINT
50730   AddParams "-dOffOptimizations=0"
50740   AddParams "-dLockDistillerParams=false"
50750   AddParams "-dGrayImageDepth=-1"
50760   AddParams "-dASCII85EncodePages=" & GS_ASCII85
50770   AddParams "-dDefaultRenderingIntent=/Default"
50780   AddParams "-dTransferFunctionInfo=/" & GS_TRANSFERFUNCTIONS
50790   AddParams "-dPreserveHalftoneInfo=" & GS_HALFTONE
50800   AddParams "-dDetectBlends=true"
50810
50820   AddAdditionalGhostscriptParameters
50830
50840   AddParams "-f"
50850   AddParams CompletePath(Options.DirectoryGhostscriptLibraries) + "pdfa_def.ps"
50860
50870   AddParamCommands
50880  End If
50890
50900  AddParams "-f"
50910  If LenB(StampFile) > 0 Then
50920   If FileExists(StampFile) Then
50930    AddParams StampFile
50940   End If
50950  End If
50960  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50970  Set isf = New clsInfoSpoolFile
50980  isf.ReadInfoFile InfoSpoolFileName
50990  For i = 1 To isf.InfoFiles.Count
51000   Set isfi = isf.InfoFiles(i)
51010   If FileExists(isfi.SpoolFileName) Then
51020    AddParams isfi.SpoolFileName
51030   End If
51040  Next i
51050  If LenB(PDFDocInfoFile) > 0 Then
51060   If FileExists(PDFDocInfoFile) Then
51070    AddParams PDFDocInfoFile
51080   End If
51090  End If
51100
51110  ShowParams
51120  CallGhostscript "PDF/A (without encryption)"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreatePDFA")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePDFX(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim FName As String, tStr As String, encPDF As EncryptData, tEnc As Boolean
50020
50030  InitParams
50040  Set ParamCommands = New Collection
50050
50060  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50070
50080  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50090   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50100  End If
50110  AddParams "-I" & tStr
50120  AddParams "-q"
50130  AddParams "-dNOPAUSE"
50140  'AddParams "-dSAFER"
50150  AddParams "-dBATCH"
50160  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50170   AddParams "-sFONTPATH=" & GetFontsDirectory
50180  End If
50190  AddParams "-sDEVICE=pdfwrite"
50200  If Options.DontUseDocumentSettings = 0 Then
50210   AddParams "-dPDFSETTINGS=/" & GS_PDFDEFAULT
50220   AddParams "-dCompatibilityLevel=" & GS_COMPATIBILITY
50230   AddParams "-r" & GS_RESOLUTION & "x" & GS_RESOLUTION
50240   AddParams "-dProcessColorModel=/Device" & GS_COLORMODEL
50250   AddParams "-dAutoRotatePages=/" & GS_AUTOROTATE
50260   AddParams "-dCompressPages=" & GS_COMPRESSPAGES
50270   AddParams "-dEmbedAllFonts=" & GS_EMBEDALLFONTS
50280   AddParams "-dSubsetFonts=" & GS_SUBSETFONTS
50290   AddParams "-dMaxSubsetPct=" & GS_SUBSETFONTPERC
50300   AddParams "-dConvertCMYKImagesToRGB=" & GS_CMYKTORGB
50310
50320   If Options.UseFixPapersize <> 0 Then
50330    If Options.UseCustomPaperSize = 0 Then
50340      If LenB(Trim$(Options.Papersize)) > 0 Then
50350       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50360       AddParams "-dFIXEDMEDIA"
50370       AddParams "-dNORANGEPAGESIZE"
50380      End If
50390     Else
50400      If Options.DeviceWidthPoints >= 1 Then
50410       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50420      End If
50430      If Options.DeviceHeightPoints >= 1 Then
50440       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50450      End If
50460    End If
50470   End If
50480
50490  End If
50500  tEnc = False
50510
50520  If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50530   GSOutputFile = Replace$(GSOutputFile, "%", "%%")
50540  End If
50550  AddParams "-sOutputFile=" & GSOutputFile
50560
50570  If Options.DontUseDocumentSettings = 0 Then
50580   SetColorParams
50590   SetGreyParams
50600   SetMonoParams
50610
50620   AddParams "-dPreserveOverprintSettings=" & GS_PRESERVEOVERPRINT
50630   AddParams "-dUCRandBGInfo=/Preserve"
50640   AddParams "-dUseFlateCompression=true"
50650   AddParams "-dParseDSCCommentsForDocInfo=true"
50660   AddParams "-dParseDSCComments=true"
50670   AddParams "-dOPM=" & GS_OVERPRINT
50680   AddParams "-dOffOptimizations=0"
50690   AddParams "-dLockDistillerParams=false"
50700   AddParams "-dGrayImageDepth=-1"
50710   AddParams "-dASCII85EncodePages=" & GS_ASCII85
50720   AddParams "-dDefaultRenderingIntent=/Default"
50730   AddParams "-dTransferFunctionInfo=/" & GS_TRANSFERFUNCTIONS
50740   AddParams "-dPreserveHalftoneInfo=" & GS_HALFTONE
50750   AddParams "-dDetectBlends=true"
50760
50770   AddAdditionalGhostscriptParameters
50780
50790   AddParams "-dPDFX"
50800   AddParams "-f"
50810   AddParams CompletePath(Options.DirectoryGhostscriptLibraries) + "pdfx_def.ps"
50820
50830   AddParamCommands
50840  End If
50850
50860  AddParams "-f"
50870  If LenB(StampFile) > 0 Then
50880   If FileExists(StampFile) Then
50890    AddParams StampFile
50900   End If
50910  End If
50920  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50930  Set isf = New clsInfoSpoolFile
50940  isf.ReadInfoFile InfoSpoolFileName
50950  For i = 1 To isf.InfoFiles.Count
50960   Set isfi = isf.InfoFiles(i)
50970   If FileExists(isfi.SpoolFileName) Then
50980    AddParams isfi.SpoolFileName
50990   End If
51000  Next i
51010  If LenB(PDFDocInfoFile) > 0 Then
51020   If FileExists(PDFDocInfoFile) Then
51030    AddParams PDFDocInfoFile
51040   End If
51050  End If
51060
51070  ShowParams
51080  CallGhostscript "PDF/X (without encryption)"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreatePDFX")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePSD(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
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
50110
50120  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50130
50140  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50150   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50160  End If
50170  AddParams "-I" & tStr
50180  AddParams "-q"
50190  AddParams "-dNOPAUSE"
50200  'AddParams "-dSAFER"
50210  AddParams "-dBATCH"
50220  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50230   AddParams "-sFONTPATH=" & GetFontsDirectory
50240  End If
50250
50260  AddParams "-sDEVICE=" & GS_PSDColorscount
50270  If Options.DontUseDocumentSettings = 0 Then
50280   If Options.UseFixPapersize <> 0 Then
50290    If Options.UseCustomPaperSize = 0 Then
50300      If LenB(Trim$(Options.Papersize)) > 0 Then
50310       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50320       AddParams "-dFIXEDMEDIA"
50330       AddParams "-dNORANGEPAGESIZE"
50340      End If
50350     Else
50360      If Options.DeviceWidthPoints >= 1 Then
50370       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50380      End If
50390      If Options.DeviceHeightPoints >= 1 Then
50400       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50410      End If
50420    End If
50430   End If
50440   AddParams "-r" & Options.PSDResolution & "x" & Options.PSDResolution
50450  End If
50460
50470  If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50480   GSOutputFile = Replace$(GSOutputFile, "%", "%%")
50490  End If
50500  AddParams "-sOutputFile=" & GSOutputFile
50510
50520  AddAdditionalGhostscriptParameters
50530
50540  AddParams "-f"
50550  If LenB(StampFile) > 0 Then
50560   If FileExists(StampFile) Then
50570    AddParams StampFile
50580   End If
50590  End If
50600  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50610  Set isf = New clsInfoSpoolFile
50620  isf.ReadInfoFile InfoSpoolFileName
50630  For i = 1 To isf.InfoFiles.Count
50640   Set isfi = isf.InfoFiles(i)
50650   If FileExists(isfi.SpoolFileName) Then
50660    AddParams isfi.SpoolFileName
50670   End If
50680  Next i
50690  If LenB(PDFDocInfoFile) > 0 Then
50700   If FileExists(PDFDocInfoFile) Then
50710    AddParams PDFDocInfoFile
50720   End If
50730  End If
50740
50750  ShowParams
50760  CallGhostscript "PSD"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreatePSD")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePCL(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
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
50110
50120  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50130
50140  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50150   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50160  End If
50170  AddParams "-I" & tStr
50180  AddParams "-q"
50190  AddParams "-dNOPAUSE"
50200  'AddParams "-dSAFER"
50210  AddParams "-dBATCH"
50220  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50230   AddParams "-sFONTPATH=" & GetFontsDirectory
50240  End If
50250
50260  AddParams "-sDEVICE=" & GS_PCLColorscount
50270  If Options.DontUseDocumentSettings = 0 Then
50280   If Options.UseFixPapersize <> 0 Then
50290    If Options.UseCustomPaperSize = 0 Then
50300      If LenB(Trim$(Options.Papersize)) > 0 Then
50310       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50320       AddParams "-dFIXEDMEDIA"
50330       AddParams "-dNORANGEPAGESIZE"
50340      End If
50350     Else
50360      If Options.DeviceWidthPoints >= 1 Then
50370       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50380      End If
50390      If Options.DeviceHeightPoints >= 1 Then
50400       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50410      End If
50420    End If
50430   End If
50440   AddParams "-r" & Options.PCLResolution & "x" & Options.PCLResolution
50450  End If
50460
50470  If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50480   GSOutputFile = Replace$(GSOutputFile, "%", "%%")
50490  End If
50500  AddParams "-sOutputFile=" & GSOutputFile
50510
50520  AddAdditionalGhostscriptParameters
50530
50540  AddParams "-f"
50550  If LenB(StampFile) > 0 Then
50560   If FileExists(StampFile) Then
50570    AddParams StampFile
50580   End If
50590  End If
50600  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50610  Set isf = New clsInfoSpoolFile
50620  isf.ReadInfoFile InfoSpoolFileName
50630  For i = 1 To isf.InfoFiles.Count
50640   Set isfi = isf.InfoFiles(i)
50650   If FileExists(isfi.SpoolFileName) Then
50660    AddParams isfi.SpoolFileName
50670   End If
50680  Next i
50690  If LenB(PDFDocInfoFile) > 0 Then
50700   If FileExists(PDFDocInfoFile) Then
50710    AddParams PDFDocInfoFile
50720   End If
50730  End If
50740
50750  ShowParams
50760  CallGhostscript "PCL"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreatePCL")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateRAW(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
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
50110
50120  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50130
50140  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50150   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50160  End If
50170  AddParams "-I" & tStr
50180  AddParams "-q"
50190  AddParams "-dNOPAUSE"
50200  'AddParams "-dSAFER"
50210  AddParams "-dBATCH"
50220  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50230   AddParams "-sFONTPATH=" & GetFontsDirectory
50240  End If
50250
50260  AddParams "-sDEVICE=" & GS_RAWColorscount
50270  If Options.DontUseDocumentSettings = 0 Then
50280   If Options.UseFixPapersize <> 0 Then
50290    If Options.UseCustomPaperSize = 0 Then
50300      If LenB(Trim$(Options.Papersize)) > 0 Then
50310       AddParams "-sPAPERSIZE=" & LCase$(Trim$(Options.Papersize))
50320       AddParams "-dFIXEDMEDIA"
50330       AddParams "-dNORANGEPAGESIZE"
50340      End If
50350     Else
50360      If Options.DeviceWidthPoints >= 1 Then
50370       AddParams "-dDEVICEWIDTHPOINTS=" & Options.DeviceWidthPoints
50380      End If
50390      If Options.DeviceHeightPoints >= 1 Then
50400       AddParams "-dDEVICEHEIGHTPOINTS=" & Options.DeviceHeightPoints
50410      End If
50420    End If
50430   End If
50440   AddParams "-r" & Options.RAWResolution & "x" & Options.RAWResolution
50450  End If
50460
50470  If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50480   GSOutputFile = Replace$(GSOutputFile, "%", "%%")
50490  End If
50500  AddParams "-sOutputFile=" & GSOutputFile
50510
50520  AddAdditionalGhostscriptParameters
50530
50540  AddParams "-f"
50550  If LenB(StampFile) > 0 Then
50560   If FileExists(StampFile) Then
50570    AddParams StampFile
50580   End If
50590  End If
50600  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50610  Set isf = New clsInfoSpoolFile
50620  isf.ReadInfoFile InfoSpoolFileName
50630  For i = 1 To isf.InfoFiles.Count
50640   Set isfi = isf.InfoFiles(i)
50650   If FileExists(isfi.SpoolFileName) Then
50660    AddParams isfi.SpoolFileName
50670   End If
50680  Next i
50690  If LenB(PDFDocInfoFile) > 0 Then
50700   If FileExists(PDFDocInfoFile) Then
50710    AddParams PDFDocInfoFile
50720   End If
50730  End If
50740
50750  ShowParams
50760  CallGhostscript "RAW"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreateRAW")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateSVG(InfoSpoolFileName As String, ByVal GSOutputFile As String, Options As tOptions, PDFDocInfoFile As String, StampFile As String)
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
50110
50120  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50130
50140  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50150   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50160  End If
50170  AddParams "-I" & tStr
50180  AddParams "-q"
50190  AddParams "-dNOPAUSE"
50200  'AddParams "-dSAFER"
50210  AddParams "-dBATCH"
50220  AddParams "-q"
50230  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50240   AddParams "-sFONTPATH=" & GetFontsDirectory
50250  End If
50260
50270  AddParams "-sDEVICE=svg"
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
50450   AddParams "-r" & Options.SVGResolution & "x" & Options.SVGResolution
50460  End If
50470
50480  If Options.AllowSpecialGSCharsInFilenames = 1 And Options.OnePagePerFile <> 1 Then
50490   GSOutputFile = Replace$(GSOutputFile, "%", "%%")
50500  End If
50510  AddParams "-sOutputFile=" & GSOutputFile
50520
50530  AddAdditionalGhostscriptParameters
50540
50550  AddParams "-f"
50560  If LenB(StampFile) > 0 Then
50570   If FileExists(StampFile) Then
50580    AddParams StampFile
50590   End If
50600  End If
50610  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50620  Set isf = New clsInfoSpoolFile
50630  isf.ReadInfoFile InfoSpoolFileName
50640  For i = 1 To isf.InfoFiles.Count
50650   Set isfi = isf.InfoFiles(i)
50660   If FileExists(isfi.SpoolFileName) Then
50670    AddParams isfi.SpoolFileName
50680   End If
50690  Next i
50700  If LenB(PDFDocInfoFile) > 0 Then
50710   If FileExists(PDFDocInfoFile) Then
50720    AddParams PDFDocInfoFile
50730   End If
50740  End If
50750
50760  ShowParams
50770  CallGhostscript "SVG"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreateSVG")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CallGScript(GSInfoSpoolFile As String, GSOutputFile As String, _
 Options As tOptions, Ghostscriptdevice As tGhostscriptDevice, PDFDocInfoFile As String, StampFile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim enc As Boolean, encPDF As EncryptData, retEnc As Boolean, _
  Tempfile As String, tL As Long, m As Object
50030
50040  GSInit Options
50051  Select Case Ghostscriptdevice
        Case 0: 'PDF
50070    With Options
50080     If .PDFOptimize = 1 And .PDFUseSecurity = 0 Then
50090       Tempfile = GetTempFile(GetTempPathApi, "~CP")
50100       KillFile Tempfile
50110       CreatePDF GSInfoSpoolFile, Tempfile, Options, PDFDocInfoFile, StampFile
50120       OptimizePDF Tempfile, GSOutputFile
50130       KillFile Tempfile
50140      Else
50150       If .PDFUseSecurity <> 0 And SecurityIsPossible = True Then
50160         If pdfforgeDllInstalled And .PDFAes128Encryption = 1 Then
50170           enc = SetEncryptionParams(encPDF, GSInfoSpoolFile, GSOutputFile)
50180           If enc = True Then
50190            Tempfile = GetTempFile(GetTempPathApi, "~CP")
50200            KillFile Tempfile
50210            currentOwnerPassword = encPDF.OwnerPass
50220            CreatePDF GSInfoSpoolFile, Tempfile, Options, PDFDocInfoFile, StampFile
50230            encPDF.InputFile = Tempfile
50240            retEnc = EncryptPDFUsingPdfforgeDll(encPDF)
50250            If retEnc = False Then
50260             IfLoggingWriteLogfile "Error with encryption - using unencrypted file"
50270             KillFile GSOutputFile
50280             If FileExists(GSOutputFile) = False Then
50290              Name Tempfile As GSOutputFile
50300             End If
50310            End If
50320           End If
50330          Else
50340           tL = .PDFOptimize
50350           .PDFOptimize = 0
50360           CreatePDF GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
50370           .PDFOptimize = tL
50380         End If
50390        Else
50400         CreatePDF GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
50410       End If
50420     End If
50430     If pdfforgeDllInstalled Then
50440      If .PDFUpdateMetadata > 0 Then
50450       If .PDFUpdateMetadata = 2 Or _
       (.PDFUpdateMetadata = 1 And (InStr(1, .AdditionalGhostscriptParameters, "dpdfa", vbTextCompare) > 0)) Then
50470        Set m = CreateObject("pdfForge.pdf.pdf")
50480        Tempfile = GetTempFile(GetTempPathApi, "~MP")
50490        KillFile Tempfile
50500        Call m.UpdateXMPMetadata(GSOutputFile, Tempfile)
50510        If FileExists(Tempfile) Then
50520         If KillFile(GSOutputFile) Then
50530          Name Tempfile As GSOutputFile
50540         End If
50550        End If
50560       End If
50570      End If
50580     End If
50590     If DotNet20Installed And pdfforgeDllInstalled Then
50600      If .PDFSigningSignPDF = 1 Then
50610       SignPDF GSOutputFile, currentOwnerPassword
50620      End If
50630     End If
50640    End With
50650   Case 1: 'PNG
50660    CreatePNG GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
50670   Case 2: 'JPEG
50680    CreateJPEG GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
50690   Case 3: 'BMP
50700    CreateBMP GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
50710   Case 4: 'PCX
50720    CreatePCX GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
50730   Case 5: 'TIFF
50740    CreateTIFF GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
50750   Case 6: 'PS
50760    CreatePS GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
50770   Case 7: 'EPS
50780    CreateEPS GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
50790   Case 8: 'TXT
50800    ' TODO
50810    ' CreateTXT GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
50820    ' CreateTextFile GSOutputFile, GS_OutStr
50830   Case 9: 'PDFA
50840    CreatePDFA GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
50850    With Options
50860     If DotNet20Installed And pdfforgeDllInstalled Then
50870      If .PDFUpdateMetadata > 0 Then
50880       Set m = CreateObject("pdfForge.pdf.pdf")
50890       Tempfile = GetTempFile(GetTempPathApi, "~MP")
50900       KillFile Tempfile
50910       Call m.UpdateXMPMetadata(GSOutputFile, Tempfile)
50920       If FileExists(Tempfile) Then
50930        If KillFile(GSOutputFile) Then
50940         Name Tempfile As GSOutputFile
50950        End If
50960       End If
50970      End If
50980      If .PDFSigningSignPDF = 1 Then
50990       SignPDF GSOutputFile
51000      End If
51010     End If
51020    End With
51030   Case 10: 'PDFX
51040    CreatePDFX GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
51050    With Options
51060     If DotNet20Installed And pdfforgeDllInstalled Then
51070      If .PDFSigningSignPDF = 1 Then
51080       SignPDF GSOutputFile
51090      End If
51100     End If
51110    End With
51120   Case 11: 'PSD
51130    CreatePSD GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
51140   Case 12: 'PCL
51150    CreatePCL GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
51160   Case 13: 'RAW
51170    CreateRAW GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
51180   Case 14: 'SVG
51190    CreateSVG GSInfoSpoolFile, GSOutputFile, Options, PDFDocInfoFile, StampFile
51200  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CallGScript")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub SignPDF(filename As String, Optional ownerPasswd As String = "")
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim m As Object, signatureVisible As Boolean, multiSignatures As Boolean, Tempfile As String, tStr As String
50020   Dim res As Long, files As Collection, certFilename As String
50030  With Options
50040   If LenB(.PDFSigningPFXFile) = 0 Then
50050     res = OpenFileDialog(files, "", _
     LanguageStrings.OptionsPDFSigningPfxP12Files & " (*.pfx,*.p12)|*.pfx;*.p12|" & _
     LanguageStrings.OptionsPDFSigningPfxFiles & " (*.pfx)|*pfx|" & _
     LanguageStrings.OptionsPDFSigningP12Files & " (*.p12|*.p12", _
     "*.pfx;*.p12", GetMyFiles, LanguageStrings.OptionsPDFSigningChooseCertifcateFile, _
     OFN_FILEMUSTEXIST Or OFN_EXPLORER Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST, 0, 1)
50110     If res > 0 Then
50120      certFilename = files(1)
50130     End If
50140    Else
50150     certFilename = .PDFSigningPFXFile
50160   End If
50170   If LenB(.PDFSigningPFXFilePassword) > 0 Then
50180     PFXPassword = .PDFSigningPFXFilePassword
50190    Else
50200     'Ask for the password
50210     frmCertificatePassword.certFilename = certFilename
50220     frmCertificatePassword.Show vbModal
50230   End If
50240   If LenB(PFXPassword) = 0 Then
50250    MsgBox LanguageStrings.OptionsPDFSigningCertificateEmptyPassword, vbCritical + vbOKOnly
50260    Exit Sub
50270   End If
50280   Tempfile = GetTempFile(GetTempPathApi, "~MP")
50290   KillFile Tempfile
50300   If .PDFSigningSignatureVisible = 0 Then
50310     signatureVisible = False
50320    Else
50330     signatureVisible = True
50340   End If
50350   If .PDFSigningMultiSignature = 0 Then
50360     multiSignatures = False
50370    Else
50380     multiSignatures = True
50390   End If
50400   If ownerPasswd = vbNullString Then
50410    ownerPasswd = ""
50420   End If
50430   Set m = CreateObject("pdfforge.pdf.pdf")
50440   Call m.SignPDFFile(filename, ownerPasswd, Tempfile, certFilename, PFXPassword, .PDFSigningSignatureReason, .PDFSigningSignatureContact, .PDFSigningSignatureLocation, _
   signatureVisible, .PDFSigningSignatureOnPage, .PDFSigningSignatureLeftX, .PDFSigningSignatureLeftY, .PDFSigningSignatureRightX, .PDFSigningSignatureRightY, multiSignatures, Nothing)
50460  End With
50470  If FileExists(Tempfile) Then
50480   If KillFile(filename) Then
50490    Name Tempfile As filename
50500   End If
50510  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "SignPDF")
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
50030
50040  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50050
50060  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50070   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50080  End If
50090  AddParams "-I" & tStr
50100  AddParams "-q"
50110  AddParams "-dNODISPLAY"
50120  'AddParams "-dSAFER"
50130  AddParams "-dDELAYSAFER"
50140  AddParams "--"
50150  AddParams "pdfopt.ps"
50160  AddParams PDFInputFilename
50170  AddParams PDFOutputFilename
50180
50190  GSParams(0) = "pdfopt"
50200   If PerformanceTimer Then
50210    c = ExactTimer_Value() - LastStop
50220    IfLoggingWriteLogfile "Time for converting: " & _
    Format$(Int(c) * (1 / 86400), "hh:nn:ss:") & Format$(((c) - Int(c)) * 1000, "000")
50240   Else
50250    IfLoggingWriteLogfile "Time for converting -> No performance timer"
50260  End If
50270
50280  If PerformanceTimer Then
50290   LastStop = ExactTimer_Value()
50300  End If
50310  OptimizePDF = CallGS(GSParams)
50320  If PerformanceTimer Then
50330    c = ExactTimer_Value() - LastStop
50340    IfLoggingWriteLogfile "Time for optimizing: " & _
    Format$(Int(c) * (1 / 86400), "hh:nn:ss:") & Format$(((c) - Int(c)) * 1000, "000")
50360   Else
50370    IfLoggingWriteLogfile "Time for optimizing: No performance timer"
50380  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "OptimizePDF")
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
Select Case ErrPtnr.OnError("modGhostScript", "Bool2Text")
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
Select Case ErrPtnr.OnError("modGhostScript", "SelectColorCompression")
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
Select Case ErrPtnr.OnError("modGhostScript", "SelectGreyCompression")
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
Select Case ErrPtnr.OnError("modGhostScript", "SelectMonoCompression")
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
Select Case ErrPtnr.OnError("modGhostScript", "InitParams")
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
Select Case ErrPtnr.OnError("modGhostScript", "ShowParams")
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
Select Case ErrPtnr.OnError("modGhostScript", "AddParams")
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
Select Case ErrPtnr.OnError("modGhostScript", "BuildPermissionString")
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
50080  strShell = PDFCreatorApplicationPath & "pdfenc.exe """ & encData.InputFile & """ """ & encData.OutputFile & """ """ & encData.UserPass & """ """ & encData.OwnerPass & """ " & strPermissions & " " & encData.EncryptionLevel
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
Select Case ErrPtnr.OnError("modGhostScript", "EncryptPDF")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function EncryptPDFUsingPdfforgeDllTest() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim pdf As Object, enc As Object
50020  Set enc = CreateObject("pdfForge.pdf.PDFEncryptor")
50030
50040  enc.AllowAssembly = False
50050  enc.AllowCopy = False
50060  enc.AllowFillIn = False
50070  enc.AllowModifyAnnotations = False
50080  enc.AllowModifyContents = False
50090  enc.AllowPrinting = True
50100  enc.AllowPrintingHighResolution = False
50110  enc.AllowScreenReaders = True
50120  enc.EncryptionMethode = 2
50130  enc.OwnerPassword = "pdfforge"
50140  enc.UserPassword = ""
50150  Set pdf = CreateObject("pdfForge.pdf.pdf")
50160  pdf.CreatePDFTestDocument "test.pdf", 10, "This is an example.", True
50170  pdf.EncryptPDFFile "test.pdf", "encrypted.pdf", enc
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "EncryptPDFUsingPdfforgeDllTest")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function EncryptPDFUsingPdfforgeDll(encData As EncryptData) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim m As Object, enc As Object, res As Long
50020  Set m = CreateObject("pdfforge.pdf.pdf")
50030  Set enc = CreateObject("pdfforge.pdf.PDFEncryptor")
50040
50050  enc.AllowAssembly = encData.AllowAssembly
50060  enc.AllowCopy = Not encData.DisallowCopy
50070  enc.AllowFillIn = encData.AllowFillIn
50080  enc.AllowModifyAnnotations = Not encData.DisallowModifyAnnotations
50090  enc.AllowModifyContents = Not encData.DisallowModifyContents
50100  enc.AllowPrinting = Not encData.DisallowPrinting
50110  enc.AllowPrintingHighResolution = Not encData.AllowDegradedPrinting
50120  enc.AllowScreenReaders = encData.AllowFillIn
50130  enc.EncryptionMethod = 2
50140  enc.OwnerPassword = encData.OwnerPass
50150  enc.UserPassword = encData.UserPass
50160
50170  res = m.EncryptPDFFile(encData.InputFile, encData.OutputFile, enc)
50180  If res = 0 Then
50190   EncryptPDFUsingPdfforgeDll = True
50200  Else
50210   EncryptPDFUsingPdfforgeDll = False
50220  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "EncryptPDFUsingPdfforgeDll")
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
Select Case ErrPtnr.OnError("modGhostScript", "CalculatePermissions")
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
Select Case ErrPtnr.OnError("modGhostScript", "SetEncryptionParams")
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
50480        AddParams "-dColorImageFilter=/DCTEncode"
50490        AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionColorCompressionJPEGManualFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50520       Case 7:
50530        AddParams "-dColorImageFilter=/FlateEncode"
50540       Case 8:
50550        AddParams "-dColorImageFilter=/LZWEncode"
50560      End Select
50570    End If
50580   Else
50590    AddParams "-dEncodeColorImages=false"
50600  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "SetColorParams")
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
50480        AddParams "-dGrayImageFilter=/DCTEncode"
50490        AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor " & _
        Replace$(CStr(Options.PDFCompressionGreyCompressionJPEGManualFactor), GetDecimalChar, ".") & _
        " /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50520       Case 7:
50530        AddParams "-dGrayImageFilter=/FlateEncode"
50540       Case 8:
50550        AddParams "-dGrayImageFilter=/LZWEncode"
50560      End Select
50570    End If
50580   Else
50590    AddParams "-dEncodeGrayImages=false"
50600  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "SetGreyParams")
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
Select Case ErrPtnr.OnError("modGhostScript", "SetMonoParams")
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
Select Case ErrPtnr.OnError("modGhostScript", "GhostScriptSecurity")
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
50200    .SubKey = tColl.Item(i)
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
50330    .SubKey = tColl.Item(i)
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
50460    .SubKey = tColl.Item(i)
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
Select Case ErrPtnr.OnError("modGhostScript", "GetAllGhostscriptversions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CreateStampFile(filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim StampPage As String, tStr As String, R As String, G As String, b As String, _
  StampFile As String, Path As String, ff As Long, files As Collection, _
  StampString As String, StampFontsize As Double, _
  StampOutlineFontthickness As Double
50050  Dim File As String
50060  StampString = RemoveLeadingAndTrailingQuotes(Trim$(Options.StampString))
50070  If Len(StampString) > 0 Then
50080   StampPage = StrConv(LoadResData(101, "STAMPPAGE"), vbUnicode)
50090   StampPage = Replace(StampPage, vbCrLf, vbCr, , , vbBinaryCompare)
50100   StampPage = Replace(StampPage, "[STAMPSTRING]", EncodeCharsOctal(StampString), , , vbTextCompare)
50110   StampPage = Replace(StampPage, "[FONTNAME]", Replace(Trim$(Options.StampFontname), " ", ""), , , vbTextCompare)
50120   StampFontsize = 48
50130   If IsNumeric(Options.StampFontsize) = True Then
50140    If CDbl(Options.StampFontsize) > 0 Then
50150     StampFontsize = CDbl(Options.StampFontsize)
50160    End If
50170   End If
50180   StampPage = Replace(StampPage, "[FONTSIZE]", StampFontsize, , , vbTextCompare)
50190   StampOutlineFontthickness = 0
50200   If IsNumeric(Options.StampOutlineFontthickness) = True Then
50210    If CDbl(Options.StampOutlineFontthickness) >= 0 Then
50220     StampOutlineFontthickness = CDbl(Options.StampOutlineFontthickness)
50230    End If
50240   End If
50250   StampPage = Replace(StampPage, "[STAMPOUTLINEFONTTHICKNESS]", StampOutlineFontthickness, , , vbTextCompare)
50260   If Options.StampUseOutlineFont <> 1 Then
50270     StampPage = Replace(StampPage, "[USEOUTLINEFONT]", "show", , , vbTextCompare)
50280    Else
50290     StampPage = Replace(StampPage, "[USEOUTLINEFONT]", "true charpath stroke", , , vbTextCompare)
50300   End If
50310   If Len(Options.StampFontColor) > 0 Then
50320     tStr = Replace$(Options.StampFontColor, "#", "&H")
50330     If IsNumeric(tStr) = True Then
50340       R = Replace$(Format(CDbl((CLng(tStr) And CLng("&HFF0000")) / 65536) / 255#, "0.00"), ",", ".", , 1)
50350       G = Replace$(Format(CDbl((CLng(tStr) And CLng("&H00FF00")) / 256) / 255#, "0.00"), ",", ".", , 1)
50360       b = Replace$(Format(CDbl(CLng(tStr) And CLng("&H0000FF")) / 255#, "0.00"), ",", ".", , 1)
50370       StampPage = Replace(StampPage, "[FONTCOLOR]", R & " " & G & " " & b, , , vbTextCompare)
50380      Else
50390       StampPage = Replace(StampPage, "[FONTCOLOR]", "1 0 0", , , vbTextCompare)
50400     End If
50410    Else
50420     StampPage = Replace(StampPage, "[FONTCOLOR]", "1 0 0", , , vbTextCompare)
50430   End If
50440   SplitPath filename, , Path, , File
50450   StampFile = CompletePath(Path) & File & ".stm"
50460   ff = FreeFile
50470   Open StampFile For Output As #ff
50480   Print #ff, StampPage
50490   Close #ff
50500   CreateStampFile = StampFile
50510  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CreateStampFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


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
Select Case ErrPtnr.OnError("modGhostScript", "AddParamCommands")
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
Select Case ErrPtnr.OnError("modGhostScript", "AddParamCommand")
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
50020  tStr = Replace$(Trim$(Options.AdditionalGhostscriptParameters), "<app>", PDFCreatorApplicationPath, , , vbTextCompare)
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
Select Case ErrPtnr.OnError("modGhostScript", "AddAdditionalGhostscriptParameters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub CheckForPrintingAfterSaving(InfoSpoolFileName As String, Options As tOptions)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String, PrintAfterSavingBitsPerPixel(2) As String, NoCancelStr As String, sPrinter1 As String, _
  sQueryUser As String, sDuplex As String, sMaxResolution As String
50030
50040  If Options.PrintAfterSaving = 0 Then
50050   Exit Sub
50060  End If
50070
50080  GSInit Options
50090  InitParams
50100  Set ParamCommands = New Collection
50110
50120  PrintAfterSavingBitsPerPixel(0) = "1": PrintAfterSavingBitsPerPixel(1) = "4"
50130  PrintAfterSavingBitsPerPixel(2) = "24"
50140
50150  tStr = Options.DirectoryGhostscriptLibraries & GetGhostscriptResourceString
50160
50170  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50180   tStr = tStr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50190  End If
50200  AddParams "-I" & tStr
50210  AddParams "-q"
50220  AddParams "-dNOPAUSE"
50230  AddParams "-dBATCH"
50240  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50250   AddParams "-sFONTPATH=" & GetFontsDirectory
50260  End If
50270
50280  NoCancelStr = "/NoCancel "
50290  If Options.PrintAfterSavingNoCancel = 1 Then
50300    NoCancelStr = NoCancelStr & "true"
50310   Else
50320    NoCancelStr = NoCancelStr & "false"
50330  End If
50340
50350  If Options.PrintAfterSavingQueryUser > 0 Then
50360    sQueryUser = " /QueryUser " & Options.PrintAfterSavingQueryUser
50370   Else
50380    If LenB(Trim$(Options.PrintAfterSavingPrinter)) > 0 Then
50390        If Mid$(Options.PrintAfterSavingPrinter, 1, 2) = "\\" Then ' network printer
50400        sPrinter1 = " /OutputFile (" & Replace$(Options.PrintAfterSavingPrinter, "\", "\\") & ") "
50410       Else
50420        sPrinter1 = " /OutputFile (" & Replace$("\\spool\" & Options.PrintAfterSavingPrinter, "\", "\\") & ") "
50430      End If
50440     Else
50450      sQueryUser = " /QueryUser 1"
50460    End If
50470  End If
50480
50490  sMaxResolution = " dup /MaxResolution " & Options.PrintAfterSavingMaxResolution & " put"
50500  sMaxResolution = ""
50510
50520  AddParamCommand "mark " & NoCancelStr & " /BitsPerPixel " & PrintAfterSavingBitsPerPixel(Options.PrintAfterSavingBitsPerPixel) & _
   sQueryUser & sPrinter1 & _
  " /UserSettings 1 dict dup /DocumentName (" & Replace$("A", "\", "\\") & ") put" & sMaxResolution & " (mswinpr2) finddevice putdeviceprops setdevice"
50550
50560  If Options.PrintAfterSavingDuplex = 1 Then
50570   If Options.PrintAfterSavingTumble = 1 Then
50580     AddParamCommand "<< /Duplex true /Tumble true >> setpagedevice"
50590    Else
50600     AddParamCommand "<< /Duplex true /Tumble false >> setpagedevice"
50610   End If
50620  End If
50630
50640  AddParamCommands
50650  AddParams "-f"
50660
50670  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long
50680  Set isf = New clsInfoSpoolFile
50690  isf.ReadInfoFile InfoSpoolFileName
50700  For i = 1 To isf.InfoFiles.Count
50710   Set isfi = isf.InfoFiles(i)
50720   If FileExists(isfi.SpoolFileName) Then
50730    AddParams isfi.SpoolFileName
50740   End If
50750  Next i
50760
50770  ShowParams
50780  CallGhostscript "mswinpr2"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "CheckForPrintingAfterSaving")
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
50010  Dim Ext As String, Tempfile As String, ivgf As Boolean, inFile As String, Path As String, File As String, psFileName As String
50020  Dim InfoSpoolFileName As String, PDFDocInfoFile As String, StampFile As String, PDFDocInfo As tPDFDocInfo, tDate As Date
50030  Dim PSHeader As tPSHeader, tStr As String, strGUID As String
50040
50050  IFIsPS = False
50060  If LenB(InputFilename) = 0 Then
50070   Exit Sub
50080  End If
50090  If FileExists(InputFilename) = False Then
50100   If LenB(InputFilename) > 0 Then
50110    MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & _
    "InputFile -IF" & vbCrLf & ">" & InputFilename & "<", vbExclamation + vbOKOnly
50130   End If
50140   Exit Sub
50150  End If
50160  ivgf = IsValidGraphicFile(InputFilename)
50170  If LenB(OutputFilename) > 0 Then
50180    If IsPostscriptFile(InputFilename) = True Or ivgf Or IsPDFFile(InputFilename) Then
50190     If GsDllLoaded = 0 Then
50200      Exit Sub
50210     End If
50220     GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50230     If GsDllLoaded = 0 Then
50240      MsgBox LanguageStrings.MessagesMsg08
50250     End If
50260     inFile = InputFilename
50270     strGUID = GetGUID
50280     File = GetPDFCreatorSpoolDirectory & strGUID
50290     If ivgf Then
50300      psFileName = File & ".ps"
50310      If Image2PS(InputFilename, psFileName) Then
50320        inFile = psFileName
50330       Else
50340        IfLoggingWriteLogfile "ConvertFile: There is a problem converting '" & InputFilename & "'!"
50350        Exit Sub
50360      End If
50370     End If
50380
50390     InfoSpoolFileName = CreateInfoSpoolFile(inFile, File & ".inf")
50400     SplitPath OutputFilename, , , , , Ext
50410
50420     With PDFDocInfo
50430      If Len(Trim$(Options.StandardTitle)) > 0 Then
50440        .Author = GetSubstFilename(inFile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardTitle)), , , True)
50450       Else
50460        .Author = GetSubstFilename(inFile, Options.SaveFilename, , , True)
50470      End If
50480      If Options.UseStandardAuthor = 1 Then
50490        .Creator = GetSubstFilename(inFile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardAuthor)), True, , True)
50500       Else
50510        .Creator = GetDocUsernameFromPostScriptFile(inFile, False)
50520      End If
50530      If Len(Trim$(Options.StandardKeywords)) > 0 Then
50540       .Keywords = GetSubstFilename(inFile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardKeywords)), , , True)
50550      End If
50560      If Len(Trim$(Options.StandardSubject)) > 0 Then
50570       .Subject = GetSubstFilename(inFile, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardSubject)), , , True)
50580      End If
50590      tDate = Now
50600      PSHeader = GetPSHeader(inFile)
50610
50620      If LenB(PSHeader.CreationDate.Comment) > 0 Then
50630        tStr = FormatPrintDocumentDate(PSHeader.CreationDate.Comment)
50640       Else
50650        tStr = CStr(tDate)
50660      End If
50670      .CreationDate = GetDocDate(Options.StandardCreationdate, Options.StandardDateformat, tStr)
50680      .ModifyDate = GetDocDate(Options.StandardModifydate, Options.StandardDateformat, tStr)
50690      .Creator = App.EXEName & " Version " & App.Major & "." & App.Minor & "." & App.Revision
50700     End With
50710
50720     PDFDocInfoFile = CreatePDFDocInfoFile(inFile, PDFDocInfo)
50730     StampFile = CreateStampFile(inFile)
50740
50751     Select Case UCase$(Ext)
           Case "PDF"
50771       Select Case UCase(SubFormat)
             Case "PDF/A-1B"
50790         CallGScript InfoSpoolFileName, OutputFilename, Options, PDFAWriter, PDFDocInfoFile, StampFile
50800        Case "PDF/X"
50810         CallGScript InfoSpoolFileName, OutputFilename, Options, PDFXWriter, PDFDocInfoFile, StampFile
50820        Case Else
50830         CallGScript InfoSpoolFileName, OutputFilename, Options, PDFWriter, PDFDocInfoFile, StampFile
50840       End Select
50850      Case "PNG"
50860       CallGScript InfoSpoolFileName, OutputFilename, Options, PNGWriter, PDFDocInfoFile, StampFile
50870      Case "JPG"
50880       CallGScript InfoSpoolFileName, OutputFilename, Options, JPEGWriter, PDFDocInfoFile, StampFile
50890      Case "BMP"
50900       CallGScript InfoSpoolFileName, OutputFilename, Options, BMPWriter, PDFDocInfoFile, StampFile
50910      Case "PCX"
50920       CallGScript InfoSpoolFileName, OutputFilename, Options, PCXWriter, PDFDocInfoFile, StampFile
50930      Case "TIF"
50940       CallGScript InfoSpoolFileName, OutputFilename, Options, TIFFWriter, PDFDocInfoFile, StampFile
50950      Case "PS"
50960       CallGScript InfoSpoolFileName, OutputFilename, Options, PSWriter, PDFDocInfoFile, StampFile
50970      Case "EPS"
50980       CallGScript InfoSpoolFileName, OutputFilename, Options, EPSWriter, PDFDocInfoFile, StampFile
50990      Case "TXT"
51000       CallGScript InfoSpoolFileName, OutputFilename, Options, TXTWriter, PDFDocInfoFile, StampFile
51010      Case "PCL"
51020       CallGScript InfoSpoolFileName, OutputFilename, Options, PCLWriter, PDFDocInfoFile, StampFile
51030      Case "PSD"
51040       CallGScript InfoSpoolFileName, OutputFilename, Options, PSDWriter, PDFDocInfoFile, StampFile
51050      Case "RAW"
51060       CallGScript InfoSpoolFileName, OutputFilename, Options, RAWWriter, PDFDocInfoFile, StampFile
51070      Case "SVG"
51080       CallGScript InfoSpoolFileName, OutputFilename, Options, SVGWriter, PDFDocInfoFile, StampFile
51090     End Select
51100
51110     KillFile InfoSpoolFileName
51120     KillFile PDFDocInfoFile
51130     KillFile StampFile
51140    End If
51150 '   If GsDllLoaded <> 0 Then
51160 '    UnloadDLLComplete GsDllLoaded
51170 '   End If
51180    ConvertedOutputFilename = OutputFilename
51190    ReadyConverting = True
51200    Exit Sub
51210   Else
51220    If FileExists(InputFilename) = True Then
51230     If IsPostscriptFile(InputFilename) = True Then
51240       IFIsPS = True
51250      Else
51260       MsgBox LanguageStrings.MessagesMsg06 & vbCrLf & vbCrLf & InputFilename
51270     End If
51280    End If
51290  End If
51300  DoEvents
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "ConvertFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetGhostscriptVersion() As tGhostscriptVersion
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim gsRev As String, tStr As String, Major As Long, Minor As Long
50020  gsRev = CStr(GSRevision.intRevision)
50030  If Len(gsRev) >= 3 Then
50040   tStr = Mid(gsRev, Len(gsRev) - 1, 2)
50050   If IsNumeric(tStr) Then
50060    Minor = CLng(tStr)
50070   End If
50080   tStr = Mid(gsRev, 1, Len(gsRev) - 2)
50090   If IsNumeric(tStr) Then
50100    Major = CLng(tStr)
50110   End If
50120   GetGhostscriptVersion.Major = Major
50130   GetGhostscriptVersion.Minor = Minor
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "GetGhostscriptVersion")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetGhostscriptResourceString() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String
50020  If (GetGhostscriptVersion.Major < 8) Or (GetGhostscriptVersion.Major = 8 And GetGhostscriptVersion.Minor <= 62) Then
50030   If LenB(LTrim(Options.DirectoryGhostscriptFonts)) > 0 Then
50040    tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptFonts)
50050   End If
50060   If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50070    tStr = tStr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50080   End If
50090  End If
50100  GetGhostscriptResourceString = tStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostScript", "GetGhostscriptResourceString")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
