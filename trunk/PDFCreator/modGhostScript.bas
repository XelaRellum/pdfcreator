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

Private ParamCommands As Collection

Public Sub GSInit(Options As tOptions)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Rotate(2) As String, Resample(2) As String, Colormodel(2) As String, _
  ColorsPreserveTransfer(1) As String
50060  Dim PNGColorscount(4) As String, JPEGColorscount(1) As String, BMPColorscount(6) As String, _
  PCXColorscount(5) As String, TIFFColorscount(7) As String, _
  PSLanguageLevel(3) As String
50090
50100  Rotate(0) = "None": Rotate(1) = "All": Rotate(2) = "PageByPage"
50110
50120  Resample(0) = "Bicubic": Resample(1) = "Subsample": Resample(2) = "Average"
50130
50140  Colormodel(0) = "RGB": Colormodel(1) = "CMYK": Colormodel(2) = "GRAY"
50150
50160  ColorsPreserveTransfer(0) = "Remove": ColorsPreserveTransfer(1) = "Preserve"
50170
50180  PNGColorscount(0) = "png16m": PNGColorscount(1) = "png256"
50190  PNGColorscount(2) = "png16": PNGColorscount(3) = "png2"
50200  PNGColorscount(4) = "pnggray"
50210
50220  JPEGColorscount(0) = "jpeg": JPEGColorscount(1) = "jpeggray"
50230
50240  BMPColorscount(0) = "bmp32b": BMPColorscount(1) = "bmp16m"
50250  BMPColorscount(2) = "bmp256": BMPColorscount(3) = "bmp16"
50260  BMPColorscount(4) = "bmpsep8": BMPColorscount(5) = "bmpsep1"
50270  BMPColorscount(6) = "bmpgray"
50280
50290  PCXColorscount(0) = "pcxcmyk": PCXColorscount(1) = "pcx24b"
50300  PCXColorscount(2) = "pcx256": PCXColorscount(3) = "pcx16"
50310  PCXColorscount(4) = "pcxmono": PCXColorscount(5) = "pcxgray"
50320
50330  TIFFColorscount(0) = "tiff24nc": TIFFColorscount(1) = "tiff12nc"
50340  TIFFColorscount(2) = "tiffcrle": TIFFColorscount(3) = "tiffg3"
50350  TIFFColorscount(4) = "tiffg32d": TIFFColorscount(5) = "tiffg4"
50360  TIFFColorscount(6) = "tifflzw": TIFFColorscount(7) = "tiffpack"
50370
50380  PSLanguageLevel(0) = "1": PSLanguageLevel(1) = "1.5"
50390  PSLanguageLevel(2) = "2": PSLanguageLevel(3) = "3"
50400
50410 With Options
50420  'General
50430
50440  GS_COMPATIBILITY = "1." & (.PDFGeneralCompatibility + 2)
50450  GS_RESOLUTION = .PDFGeneralResolution
50460  GS_AUTOROTATE = Rotate(.PDFGeneralAutorotate)
50470  GS_OVERPRINT = .PDFGeneralOverprint
50480  GS_ASCII85 = Bool2Text(.PDFGeneralASCII85)
50490
50500  'Compression
50510  GS_COMPRESSPAGES = Bool2Text(.PDFCompressionTextCompression)
50520  GS_COMPRESSCOLOR = Bool2Text(.PDFCompressionColorCompression)
50530  GS_COMPRESSGREY = Bool2Text(.PDFCompressionGreyCompression)
50540  GS_COMPRESSMONO = Bool2Text(.PDFCompressionMonoCompression)
50550
50560  SelectColorCompression .PDFCompressionColorCompressionChoice
50570  SelectGreyCompression .PDFCompressionGreyCompressionChoice
50580  SelectMonoCompression .PDFCompressionMonoCompressionChoice
50590
50600  GS_COMPRESSCOLORVALUE = Bool2Text(.PDFCompressionColorCompression)
50610  GS_COMPRESSGREYVALUE = Bool2Text(.PDFCompressionGreyCompression)
50620  GS_COMPRESSMONOVALUE = Bool2Text(.PDFCompressionMonoCompression)
50630
50640  GS_COLORRESOLUTION = .PDFCompressionColorResolution
50650  GS_GREYRESOLUTION = .PDFCompressionGreyResolution
50660  GS_MONORESOLUTION = .PDFCompressionMonoResolution
50670
50680  GS_COLORRESAMPLE = Bool2Text(.PDFCompressionColorResample)
50690  GS_GREYRESAMPLE = Bool2Text(.PDFCompressionGreyResample)
50700  GS_MONORESAMPLE = Bool2Text(.PDFCompressionMonoResample)
50710
50720  GS_COLORRESAMPLEMETHOD = Resample(.PDFCompressionColorResampleChoice)
50730  GS_GREYRESAMPLEMETHOD = Resample(.PDFCompressionGreyResampleChoice)
50740  GS_MONORESAMPLEMETHOD = Resample(.PDFCompressionMonoResampleChoice)
50750
50760  'Fonts
50770  GS_EMBEDALLFONTS = Bool2Text(.PDFFontsEmbedAll)
50780  GS_SUBSETFONTS = Bool2Text(.PDFFontsSubSetFonts)
50790  GS_SUBSETFONTPERC = .PDFFontsSubSetFontsPercent
50800
50810  'Colors
50820  GS_COLORMODEL = Colormodel(.PDFColorsColorModel)
50830  GS_CMYKTORGB = Bool2Text(.PDFColorsCMYKToRGB)
50840  GS_PRESERVEOVERPRINT = Bool2Text(.PDFColorsPreserveOverprint)
50850  GS_TRANSFERFUNCTIONS = ColorsPreserveTransfer(.PDFColorsPreserveTransfer)
50860  GS_HALFTONE = Bool2Text(.PDFColorsPreserveHalftone)
50870
50880  'Bitmap
50890  GS_BitmapRESOLUTION = .BitmapResolution
50900  GS_PNGColorscount = PNGColorscount(.PNGColorscount)
50910  GS_JPEGColorscount = JPEGColorscount(.JPEGColorscount)
50920  GS_BMPColorscount = BMPColorscount(.BMPColorscount)
50930  GS_PCXColorscount = PCXColorscount(.PCXColorscount)
50940  GS_TIFFColorscount = TIFFColorscount(.TIFFColorscount)
50950  GS_JPEGQuality = .JPEGQuality
50960  GS_PSLanguageLevel = PSLanguageLevel(.PSLanguageLevel)
50970  GS_EPSLanguageLevel = PSLanguageLevel(.EPSLanguageLevel)
50980 End With
50990 'Other
51000 GS_ERROR = 0
51010 UseReturnPipe = 1
51020 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
51030 Exit Sub
ErrPtnr_OnError:
51051 Select Case ErrPtnr.OnError("modGhostscript", "GSInit")
      Case 0: Resume
51070 Case 1: Resume Next
51080 Case 2: Exit Sub
51090 Case 3: End
51100 End Select
51110 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function CallGhostscript(Comment As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim LastStop As Currency, res As Boolean
50050  LastStop = ExactTimer_Value()
50060  res = CallGS(GSParams)
50070  IfLoggingWriteLogfile "Time for converting in seconds [" & Comment & "]: " & CStr(ExactTimer_Value() - LastStop)
50080 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50090 Exit Function
ErrPtnr_OnError:
50111 Select Case ErrPtnr.OnError("modGhostscript", "CallGhostscript")
      Case 0: Resume
50130 Case 1: Resume Next
50140 Case 2: Exit Function
50150 Case 3: End
50160 End Select
50170 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePDF(GSInputFile As String, GSOutputFile As String, Options As tOptions)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim FName As String, tstr As String, encPDF As EncryptData, tEnc As Boolean
50050
50060  InitParams
50070  Set ParamCommands = New Collection
50080
50090  tstr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50100  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50110   tstr = tstr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50120  End If
50130  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50140   tstr = tstr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50150  End If
50160  AddParams "-I" & tstr
50170  AddParams "-q"
50180  AddParams "-dNOPAUSE"
50190  AddParams "-dSAFER"
50200  AddParams "-dBATCH"
50210  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50220   AddParams "-sFONTPATH=" & GetFontsDirectory
50230  End If
50240  AddParams "-sDEVICE=pdfwrite"
50250  If Options.DontUseDocumentSettings = 0 Then
50260   AddParams "-dCompatibilityLevel=" & GS_COMPATIBILITY
50270   AddParams "-r" & GS_RESOLUTION & "x" & GS_RESOLUTION
50280   AddParams "-dProcessColorModel=/Device" & GS_COLORMODEL
50290   AddParams "-dAutoRotatePages=/" & GS_AUTOROTATE
50300   AddParams "-dCompressPages=" & GS_COMPRESSPAGES
50310   AddParams "-dEmbedAllFonts=" & GS_EMBEDALLFONTS
50320   AddParams "-dSubsetFonts=" & GS_SUBSETFONTS
50330   AddParams "-dMaxSubsetPct=" & GS_SUBSETFONTPERC
50340   AddParams "-dConvertCMYKImagesToRGB=" & GS_CMYKTORGB
50350  End If
50360  tEnc = False
50370  If Options.PDFOptimize = 0 And Options.PDFUseSecurity <> 0 And _
    SecurityIsPossible = True And Options.PDFEncryptor = 0 Then
50390    If SetEncryptionParams(encPDF, "", "") = True Then
50400     tEnc = True
50410    If Len(encPDF.OwnerPass) > 0 Then
50420      AddParams "-sOwnerPassword=" & encPDF.OwnerPass & ""
50430     End If
50440     If Len(encPDF.UserPass) > 0 Then
50450      AddParams "-sUserPassword=" & encPDF.UserPass
50460     End If
50470     AddParams "-dPermissions=" & CalculatePermissions(encPDF)
50480     If GS_COMPATIBILITY = "1.4" Then
50490       AddParams "-dEncryptionR=3"
50500      Else
50510       AddParams "-dEncryptionR=2"
50520     End If
50530     If encPDF.EncryptionLevel = encLow Then
50540       AddParams "-dKeyLength=40"
50550      Else
50560       AddParams "-dKeyLength=128"
50570     End If
50580    Else
50590     If Options.UseAutosave = 0 Then
50600      MsgBox LanguageStrings.MessagesMsg23, vbCritical
50610     End If
50620   End If
50630  End If
50640  AddParams "-sOutputFile=" & GSOutputFile
50650
50660  If Options.DontUseDocumentSettings = 0 Then
50670   SetColorParams
50680   SetGreyParams
50690   SetMonoParams
50700
50710   AddParams "-dPreserveOverprintSettings=" & GS_PRESERVEOVERPRINT
50720   AddParams "-dUCRandBGInfo=/Preserve"
50730   AddParams "-dUseFlateCompression=true"
50740   AddParams "-dParseDSCCommentsForDocInfo=true"
50750   AddParams "-dParseDSCComments=true"
50760   AddParams "-dOPM=" & GS_OVERPRINT
50770   AddParams "-dOffOptimizations=0"
50780   AddParams "-dLockDistillerParams=false"
50790   AddParams "-dGrayImageDepth=-1"
50800   AddParams "-dASCII85EncodePages=" & GS_ASCII85
50810   AddParams "-dDefaultRenderingIntent=/Default"
50820   AddParams "-dTransferFunctionInfo=/" & GS_TRANSFERFUNCTIONS
50830   AddParams "-dPreserveHalftoneInfo=" & GS_HALFTONE
50840   AddParams "-dDetectBlends=true"
50850   AddParamCommands
50860  End If
50870
50880  AddParams "-f"
50890  AddParams GSInputFile
50900  ShowParams
50910  If tEnc = True Then
50920    CallGhostscript "PDF with encryption"
50930   Else
50940    CallGhostscript "PDF without encryption"
50950  End If
50960 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50970 Exit Function
ErrPtnr_OnError:
50991 Select Case ErrPtnr.OnError("modGhostscript", "CreatePDF")
      Case 0: Resume
51010 Case 1: Resume Next
51020 Case 2: Exit Function
51030 Case 3: End
51040 End Select
51050 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePNG(GSInputFile As String, GSOutputFile As String, Options As tOptions)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Path As String, FName As String, Ext As String, tstr As String
50050
50060  GSInit Options
50070  InitParams
50080  Set ParamCommands = New Collection
50090
50100  If Options.OnePagePerFile = 1 Then
50110   SplitPath GSOutputFile, , Path, , FName, Ext
50120   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50130  End If
50140  tstr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50150  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50160   tstr = tstr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50170  End If
50180  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50190   tstr = tstr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50200  End If
50210  AddParams "-I" & tstr
50220  AddParams "-q"
50230  AddParams "-dNOPAUSE"
50240  AddParams "-dSAFER"
50250  AddParams "-dBATCH"
50260  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50270   AddParams "-sFONTPATH=" & GetFontsDirectory
50280  End If
50290
50300  If Options.OnePagePerFile = 1 Then
50310   SplitPath GSOutputFile, , Path, , FName, Ext
50320   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50330  End If
50340
50350  AddParams "-sDEVICE=" & GS_PNGColorscount
50360
50370  If Options.DontUseDocumentSettings = 0 Then
50380   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
50390   AddParams "-sOutputFile=" & GSOutputFile
50400  End If
50410
50420  AddParams "-f"
50430  AddParams GSInputFile
50440  ShowParams
50450  CallGhostscript "PNG"
50460 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50470 Exit Function
ErrPtnr_OnError:
50491 Select Case ErrPtnr.OnError("modGhostscript", "CreatePNG")
      Case 0: Resume
50510 Case 1: Resume Next
50520 Case 2: Exit Function
50530 Case 3: End
50540 End Select
50550 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateJPEG(GSInputFile As String, GSOutputFile As String, Options As tOptions)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Path As String, FName As String, Ext As String, tstr As String
50050
50060  GSInit Options
50070  InitParams
50080  Set ParamCommands = New Collection
50090
50100  If Options.OnePagePerFile = 1 Then
50110   SplitPath GSOutputFile, , Path, , FName, Ext
50120   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50130  End If
50140  tstr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50150  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50160   tstr = tstr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50170  End If
50180  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50190   tstr = tstr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50200  End If
50210  AddParams "-I" & tstr
50220  AddParams "-q"
50230  AddParams "-dNOPAUSE"
50240  AddParams "-dSAFER"
50250  AddParams "-dBATCH"
50260  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50270   AddParams "-sFONTPATH=" & GetFontsDirectory
50280  End If
50290
50300  If Options.OnePagePerFile = 1 Then
50310   SplitPath GSOutputFile, , Path, , FName, Ext
50320   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50330  End If
50340
50350  AddParams "-sDEVICE=" & GS_JPEGColorscount
50360  If Options.DontUseDocumentSettings = 0 Then
50370   AddParams "-dJPEGQ=" & GS_JPEGQuality
50380   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
50390   AddParams "-sOutputFile=" & GSOutputFile
50400  End If
50410
50420  AddParams "-f"
50430  AddParams GSInputFile
50440  ShowParams
50450  CallGhostscript "JPEG"
50460 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50470 Exit Function
ErrPtnr_OnError:
50491 Select Case ErrPtnr.OnError("modGhostscript", "CreateJPEG")
      Case 0: Resume
50510 Case 1: Resume Next
50520 Case 2: Exit Function
50530 Case 3: End
50540 End Select
50550 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateBMP(GSInputFile As String, GSOutputFile As String, Options As tOptions)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Path As String, FName As String, Ext As String, tstr As String
50050
50060  GSInit Options
50070  InitParams
50080  Set ParamCommands = New Collection
50090
50100  If Options.OnePagePerFile = 1 Then
50110   SplitPath GSOutputFile, , Path, , FName, Ext
50120   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50130  End If
50140  tstr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50150  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50160   tstr = tstr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50170  End If
50180  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50190   tstr = tstr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50200  End If
50210  AddParams "-I" & tstr
50220  AddParams "-q"
50230  AddParams "-dNOPAUSE"
50240  AddParams "-dSAFER"
50250  AddParams "-dBATCH"
50260  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50270   AddParams "-sFONTPATH=" & GetFontsDirectory
50280  End If
50290
50300  If Options.OnePagePerFile = 1 Then
50310   SplitPath GSOutputFile, , Path, , FName, Ext
50320   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50330  End If
50340
50350  AddParams "-sDEVICE=" & GS_BMPColorscount
50360  If Options.DontUseDocumentSettings = 0 Then
50370   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
50380  End If
50390  AddParams "-sOutputFile=" & GSOutputFile
50400
50410  AddParams "-f"
50420  AddParams GSInputFile
50430  ShowParams
50440  CallGhostscript "BMP"
50450 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50460 Exit Function
ErrPtnr_OnError:
50481 Select Case ErrPtnr.OnError("modGhostscript", "CreateBMP")
      Case 0: Resume
50500 Case 1: Resume Next
50510 Case 2: Exit Function
50520 Case 3: End
50530 End Select
50540 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePCX(GSInputFile As String, GSOutputFile As String, Options As tOptions)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Path As String, FName As String, Ext As String, tstr As String
50050
50060  GSInit Options
50070  InitParams
50080  Set ParamCommands = New Collection
50090
50100  If Options.OnePagePerFile = 1 Then
50110   SplitPath GSOutputFile, , Path, , FName, Ext
50120   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50130  End If
50140  tstr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50150  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50160   tstr = tstr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50170  End If
50180  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50190   tstr = tstr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50200  End If
50210  AddParams "-I" & tstr
50220  AddParams "-q"
50230  AddParams "-dNOPAUSE"
50240  AddParams "-dSAFER"
50250  AddParams "-dBATCH"
50260  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50270   AddParams "-sFONTPATH=" & GetFontsDirectory
50280  End If
50290
50300  If Options.OnePagePerFile = 1 Then
50310   SplitPath GSOutputFile, , Path, , FName, Ext
50320   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50330  End If
50340
50350  AddParams "-sDEVICE=" & GS_PCXColorscount
50360  If Options.DontUseDocumentSettings = 0 Then
50370   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
50380  End If
50390  AddParams "-sOutputFile=" & GSOutputFile
50400
50410  AddParams "-f"
50420  AddParams GSInputFile
50430  ShowParams
50440  CallGhostscript "PCX"
50450 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50460 Exit Function
ErrPtnr_OnError:
50481 Select Case ErrPtnr.OnError("modGhostscript", "CreatePCX")
      Case 0: Resume
50500 Case 1: Resume Next
50510 Case 2: Exit Function
50520 Case 3: End
50530 End Select
50540 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateTIFF(GSInputFile As String, GSOutputFile As String, Options As tOptions)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Path As String, FName As String, Ext As String, tstr As String
50050
50060  GSInit Options
50070  InitParams
50080  Set ParamCommands = New Collection
50090
50100  If Options.OnePagePerFile = 1 Then
50110   SplitPath GSOutputFile, , Path, , FName, Ext
50120   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50130  End If
50140  tstr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50150  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50160   tstr = tstr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50170  End If
50180  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50190   tstr = tstr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50200  End If
50210  AddParams "-I" & tstr
50220  AddParams "-q"
50230  AddParams "-dNOPAUSE"
50240  AddParams "-dSAFER"
50250  AddParams "-dBATCH"
50260  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50270   AddParams "-sFONTPATH=" & GetFontsDirectory
50280  End If
50290
50300  If Options.OnePagePerFile = 1 Then
50310   SplitPath GSOutputFile, , Path, , FName, Ext
50320   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50330  End If
50340
50350  AddParams "-sDEVICE=" & GS_TIFFColorscount
50360  If Options.DontUseDocumentSettings = 0 Then
50370   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
50380  End If
50390  AddParams "-sOutputFile=" & GSOutputFile
50400
50410  AddParams "-f"
50420  AddParams GSInputFile
50430  ShowParams
50440  CallGhostscript "TIFF"
50450 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50460 Exit Function
ErrPtnr_OnError:
50481 Select Case ErrPtnr.OnError("modGhostscript", "CreateTIFF")
      Case 0: Resume
50500 Case 1: Resume Next
50510 Case 2: Exit Function
50520 Case 3: End
50530 End Select
50540 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreatePS(GSInputFile As String, GSOutputFile As String, Options As tOptions)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Path As String, FName As String, Ext As String, tstr As String
50050
50060  GSInit Options
50070  InitParams
50080  Set ParamCommands = New Collection
50090
50100  If Options.OnePagePerFile = 1 Then
50110   SplitPath GSOutputFile, , Path, , FName, Ext
50120   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50130  End If
50140  tstr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50150  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50160   tstr = tstr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50170  End If
50180  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50190   tstr = tstr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50200  End If
50210  AddParams "-I" & tstr
50220  AddParams "-q"
50230  AddParams "-dNOPAUSE"
50240  AddParams "-dSAFER"
50250  AddParams "-dBATCH"
50260  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50270   AddParams "-sFONTPATH=" & GetFontsDirectory
50280  End If
50290
50300  AddParams "-sDEVICE=pswrite"
50310  If Options.DontUseDocumentSettings = 0 Then
50320   AddParams "-dLanguageLevel=" & GS_PSLanguageLevel
50330  End If
50340  AddParams "-sOutputFile=" & GSOutputFile
50350
50360  AddParams "-f"
50370  AddParams GSInputFile
50380  ShowParams
50390  CallGhostscript "PS"
50400 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50410 Exit Function
ErrPtnr_OnError:
50431 Select Case ErrPtnr.OnError("modGhostscript", "CreatePS")
      Case 0: Resume
50450 Case 1: Resume Next
50460 Case 2: Exit Function
50470 Case 3: End
50480 End Select
50490 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CreateEPS(GSInputFile As String, GSOutputFile As String, Options As tOptions)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Path As String, FName As String, Ext As String, tstr As String
50050
50060  GSInit Options
50070  InitParams
50080  Set ParamCommands = New Collection
50090
50100  If Options.OnePagePerFile = 1 Then
50110   SplitPath GSOutputFile, , Path, , FName, Ext
50120   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50130  End If
50140  tstr = Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50150  If LenB(LTrim(Options.DirectoryGhostscriptResource)) > 0 Then
50160   tstr = tstr & ";" & LTrim(Options.DirectoryGhostscriptResource)
50170  End If
50180  If LenB(LTrim(Options.AdditionalGhostscriptSearchpath)) > 0 Then
50190   tstr = tstr & ";" & LTrim(Options.AdditionalGhostscriptSearchpath)
50200  End If
50210  AddParams "-I" & tstr
50220  AddParams "-q"
50230  AddParams "-dNOPAUSE"
50240  AddParams "-dSAFER"
50250  AddParams "-dBATCH"
50260  If LenB(GetFontsDirectory) > 0 And Options.AddWindowsFontpath = 1 Then
50270   AddParams "-sFONTPATH=" & GetFontsDirectory
50280  End If
50290
50300  If Options.OnePagePerFile = 1 Then
50310   SplitPath GSOutputFile, , Path, , FName, Ext
50320   GSOutputFile = CompletePath(Path) & FName & "%d." & Ext
50330  End If
50340
50350  AddParams "-sDEVICE=epswrite"
50360  If Options.DontUseDocumentSettings = 0 Then
50370   AddParams "-dLanguageLevel=" & GS_EPSLanguageLevel
50380  End If
50390  AddParams "-sOutputFile=" & GSOutputFile
50400
50410  AddParams "-f"
50420  AddParams GSInputFile
50430  ShowParams
50440  CallGhostscript "EPS"
50450 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50460 Exit Function
ErrPtnr_OnError:
50481 Select Case ErrPtnr.OnError("modGhostscript", "CreateEPS")
      Case 0: Resume
50500 Case 1: Resume Next
50510 Case 2: Exit Function
50520 Case 3: End
50530 End Select
50540 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CallGScript(GSInputFile As String, GSOutputFile As String, _
 Options As tOptions, Ghostscriptdevice As tGhostscriptDevice)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim enc As Boolean, encPDF As EncryptData, retEnc As Boolean, _
  Tempfile As String, tL As Long
50060  GSInit Options
50071  Select Case Ghostscriptdevice
        Case 0: 'PDF
50090    With Options
50100     If .PDFOptimize = 1 And .PDFUseSecurity = 0 Then
50110       Tempfile = GetTempFile(GetTempPath, "~CP")
50120       KillFile Tempfile
50130       CreatePDF GSInputFile, Tempfile, Options
50140       OptimizePDF Tempfile, GSOutputFile
50150       KillFile Tempfile
50160      Else
50170       If .PDFUseSecurity <> 0 And SecurityIsPossible = True Then
50180         If .PDFEncryptor = 1 Then
50190           enc = SetEncryptionParams(encPDF, GSInputFile, GSOutputFile)
50200           If enc = True Then
50210            retEnc = EncryptPDF(encPDF)
50220            If retEnc = False Then
50230             IfLoggingWriteLogfile "Error with encryption - using unencrypted file"
50240             Name GSInputFile As GSOutputFile
50250            End If
50260           End If
50270          Else
50280           tL = .PDFOptimize
50290           .PDFOptimize = 0
50300           CreatePDF GSInputFile, GSOutputFile, Options
50310           .PDFOptimize = tL
50320         End If
50330        Else
50340         CreatePDF GSInputFile, GSOutputFile, Options
50350       End If
50360     End If
50370    End With
50380   Case 1: 'PNG
50390    CreatePNG GSInputFile, GSOutputFile, Options
50400   Case 2: 'JPEG
50410    CreateJPEG GSInputFile, GSOutputFile, Options
50420   Case 3: 'BMP
50430    CreateBMP GSInputFile, GSOutputFile, Options
50440   Case 4: 'PCX
50450    CreatePCX GSInputFile, GSOutputFile, Options
50460   Case 5: 'TIFF
50470    CreateTIFF GSInputFile, GSOutputFile, Options
50480   Case 6: 'PS
50490    CreatePS GSInputFile, GSOutputFile, Options
50500   Case 7: 'EPS
50510    CreateEPS GSInputFile, GSOutputFile, Options
50520  End Select
50530 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50540 Exit Function
ErrPtnr_OnError:
50561 Select Case ErrPtnr.OnError("modGhostscript", "CallGScript")
      Case 0: Resume
50580 Case 1: Resume Next
50590 Case 2: Exit Function
50600 Case 3: End
50610 End Select
50620 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function OptimizePDF(PDFInputFilename As String, PDFOutputFilename As String) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim LastStop As Currency
50050  InitParams
50060  AddParams "-q"
50070  AddParams "-dNODISPLAY"
50080  AddParams "-dSAFER"
50090  AddParams "-dDELAYSAFER"
50100  AddParams "--"
50110  AddParams "pdfopt.ps"
50120  AddParams PDFInputFilename
50130  AddParams PDFOutputFilename
50140
50150  GSParams(0) = "pdfopt"
50160  LastStop = ExactTimer_Value()
50170  OptimizePDF = CallGS(GSParams)
50180  IfLoggingWriteLogfile "Time for optimizing in seconds: " & CStr(ExactTimer_Value() - LastStop)
50190 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50200 Exit Function
ErrPtnr_OnError:
50221 Select Case ErrPtnr.OnError("modGhostscript", "OptimizePDF")
      Case 0: Resume
50240 Case 1: Resume Next
50250 Case 2: Exit Function
50260 Case 3: End
50270 End Select
50280 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub ReturnValue(data As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim newData As String
50050  newData = Replace(data, vbLf, "; ")
50060  IfLoggingWriteLogfile "Error: " & newData
50070 ' IfLoggingShowLogfile frmLog, frmMain
50080 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50090 Exit Sub
ErrPtnr_OnError:
50111 Select Case ErrPtnr.OnError("modGhostscript", "ReturnValue")
      Case 0: Resume
50130 Case 1: Resume Next
50140 Case 2: Exit Sub
50150 Case 3: End
50160 End Select
50170 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function Bool2Text(Number As Long)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If Number = 1 Then
50050   Bool2Text = "true"
50060  Else
50070   Bool2Text = "false"
50080  End If
50090 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50100 Exit Function
ErrPtnr_OnError:
50121 Select Case ErrPtnr.OnError("modGhostscript", "Bool2Text")
      Case 0: Resume
50140 Case 1: Resume Next
50150 Case 2: Exit Function
50160 Case 3: End
50170 End Select
50180 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub SelectColorCompression(ByVal gsMethod)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  GS_COMPRESSCOLORAUTO = "false"
50051  Select Case gsMethod
        Case 0
50070    GS_COMPRESSCOLORAUTO = "true"
50080    GS_COMPRESSCOLORMETHOD = "null"
50090    GS_COMPRESSCOLORLEVEL = "null"
50100   Case 1
50110    GS_COMPRESSCOLORMETHOD = "DCTEncode"
50120    GS_COMPRESSCOLORLEVEL = "Maximum"
50130   Case 2
50140    GS_COMPRESSCOLORMETHOD = "DCTEncode"
50150    GS_COMPRESSCOLORLEVEL = "High"
50160   Case 3
50170    GS_COMPRESSCOLORMETHOD = "DCTEncode"
50180    GS_COMPRESSCOLORLEVEL = "Medium"
50190   Case 4
50200    GS_COMPRESSCOLORMETHOD = "DCTEncode"
50210    GS_COMPRESSCOLORLEVEL = "Low"
50220   Case 5
50230    GS_COMPRESSCOLORMETHOD = "DCTEncode"
50240    GS_COMPRESSCOLORLEVEL = "Minimum"
50250   Case 6
50260    GS_COMPRESSCOLORMETHOD = "FlateEncode"
50270    GS_COMPRESSCOLORLEVEL = "Maximum"
50280  End Select
50290 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50300 Exit Sub
ErrPtnr_OnError:
50321 Select Case ErrPtnr.OnError("modGhostscript", "SelectColorCompression")
      Case 0: Resume
50340 Case 1: Resume Next
50350 Case 2: Exit Sub
50360 Case 3: End
50370 End Select
50380 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SelectGreyCompression(ByVal gsMethod)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  GS_COMPRESSGREYAUTO = "false"
50051  Select Case gsMethod
        Case 0
50070    GS_COMPRESSGREYAUTO = "true"
50080    GS_COMPRESSGREYMETHOD = "null"
50090    GS_COMPRESSGREYLEVEL = "null"
50100   Case 1
50110    GS_COMPRESSGREYMETHOD = "DCTEncode"
50120    GS_COMPRESSGREYLEVEL = "Maximum"
50130   Case 2
50140    GS_COMPRESSGREYMETHOD = "DCTEncode"
50150    GS_COMPRESSGREYLEVEL = "High"
50160   Case 3
50170    GS_COMPRESSGREYMETHOD = "DCTEncode"
50180    GS_COMPRESSGREYLEVEL = "Medium"
50190   Case 4
50200    GS_COMPRESSGREYMETHOD = "DCTEncode"
50210    GS_COMPRESSGREYLEVEL = "Low"
50220   Case 5
50230    GS_COMPRESSGREYMETHOD = "DCTEncode"
50240    GS_COMPRESSGREYLEVEL = "Minimum"
50250   Case 6
50260    GS_COMPRESSGREYMETHOD = "FlateEncode"
50270    GS_COMPRESSGREYLEVEL = "Maximum"
50280  End Select
50290 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50300 Exit Sub
ErrPtnr_OnError:
50321 Select Case ErrPtnr.OnError("modGhostscript", "SelectGreyCompression")
      Case 0: Resume
50340 Case 1: Resume Next
50350 Case 2: Exit Sub
50360 Case 3: End
50370 End Select
50380 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SelectMonoCompression(ByVal gsMethod)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50041  Select Case gsMethod
        Case 0
50060    GS_COMPRESSMONOMETHOD = "CCITTFaxEncode"
50070   Case 1
50080    GS_COMPRESSMONOMETHOD = "FlateEncode"
50090   Case 2
50100    GS_COMPRESSMONOMETHOD = "LZWEncode"
50110   Case 3
50120    GS_COMPRESSMONOMETHOD = "RunLengthEncode"
50130  End Select
50140 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50150 Exit Sub
ErrPtnr_OnError:
50171 Select Case ErrPtnr.OnError("modGhostscript", "SelectMonoCompression")
      Case 0: Resume
50190 Case 1: Resume Next
50200 Case 2: Exit Sub
50210 Case 3: End
50220 End Select
50230 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub InitParams()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  GSParamsIndex = 0
50050  ReDim GSParams(GSParamsIndex)
50060 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50070 Exit Sub
ErrPtnr_OnError:
50091 Select Case ErrPtnr.OnError("modGhostscript", "InitParams")
      Case 0: Resume
50110 Case 1: Resume Next
50120 Case 2: Exit Sub
50130 Case 3: End
50140 End Select
50150 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function ShowParams() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, tstr As String
50050  tstr = GSParams(LBound(GSParams))
50060  For i = LBound(GSParams) + 1 To UBound(GSParams)
50070   tstr = tstr & vbCrLf & GSParams(i)
50080  Next i
50090  IfLoggingWriteLogfile "Ghostscriptparameter:" & vbCrLf & tstr
50100  ShowParams = "Ghostscriptparameter:" & vbCrLf & tstr
50110 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50120 Exit Function
ErrPtnr_OnError:
50141 Select Case ErrPtnr.OnError("modGhostscript", "ShowParams")
      Case 0: Resume
50160 Case 1: Resume Next
50170 Case 2: Exit Function
50180 Case 3: End
50190 End Select
50200 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub AddParams(strValue As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  GSParamsIndex = GSParamsIndex + 1
50050  ReDim Preserve GSParams(GSParamsIndex)
50060  GSParams(GSParamsIndex) = strValue
50070 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50080 Exit Sub
ErrPtnr_OnError:
50101 Select Case ErrPtnr.OnError("modGhostscript", "AddParams")
      Case 0: Resume
50120 Case 1: Resume Next
50130 Case 2: Exit Sub
50140 Case 3: End
50150 End Select
50160 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function BuildPermissionString(encData As EncryptData) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim strPermissions As String
50050
50060  strPermissions = vbNullString
50070  strPermissions = strPermissions & Abs(Int(Not encData.DisallowPrinting))
50080  strPermissions = strPermissions & Abs(Int(Not encData.DisallowModifyContents))
50090  strPermissions = strPermissions & Abs(Int(Not encData.DisallowCopy))
50100  strPermissions = strPermissions & Abs(Int(Not encData.DisallowModifyAnnotations))
50110  If Options.PDFHighEncryption Then
50120    strPermissions = strPermissions & Abs(Int(encData.AllowFillIn)) '(128 bit only)
50130    strPermissions = strPermissions & Abs(Int(encData.AllowScreenReaders)) '(128 bit only)
50140    strPermissions = strPermissions & Abs(Int(encData.AllowAssembly)) '(128 bit only)
50150    strPermissions = strPermissions & Abs(Int(encData.AllowDegradedPrinting)) '(128 bit only)
50160   Else
50170    strPermissions = strPermissions & "0000"
50180  End If
50190  BuildPermissionString = strPermissions
50200 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50210 Exit Function
ErrPtnr_OnError:
50231 Select Case ErrPtnr.OnError("modGhostscript", "BuildPermissionString")
      Case 0: Resume
50250 Case 1: Resume Next
50260 Case 2: Exit Function
50270 Case 3: End
50280 End Select
50290 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function EncryptPDF(encData As EncryptData) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim strPermissions As String, strShell As String, ret As Double
50050
50060  strPermissions = BuildPermissionString(encData)
50070
50080 ' strShell = App.Path & "\pdfencrypt.exe """ & encData.InputFile & """ """ & encData.OutputFile & """ User=" & encData.UserPass & " Owner=" & encData.OwnerPass & " " & strPermissions & " " & encData.EncryptionLevel
50090 ' strShell = CompletePath(Options.DirectoryJava) & "Java.exe -cp """ & CompletePath(App.Path) & "iText.jar"" com.lowagie.tools.encrypt_pdf """ & encData.InputFile & """ """ & encData.OutputFile & """ """ & encData.UserPass & """ """ & encData.OwnerPass & """ " & strPermissions & " " & encData.EncryptionLevel
50100
50110  strShell = CompletePath(App.Path) & "pdfenc.exe """ & encData.InputFile & """ """ & encData.OutputFile & """ """ & encData.UserPass & """ """ & encData.OwnerPass & """ " & strPermissions & " " & encData.EncryptionLevel
50120
50130  IfLoggingWriteLogfile strShell
50140
50150  ret = RunProgramWait(strShell, False)
50160
50170  If Dir$(encData.OutputFile) <> vbNullString Then
50180   EncryptPDF = True
50190  End If
50200 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50210 Exit Function
ErrPtnr_OnError:
50231 Select Case ErrPtnr.OnError("modGhostscript", "EncryptPDF")
      Case 0: Resume
50250 Case 1: Resume Next
50260 Case 2: Exit Function
50270 Case 3: End
50280 End Select
50290 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CalculatePermissions(ByRef encData As EncryptData) As Long
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  With encData
50050   If .EncryptionLevel = encStrong Then
50060     CalculatePermissions = -4 - (Abs(.DisallowPrinting) * SEC_PRINT _
     + Abs(.DisallowModifyContents) * SEC_MODIFY _
     + Abs(.DisallowCopy) * SEC_COPY _
     + Abs(.DisallowModifyAnnotations) * SEC_FORM _
     + Abs(.AllowFillIn) * SEC_FILLFORM _
     + Abs(.AllowScreenReaders) * SEC_SCREENREADERS _
     + Abs(.AllowAssembly) * SEC_ASSEMBLY _
     + Abs(.AllowDegradedPrinting) * SEC_HQPRINT)
50140    Else
50150     CalculatePermissions = -4 - (Abs(.DisallowPrinting) * SEC_PRINT _
     + Abs(.DisallowModifyContents) * SEC_MODIFY _
     + Abs(.DisallowCopy) * SEC_COPY _
     + Abs(.DisallowModifyAnnotations) * SEC_FORM)
50190   End If
50200  End With
50210 ' Debug.Print "CP:" & CalculatePermissions
50220 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50230 Exit Function
ErrPtnr_OnError:
50251 Select Case ErrPtnr.OnError("modGhostscript", "CalculatePermissions")
      Case 0: Resume
50270 Case 1: Resume Next
50280 Case 2: Exit Function
50290 Case 3: End
50300 End Select
50310 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function SetEncryptionParams(ByRef encData As EncryptData, InputFile As String, OutputFile As String) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim retPasswd As Boolean
50050
50060  encData.InputFile = InputFile
50070  encData.OutputFile = OutputFile
50080
50090  If Len(Options.PDFOwnerPasswordString) > 0 Then
50100    encData.OwnerPass = Options.PDFOwnerPasswordString
50110    encData.UserPass = Options.PDFUserPasswordString
50120    OwnerPassword = Options.PDFOwnerPasswordString
50130    UserPassword = Options.PDFUserPasswordString
50140    retPasswd = True
50150   Else
50160    If SavePasswordsForThisSession = False Then
50170      If Options.UseAutosave = 0 Then
50180        retPasswd = EnterPasswords(encData.UserPass, encData.OwnerPass, frmPassword)
50190       Else
50200        retPasswd = False
50210      End If
50220     Else
50230      encData.OwnerPass = OwnerPassword: encData.UserPass = UserPassword
50240    End If
50250  End If
50260  If retPasswd = True Or SavePasswordsForThisSession = True Then
50270    With encData
50280     .DisallowPrinting = Options.PDFDisallowPrinting
50290     .DisallowModifyContents = Options.PDFDisallowModifyContents
50300     .DisallowCopy = Options.PDFDisallowCopy
50310     .DisallowModifyAnnotations = Options.PDFDisallowModifyAnnotations
50320     .AllowFillIn = Options.PDFAllowFillIn
50330     .AllowScreenReaders = Options.PDFAllowScreenReaders
50340     .AllowAssembly = Options.PDFAllowAssembly
50350     .AllowDegradedPrinting = Options.PDFAllowDegradedPrinting
50360     If Options.PDFHighEncryption = 1 Then
50370       .EncryptionLevel = encStrong
50380      Else
50390       .EncryptionLevel = encLow
50400     End If
50410    End With
50420    SetEncryptionParams = True
50430    encData.UserPass = UserPassword
50440    encData.OwnerPass = OwnerPassword
50450   Else
50460    SetEncryptionParams = False
50470  End If
50480 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50490 Exit Function
ErrPtnr_OnError:
50511 Select Case ErrPtnr.OnError("modGhostscript", "SetEncryptionParams")
      Case 0: Resume
50530 Case 1: Resume Next
50540 Case 2: Exit Function
50550 Case 3: End
50560 End Select
50570 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub SetColorParams()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If Options.PDFCompressionColorCompression = 1 Then
50050    AddParams "-dEncodeColorImages=true"
50060    If Options.PDFCompressionColorCompressionChoice = 0 Then
50070      AddParams "-dAutoFilterColorImages=true"
50080     Else
50090      AddParams "-dAutoFilterColorImages=false"
50100      If Options.PDFCompressionColorResample = 1 Then
50110        AddParams "-dDownsampleColorImages=true"
50121        Select Case Options.PDFCompressionColorResampleChoice
              Case 0:
50140          AddParams "-dColorImageDownsampleType=/Bicubic"
50150         Case 1:
50160          AddParams "-dColorImageDownsampleType=/Subsample"
50170         Case 2:
50180          AddParams "-dColorImageDownsampleType=/Average"
50190        End Select
50200        AddParams "-dColorImageResolution=" & Options.PDFCompressionColorResolution
50210       Else
50220        AddParams "-dDownsampleColorImages=false"
50230      End If
50241      Select Case Options.PDFCompressionColorCompressionChoice
            Case 1:
50260        AddParams "-dColorImageFilter=/DCTEncode"
50270        AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor 2 /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50280       Case 2:
50290        AddParams "-dColorImageFilter=/DCTEncode"
50300        AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor 0.9 /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50310       Case 3:
50320        AddParams "-dColorImageFilter=/DCTEncode"
50330        AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor 0.5 /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50340       Case 4:
50350        AddParams "-dColorImageFilter=/DCTEncode"
50360        AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor 0.25 /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
50370       Case 5:
50380        AddParams "-dColorImageFilter=/DCTEncode"
50390        AddParamCommand ".setpdfwrite << /ColorImageDict <</QFactor 0.1 /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
50400       Case 6:
50410        AddParams "-dColorImageFilter=/FlateEncode"
50420       Case 7:
50430        AddParams "-dColorImageFilter=/LZWEncode"
50440      End Select
50450    End If
50460   Else
50470    AddParams "-dEncodeColorImages=false"
50480  End If
50490 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50500 Exit Sub
ErrPtnr_OnError:
50521 Select Case ErrPtnr.OnError("modGhostscript", "SetColorParams")
      Case 0: Resume
50540 Case 1: Resume Next
50550 Case 2: Exit Sub
50560 Case 3: End
50570 End Select
50580 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetGreyParams()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If Options.PDFCompressionGreyCompression = 1 Then
50050    AddParams "-dEncodeGrayImages=true"
50060    If Options.PDFCompressionGreyCompressionChoice = 0 Then
50070      AddParams "-dAutoFilterGrayImages=true"
50080     Else
50090      AddParams "-dAutoFilterGrayImages=false"
50100      If Options.PDFCompressionGreyResample = 1 Then
50110        AddParams "-dDownsampleGrayImages=true"
50121        Select Case Options.PDFCompressionGreyResampleChoice
              Case 0:
50140          AddParams "-dGrayImageDownsampleType=/Bicubic"
50150         Case 1:
50160          AddParams "-dGrayImageDownsampleType=/Subsample"
50170         Case 2:
50180          AddParams "-dGrayImageDownsampleType=/Average"
50190        End Select
50200        AddParams "-dGrayImageResolution=" & Options.PDFCompressionGreyResolution
50210       Else
50220        AddParams "-dDownsampleGrayImages=false"
50230      End If
50241      Select Case Options.PDFCompressionGreyCompressionChoice
            Case 1:
50260        AddParams "-dGrayImageFilter=/DCTEncode"
50270        AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor 2 /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50280       Case 2:
50290        AddParams "-dGrayImageFilter=/DCTEncode"
50300        AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor 0.9 /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50310       Case 3:
50320        AddParams "-dGrayImageFilter=/DCTEncode"
50330        AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor 0.5 /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50340       Case 4:
50350        AddParams "-dGrayImageFilter=/DCTEncode"
50360        AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor 0.25 /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
50370       Case 5:
50380        AddParams "-dGrayImageFilter=/DCTEncode"
50390        AddParamCommand ".setpdfwrite << /GrayImageDict <</QFactor 0.1 /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
50400       Case 6:
50410        AddParams "-dGrayImageFilter=/FlateEncode"
50420       Case 7:
50430        AddParams "-dGrayImageFilter=/LZWEncode"
50440      End Select
50450    End If
50460   Else
50470    AddParams "-dEncodeGrayImages=false"
50480  End If
50490 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50500 Exit Sub
ErrPtnr_OnError:
50521 Select Case ErrPtnr.OnError("modGhostscript", "SetGreyParams")
      Case 0: Resume
50540 Case 1: Resume Next
50550 Case 2: Exit Sub
50560 Case 3: End
50570 End Select
50580 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetMonoParams()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If Options.PDFCompressionMonoCompression = 1 Then
50050    AddParams "-dEncodeMonoImages=true"
50061    Select Case Options.PDFCompressionMonoCompressionChoice
          Case 0:
50080      AddParams "-dMonoImageFilter=/CCITTFaxEncode"
50090     Case 1:
50100      AddParams "-dMonoImageFilter=/FlateEncode"
50110     Case 2:
50120      AddParams "-dMonoImageFilter=/LZWEncode"
50130     Case 3:
50140      AddParams "-dMonoImageFilter=/RunLengthEncode"
50150    End Select
50160    If Options.PDFCompressionMonoResample = 1 Then
50170      AddParams "-dDownsampleMonoImages=true"
50181      Select Case Options.PDFCompressionMonoResampleChoice
            Case 0:
50200        AddParams "-dMonoImageDownsampleType=/Bicubic"
50210       Case 1:
50220        AddParams "-dMonoImageDownsampleType=/Subsample"
50230       Case 2:
50240        AddParams "-dMonoImageDownsampleType=/Average"
50250      End Select
50260      AddParams "-dMonoImageResolution=" & Options.PDFCompressionMonoResolution
50270     Else
50280      AddParams "-dDownsampleMonoImages=false"
50290    End If
50300   Else
50310    AddParams "-dEncodeMonoImages=false"
50320  End If
50330 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50340 Exit Sub
ErrPtnr_OnError:
50361 Select Case ErrPtnr.OnError("modGhostscript", "SetMonoParams")
      Case 0: Resume
50380 Case 1: Resume Next
50390 Case 2: Exit Sub
50400 Case 3: End
50410 End Select
50420 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GhostScriptSecurity() As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  GhostScriptSecurity = False
50050  If LenB(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll) = 0 Then
50060   Exit Function
50070  End If
50080 ' GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50090  If GsDllLoaded = 0 Then
50100   Exit Function
50110  End If
50120  GSRevision = GetGhostscriptRevision
50130 ' UnLoadDLL GsDllLoaded
50140  If InStr(UCase$(GSRevision.strProduct), "AFPL") > 0 Then
50150   If GSRevision.intRevision < 814 Then
50160    Exit Function
50170   End If
50180   GhostScriptSecurity = True
50190   Exit Function
50200  End If
50210  If InStr(UCase$(GSRevision.strProduct), "GPL") > 0 Then
50220   If GSRevision.intRevision < 815 Then
50230    Exit Function
50240   End If
50250   GhostScriptSecurity = True
50260   Exit Function
50270  End If
50280 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50290 Exit Function
ErrPtnr_OnError:
50311 Select Case ErrPtnr.OnError("modGhostscript", "GhostScriptSecurity")
      Case 0: Resume
50330 Case 1: Resume Next
50340 Case 2: Exit Function
50350 Case 3: End
50360 End Select
50370 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetAllGhostscriptversions() As Collection
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim reg As clsRegistry, tstr As String, tColl As Collection, i As Long, _
  tf() As String, GS_DLL As String, GS_LIB As String, tB As Boolean, j As Long
50060  Set reg = New clsRegistry
50070  Set GetAllGhostscriptversions = New Collection
50080  With reg
50090   .hkey = HKEY_LOCAL_MACHINE
50100   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50110   If .KeyExists = True Then
50120    tstr = Trim$(.GetRegistryValue("GhostscriptCopyright"))
50130    If LenB(tstr) > 0 Then
50140     tstr = Replace$(LanguageStrings.OptionsGhostscriptInternal, "%1", tstr)
50150     tstr = Replace$(tstr, "%2", Trim$(.GetRegistryValue("GhostscriptVersion")))
50160     GetAllGhostscriptversions.Add tstr
50170    End If
50180   End If
50190   tstr = "AFPL Ghostscript"
50200   .KeyRoot = "SOFTWARE\" & tstr
50210   Set tColl = .EnumRegistryKeys(HKEY_LOCAL_MACHINE, .KeyRoot)
50220   For i = 1 To tColl.Count
50230    .Subkey = tColl.Item(i)
50240    GS_DLL = .GetRegistryValue("GS_DLL")
50250    GS_LIB = .GetRegistryValue("GS_LIB")
50260    If Len(GS_DLL) > 0 Then
50270     If FileExists(GS_DLL) = True Then
50280      If Len(GS_LIB) > 0 Then
50290       If InStr(GS_LIB, ";") > 0 Then
50300        tf = Split(GS_LIB, ";")
50310        tB = False
50320        For j = 0 To UBound(tf)
50330         If DirExists(tf(j)) = False Then
50340          tB = True
50350         End If
50360        Next j
50370        If tB = False Then
50380         GetAllGhostscriptversions.Add tstr & " " & tColl.Item(i)
50390        End If
50400       End If
50410      End If
50420     End If
50430    End If
50440   Next i
50450   tstr = "GNU Ghostscript"
50460   .KeyRoot = "SOFTWARE\" & tstr
50470   Set tColl = .EnumRegistryKeys(HKEY_LOCAL_MACHINE, .KeyRoot)
50480   For i = 1 To tColl.Count
50490    .Subkey = tColl.Item(i)
50500    GS_DLL = .GetRegistryValue("GS_DLL")
50510    GS_LIB = .GetRegistryValue("GS_LIB")
50520    If Len(GS_DLL) > 0 Then
50530     If FileExists(GS_DLL) = True Then
50540      If Len(GS_LIB) > 0 Then
50550       If InStr(GS_LIB, ";") > 0 Then
50560        tf = Split(GS_LIB, ";")
50570        tB = False
50580        For j = 0 To UBound(tf)
50590         If DirExists(tf(j)) = False Then
50600          tB = True
50610         End If
50620        Next j
50630        If tB = False Then
50640         GetAllGhostscriptversions.Add tstr & " " & tColl.Item(i)
50650        End If
50660       End If
50670      End If
50680     End If
50690    End If
50700   Next i
50710   tstr = "GPL Ghostscript"
50720   .KeyRoot = "SOFTWARE\" & tstr
50730   Set tColl = .EnumRegistryKeys(HKEY_LOCAL_MACHINE, .KeyRoot)
50740   For i = 1 To tColl.Count
50750    .Subkey = tColl.Item(i)
50760    GS_DLL = .GetRegistryValue("GS_DLL")
50770    GS_LIB = .GetRegistryValue("GS_LIB")
50780    If Len(GS_DLL) > 0 Then
50790     If FileExists(GS_DLL) = True Then
50800      If Len(GS_LIB) > 0 Then
50810       If InStr(GS_LIB, ";") > 0 Then
50820        tf = Split(GS_LIB, ";")
50830        tB = False
50840        For j = 0 To UBound(tf)
50850         If DirExists(tf(j)) = False Then
50860          tB = True
50870         End If
50880        Next j
50890        If tB = False Then
50900         GetAllGhostscriptversions.Add tstr & " " & tColl.Item(i)
50910        End If
50920       End If
50930      End If
50940     End If
50950    End If
50960   Next i
50970  End With
50980  Set reg = Nothing
50990 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
51000 Exit Function
ErrPtnr_OnError:
51021 Select Case ErrPtnr.OnError("modGhostscript", "GetAllGhostscriptversions")
      Case 0: Resume
51040 Case 1: Resume Next
51050 Case 2: Exit Function
51060 Case 3: End
51070 End Select
51080 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub CheckForStamping(Filename As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim StampPage As String, tstr As String, R As String, G As String, B As String, _
  Stampfile As String, Path As String, ff As Long, Files As Collection, _
  StampString As String, StampFontsize As Double, _
  StampOutlineFontthickness As Double
50080  StampString = RemoveLeadingAndTrailingQuotes(Trim$(Options.StampString))
50090  If Len(StampString) > 0 Then
50100   StampPage = StrConv(LoadResData(101, "STAMPPAGE"), vbUnicode)
50110   StampPage = Replace(StampPage, vbCrLf, vbCr, , , vbBinaryCompare)
50120   StampPage = Replace(StampPage, "[STAMPSTRING]", StampString, , , vbTextCompare)
50130   StampPage = Replace(StampPage, "[FONTNAME]", Trim$(Options.StampFontname), , , vbTextCompare)
50140   StampFontsize = 48
50150   If IsNumeric(Options.StampFontsize) = True Then
50160    If CDbl(Options.StampFontsize) > 0 Then
50170     StampFontsize = CDbl(Options.StampFontsize)
50180    End If
50190   End If
50200   StampPage = Replace(StampPage, "[FONTSIZE]", StampFontsize, , , vbTextCompare)
50210   StampOutlineFontthickness = 0
50220   If IsNumeric(Options.StampOutlineFontthickness) = True Then
50230    If CDbl(Options.StampOutlineFontthickness) >= 0 Then
50240     StampOutlineFontthickness = CDbl(Options.StampOutlineFontthickness)
50250    End If
50260   End If
50270   StampPage = Replace(StampPage, "[STAMPOUTLINEFONTTHICKNESS]", StampOutlineFontthickness, , , vbTextCompare)
50280   If Options.StampUseOutlineFont <> 1 Then
50290     StampPage = Replace(StampPage, "[USEOUTLINEFONT]", "show", , , vbTextCompare)
50300    Else
50310     StampPage = Replace(StampPage, "[USEOUTLINEFONT]", "true charpath stroke", , , vbTextCompare)
50320   End If
50330   If Len(Options.StampFontColor) > 0 Then
50340     tstr = Replace$(Options.StampFontColor, "#", "&H")
50350     If IsNumeric(tstr) = True Then
50360       R = Replace$(Format(CDbl((CLng(tstr) And CLng("&HFF0000")) / 65536) / 255#, "0.00"), ",", ".", , 1)
50370       G = Replace$(Format(CDbl((CLng(tstr) And CLng("&H00FF00")) / 256) / 255#, "0.00"), ",", ".", , 1)
50380       B = Replace$(Format(CDbl(CLng(tstr) And CLng("&H0000FF")) / 255#, "0.00"), ",", ".", , 1)
50390       StampPage = Replace(StampPage, "[FONTCOLOR]", R & " " & G & " " & B, , , vbTextCompare)
50400      Else
50410       StampPage = Replace(StampPage, "[FONTCOLOR]", "1 0 0", , , vbTextCompare)
50420     End If
50430    Else
50440     StampPage = Replace(StampPage, "[FONTCOLOR]", "1 0 0", , , vbTextCompare)
50450   End If
50460   Path = CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & GetUsername
50470   If DirExists(Path) = False Then
50480    MakePath Path
50490   End If
50500   Stampfile = GetTempFile(Path, "~ST")
50510   ff = FreeFile
50520   Open Stampfile For Output As #ff
50530   Print #ff, StampPage
50540   Close #ff
50550   Set Files = New Collection
50560   Files.Add Stampfile
50570   Files.Add Filename
50580   Stampfile = GetTempFile(Path, "~ST")
50590   KillFile Stampfile
50600   CombineFiles Stampfile, Files
50610   Name Stampfile As Filename
50620  End If
50630 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50640 Exit Sub
ErrPtnr_OnError:
50661 Select Case ErrPtnr.OnError("modGhostscript", "CheckForStamping")
      Case 0: Resume
50680 Case 1: Resume Next
50690 Case 2: Exit Sub
50700 Case 3: End
50710 End Select
50720 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AddParamCommands()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long
50050  If ParamCommands.Count > 0 Then
50060   AddParams "-c"
50070   For i = 1 To ParamCommands.Count
50080    AddParams ParamCommands(i)
50090   Next i
50100  End If
50110 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50120 Exit Sub
ErrPtnr_OnError:
50141 Select Case ErrPtnr.OnError("modGhostscript", "AddParamCommands")
      Case 0: Resume
50160 Case 1: Resume Next
50170 Case 2: Exit Sub
50180 Case 3: End
50190 End Select
50200 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AddParamCommand(GhostscriptCommand As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  ParamCommands.Add GhostscriptCommand
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("modGhostscript", "AddParamCommand")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetTestpageFromRessource() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GetTestpageFromRessource = _
  Replace(StrConv(LoadResData(101, "TESTPAGE"), vbUnicode), vbCrLf, vbLf, , , vbBinaryCompare)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGhostscript", "GetTestpageFromRessource")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
