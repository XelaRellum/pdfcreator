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
50010  Dim gsret As Long, i As Long, enc As Boolean
50020
50030 GSInit Options
50040 InitParams
50050
50060 AddParams "-I" & Options.DirectoryGhostscriptLibraries & ";" & Options.DirectoryGhostscriptFonts
50070 AddParams "-q"
50080 AddParams "-dNOPAUSE"
50090 AddParams "-dSAFER"
50100 AddParams "-dBATCH"
50111 Select Case Ghostscriptdevice
       Case 0: 'PDF
50130   AddParams "-sDEVICE=pdfwrite"
50140 '  AddParams "-dPDFSETTINGS=/printer"
50150   AddParams "-dCompatibilityLevel=" & GS_COMPATIBILITY
50160   AddParams "-r" & GS_RESOLUTION & "x" & GS_RESOLUTION
50170   AddParams "-dProcessColorModel=/Device" & GS_COLORMODEL
50180   AddParams "-dAutoRotatePages=/" & GS_AUTOROTATE
50190   AddParams "-dCompressPages=" & GS_COMPRESSPAGES
50200   AddParams "-dEmbedAllFonts=" & GS_EMBEDALLFONTS
50210   AddParams "-dSubsetFonts=" & GS_SUBSETFONTS
50220   AddParams "-dMaxSubsetPct=" & GS_SUBSETFONTPERC
50230   AddParams "-dConvertCMYKImagesToRGB=" & GS_CMYKTORGB
50240
50250   If Options.PDFUseSecurity <> 0 And SecurityIsPossible = True Then
50260     If Options.PDFEncryptor > 0 Then
50270       Dim Tempfile As String, Temppath As String, encPDF As EncryptData, retEnc As Boolean
50280       Tempfile = GetTempFile(GetTempPath, "~PDF")
50290       AddParams "-sOutputFile=" & Tempfile
50300      Else
50310       If SetEncryptionParams(encPDF, "", "") = True Then
50320         If Len(encPDF.OwnerPass) > 0 Then
50330          AddParams "-sOwnerPassword=" & encPDF.OwnerPass & ""
50340         End If
50350         If Len(encPDF.UserPass) > 0 Then
50360          AddParams "-sUserPassword=" & encPDF.UserPass
50370         End If
50380         AddParams "-dPermissions=" & CalculatePermissions(encPDF)
50390 '        Debug.Print BuildPermissionString(encPDF), CalculatePermissions(encPDF)
50400         If GS_COMPATIBILITY = "1.4" Then
50410           AddParams "-dEncryptionR=3"
50420          Else
50430           AddParams "-dEncryptionR=2"
50440         End If
50450         If encPDF.EncryptionLevel = encLow Then
50460           AddParams "-dKeyLength=40"
50470          Else
50480           AddParams "-dKeyLength=128"
50490         End If
50500        Else
50510         MsgBox LanguageStrings.MessagesMsg23, vbCritical
50520       End If
50530
50540       AddParams "-sOutputFile=" & GSOutputFile
50550 '      AddParams "-c .setpdfwrite"
50560     End If
50570    Else
50580     AddParams "-sOutputFile=" & GSOutputFile
50590   End If
50600
50610   SetColorParams
50620   SetGreyParams
50630   SetMonoParams
50640
50650
50660 '  AddParams "-dGrayACSImageDict " & GS_COMPRESSGREYLEVEL
50670
50680   AddParams "-dPreserveOverprintSettings=" & GS_PRESERVEOVERPRINT
50690   AddParams "-dUCRandBGInfo=/Preserve"
50700   AddParams "-dUseFlateCompression=true"
50710   AddParams "-dParseDSCCommentsForDocInfo=true"
50720   AddParams "-dParseDSCComments=true"
50730   AddParams "-dOPM=" & GS_OVERPRINT
50740   AddParams "-dOffOptimizations=0"
50750   AddParams "-dLockDistillerParams=false"
50760   AddParams "-dGrayImageDepth=-1"
50770   AddParams "-dASCII85EncodePages=" & GS_ASCII85
50780   AddParams "-dDefaultRenderingIntent=/Default"
50790   AddParams "-dTransferFunctionInfo=/" & GS_TRANSFERFUNCTIONS
50800   AddParams "-dPreserveHalftoneInfo=" & GS_HALFTONE
50810   AddParams "-dOptimize=true"
50820   AddParams "-dDetectBlends=true"
50830
50840   AddParams "-f"
50850   AddParams GSInputFile
50860   ShowParams
50870  Case 1: 'PNG
50880   AddParams "-sDEVICE=" & GS_PNGColorscount
50890   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
50900   AddParams "-sOutputFile=" & GSOutputFile
50910   AddParams GSInputFile
50920  Case 2: 'JPEG
50930   AddParams "-sDEVICE=" & GS_JPEGColorscount
50940   AddParams "-dJPEGQ=" & GS_JPEGQuality
50950   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
50960   AddParams "-sOutputFile=" & GSOutputFile
50970   AddParams GSInputFile
50980  Case 3: 'BMP
50990   AddParams "-sDEVICE=" & GS_BMPColorscount
51000   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
51010   AddParams "-q"
51020   AddParams "-sOutputFile=" & GSOutputFile
51030   AddParams GSInputFile
51040  Case 4: 'PCX
51050   AddParams "-sDEVICE=" & GS_PCXColorscount
51060   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
51070   AddParams "-sOutputFile=" & GSOutputFile
51080   AddParams GSInputFile
51090  Case 5: 'TIFF
51100   AddParams "-sDEVICE=" & GS_TIFFColorscount
51110   AddParams "-r" & GS_BitmapRESOLUTION & "x" & GS_BitmapRESOLUTION
51120   AddParams "-sOutputFile=" & GSOutputFile
51130   AddParams GSInputFile
51140  Case 6: 'PS
51150   AddParams "-dLanguageLevel=" & GS_PSLanguageLevel
51160   AddParams "-sDEVICE=pswrite"
51170   AddParams "-sOutputFile=" & GSOutputFile
51180   AddParams GSInputFile
51190  Case 7: 'EPS
51200   AddParams "-dLanguageLevel=" & GS_EPSLanguageLevel
51210   AddParams "-sDEVICE=epswrite"
51220   AddParams "-sOutputFile=" & GSOutputFile
51230   AddParams GSInputFile
51240 End Select
51250
51260 gsret = CallGS(GSParams)
51270
51280 If (Options.PDFUseSecurity <> 0) And (Ghostscriptdevice = PDFWriter) And _
   (Options.PDFEncryptor > 0) And SecurityIsPossible = True Then
51300  If Len(Dir(GSOutputFile)) > 0 Then
51310   Kill GSOutputFile
51320  End If
51330  enc = SetEncryptionParams(encPDF, Tempfile, GSOutputFile)
51340  If enc = True Then
51350    retEnc = EncryptPDF(encPDF)
51360    If retEnc = False Then
51370     FileCopy Tempfile, GSOutputFile
51380     IfLoggingWriteLogfile "Error with encryption - using unencrypted file"
51390    End If
51400    Kill Tempfile
51410   Else
51420    MsgBox LanguageStrings.MessagesMsg23, vbCritical
51430    If Len(Dir(GSOutputFile)) > 0 Then
51440     Kill GSOutputFile
51450    End If
51460    Name Tempfile As GSOutputFile
51470  End If
51480 End If
51490
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
50010  Dim i As Long, tStr As String, fn As Long
50020  tStr = GSParams(LBound(GSParams))
50030  For i = LBound(GSParams) + 1 To UBound(GSParams)
50040   tStr = tStr & vbCrLf & GSParams(i)
50050  Next i
50060  IfLoggingWriteLogfile "Ghostscriptparameter:" & vbCrLf & tStr
50070 ' fn = FreeFile
50080 ' Open CompletePath(App.Path) & "\params.txt" For Output As #1
50090 ' Close #1
50100 ' OpenDocument CompletePath(App.Path) & "\params.txt"
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
50080  strShell = CompletePath(App.Path) & "pdfenc.exe """ & encData.InputFile & """ """ & encData.OutputFile & """ """ & encData.UserPass & """ """ & encData.OwnerPass & """ " & strPermissions & " " & encData.EncryptionLevel
50090
50100  IfLoggingWriteLogfile strShell
50110
50120  ret = RunProgramWait(strShell, False)
50130
50140  If Dir$(encData.OutputFile) <> "" Then
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
50010  With encData
50020   If .EncryptionLevel = encStrong Then
50030     CalculatePermissions = -4 - (Abs(.DisallowPrinting) * SEC_PRINT _
     + Abs(.DisallowModifyContents) * SEC_MODIFY _
     + Abs(.DisallowCopy) * SEC_COPY _
     + Abs(.DisallowModifyAnnotations) * SEC_FORM _
     + Abs(.AllowFillIn) * SEC_FILLFORM _
     + Abs(.AllowScreenReaders) * SEC_SCREENREADERS _
     + Abs(.AllowAssembly) * SEC_ASSEMBLY _
     + Abs(.AllowDegradedPrinting) * SEC_HQPRINT)
50110    Else
50120     CalculatePermissions = -4 - (Abs(.DisallowPrinting) * SEC_PRINT _
     + Abs(.DisallowModifyContents) * SEC_MODIFY _
     + Abs(.DisallowCopy) * SEC_COPY _
     + Abs(.DisallowModifyAnnotations) * SEC_FORM)
50160   End If
50170  End With
50180 ' Debug.Print "CP:" & CalculatePermissions
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
50060  If SavePasswordsForThisSession = False Then
50070    retPasswd = EnterPasswords(encData.UserPass, encData.OwnerPass, frmPassword)
50080   Else
50090    encData.OwnerPass = OwnerPassword: encData.UserPass = UserPassword
50100  End If
50110  If retPasswd = True Or SavePasswordsForThisSession = True Then
50120    With encData
50130     .DisallowPrinting = Options.PDFDisallowPrinting
50140     .DisallowModifyContents = Options.PDFDisallowModifyContents
50150     .DisallowCopy = Options.PDFDisallowCopy
50160     .DisallowModifyAnnotations = Options.PDFDisallowModifyAnnotations
50170     .AllowFillIn = Options.PDFAllowFillIn
50180     .AllowScreenReaders = Options.PDFAllowScreenReaders
50190     .AllowAssembly = Options.PDFAllowAssembly
50200     .AllowDegradedPrinting = Options.PDFAllowDegradedPrinting
50210     If Options.PDFHighEncryption = 1 Then
50220       .EncryptionLevel = encStrong
50230      Else
50240       .EncryptionLevel = encLow
50250     End If
50260    End With
50270    SetEncryptionParams = True
50280    encData.UserPass = UserPassword
50290    encData.OwnerPass = OwnerPassword
50300   Else
50310    SetEncryptionParams = False
50320  End If
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
50071      Select Case Options.PDFCompressionColorCompressionChoice
            Case 1:
50090        AddParams "-dColorImageFilter=/DCTEncode"
50100        AddParams "-c"
50110        AddParams ".setpdfwrite << /ColorImageDict <</QFactor 1.3 /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50120       Case 2:
50130        AddParams "-dColorImageFilter=/DCTEncode"
50140        AddParams "-c"
50150        AddParams ".setpdfwrite << /ColorImageDict <</QFactor 0.9 /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50160       Case 3:
50170        AddParams "-dColorImageFilter=/DCTEncode"
50180        AddParams "-c"
50190        AddParams ".setpdfwrite << /ColorImageDict <</QFactor 0.5 /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50200       Case 4:
50210        AddParams "-dColorImageFilter=/DCTEncode"
50220        AddParams "-c"
50230        AddParams ".setpdfwrite << /ColorImageDict <</QFactor 0.25 /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
50240       Case 5:
50250        AddParams "-dColorImageFilter=/DCTEncode"
50260        AddParams "-c"
50270        AddParams ".setpdfwrite << /ColorImageDict <</QFactor 0.1 /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
50280       Case 6:
50290        AddParams "-dColorImageFilter=/FlateEncode"
50300       Case 7:
50310        AddParams "-dColorImageFilter=/LZWEncode"
50320      End Select
50330      If Options.PDFCompressionColorResample = 1 Then
50340        AddParams "-dDownsampleColorImages=true"
50351        Select Case Options.PDFCompressionColorResampleChoice
              Case 0:
50370          AddParams "-dColorImageDownsampleType=/Bicubic"
50380         Case 1:
50390          AddParams "-dColorImageDownsampleType=/Subsample"
50400         Case 2:
50410          AddParams "-dColorImageDownsampleType=/Average"
50420        End Select
50430        AddParams "-dColorImageResolution=" & Options.PDFCompressionColorResolution
50440       Else
50450        AddParams "-dDownsampleColorImages=false"
50460      End If
50470    End If
50480   Else
50490    AddParams "-dEncodeColorImages=false"
50500  End If
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
50071      Select Case Options.PDFCompressionGreyCompressionChoice
            Case 1:
50090        AddParams "-dGrayImageFilter=/DCTEncode"
50100        AddParams "-c"
50110        AddParams ".setpdfwrite << /GrayImageDict <</QFactor 1.3 /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50120       Case 2:
50130        AddParams "-dGrayImageFilter=/DCTEncode"
50140        AddParams "-c"
50150        AddParams ".setpdfwrite << /GrayImageDict <</QFactor 0.9 /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50160       Case 3:
50170        AddParams "-dGrayImageFilter=/DCTEncode"
50180        AddParams "-c"
50190        AddParams ".setpdfwrite << /GrayImageDict <</QFactor 0.5 /Blend 1 /HSample [2 1 1 2] /VSample [2 1 1 2]>> >> setdistillerparams"
50200       Case 4:
50210        AddParams "-dGrayImageFilter=/DCTEncode"
50220        AddParams "-c"
50230        AddParams ".setpdfwrite << /GrayImageDict <</QFactor 0.25 /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
50240       Case 5:
50250        AddParams "-dGrayImageFilter=/DCTEncode"
50260        AddParams "-c"
50270        AddParams ".setpdfwrite << /GrayImageDict <</QFactor 0.1 /Blend 0 /HSample [1 1 1 1] /VSample [1 1 1 1]>> >> setdistillerparams"
50280       Case 6:
50290        AddParams "-dGrayImageFilter=/FlateEncode"
50300       Case 7:
50310        AddParams "-dGrayImageFilter=/LZWEncode"
50320      End Select
50330      If Options.PDFCompressionGreyResample = 1 Then
50340        AddParams "-dDownsampleGrayImages=true"
50351        Select Case Options.PDFCompressionGreyResampleChoice
              Case 0:
50370          AddParams "-dGrayImageDownsampleType=/Bicubic"
50380         Case 1:
50390          AddParams "-dGrayImageDownsampleType=/Subsample"
50400         Case 2:
50410          AddParams "-dGrayImageDownsampleType=/Average"
50420        End Select
50430        AddParams "-dGrayImageResolution=" & Options.PDFCompressionGreyResolution
50440       Else
50450        AddParams "-dDownsampleGrayImages=false"
50460      End If
50470    End If
50480   Else
50490    AddParams "-dEncodeGrayImages=false"
50500  End If
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
50030    If Options.PDFCompressionMonoCompressionChoice = 0 Then
50040      AddParams "-dAutoFilterMonoImages=true"
50050     Else
50060      AddParams "-dAutoFilterMonoImages=false"
50071      Select Case Options.PDFCompressionMonoCompressionChoice
            Case 1:
50090        AddParams "-dMonoImageFilter=/CCITTFaxEncode"
50100       Case 2:
50110        AddParams "-dMonoImageFilter=/FlateEncode"
50120       Case 3:
50130        AddParams "-dMonoImageFilter=/LZWEncode"
50140       Case 4:
50150        AddParams "-dMonoImageFilter=/RunLengthEncode"
50160      End Select
50170      If Options.PDFCompressionMonoResample = 1 Then
50180        AddParams "-dDownsampleMonoImages=true"
50191        Select Case Options.PDFCompressionMonoResampleChoice
              Case 0:
50210          AddParams "-dMonoImageDownsampleType=/Bicubic"
50220         Case 1:
50230          AddParams "-dMonoImageDownsampleType=/Subsample"
50240         Case 2:
50250          AddParams "-dMonoImageDownsampleType=/Average"
50260        End Select
50270        AddParams "-dMonoImageResolution=" & Options.PDFCompressionMonoResolution
50280       Else
50290        AddParams "-dDownsampleMonoImages=false"
50300      End If
50310    End If
50320   Else
50330    AddParams "-dEncodeMonoImages=false"
50340  End If
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
50010  Dim reg As clsRegistry, tStr As String, tColl As Collection, i As Long
50020  Set reg = New clsRegistry
50030  Set GetAllGhostscriptversions = New Collection
50040  With reg
50050   .hkey = HKEY_LOCAL_MACHINE
50060   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50070   If .KeyExists = True Then
50080    tStr = Trim$(.GetRegistryValue("GhostscriptCopyright"))
50090    If LenB(tStr) > 0 Then
50100     tStr = Replace$(LanguageStrings.OptionsGhostscriptInternal, "%1", tStr)
50110     tStr = Replace$(tStr, "%2", Trim$(.GetRegistryValue("GhostscriptVersion")))
50120     GetAllGhostscriptversions.Add tStr
50130    End If
50140   End If
50150   tStr = "AFPL Ghostscript"
50160   .KeyRoot = "SOFTWARE\" & tStr
50170   Set tColl = .EnumRegistryKeys(HKEY_LOCAL_MACHINE, .KeyRoot)
50180   For i = 1 To tColl.Count
50190    GetAllGhostscriptversions.Add tStr & " " & tColl.item(i)
50200   Next i
50210   tStr = "GNU Ghostscript"
50220   .KeyRoot = "SOFTWARE\" & tStr
50230   Set tColl = .EnumRegistryKeys(HKEY_LOCAL_MACHINE, .KeyRoot)
50240   For i = 1 To tColl.Count
50250    GetAllGhostscriptversions.Add tStr & " " & tColl.item(i)
50260   Next i
50270   tStr = "GPL Ghostscript"
50280   .KeyRoot = "SOFTWARE\" & tStr
50290   Set tColl = .EnumRegistryKeys(HKEY_LOCAL_MACHINE, .KeyRoot)
50300   For i = 1 To tColl.Count
50310    GetAllGhostscriptversions.Add tStr & " " & tColl.item(i)
50320   Next i
50330  End With
50340
50350  Set reg = Nothing
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

