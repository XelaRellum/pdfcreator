Attribute VB_Name = "modGhostScript"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetTempPath Lib "kernel32" Alias _
  "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer _
  As String) As Long

Public Declare Function GetTempFileName Lib "kernel32" Alias _
  "GetTempFileNameA" (ByVal lpszPath As String, ByVal _
  lpPrefixString As String, ByVal wUnique As Long, ByVal _
  lpTempFileName As String) As Long
  
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOSIZE = &H1

Global DOKUMENT_PS As String
Global DOKUMENT_PDF As String
Global DOKUMENT_ALL As String

Global EMAIL_NAME As String

Global SELECT_FILE As String
Global SAVE_FILE As String

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

Public Sub GSInit()
'General
GS_COMPATIBILITY = "1." & (frmOptions.cmbCompat.ListIndex + 2)
GS_RESOLUTION = frmOptions.txtRes
GS_AUTOROTATE = Tag2Text(frmOptions.cmbRotate.Tag, frmOptions.cmbRotate.ListIndex)
GS_OVERPRINT = frmOptions.cmbOverprint.ListIndex
GS_ASCII85 = Bool2Text(frmOptions.chkASCII85.Value)

'Compression
GS_COMPRESSPAGES = Bool2Text(frmOptions.chkTextComp.Value)
GS_COMPRESSCOLOR = Bool2Text(frmOptions.chkColorComp.Value)
GS_COMPRESSGREY = Bool2Text(frmOptions.chkGreyComp.Value)
GS_COMPRESSMONO = Bool2Text(frmOptions.chkMonoComp.Value)

SelectColorCompression frmOptions.cmbColorComp.ListIndex
SelectGreyCompression frmOptions.cmbGreyComp.ListIndex
'GS_COMPRESSMONOMETHOD = Bool2Text(frmOptions.chkMonoComp.Value)

GS_COMPRESSCOLORVALUE = Bool2Text(frmOptions.chkColorComp.Value)
GS_COMPRESSGREYVALUE = Bool2Text(frmOptions.chkGreyComp.Value)
GS_COMPRESSMONOVALUE = Bool2Text(frmOptions.chkMonoComp.Value)

GS_COLORRESOLUTION = frmOptions.txtColorRes.Text
GS_GREYRESOLUTION = frmOptions.txtGreyRes.Text
GS_MONORESOLUTION = frmOptions.txtMonoRes.Text

GS_COLORRESAMPLE = Bool2Text(frmOptions.chkColorResample.Value)
GS_GREYRESAMPLE = Bool2Text(frmOptions.chkGreyResample.Value)
GS_MONORESAMPLE = Bool2Text(frmOptions.chkMonoResample.Value)

GS_COLORRESAMPLEMETHOD = Tag2Text(frmOptions.cmbColorResample.Tag, frmOptions.cmbColorResample.ListIndex)
GS_GREYRESAMPLEMETHOD = Tag2Text(frmOptions.cmbGreyResample.Tag, frmOptions.cmbGreyResample.ListIndex)
GS_MONORESAMPLEMETHOD = Tag2Text(frmOptions.cmbMonoResample.Tag, frmOptions.cmbMonoResample.ListIndex)

'Fonts
GS_EMBEDALLFONTS = Bool2Text(frmOptions.chkEmbedAll.Value)
GS_SUBSETFONTS = Bool2Text(frmOptions.chkSubSetFonts.Value)
GS_SUBSETFONTPERC = frmOptions.txtSubSetPerc.Text

'Colors
GS_COLORMODEL = Tag2Text(frmOptions.cmbColorModel.Tag, frmOptions.cmbColorModel.ListIndex)
GS_CMYKTORGB = Bool2Text(frmOptions.chkCMYKtoRGB.Value)
GS_PRESERVEOVERPRINT = Bool2Text(frmOptions.chkPreserveOverprint.Value)
GS_TRANSFERFUNCTIONS = Tag2Text(frmOptions.chkPreserveTransfer.Tag, frmOptions.chkPreserveTransfer.Value)
GS_HALFTONE = Bool2Text(frmOptions.chkPreserverHalftone.Value)

'Other
GS_ERROR = 0
UseReturnPipe = 1
End Sub

Public Function CallGScript(GSInputFile As String, GSOutputFile As String)
Dim gsret

GSInit
InitParams

AddParams "-I" & App.Path & "\lib;" & App.Path & "\fonts"
AddParams "-q"
AddParams "-dNOPAUSE"
AddParams "-dSAFER"
AddParams "-dBATCH"
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
AddParams "-q"
AddParams "-dNOPAUSE"
AddParams "-dSAFER"
AddParams "-sOutputFile=" & GSOutputFile
AddParams GSInputFile
AddParams "-c quit"

frmMain.Refresh
gsret = CallGS(GSParams)
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


Public Function GetTitle(FileName As String)
Dim PosTitle
Dim PosTitleEnd
Dim strTM As String

Open FileName For Input As #1
   strTM = Input(5000, #1)
Close #1   ' Datei schlieﬂen.

PosTitle = InStr(1, strTM, "%%Title: ", vbTextCompare)
PosTitleEnd = InStr(PosTitle, strTM, vbLf, vbTextCompare)
If Mid$(strTM, PosTitleEnd - 1, 1) = vbCr Then PosTitleEnd = PosTitleEnd - 1

If PosTitle = 0 Then Exit Function
GetTitle = Mid$(strTM, PosTitle + 9, PosTitleEnd - PosTitle - 9)
End Function

Public Sub SaveTitle(FileName, NewTitle)
Dim PosTitle
Dim PosTitleEnd
Dim PosTitleEnd2
Dim strTM As String
Dim OldTitle As String

Open FileName For Input As #1
   strTM = Input(LOF(1) - 1, #1)
Close #1   ' Datei schlieﬂen.

PosTitle = InStr(1, strTM, "%%Title: ", vbTextCompare)
PosTitleEnd = InStr(PosTitle, strTM, vbCr, vbTextCompare)
PosTitleEnd2 = InStr(PosTitle, strTM, vbLf, vbTextCompare)

If PosTitleEnd = 0 Then PosTitleEnd = PosTitleEnd2

OldTitle = Mid$(strTM, PosTitle + 9, PosTitleEnd - PosTitle - 9)

If PosTitle = 0 Then Exit Sub
strTM = Replace$(strTM, OldTitle, NewTitle, 1, 1)

Open FileName For Output As #1
Print #1, strTM
Close #1
End Sub

Public Function Bool2Text(number As Integer)
If number = 1 Then
Bool2Text = "true"
Else
Bool2Text = "false"
End If
End Function

Public Function Tag2Text(Tag As String, lIndex As Integer)
Dim RetSplit As Variant
RetSplit = Split(Tag, "|")
Tag2Text = RetSplit(lIndex)
End Function

Public Sub ReturnValue(data As String)
Dim newData As String
newData = Replace(data, vbLf, vbCrLf)

frmMain.Text1.Text = frmMain.Text1.Text & newData
frmMain.Text1.Refresh
End Sub

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

Private Sub InitParams()
GSParamsIndex = 0
ReDim GSParams(GSParamsIndex)
End Sub

Private Sub AddParams(strValue As String)
GSParamsIndex = GSParamsIndex + 1
ReDim Preserve GSParams(GSParamsIndex)
GSParams(GSParamsIndex) = strValue
End Sub
