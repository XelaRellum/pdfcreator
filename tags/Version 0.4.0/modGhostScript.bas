Attribute VB_Name = "modGhostScript"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOSIZE = &H1


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
Dim GSParams(45) As String
Dim GSRet

GSInit

GSParams(0) = vbNullString
GSParams(1) = "-I" & App.Path & "\lib;" & App.Path & "\fonts"
GSParams(2) = "-q"
GSParams(3) = "-dNOPAUSE"
GSParams(4) = "-dSAFER"
GSParams(5) = "-dBATCH"
GSParams(6) = "-sDEVICE=pdfwrite"
GSParams(7) = "-dPDFSETTINGS=/printer"
GSParams(8) = "-dCompatibilityLevel=" & GS_COMPATIBILITY
GSParams(9) = "-r" & GS_RESOLUTION & "x" & GS_RESOLUTION
GSParams(10) = "-dProcessColorModel=/Device" & GS_COLORMODEL
GSParams(11) = "-dAutoRotatePages=/" & GS_AUTOROTATE
GSParams(12) = "-dCompressPages=" & GS_COMPRESSPAGES
GSParams(13) = "-dEmbedAllFonts=" & GS_EMBEDALLFONTS
GSParams(14) = "-dSubsetFonts=" & GS_SUBSETFONTS
GSParams(15) = "-dMaxSubsetPct=" & GS_SUBSETFONTPERC
GSParams(16) = "-dConvertCMYKImagesToRGB=" & GS_CMYKTORGB
GSParams(17) = "-dColorImageResolution=" & GS_COLORRESOLUTION
GSParams(18) = "-dGrayImageResolution=" & GS_GREYRESOLUTION
GSParams(19) = "-dMonoImageResolution=" & GS_MONORESOLUTION
GSParams(20) = "-dDownsampleColorImages=" & GS_COLORRESAMPLE
GSParams(21) = "-dDownsampleGrayImages=" & GS_GREYRESAMPLE
GSParams(22) = "-dDownsampleMonoImages=" & GS_MONORESAMPLE
GSParams(23) = "-dColorImageDownsampleType=/" & GS_COLORRESAMPLEMETHOD
GSParams(24) = "-dGrayImageDownsampleType=/" & GS_GREYRESAMPLEMETHOD
GSParams(25) = "-dMonoImageDownsampleType=/" & GS_MONORESAMPLEMETHOD
GSParams(26) = "-dPreserveOverprintSettings=" & GS_PRESERVEOVERPRINT
GSParams(27) = "-dUCRandBGInfo=/Preserve"
GSParams(28) = "-dUseFlateCompression=true"
GSParams(29) = "-dParseDSCCommentsForDocInfo=true"
GSParams(30) = "-dParseDSCComments=true"
GSParams(31) = "-dOPM=" & GS_OVERPRINT
GSParams(32) = "-dOffOptimizations=0"
GSParams(33) = "-dLockDistillerParams=false"
GSParams(34) = "-dGrayImageDepth=-1"
GSParams(35) = "-dColorImageFilter=/DCTEncode"
GSParams(36) = "-dASCII85EncodePages=" & GS_ASCII85
GSParams(37) = "-dDefaultRenderingIntent=/Default"
GSParams(38) = "-dTransferFunctionInfo=/" & GS_TRANSFERFUNCTIONS
GSParams(39) = "-dPreserveHalftoneInfo=" & GS_HALFTONE
GSParams(40) = "-q"
GSParams(41) = "-dNOPAUSE"
GSParams(42) = "-dSAFER"
GSParams(43) = "-sOutputFile=" & GSOutputFile
GSParams(44) = GSInputFile
GSParams(45) = "-c quit"

frmMain.Refresh
'ShellRet = Shell(Chr(34) & App.Path & "\bin\GSWIN32C.EXE" & Chr(34) & GSParams)
'ShellRet = ExecCmdPipe(Chr(34) & App.Path & "\bin\GSWIN32C.EXE" & Chr(34) & GSParams)
    
GSRet = CallGS(GSParams)

'Debug.Print GSRet

If frmMain.chkDebug.Value = 1 Then
    'frmMain.Text1.Text = "Parameter:" & vbCrLf & GSParams & vbCrLf & App.Path & vbCrLf & Chr(34) & App.Path & "\bin\GSWIN32C.EXE" & Chr(34)
    'frmMain.Text2.Text = strGhostscriptErrors
End If

End Function

Public Function GetTitle(FileName As String)
Dim PosTitle
Dim PosTitleEnd
Dim strTM As String

Open FileName For Input As #1
   strTM = Input(1000, #1)
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
Dim strTM As String
Dim OldTitle As String

Open FileName For Input As #1
   strTM = Input(LOF(1) - 1, #1)
Close #1   ' Datei schlieﬂen.

PosTitle = InStr(1, strTM, "%%Title: ", vbTextCompare)
PosTitleEnd = InStr(PosTitle, strTM, vbCr, vbTextCompare)

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
