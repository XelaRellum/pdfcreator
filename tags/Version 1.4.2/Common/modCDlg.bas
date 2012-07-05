Attribute VB_Name = "modCDlg"
Option Explicit

Public Function OpenFileDialog(files As Collection, Optional InitFilename As String = "", _
 Optional Filter As String, Optional DefaultFileExtension As String = "*.*", _
 Optional InitDir As String = "", Optional DialogTitle As String = "", _
 Optional Flags As OpenSaveFlags, Optional hwnd As Long = 0, Optional FilterIndex As Long = 1) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'res = -1 : UserCancel
50020  'res > 0  : Count of Files
50030
50040  Dim buff As String, tFil As String, buffA() As String, i As Long, ofn As OPENFILENAME
50050  If Len(Filter) > 0 Then
50060    If InStr(Filter, "|") > 0 Then
50070      tFil = Replace(Filter, "|", vbNullChar)
50080      tFil = tFil & vbNullChar & vbNullChar
50090     Else
50100      tFil = Filter & vbNullChar & vbNullChar
50110    End If
50120   Else
50130    tFil = vbNullChar & vbNullChar
50140  End If
50150
50160  With ofn
50170   .nStructSize = Len(ofn)
50180   .hWndOwner = hwnd
50190   .sFilter = tFil
50200   .nFilterIndex = FilterIndex
50210   .sFile = InitFilename & Space$(1024) & vbNullChar & vbNullChar
50220   .nMaxFile = Len(.sFile)
50230   .sDefFileExt = DefaultFileExtension & vbNullChar & vbNullChar
50240   .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
50250   .nMaxTitle = Len(ofn.sFileTitle)
50260   If InitDir = vbNullString Then
50270     .sInitialDir = PDFCreatorApplicationPath & vbNullChar & vbNullChar
50280    Else
50290     .sInitialDir = InitDir & vbNullChar & vbNullChar
50300   End If
50310   .sDialogTitle = DialogTitle
50320   .Flags = Flags
50330  End With
50340
50350  Set files = New Collection
50360  If GetOpenFileName(ofn) <> 0 Then
50370    buff = Trim$(Replace$(Left$(ofn.sFile, Len(ofn.sFile) - 2), vbNullChar & vbNullChar, ""))
50380    Do While Right$(buff, 1) = vbNullChar
50390     buff = Mid(buff, 1, Len(buff) - 1)
50400     DoEvents
50410    Loop
50420    If Len(buff) > 3 Then
50430     If InStr(buff, vbNullChar) > 0 Then
50440       buffA = Split(buff, vbNullChar)
50450       For i = LBound(buffA) + 1 To UBound(buffA)
50460        If Len(buffA(i)) > 0 Then
50470         files.Add CompletePath(buffA(0)) & buffA(i)
50480        End If
50490       Next i
50500      Else
50510       files.Add buff
50520     End If
50530    End If
50540    OpenFileDialog = files.Count
50550   Else
50560    OpenFileDialog = -1
50570  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modCDlg", "OpenFileDialog")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function SaveFileDialog(filename As String, Optional InitFilename As String = "", _
 Optional Filter As String, Optional DefaultFileExtension As String = "*.*", _
 Optional InitDir As String = "", Optional DialogTitle As String = "", _
 Optional Flags As OpenSaveFlags, Optional hwnd As Long = 0, Optional FilterIndex As Long = 1) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'res = -1 : UserCancel
50020  'res > 0  : Ok
50030
50040  Dim buff As String, tFil As String, buffA() As String, i As Long, ofn As OPENFILENAME
50050  If Len(Filter) > 0 Then
50060    If InStr(Filter, "|") > 0 Then
50070      tFil = Replace(Filter, "|", vbNullChar)
50080      tFil = tFil & vbNullChar & vbNullChar
50090     Else
50100      tFil = Filter & vbNullChar & vbNullChar
50110    End If
50120   Else
50130    tFil = vbNullChar & vbNullChar
50140  End If
50150
50160  With ofn
50170   .nStructSize = Len(ofn)
50180   .hWndOwner = hwnd
50190   .sFilter = tFil
50200   .nFilterIndex = FilterIndex
50210   .sFile = InitFilename & Space$(1024) & vbNullChar & vbNullChar
50220   .nMaxFile = Len(.sFile)
50230   .sDefFileExt = DefaultFileExtension & vbNullChar & vbNullChar
50240   .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
50250   .nMaxTitle = Len(ofn.sFileTitle)
50260   If InitDir = vbNullString Then
50270     .sInitialDir = PDFCreatorApplicationPath & vbNullChar & vbNullChar
50280    Else
50290     .sInitialDir = InitDir & vbNullChar & vbNullChar
50300   End If
50310   .sDialogTitle = DialogTitle
50320   .Flags = Flags
50330  End With
50340
50350  If GetSaveFileName(ofn) <> 0 Then
50360    buff = Trim$(Replace$(Left$(ofn.sFile, Len(ofn.sFile) - 2), vbNullChar & vbNullChar, ""))
50370    Do While Right$(buff, 1) = vbNullChar
50380     buff = Mid(buff, 1, Len(buff) - 1)
50390     DoEvents
50400    Loop
50410    If Len(buff) > 3 Then
50420     filename = buff
50430    End If
50440    filename = Replace$(filename, "?", "_", , , vbTextCompare)
50450    SaveFileDialog = ofn.nFilterIndex
50460   Else
50470    SaveFileDialog = -1
50480  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modCDlg", "SaveFileDialog")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function OpenFontDialog(Font As tFont, Optional hwnd As Long = 0) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, TmpFName As String
50020  Static LFnt As LOGFONT, CF_T As tCHOOSEFONT
50030
50040  LFnt.lfUnderline = Font.Underline
50050  LFnt.lfStrikeOut = Font.Strikethrough
50060  LFnt.lfItalic = Font.Italic
50070  LFnt.lfHeight = Font.Size / (72 / (1440 / Screen.TwipsPerPixelY)) * -1
50080  If Font.Bold Then
50090    LFnt.lfWeight = 700
50100   Else
50110    LFnt.lfWeight = 300
50120  End If
50130
50140  With CF_T
50150   .Flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
50160   .hWndOwner = hwnd
50170   .lStructSize = Len(CF_T)
50180   .lpLogFont = VarPtr(LFnt)
50190   .hInstance = App.hInstance
50200   .hdc = Printer.hdc
50210   .nFontType = CF_SCREENFONTS
50220   .rgbColors = Font.Color
50230  End With
50240
50250  MoveMemory LFnt.lfFaceName(0), ByVal Font.Name, Len(Font.Name) + 1
50260
50270  res = CHOOSEFONT(CF_T)
50280  If res = 0 Then
50290   OpenFontDialog = -1
50300   Exit Function
50310  End If
50320
50330  TmpFName = StrConv(LFnt.lfFaceName, vbUnicode)
50340  With Font
50350   .Name = Left$(TmpFName, InStr(1, TmpFName, vbNullChar) - 1)
50360   .Bold = CBool(LFnt.lfWeight >= FW_BOLD)
50370   .Italic = CBool(LFnt.lfItalic)
50380   .Underline = CBool(LFnt.lfUnderline)
50390   .Strikethrough = CBool(LFnt.lfStrikeOut)
50400   .Size = CF_T.iPointSize / 10
50410   .Color = CF_T.rgbColors
50420  End With
50430  OpenFontDialog = 1
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modCDlg", "OpenFontDialog")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function OpenColorDialog(Color As OLE_COLOR, Optional hwnd As Long = 0) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nColor As OLE_COLOR, res As Long, cc As tCHOOSECOLOR, i As Long, tColor As OLE_COLOR
50020  Static BDF(15) As Long, initColors As Boolean
50030
50040  If OleTranslateColor(Color, 0, nColor) Then
50050   nColor = 0
50060  End If
50070
50080  If Not initColors Then
50090   BDF(0) = RGB(255, 0, 0)
50100   BDF(1) = RGB(255, 64, 64)
50110   BDF(2) = RGB(255, 128, 128)
50120   BDF(3) = RGB(255, 192, 192)
50130   BDF(4) = RGB(0, 255, 0)
50140   BDF(5) = RGB(64, 255, 64)
50150   BDF(6) = RGB(128, 255, 128)
50160   BDF(7) = RGB(192, 255, 192)
50170   BDF(8) = RGB(0, 0, 255)
50180   BDF(9) = RGB(64, 64, 255)
50190   BDF(10) = RGB(128, 128, 255)
50200   BDF(11) = RGB(192, 192, 255)
50210   BDF(12) = RGB(0, 0, 0)
50220   BDF(13) = RGB(64, 64, 64)
50230   BDF(14) = RGB(128, 128, 128)
50240   BDF(15) = RGB(192, 192, 192)
50250   initColors = True
50260  End If
50270
50280  With cc
50290   .lStructSize = Len(cc)
50300   .hInstance = App.hInstance
50310   .hWndOwner = hwnd
50320   .Flags = CC_RGBINIT Or CC_ANYCOLOR Or CC_FULLOPEN Or CC_PREVENTFULLOPEN
50330   .rgbResult = nColor
50340   .lpCustColors = VarPtr(BDF(0))
50350  End With
50360
50370  res = CHOOSECOLOR(cc)
50380
50390  If res <> 0 Then
50400    Color = cc.rgbResult
50410    OpenColorDialog = 1
50420   Else
50430    OpenColorDialog = -1
50440  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modCDlg", "OpenColorDialog")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
