Attribute VB_Name = "modRegionFromBitmap"
'Code von Benjamin Wilger
'Benjamin@ActiveVB.de
'Copyright (C) 2001
Option Explicit
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Const RGN_OR = 2
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" ( _
     ByVal clr As Long, _
     ByVal hpal As Long, _
     ByRef lpcolorref As Long)

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long

End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte

End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD

End Type
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0

Private Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long

Private Const LWA_COLORKEY = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Function MakeFormTransparent(Frm As Form, ByVal lngTransColor As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim hRegion As Long
50020     Dim WinStyle As Long
50030
50040     'Systemfarben ggf. in RGB-Werte �bersetzen
50050     If lngTransColor < 0 Then OleTranslateColor lngTransColor, 0&, lngTransColor
50060
50070     'Ab Windows 2000/98 geht das relativ einfach per API
50080     'Mit IsFunctionExported wird gepr�ft, ob die Funktion
50090     'SetLayeredWindowAttributes unter diesem Betriebsystem unterst�tzt wird.
50100     If IsFunctionExported("SetLayeredWindowAttributes", "user32") Then
50110         'Den Fenster-Stil auf "Layered" setzen
50120         WinStyle = GetWindowLong(Frm.hwnd, GWL_EXSTYLE)
50130         WinStyle = WinStyle Or WS_EX_LAYERED
50140         SetWindowLong Frm.hwnd, GWL_EXSTYLE, WinStyle
50150         SetLayeredWindowAttributes Frm.hwnd, lngTransColor, 0&, LWA_COLORKEY
50160
50170     Else 'Manuell die Region erstellen und �bernehmen
50180         hRegion = RegionFromBitmap(Frm, lngTransColor)
50190         SetWindowRgn Frm.hwnd, hRegion, True
50200         DeleteObject hRegion
50210     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRegionFromBitmap", "MakeFormTransparent")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function RegionFromBitmap(picSource As Object, ByVal lngTransColor As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
50020     Dim lngRgnFinal As Long, lngRgnTmp As Long
50030     Dim lngStart As Long
50040     Dim X As Long, Y As Long
50050     Dim hDC As Long
50060
50070     Dim bi24BitInfo As BITMAPINFO
50080     Dim iBitmap As Long
50090     Dim BWidth As Long
50100     Dim BHeight As Long
50110     Dim iDC As Long
50120     Dim PicBits() As Byte
50130     Dim Col As Long
50140     Dim OldScaleMode As ScaleModeConstants
50150
50160     OldScaleMode = picSource.ScaleMode
50170     picSource.ScaleMode = vbPixels
50180
50190     hDC = picSource.hDC
50200     lngWidth = picSource.ScaleWidth '- 1
50210     lngHeight = picSource.ScaleHeight - 1
50220
50230     BWidth = (picSource.ScaleWidth \ 4) * 4 + 4
50240     BHeight = picSource.ScaleHeight
50250
50260     'Bitmap-Header
50270     With bi24BitInfo.bmiHeader
50280         .biBitCount = 24
50290         .biCompression = BI_RGB
50300         .biPlanes = 1
50310         .biSize = Len(bi24BitInfo.bmiHeader)
50320         .biWidth = BWidth
50330         .biHeight = BHeight + 1
50340     End With
50350     'ByteArrays in der erforderlichen Gr��e anlegen
50360     ReDim PicBits(0 To bi24BitInfo.bmiHeader.biWidth * 3 - 1, 0 To bi24BitInfo.bmiHeader.biHeight - 1)
50370
50380     iDC = CreateCompatibleDC(hDC)
50390     'Ger�tekontextunabh�ngige Bitmap (DIB) erzeugen
50400     iBitmap = CreateDIBSection(iDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
50410     'iBitmap in den neuen DIB-DC w�hlen
50420     Call SelectObject(iDC, iBitmap)
50430     'hDC des Quell-Fensters in den hDC der DIB kopieren
50440     Call BitBlt(iDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, hDC, 0, 0, vbSrcCopy)
50450     'Ger�tekontextunabh�ngige Bitmap in ByteArrays kopieren
50460     Call GetDIBits(hDC, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, PicBits(0, 0), bi24BitInfo, DIB_RGB_COLORS)
50470
50480     'Wir brauchen nur den Array, also k�nnen wir die Bitmap direkt wieder l�schen.
50490
50500     'DIB-DC
50510     Call DeleteDC(iDC)
50520     'Bitmap
50530     Call DeleteObject(iBitmap)
50540
50550     lngRgnFinal = CreateRectRgn(0, 0, 0, 0)
50560     For Y = 0 To lngHeight
50570         X = 0
50580         Do While X < lngWidth
50590             Do While X < lngWidth And _
                RGB(PicBits(X * 3 + 2, lngHeight - Y + 1), _
                    PicBits(X * 3 + 1, lngHeight - Y + 1), _
                    PicBits(X * 3, lngHeight - Y + 1) _
                    ) = lngTransColor
50640
50650                 X = X + 1
50660             Loop
50670             If X <= lngWidth Then
50680                 lngStart = X
50690                 Do While X < lngWidth And _
                    RGB(PicBits(X * 3 + 2, lngHeight - Y + 1), _
                        PicBits(X * 3 + 1, lngHeight - Y + 1), _
                        PicBits(X * 3, lngHeight - Y + 1) _
                        ) <> lngTransColor
50740                     X = X + 1
50750                 Loop
50760                 If X + 1 > lngWidth Then X = lngWidth
50770                 lngRgnTmp = CreateRectRgn(lngStart, Y, X, Y + 1)
50780                 lngRetr = CombineRgn(lngRgnFinal, lngRgnFinal, lngRgnTmp, RGN_OR)
50790                 DeleteObject lngRgnTmp
50800             End If
50810         Loop
50820     Next
50830
50840     picSource.ScaleMode = OldScaleMode
50850     RegionFromBitmap = lngRgnFinal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRegionFromBitmap", "RegionFromBitmap")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

'Code von vbVision:
'Diese Funktion �berpr�ft, ob die angegebene Function von einer DLL exportiert wird.
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim hMod As Long, lpFunc As Long, bLibLoaded As Boolean
50020
50030     'Handle der DLL erhalten
50040     hMod = GetModuleHandle(sModule)
50050     If hMod = 0 Then 'Falls DLL nicht registriert ...
50060         hMod = LoadLibrary(sModule) 'DLL in den Speicher laden.
50070         If hMod Then bLibLoaded = True
50080     End If
50090
50100     If hMod Then
50110         If GetProcAddress(hMod, sFunction) Then IsFunctionExported = True
50120     End If
50130
50140     If bLibLoaded Then Call FreeLibrary(hMod)
50150
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRegionFromBitmap", "IsFunctionExported")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

