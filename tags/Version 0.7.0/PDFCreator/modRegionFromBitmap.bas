Attribute VB_Name = "modRegionFromBitmap"
'Code von Benjamin Wilger
'Benjamin@ActiveVB.de
'Copyright (C) 2001
Option Explicit
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
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
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Const LWA_COLORKEY = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Function MakeFormTransparent(Frm As Form, ByVal lngTransColor As Long)
    Dim hRegion As Long
    Dim WinStyle As Long

    'Systemfarben ggf. in RGB-Werte übersetzen
    If lngTransColor < 0 Then OleTranslateColor lngTransColor, 0&, lngTransColor

    'Ab Windows 2000/98 geht das relativ einfach per API
    'Mit IsFunctionExported wird geprüft, ob die Funktion
    'SetLayeredWindowAttributes unter diesem Betriebsystem unterstützt wird.
    If IsFunctionExported("SetLayeredWindowAttributes", "user32") Then
        'Den Fenster-Stil auf "Layered" setzen
        WinStyle = GetWindowLong(Frm.hWnd, GWL_EXSTYLE)
        WinStyle = WinStyle Or WS_EX_LAYERED
        SetWindowLong Frm.hWnd, GWL_EXSTYLE, WinStyle
        SetLayeredWindowAttributes Frm.hWnd, lngTransColor, 0&, LWA_COLORKEY

    Else 'Manuell die Region erstellen und übernehmen
        hRegion = RegionFromBitmap(Frm, lngTransColor)
        SetWindowRgn Frm.hWnd, hRegion, True
        DeleteObject hRegion
    End If
End Function

Private Function RegionFromBitmap(picSource As Object, ByVal lngTransColor As Long) As Long
    Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
    Dim lngRgnFinal As Long, lngRgnTmp As Long
    Dim lngStart As Long
    Dim x As Long, y As Long
    Dim hDC As Long

    Dim bi24BitInfo As BITMAPINFO
    Dim iBitmap As Long
    Dim BWidth As Long
    Dim BHeight As Long
    Dim iDC As Long
    Dim PicBits() As Byte
    Dim Col As Long
    Dim OldScaleMode As ScaleModeConstants

    OldScaleMode = picSource.ScaleMode
    picSource.ScaleMode = vbPixels

    hDC = picSource.hDC
    lngWidth = picSource.ScaleWidth '- 1
    lngHeight = picSource.ScaleHeight - 1

    BWidth = (picSource.ScaleWidth \ 4) * 4 + 4
    BHeight = picSource.ScaleHeight

    'Bitmap-Header
    With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = BWidth
        .biHeight = BHeight + 1
    End With
    'ByteArrays in der erforderlichen Größe anlegen
    ReDim PicBits(0 To bi24BitInfo.bmiHeader.biWidth * 3 - 1, 0 To bi24BitInfo.bmiHeader.biHeight - 1)

    iDC = CreateCompatibleDC(hDC)
    'Gerätekontextunabhängige Bitmap (DIB) erzeugen
    iBitmap = CreateDIBSection(iDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    'iBitmap in den neuen DIB-DC wählen
    Call SelectObject(iDC, iBitmap)
    'hDC des Quell-Fensters in den hDC der DIB kopieren
    Call BitBlt(iDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, hDC, 0, 0, vbSrcCopy)
    'Gerätekontextunabhängige Bitmap in ByteArrays kopieren
    Call GetDIBits(hDC, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, PicBits(0, 0), bi24BitInfo, DIB_RGB_COLORS)

    'Wir brauchen nur den Array, also können wir die Bitmap direkt wieder löschen.

    'DIB-DC
    Call DeleteDC(iDC)
    'Bitmap
    Call DeleteObject(iBitmap)

    lngRgnFinal = CreateRectRgn(0, 0, 0, 0)
    For y = 0 To lngHeight
        x = 0
        Do While x < lngWidth
            Do While x < lngWidth And _
                RGB(PicBits(x * 3 + 2, lngHeight - y + 1), _
                    PicBits(x * 3 + 1, lngHeight - y + 1), _
                    PicBits(x * 3, lngHeight - y + 1) _
                    ) = lngTransColor

                x = x + 1
            Loop
            If x <= lngWidth Then
                lngStart = x
                Do While x < lngWidth And _
                    RGB(PicBits(x * 3 + 2, lngHeight - y + 1), _
                        PicBits(x * 3 + 1, lngHeight - y + 1), _
                        PicBits(x * 3, lngHeight - y + 1) _
                        ) <> lngTransColor
                    x = x + 1
                Loop
                If x + 1 > lngWidth Then x = lngWidth
                lngRgnTmp = CreateRectRgn(lngStart, y, x, y + 1)
                lngRetr = CombineRgn(lngRgnFinal, lngRgnFinal, lngRgnTmp, RGN_OR)
                DeleteObject lngRgnTmp
            End If
        Loop
    Next

    picSource.ScaleMode = OldScaleMode
    RegionFromBitmap = lngRgnFinal
End Function

'Code von vbVision:
'Diese Funktion überprüft, ob die angegebene Function von einer DLL exportiert wird.
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    Dim hMod As Long, lpFunc As Long, bLibLoaded As Boolean

    'Handle der DLL erhalten
    hMod = GetModuleHandle(sModule)
    If hMod = 0 Then 'Falls DLL nicht registriert ...
        hMod = LoadLibrary(sModule) 'DLL in den Speicher laden.
        If hMod Then bLibLoaded = True
    End If

    If hMod Then
        If GetProcAddress(hMod, sFunction) Then IsFunctionExported = True
    End If

    If bLibLoaded Then Call FreeLibrary(hMod)

End Function

