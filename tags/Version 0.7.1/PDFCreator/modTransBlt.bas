Attribute VB_Name = "modTransBlt"
'*************************************************************
'
'  CreditsScroller
'
'  © Copyright 2000-2001 by ActiveVB.de
'    Geschrieben von Herfried K. Wagner     02/02/01
'
'  Ich übernehmen keine Haftung für Schäden, die durch die
'  Verwendung dieses Codes verursacht wurden. Wenn Sie Fehler
'  im Code entdecken, Verbesserungsvorschläge haben, eine
'  interessante Weiterentwicklung gemacht haben, die Sie der
'  Öffentlichkeit zugänglich machen wollen, oder wenn Sie
'  Fragen zur Funktion haben, dann wenden Sie sich an eine
'  der folgenden Adressen oder posten Sie die Frage in
'  unserem Forum.
'
'  Hirf@ActiveVB.de
'  http://www.ActiveVB.de
'
'*************************************************************

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As _
        Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth _
        As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop _
        As Long) As Long
        
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC _
        As Long, ByVal crColor As Long) As Long
        
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal _
        hDC As Long) As Long
        
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As _
        Long) As Long
        
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth _
        As Long, ByVal nHeight As Long, ByVal nPlanes As Long, _
        ByVal nBitCount As Long, lpBits As Any) As Long
        
Private Declare Function CreateCompatibleBitmap Lib "gdi32" _
        (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight _
        As Long) As Long
        
Private Declare Function GetObj Lib "gdi32" Alias "GetObjectA" _
        (ByVal hObject As Long, ByVal nCount As Long, lpObject _
        As Any) As Long
        
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC _
        As Long, ByVal hObject As Long) As Long
        
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject _
        As Long) As Long
        
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC _
        As Long, ByVal nIndex As Long) As Long
        
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, _
        ByVal nBkMode As Long) As Long
        
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) _
        As Long
        
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As _
        Long, ByVal hDC As Long) As Long

'Constants used by new transparent support in NT.
Private Const CAPS1 = 94             'Other caps.
Private Const C1_TRANSPARENT = &H1   'New raster cap.
Private Const NEWTRANSPARENT = 3     'Use with SetBkMode().
Private Const OBJ_BITMAP = 7         'Used to retrieve HBITMAP from hDC.

'Ternary raster operations.
Private Const SRCCOPY = &HCC0020     '(DWORD) dest = source
Private Const SRCPAINT = &HEE0086    '(DWORD) dest = source OR dest
Private Const SRCAND = &H8800C6      '(DWORD) dest = source AND dest
Private Const NOTSRCCOPY = &H330008  '(DWORD) dest = (NOT source)


' 32-Bit Transparent BitBlt Function.
'
' Parameters
'
'   hDestDC         -  Destination device context.
'   x, y            -  Upper-left destination coordinates (pixels).
'   nWidth          -  Width of destination.
'   nHeight         -  Height of destination.
'   hSrcDC          -  Source device context.
'   xSrc, ySrc      -  Upper-left source coordinates (pixels).
'   lngTransColor   -  RGB value for transparent pixels, typically &HC0C0C0.
'
' Win98/2K have a TransparentBlt API function located in msimg32.dll
' but it doesn't deallocate system ressources after finishing (?!).

Public Sub TransBlt(ByVal hDestDC As Long, ByVal x As Long, _
                    ByVal Y As Long, ByVal nWidth As Long, _
                    ByVal nHeight As Long, ByVal hSrcDC As _
                    Long, ByVal xSrc As Long, ByVal ySrc As _
                    Long, ByVal lngTransColor As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020   Dim lngOrigColor As Long ' Holds original background color.
50030   Dim lngOrigMode As Long  ' Holds original background drawing mode.
50040
50050     If (GetDeviceCaps(hDestDC, CAPS1) And C1_TRANSPARENT) Then
50060
50070       ' Some NT machines support this *super* simple method!
50080       ' Save original settings, Blt, restore settings.
50090       lngOrigMode = SetBkMode(hDestDC, NEWTRANSPARENT)
50100       lngOrigColor = SetBkColor(hDestDC, lngTransColor)
50110       Call BitBlt(hDestDC, x, Y, nWidth, nHeight, hSrcDC, xSrc, ySrc, SRCCOPY)
50120       Call SetBkColor(hDestDC, lngOrigColor)
50130       Call SetBkMode(hDestDC, lngOrigMode)
50140     Else
50150       Dim lngSaveDC As Long           ' Backup copy of source bitmap.
50160       Dim lngMaskDC As Long           ' Mask bitmap (monochrome).
50170       Dim lngInvDC As Long            ' Inverse of mask bitmap (monochrome).
50180       Dim lngResultDC As Long         ' Combination of source bitmap & background.
50190       Dim lnghSaveBmp As Long         ' Bitmap stores backup copy of source bitmap.
50200       Dim lnghMaskBmp As Long         ' Bitmap stores mask (monochrome).
50210       Dim lnghInvBmp As Long          ' Bitmap holds inverse of mask (monochrome).
50220       Dim lnghResultBmp As Long       ' Bitmap combination of source & background.
50230       Dim lnghSavePrevBmp As Long     ' Holds previous bitmap in saved DC.
50240       Dim lnghMaskPrevBmp As Long     ' Holds previous bitmap in the mask DC.
50250       Dim lnghInvPrevBmp As Long      ' Holds previous bitmap in inverted mask DC.
50260       Dim lnghDestPrevBmp As Long     ' Holds previous bitmap in destination DC.
50270       Dim lngOriginalColor As Long    ' // Holds src's original BkColor.
50280
50290       ' // Not included in Karl E. Petersons example...
50300       ' // We need this to blit using the original colors.
50310       lngOriginalColor = SetBkColor(hSrcDC, vbWhite)
50320
50330       ' Create DCs to hold various stages of transformation.
50340       lngSaveDC = CreateCompatibleDC(hDestDC)
50350       lngMaskDC = CreateCompatibleDC(hDestDC)
50360       lngInvDC = CreateCompatibleDC(hDestDC)
50370       lngResultDC = CreateCompatibleDC(hDestDC)
50380
50390       ' Create monochrome bitmaps for the mask-related bitmaps.
50400       lnghMaskBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
50410       lnghInvBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
50420
50430       ' Create color bitmaps for final result & stored copy of source.
50440       lnghResultBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
50450       lnghSaveBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
50460
50470       ' Select bitmaps into DCs.
50480       lnghSavePrevBmp = SelectObject(lngSaveDC, lnghSaveBmp)
50490       lnghMaskPrevBmp = SelectObject(lngMaskDC, lnghMaskBmp)
50500       lnghInvPrevBmp = SelectObject(lngInvDC, lnghInvBmp)
50510       lnghDestPrevBmp = SelectObject(lngResultDC, lnghResultBmp)
50520
50530       ' Make backup of source bitmap to restore later.
50540       Call BitBlt(lngSaveDC, 0, 0, nWidth, nHeight, hSrcDC, xSrc, ySrc, SRCCOPY)
50550
50560       ' Create mask: set background color of source to transparent color.
50570       lngOrigColor = SetBkColor(hSrcDC, lngTransColor)
50580       Call BitBlt(lngMaskDC, 0, 0, nWidth, nHeight, hSrcDC, xSrc, ySrc, SRCCOPY)
50590       lngTransColor = SetBkColor(hSrcDC, lngOrigColor)
50600
50610       ' Create inverse of mask to AND w/ source & combine w/ background.
50620       Call BitBlt(lngInvDC, 0, 0, nWidth, nHeight, lngMaskDC, 0, 0, NOTSRCCOPY)
50630
50640       ' Copy background bitmap to result & create final transparent bitmap.
50650       Call BitBlt(lngResultDC, 0, 0, nWidth, nHeight, hDestDC, x, Y, SRCCOPY)
50660
50670       ' AND mask bitmap w/ result DC to punch hole in the background by painting black area for
50680       ' non-transparent portion of source bitmap.
50690       Call BitBlt(lngResultDC, 0, 0, nWidth, nHeight, lngMaskDC, 0, 0, SRCAND)
50700
50710       ' AND inverse mask w/ source bitmap to turn off bits associated with transparent area of
50720       ' source bitmap by making it black.
50730       Call BitBlt(hSrcDC, xSrc, ySrc, nWidth, nHeight, lngInvDC, 0, 0, SRCAND)
50740
50750       ' XOR result w/ source bitmap to make background show through.
50760       Call BitBlt(lngResultDC, 0, 0, nWidth, nHeight, hSrcDC, xSrc, ySrc, SRCPAINT)
50770
50780       ' Display transparent bitmap on background.
50790       Call BitBlt(hDestDC, x, Y, nWidth, nHeight, lngResultDC, 0, 0, SRCCOPY)
50800
50810       ' Restore backup of original bitmap.
50820       Call BitBlt(hSrcDC, xSrc, ySrc, nWidth, nHeight, lngSaveDC, 0, 0, SRCCOPY)
50830
50840       ' // Reset BkColor.
50850       Call SetBkColor(hSrcDC, lngOriginalColor)
50860
50870       ' Select original objects back.
50880       Call SelectObject(lngSaveDC, lnghSavePrevBmp)
50890       Call SelectObject(lngResultDC, lnghDestPrevBmp)
50900       Call SelectObject(lngMaskDC, lnghMaskPrevBmp)
50910       Call SelectObject(lngInvDC, lnghInvPrevBmp)
50920
50930       ' Deallocate system resources.
50940       Call DeleteObject(lnghSaveBmp)
50950       Call DeleteObject(lnghMaskBmp)
50960       Call DeleteObject(lnghInvBmp)
50970       Call DeleteObject(lnghResultBmp)
50980       Call DeleteDC(lngSaveDC)
50990       Call DeleteDC(lngInvDC)
51000       Call DeleteDC(lngMaskDC)
51010       Call DeleteDC(lngResultDC)
51020     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modTransBlt", "TransBlt")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
