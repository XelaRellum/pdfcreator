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
        Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth _
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

Public Sub TransBlt(ByVal hDestDC As Long, ByVal X As Long, _
                    ByVal Y As Long, ByVal nWidth As Long, _
                    ByVal nHeight As Long, ByVal hSrcDC As _
                    Long, ByVal xSrc As Long, ByVal ySrc As _
                    Long, ByVal lngTransColor As Long)

  Dim lngOrigColor As Long ' Holds original background color.
  Dim lngOrigMode As Long  ' Holds original background drawing mode.

    If (GetDeviceCaps(hDestDC, CAPS1) And C1_TRANSPARENT) Then

      ' Some NT machines support this *super* simple method!
      ' Save original settings, Blt, restore settings.
      lngOrigMode = SetBkMode(hDestDC, NEWTRANSPARENT)
      lngOrigColor = SetBkColor(hDestDC, lngTransColor)
      Call BitBlt(hDestDC, X, Y, nWidth, nHeight, hSrcDC, xSrc, ySrc, SRCCOPY)
      Call SetBkColor(hDestDC, lngOrigColor)
      Call SetBkMode(hDestDC, lngOrigMode)
    Else
      Dim lngSaveDC As Long           ' Backup copy of source bitmap.
      Dim lngMaskDC As Long           ' Mask bitmap (monochrome).
      Dim lngInvDC As Long            ' Inverse of mask bitmap (monochrome).
      Dim lngResultDC As Long         ' Combination of source bitmap & background.
      Dim lnghSaveBmp As Long         ' Bitmap stores backup copy of source bitmap.
      Dim lnghMaskBmp As Long         ' Bitmap stores mask (monochrome).
      Dim lnghInvBmp As Long          ' Bitmap holds inverse of mask (monochrome).
      Dim lnghResultBmp As Long       ' Bitmap combination of source & background.
      Dim lnghSavePrevBmp As Long     ' Holds previous bitmap in saved DC.
      Dim lnghMaskPrevBmp As Long     ' Holds previous bitmap in the mask DC.
      Dim lnghInvPrevBmp As Long      ' Holds previous bitmap in inverted mask DC.
      Dim lnghDestPrevBmp As Long     ' Holds previous bitmap in destination DC.
      Dim lngOriginalColor As Long    ' // Holds src's original BkColor.

      ' // Not included in Karl E. Petersons example...
      ' // We need this to blit using the original colors.
      lngOriginalColor = SetBkColor(hSrcDC, vbWhite)

      ' Create DCs to hold various stages of transformation.
      lngSaveDC = CreateCompatibleDC(hDestDC)
      lngMaskDC = CreateCompatibleDC(hDestDC)
      lngInvDC = CreateCompatibleDC(hDestDC)
      lngResultDC = CreateCompatibleDC(hDestDC)

      ' Create monochrome bitmaps for the mask-related bitmaps.
      lnghMaskBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
      lnghInvBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)

      ' Create color bitmaps for final result & stored copy of source.
      lnghResultBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
      lnghSaveBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)

      ' Select bitmaps into DCs.
      lnghSavePrevBmp = SelectObject(lngSaveDC, lnghSaveBmp)
      lnghMaskPrevBmp = SelectObject(lngMaskDC, lnghMaskBmp)
      lnghInvPrevBmp = SelectObject(lngInvDC, lnghInvBmp)
      lnghDestPrevBmp = SelectObject(lngResultDC, lnghResultBmp)

      ' Make backup of source bitmap to restore later.
      Call BitBlt(lngSaveDC, 0, 0, nWidth, nHeight, hSrcDC, xSrc, ySrc, SRCCOPY)

      ' Create mask: set background color of source to transparent color.
      lngOrigColor = SetBkColor(hSrcDC, lngTransColor)
      Call BitBlt(lngMaskDC, 0, 0, nWidth, nHeight, hSrcDC, xSrc, ySrc, SRCCOPY)
      lngTransColor = SetBkColor(hSrcDC, lngOrigColor)

      ' Create inverse of mask to AND w/ source & combine w/ background.
      Call BitBlt(lngInvDC, 0, 0, nWidth, nHeight, lngMaskDC, 0, 0, NOTSRCCOPY)

      ' Copy background bitmap to result & create final transparent bitmap.
      Call BitBlt(lngResultDC, 0, 0, nWidth, nHeight, hDestDC, X, Y, SRCCOPY)

      ' AND mask bitmap w/ result DC to punch hole in the background by painting black area for
      ' non-transparent portion of source bitmap.
      Call BitBlt(lngResultDC, 0, 0, nWidth, nHeight, lngMaskDC, 0, 0, SRCAND)

      ' AND inverse mask w/ source bitmap to turn off bits associated with transparent area of
      ' source bitmap by making it black.
      Call BitBlt(hSrcDC, xSrc, ySrc, nWidth, nHeight, lngInvDC, 0, 0, SRCAND)

      ' XOR result w/ source bitmap to make background show through.
      Call BitBlt(lngResultDC, 0, 0, nWidth, nHeight, hSrcDC, xSrc, ySrc, SRCPAINT)

      ' Display transparent bitmap on background.
      Call BitBlt(hDestDC, X, Y, nWidth, nHeight, lngResultDC, 0, 0, SRCCOPY)

      ' Restore backup of original bitmap.
      Call BitBlt(hSrcDC, xSrc, ySrc, nWidth, nHeight, lngSaveDC, 0, 0, SRCCOPY)

      ' // Reset BkColor.
      Call SetBkColor(hSrcDC, lngOriginalColor)

      ' Select original objects back.
      Call SelectObject(lngSaveDC, lnghSavePrevBmp)
      Call SelectObject(lngResultDC, lnghDestPrevBmp)
      Call SelectObject(lngMaskDC, lnghMaskPrevBmp)
      Call SelectObject(lngInvDC, lnghInvPrevBmp)

      ' Deallocate system resources.
      Call DeleteObject(lnghSaveBmp)
      Call DeleteObject(lnghMaskBmp)
      Call DeleteObject(lnghInvBmp)
      Call DeleteObject(lnghResultBmp)
      Call DeleteDC(lngSaveDC)
      Call DeleteDC(lngInvDC)
      Call DeleteDC(lngMaskDC)
      Call DeleteDC(lngResultDC)
    End If
End Sub
